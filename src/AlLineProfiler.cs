// Приёмник событий трассировки C/AL для профайлера строк.
//
// Зачем нужен: штатный Microsoft.Dynamics.Nav.EtwListener отдаёт наружу только
// NavPermissionEventReceiver (события проверки прав, 801) - ни номера строки, ни имени
// функции в его payload нет. Здесь тот же приём, что у Microsoft в Codeunit 9800:
// объект живёт внутри процесса NST, слушает ETW и отдаёт результат в C/AL.
//
// Главное отличие от образца: событие оператора C/AL приходит на КАЖДЫЙ выполненный
// оператор - в живом прогоне это 167 тысяч событий за 80 секунд. Отдавать их в C/AL
// поштучно нельзя (в том же CU 9800 Microsoft прямо предупреждает не открывать
// транзакций в обработчике), поэтому агрегация идёт здесь, а в C/AL уходят готовые
// строки - их сотни.
//
// Сборка: csc /target:library, ссылка на Microsoft.Diagnostics.Tracing.TraceEvent.dll
// из поставки NAV (Service\Add-ins\NavEtwReceiver). Компилятор - из состава .NET
// Framework, он понимает C# 5: без интерполяции строк, без ?. и без nameof.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Threading;
using Microsoft.Diagnostics.Tracing;
using Microsoft.Diagnostics.Tracing.Session;

[assembly: AssemblyTitle("AL Line Profiler ETW receiver")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace LineProfiler
{
    /// <summary>Род события трассировки. Определяется по СОСТАВУ ПОЛЕЙ payload,
    /// а не по номеру: на другой сборке платформы номера могут разъехаться.</summary>
    internal enum EventKind
    {
        Unknown = 0,
        Statement = 1,   // оператор C/AL: есть lineNumber
        FuncStart = 2,   // вход в функцию: есть functionName, opcode Start
        FuncStop = 3,    // выход из функции: есть functionName, opcode Stop
        SqlStart = 4,    // начало операции SQL: есть sqlStatement, opcode Start
        SqlStop = 5,     // конец операции SQL
        FuncFail = 6,    // функция оборвалась ошибкой: есть functionName, opcode не Start/Stop
        Ignore = 7
    }

    /// <summary>Накопленные метрики одной строки кода.</summary>
    internal sealed class LineStat
    {
        public int ObjectType;
        public int ObjectId;
        public int LineNo;
        public string FunctionName;
        public long Hits;
        public double TotalMs;
        public double SelfMs;
        public double SqlMs;
        public int SqlCount;
        public double MinMs;
        public double MaxMs;
        public long OuterHits;      // выполнений, не вложенных в своё же выполнение той же
                                    // строки: только они дают полное время
        public string Text;         // текст оператора из ETW: запасной листинг
    }

    /// <summary>Один РАЗЛИЧНЫЙ запрос SQL, порождённый строкой кода. Одинаковый текст
    /// с одной строки складывается в одну запись: в цикле запрос повторяется тысячи раз.</summary>
    internal sealed class SqlStat
    {
        public int ObjectType;
        public int ObjectId;
        public int LineNo;
        public string FunctionName;
        public string Statement;
        public bool Truncated;
        public long Executions;
        public double TotalMs;
        public double MaxMs;
    }

    /// <summary>Кадр вызова: одна выполняющаяся функция.</summary>
    internal sealed class Frame
    {
        public int ObjectType;
        public int ObjectId;
        public string Func;
        public double EnterMs;
        public bool OpenedByCall;   // кадр заведён событием входа, а не догадкой по оператору

        public bool HasStmt;        // в кадре есть незакрытый оператор
        public int StmtLine;
        public double StmtStart;
        public double StmtChildMs;  // время вложенных вызовов внутри текущего оператора
        public double SelfAccMs;    // своё время операторов ЭТОГО захода: доля ребра вызова

        public int SqlDepth;        // вложенность операций SQL
        public double SqlStart;
        public string SqlText;      // текст ВНЕШНЕЙ операции SQL: её длительность и меряем
        public string StmtText;     // текст текущего оператора, если его брали из ETW
    }

    /// <summary>Ключ вызванной функции внутри ребра. Регистр значим: сопоставление
    /// кадров в OnFuncStop идёт точным сравнением имён, и разойтись эти два места не должны.</summary>
    internal struct CalleeKey : IEquatable<CalleeKey>
    {
        public int Type;
        public int Id;
        public string Func;

        public bool Equals(CalleeKey o)
        {
            return Type == o.Type && Id == o.Id && string.Equals(Func, o.Func, StringComparison.Ordinal);
        }

        public override bool Equals(object o) { return o is CalleeKey && Equals((CalleeKey)o); }

        public override int GetHashCode()
        {
            unchecked { return (((Type * 397) ^ Id) * 397) ^ (Func == null ? 0 : Func.GetHashCode()); }
        }
    }

    /// <summary>Ребро вызова: что набрала вызванная функция при вызовах ИЗ ОДНОЙ строки.
    /// Построчная статистика этого не знает - там всё сложено по всем местам вызова.</summary>
    internal sealed class EdgeStat
    {
        public int CallerObjectType;
        public int CallerObjectId;
        public int CallerLineNo;
        public string CallerFunc;   // функция, из которой идёт вызов: по ней проходят сквозь
                                    // заглушку события, у которой нет ни одной строки в листинге
        public int CalleeObjectType;
        public int CalleeObjectId;
        public string CalleeFunc;
        public long Calls;
        public double TotalMs;
        public double SelfMs;
        public double MaxMs;
        public bool Recursive;      // хоть раз вызвана, когда её же кадр был в стеке
        public long CalleeTotalCalls;   // вызовов этой функции ОТОВСЮДУ: заполняется снимком
    }

    internal sealed class SessionState
    {
        public readonly List<Frame> Stack = new List<Frame>();
        public bool Contributed;    // из этой сессии в результат попала хотя бы одна строка
    }

    /// <summary>
    /// Слушает провайдер Microsoft-DynamicsNAV-Server и складывает время по строкам кода.
    /// Создаётся и управляется из C/AL.
    /// </summary>
    public sealed class AlLineProfilerReceiver
    {
        // Провайдер сервера NAV. GUID один и тот же начиная с NAV 2013.
        private static readonly Guid NavServerProvider =
            new Guid("85423fd1-c021-5a63-f214-c4819f8809f3");

        // ALFunctionCallTracing (0x8) + SqlTracing (0x2).
        private const ulong Keywords = 0xA;

        private readonly object _sync = new object();
        private readonly Dictionary<long, LineStat> _stats = new Dictionary<long, LineStat>();
        private readonly Dictionary<int, SessionState> _sessions = new Dictionary<int, SessionState>();
        private readonly Dictionary<int, EventKind> _kindById = new Dictionary<int, EventKind>();

        private TraceEventSession _session;
        private Thread _worker;
        private volatile bool _running;
        private volatile bool _replaying;   // идёт разбор файла .etl, живой сессии нет
        private volatile bool _stopping;    // идёт остановка: события ещё принимаем, ошибку разбора - нет

        private long _eventsSeen;
        private long _statementEvents;
        private long _statementEventsAll;   // до отбора по сессии: признак живой трассировки
        private long _sqlEvents;
        private long _orphanStops;
        private long _preTraceStops;     // выходов из кадров, открытых ДО начала трассы
        private long _unnamedEvents;
        private long _forcedPops;        // кадров снято чужим выходом: следы оборванных вызовов
        private long _unclosedFrames;    // кадров осталось открытыми: стек на момент остановки
        private long _unpairedSql;       // конец операции SQL без начала
        private long _failedFunctions;   // функций оборвалось ошибкой (событие 402)
        private int _staleSessionsKilled;
        private int _lostEvents = -1;   // -1: ни одного удачного чтения, значение неизвестно
        private string _lastError = "";
        private bool _providerEnabled;

        private int _sessionFilter;      // 0 - все сессии NAV
        private int _drainTimeoutMs = 8000;
        private int _targetObjectType = -1;
        private int _targetObjectId = -1;

        private List<LineStat> _snapshot = new List<LineStat>();
        private int _cursor = -1;

        // Тексты запросов: по ключу строки - словарь по тексту запроса. Двухуровнево,
        // чтобы не склеивать ключ из номера строки и текста в 2 КБ на каждом событии.
        private readonly Dictionary<long, Dictionary<string, SqlStat>> _sqlByLine =
            new Dictionary<long, Dictionary<string, SqlStat>>();
        private List<SqlStat> _querySnapshot = new List<SqlStat>();
        private int _queryCursor = -1;

        // Рёбра вызова: по ключу ВЫЗЫВАЮЩЕЙ строки - словарь по вызванной функции.
        // Двухуровнево по той же причине, что и тексты запросов: склеивать имя функции
        // в общий ключ на каждом снятии кадра дорого.
        private readonly Dictionary<long, Dictionary<CalleeKey, EdgeStat>> _edges =
            new Dictionary<long, Dictionary<CalleeKey, EdgeStat>>();
        private List<EdgeStat> _edgeSnapshot = new List<EdgeStat>();
        private int _edgeCursor = -1;
        private int _edgeRowsCollected;
        private int _edgeRowsDropped;
        private int _maxEdgeRows = 5000;
        private long _edgelessCalls;     // вызовов, у которых вызывающая строка неизвестна
        private int _sqlRowsCollected;
        private int _sqlRowsDropped;
        private bool _collectSqlText = true;
        private int _maxSqlTextLen = 8000;
        private int _maxSqlRows = 2000;
        private readonly List<string> _catalog = new List<string>();

        // Ключи строк, для которых текст оператора уже взят: читать строковое поле
        // payload на каждом из 167 тысяч событий незачем, хватает первого попадания.
        private readonly HashSet<long> _haveText = new HashSet<long>();
        private bool _collectLineText = true;
        private int _lineTextRows;

        // Имена типов объектов: в живых событиях objectType приходит СТРОКОЙ
        // («Report», «CodeUnit»), а не числом. Значения - как в dbo.[Object].[Type].
        private static readonly Dictionary<string, int> ObjTypeByName =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "TableData", 0 }, { "Table", 1 }, { "Form", 2 }, { "Report", 3 },
                { "Dataport", 4 }, { "CodeUnit", 5 }, { "XMLport", 6 }, { "MenuSuite", 7 },
                { "Page", 8 }, { "Query", 9 }, { "System", 10 }
            };

        // ---------------------------------------------------------------- управление

        /// <summary>Слушать только эту сессию NAV. 0 - все (тогда в замер попадёт
        /// и фоновая очередь заданий).</summary>
        /// <summary>Предел ожидания хвоста событий при остановке, мс. При фоновой нагрузке
        /// на инстанции поток событий не затихает, и остановка упирается в этот предел.</summary>
        public int DrainTimeoutMs
        {
            get { return _drainTimeoutMs; }
            set { _drainTimeoutMs = value < 500 ? 500 : value; }
        }

        public int SessionFilter
        {
            get { return _sessionFilter; }
            set { _sessionFilter = value; }
        }

        /// <summary>Собирать тексты запросов SQL. Выключение экономит память и время
        /// разбора: текст приезжает отдельной строкой на КАЖДОЕ обращение к базе.</summary>
        public bool CollectSqlText
        {
            get { return _collectSqlText; }
            set { _collectSqlText = value; }
        }

        /// <summary>Предел длины сохраняемого текста запроса, символов. Длинный SELECT
        /// с сотней полей иначе занял бы больше места, чем весь остальной замер.</summary>
        public int MaxSqlTextLength
        {
            get { return _maxSqlTextLen; }
            set { _maxSqlTextLen = value < 100 ? 100 : (value > 30000 ? 30000 : value); }
        }

        /// <summary>Предел числа РАЗЛИЧНЫХ запросов в замере.</summary>
        public int MaxSqlRows
        {
            get { return _maxSqlRows; }
            set { _maxSqlRows = value < 100 ? 100 : (value > 20000 ? 20000 : value); }
        }

        public int SqlRowsCollected { get { lock (_sync) { return _sqlRowsCollected; } } }

        /// <summary>Предел числа РАЗЛИЧНЫХ рёбер вызова в замере.</summary>
        public int MaxEdgeRows
        {
            get { return _maxEdgeRows; }
            set { _maxEdgeRows = value < 200 ? 200 : (value > 50000 ? 50000 : value); }
        }

        public int EdgeRowsCollected { get { lock (_sync) { return _edgeRowsCollected; } } }

        /// <summary>Сколько рёбер не поместилось в предел. Больше нуля - разбор по местам
        /// вызова неполон, и об этом надо сказать вслух.</summary>
        public int EdgeRowsDropped { get { lock (_sync) { return _edgeRowsDropped; } } }

        /// <summary>Вызовов, у которых вызывающий кадр был без открытого оператора: строку
        /// вызова приписать нечему. Много таких - события операторов терялись.</summary>
        public long EdgelessCalls { get { return Interlocked.Read(ref _edgelessCalls); } }

        /// <summary>Сколько различных запросов не поместилось в предел. Больше нуля -
        /// список запросов неполон, и об этом надо сказать вслух.</summary>
        public int SqlRowsDropped { get { lock (_sync) { return _sqlRowsDropped; } } }

        /// <summary>Перечень ВСЕХ разобранных родов событий: номер, опкод, состав полей.
        /// Провайдер своего манифеста наружу не отдаёт (TraceEvent из поставки NAV этого
        /// не умеет), поэтому единственный честный способ узнать имена полей - посмотреть
        /// на живые события.</summary>
        public string EventCatalog
        {
            get { lock (_sync) { return "filter=" + _sessionFilter + " || " + string.Join(" || ", _catalog.ToArray()); } }
        }

        public bool IsRunning { get { return _running; } }
        public long EventsSeen { get { return Interlocked.Read(ref _eventsSeen); } }
        public long StatementEvents { get { return Interlocked.Read(ref _statementEvents); } }

        /// <summary>Операторов C/AL во всём потоке, без отбора по сессии.</summary>
        public long StatementEventsAll { get { return Interlocked.Read(ref _statementEventsAll); } }
        public long SqlEvents { get { return Interlocked.Read(ref _sqlEvents); } }
        public long OrphanStops { get { return Interlocked.Read(ref _orphanStops); } }

        /// <summary>Выходов из функций, вход в которые случился до начала трассы.
        /// Это граница ее начала, а не брак: стек C/AL на старте уже не пуст.</summary>
        public long PreTraceStops { get { return Interlocked.Read(ref _preTraceStops); } }
        public long UnnamedEvents { get { return Interlocked.Read(ref _unnamedEvents); } }
        public long ForcedPops { get { return Interlocked.Read(ref _forcedPops); } }
        public long UnclosedFrames { get { return Interlocked.Read(ref _unclosedFrames); } }
        public long UnpairedSql { get { return Interlocked.Read(ref _unpairedSql); } }
        public long FailedFunctions { get { return Interlocked.Read(ref _failedFunctions); } }
        public int StaleSessionsKilled { get { return _staleSessionsKilled; } }

        /// <summary>Сколько РАЗНЫХ сессий NAV дало строки в результат. Больше одной -
        /// в замер попала чужая работа, и "Выполнений" суммарное по сессиям.</summary>
        public int DistinctSessions
        {
            get
            {
                lock (_sync)
                {
                    int n = 0;
                    foreach (KeyValuePair<int, SessionState> kv in _sessions)
                        if (kv.Value.Contributed) n++;
                    return n;
                }
            }
        }

        /// <summary>Брать ли текст оператора из ETW. Он нужен только как запасной
        /// листинг: штатный листинг даёт Code Coverage, и он полнее.</summary>
        public bool CollectLineText
        {
            get { return _collectLineText; }
            set { _collectLineText = value; }
        }

        public int LineTextRows { get { lock (_sync) { return _lineTextRows; } } }
        /// <summary>Событий, потерянных буферами ETW. -1 - неизвестно: ни одного удачного
        /// чтения не было. Ноль и «неизвестно» - разные вещи, и путать их нельзя.</summary>
        public int LostEvents { get { return _lostEvents; } }
        public string LastError { get { return _lastError; } }

        /// <summary>Провайдер удалось включить. Если false, сессия есть, а событий не будет.</summary>
        public bool ProviderEnabled { get { return _providerEnabled; } }

        /// <summary>Сессия ETW жива по мнению самой системы.</summary>
        public bool SessionActive
        {
            get { try { return _session != null && _session.IsActive; } catch { return false; } }
        }

        /// <summary>Копить строки только по одному объекту. Без вызова копятся все.</summary>
        public void SetTarget(int objectType, int objectId)
        {
            _targetObjectType = objectType;
            _targetObjectId = objectId;
        }

        public void ClearTarget()
        {
            _targetObjectType = -1;
            _targetObjectId = -1;
        }

        /// <summary>Поднимает сессию ETW и начинает слушать. Имя сессии уникально по
        /// идентификатору процесса: два NST на машине не подерутся за одно имя.</summary>
        public void Start()
        {
            if (_running) return;
            _lastError = "";

            // Словарь родов событий трогает только поток разбора, и только после этой
            // строки: чистить его из C/AL нельзя - Dictionary не потокобезопасен, и
            // читатель без блокировки может зациклиться внутри поиска.
            _kindById.Clear();
            _catalog.Clear();
            try
            {
                Clear();
                string name = "AlLineProfiler-" + System.Diagnostics.Process.GetCurrentProcess().Id;
                KillStaleSessions(name);

                // Create, а не Attach: чужую сессию не трогаем. StopOnDispose - чтобы
                // сессия не пережила процесс, если Stop не дойдёт.
                _session = new TraceEventSession(name, TraceEventSessionOptions.Create);
                _session.StopOnDispose = true;
                _providerEnabled = _session.EnableProvider(NavServerProvider, TraceEventLevel.Informational, Keywords, null);
                _session.Source.Dynamic.All += OnEvent;

                _running = true;
                _worker = new Thread(RunLoop);
                _worker.IsBackground = true;   // не держит остановку службы
                _worker.Name = "AlLineProfiler";
                _worker.Start();
            }
            catch (Exception ex)
            {
                _running = false;
                _lastError = ex.Message;
                SafeDisposeSession();
                throw;
            }
        }

        /// <summary>Гасит наши же сессии ETW, оставшиеся от прошлых запусков. Имя несёт
        /// идентификатор процесса, поэтому после падения NST или жёсткой остановки сессия
        /// с прежним номером живёт в системе и продолжает жечь буферы ядра, а имени этого
        /// уже никто не займёт. Чужие сессии не трогаем - только с нашей приставкой.</summary>
        private void KillStaleSessions(string keepName)
        {
            IList<string> names;
            try { names = TraceEventSession.GetActiveSessionNames(); }
            catch { return; }
            if (names == null) return;

            for (int i = 0; i < names.Count; i++)
            {
                string n = names[i];
                if (string.IsNullOrEmpty(n)) continue;
                if (!n.StartsWith("AlLineProfiler-", StringComparison.OrdinalIgnoreCase)) continue;
                if (string.Equals(n, keepName, StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    // Attach, а не Create: сессия уже есть, нам нужно только её погасить.
                    TraceEventSession stale = new TraceEventSession(n, TraceEventSessionOptions.Attach);
                    try { stale.Stop(true); }
                    finally { stale.Dispose(); }
                    _staleSessionsKilled++;
                }
                catch { }
            }
        }

        /// <summary>Останавливает сессию, дожидается разбора и готовит выборку строк.</summary>
        /// <summary>Ждёт, пока уже порождённые события доедут до приёмника, и только потом
        /// разрешает сбросить накопленное. Отсекать разогрев по МЕТКЕ ВРЕМЕНИ события
        /// нельзя: доставка ETW не упорядочена глобально, фоновое событие приезжает
        /// раньше своей очереди и задирает границу вперёд - тогда настоящие события
        /// замера отбрасываются как разогрев, а оператор, открытый до границы и закрытый
        /// после, получает время всего простоя. Границу проводит сам сброс: Clear стирает
        /// и накопленное, и стеки кадров, поэтому незакрытых операторов через неё не
        /// переходит.</summary>
        public void Settle()
        {
            if (!_running) return;
            // Flush здесь звать НЕЛЬЗЯ, хотя напрашивается. Проверено на стенде: принуди-
            // тельный сброс буферов ПОСРЕДИ живой сессии глушит доставку - события,
            // записанные после него, до потребителя больше не доезжают, сколько ни жди.
            // Замер выходил пустым примерно в половине прогонов, причём выглядел удачным:
            // ошибок нет, потерь нет, просто ноль времени. Хвост разогрева и так доедет
            // сам - доставка ETW отстаёт около секунды, и DrainQuiet её дожидается.
            DrainQuiet(_drainTimeoutMs);
        }

        public void Stop()
        {
            if (!_running) { BuildSnapshot(); return; }

            // Буферы ETW уходят потребителю по таймеру, около секунды, поэтому хвост
            // прогона надо дождаться. Принудительный Flush для этого НЕ годится, хотя
            // и напрашивается: проверено на стенде - он глушит доставку не только этой
            // сессии, но и следующей, и замер сразу после предыдущего выходил почти
            // пустым. Ждём естественного сброса: DrainQuiet требует, чтобы поток сперва
            // подрос, и только потом засчитывает тишину.
            DrainQuiet(_drainTimeoutMs);

            // Ворота приёмника (OnEvent) держим ОТКРЫТЫМИ до конца разбора: гасить
            // _running здесь нельзя. Хвост прогона приезжает как раз при остановке
            // сессии, и прежний порядок выбрасывал его прямо на входе в обработчик -
            // события доезжали, а их никто не принимал. Со стороны это выглядело как
            // "часть событий не дошла": счётчик выполнений у строки есть, времени нет.
            // Мерено опытом на демонстрационном отчёте: окно дожидания 500 мс - 63
            // строки из 147 без времени и 955 тысяч событий из 1,69 миллиона мимо,
            // причём сама ETW не потеряла НИ ОДНОГО.
            _stopping = true;
            // Порядок здесь принципиален. StopProcessing обрывает разбор НЕМЕДЛЕННО -
            // вместе с тем, что ещё лежит в буферах, - поэтому первым идёт Stop самой
            // сессии: он гасит поставщика и досылает потребителю остаток, после чего
            // Process() возвращается сам. Прежний порядок (сперва StopProcessing) резал
            // хвост прогона: на демонстрационном отчёте пропадали целые шаги, у которых
            // счётчик выполнений был, а времени не было вовсе.
            try
            {
                if (_session != null) _session.Stop(false);
            }
            catch (Exception ex) { _lastError = ex.Message; }

            if (_worker != null)
            {
                // Ждём, пока разбор ИДЁТ, а не фиксированные пять секунд: очередь в
                // миллион событий за них не разбирается, и обрыв по таймеру - это ровно
                // та же потеря хвоста, только на шаг позже.
                if (!JoinWhileParsing(_worker))
                {
                    try
                    {
                        if (_session != null && _session.Source != null) _session.Source.StopProcessing();
                    }
                    catch { }
                    if (!_worker.Join(10000)) _lastError = "разбор событий не завершился";
                }
                _worker = null;
            }
            _running = false;
            _stopping = false;
            SafeDisposeSession();
            CloseOpenFrames();
            BuildSnapshot();
        }

        /// <summary>Ждёт выхода потока разбора, пока тот делает работу. Признак работы -
        /// растущий счётчик событий: сессия уже остановлена, новых событий не появится,
        /// значит счётчик растёт ровно до тех пор, пока разбирается остаток буферов.
        /// Возвращает true, если разбор вышел сам.</summary>
        private bool JoinWhileParsing(Thread worker)
        {
            const int Step = 250;
            const int QuietMs = 3000;       // столько тишины считаем концом разбора
            const int HardCapMs = 120000;   // страховка от вечного ожидания

            long last = -1;
            int quiet = 0;
            int waited = 0;
            while (waited < HardCapMs)
            {
                if (worker.Join(Step)) return true;
                waited += Step;
                long now = Interlocked.Read(ref _eventsSeen);
                if (now == last) quiet += Step; else quiet = 0;
                last = now;
                if (quiet >= QuietMs) return false;
            }
            return false;
        }

        public void Clear()
        {
            lock (_sync)
            {
                _stats.Clear();
                _sessions.Clear();
                _snapshot = new List<LineStat>();
                _cursor = -1;
                _sqlByLine.Clear();
                _querySnapshot = new List<SqlStat>();
                _queryCursor = -1;
                _sqlRowsCollected = 0;
                _sqlRowsDropped = 0;
                _edges.Clear();
                _edgeSnapshot = new List<EdgeStat>();
                _edgeCursor = -1;
                _edgeRowsCollected = 0;
                _edgeRowsDropped = 0;
                _edgelessCalls = 0;
                _eventsSeen = 0;
                _statementEvents = 0;
                _statementEventsAll = 0;
                _sqlEvents = 0;
                _orphanStops = 0;
                _preTraceStops = 0;
                _unnamedEvents = 0;
                _forcedPops = 0;
                _unclosedFrames = 0;
                _unpairedSql = 0;
                _failedFunctions = 0;
                _haveText.Clear();
                _lineTextRows = 0;
                _lostEvents = -1;
            }
        }

        /// <summary>Ждёт, пока счётчик событий перестанет расти: значит хвост разобран.</summary>
        private void DrainQuiet(int maxMs)
        {
            // Событие едет к потребителю около секунды, поэтому "счётчик не вырос за
            // 100 мс" сразу после Flush ничего не доказывает: хвост прогона в этот момент
            // только в пути. Ждём не меньше минимума и требуем несколько тихих замеров
            // подряд - иначе последние строки сценария теряются.
            const int Step = 100;
            const int MinMs = 1500;
            // Буферы ETW уходят потребителю по таймеру, около секунды. Пауза КОРОЧЕ
            // этого периода - не конец потока, а обычное ожидание сброса. Прежние
            // четыре замера по 100 мс давали окно тишины 400 мс, то есть заведомо
            // меньше периода: приёмник считал прогон законченным и отбрасывал хвост.
            // Мерено на демонстрационном отчёте: чем КОРОЧЕ прогон, тем больше строк
            // оставалось без времени (139 мс прогона - 49 строк, 2,1 с - 32 строки).
            const int QuietSamples = 12;

            long last = -1;
            int quiet = 0;
            int waited = 0;
            // «Тихо» и «ещё ничего не приехало» - разные вещи, и раньше они путались.
            // Доставка ETW отстаёт примерно на секунду: сразу после Flush счётчик стоит
            // на месте, четыре тихих замера подряд набираются на пустом месте, и мы
            // закрывали сессию РАНЬШЕ, чем приезжали события прогона. Внешне это
            // выглядело как удачный замер с нулём времени - самый скверный вид отказа.
            // Теперь тишина засчитывается только ПОСЛЕ того, как поток хоть раз подрос.
            bool sawGrowth = false;
            while (waited < maxMs)
            {
                // Счётчик потерь снимаем ЗДЕСЬ, пока сессия жива: после Stop обращение
                // к EventsLost бросает COMException 0x80071069 («имя экземпляра не
                // распознано поставщиком WMI»), и прежнее чтение в SafeDisposeSession
                // молча съедалось пустым catch - индикатор потерь не загорался никогда.
                try { if (_session != null) _lostEvents = _session.EventsLost; }
                catch { }

                long now = Interlocked.Read(ref _eventsSeen);
                if (now == last)
                {
                    quiet++;
                }
                else
                {
                    quiet = 0;
                    if (last >= 0) sawGrowth = true;
                }
                last = now;
                if (waited >= MinMs && quiet >= QuietSamples && sawGrowth) return;
                Thread.Sleep(Step);
                waited += Step;
            }
        }

        private void RunLoop()
        {
            try { _session.Source.Process(); }
            catch (Exception ex)
            {
                // Штатная остановка роняет разбор изнутри - это не ошибка замера.
                // Раньше такое исключение показывалось пользователю после удачного прогона.
                // Проверяем и признак остановки: ворота теперь закрываются ПОСЛЕ разбора,
                // и на штатном выходе _running ещё поднят.
                if (_running && !_stopping) _lastError = ex.Message;
            }
        }

        private void SafeDisposeSession()
        {
            // Dispose в своём try: при ошибке чтения любого свойства сессии он раньше
            // не выполнялся вовсе, и сессия ETW переживала процесс.
            try
            {
                if (_session != null) _session.Dispose();
            }
            catch { }
            _session = null;
        }

        /// <summary>Прогоняет сохранённый файл трассировки через тот же обработчик, что и
        /// живую сессию. Нужен для двух вещей: проверить агрегатор без живой базы (сегодня
        /// склейку кадров нечем проверить, кроме прогона на стенде) и снять замер там, где
        /// класть свою сборку в службу не разрешили - трассу собирает logman, разбор идёт
        /// потом. Возвращает число строк в результате.</summary>
        public int ProcessEtlFile(string path)
        {
            if (_running) throw new InvalidOperationException("идёт живой замер: сначала остановите его");
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("путь к файлу трассировки пуст");
            if (!System.IO.File.Exists(path)) throw new System.IO.FileNotFoundException("файл трассировки не найден", path);

            _lastError = "";
            _kindById.Clear();
            _catalog.Clear();
            Clear();

            try
            {
                using (ETWTraceEventSource src = new ETWTraceEventSource(path))
                {
                    _replaying = true;
                    src.Dynamic.All += OnEvent;
                    src.Process();
                    // У файла потери считает сам источник, а не сессия: сессии уже нет.
                    _lostEvents = src.EventsLost;
                }
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                throw;
            }
            finally
            {
                _replaying = false;
            }

            CloseOpenFrames();
            BuildSnapshot();
            return RowCount;
        }

        // ---------------------------------------------------------------- разбор событий

        private void OnEvent(TraceEvent data)
        {
            if (!_running && !_replaying) return;
            if (data.ProviderGuid != NavServerProvider) return;

            Interlocked.Increment(ref _eventsSeen);

            EventKind kind = ClassifyCached(data);
            if (kind == EventKind.Unknown || kind == EventKind.Ignore) return;

            // Операторы считаем ДО отбора по сессии тоже: ожидание старта трассировки
            // проверяет по этому счётчику, пошёл ли поток вообще. При замере ЧУЖОЙ
            // сессии свои операторы отбор выбрасывает, и счётчик замера остался бы
            // нулём - ожидание решило бы, что трассировка выключена, и сорвало прогон.
            if (kind == EventKind.Statement) Interlocked.Increment(ref _statementEventsAll);

            int sid = ToInt(SafePayload(data, "sessionId"), -1);
            if (_sessionFilter != 0 && sid != _sessionFilter) return;

            double now = data.TimeStampRelativeMSec;

            lock (_sync)
            {
                SessionState st = GetSession(sid);
                switch (kind)
                {
                    case EventKind.Statement:
                        Interlocked.Increment(ref _statementEvents);
                        OnStatement(st, data, now);
                        break;
                    case EventKind.FuncStart:
                        OnFuncStart(st, data, now);
                        break;
                    case EventKind.FuncStop:
                        OnFuncStop(st, data, now);
                        break;
                    case EventKind.SqlStart:
                        Interlocked.Increment(ref _sqlEvents);
                        OnSqlStart(st, data, now);
                        break;
                    case EventKind.SqlStop:
                        OnSqlStop(st, data, now);
                        break;
                    case EventKind.FuncFail:
                        // Функция оборвалась ошибкой C/AL. Кадр надо снять тем же путём,
                        // что и при обычном выходе, иначе время до конца трассы осядет
                        // на вызывающей строке. Считаем отдельно: замер с обрывами
                        // доверия стоит меньше, и это должно быть видно на странице.
                        Interlocked.Increment(ref _failedFunctions);
                        OnFuncStop(st, data, now);
                        break;
                }
            }
        }

        /// <summary>Род события определяется один раз на номер события и кэшируется:
        /// перебирать имена полей 167 тысяч раз незачем.</summary>
        private EventKind ClassifyCached(TraceEvent data)
        {
            int id = (int)data.ID;
            EventKind kind;
            if (_kindById.TryGetValue(id, out kind)) return kind;

            string[] names;
            try { names = data.PayloadNames; }
            catch { names = null; }

            if (names == null || names.Length == 0)
            {
                // Схема провайдера приезжает в поток отдельным манифестом, и первые
                // события успевают его обогнать - имён полей у них ещё нет. Такой
                // вердикт кэшировать нельзя: номер события навсегда остался бы
                // "неинтересным", и весь прогон вышел бы пустым. Пропускаем событие
                // и определяем род заново на следующем с тем же номером.
                Interlocked.Increment(ref _unnamedEvents);
                return EventKind.Unknown;
            }

            kind = Classify(names, data.Opcode);
            lock (_sync)
            {
                _kindById[id] = kind;
                // Только имена полей: в значениях едут тексты запросов и имена
                // пользователей, а это диагностическая строка, а не журнал.
                if (_catalog.Count < 64)
                    _catalog.Add("id=" + id + " op=" + ((int)data.Opcode) +
                        " task=" + ((int)data.Task) + " kind=" + kind +
                        " fields=" + string.Join(",", names));
            }
            return kind;
        }

        private static EventKind Classify(string[] names, TraceEventOpcode opcode)
        {
            bool hasLine = false, hasFunc = false, hasSql = false;
            for (int i = 0; i < names.Length; i++)
            {
                string n = names[i];
                if (string.Equals(n, "lineNumber", StringComparison.OrdinalIgnoreCase)) hasLine = true;
                else if (string.Equals(n, "functionName", StringComparison.OrdinalIgnoreCase)) hasFunc = true;
                else if (string.Equals(n, "sqlStatement", StringComparison.OrdinalIgnoreCase)) hasSql = true;
            }

            if (hasLine && hasFunc) return EventKind.Statement;
            if (hasFunc)
            {
                if (opcode == TraceEventOpcode.Start) return EventKind.FuncStart;
                if (opcode == TraceEventOpcode.Stop) return EventKind.FuncStop;
                // Третий род события функции: обрыв ошибкой (у платформы это 402).
                // Опознаём по форме, а не по номеру: под нашей маской 0xA имя функции
                // без номера строки несут только события входа, выхода и обрыва.
                if (!hasSql) return EventKind.FuncFail;
                return EventKind.Ignore;
            }
            if (hasSql)
            {
                if (opcode == TraceEventOpcode.Start) return EventKind.SqlStart;
                if (opcode == TraceEventOpcode.Stop) return EventKind.SqlStop;
                return EventKind.Ignore;
            }

            // Транзакционные операции - COMMIT, BEGIN TRANSACTION, ROLLBACK, открытие
            // соединения - идут под тем же ключевым словом SqlTracing, но поля с текстом
            // запроса у них НЕТ. Раньше они уходили в Ignore, и строка, вся стоимость
            // которой в коммите, показывала ноль времени SQL. Опознаём по форме события:
            // под нашей маской 0xA всё, что имеет пару Start/Stop и не относится ни к
            // оператору C/AL, ни к функции, - это операция с базой.
            if (opcode == TraceEventOpcode.Start) return EventKind.SqlStart;
            if (opcode == TraceEventOpcode.Stop) return EventKind.SqlStop;
            return EventKind.Ignore;
        }

        private void OnStatement(SessionState st, TraceEvent data, double now)
        {
            int objType = ToObjType(SafePayload(data, "objectType"));
            int objId = ToInt(SafePayload(data, "objectId"), -1);
            string func = ToStr(SafePayload(data, "functionName"));
            int line = ToInt(SafePayload(data, "lineNumber"), -1);
            if (objId < 0 || line < 0) return;

            Frame f = Top(st);
            if (f == null || f.ObjectType != objType || f.ObjectId != objId || f.Func != func)
            {
                // Оператор пришёл из другой функции, а события входа не было. Такое бывает,
                // если трассировка вызовов выключена или начало прогона срезано кольцевым
                // буфером. Кадр заводим сами и помечаем, что он не от события входа.
                if (f != null && !f.OpenedByCall)
                {
                    CloseStatement(st, st.Stack.Count - 1, now);
                    st.Stack.RemoveAt(st.Stack.Count - 1);
                }
                f = new Frame();
                f.ObjectType = objType;
                f.ObjectId = objId;
                f.Func = func;
                f.EnterMs = now;
                f.OpenedByCall = false;
                st.Stack.Add(f);
                TrimStack(st, now);
            }
            else
            {
                CloseStatement(st, st.Stack.Count - 1, now);
            }

            f.StmtLine = line;
            f.StmtStart = now;
            f.StmtChildMs = 0;
            f.HasStmt = true;
            f.StmtText = null;

            if (Interesting(objType, objId))
            {
                st.Contributed = true;
                // Текст оператора берём ОДИН раз на строку: строковое поле payload на
                // каждом событии - это лишние сотни тысяч разборов ради того, что
                // уже известно. Нужен он только там, где листинга из Code Coverage нет.
                if (_collectLineText)
                {
                    long key = LineKey(objType, objId, line);
                    if (!_haveText.Contains(key))
                    {
                        _haveText.Add(key);
                        f.StmtText = ToStr(SafePayload(data, "statement"));
                    }
                }
            }
        }

        private void OnFuncStart(SessionState st, TraceEvent data, double now)
        {
            Frame f = new Frame();
            f.ObjectType = ToObjType(SafePayload(data, "objectType"));
            f.ObjectId = ToInt(SafePayload(data, "objectId"), -1);
            f.Func = ToStr(SafePayload(data, "functionName"));
            f.EnterMs = now;
            f.OpenedByCall = true;
            st.Stack.Add(f);
            TrimStack(st, now);
        }

        private void OnFuncStop(SessionState st, TraceEvent data, double now)
        {
            int objType = ToObjType(SafePayload(data, "objectType"));
            int objId = ToInt(SafePayload(data, "objectId"), -1);
            string func = ToStr(SafePayload(data, "functionName"));

            // Ищем СВОЙ кадр сверху вниз и снимаем всё, что осталось над ним. Без этого
            // кадр, брошенный неперехваченной ошибкой C/AL, не снимается никогда, и время
            // вызывающей строки вбирает весь остаток трассы вместе с простоем.
            int idx = -1;
            for (int i = st.Stack.Count - 1; i >= 0; i--)
            {
                Frame c = st.Stack[i];
                if (c.ObjectType == objType && c.ObjectId == objId && c.Func == func) { idx = i; break; }
            }
            if (idx < 0)
            {
                // Выход из функции, в которую вошли ДО начала трассы: на старте стек C/AL
                // уже был не пуст, и тех входов мы не видели. Узнаётся по пустому стеку
                // сессии - раскручивается то, что лежало ниже точки старта. Браком это не
                // является, и в здоровье замера ему не место. Если же стек НЕ пуст, а
                // своего кадра в нём нет, - это потерянный вход, и вот он уже брак.
                if (st.Stack.Count == 0) Interlocked.Increment(ref _preTraceStops);
                else Interlocked.Increment(ref _orphanStops);
                return;
            }

            // Всё, что выше найденного кадра, снимается принудительно - это следы
            // вызовов, чей выход до нас не доехал. Считаем их: много таких снятий
            // означает рваную трассу и заниженное время вложенных строк.
            if (st.Stack.Count - 1 > idx)
                Interlocked.Add(ref _forcedPops, st.Stack.Count - 1 - idx);

            while (st.Stack.Count - 1 >= idx) PopFrame(st, now);
        }

        private void PopFrame(SessionState st, double now)
        {
            int last = st.Stack.Count - 1;
            if (last < 0) return;
            Frame f = st.Stack[last];
            CloseStatement(st, last, now);

            double subtree = now - f.EnterMs;
            if (subtree < 0) subtree = 0;
            st.Stack.RemoveAt(last);

            Frame parent = Top(st);
            if (parent != null && parent.HasStmt)
            {
                parent.StmtChildMs += subtree;
                // Ребро пишем, только если ОБА конца попали в отбор: вызов в чужой объект
                // показать негде, а место в словаре он занял бы.
                if (Interesting(parent.ObjectType, parent.ObjectId) &&
                    Interesting(f.ObjectType, f.ObjectId))
                    RecordEdge(st, parent, f, subtree);
            }
            else if (parent != null)
                // Вызывающий кадр есть, а открытого оператора в нём нет: строку вызова
                // приписать нечему. Пустой стек (parent == null) - это снятие точки входа,
                // норма на каждом прогоне, и в счётчик оно не идёт.
                Interlocked.Increment(ref _edgelessCalls);
        }

        /// <summary>Приписывает время снятого кадра ребру «строка вызова -> функция».</summary>
        private void RecordEdge(SessionState st, Frame parent, Frame callee, double subtree)
        {
            long key = LineKey(parent.ObjectType, parent.ObjectId, parent.StmtLine);
            CalleeKey ck;
            ck.Type = callee.ObjectType;
            ck.Id = callee.ObjectId;
            ck.Func = callee.Func;

            Dictionary<CalleeKey, EdgeStat> byCallee;
            if (!_edges.TryGetValue(key, out byCallee))
            {
                // Предел проверяем ДО того, как завести словарь: иначе после переполнения
                // каждая новая строка навсегда оставляет за собой пустую заготовку.
                if (_edgeRowsCollected >= _maxEdgeRows) { _edgeRowsDropped++; return; }
                byCallee = new Dictionary<CalleeKey, EdgeStat>();
                _edges[key] = byCallee;
            }

            EdgeStat e;
            if (!byCallee.TryGetValue(ck, out e))
            {
                if (_edgeRowsCollected >= _maxEdgeRows) { _edgeRowsDropped++; return; }
                e = new EdgeStat();
                e.CallerObjectType = parent.ObjectType;
                e.CallerObjectId = parent.ObjectId;
                e.CallerLineNo = parent.StmtLine;
                e.CallerFunc = parent.Func;
                e.CalleeObjectType = callee.ObjectType;
                e.CalleeObjectId = callee.ObjectId;
                e.CalleeFunc = callee.Func;
                byCallee[ck] = e;
                _edgeRowsCollected++;
            }
            // Кадр той же функции ещё в стеке - это виток рекурсии, и его время внешний
            // виток уже содержит целиком. Полное время ребра копим только с ВНЕШНЕГО
            // витка, иначе оно вышло бы кратным глубине. Вызовы и своё время идут все:
            // вызов был настоящий, а своё время витков не пересекается.
            //
            // Внимание: у СТРОКИ (CloseStatement) правило витка ДРУГОЕ - там ищется та же
            // СТРОКА в кадре-предке, а здесь та же ФУНКЦИЯ где угодно в стеке. Расходятся
            // они ровно на самом внешнем витке: строка рекурсивного вызова в предках себя
            // не находит и отдаёт всё поддерево, а ребро из неё в ту же функцию находит её
            // кадр всегда и не отдаёт ничего. Отсюда экран, где у строки вызова полное
            // время есть, а у ребра из неё - ноль. Числа посчитаны на разных множествах, и
            // складывать или сравнивать их нельзя.
            bool nested = false;
            for (int i = 0; i < st.Stack.Count; i++)
            {
                Frame c = st.Stack[i];
                if (c.ObjectType == callee.ObjectType && c.ObjectId == callee.ObjectId &&
                    string.Equals(c.Func, callee.Func, StringComparison.Ordinal))
                { nested = true; break; }
            }

            e.Calls++;
            e.SelfMs += callee.SelfAccMs;
            if (nested) e.Recursive = true;
            else
            {
                e.TotalMs += subtree;
                if (subtree > e.MaxMs) e.MaxMs = subtree;
            }
        }

        /// <summary>Закрывает текущий оператор кадра и разносит его время по строке.</summary>
        private void CloseStatement(SessionState st, int index, double now)
        {
            Frame f = st.Stack[index];
            if (!f.HasStmt) return;

            double total = now - f.StmtStart;
            if (total < 0) total = 0;
            double self = total - f.StmtChildMs;
            if (self < 0) self = 0;

            // Виток рекурсии: та же строка того же объекта уже открыта в кадре-предке.
            // Своё время у каждого витка своё - они не пересекаются и складываются
            // честно. А вот ПОЛНОЕ время внешнего витка содержит все внутренние целиком,
            // и складывая их, строка получала время, кратное глубине: на глубине 30 оно
            // обгоняло собственный вызов, внутри которого вся рекурсия и происходит.
            // Поэтому полное время берём только с ВНЕШНЕГО витка.
            bool nested = false;
            for (int i = 0; i < index; i++)
            {
                Frame c = st.Stack[i];
                if (c.HasStmt && c.StmtLine == f.StmtLine &&
                    c.ObjectType == f.ObjectType && c.ObjectId == f.ObjectId)
                { nested = true; break; }
            }
            // Копим ДО отбора по объекту: ребро может вести в функцию, которая сама в
            // отбор не попала, а её долю в строке вызова показать всё равно надо.
            f.SelfAccMs += self;

            if (Interesting(f.ObjectType, f.ObjectId))
            {
                LineStat s = GetStat(f.ObjectType, f.ObjectId, f.StmtLine, f.Func);
                s.Hits++;
                s.SelfMs += self;
                if (!nested)
                {
                    s.OuterHits++;
                    s.TotalMs += total;
                    // Минимум заводится по ВНЕШНИМ выполнениям: у вложенных полного
                    // времени нет вовсе, и нулём они сбили бы и минимум, и максимум.
                    if (s.OuterHits == 1 || total < s.MinMs) s.MinMs = total;
                    if (total > s.MaxMs) s.MaxMs = total;
                }
                if (s.Text == null && !string.IsNullOrEmpty(f.StmtText))
                {
                    s.Text = f.StmtText.Length > 500 ? f.StmtText.Substring(0, 500) : f.StmtText;
                    _lineTextRows++;
                }
            }

            f.HasStmt = false;
            f.StmtText = null;
            f.StmtChildMs = 0;
            // Ожидающее начало SQL живёт не дольше своего оператора: иначе одна
            // потерянная пара Start/Stop навсегда сбивает учёт SQL в этой сессии.
            f.SqlDepth = 0;
            f.SqlText = null;
        }

        private void OnSqlStart(SessionState st, TraceEvent data, double now)
        {
            Frame f = Top(st);
            if (f == null) return;
            if (f.SqlDepth == 0)
            {
                f.SqlStart = now;
                // Текст берём только у ВНЕШНЕЙ операции: именно её длительность меряется.
                f.SqlText = _collectSqlText ? SqlTextOf(data) : null;
            }
            f.SqlDepth++;
        }

        private void OnSqlStop(SessionState st, TraceEvent data, double now)
        {
            Frame f = Top(st);
            // Кадра нет вовсе - это запрос без контекста C/AL: фоновые задания, работа
            // самой платформы, чужие сессии. Таких на полной трассе десятки тысяч, и
            // считать их отказом нельзя - счётчик перестал бы что-либо значить.
            if (f == null) return;
            if (f.SqlDepth <= 0)
            {
                Interlocked.Increment(ref _unpairedSql);   // хвост без начала ВНУТРИ кадра
                return;
            }
            f.SqlDepth--;
            if (f.SqlDepth > 0) return;                 // внутренняя операция: время уже учтено внешней

            double d = now - f.SqlStart;
            if (d < 0) d = 0;
            if (f.HasStmt && Interesting(f.ObjectType, f.ObjectId))
            {
                LineStat s = GetStat(f.ObjectType, f.ObjectId, f.StmtLine, f.Func);
                s.SqlMs += d;
                s.SqlCount++;
                // Текст запроса несёт начало пары, но если там пусто - пробуем конец.
                if (_collectSqlText && string.IsNullOrEmpty(f.SqlText))
                    f.SqlText = SqlTextOf(data);
                RecordSql(f, d);
            }
            f.SqlText = null;
        }

        /// <summary>Текст операции с базой. У транзакционных событий поля с запросом нет,
        /// и вместо пустоты подставляем имя операции - иначе запись отбрасывается по
        /// пустому тексту, и в списке запросов строки просто не будет.</summary>
        private static string SqlTextOf(TraceEvent data)
        {
            string text = ToStr(SafePayload(data, "sqlStatement"));
            if (!string.IsNullOrEmpty(text)) return text;
            string task = "";
            try { task = data.TaskName; }
            catch { }
            if (string.IsNullOrEmpty(task)) return "";
            return "-- " + task.ToUpperInvariant();
        }

        /// <summary>Складывает запрос в копилку своей строки кода. Одинаковый текст
        /// с одной строки - одна запись со счётчиком выполнений.</summary>
        private void RecordSql(Frame f, double ms)
        {
            if (!_collectSqlText) return;
            string text = f.SqlText;
            if (string.IsNullOrEmpty(text)) return;

            bool cut = false;
            if (text.Length > _maxSqlTextLen)
            {
                text = text.Substring(0, _maxSqlTextLen);
                cut = true;
            }

            long key = LineKey(f.ObjectType, f.ObjectId, f.StmtLine);
            Dictionary<string, SqlStat> byText;
            if (!_sqlByLine.TryGetValue(key, out byText))
            {
                // Предел не про экономию, а про выживание процесса NST: без него прогон
                // по чужому коду с уникальными запросами съел бы всю его память. Проверяем
                // ДО заведения словаря: раньше после переполнения каждая новая строка
                // навсегда оставляла за собой пустую заготовку.
                if (_sqlRowsCollected >= _maxSqlRows) { _sqlRowsDropped++; return; }
                byText = new Dictionary<string, SqlStat>(StringComparer.Ordinal);
                _sqlByLine[key] = byText;
            }

            SqlStat q;
            if (!byText.TryGetValue(text, out q))
            {
                if (_sqlRowsCollected >= _maxSqlRows) { _sqlRowsDropped++; return; }
                q = new SqlStat();
                q.ObjectType = f.ObjectType;
                q.ObjectId = f.ObjectId;
                q.LineNo = f.StmtLine;
                q.FunctionName = f.Func;
                q.Statement = text;
                q.Truncated = cut;
                byText[text] = q;
                _sqlRowsCollected++;
            }
            q.Executions++;
            q.TotalMs += ms;
            if (ms > q.MaxMs) q.MaxMs = ms;
        }

        /// <summary>Аварийный ограничитель: без событий выхода стек рос бы бесконечно.</summary>
        private void TrimStack(SessionState st, double now)
        {
            const int MaxDepth = 512;
            while (st.Stack.Count > MaxDepth)
            {
                CloseStatement(st, 0, now);
                st.Stack.RemoveAt(0);
            }
        }

        private void CloseOpenFrames()
        {
            lock (_sync)
            {
                foreach (KeyValuePair<int, SessionState> kv in _sessions)
                {
                    SessionState st = kv.Value;
                    if (st.Stack.Count > 0)
                        Interlocked.Add(ref _unclosedFrames, st.Stack.Count);
                    while (st.Stack.Count > 0) PopFrame(st, LastSeenMs(st));
                }
            }
        }

        private static double LastSeenMs(SessionState st)
        {
            // Прогон оборван - берём момент входа самого верхнего кадра: приписывать
            // строке время до «сейчас» нечестно, событий после обрыва не было.
            double m = 0;
            for (int i = 0; i < st.Stack.Count; i++)
            {
                Frame f = st.Stack[i];
                if (f.EnterMs > m) m = f.EnterMs;
                if (f.HasStmt && f.StmtStart > m) m = f.StmtStart;
            }
            return m;
        }

        // ---------------------------------------------------------------- выборка для C/AL

        private void BuildSnapshot()
        {
            lock (_sync)
            {
                List<LineStat> list = new List<LineStat>(_stats.Values);
                list.Sort(CompareStat);
                _snapshot = list;
                _cursor = -1;
            }
        }

        private static int CompareStat(LineStat a, LineStat b)
        {
            if (a.ObjectType != b.ObjectType) return a.ObjectType.CompareTo(b.ObjectType);
            if (a.ObjectId != b.ObjectId) return a.ObjectId.CompareTo(b.ObjectId);
            return a.LineNo.CompareTo(b.LineNo);
        }

        public int RowCount { get { lock (_sync) { return _snapshot.Count; } } }

        /// <summary>Готовит выборку запросов SQL и возвращает её размер. Отрицательное
        /// значение отбора означает «не фильтровать»: код объекта 0 и строка 0 существуют.
        /// Номер строки - как его прислал ETW, без смещения на листинг.</summary>
        public int SelectQueries(int objectType, int objectId, int lineNo, string functionName)
        {
            lock (_sync)
            {
                List<SqlStat> list = new List<SqlStat>();
                foreach (KeyValuePair<long, Dictionary<string, SqlStat>> kv in _sqlByLine)
                {
                    foreach (SqlStat q in kv.Value.Values)
                    {
                        if (objectId >= 0 && (q.ObjectId != objectId || q.ObjectType != objectType)) continue;
                        if (lineNo >= 0 && q.LineNo != lineNo) continue;
                        if (!string.IsNullOrEmpty(functionName) &&
                            !string.Equals(q.FunctionName, functionName, StringComparison.OrdinalIgnoreCase)) continue;
                        list.Add(q);
                    }
                }
                list.Sort(CompareSql);
                _querySnapshot = list;
                _queryCursor = -1;
                return list.Count;
            }
        }

        /// <summary>Отбирает рёбра вызова. Отрицательный код объекта и пустое имя
        /// означают «не сужать». Номер строки - в нумерации ETW, как его прислал провайдер.</summary>
        /// <summary>Отбирает рёбра, выходящие из ФУНКЦИИ, а не из строки. Нужно там, где
        /// вызванная функция не имеет ни одной строки в листинге и потому не может быть целью
        /// проваливания: так устроена пустая функция-издатель события. Подписчик висит на её
        /// операторе, и добраться до него можно только через её имя.</summary>
        public int SelectEdgesFromFunction(int callerObjectType, int callerObjectId, string callerFunctionName)
        {
            lock (_sync)
            {
                List<EdgeStat> list = new List<EdgeStat>();
                foreach (KeyValuePair<long, Dictionary<CalleeKey, EdgeStat>> kv in _edges)
                    foreach (KeyValuePair<CalleeKey, EdgeStat> ev in kv.Value)
                    {
                        EdgeStat e = ev.Value;
                        if (e.CallerObjectId != callerObjectId || e.CallerObjectType != callerObjectType) continue;
                        if (!string.Equals(e.CallerFunc, callerFunctionName, StringComparison.OrdinalIgnoreCase)) continue;
                        list.Add(e);
                    }
                list.Sort(CompareEdge);
                _edgeSnapshot = list;
                _edgeCursor = -1;
                return list.Count;
            }
        }

        public int SelectEdges(int callerObjectType, int callerObjectId, int lineNo, string calleeFunctionName)
        {
            lock (_sync)
            {
                // Сперва считаем вызовы каждой функции ОТОВСЮДУ - это то самое «из M»,
                // ради которого всё и затевалось: без него не отличить единственного
                // вызывателя от одного из многих.
                Dictionary<CalleeKey, long> totals = new Dictionary<CalleeKey, long>();
                foreach (KeyValuePair<long, Dictionary<CalleeKey, EdgeStat>> kv in _edges)
                    foreach (KeyValuePair<CalleeKey, EdgeStat> ev in kv.Value)
                    {
                        long had;
                        totals.TryGetValue(ev.Key, out had);
                        totals[ev.Key] = had + ev.Value.Calls;
                    }

                List<EdgeStat> list = new List<EdgeStat>();
                foreach (KeyValuePair<long, Dictionary<CalleeKey, EdgeStat>> kv in _edges)
                    foreach (KeyValuePair<CalleeKey, EdgeStat> ev in kv.Value)
                    {
                        EdgeStat e = ev.Value;
                        if (callerObjectId >= 0 &&
                            (e.CallerObjectId != callerObjectId || e.CallerObjectType != callerObjectType)) continue;
                        if (lineNo >= 0 && e.CallerLineNo != lineNo) continue;
                        if (!string.IsNullOrEmpty(calleeFunctionName) &&
                            !string.Equals(e.CalleeFunc, calleeFunctionName, StringComparison.OrdinalIgnoreCase)) continue;
                        // Итог по функции держим в самой записи: заводить параллельный тип
                        // строки снимка ради одного числа не стоит, а пересчитывается он
                        // на каждом отборе заново.
                        e.CalleeTotalCalls = totals[ev.Key];
                        list.Add(e);
                    }
                list.Sort(CompareEdge);
                _edgeSnapshot = list;
                _edgeCursor = -1;
                return list.Count;
            }
        }

        private static int CompareEdge(EdgeStat a, EdgeStat b)
        {
            int c = b.TotalMs.CompareTo(a.TotalMs);
            if (c != 0) return c;
            return b.Calls.CompareTo(a.Calls);
        }

        /// <summary>Ставит курсор на ребро выборки (нумерация с нуля).</summary>
        public bool SelectEdgeRow(int index)
        {
            lock (_sync)
            {
                if (index < 0 || index >= _edgeSnapshot.Count) return false;
                _edgeCursor = index;
                return true;
            }
        }

        private EdgeStat CurEdge
        {
            get
            {
                lock (_sync)
                {
                    if (_edgeCursor < 0 || _edgeCursor >= _edgeSnapshot.Count) return null;
                    return _edgeSnapshot[_edgeCursor];
                }
            }
        }

        public int CurrentEdgeCallerObjectType { get { EdgeStat e = CurEdge; return e == null ? 0 : e.CallerObjectType; } }
        public int CurrentEdgeCallerObjectId { get { EdgeStat e = CurEdge; return e == null ? 0 : e.CallerObjectId; } }
        public int CurrentEdgeCallerLineNo { get { EdgeStat e = CurEdge; return e == null ? 0 : e.CallerLineNo; } }
        public string CurrentEdgeCallerFunctionName { get { EdgeStat e = CurEdge; return e == null ? "" : (e.CallerFunc ?? ""); } }
        public int CurrentEdgeCalleeObjectType { get { EdgeStat e = CurEdge; return e == null ? 0 : e.CalleeObjectType; } }
        public int CurrentEdgeCalleeObjectId { get { EdgeStat e = CurEdge; return e == null ? 0 : e.CalleeObjectId; } }
        public string CurrentEdgeFunctionName { get { EdgeStat e = CurEdge; return e == null || e.CalleeFunc == null ? "" : e.CalleeFunc; } }
        public int CurrentEdgeCalls { get { EdgeStat e = CurEdge; return e == null ? 0 : (int)e.Calls; } }

        /// <summary>Вызовов этой функции ОТОВСЮДУ. Равно CurrentEdgeCalls - значит из этой
        /// строки её и зовут единственно.</summary>
        public int CurrentEdgeCalleeTotalCalls { get { EdgeStat e = CurEdge; return e == null ? 0 : (int)e.CalleeTotalCalls; } }

        public bool CurrentEdgeRecursive { get { EdgeStat e = CurEdge; return e != null && e.Recursive; } }

        public decimal CurrentEdgeTotalMs { get { return EdgeMs(CurEdge, 1); } }
        public decimal CurrentEdgeSelfMs { get { return EdgeMs(CurEdge, 2); } }
        public decimal CurrentEdgeMaxMs { get { return EdgeMs(CurEdge, 3); } }

        private static decimal EdgeMs(EdgeStat e, int what)
        {
            if (e == null) return 0m;
            double v;
            switch (what)
            {
                case 1: v = e.TotalMs; break;
                case 2: v = e.SelfMs; break;
                default: v = e.MaxMs; break;
            }
            if (double.IsNaN(v) || double.IsInfinity(v) || v < 0) return 0m;
            return Math.Round((decimal)v, 3);
        }

        private static int CompareSql(SqlStat a, SqlStat b)
        {
            int c = b.TotalMs.CompareTo(a.TotalMs);
            if (c != 0) return c;
            return b.Executions.CompareTo(a.Executions);
        }

        /// <summary>Ставит курсор на запрос выборки (нумерация с нуля).</summary>
        public bool SelectQueryRow(int index)
        {
            lock (_sync)
            {
                if (index < 0 || index >= _querySnapshot.Count) return false;
                _queryCursor = index;
                return true;
            }
        }

        private SqlStat CurQuery
        {
            get
            {
                lock (_sync)
                {
                    if (_queryCursor < 0 || _queryCursor >= _querySnapshot.Count) return null;
                    return _querySnapshot[_queryCursor];
                }
            }
        }

        public int CurrentQueryObjectType { get { SqlStat q = CurQuery; return q == null ? 0 : q.ObjectType; } }
        public int CurrentQueryObjectId { get { SqlStat q = CurQuery; return q == null ? 0 : q.ObjectId; } }
        public int CurrentQueryLineNo { get { SqlStat q = CurQuery; return q == null ? 0 : q.LineNo; } }
        public int CurrentQueryExecutions { get { SqlStat q = CurQuery; return q == null ? 0 : (int)q.Executions; } }
        public bool CurrentQueryTruncated { get { SqlStat q = CurQuery; return q != null && q.Truncated; } }

        public string CurrentQueryFunctionName
        {
            get { SqlStat q = CurQuery; return q == null || q.FunctionName == null ? "" : q.FunctionName; }
        }

        public string CurrentQueryText
        {
            get { SqlStat q = CurQuery; return q == null || q.Statement == null ? "" : q.Statement; }
        }

        public decimal CurrentQueryTotalMs { get { return SqlMs(CurQuery, true); } }
        public decimal CurrentQueryMaxMs { get { return SqlMs(CurQuery, false); } }

        private static decimal SqlMs(SqlStat q, bool total)
        {
            if (q == null) return 0m;
            double v = total ? q.TotalMs : q.MaxMs;
            if (double.IsNaN(v) || double.IsInfinity(v) || v < 0) return 0m;
            return Math.Round((decimal)v, 3);
        }

        /// <summary>Ставит курсор на строку (нумерация с нуля).</summary>
        public bool SelectRow(int index)
        {
            lock (_sync)
            {
                if (index < 0 || index >= _snapshot.Count) return false;
                _cursor = index;
                return true;
            }
        }

        private LineStat Cur
        {
            get
            {
                lock (_sync)
                {
                    if (_cursor < 0 || _cursor >= _snapshot.Count) return null;
                    return _snapshot[_cursor];
                }
            }
        }

        public int CurrentObjectType { get { LineStat s = Cur; return s == null ? 0 : s.ObjectType; } }
        public int CurrentObjectId { get { LineStat s = Cur; return s == null ? 0 : s.ObjectId; } }
        public int CurrentLineNo { get { LineStat s = Cur; return s == null ? 0 : s.LineNo; } }
        public string CurrentFunctionName { get { LineStat s = Cur; return s == null ? "" : (s.FunctionName == null ? "" : s.FunctionName); } }
        public int CurrentHits { get { LineStat s = Cur; return s == null ? 0 : (int)s.Hits; } }
        public string CurrentLineText { get { LineStat s = Cur; return s == null || s.Text == null ? "" : s.Text; } }
        public int CurrentSqlCount { get { LineStat s = Cur; return s == null ? 0 : s.SqlCount; } }

        // Время отдаётся в decimal: у C/AL это родной Decimal, double он принимает хуже.
        // ВАЖНО: CurrentSqlMs - ЧАСТЬ CurrentSelfMs, а не добавка к нему. Складывать нельзя.
        public decimal CurrentTotalMs { get { return Ms(Cur, 1); } }
        public decimal CurrentSelfMs { get { return Ms(Cur, 2); } }
        public decimal CurrentSqlMs { get { return Ms(Cur, 3); } }
        public decimal CurrentMinMs { get { return Ms(Cur, 4); } }
        public decimal CurrentMaxMs { get { return Ms(Cur, 5); } }

        private static decimal Ms(LineStat s, int what)
        {
            if (s == null) return 0m;
            double v;
            switch (what)
            {
                case 1: v = s.TotalMs; break;
                case 2: v = s.SelfMs; break;
                case 3: v = s.SqlMs; break;
                case 4: v = s.MinMs; break;
                default: v = s.MaxMs; break;
            }
            if (double.IsNaN(v) || double.IsInfinity(v) || v < 0) return 0m;
            return Math.Round((decimal)v, 3);
        }

        // ---------------------------------------------------------------- мелочи

        private bool Interesting(int objType, int objId)
        {
            if (_targetObjectId < 0) return true;
            return objType == _targetObjectType && objId == _targetObjectId;
        }

        private static long LineKey(int objType, int objId, int line)
        {
            return ((long)objType << 56) ^ ((long)objId << 24) ^ (uint)line;
        }

        private LineStat GetStat(int objType, int objId, int line, string func)
        {
            long key = LineKey(objType, objId, line);
            LineStat s;
            if (!_stats.TryGetValue(key, out s))
            {
                s = new LineStat();
                s.ObjectType = objType;
                s.ObjectId = objId;
                s.LineNo = line;
                s.FunctionName = func;
                _stats[key] = s;
            }
            else if (string.IsNullOrEmpty(s.FunctionName)) s.FunctionName = func;
            return s;
        }

        private SessionState GetSession(int sid)
        {
            SessionState st;
            if (!_sessions.TryGetValue(sid, out st))
            {
                st = new SessionState();
                _sessions[sid] = st;
            }
            return st;
        }

        private static Frame Top(SessionState st)
        {
            return st.Stack.Count == 0 ? null : st.Stack[st.Stack.Count - 1];
        }

        private static object SafePayload(TraceEvent data, string name)
        {
            try { return data.PayloadByName(name); }
            catch { return null; }
        }

        private static int ToInt(object v, int dflt)
        {
            if (v == null) return dflt;
            try { return Convert.ToInt32(v); }
            catch { }
            int n;
            if (int.TryParse(Convert.ToString(v), out n)) return n;
            return dflt;
        }

        private static string ToStr(object v)
        {
            return v == null ? "" : Convert.ToString(v);
        }

        /// <summary>objectType приходит то числом, то именем типа - принимаем оба вида.</summary>
        private static int ToObjType(object v)
        {
            if (v == null) return -1;
            int n;
            string s = Convert.ToString(v);
            if (int.TryParse(s, out n)) return n;
            if (ObjTypeByName.TryGetValue(s, out n)) return n;
            return -1;
        }
    }
}