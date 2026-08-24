using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace SusScripting_1._0;

internal sealed class Runner : IDisposable
{
    private readonly string _dllPath;
    private readonly string _language;
    private readonly string _code;

    private IntPtr _dll;
    private IActiveScript? _script;
    private IActiveScriptParse32? _parse;
    private ScriptSite? _site;

    private bool _initialized;
    private bool _disposed;
    private object? _dispatch;
    private WScriptHost? _wscript;
    internal WScriptHost? WScript =>
    _wscript;

    public string? LastError { get; private set; }

    public Runner(
        string dllPath,
        string language,
        string code)
    {
        _dllPath = dllPath;
        _language = language;
        _code = code;
    }

    public bool Initialize()
    {
        try
        {
            if (!OperatingSystem.IsWindows())
            {
                throw new PlatformNotSupportedException(
                    "Active Scripting requires Windows.");
            }

            /*
             * Этот Runner рассчитан на x86.
             */
            if (IntPtr.Size != 4)
            {
                throw new PlatformNotSupportedException(
                    "This Runner is configured for x86. " +
                    "Build the project with PlatformTarget=x86.");
            }

            LastError = null;

            Guid clsid = GetClsid(_language);

            /*
             * Загружаем DLL явно.
             */
            _dll = Native.LoadLibrary(_dllPath);

            if (_dll == IntPtr.Zero)
            {
                throw new Win32Exception(
                    Marshal.GetLastWin32Error(),
                    $"LoadLibrary failed: {_dllPath}");
            }

            /*
             * Получаем IActiveScript.
             */
            _script =
                CreateComObject<IActiveScript>(
                    _dll,
                    clsid);

            /*
             * IActiveScriptParse32 является интерфейсом
             * ТОГО ЖЕ COM-объекта.
             *
             * Нельзя делать второй CreateInstance().
             */
            _parse =
                GetInterface<IActiveScriptParse32>(
                    _script);

            /*
             * Очень важно:
             * ScriptSite должен жить всё время жизни engine.
             */
            _site =
                new ScriptSite(this);

            /*
             * Передаём ActiveScriptSite движку.
             */
            int hr =
                _script.SetScriptSite(_site);

            CheckHR(
                hr,
                "IActiveScript.SetScriptSite");

            /*
             * Инициализируем parser.
             */
            hr =
                _parse.InitNew();

            CheckHR(
                hr,
                "IActiveScriptParse32.InitNew");

            _wscript =
    new WScriptHost();

            hr =
                _script.AddNamedItem(
                    "WScript",
                    ScriptItemFlags.IsVisible);

            CheckHR(
                hr,
                "IActiveScript.AddNamedItem");

            /*
             * Загружаем основной код.
             */
            IntPtr variant =
                Marshal.AllocCoTaskMem(
                    16);

            try
            {
                ClearVariant(
                    variant);

                hr =
                    _parse.ParseScriptText(
                        _code,
                        null,
                        IntPtr.Zero,
                        null,
                        0,
                        0,
                        ScriptTextFlags.HostManagesSource,
                        variant,
                        IntPtr.Zero);

                if (hr < 0)
                {
                    if (!string.IsNullOrWhiteSpace(
                            LastError))
                    {
                        throw new ActiveScriptException(
                            LastError!,
                            hr);
                    }

                    throw new COMException(
                        $"ParseScriptText failed. " +
                        $"HRESULT=0x{hr:X8}",
                        hr);
                }
            }
            finally
            {
                Marshal.FreeCoTaskMem(
                    variant);
            }

            /*
             * Script engine мог сообщить ошибку через
             * IActiveScriptSite.OnScriptError().
             */
            if (!string.IsNullOrWhiteSpace(
                    LastError))
            {
                throw new ActiveScriptException(
                    LastError!,
                    HResult.E_FAIL);
            }

            /*
             * Переводим engine в Connected.
             */
            hr =
                _script.SetScriptState(
                    ScriptState.Connected);

            CheckHR(
                hr,
                "IActiveScript.SetScriptState");

            _initialized = true;

            return true;
        }
        catch (Exception ex)
        {
            SetError(ex);

            Dispose();

            return false;
        }
    }

    public int Execute(
        int a,
        int b)
    {
        if (!_initialized ||
            _parse == null)
        {
            throw new InvalidOperationException(
                "Runner is not initialized.");
        }

        LastError = null;

        string invocation =
            BuildInvocation(a, b);

        IntPtr variant =
            Marshal.AllocCoTaskMem(
                16);

        try
        {
            ClearVariant(
                variant);

            int hr =
                _parse.ParseScriptText(
                    invocation,
                    null,
                    IntPtr.Zero,
                    null,
                    0,
                    0,
                    ScriptTextFlags.IsExpression |
                    ScriptTextFlags.HostManagesSource,
                    variant,
                    IntPtr.Zero);

            if (hr < 0)
            {
                if (!string.IsNullOrWhiteSpace(
                        LastError))
                {
                    throw new ActiveScriptException(
                        LastError!,
                        hr);
                }

                throw new COMException(
                    $"ParseScriptText failed. " +
                    $"HRESULT=0x{hr:X8}",
                    hr);
            }

            if (!string.IsNullOrWhiteSpace(
                    LastError))
            {
                throw new ActiveScriptException(
                    LastError!,
                    HResult.E_FAIL);
            }

            object? value =
                Marshal.GetObjectForNativeVariant(
                    variant);

            if (value == null)
            {
                throw new InvalidOperationException(
                    "ActiveScript returned NULL.");
            }

            return Convert.ToInt32(
                value,
                CultureInfo.InvariantCulture);
        }
        finally
        {
            Marshal.FreeCoTaskMem(
                variant);
        }
    }

    private static string BuildInvocation(
    int a,
    int b)
    {
        return
            "Execute(" +
            a.ToString(CultureInfo.InvariantCulture) +
            ", " +
            b.ToString(CultureInfo.InvariantCulture) +
            ")";
    }

    private static T CreateComObject<T>(
        IntPtr dll,
        Guid clsid)
        where T : class
    {
        IntPtr proc =
            Native.GetProcAddress(
                dll,
                "DllGetClassObject");

        if (proc == IntPtr.Zero)
        {
            throw new Win32Exception(
                Marshal.GetLastWin32Error(),
                "DllGetClassObject not found.");
        }

        var getClassObject =
            Marshal.GetDelegateForFunctionPointer<
                Native.DllGetClassObjectDelegate>(
                proc);

        Guid iidFactory =
            typeof(IClassFactory).GUID;

        int hr =
            getClassObject(
                ref clsid,
                ref iidFactory,
                out IntPtr factoryPtr);

        if (hr < 0)
        {
            throw new COMException(
                $"DllGetClassObject failed. " +
                $"CLSID={clsid}, " +
                $"HRESULT=0x{hr:X8}",
                hr);
        }

        if (factoryPtr == IntPtr.Zero)
        {
            throw new COMException(
                "DllGetClassObject returned NULL.",
                HResult.E_POINTER);
        }

        try
        {
            object factoryObject =
                Marshal.GetObjectForIUnknown(
                    factoryPtr);

            if (factoryObject is not IClassFactory factory)
            {
                throw new COMException(
                    "IClassFactory cast failed.",
                    HResult.E_NOINTERFACE);
            }

            Guid iid =
                typeof(T).GUID;

            hr =
                factory.CreateInstance(
                    IntPtr.Zero,
                    ref iid,
                    out IntPtr objectPtr);

            if (hr < 0)
            {
                throw new COMException(
                    $"IClassFactory.CreateInstance failed. " +
                    $"IID={iid}, " +
                    $"HRESULT=0x{hr:X8}",
                    hr);
            }

            if (objectPtr == IntPtr.Zero)
            {
                throw new COMException(
                    "CreateInstance returned NULL.",
                    HResult.E_POINTER);
            }

            try
            {
                object obj =
                    Marshal.GetObjectForIUnknown(
                        objectPtr);

                if (obj is T result)
                {
                    return result;
                }

                /*
                 * Дополнительный QI.
                 */
                IntPtr queriedPtr =
                    IntPtr.Zero;

                try
                {
                    Guid requestedIid =
                        typeof(T).GUID;

                    hr =
                        Marshal.QueryInterface(
                            objectPtr,
                            ref requestedIid,
                            out queriedPtr);

                    if (hr < 0)
                    {
                        throw new COMException(
                            $"QueryInterface failed. " +
                            $"IID={requestedIid}, " +
                            $"HRESULT=0x{hr:X8}",
                            hr);
                    }

                    object queriedObject =
                        Marshal.GetObjectForIUnknown(
                            queriedPtr);

                    if (queriedObject is T typed)
                    {
                        return typed;
                    }

                    throw new COMException(
                        $"Interface {typeof(T).Name} " +
                        "is not supported.",
                        HResult.E_NOINTERFACE);
                }
                finally
                {
                    if (queriedPtr != IntPtr.Zero)
                    {
                        Marshal.Release(
                            queriedPtr);
                    }
                }
            }
            finally
            {
                Marshal.Release(
                    objectPtr);
            }
        }
        finally
        {
            Marshal.Release(
                factoryPtr);
        }
    }

    private static T GetInterface<T>(
        object source)
        where T : class
    {
        IntPtr unknown =
            Marshal.GetIUnknownForObject(
                source);

        try
        {
            Guid iid =
                typeof(T).GUID;

            int hr =
                Marshal.QueryInterface(
                    unknown,
                    ref iid,
                    out IntPtr interfacePtr);

            if (hr < 0)
            {
                throw new COMException(
                    $"QueryInterface failed. " +
                    $"IID={iid}, " +
                    $"HRESULT=0x{hr:X8}",
                    hr);
            }

            if (interfacePtr == IntPtr.Zero)
            {
                throw new COMException(
                    "QueryInterface returned NULL.",
                    HResult.E_POINTER);
            }

            try
            {
                object obj =
                    Marshal.GetObjectForIUnknown(
                        interfacePtr);

                if (obj is T result)
                {
                    return result;
                }

                throw new COMException(
                    $"Interface {typeof(T).Name} " +
                    "could not be obtained.",
                    HResult.E_NOINTERFACE);
            }
            finally
            {
                Marshal.Release(
                    interfacePtr);
            }
        }
        finally
        {
            Marshal.Release(
                unknown);
        }
    }

    private static Guid GetClsid(
        string language)
    {
        if (language.Equals(
                "VBScript",
                StringComparison.OrdinalIgnoreCase))
        {
            return new Guid(
                "B54F3741-5B07-11CF-A4B0-00AA004A55E8");
        }

        if (language.Equals(
                "JScript",
                StringComparison.OrdinalIgnoreCase))
        {
            return new Guid(
                "F414C260-6AC0-11CF-B6D1-00AA00BBBB58");
        }

        if (Guid.TryParse(
                language,
                out Guid guid))
        {
            return guid;
        }

        throw new NotSupportedException(
            $"Unknown Active Script engine: {language}");
    }

    internal void SetError(
        Exception ex)
    {
        LastError =
            ex.ToString();
    }

    internal void SetError(
        string error)
    {
        LastError =
            error;
    }

    private static void CheckHR(
        int hr,
        string operation)
    {
        if (hr < 0)
        {
            throw new COMException(
                $"{operation} failed. " +
                $"HRESULT=0x{hr:X8}",
                hr);
        }
    }

    private static void ClearVariant(
        IntPtr ptr)
    {
        for (int i = 0; i < 16; i++)
        {
            Marshal.WriteByte(
                ptr,
                i,
                0);
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        try
        {
            if (_script != null)
            {
                try
                {
                    _script.SetScriptState(
                        ScriptState.Closed);
                }
                catch
                {
                }

                try
                {
                    _script.Close();
                }
                catch
                {
                }
            }
        }
        finally
        {
            _parse = null;
            _script = null;
            _site = null;

            if (_dll != IntPtr.Zero)
            {
                Native.FreeLibrary(
                    _dll);

                _dll = IntPtr.Zero;
            }

            _initialized = false;
        }

        GC.SuppressFinalize(this);
    }

    ~Runner()
    {
        Dispose();
    }
}


internal sealed class ActiveScriptException :
    Exception
{
    public int ErrorHResult { get; }

    public ActiveScriptException(
        string message,
        int hresult)
        : base(message)
    {
        ErrorHResult = hresult;
    }
}


/*
 * Active Script Site
 */
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
internal sealed class ScriptSite :
    IActiveScriptSite
{
    private readonly Runner _runner;

    public ScriptSite(
        Runner runner)
    {
        _runner = runner;
    }

    public int GetLCID(
        out uint plcid)
    {
        plcid = 0x0409;

        return HResult.S_OK;
    }

    /*
     * ВАЖНО:
     *
     * Native signature:
     *
     * HRESULT GetItemInfo(
     *     LPCOLESTR pstrName,
     *     DWORD dwReturnMask,
     *     IUnknown **ppiunkItem,
     *     ITypeInfo **ppti
     * );
     *
     * Поэтому здесь используются IntPtr,
     * а НЕ out object.
     */
    public int GetItemInfo(
    string pstrName,
    uint dwReturnMask,
    out IntPtr ppiunkItem,
    IntPtr ppti)
    {
        ppiunkItem = IntPtr.Zero;

        const uint SCRIPTINFO_IUNKNOWN =
            0x00000001;

        const uint SCRIPTINFO_ITYPEINFO =
            0x00000002;

        if (!string.Equals(
                pstrName,
                "WScript",
                StringComparison.OrdinalIgnoreCase))
        {
            return HResult.TYPE_E_ELEMENTNOTFOUND;
        }

        if (_runner.WScript == null)
        {
            return HResult.E_FAIL;
        }

        /*
         * ActiveScript просит IUnknown.
         */
        if ((dwReturnMask &
             SCRIPTINFO_IUNKNOWN) != 0)
        {
            ppiunkItem =
                Marshal.GetIUnknownForObject(
                    _runner.WScript);

            if (ppiunkItem == IntPtr.Zero)
            {
                return HResult.E_POINTER;
            }
        }

        /*
         * TypeInfo нам для этого теста не нужен.
         *
         * Если engine запросит только ITypeInfo,
         * оставляем ppti = NULL.
         */
        if ((dwReturnMask &
             SCRIPTINFO_ITYPEINFO) != 0)
        {
            /*
             * ppti уже передан как IntPtr,
             * поэтому здесь ничего делать не нужно.
             */
        }

        return HResult.S_OK;
    }
    public int GetDocVersionString(
        out string pbstrVersion)
    {
        pbstrVersion = "1.0";

        return HResult.S_OK;
    }

    public int OnScriptTerminate(
        IntPtr pvarResult,
        IntPtr pexcepinfo)
    {
        return HResult.S_OK;
    }

    public int OnStateChange(
        ScriptState ssScriptState)
    {
        return HResult.S_OK;
    }

    public int OnScriptError(
        IActiveScriptError pscripterror)
    {
        try
        {
            int hr =
                pscripterror.GetExceptionInfo(
                    out EXCEPINFO info);

            if (hr < 0)
            {
                _runner.SetError(
                    $"GetExceptionInfo failed. " +
                    $"HRESULT=0x{hr:X8}");

                return HResult.S_OK;
            }

            pscripterror.GetSourcePosition(
                out _,
                out uint line,
                out int character);

            string? sourceLine = null;

            try
            {
                pscripterror.GetSourceLineText(
                    out sourceLine);
            }
            catch
            {
            }

            int errorCode =
                info.scode != 0
                    ? info.scode
                    : info.wCode;

            var sb =
                new StringBuilder();

            sb.AppendLine(
                $"HRESULT: 0x{errorCode:X8}");

            if (!string.IsNullOrEmpty(
                    info.bstrSource))
            {
                sb.AppendLine(
                    $"Source: {info.bstrSource}");
            }

            if (!string.IsNullOrEmpty(
                    info.bstrDescription))
            {
                sb.AppendLine(
                    $"Description: {info.bstrDescription}");
            }

            sb.AppendLine(
                $"Line: {line + 1}");

            sb.AppendLine(
                $"Character: {character}");

            if (!string.IsNullOrEmpty(
                    sourceLine))
            {
                sb.AppendLine(
                    $"Code: {sourceLine}");
            }

            _runner.SetError(
                sb.ToString());
        }
        catch (Exception ex)
        {
            _runner.SetError(ex);
        }

        /*
         * Active Script ожидает S_OK,
         * даже если мы сохранили ошибку.
         */
        return HResult.S_OK;
    }

    public int OnEnterScript()
    {
        return HResult.S_OK;
    }

    public int OnLeaveScript()
    {
        return HResult.S_OK;
    }
}


/*
 * IActiveScript
 */
[ComImport]
[Guid(
    "BB1A2AE1-A4F9-11CF-8F20-00805F2CD064")]
[InterfaceType(
    ComInterfaceType.InterfaceIsIUnknown)]
internal interface IActiveScript
{
    [PreserveSig]
    int SetScriptSite(
        IActiveScriptSite pass);

    [PreserveSig]
    int GetScriptSite(
        ref Guid riid,
        out object ppvObject);

    [PreserveSig]
    int SetScriptState(
        ScriptState ss);

    [PreserveSig]
    int GetScriptState(
        out ScriptState pssState);

    [PreserveSig]
    int Close();

    [PreserveSig]
    int AddNamedItem(
        string pstrName,
        ScriptItemFlags dwFlags);

    [PreserveSig]
    int AddTypeLib(
        ref Guid rguidTypeLib,
        uint dwMajor,
        uint dwMinor,
        uint dwFlags);

    [PreserveSig]
    int GetScriptDispatch(
        string pstrItemName,
        out object ppdisp);

    [PreserveSig]
    int GetCurrentScriptThreadID(
        out uint pstidThread);

    [PreserveSig]
    int GetScriptThreadID(
        uint dwWin32ThreadId,
        out uint pdwScriptThreadID);

    [PreserveSig]
    int GetScriptThreadState(
        uint stidThread,
        out ScriptThreadState pstsState);

    [PreserveSig]
    int InterruptScriptThread(
        uint stidThread,
        IntPtr pexcepinfo,
        uint dwFlags);

    [PreserveSig]
    int Clone(
        out IActiveScript ppscript);
}


/*
 * IActiveScriptParse32
 *
 * Именно этот интерфейс используется x86.
 */
[ComImport]
[Guid(
    "BB1A2AE2-A4F9-11CF-8F20-00805F2CD064")]
[InterfaceType(
    ComInterfaceType.InterfaceIsIUnknown)]
internal interface IActiveScriptParse32
{
    [PreserveSig]
    int InitNew();

    [PreserveSig]
    int AddScriptlet(
        string pstrDefaultName,
        string pstrCode,
        string pstrItemName,
        string pstrSubItemName,
        string pstrEventName,
        string pstrDelimiter,
        uint dwSourceContextCookie,
        uint ulStartingLineNumber,
        uint dwFlags,
        out string pbstrName,
        IntPtr pexcepinfo);

    [PreserveSig]
    int ParseScriptText(
        string pstrCode,
        string? pstrItemName,
        IntPtr punkContext,
        string? pstrDelimiter,
        uint dwSourceContextCookie,
        uint ulStartingLineNumber,
        ScriptTextFlags dwFlags,
        IntPtr pvarResult,
        IntPtr pexcepinfo);
}


/*
 * IActiveScriptSite
 */
[ComImport]
[Guid(
    "DB01A1E3-A42B-11CF-8F20-00805F2CD064")]
[InterfaceType(
    ComInterfaceType.InterfaceIsIUnknown)]
internal interface IActiveScriptSite
{
    [PreserveSig]
    int GetLCID(
        out uint plcid);

    [PreserveSig]
    int GetItemInfo(
        string pstrName,
        uint dwReturnMask,
        out IntPtr ppiunkItem,
        IntPtr ppti);

    [PreserveSig]
    int GetDocVersionString(
        [MarshalAs(UnmanagedType.BStr)]
        out string pbstrVersion);

    [PreserveSig]
    int OnScriptTerminate(
        IntPtr pvarResult,
        IntPtr pexcepinfo);

    [PreserveSig]
    int OnStateChange(
        ScriptState ssScriptState);

    [PreserveSig]
    int OnScriptError(
        IActiveScriptError pscripterror);

    [PreserveSig]
    int OnEnterScript();

    [PreserveSig]
    int OnLeaveScript();
}


/*
 * IActiveScriptError
 */
[ComImport]
[Guid(
    "EAE1BA61-A4ED-11CF-8F20-00805F2CD064")]
[InterfaceType(
    ComInterfaceType.InterfaceIsIUnknown)]
internal interface IActiveScriptError
{
    [PreserveSig]
    int GetExceptionInfo(
        out EXCEPINFO pexcepinfo);

    [PreserveSig]
    int GetSourcePosition(
        out uint pdwSourceContext,
        out uint pulLineNumber,
        out int plCharacterPosition);

    [PreserveSig]
    int GetSourceLineText(
        [MarshalAs(UnmanagedType.BStr)]
        out string pbstrSourceLine);
}


/*
 * IClassFactory
 */
[ComImport]
[Guid(
    "00000001-0000-0000-C000-000000000046")]
[InterfaceType(
    ComInterfaceType.InterfaceIsIUnknown)]
internal interface IClassFactory
{
    [PreserveSig]
    int CreateInstance(
        IntPtr pUnkOuter,
        ref Guid riid,
        out IntPtr ppvObject);

    [PreserveSig]
    int LockServer(
        [MarshalAs(UnmanagedType.Bool)]
        bool fLock);
}


/*
 * Native Win32
 */
internal static class Native
{
    [DllImport(
        "kernel32.dll",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    internal static extern IntPtr LoadLibrary(
        string lpFileName);

    [DllImport(
        "kernel32.dll",
        SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool FreeLibrary(
        IntPtr hModule);

    [DllImport(
        "kernel32.dll",
        CharSet = CharSet.Ansi,
        SetLastError = true)]
    internal static extern IntPtr GetProcAddress(
        IntPtr hModule,
        string lpProcName);

    [UnmanagedFunctionPointer(
        CallingConvention.StdCall)]
    internal delegate int DllGetClassObjectDelegate(
        ref Guid rclsid,
        ref Guid riid,
        out IntPtr ppv);
}


/*
 * EXCEPINFO
 */
[StructLayout(
    LayoutKind.Sequential)]
internal struct EXCEPINFO
{
    public ushort wCode;

    public ushort wReserved;

    [MarshalAs(
        UnmanagedType.BStr)]
    public string? bstrSource;

    [MarshalAs(
        UnmanagedType.BStr)]
    public string? bstrDescription;

    [MarshalAs(
        UnmanagedType.BStr)]
    public string? bstrHelpFile;

    public uint dwHelpContext;

    public IntPtr pvReserved;

    public IntPtr pfnDeferredFillIn;

    public int scode;
}


/*
 * Script flags
 */
[Flags]
internal enum ScriptTextFlags : uint
{
    None = 0x00000000,

    IsExpression = 0x00000020,

    IsPersistent = 0x00000040,

    HostManagesSource = 0x00000080,

    IsVisible = 0x00000002,

    IgnoreScriptError = 0x00000001
}


/*
 * Script item flags
 */
[Flags]
internal enum ScriptItemFlags : uint
{
    IsVisible = 0x00000002,

    IsSource = 0x00000004,

    GlobalMembers = 0x00000008,

    IsPersistent = 0x00000040,

    CodeOnly = 0x00000200,

    NoCode = 0x00000400
}


/*
 * Script state
 */
internal enum ScriptState : uint
{
    Uninitialized = 0,

    Started = 1,

    Connected = 2,

    Disconnected = 3,

    Closed = 4,

    Initialized = 5
}


/*
 * Script thread state
 */
internal enum ScriptThreadState : uint
{
    NotInScript = 0,

    Running = 1
}


/*
 * HRESULT
 */
internal static class HResult
{
    public const int S_OK = 0;

    public const int E_FAIL =
        unchecked(
            (int)0x80004005);

    public const int E_POINTER =
        unchecked(
            (int)0x80004003);

    public const int E_NOINTERFACE =
        unchecked(
            (int)0x80004002);

    public const int E_HANDLE =
        unchecked(
            (int)0x80070006);

    public const int TYPE_E_ELEMENTNOTFOUND =
        unchecked(
            (int)0x8002802B);
}