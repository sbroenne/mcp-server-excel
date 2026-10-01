using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacOsaScriptRuntime
{
    private const string ObjectiveC = "/usr/lib/libobjc.A.dylib";

    public static string Execute(string source, string command, string arguments)
    {
        if (!OperatingSystem.IsMacOS())
        {
            throw new PlatformNotSupportedException("OSAKit requires macOS.");
        }

        var invocation = JsonSerializer.Serialize(new[] { command, arguments });
        var callableSource = source.Replace(
            "function run(argv)",
            "function excelMcpRun(argv)",
            StringComparison.Ordinal);
        if (ReferenceEquals(callableSource, source))
        {
            throw new InvalidOperationException("Embedded Excel automation entry point was not found.");
        }
        return ExecuteSource($"{callableSource}\nexcelMcpRun({invocation});", "JavaScript");
    }

    public static string ExecuteAppleScript(string source) => ExecuteSource(source, "AppleScript");

    private static string ExecuteSource(string source, string languageNameValue)
    {
        NativeLibrary.Load("/System/Library/Frameworks/Foundation.framework/Foundation");
        NativeLibrary.Load("/System/Library/Frameworks/OSAKit.framework/OSAKit");
        var pool = Send(GetClass("NSAutoreleasePool"), GetSelector("alloc"));
        pool = Send(pool, GetSelector("init"));
        try
        {
            var nsString = GetClass("NSString");
            var languageName = SendString(
                nsString,
                GetSelector("stringWithUTF8String:"),
                languageNameValue);
            var language = Send(
                GetClass("OSALanguage"),
                GetSelector("languageForName:"),
                languageName);
            if (language == IntPtr.Zero)
            {
                throw new InvalidOperationException("OSAKit JavaScript language is unavailable.");
            }

            var nativeSource = SendString(
                nsString,
                GetSelector("stringWithUTF8String:"),
                source);
            var script = Send(GetClass("OSAScript"), GetSelector("alloc"));
            script = Send(
                script,
                GetSelector("initWithSource:language:"),
                nativeSource,
                language);
            if (script == IntPtr.Zero)
            {
                throw new InvalidOperationException("OSAKit could not initialize the Excel automation script.");
            }

            try
            {
                var error = IntPtr.Zero;
                var result = Send(script, GetSelector("executeAndReturnError:"), ref error);
                if (result == IntPtr.Zero)
                {
                    throw new InvalidOperationException(
                        $"OSAKit execution failed: {Describe(error)}");
                }

                var value = Send(result, GetSelector("stringValue"));
                return ToManagedString(value);
            }
            finally
            {
                Send(script, GetSelector("release"));
            }
        }
        finally
        {
            Send(pool, GetSelector("drain"));
        }
    }

    private static string Describe(IntPtr value)
    {
        if (value == IntPtr.Zero)
        {
            return "Unknown Apple Events error.";
        }

        return ToManagedString(Send(value, GetSelector("description")));
    }

    private static string ToManagedString(IntPtr value)
    {
        if (value == IntPtr.Zero)
        {
            return "";
        }

        var utf8 = Send(value, GetSelector("UTF8String"));
        return Marshal.PtrToStringUTF8(utf8) ?? "";
    }

    private static IntPtr GetClass(string name) =>
        WithCString(name, ObjectiveCGetClass);

    private static IntPtr GetSelector(string name) =>
        WithCString(name, SelectorRegisterName);

    private static IntPtr Send(IntPtr receiver, IntPtr selector) =>
        ObjectiveCMessageSend(receiver, selector);

    private static IntPtr Send(IntPtr receiver, IntPtr selector, IntPtr argument) =>
        ObjectiveCMessageSendOne(receiver, selector, argument);

    private static IntPtr Send(
        IntPtr receiver,
        IntPtr selector,
        IntPtr firstArgument,
        IntPtr secondArgument) =>
        ObjectiveCMessageSendTwo(receiver, selector, firstArgument, secondArgument);

    private static IntPtr Send(IntPtr receiver, IntPtr selector, ref IntPtr argument) =>
        ObjectiveCMessageSendReference(receiver, selector, ref argument);

    private static IntPtr SendString(
        IntPtr receiver,
        IntPtr selector,
        string argument) =>
        WithCString(argument, pointer => ObjectiveCMessageSendString(receiver, selector, pointer));

    [DllImport(ObjectiveC, EntryPoint = "objc_getClass")]
    private static extern IntPtr ObjectiveCGetClass(IntPtr name);

    [DllImport(ObjectiveC, EntryPoint = "sel_registerName")]
    private static extern IntPtr SelectorRegisterName(IntPtr name);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    private static extern IntPtr ObjectiveCMessageSend(IntPtr receiver, IntPtr selector);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    private static extern IntPtr ObjectiveCMessageSendOne(
        IntPtr receiver,
        IntPtr selector,
        IntPtr argument);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    private static extern IntPtr ObjectiveCMessageSendTwo(
        IntPtr receiver,
        IntPtr selector,
        IntPtr firstArgument,
        IntPtr secondArgument);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    private static extern IntPtr ObjectiveCMessageSendReference(
        IntPtr receiver,
        IntPtr selector,
        ref IntPtr argument);

    [DllImport(ObjectiveC, EntryPoint = "objc_msgSend")]
    private static extern IntPtr ObjectiveCMessageSendString(
        IntPtr receiver,
        IntPtr selector,
        IntPtr argument);

    private static IntPtr WithCString(string value, Func<IntPtr, IntPtr> action)
    {
        var bytes = new byte[Encoding.UTF8.GetByteCount(value) + 1];
        Encoding.UTF8.GetBytes(value, bytes);
        var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
        try
        {
            return action(handle.AddrOfPinnedObject());
        }
        finally
        {
            handle.Free();
        }
    }
}
