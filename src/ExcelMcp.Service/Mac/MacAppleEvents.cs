using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacAppleEvents
{
    private const string Framework =
        "/System/Library/Frameworks/CoreServices.framework/Frameworks/AE.framework/AE";

    internal const int MissingParameter = -1701;

    internal static TimeSpan Remaining(TimeSpan timeout, long started)
    {
        var remaining = timeout - Stopwatch.GetElapsedTime(started);
        return remaining > TimeSpan.Zero ? remaining
            : throw new TimeoutException("Mac Excel operation timed out; its outcome may be uncertain.");
    }

    internal static uint Code(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        if (value.Length != 4 || value.Any(character => character > 0x7f))
        {
            throw new ArgumentException("Apple Event codes require exactly four ASCII characters.", nameof(value));
        }

        return ((uint)value[0] << 24) | ((uint)value[1] << 16) | ((uint)value[2] << 8) | value[3];
    }

    internal static nint TimeoutTicks(TimeSpan timeout)
    {
        if (timeout <= TimeSpan.Zero || timeout.TotalSeconds > (double)nint.MaxValue / 60)
        {
            throw new ArgumentOutOfRangeException(nameof(timeout), "Apple Event timeout must be positive and bounded.");
        }

        return checked((nint)Math.Ceiling(timeout.TotalSeconds * 60));
    }

    internal static void Check(int status, string operation, bool isExcelReply = false)
    {
        if (status == -1712)
        {
            throw new TimeoutException($"Excel Apple Event '{operation}' timed out (OSStatus {status}); its outcome may be uncertain.");
        }
        if (status != 0)
        {
            throw new MacAppleEventException(status, operation, isExcelReply);
        }
    }

    internal static Descriptor Create(uint type, byte[] data)
    {
        if (!OperatingSystem.IsMacOS())
        {
            throw new PlatformNotSupportedException("Native Apple Events require macOS.");
        }

        return Acquire(result => AECreateDesc(type, data, data.Length, out result.Value), "AECreateDesc");
    }

    internal static Descriptor Text(string value) => Create(Code("utxt"), Encoding.Unicode.GetBytes(value));

    internal static Descriptor List()
    {
        if (!OperatingSystem.IsMacOS())
        {
            throw new PlatformNotSupportedException("Native Apple Events require macOS.");
        }
        return Acquire(value => AECreateList(IntPtr.Zero, 0, 0, out value.Value), "AECreateList");
    }

    internal static void Append(Descriptor list, Descriptor value)
    {
        list.RequireAlive();
        value.RequireAlive();
        if (list.Value.DescriptorType != Code("list"))
        {
            throw new ArgumentException("An Apple Event list is required.", nameof(list));
        }
        Check(AEPutDesc(ref list.Value, 0, ref value.Value), "AEPutDesc");
    }

    internal static Descriptor Record(uint type)
    {
        if (!OperatingSystem.IsMacOS())
        {
            throw new PlatformNotSupportedException("Native Apple Events require macOS.");
        }
        var result = Acquire(value => AECreateList(IntPtr.Zero, 0, 1, out value.Value), "AECreateList");
        result.Value.DescriptorType = type;
        return result;
    }

    internal static void PutKey(Descriptor record, uint keyword, Descriptor value)
    {
        record.RequireAlive();
        value.RequireAlive();
        // AEPutKeyDesc is an SDK macro alias, not an exported native entry point.
        Check(AEPutParamDesc(ref record.Value, keyword, ref value.Value), "AEPutKeyDesc");
    }

    internal static Descriptor Object(uint objectClass, Descriptor container, uint form, Descriptor selector)
    {
        container.RequireAlive();
        selector.RequireAlive();
        return Acquire(result => CreateObjSpecifier(
            objectClass, ref container.Value, form, ref selector.Value, 0, out result.Value), "CreateObjSpecifier");
    }

    internal static Descriptor Property(Descriptor container, uint property)
    {
        using var selector = Create(Code("type"), BitConverter.GetBytes(property));
        return Object(Code("prop"), container, Code("prop"), selector);
    }

    internal static Descriptor Event(uint eventClass, uint eventId)
    {
        using var target = Create(Code("bund"), Encoding.UTF8.GetBytes("com.microsoft.Excel"));
        return Acquire(result => AECreateAppleEvent(
            eventClass, eventId, ref target.Value, -1, 0, out result.Value), "AECreateAppleEvent");
    }

    internal static void Put(Descriptor appleEvent, uint keyword, Descriptor value)
    {
        appleEvent.RequireAlive();
        value.RequireAlive();
        Check(AEPutParamDesc(ref appleEvent.Value, keyword, ref value.Value), "AEPutParamDesc");
    }

    internal static JsonNode? Send(Descriptor appleEvent, TimeSpan timeout) =>
        SendCore(appleEvent, timeout, requiresResult: true);

    internal static void SendCommand(Descriptor appleEvent, TimeSpan timeout) =>
        SendCore(appleEvent, timeout, requiresResult: false);

    internal static Descriptor SendSpecifier(Descriptor appleEvent, TimeSpan timeout)
    {
        using var reply = SendReply(appleEvent, timeout);
        return Parameter(reply, Code("----"), optional: false, desiredType: Code("obj "))!;
    }

    private static JsonNode? SendCore(Descriptor appleEvent, TimeSpan timeout, bool requiresResult)
    {
        using var reply = SendReply(appleEvent, timeout);
        if (!requiresResult)
        {
            return null;
        }
        using var value = Parameter(reply, Code("----"), optional: false);
        return value!.Decode();
    }

    private static Descriptor SendReply(Descriptor appleEvent, TimeSpan timeout)
    {
        appleEvent.RequireAlive();
        var ticks = TimeoutTicks(timeout);
        // Wait for the reply, but never allow Excel to interact with a dialog.
        var reply = Acquire(result => AESendMessage(
            ref appleEvent.Value, out result.Value, 3 | 0x10, ticks), "AESendMessage");
        try
        {
            using var error = Parameter(reply, Code("errn"), optional: true, desiredType: Code("long"));
            if (error is not null)
            {
                var status = error.Decode()?.GetValue<int>()
                    ?? throw new InvalidOperationException("Excel returned an empty Apple Event error code.");
                if (status != 0)
                {
                    using var description = Parameter(reply, Code("errs"), optional: true, desiredType: Code("utxt"));
                    var message = description?.Decode()?.GetValue<string>();
                    Check(status, string.IsNullOrEmpty(message) ? "Excel reply" : $"Excel reply: {message}", isExcelReply: true);
                }
            }

            return reply;
        }
        catch
        {
            reply.Dispose();
            throw;
        }
    }

    private static Descriptor? Parameter(Descriptor appleEvent, uint keyword, bool optional, uint? desiredType = null)
    {
        var result = new Descriptor();
        var status = AEGetParamDesc(ref appleEvent.Value, keyword, desiredType ?? Code("****"), out result.Value);
        if (optional && status == MissingParameter)
        {
            result.Dispose();
            return null;
        }

        return FinishAcquire(result, status, "AEGetParamDesc");
    }

    private static Descriptor Acquire(Func<Descriptor, int> operation, string name)
    {
        var result = new Descriptor();
        try
        {
            return FinishAcquire(result, operation(result), name);
        }
        catch
        {
            result.Dispose();
            throw;
        }
    }

    private static Descriptor FinishAcquire(Descriptor result, int status, string name)
    {
        if (status != 0)
        {
            result.Dispose();
            Check(status, name);
        }
        return result;
    }

    internal sealed class Descriptor : IDisposable
    {
        internal MacAutomationAccess.Descriptor Value;
        private bool _disposed;

        internal void RequireAlive() => ObjectDisposedException.ThrowIf(_disposed, this);

        internal JsonNode? Decode()
        {
            RequireAlive();
            var type = Value.DescriptorType;
            if (type == Code("null") || type == Code("msng"))
            {
                return null;
            }
            if (type == Code("list"))
            {
                Check(AECountItems(ref Value, out var count), "AECountItems");
                if (count < 0 || count > int.MaxValue)
                {
                    throw new InvalidOperationException("Excel returned an invalid Apple Event list length.");
                }
                var values = new JsonArray();
                for (nint index = 1; index <= count; index++)
                {
                    using var item = Acquire(result => AEGetNthDesc(
                        ref Value, index, Code("****"), out _, out result.Value), "AEGetNthDesc");
                    values.Add(item.Decode());
                }
                return values;
            }

            var size = AEGetDescDataSize(ref Value);
            if (size < 0 || size > int.MaxValue)
            {
                throw new InvalidOperationException("Excel returned an invalid Apple Event descriptor size.");
            }
            var data = new byte[(int)size];
            Check(AEGetDescData(ref Value, data, size), "AEGetDescData");
            return DecodeScalar(type, data);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (Value.DataHandle != IntPtr.Zero)
            {
                Check(AEDisposeDesc(ref Value), "AEDisposeDesc");
            }
        }
    }

    internal static JsonNode? DecodeScalar(uint type, byte[] data)
    {
        if (type == Code("type") && data.Length == sizeof(uint)
            && BitConverter.ToUInt32(data) == Code("msng"))
        {
            return null;
        }
        if (type == Code("utxt") && data.Length % 2 == 0)
        {
            return JsonValue.Create(new UnicodeEncoding(false, false, true).GetString(data));
        }
        if (type == Code("utf8"))
        {
            return JsonValue.Create(new UTF8Encoding(false, true).GetString(data));
        }
        if (type == Code("long") && data.Length == sizeof(int))
        {
            return JsonValue.Create(BitConverter.ToInt32(data));
        }
        if (type == Code("enum") && data.Length == sizeof(uint))
        {
            return JsonValue.Create(BitConverter.ToUInt32(data));
        }
        if (type == Code("comp") && data.Length == sizeof(long))
        {
            return JsonValue.Create(BitConverter.ToInt64(data));
        }
        if (type == Code("doub") && data.Length == sizeof(double))
        {
            return JsonValue.Create(BitConverter.ToDouble(data));
        }
        if (type == Code("bool") && data.Length == 1 && data[0] <= 1)
        {
            return JsonValue.Create(data[0] == 1);
        }
        throw new InvalidOperationException($"Unsupported or malformed Excel Apple Event descriptor 0x{type:X8} ({data.Length} bytes).");
    }

    [DllImport(Framework)]
    private static extern int AECreateDesc(uint type, byte[] data, nint size, out MacAutomationAccess.Descriptor result);

    [DllImport(Framework)]
    private static extern int AECreateList(IntPtr data, nint size, byte isRecord, out MacAutomationAccess.Descriptor result);

    [DllImport(Framework)]
    private static extern int AECreateAppleEvent(
        uint eventClass, uint eventId, ref MacAutomationAccess.Descriptor target, short returnId,
        int transactionId, out MacAutomationAccess.Descriptor result);

    [DllImport(Framework)]
    private static extern int CreateObjSpecifier(
        uint objectClass, ref MacAutomationAccess.Descriptor container, uint form,
        ref MacAutomationAccess.Descriptor selector, byte disposeInputs, out MacAutomationAccess.Descriptor result);

    [DllImport(Framework)]
    private static extern int AEPutParamDesc(
        ref MacAutomationAccess.Descriptor appleEvent, uint keyword, ref MacAutomationAccess.Descriptor value);

    [DllImport(Framework)]
    private static extern int AEPutDesc(
        ref MacAutomationAccess.Descriptor list, nint index, ref MacAutomationAccess.Descriptor value);

    [DllImport(Framework)]
    private static extern int AESendMessage(
        ref MacAutomationAccess.Descriptor appleEvent, out MacAutomationAccess.Descriptor reply, int mode, nint timeout);

    [DllImport(Framework)]
    private static extern int AEGetParamDesc(
        ref MacAutomationAccess.Descriptor appleEvent, uint keyword, uint desiredType, out MacAutomationAccess.Descriptor result);

    [DllImport(Framework)]
    private static extern nint AEGetDescDataSize(ref MacAutomationAccess.Descriptor descriptor);

    [DllImport(Framework)]
    private static extern int AEGetDescData(ref MacAutomationAccess.Descriptor descriptor, [Out] byte[] data, nint size);

    [DllImport(Framework)]
    private static extern int AECountItems(ref MacAutomationAccess.Descriptor descriptor, out nint count);

    [DllImport(Framework)]
    private static extern int AEGetNthDesc(
        ref MacAutomationAccess.Descriptor descriptor, nint index, uint desiredType,
        out uint keyword, out MacAutomationAccess.Descriptor result);

    [DllImport(Framework)]
    private static extern int AEDisposeDesc(ref MacAutomationAccess.Descriptor descriptor);
}
