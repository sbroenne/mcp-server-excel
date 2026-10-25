using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Owns access to one workbook document-property collection.
/// </summary>
internal sealed partial class ExcelDocumentProperties : IDisposable
{
    private object? _collection;

    private ExcelDocumentProperties(object collection) => _collection = collection;

    internal static ExcelDocumentProperties Open(Excel.Workbook workbook, bool builtIn)
    {
        object collection = builtIn ? workbook.BuiltinDocumentProperties : workbook.CustomDocumentProperties;
        return new ExcelDocumentProperties(collection);
    }

    internal int Count => Convert.ToInt32(ReadCollection(Member.Count) ??
        throw new InvalidOperationException("Excel returned no document-property count."), CultureInfo.InvariantCulture);

    internal Property Item(int index) => GetItem(index);

    internal Property Item(string name) => GetItem(name);

    private Property GetItem(object index) =>
        TakeProperty(InvokeCollection(Member.Item, INVOKEKIND.INVOKE_PROPERTYGET, [index]));

    internal Property Add(string name, Office.MsoDocProperties type, object value) =>
        TakeProperty(InvokeCollection(Member.Add, INVOKEKIND.INVOKE_FUNC, [name, false, (int)type, value]));

    private static Property TakeProperty(object? value)
    {
        if (value is null)
            throw new InvalidOperationException("Excel returned no document property.");
        try
        {
            var property = new Property(value);
            value = null;
            return property;
        }
        finally
        {
            ComUtilities.Release(ref value);
        }
    }

    private object? ReadCollection(Member member) =>
        InvokeCollection(member, INVOKEKIND.INVOKE_PROPERTYGET, []);

    private object? InvokeCollection(Member member, INVOKEKIND kind, object?[] arguments)
    {
        ObjectDisposedException.ThrowIf(_collection is null, this);
        return Invoke(_collection, member, kind, arguments);
    }

    public void Dispose()
    {
        ComUtilities.Release(ref _collection);
        GC.SuppressFinalize(this);
    }

    internal sealed class Property : IDisposable
    {
        private object? _property;

        internal Property(object property)
        {
            ValidateProperty(property);
            _property = property;
        }

        internal string Name => (string)(Read(Member.Name) ??
            throw new InvalidOperationException("Excel returned no document-property name."));

        internal Office.MsoDocProperties Type =>
            (Office.MsoDocProperties)Convert.ToInt32(Read(Member.Type) ??
                throw new InvalidOperationException("Excel returned no document-property type."), CultureInfo.InvariantCulture);

        internal bool LinkToContent => Convert.ToBoolean(Read(Member.LinkToContent) ??
            throw new InvalidOperationException("Excel returned no document-property link flag."), CultureInfo.InvariantCulture);

        internal object? Value => Read(Member.Value);

        internal void SetValue(string value) => Call(Member.Value, INVOKEKIND.INVOKE_PROPERTYPUT, [value]);

        internal void Delete() => Call(Member.Delete, INVOKEKIND.INVOKE_FUNC, []);

        private object? Read(Member member) => Call(member, INVOKEKIND.INVOKE_PROPERTYGET, []);

        private object? Call(Member member, INVOKEKIND kind, object?[] arguments)
        {
            ObjectDisposedException.ThrowIf(_property is null, this);
            return Invoke(_property, member, kind, arguments);
        }

        public void Dispose()
        {
            ComUtilities.Release(ref _property);
            GC.SuppressFinalize(this);
        }
    }

    // Only this fixed document-property surface may use the approved native adapter.
    private enum Member { Count, Item, Add, Name, Type, LinkToContent, Value, Delete }

    private static void ValidateProperty(object property)
    {
        ITypeInfo? information = null;
        try
        {
            var dispatch = (IPropertyDispatch)property;
            int result = dispatch.GetTypeInfo(0, 0, out information);
            if (result < 0)
                throw new NativePropertyException("Reading document-property type information failed.",
                    result, Marshal.GetExceptionForHR(result));
            information.GetTypeAttr(out nint attributes);
            try
            {
                var type = Marshal.PtrToStructure<TYPEATTR>(attributes);
                if (type.guid != typeof(Office.DocumentProperty).GUID)
                    throw new ArgumentException("Only Excel document-property items are supported by this adapter.", nameof(property));
            }
            finally
            {
                information.ReleaseTypeAttr(attributes);
            }
        }
        finally
        {
            ComUtilities.Release(ref information);
        }
    }

    private static unsafe object? Invoke(object target, Member member, INVOKEKIND kind, object?[] arguments)
    {
        var dispatch = (IPropertyDispatch)target;
        var reserved = Guid.Empty;
        uint locale = (uint)CultureInfo.CurrentCulture.LCID;
        int[] identifiers = new int[1];
        int resolved = dispatch.GetIDsOfNames(ref reserved, [member.ToString()], 1, locale, identifiers);
        if (resolved < 0)
            throw CreateNativeFailure(member, resolved, default, 0);

        int variantSize = IntPtr.Size == 8 ? 24 : 16;
        nint memory = 0;
        int initialized = 0;
        Exception? operationFailure = null;
        object? result = null;
        try
        {
            if (arguments.Length > 0)
            {
                int bytes = checked(variantSize * arguments.Length);
                memory = Marshal.AllocCoTaskMem(bytes);
                new Span<byte>((void*)memory, bytes).Clear();
                for (int index = 0; index < arguments.Length; index++)
                {
                    initialized++;
                    Marshal.GetNativeVariantForObject(arguments[arguments.Length - index - 1], memory + index * variantSize);
                }
            }
            int propertyPut = -3; // Standard DISPID_PROPERTYPUT, not an Excel member identifier.
            var parameters = new DISPPARAMS
            {
                rgvarg = memory,
                cArgs = arguments.Length,
                cNamedArgs = kind == INVOKEKIND.INVOKE_PROPERTYPUT ? 1 : 0,
                rgdispidNamedArgs = kind == INVOKEKIND.INVOKE_PROPERTYPUT ? (nint)(&propertyPut) : 0
            };
            int invoked = dispatch.Invoke(identifiers[0], ref reserved, locale, (ushort)kind,
                ref parameters, out result, out EXCEPINFO error, out uint argumentError);
            if (invoked < 0)
            {
                if (error.pfnDeferredFillIn != 0)
                {
                    var fill = Marshal.GetDelegateForFunctionPointer<FillException>(error.pfnDeferredFillIn);
                    int filled = fill(ref error);
                    if (filled < 0)
                        throw new AggregateException(CreateNativeFailure(member, invoked, error, argumentError),
                            new NativePropertyException("Excel could not supply deferred document-property error information.",
                                filled, Marshal.GetExceptionForHR(filled)));
                }
                throw CreateNativeFailure(member, invoked, error, argumentError);
            }
            return result;
        }
        catch (Exception failure)
        {
            operationFailure = failure;
            throw;
        }
        finally
        {
            ClearArguments(memory, initialized, variantSize, ref result, operationFailure);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Collects result-release failures only to throw them together with every argument-cleanup failure and the original operation failure; none are discarded.")]
    private static void ClearArguments(nint memory, int initialized, int variantSize,
        ref object? result, Exception? operationFailure)
    {
        List<Exception>? cleanupFailures = null;
        for (int index = initialized - 1; index >= 0; index--)
        {
            int cleared = VariantClear(memory + index * variantSize);
            if (cleared < 0)
                (cleanupFailures ??= []).Add(new NativePropertyException("Releasing a document-property argument failed.",
                    cleared, Marshal.GetExceptionForHR(cleared)));
        }
        Marshal.FreeCoTaskMem(memory);
        if (operationFailure is not null || cleanupFailures is not null)
        {
            try
            {
                ComUtilities.Release(ref result);
            }
            catch (Exception failure)
            {
                (cleanupFailures ??= []).Add(failure);
            }
        }
        if (cleanupFailures is not null)
        {
            if (operationFailure is not null)
                cleanupFailures.Insert(0, operationFailure);
            throw new AggregateException("Document-property operation cleanup failed.", cleanupFailures);
        }
    }

    private static NativePropertyException CreateNativeFailure(Member member, int result, EXCEPINFO error, uint argumentError)
    {
        int code = result;
        if (result == unchecked((int)0x80020009)) // Standard DISP_E_EXCEPTION.
        {
            if (error.scode < 0)
                code = error.scode;
            else if (error.wCode != 0)
                code = unchecked((int)0x800A0000) | unchecked((ushort)error.wCode);
        }
        var cause = Marshal.GetExceptionForHR(code);
        string description = error.bstrDescription ?? cause?.Message ??
            "Excel returned an unspecified native error.";
        var failure = new NativePropertyException(
            $"Excel document-property member '{member}' failed (0x{code:X8}): {description}",
            code, cause);
        failure.Data["DispatchHResult"] = result;
        failure.Data["ArgumentIndex"] = argumentError;
        if (!string.IsNullOrEmpty(error.bstrSource))
            failure.Source = error.bstrSource;
        return failure;
    }

    [LibraryImport("oleaut32.dll")]
    private static partial int VariantClear(nint value);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int FillException(ref EXCEPINFO information);

    private sealed class NativePropertyException : COMException
    {
        internal NativePropertyException(string message, int result, Exception? cause) : base(message, cause)
        {
            HResult = result;
        }
    }

    // Microsoft's standard IDispatch layout (oaidl.h); member IDs are resolved from Excel.
    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyDispatch
    {
        [PreserveSig]
        int GetTypeInfoCount(out uint count);

        [PreserveSig]
        int GetTypeInfo(uint index, uint locale, out ITypeInfo information);

        [PreserveSig]
        int GetIDsOfNames(ref Guid reserved,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)] string[] names,
            uint count, uint locale,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] int[] identifiers);

        [PreserveSig]
        int Invoke(int identifier, ref Guid reserved, uint locale, ushort flags,
            ref DISPPARAMS parameters, [MarshalAs(UnmanagedType.Struct)] out object? result,
            out EXCEPINFO error, out uint argumentError);
    }
}
