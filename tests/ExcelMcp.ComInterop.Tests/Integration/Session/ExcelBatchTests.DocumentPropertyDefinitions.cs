using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

public partial class ExcelBatchTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_DocumentPropertyDefinitions_RecordNativeTypeLibrary(bool builtIn)
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            object? properties = null;
            ITypeInfo? typeInfo = null;
            ITypeLib? library = null;
            try
            {
                properties = builtIn ? ctx.Book.BuiltinDocumentProperties : ctx.Book.CustomDocumentProperties;
                var metadata = (IAutomationTypeInformation)properties;
                metadata.GetTypeInfoCount(out uint count);
                Assert.Equal(1u, count);
                metadata.GetTypeInfo(0, 0, out typeInfo);
                typeInfo.GetDocumentation(-1, out string typeName, out _, out _, out _);
                typeInfo.GetTypeAttr(out nint attributes);
                try
                {
                    var type = Marshal.PtrToStructure<TYPEATTR>(attributes);
                    Assert.Equal(typeof(Microsoft.Office.Core.DocumentProperties).GUID, type.guid);
                    _output.WriteLine(
                        $"{(builtIn ? "Built-in" : "Custom")} collection: {typeName}; " +
                        $"GUID={type.guid}; kind={type.typekind}; flags={type.wTypeFlags}; " +
                        $"functions={type.cFuncs}; interfaces={type.cImplTypes}");
                    for (int index = 0; index < type.cImplTypes; index++)
                    {
                        ITypeInfo? baseInfo = null;
                        try
                        {
                            typeInfo.GetRefTypeOfImplType(index, out int reference);
                            typeInfo.GetImplTypeFlags(index, out IMPLTYPEFLAGS flags);
                            typeInfo.GetRefTypeInfo(reference, out baseInfo);
                            baseInfo.GetDocumentation(-1, out string baseName, out _, out _, out _);
                            baseInfo.GetTypeAttr(out nint baseAttributes);
                            try
                            {
                                var baseType = Marshal.PtrToStructure<TYPEATTR>(baseAttributes);
                                _output.WriteLine(
                                    $"Base {baseName}; GUID={baseType.guid}; kind={baseType.typekind}; " +
                                    $"flags={flags}; type flags={baseType.wTypeFlags}");
                            }
                            finally
                            {
                                baseInfo.ReleaseTypeAttr(baseAttributes);
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref baseInfo);
                        }
                    }
                    for (int index = 0; index < type.cFuncs; index++)
                    {
                        typeInfo.GetFuncDesc(index, out nint descriptor);
                        try
                        {
                            var function = Marshal.PtrToStructure<FUNCDESC>(descriptor);
                            typeInfo.GetDocumentation(function.memid, out string memberName, out _, out _, out _);
                            _output.WriteLine(
                                $"Member {function.memid}: {memberName}; {function.invkind}; " +
                                $"parameters={function.cParams}; return={function.elemdescFunc.tdesc.vt}");
                        }
                        finally
                        {
                            typeInfo.ReleaseFuncDesc(descriptor);
                        }
                    }
                }
                finally
                {
                    typeInfo.ReleaseTypeAttr(attributes);
                }

                typeInfo.GetContainingTypeLib(out library, out int typeIndex);
                library.GetDocumentation(-1, out string libraryName, out _, out _, out _);
                library.GetLibAttr(out nint libraryAttributes);
                try
                {
                    var identity = Marshal.PtrToStructure<TYPELIBATTR>(libraryAttributes);
                    _output.WriteLine(
                        $"Library {libraryName}; GUID={identity.guid}; " +
                        $"version={identity.wMajorVerNum}.{identity.wMinorVerNum}; " +
                        $"LCID={identity.lcid}; platform={identity.syskind}; type index={typeIndex}");
                }
                finally
                {
                    library.ReleaseTLibAttr(libraryAttributes);
                }
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref library);
                ComUtilities.Release(ref typeInfo);
                ComUtilities.Release(ref properties);
            }
        });
    }

    [Fact]
    public void Execute_DocumentPropertyDefinitions_RecordBuiltInItemInterface()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            object? properties = null;
            System.Collections.IEnumerator? enumerator = null;
            object? property = null;
            ITypeInfo? typeInfo = null;
            try
            {
                properties = ctx.Book.BuiltinDocumentProperties;
                enumerator = ((System.Collections.IEnumerable)properties).GetEnumerator();
                Assert.True(enumerator.MoveNext(), "The built-in property collection must contain an item.");
                property = enumerator.Current;
                var metadata = (IAutomationTypeInformation)property;
                metadata.GetTypeInfo(0, 0, out typeInfo);
                typeInfo.GetDocumentation(-1, out string typeName, out _, out _, out _);
                typeInfo.GetTypeAttr(out nint attributes);
                try
                {
                    var type = Marshal.PtrToStructure<TYPEATTR>(attributes);
                    Assert.Equal(typeof(Microsoft.Office.Core.DocumentProperty).GUID, type.guid);
                    _output.WriteLine($"Item {typeName}; GUID={type.guid}; kind={type.typekind}; flags={type.wTypeFlags}");
                }
                finally
                {
                    typeInfo.ReleaseTypeAttr(attributes);
                }

                var failure = Assert.Throws<InvalidCastException>(
                    () => (Microsoft.Office.Core.DocumentProperty)property);
                Assert.Equal(unchecked((int)0x80004002), failure.HResult);
                _output.WriteLine("The individual item also rejects its Office definition interface: 0x80004002.");
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref typeInfo);
                ComUtilities.Release(ref property);
                ComUtilities.Release(ref enumerator);
                ComUtilities.Release(ref properties);
            }
        });
    }

    // The documented IDispatch metadata prefix from Microsoft's oaidl.h; no member lookup or invocation.
    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IAutomationTypeInformation
    {
        void GetTypeInfoCount(out uint count);
        void GetTypeInfo(uint index, uint locale, out ITypeInfo typeInfo);
    }
}
