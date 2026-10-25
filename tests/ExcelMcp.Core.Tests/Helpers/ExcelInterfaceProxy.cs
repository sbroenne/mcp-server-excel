using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

public class ExcelInterfaceProxy : DispatchProxy
{
    public Func<MethodInfo, object?[]?, object?> Read { get; set; } = null!;

    public static T Create<T>(Func<MethodInfo, object?[]?, object?> read, Assembly? owner = null) where T : class
    {
        var interfaceType = (owner ?? typeof(ConnectionCommands).Assembly).GetType(typeof(T).FullName!) ?? typeof(T);
        var value = (T)Create(interfaceType, typeof(ExcelInterfaceProxy));
        ((ExcelInterfaceProxy)(object)value).Read = read;
        return value;
    }

    protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
        Read(targetMethod ?? throw new InvalidOperationException("No intercepted interface member."), args);
}
