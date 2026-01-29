using Microsoft.Extensions.DependencyInjection;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal sealed class TypeRegistrar : ITypeRegistrar, IDisposable
{
    private readonly IServiceCollection _services;
    private TypeResolver? _resolver;
    private bool _disposed;

    public TypeRegistrar(IServiceCollection services)
    {
        _services = services ?? throw new ArgumentNullException(nameof(services));
    }

    public ITypeResolver Build()
    {
        _resolver = new TypeResolver(_services.BuildServiceProvider());
        return _resolver;
    }

    public void Register(Type service, Type implementation)
    {
        _services.AddSingleton(service, implementation);
    }

    public void RegisterInstance(Type service, object implementation)
    {
        _services.AddSingleton(service, implementation);
    }

    public void RegisterLazy(Type service, Func<object> factory)
    {
        _services.AddSingleton(service, _ => factory());
    }

    public void Dispose()
    {
        if (_disposed) return;
        _resolver?.Dispose();
        _disposed = true;
    }
}

internal sealed class TypeResolver : ITypeResolver, IDisposable
{
    private readonly ServiceProvider _provider;
    private bool _disposed;

    public TypeResolver(ServiceProvider provider)
    {
        _provider = provider;
    }

    public object? Resolve(Type? type)
    {
        if (type == null)
        {
            return null;
        }

        return _provider.GetService(type);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _provider.Dispose();
        _disposed = true;
    }
}
