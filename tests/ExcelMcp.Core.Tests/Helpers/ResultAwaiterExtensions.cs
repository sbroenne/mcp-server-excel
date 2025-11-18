using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Enables awaiting synchronous Core command results in tests without changing production APIs.
/// Wraps the result inside a completed task so existing async test patterns continue to compile.
/// </summary>
public static class ResultAwaiterExtensions
{
    /// <summary>
    /// Provides an awaiter for any result derived from <see cref="ResultBase"/>.
    /// </summary>
    /// <typeparam name="T">Result type</typeparam>
    /// <param name="result">Result instance to await</param>
    /// <returns>Awaiter that immediately returns the provided result</returns>
    public static TaskAwaiter<T?> GetAwaiter<T>(this T? result)
        where T : ResultBase
    {
        return Task.FromResult(result).GetAwaiter();
    }
}
