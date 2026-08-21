using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class DaemonPipeIdentity
{
    internal static string GetHash(string pipeName) =>
        Hash(pipeName.ToUpperInvariant());

    internal static string GetCaseSensitiveHash(string pipeName) =>
        Hash(pipeName);

    internal static IReadOnlyList<string> GetLegacyCaseVariants(string pipeName) =>
        [
            .. new[]
            {
                pipeName,
                pipeName.ToUpperInvariant(),
                pipeName.ToLowerInvariant()
            }.Distinct(StringComparer.Ordinal)
        ];

    private static string Hash(string pipeName) =>
        Convert.ToHexString(SHA256.HashData(
            Encoding.UTF8.GetBytes(pipeName)));
}
