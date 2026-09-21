using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class DaemonPipeIdentity
{
    internal static string GetHash(string pipeName) =>
        Hash(pipeName.ToUpperInvariant());

    private static string Hash(string pipeName) =>
        Convert.ToHexString(SHA256.HashData(
            Encoding.UTF8.GetBytes(pipeName)));
}
