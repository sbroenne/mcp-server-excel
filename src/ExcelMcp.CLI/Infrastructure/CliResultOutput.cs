using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class CliResultOutput
{
    internal static int WriteJson(string json, string? command = null) =>
        Write(new ServiceResponse { Success = true, Result = json }, command);

    internal static ResultJson Read(ServiceResponse response, string command) =>
        ResultJson.Read(response.Result, command, response.Success, response.ErrorMessage,
            response.ErrorCategory, response.SessionId, response.ExceptionType, response.HResult,
            response.InnerError, response.PartialState, ServiceRegistry.GetResultContract(command));

    internal static int Write(
        ServiceResponse response, string? command = null, long? durationMs = null,
        int? index = null)
    {
        command ??= response.Command ?? CliCommandRuntime.Current.CommandName;
        var result = Read(response, command);
        return Write(result, command, durationMs, index);
    }

    internal static int Write(ResultJson result, string command, long? durationMs = null, int? index = null)
    {
        CliCommandRuntime.Current.Output.WriteLine(result.ToCliJson(
            command, durationMs ?? CliCommandRuntime.Current.OperationTimer.ElapsedMilliseconds, index));
        return result.Success ? 0 : 1;
    }
}
