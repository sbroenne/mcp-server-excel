using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal static class ServiceCommandProxy
{
    private static readonly JsonSerializerOptions ResultJsonOptions = CreateResultJsonOptions();

    internal static T Create<T>(PersistentServiceWorkbookFixture fixture)
        where T : class
    {
        var proxy = DispatchProxy.Create<T, ServiceCommandDispatchProxy>();
        ((ServiceCommandDispatchProxy)(object)proxy).Fixture = fixture;
        return proxy;
    }

    internal static T DeserializeResult<T>(string json) =>
        JsonSerializer.Deserialize<T>(json, ResultJsonOptions)
        ?? throw new InvalidOperationException(
            $"Result payload could not be deserialized as {typeof(T).Name}.");

    [System.Diagnostics.CodeAnalysis.SuppressMessage(
        "Performance",
        "CA1852:Seal internal types",
        Justification = "DispatchProxy requires the proxy base type to remain unsealed.")]
    private class ServiceCommandDispatchProxy : DispatchProxy
    {
        internal PersistentServiceWorkbookFixture Fixture { get; set; } = null!;

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            ArgumentNullException.ThrowIfNull(targetMethod);
            args ??= [];

            var declaringType = targetMethod.DeclaringType
                ?? throw new InvalidOperationException(
                    $"Service command method '{targetMethod.Name}' has no declaring type.");
            var category = declaringType.GetCustomAttribute<ServiceCategoryAttribute>()?.Category
                ?? throw new InvalidOperationException(
                    $"Service command interface '{declaringType.FullName}' has no ServiceCategory.");
            var action = GetActionName(targetMethod);

            var parameters = targetMethod.GetParameters();
            var requestArgs = new Dictionary<string, object?>(StringComparer.Ordinal);
            for (var index = 0; index < parameters.Length; index++)
            {
                var parameter = parameters[index];
                if (typeof(IExcelBatch).IsAssignableFrom(parameter.ParameterType))
                {
                    Fixture.ValidateBatchToken(args[index]);
                    continue;
                }
                if (IsAmbientProgressParameter(parameter))
                {
                    continue;
                }

                requestArgs.Add(
                    GetParameterName(parameter, index),
                    NormalizeRequestValue(parameter, args[index]));
            }

            var response = Fixture.SendAsync($"{category}.{action}", requestArgs)
                .GetAwaiter()
                .GetResult();
            if (targetMethod.ReturnType == typeof(void))
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(response.Result))
            {
                throw new InvalidOperationException(
                    $"{category}.{action} succeeded without a result payload.");
            }

            return JsonSerializer.Deserialize(
                       response.Result,
                       targetMethod.ReturnType,
                       ResultJsonOptions)
                   ?? throw new InvalidOperationException(
                       $"{category}.{action} returned an empty " +
                       $"{targetMethod.ReturnType.Name} payload.");
        }
    }

    private static JsonSerializerOptions CreateResultJsonOptions()
    {
        var options = new JsonSerializerOptions(ServiceProtocol.JsonOptions);
        options.Converters.Insert(0, new InferredObjectJsonConverter());
        return options;
    }

    internal static string GetActionName(MethodInfo method) =>
        method.GetCustomAttribute<ServiceActionAttribute>()?.Action
        ?? System.Text.RegularExpressions.Regex.Replace(
                method.Name,
                "([a-z0-9])([A-Z])",
                "$1-$2",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant)
            .ToLowerInvariant();

    internal static string GetParameterName(ParameterInfo parameter, int index) =>
        parameter.GetCustomAttribute<FromStringAttribute>()?.ExposedName
        ?? parameter.Name
        ?? throw new InvalidOperationException(
            $"Service command parameter {index} has no name.");

    internal static bool IsAmbientProgressParameter(ParameterInfo parameter) =>
        parameter.ParameterType.IsGenericType
        && parameter.ParameterType.GetGenericTypeDefinition() == typeof(IProgress<>);

    internal static object? NormalizeRequestValue(
        ParameterInfo parameter,
        object? value)
    {
        var parameterType = Nullable.GetUnderlyingType(parameter.ParameterType)
            ?? parameter.ParameterType;
        return parameterType == typeof(TimeSpan) && value is TimeSpan timeout
            ? checked((int)timeout.TotalSeconds)
            : value;
    }

    private sealed class InferredObjectJsonConverter : JsonConverter<object>
    {
        public override object? Read(
            ref Utf8JsonReader reader,
            Type typeToConvert,
            JsonSerializerOptions options) =>
            reader.TokenType switch
            {
                JsonTokenType.Null => null,
                JsonTokenType.True => true,
                JsonTokenType.False => false,
                JsonTokenType.String => reader.GetString(),
                JsonTokenType.Number when reader.TryGetInt32(out var integer) => integer,
                JsonTokenType.Number when reader.TryGetInt64(out var longInteger) => longInteger,
                JsonTokenType.Number => reader.GetDouble(),
                _ => JsonDocument.ParseValue(ref reader).RootElement.Clone()
            };

        public override void Write(
            Utf8JsonWriter writer,
            object value,
            JsonSerializerOptions options) =>
            JsonSerializer.Serialize(writer, value, value.GetType(), options);
    }
}
