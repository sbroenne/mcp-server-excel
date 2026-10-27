using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

[JsonConverter(typeof(JsonStringEnumConverter<FileWriteAction>))]
public enum FileWriteAction
{
    [JsonStringEnumMemberName("open")]
    Open,
    [JsonStringEnumMemberName("close")]
    Close,
    [JsonStringEnumMemberName("create")]
    Create,
    [JsonStringEnumMemberName("save")]
    Save,
    [JsonStringEnumMemberName("attach")]
    Attach,
    [JsonStringEnumMemberName("enable-editing")]
    EnableEditing,
    [JsonStringEnumMemberName("enable-connections")]
    EnableConnections,
    [JsonStringEnumMemberName("enable-macros")]
    EnableMacros
}

[JsonConverter(typeof(JsonStringEnumConverter<FileReadAction>))]
public enum FileReadAction
{
    [JsonStringEnumMemberName("list")]
    List,
    [JsonStringEnumMemberName("test")]
    Test,
    [JsonStringEnumMemberName("discover")]
    Discover,
    [JsonStringEnumMemberName("notices")]
    Notices
}
