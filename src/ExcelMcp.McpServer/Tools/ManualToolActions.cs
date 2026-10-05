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
    Create
}

[JsonConverter(typeof(JsonStringEnumConverter<FileReadAction>))]
public enum FileReadAction
{
    [JsonStringEnumMemberName("list")]
    List,
    [JsonStringEnumMemberName("test")]
    Test
}

[JsonConverter(typeof(JsonStringEnumConverter<WorksheetWriteAction>))]
public enum WorksheetWriteAction
{
    [JsonStringEnumMemberName("create")]
    Create,
    [JsonStringEnumMemberName("rename")]
    Rename,
    [JsonStringEnumMemberName("copy")]
    Copy,
    [JsonStringEnumMemberName("delete")]
    Delete,
    [JsonStringEnumMemberName("move")]
    Move,
    [JsonStringEnumMemberName("copy-to-file")]
    CopyToFile,
    [JsonStringEnumMemberName("move-to-file")]
    MoveToFile
}

[JsonConverter(typeof(JsonStringEnumConverter<WorksheetReadAction>))]
public enum WorksheetReadAction
{
    [JsonStringEnumMemberName("list")]
    List
}
