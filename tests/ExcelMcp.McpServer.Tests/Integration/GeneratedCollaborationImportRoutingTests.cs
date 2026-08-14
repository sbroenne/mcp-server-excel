using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

[Trait("Category", "Integration")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "CollaborationImport")]
public sealed class GeneratedCollaborationImportRoutingTests
{
    [Fact]
    public void RangeLinkGeneratedRouting_ContainsAllThreadedCommentActions()
    {
        Assert.Contains("add-threaded-comment", ServiceRegistry.RangeLink.ValidActions);
        Assert.Contains("list-threaded-comments", ServiceRegistry.RangeLink.ValidActions);
        Assert.Contains("add-threaded-comment-reply", ServiceRegistry.RangeLink.ValidActions);
        Assert.Contains("delete-threaded-comment", ServiceRegistry.RangeLink.ValidActions);
    }

    [Fact]
    public void QueryTableGeneratedRouting_ContainsAllActions()
    {
        Assert.Equal(
            [
                "list",
                "view",
                "create-text",
                "create-web",
                "set-properties",
                "refresh",
                "get-refresh-status",
                "cancel-refresh",
                "delete"
            ],
            ServiceRegistry.QueryTable.ValidActions);
    }

    [Fact]
    public void ConnectionGeneratedRouting_ContainsRefreshControlActions()
    {
        Assert.Contains("get-refresh-status", ServiceRegistry.Connection.ValidActions);
        Assert.Contains("cancel-refresh", ServiceRegistry.Connection.ValidActions);
    }
}
