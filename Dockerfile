# Dockerfile for mcp-server-excel
# ============================================================================
# PURPOSE: Glama.ai Tool Discovery and MCP Registry Inspection ONLY
# ============================================================================
#
# This Dockerfile exists solely to enable Glama.ai (https://glama.ai) to:
#   - Inspect the MCP server and discover available tools
#   - Display tool schemas, parameters, and descriptions in their registry
#   - Allow users to browse capabilities before installing locally
#
# IMPORTANT LIMITATIONS:
#   - This container CANNOT perform actual Excel operations
#   - Excel COM automation requires Windows + Microsoft Excel installed
#   - For real Excel automation, install the MCP server locally on Windows
#
# See: https://glama.ai/mcp/servers/@sbroenne/mcp-server-excel
# ============================================================================

# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Create a clean NuGet config to avoid Windows-specific fallback folder paths
RUN echo '<?xml version="1.0" encoding="utf-8"?><configuration><packageSources><clear /><add key="nuget.org" value="https://api.nuget.org/v3/index.json" /></packageSources></configuration>' > /src/nuget.config

# Copy project files for restore
COPY Directory.Build.props Directory.Packages.props ./
COPY src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj src/ExcelMcp.ComInterop/
COPY src/ExcelMcp.Core/ExcelMcp.Core.csproj src/ExcelMcp.Core/
COPY src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj src/ExcelMcp.McpServer/

# Restore dependencies
RUN dotnet restore src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj --configfile /src/nuget.config

# Copy source code
COPY src/ src/

# Build and publish
RUN dotnet publish src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj \
    -c Release \
    -o /app/publish

# Runtime stage
FROM mcr.microsoft.com/dotnet/runtime:8.0
WORKDIR /app

# Copy published application
COPY --from=build /app/publish .

# MCP servers communicate via stdio
ENTRYPOINT ["dotnet", "Sbroenne.ExcelMcp.McpServer.dll"]
