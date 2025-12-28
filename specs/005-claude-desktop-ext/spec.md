# Feature Specification: Claude Desktop Extension (MCPB) Packaging

**Feature Branch**: `005-claude-desktop-ext`  
**Created**: 2024-12-28  
**Status**: Draft  
**Input**: User description: "I want to package the MCP Server as Desktop Extension for Claude"

## User Scenarios & Testing *(mandatory)*

### User Story 1 - One-Click Installation (Priority: P1)

As a Claude Desktop user, I want to install the Excel MCP Server by simply double-clicking a downloaded `.mcpb` file and clicking "Install", so that I can use Excel automation with Claude without any technical setup.

**Why this priority**: This is the core value proposition - eliminating the friction of manual configuration, .NET installation, and JSON editing that currently blocks non-technical users from using the MCP Server.

**Independent Test**: Can be fully tested by downloading the `.mcpb` file, double-clicking it in Claude Desktop, and verifying that Excel tools appear in the Claude interface.

**Acceptance Scenarios**:

1. **Given** a user has Claude Desktop installed, **When** they double-click the `excel-mcp.mcpb` file, **Then** Claude Desktop displays an installation dialog with extension name, description, author, and an "Install" button.
2. **Given** a user is viewing the installation dialog, **When** they click "Install", **Then** the extension installs and Excel tools become available in Claude Desktop.
3. **Given** a user has installed the extension, **When** they start a new Claude conversation, **Then** they can see and use Excel tools (e.g., create workbooks, read/write data).

---

### User Story 2 - Cross-Platform Windows Support (Priority: P2)

As a Windows user, I want the Desktop Extension to work on my system without requiring me to install .NET separately, so that installation remains truly one-click.

**Why this priority**: Windows is the primary (and only) platform for this MCP Server due to Excel COM interop requirements. The extension must bundle or reference the .NET runtime appropriately.

**Independent Test**: Can be tested by installing the extension on a clean Windows machine without .NET 8 pre-installed.

**Acceptance Scenarios**:

1. **Given** a Windows user without .NET 8 installed, **When** they install the extension, **Then** the extension either works with bundled runtime or provides clear guidance for .NET installation.
2. **Given** a Windows user with .NET 8 installed, **When** they install the extension, **Then** the extension uses the existing runtime and works correctly.

---

### User Story 3 - Extension Metadata and Discovery (Priority: P3)

As a Claude Desktop user browsing available extensions, I want to see clear information about what the Excel MCP Server does, so that I can decide whether to install it.

**Why this priority**: Good metadata helps users discover and understand the extension's capabilities, increasing adoption.

**Independent Test**: Can be tested by viewing the extension in Claude Desktop's extension directory (if submitted) or the installation dialog.

**Acceptance Scenarios**:

1. **Given** a user views the extension installation dialog, **When** reading the extension details, **Then** they see the extension name, description, author, homepage, and list of available tools.
2. **Given** the extension is submitted to the Claude extension directory, **When** a user searches for "Excel", **Then** the extension appears in search results with accurate metadata.

---

### User Story 4 - Documentation and Guidance (Priority: P4)

As a user discovering Excel MCP, I want clear documentation explaining all installation options (Desktop Extension, NuGet, dotnet tool), so that I can choose the method that best fits my workflow and technical expertise.

**Why this priority**: Clear documentation is essential for adoption. Users need to understand the differences between installation methods.

**Independent Test**: Can be tested by reading the documentation and installing the extension following the documented steps.

**Acceptance Scenarios**:

1. **Given** a user visits the README.md, **When** they look for installation instructions, **Then** they find clear guidance on using the Desktop Extension.
2. **Given** a user visits the GitHub Pages documentation, **When** they navigate to installation, **Then** they see a comparison of installation methods with the Desktop Extension highlighted as the easiest option.
3. **Given** a user has questions about the Desktop Extension, **When** they check the documentation, **Then** they find troubleshooting information and FAQs.

---

### Edge Cases

- What happens when the user tries to install on macOS or Linux? (Extension should indicate Windows-only compatibility and Claude Desktop should prevent installation)
- What happens if Excel is not installed on the user's Windows machine? (Extension installs successfully, but tools fail gracefully with clear error messages explaining Excel is required)
- What happens if the user already has a different version installed? (Claude Desktop handles update flow - newer version replaces older)
- What happens if the user wants to uninstall? (Standard Claude Desktop uninstall flow via Settings > Extensions)

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: System MUST create a valid `.mcpb` file (ZIP archive) containing the MCP Server binary, dependencies, and manifest.json
- **FR-002**: System MUST include a `manifest.json` following the MCPB specification (version 0.3 or later)
- **FR-003**: Manifest MUST declare `server.type = "binary"` since the MCP Server is a compiled .NET executable
- **FR-004**: Manifest MUST specify `platforms: ["win32"]` since the MCP Server only works on Windows (Excel COM interop requirement)
- **FR-005**: System MUST include the self-contained .NET publish output (all required .dll files and the .exe)
- **FR-006**: Manifest MUST declare all available tools with names and descriptions for user visibility
- **FR-007**: System MUST include appropriate metadata (name, display_name, version, description, author, repository, homepage)
- **FR-008**: System SHOULD include an icon for the extension (PNG format)
- **FR-009**: Build process MUST be automated (integrated into existing release workflow)
- **FR-010**: System MUST handle platform-specific executable naming (`.exe` suffix on Windows)
- **FR-011**: Manifest MUST use `mcp_config` to specify how Claude Desktop launches the server executable

### Documentation Requirements

- **DR-001**: README.md MUST be updated to document the Desktop Extension as an installation option
- **DR-002**: gh-pages/index.md MUST include Desktop Extension installation instructions
- **DR-003**: gh-pages/installation.md MUST document the one-click installation process
- **DR-004**: docs/INSTALLATION.md MUST be updated with Desktop Extension installation steps
- **DR-005**: FEATURES.md MUST mention Desktop Extension availability
- **DR-006**: Documentation MUST clearly indicate Windows-only platform support
- **DR-007**: Documentation MUST explain when to use Desktop Extension vs NuGet vs dotnet tool
- **DR-008**: GitHub Releases MUST include the `.mcpb` file as a downloadable asset
- **DR-009**: A new dedicated page or section SHOULD explain Desktop Extension features and troubleshooting
- **DR-010**: Documentation SHOULD include Anthropic extension directory submission instructions

### Anthropic Extension Directory Submission

To submit the Excel MCP Desktop Extension to Anthropic's official extension directory:

1. **Submission Form**: [Anthropic Extension Submission Form](https://docs.google.com/forms/d/e/1FAIpQLScHtjkiCNjpqnWtFLIQStChXlvVcvX8NPXkMfjtYPDPymgang/viewform)

2. **Prerequisites before submission**:
   - Valid `.mcpb` package that passes validation
   - Complete manifest.json with all required metadata
   - Extension tested and working on Windows with Claude Desktop
   - Public GitHub repository with documentation
   - `.mcpb` file available for download (GitHub Releases)

3. **Information to prepare**:
   - Extension name and display name
   - Description of capabilities
   - Author/maintainer contact
   - Repository and homepage URLs
   - Platform support (Windows only)
   - List of tools exposed by the extension

### Key Entities

- **MCPB Package**: The `.mcpb` ZIP archive containing all extension files
- **Manifest**: The `manifest.json` file describing the extension to Claude Desktop
- **Binary Server**: The compiled .NET MCP Server executable and its dependencies (self-contained publish)
- **Tools Metadata**: Declarations of Excel tools (excel_file, excel_range, excel_powerquery, etc.) exposed by the server

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: Users can install the extension in under 30 seconds (download + double-click + Install button)
- **SC-002**: Extension installs successfully on Windows 10/11 systems with Claude Desktop
- **SC-003**: All 11 Excel tools are visible and functional after installation
- **SC-004**: Extension manifest passes MCPB validation (`mcpb validate` command)
- **SC-005**: Package size is reasonable (target: under 100MB for self-contained .NET publish)
- **SC-006**: Zero configuration required from users - extension works immediately after installation
- **SC-007**: All documentation (README, GitHub Pages, installation guides) is updated and accurate
- **SC-008**: Users can find Desktop Extension installation instructions within 2 clicks from the main README

### Assumptions

- Claude Desktop is installed and updated to a version supporting Desktop Extensions (MCPB format)
- The MCP Server will be published as a self-contained .NET application (bundling the .NET runtime)
- The extension will be distributed via direct download initially (GitHub Releases), with potential submission to Claude's extension directory later
- Windows is the only supported platform due to Excel COM interop requirements
- The existing MCP Server stdio transport is compatible with Claude Desktop's extension execution model

### .NET Publish Configuration

The extension requires self-contained publishing to ensure one-click installation without .NET prerequisites:

**Recommended publish command**:
```bash
dotnet publish src/ExcelMcp.McpServer -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false
```

**Configuration options**:

| Option | Value | Rationale |
|--------|-------|-----------|
| `--self-contained true` | Required | Bundles .NET runtime - no pre-installed .NET needed |
| `-r win-x64` | Required | Windows 64-bit runtime identifier |
| `-p:PublishSingleFile=false` | Recommended | MCPB expects directory structure with executable |
| `-p:PublishTrimmed=true` | Optional | Reduces size by removing unused code (~40-50MB savings) |
| `-p:PublishReadyToRun=true` | Optional | Improves startup time via ahead-of-time compilation |

**Size considerations**:

| Configuration | Approximate Size |
|---------------|------------------|
| Framework-dependent | 5-10 MB |
| Self-contained | 60-80 MB |
| Self-contained + trimmed | 30-40 MB |
| Self-contained + trimmed + compressed (.mcpb) | 15-25 MB |

**Trimming notes**:
- Requires testing to ensure no reflection-based code breaks
- May need `<TrimmerRootAssembly>` entries for dynamically loaded types
- COM interop code should be verified after trimming
