import { ADDIN_ID, ADDIN_VERSION } from "./constants.mjs";

export function createManifest({ port, token }) {
  const source = `https://localhost:${port}/taskpane.html#token=${encodeURIComponent(token)}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="TaskPaneApp">
  <Id>${ADDIN_ID}</Id>
  <Version>${ADDIN_VERSION}</Version>
  <ProviderName>ExcelMcp</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="ExcelMcp capability bridge"/>
  <Description DefaultValue="Optional capability negotiation for ExcelMcp on macOS."/>
  <IconUrl DefaultValue="https://localhost:${port}/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:${port}/icon-64.png"/>
  <SupportUrl DefaultValue="https://github.com/sbroenne/mcp-server-excel"/>
  <AppDomains>
    <AppDomain>https://localhost:${port}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="ExcelApi" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${source}"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
`;
}
