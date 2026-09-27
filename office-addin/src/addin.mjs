import { PROTOCOL_VERSION, REQUIREMENT_SETS } from "./constants.mjs";

const status = document.querySelector("#status");
const token = new URLSearchParams(location.hash.slice(1)).get("token");
history.replaceState(null, "", location.pathname);
const instanceId = crypto.randomUUID();

Office.onReady(async (info) => {
  try {
    if (!token) {
      throw new Error("The bridge activation token is missing.");
    }
    const workbookUrl = await getDocumentUrl();
    const requirementSets = REQUIREMENT_SETS.filter((version) =>
      Office.context.requirements.isSetSupported("ExcelApi", version));
    const diagnostics = Office.context.diagnostics;
    const session = await call("/v1/office/activate", {
      protocolVersion: PROTOCOL_VERSION,
      workbookUrl,
      instanceId,
      host: String(info.host),
      platform: String(info.platform),
      officeVersion: diagnostics.version,
      requirementSets
    });
    status.textContent = `Connected to ExcelMcp (ExcelApi ${requirementSets.at(-1) ?? "unavailable"}).`;
    await poll(session.sessionId, workbookUrl);
  } catch (error) {
    status.textContent = `ExcelMcp bridge unavailable: ${error.message}`;
  }
});

async function poll(sessionId, workbookUrl) {
  while (true) {
    const request = await call("/v1/office/next", { sessionId, workbookUrl, instanceId });
    if (request) {
      await call("/v1/office/results", {
        sessionId,
        workbookUrl,
        instanceId,
        requestId: request.requestId,
        success: true,
        value: {
          officeVersion: Office.context.diagnostics.version,
          requirementSets: REQUIREMENT_SETS.filter((version) =>
            Office.context.requirements.isSetSupported("ExcelApi", version))
        }
      });
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
}

async function getDocumentUrl() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFilePropertiesAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value?.url) {
        reject(new Error("Excel did not provide an exact saved workbook URL."));
        return;
      }
      resolve(result.value.url);
    });
  });
}

async function call(path, body) {
  const response = await fetch(path, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });
  const value = await response.json();
  if (!response.ok) {
    throw new Error(value.errorMessage ?? `Bridge request failed (${response.status}).`);
  }
  return value;
}
