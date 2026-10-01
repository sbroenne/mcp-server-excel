import { executeOfficeAction, negotiateRequirementSets } from "./actions.mjs";
import { PROTOCOL_VERSION } from "./constants.mjs";

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
    const requirementSets = negotiateRequirementSets(Office.context.requirements);
    const diagnostics = Office.context.diagnostics;
    const activation = {
      protocolVersion: PROTOCOL_VERSION,
      workbookUrl,
      instanceId,
      host: String(info.host),
      platform: String(info.platform),
      officeVersion: diagnostics.version,
      requirementSets: requirementSets.excelApi,
      desktopRequirementSets: requirementSets.excelApiDesktop
    };
    const session = await activateWhenRegistered(activation);
    status.textContent =
      `Connected to ExcelMcp (ExcelApi ${requirementSets.excelApi.at(-1) ?? "unavailable"}, ` +
      `ExcelApiDesktop ${requirementSets.excelApiDesktop.at(-1) ?? "unavailable"}).`;
    await poll(session.sessionId, workbookUrl, requirementSets, session.enabledActions);
  } catch (error) {
    status.textContent = `ExcelMcp bridge unavailable: ${error.message}`;
  }
});

async function activateWhenRegistered(activation) {
  while (true) {
    try {
      return await call("/v1/office/activate", activation);
    } catch (error) {
      if (error.message !== "No ExcelMcp session is registered for this exact workbook.") {
        throw error;
      }
      status.textContent = "Waiting for ExcelMcp to open this exact workbook.";
      await new Promise((resolve) => setTimeout(resolve, 500));
    }
  }
}

async function poll(sessionId, workbookUrl, requirementSets, enabledActions) {
  while (true) {
    const request = await call("/v1/office/next", { sessionId, workbookUrl, instanceId });
    if (request) {
      let result;
      try {
        result = {
          success: true,
          value: request.action === "bridge.health"
            ? {
                success: true,
                errorMessage: null,
                officeVersion: Office.context.diagnostics.version,
                requirementSets,
                enabledActions
              }
            : await executeOfficeAction(request, {
                requirementSets,
                run: Excel.run
              })
        };
      } catch (error) {
        result = {
          success: false,
          errorMessage: error.message
        };
      }
      await call("/v1/office/results", {
        sessionId,
        workbookUrl,
        instanceId,
        requestId: request.requestId,
        ...result
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
