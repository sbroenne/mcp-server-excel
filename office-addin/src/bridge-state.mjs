import { randomUUID } from "node:crypto";
import {
  MAX_REQUEST_TIMEOUT_MS,
  PROTOCOL_VERSION,
  SUPPORTED_ACTIONS
} from "./constants.mjs";

function requireString(value, name, maxLength = 4096) {
  if (typeof value !== "string" || value.length === 0 || value.length > maxLength) {
    throw new TypeError(`${name} must be a non-empty string no longer than ${maxLength} characters.`);
  }
  return value;
}

function normalizeWorkbookUrl(value) {
  const text = requireString(value, "workbookUrl");
  const url = new URL(text);
  if (url.protocol !== "file:") {
    throw new TypeError("workbookUrl must be an exact file URL.");
  }
  url.hash = "";
  return url.href;
}

export class OfficeBridgeState {
  #clock;
  #sessions = new Map();
  #requests = new Map();

  constructor(clock = () => Date.now()) {
    this.#clock = clock;
  }

  registerSession(input) {
    const sessionId = requireString(input.sessionId, "sessionId", 128);
    const workbookUrl = normalizeWorkbookUrl(input.workbookUrl);
    const current = this.#sessions.get(sessionId);
    if (current) {
      if (current.workbookUrl !== workbookUrl) {
        throw new Error("A bridge session cannot be rebound to another workbook.");
      }
      return this.describeSession(current);
    }
    const existing = [...this.#sessions.values()]
      .find((session) => session.workbookUrl === workbookUrl && session.sessionId !== sessionId);
    if (existing) {
      throw new Error("The workbook is already bound to another bridge session.");
    }

    const session = {
      sessionId,
      workbookUrl,
      instanceId: null,
      runtime: null,
      queue: [],
      activeRequestId: null
    };
    this.#sessions.set(sessionId, session);
    return this.describeSession(session);
  }

  activate(input) {
    const workbookUrl = normalizeWorkbookUrl(input.workbookUrl);
    const instanceId = requireString(input.instanceId, "instanceId", 128);
    const session = [...this.#sessions.values()]
      .find((candidate) => candidate.workbookUrl === workbookUrl);
    if (!session) {
      throw new Error("No ExcelMcp session is registered for this exact workbook.");
    }

    session.instanceId = instanceId;
    session.runtime = {
      protocolVersion: Number(input.protocolVersion),
      host: requireString(input.host, "host", 32),
      platform: requireString(input.platform, "platform", 32),
      officeVersion: requireString(input.officeVersion, "officeVersion", 64),
      requirementSets: validateRequirementSets(input.requirementSets)
    };
    if (session.runtime.protocolVersion !== PROTOCOL_VERSION) {
      session.runtime = null;
      session.instanceId = null;
      throw new Error(`Unsupported bridge protocol ${input.protocolVersion}.`);
    }
    if (session.runtime.host !== "Excel") {
      session.runtime = null;
      session.instanceId = null;
      throw new Error("The Office.js runtime is not hosted by Excel.");
    }
    return this.describeSession(session);
  }

  createRequest(input) {
    this.expireRequests();
    const session = this.#requireSession(input.sessionId);
    if (!session.runtime) {
      throw new Error("The Office.js add-in is not active for this workbook.");
    }
    const action = requireString(input.action, "action", 128);
    if (!SUPPORTED_ACTIONS.includes(action)) {
      throw new Error(`Office.js action '${action}' is not enabled.`);
    }
    const timeoutMs = Number(input.timeoutMs);
    if (!Number.isInteger(timeoutMs) || timeoutMs < 1 || timeoutMs > MAX_REQUEST_TIMEOUT_MS) {
      throw new TypeError(`timeoutMs must be between 1 and ${MAX_REQUEST_TIMEOUT_MS}.`);
    }

    const request = {
      requestId: randomUUID(),
      sessionId: session.sessionId,
      workbookUrl: session.workbookUrl,
      action,
      payload: validatePayload(action, input.payload),
      createdAt: this.#clock(),
      deadlineAt: this.#clock() + timeoutMs,
      status: "queued",
      result: null
    };
    this.#requests.set(request.requestId, request);
    session.queue.push(request.requestId);
    return { ...request };
  }

  takeNext(input) {
    this.expireRequests();
    const session = this.#requireBoundOfficeSession(input);
    if (session.activeRequestId) {
      return null;
    }
    while (session.queue.length > 0) {
      const request = this.#requests.get(session.queue.shift());
      if (request?.status === "queued") {
        request.status = "active";
        session.activeRequestId = request.requestId;
        return { ...request };
      }
    }
    return null;
  }

  complete(input) {
    this.expireRequests();
    const session = this.#requireBoundOfficeSession(input);
    const requestId = requireString(input.requestId, "requestId", 128);
    const request = this.#requests.get(requestId);
    if (!request || request.sessionId !== session.sessionId
        || session.activeRequestId !== requestId || request.status !== "active") {
      throw new Error("Request correlation or workbook binding failed.");
    }
    request.status = input.success === true ? "completed" : "failed";
    request.result = input.success === true
      ? { success: true, value: input.value ?? null, errorMessage: null }
      : {
          success: false,
          value: null,
          errorMessage: requireString(input.errorMessage, "errorMessage", 4096)
        };
    session.activeRequestId = null;
    return { ...request };
  }

  cancel(requestId) {
    const request = this.#requests.get(requireString(requestId, "requestId", 128));
    if (!request || request.status === "completed" || request.status === "failed") {
      return false;
    }
    request.status = "cancelled";
    const session = this.#sessions.get(request.sessionId);
    if (session?.activeRequestId === request.requestId) {
      session.activeRequestId = null;
    }
    return true;
  }

  expireRequests() {
    const now = this.#clock();
    for (const request of this.#requests.values()) {
      if ((request.status === "queued" || request.status === "active")
          && request.deadlineAt <= now) {
        request.status = "expired";
        const session = this.#sessions.get(request.sessionId);
        if (session?.activeRequestId === request.requestId) {
          session.activeRequestId = null;
        }
      }
    }
  }

  getRequest(requestId) {
    this.expireRequests();
    const request = this.#requests.get(requestId);
    return request ? { ...request } : null;
  }

  describeSession(session) {
    return {
      sessionId: session.sessionId,
      workbookUrl: session.workbookUrl,
      available: session.runtime !== null,
      runtime: session.runtime,
      enabledActions: [...SUPPORTED_ACTIONS]
    };
  }

  #requireSession(sessionId) {
    const session = this.#sessions.get(requireString(sessionId, "sessionId", 128));
    if (!session) {
      throw new Error("Bridge session is not registered.");
    }
    return session;
  }

  #requireBoundOfficeSession(input) {
    const session = this.#requireSession(input.sessionId);
    if (session.instanceId !== requireString(input.instanceId, "instanceId", 128)
        || session.workbookUrl !== normalizeWorkbookUrl(input.workbookUrl)) {
      throw new Error("Office.js runtime is not bound to this exact workbook session.");
    }
    return session;
  }
}

function validateRequirementSets(value) {
  if (!Array.isArray(value) || value.some((item) => typeof item !== "string")) {
    throw new TypeError("requirementSets must be an array of strings.");
  }
  return [...new Set(value)].sort((left, right) =>
    left.localeCompare(right, undefined, { numeric: true }));
}

function validatePayload(action, value) {
  const payload = value ?? {};
  if (Array.isArray(payload) || typeof payload !== "object") {
    throw new TypeError("payload must be a JSON object.");
  }
  if (action === "bridge.health" && Object.keys(payload).length !== 0) {
    throw new TypeError("bridge.health does not accept payload properties.");
  }
  return payload;
}
