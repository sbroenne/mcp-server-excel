ObjC.import("Foundation");
ObjC.import("CoreServices");

function run() {
    try {
        return checkPermission();
    } catch (error) {
        return JSON.stringify({
            implementation: "JXA",
            status: "InteropUnavailable",
            osStatus: null,
            requestedConsent: false,
            errorMessage: String(error.message || error)
        });
    }
}

function checkPermission() {
    const bundleId = $("com.microsoft.Excel");
    const target = Ref();
    const created = $.AECreateDesc(0x62756E64, bundleId.UTF8String, 19, target);
    if (created !== 0) {
        throw new Error("AECreateDesc failed with OSStatus " + created);
    }
    let status;
    let disposed;
    try {
        status = $.AEDeterminePermissionToAutomateTarget(
            target, 0x2A2A2A2A, 0x2A2A2A2A, false);
    } finally {
        disposed = $.AEDisposeDesc(target);
    }
    if (disposed !== 0) {
        throw new Error("AEDisposeDesc failed with OSStatus " + disposed);
    }
    const descriptions = {
        "0": "Allowed",
        "-1743": "Denied",
        "-1744": "ConsentRequired",
        "-600": "ExcelNotRunning"
    };
    return JSON.stringify({
        implementation: "JXA",
        status: descriptions[String(status)] || "Error",
        osStatus: status,
        requestedConsent: false
    });
}
