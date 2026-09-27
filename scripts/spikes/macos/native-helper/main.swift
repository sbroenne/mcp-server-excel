import Cocoa
import Carbon
import CryptoKit

enum ProbeFailure: Error, CustomStringConvertible {
    case failed(String)
    var description: String {
        switch self { case .failed(let message): return message }
    }
}

final class Probe: NSObject, NSApplicationDelegate {
    let mode: String
    let output: URL
    var watchdog: DispatchSourceTimer?
    var script: NSAppleScript?
    var timedOut = false
    var checks: [String] = []

    init(mode: String, output: URL) {
        self.mode = mode
        self.output = output
    }

    func finish(_ result: [String: Any], code: Int32) -> Never {
        var report = result
        report["mode"] = mode
        report["bundleIdentifier"] = Bundle.main.bundleIdentifier ?? ""
        report["pid"] = ProcessInfo.processInfo.processIdentifier
        do {
            let data = try JSONSerialization.data(withJSONObject: report, options: [.prettyPrinted, .sortedKeys])
            try data.write(to: output, options: .atomic)
        } catch {
            fputs("Could not write probe result: \(error)\n", stderr)
            exit(1)
        }
        exit(code)
    }

    func applicationDidFinishLaunching(_ notification: Notification) {
        let timer = DispatchSource.makeTimerSource(queue: .global())
        timer.schedule(deadline: .now() + 120)
        timer.setEventHandler { [self] in
            finish([
                "success": false,
                "errorCategory": "Timeout",
                "errorMessage": "Helper exceeded 120 seconds. Excel was not stopped; inspect owned fixtures before retrying."
            ], code: 124)
        }
        watchdog = timer
        timer.resume()
        DispatchQueue.global().async { [self] in
            let target = NSAppleEventDescriptor(bundleIdentifier: "com.microsoft.Excel")
            let status = AEDeterminePermissionToAutomateTarget(
                target.aeDesc, AEEventClass(typeWildCard), AEEventID(typeWildCard), mode == "setup")
            if mode == "check" {
                finish([
                    "success": status == noErr,
                    "osStatus": Int(status),
                    "errorMessage": status == noErr ? "" : "Automation permission is not ready; no consent was requested.",
                    "requestedConsent": false,
                    "scope": "Automation permission only. No protected file access."
                ], code: status == noErr ? 0 : 2)
            }
            guard status == noErr else {
                finish([
                    "success": false,
                    "osStatus": Int(status),
                    "errorCategory": "Prerequisite",
                    "errorMessage": "Excel must be running and this helper must have Automation permission."
                ], code: 2)
            }
            DispatchQueue.main.async { [self] in
                do { try runFixtures() }
                catch {
                    finish([
                        "success": false,
                        "errorCategory": timedOut ? "Timeout" : "ProbeFailure",
                        "errorMessage": String(describing: error),
                        "checksCompleted": checks
                    ], code: 1)
                }
            }
        }
    }

    func invoke(_ action: String, _ path: URL) throws -> [String: Any] {
        guard let script else { throw ProbeFailure.failed("Probe script was not loaded.") }
        let event = NSAppleEventDescriptor(
            eventClass: AEEventClass(kASAppleScriptSuite),
            eventID: AEEventID(kASSubroutineEvent),
            targetDescriptor: nil,
            returnID: AEReturnID(kAutoGenerateReturnID),
            transactionID: AETransactionID(kAnyTransactionID))
        event.setParam(NSAppleEventDescriptor(string: "invokeprobe"), forKeyword: AEKeyword(keyASSubroutineName))
        let args = NSAppleEventDescriptor.list()
        args.insert(NSAppleEventDescriptor(string: action), at: 1)
        args.insert(NSAppleEventDescriptor(string: path.path), at: 2)
        event.setParam(args, forKeyword: AEKeyword(keyDirectObject))
        var error: NSDictionary?
        let result = script.executeAppleEvent(event, error: &error)
        if let error {
            let code = (error[NSAppleScript.errorNumber] as? NSNumber)?.intValue ?? 0
            if code == -1712 { timedOut = true }
            if action == "missing-sheet" && code == 9006 {
                return ["errorCode": code]
            }
            throw ProbeFailure.failed("AppleScript \(action) failed (\(code)): \(error)")
        }
        if action == "missing-sheet" {
            throw ProbeFailure.failed("Missing worksheet unexpectedly succeeded.")
        }
        guard let text = result.stringValue,
              let data = text.data(using: .utf8),
              let object = try JSONSerialization.jsonObject(with: data) as? [String: Any] else {
            throw ProbeFailure.failed("AppleScript \(action) returned invalid JSON.")
        }
        if let success = object["success"] as? Bool,
           !success || (object["errorMessage"] as? String) != "" {
            throw ProbeFailure.failed("AppleScript \(action) violated the success/error contract.")
        }
        return object
    }

    func assertData(_ data: [String: Any]) throws {
        guard let values = data["values"] as? [[Any]],
              values.count == 3, values.allSatisfy({ $0.count == 2 }),
              values[0][0] as? String == "Item", values[0][1] as? String == "Amount",
              values[1][0] as? String == "Alpha", values[1][1] as? Double == 10,
              values[2][0] as? String == "Beta", values[2][1] as? Double == 20,
              let calculated = data["calculated"] as? [[Double]], calculated == [[30]],
              let formulas = data["formulas"] as? [[String]], formulas == [["=SUM(B2:B3)"]],
              data["numberFormat"] as? String == "0.00",
              data["text"] as? String == "quote \" slash \\ newline\n\u{03A9}" else {
            throw ProbeFailure.failed("Workbook values, shapes, formulas, text or formatting did not match.")
        }
    }

    func runFixtures() throws {
        guard let identifier = Bundle.main.bundleIdentifier,
              NSRunningApplication.runningApplications(withBundleIdentifier: identifier).count == 1 else {
            throw ProbeFailure.failed("Another helper instance is running. Mac Excel tests must be serialized.")
        }
        let fm = FileManager.default
        let support = fm.homeDirectoryForCurrentUser.appendingPathComponent(
            "Library/Application Support/ExcelMcpMacPermissionProbe", isDirectory: true)
        let receipt = support.appendingPathComponent("setup.json")
        guard let executable = Bundle.main.executableURL else {
            throw ProbeFailure.failed("Helper executable identity is unavailable.")
        }
        let fingerprint = SHA256.hash(data: try Data(contentsOf: executable)).map { String(format: "%02x", $0) }.joined()
        let identity = [
            "binary": fingerprint,
            "bundle": Bundle.main.bundleURL.path,
            "os": ProcessInfo.processInfo.operatingSystemVersionString
        ]
        if mode == "run" {
            guard fm.fileExists(atPath: receipt.path),
                  let saved = try JSONSerialization.jsonObject(with: Data(contentsOf: receipt)) as? [String: String],
                  saved == identity else {
                throw ProbeFailure.failed("Explicit setup is missing or belongs to a different helper build, path or OS.")
            }
        } else if fm.fileExists(atPath: receipt.path) {
            try fm.removeItem(at: receipt)
        }
        guard let source = Bundle.main.url(forResource: "ExcelSpike", withExtension: "applescript") else {
            throw ProbeFailure.failed("Bundled AppleScript is missing.")
        }
        script = NSAppleScript(source: try String(contentsOf: source, encoding: .utf8))
        guard script != nil else { throw ProbeFailure.failed("Could not create the AppleScript instance.") }

        let root = fm.homeDirectoryForCurrentUser.appendingPathComponent(
            "Library/Containers/com.microsoft.Excel/Data/Documents", isDirectory: true)
        let runId = UUID().uuidString.lowercased()
        let directory = root.appendingPathComponent("excelmcp-native-probe-\(runId)", isDirectory: true)
        try fm.createDirectory(at: directory, withIntermediateDirectories: false)
        let main = directory.appendingPathComponent("main \(runId).xlsx")
        let sentinel = directory.appendingPathComponent("sentinel \(runId).xlsx")
        let owned = [main, sentinel]
        var failure: Error?
        var version: String = ""
        do {
            version = try invoke("version", main)["version"] as? String ?? ""
            guard !version.isEmpty else { throw ProbeFailure.failed("Excel version was empty.") }
            _ = try invoke("create", main)
            _ = try invoke("create", sentinel)
            checks.append("create-save")
            try assertData(invoke("read", main))
            checks.append(contentsOf: ["explicit-workbook-targeting", "2d-values", "1x1-formula-result", "text-round-trip", "number-format"])
            _ = try invoke("missing-sheet", main)
            checks.append("missing-sheet-error")
            try assertData(invoke("read", main))
            checks.append("read-after-error")
            let bulk = try invoke("bulk", main)
            guard let matrix = bulk["values"] as? [[Double]], matrix.count == 1000,
                  bulk["total"] as? Double == 5_005_000 else {
                throw ProbeFailure.failed("Bulk shape or calculation did not match.")
            }
            for row in 0..<1000 {
                guard matrix[row].count == 10 else { throw ProbeFailure.failed("Bulk column count did not match.") }
                for column in 0..<10 where matrix[row][column] != Double((row + 1) * (column + 1)) {
                    throw ProbeFailure.failed("Bulk value did not match.")
                }
            }
            checks.append("10000-cell-values-and-calculation")
            _ = try invoke("discard", main)
            try assertData(invoke("read", sentinel))
            guard fm.fileExists(atPath: main.path) else { throw ProbeFailure.failed("Saved workbook is missing.") }
            _ = try invoke("open", main)
            try assertData(invoke("read", main))
            checks.append(contentsOf: ["discard-unsaved-changes", "save-reopen"])
            _ = try invoke("close", main)
            try assertData(invoke("read", sentinel))
            checks.append("sentinel-preserved")
        } catch { failure = error }
        if timedOut {
            throw failure ?? ProbeFailure.failed("Timed out. Retained synthetic fixtures; Excel was not stopped.")
        }
        for path in owned {
            do { _ = try invoke("cleanup", path) }
            catch { throw ProbeFailure.failed("Initial failure: \(String(describing: failure)); cleanup failed: \(error). Fixtures retained.") }
        }
        if let failure { throw failure }
        let entries = try fm.contentsOfDirectory(at: directory, includingPropertiesForKeys: [.isSymbolicLinkKey])
        guard Set(entries.map(\.lastPathComponent)) == Set(owned.map(\.lastPathComponent)),
              try directory.resourceValues(forKeys: [.isSymbolicLinkKey]).isSymbolicLink != true else {
            throw ProbeFailure.failed("Unexpected fixture entries or linked directory; refusing cleanup.")
        }
        for path in owned { try fm.removeItem(at: path) }
        // rmdir never recursively removes unexpected contents.
        guard directory.path.withCString({ rmdir($0) }) == 0 else {
            throw ProbeFailure.failed("Could not remove the empty fixture directory.")
        }
        checks.append("owned-cleanup")
        if mode == "setup" {
            try fm.createDirectory(at: support, withIntermediateDirectories: true)
            try JSONSerialization.data(withJSONObject: identity).write(to: receipt, options: .atomic)
        }
        finish([
            "success": true,
            "errorMessage": "",
            "excelVersion": version,
            "bulkCellsVerified": 10_000,
            "checks": checks,
            "scope": "Standalone native helper. Not production MCP/CLI E2E."
        ], code: 0)
    }
}

let args = CommandLine.arguments
guard args.count == 5, args[1] == "--mode", ["check", "setup", "run"].contains(args[2]),
      args[3] == "--output", args[4].hasPrefix("/") else {
    fputs("Expected --mode check|setup|run --output absolute-result-path\n", stderr)
    exit(2)
}
let app = NSApplication.shared
let delegate = Probe(mode: args[2], output: URL(fileURLWithPath: args[4]))
app.delegate = delegate
app.setActivationPolicy(.accessory)
app.run()
