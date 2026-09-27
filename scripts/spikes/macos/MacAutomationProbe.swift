import Foundation
import Carbon

let target = NSAppleEventDescriptor(bundleIdentifier: "com.microsoft.Excel")
let status = AEDeterminePermissionToAutomateTarget(
    target.aeDesc, AEEventClass(typeWildCard), AEEventID(typeWildCard), false)
let description: String
switch status {
case noErr: description = "Allowed"
case -1743: description = "Denied"
case -1744: description = "ConsentRequired"
case -600: description = "ExcelNotRunning"
default: description = "Error"
}
let result: [String: Any] = [
    "implementation": "Swift",
    "status": description,
    "osStatus": Int(status),
    "descriptorSize": MemoryLayout<AEDesc>.size,
    "requestedConsent": false
]
do {
    let data = try JSONSerialization.data(withJSONObject: result, options: [.sortedKeys])
    FileHandle.standardOutput.write(data)
    FileHandle.standardOutput.write(Data("\n".utf8))
    exit(status == noErr ? 0 : 2)
} catch {
    FileHandle.standardError.write(Data("Permission probe JSON serialization failed: \(error)\n".utf8))
    exit(1)
}
