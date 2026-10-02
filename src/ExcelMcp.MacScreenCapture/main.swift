import AppKit
import CoreGraphics
import Foundation
import ImageIO
import ScreenCaptureKit
import UniformTypeIdentifiers

private struct PixelRect: Codable {
    let x: Int
    let y: Int
    let width: Int
    let height: Int
}

private struct CaptureRequest: Codable {
    let version: Int
    let processId: pid_t
    let windowId: CGWindowID
    let capturedWidth: Int
    let capturedHeight: Int
    let cropPixels: PixelRect
    let quality: String
}

private struct CaptureError: Codable {
    let code: String
    let message: String
}

private struct CaptureResponse: Encodable {
    let version = 1
    let success: Bool
    let imageBase64: String
    let mimeType: String
    let width: Int
    let height: Int
    let error: CaptureError?
}

private enum HelperFailure: Error {
    case invalidRequest(String)
    case failed(String, String)

    var payload: CaptureError {
        switch self {
        case .invalidRequest(let message):
            CaptureError(code: "invalid_request", message: message)
        case .failed(let code, let message):
            CaptureError(code: code, message: message)
        }
    }
}

@main
private enum ExcelMcpScreenCapture {
    static func main() async {
        do {
            guard #available(macOS 14.0, *) else {
                throw HelperFailure.failed(
                    "macos_version_unsupported",
                    "ExcelMcp window capture requires macOS 14.0 or later.")
            }
            let request = try decodeRequest()
            let response = try await capture(request)
            write(response)
        } catch let failure as HelperFailure {
            write(CaptureResponse(
                success: false,
                imageBase64: "",
                mimeType: "",
                width: 0,
                height: 0,
                error: failure.payload))
        } catch {
            write(CaptureResponse(
                success: false,
                imageBase64: "",
                mimeType: "",
                width: 0,
                height: 0,
                error: CaptureError(code: "capture_failed", message: error.localizedDescription)))
        }
    }

    private static func decodeRequest() throws -> CaptureRequest {
        let data = FileHandle.standardInput.readDataToEndOfFile()
        let request: CaptureRequest
        do {
            request = try JSONDecoder().decode(CaptureRequest.self, from: data)
        } catch {
            throw HelperFailure.invalidRequest("Screenshot request is not valid protocol JSON.")
        }
        guard request.version == 1 else {
            throw HelperFailure.invalidRequest("Screenshot protocol version must be 1.")
        }
        guard request.processId > 0, request.windowId > 0 else {
            throw HelperFailure.invalidRequest("Excel processId and windowId must be positive.")
        }
        guard request.capturedWidth > 0, request.capturedHeight > 0 else {
            throw HelperFailure.invalidRequest("Captured window dimensions must be positive.")
        }
        let crop = request.cropPixels
        guard crop.x >= 0, crop.y >= 0,
              crop.x <= request.capturedWidth, crop.y <= request.capturedHeight,
              crop.width > 0, crop.height > 0,
              crop.width <= request.capturedWidth - crop.x,
              crop.height <= request.capturedHeight - crop.y else {
            throw HelperFailure.invalidRequest(
                "Screenshot crop must be a positive rectangle inside the captured Excel window.")
        }
        guard ["medium", "high", "low"].contains(request.quality.lowercased()) else {
            throw HelperFailure.invalidRequest("Screenshot quality must be medium, high, or low.")
        }
        return request
    }

    @available(macOS 14.0, *)
    private static func capture(_ request: CaptureRequest) async throws -> CaptureResponse {
        guard CGPreflightScreenCaptureAccess() else {
            throw HelperFailure.failed(
                "screen_recording_required",
                "Screen Recording permission is not ready; no consent was requested.")
        }

        let content: SCShareableContent
        do {
            content = try await SCShareableContent.excludingDesktopWindows(
                true,
                onScreenWindowsOnly: true)
        } catch {
            throw HelperFailure.failed(
                "window_inventory_failed",
                "ScreenCaptureKit could not enumerate shareable windows: \(error.localizedDescription)")
        }
        guard let window = content.windows.first(where: {
            $0.windowID == request.windowId
                && $0.owningApplication?.processID == request.processId
                && $0.owningApplication?.bundleIdentifier == "com.microsoft.Excel"
        }) else {
            throw HelperFailure.failed(
                "excel_window_not_found",
                "The exact Excel window ID and process ID pair is not currently shareable.")
        }

        let filter = SCContentFilter(desktopIndependentWindow: window)
        let configuration = SCStreamConfiguration()
        let pointPixelScale = CGFloat(filter.pointPixelScale)
        configuration.width = Int(filter.contentRect.width * pointPixelScale)
        configuration.height = Int(filter.contentRect.height * pointPixelScale)
        configuration.showsCursor = false
        configuration.ignoreShadowsSingleWindow = true
        guard configuration.width == request.capturedWidth,
              configuration.height == request.capturedHeight else {
            throw HelperFailure.failed(
                "window_geometry_changed",
                "The Excel window pixel dimensions changed after range geometry was prepared.")
        }

        let windowImage: CGImage
        do {
            windowImage = try await SCScreenshotManager.captureImage(
                contentFilter: filter,
                configuration: configuration)
        } catch {
            throw HelperFailure.failed(
                "capture_failed",
                "ScreenCaptureKit failed to capture the Excel window: \(error.localizedDescription)")
        }
        let crop = request.cropPixels
        guard let cropped = windowImage.cropping(to: CGRect(
            x: crop.x,
            y: crop.y,
            width: crop.width,
            height: crop.height)) else {
            throw HelperFailure.failed("crop_failed", "ScreenCaptureKit returned an invalid crop rectangle.")
        }

        let quality = request.quality.lowercased()
        let scale = quality == "low" ? 0.5 : quality == "medium" ? 0.75 : 1.0
        let output = scale == 1.0 ? cropped : try resize(cropped, scale: scale)
        let png = quality == "high"
        let data = try encode(output, png: png)
        return CaptureResponse(
            success: true,
            imageBase64: data.base64EncodedString(),
            mimeType: png ? "image/png" : "image/jpeg",
            width: output.width,
            height: output.height,
            error: nil)
    }

    private static func resize(_ image: CGImage, scale: Double) throws -> CGImage {
        let width = max(1, Int((Double(image.width) * scale).rounded()))
        let height = max(1, Int((Double(image.height) * scale).rounded()))
        guard let colorSpace = image.colorSpace ?? CGColorSpace(name: CGColorSpace.sRGB),
              let context = CGContext(
                data: nil,
                width: width,
                height: height,
                bitsPerComponent: 8,
                bytesPerRow: 0,
                space: colorSpace,
                bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue) else {
            throw HelperFailure.failed("resize_failed", "Could not allocate the screenshot resize buffer.")
        }
        context.interpolationQuality = .high
        context.draw(image, in: CGRect(x: 0, y: 0, width: width, height: height))
        guard let resized = context.makeImage() else {
            throw HelperFailure.failed("resize_failed", "Could not create the resized screenshot.")
        }
        return resized
    }

    private static func encode(_ image: CGImage, png: Bool) throws -> Data {
        let data = NSMutableData()
        let type = png ? UTType.png.identifier : UTType.jpeg.identifier
        guard let destination = CGImageDestinationCreateWithData(
            data,
            type as CFString,
            1,
            nil) else {
            throw HelperFailure.failed("encode_failed", "Could not create the screenshot encoder.")
        }
        let properties: CFDictionary? = png
            ? nil
            : [kCGImageDestinationLossyCompressionQuality: 0.85] as CFDictionary
        CGImageDestinationAddImage(destination, image, properties)
        guard CGImageDestinationFinalize(destination) else {
            throw HelperFailure.failed("encode_failed", "Could not encode the screenshot.")
        }
        return data as Data
    }

    private static func write(_ response: CaptureResponse) {
        do {
            let data = try JSONEncoder().encode(response)
            FileHandle.standardOutput.write(data)
            FileHandle.standardOutput.write(Data("\n".utf8))
        } catch {
            fputs("Could not encode screenshot response: \(error)\n", stderr)
            exit(1)
        }
    }
}
