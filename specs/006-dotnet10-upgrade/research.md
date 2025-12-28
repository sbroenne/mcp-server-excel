# Research: .NET 10 Framework Upgrade

**Branch**: `006-dotnet10-upgrade` | **Date**: 2025-01-12

## Overview

This document captures research findings for upgrading ExcelMcp from .NET 8 to .NET 10. Since .NET 10 is a standard Long-Term Support (LTS) release with well-documented migration paths, research requirements are minimal.

---

## R-001: .NET 10 GA SDK Version

**Question**: What is the correct SDK version for `global.json`?

**Decision**: Use SDK version `10.0.100`

**Rationale**: 
- .NET 10 GA released November 2024
- Version `10.0.100` is the initial GA SDK release
- `rollForward: latestFeature` allows automatic updates to patch releases

**Alternatives Considered**:
- Preview versions: Rejected per spec constraint "GA only, no preview"
- Specific patch version: Rejected in favor of `latestFeature` rollforward

**Source**: [.NET 10 Download](https://dotnet.microsoft.com/download/dotnet/10.0)

---

## R-002: NuGet Package Compatibility

**Question**: Do existing NuGet dependencies support .NET 10?

**Decision**: All packages compatible, no version changes needed

**Rationale**:
- `ModelContextProtocol` targets `netstandard2.0+` (compatible with all .NET versions)
- `Microsoft.Extensions.*` packages have .NET 10 specific builds
- Test packages (`xunit`, `Moq`) support all modern .NET versions

**Verification Steps**:
1. Run `dotnet restore` after updating target framework
2. Build should succeed without package errors
3. If issues, update to latest package versions

---

## R-003: Docker Base Images

**Question**: Are .NET 10 Docker images available?

**Decision**: Use official Microsoft container images

**Rationale**:
- Microsoft publishes images same day as GA release
- SDK image: `mcr.microsoft.com/dotnet/sdk:10.0`
- Runtime image: `mcr.microsoft.com/dotnet/runtime:10.0`

**Source**: [.NET Container Images](https://hub.docker.com/_/microsoft-dotnet)

---

## R-004: GitHub Actions Setup

**Question**: Does `actions/setup-dotnet` support .NET 10?

**Decision**: Use `dotnet-version: 10.0.x` with `actions/setup-dotnet@v4`

**Rationale**:
- GitHub Actions `setup-dotnet@v4` supports all released .NET versions
- Pattern `10.0.x` installs latest available patch version

**Source**: [setup-dotnet Action](https://github.com/actions/setup-dotnet)

---

## R-005: C# 14 Language Features

**Question**: What C# 14 features could improve the codebase?

**Decision**: Document in spec, implement as separate follow-up work

**Rationale**:
- Spec explicitly states: "No code changes required for upgrade"
- C# 14 features are optional enhancements, not upgrade requirements
- Features like `field` keyword, extension types, and improved `params` are beneficial but scope creep

**Features Identified** (for future consideration):
- `field` keyword in properties
- Extension types (replacing static extension methods)
- `params` with Span types
- Improved pattern matching
- Lock object improvements

**Source**: [C# 14 What's New](https://learn.microsoft.com/dotnet/csharp/whats-new/csharp-14)

---

## R-006: Breaking Changes

**Question**: Are there any .NET 10 breaking changes affecting ExcelMcp?

**Decision**: No blocking breaking changes identified

**Rationale**:
- ExcelMcp uses COM interop (unchanged between .NET versions)
- No deprecated APIs used that are removed in .NET 10
- `TreatWarningsAsErrors=true` will surface any issues during build

**Verification**: Build with .NET 10 and address any warnings/errors

**Source**: [.NET 10 Breaking Changes](https://learn.microsoft.com/dotnet/core/compatibility/10.0)

---

## Summary

| Research Item | Status | Action |
|---------------|--------|--------|
| SDK Version | ✅ Complete | Use `10.0.100` |
| NuGet Packages | ✅ Complete | No changes needed |
| Docker Images | ✅ Complete | Update to `:10.0` tags |
| GitHub Actions | ✅ Complete | Use `10.0.x` pattern |
| C# 14 Features | ✅ Complete | Document for future, out of scope |
| Breaking Changes | ✅ Complete | None affecting ExcelMcp |

**All research items resolved. Proceed to implementation.**
