# Chart Style & Axis Type Safety

## Summary

Improve the excel_chart tool’s reliability and usability by introducing a typed, validated style identifier and clarifying axis targeting. Today, style is an untyped integer and axis selection can be ambiguous. This feature defines a constrained, human-readable style catalog and explicit axis targeting semantics so users can set styles and axis attributes confidently with clear validation and helpful errors.

## Goals

- Provide an explicit, curated list of built‑in chart styles as a first‑class option set.
- Prevent invalid style values and ambiguous axis targeting at the source.
- Maintain a clean, predictable user experience for assistants and scripts using excel_chart.
- Offer clear validation messages and workflow guidance when inputs are not valid.

## Non‑Goals

- No new chart rendering capabilities beyond selecting existing built‑in styles and targeting existing axes.
- No custom .crtx template management in this iteration.
- No changes to default chart creation behavior aside from optional style selection.

## Actors

- Primary: AI assistants and automation clients invoking the excel_chart tool.
- Secondary: Developers/operators reviewing logs and results for troubleshooting.

## User Scenarios & Testing

1) Set built‑in style at creation
   - Given a chart is created from a cell range, when the user specifies a valid style by name/value from the catalog, then the chart appears with that style applied.

2) Update style on existing chart
   - Given an existing chart, when the user sets a different valid style, then the chart updates to the new style and the result confirms the applied style.

3) Reject invalid style
   - Given a request with a style value not present in the catalog, when the tool validates parameters, then the request is rejected with a clear error describing acceptable values.

4) Target axis explicitly
   - Given a chart with primary and secondary axes, when the user sets an axis title specifying both axis dimension and lane (e.g., Primary Value), then only that axis is updated.

5) Errors guide next steps
   - Given any validation failure (style or axis), then the response includes a concise message and suggested next actions (e.g., list available styles; read chart details for existing axes).

## Functional Requirements

- FR‑1: The tool shall expose a finite catalog of built‑in chart styles (names + stable identifiers) that users can select from when setting style.
- FR‑2: The tool shall validate style input against this catalog and reject non‑catalog values with a descriptive error. Numeric style IDs are not accepted; only catalog values are valid.
- FR‑3: The tool shall support applying style both at creation time and for existing charts via a dedicated action.
- FR‑4: The tool shall provide explicit axis targeting via a single combined parameter that identifies both lane and dimension (e.g., PrimaryValue, SecondaryCategory). Separate lane/dimension parameters are not supported.
- FR‑5: The tool shall reject requests that specify an axis combination not present on the chart (e.g., secondary axis when none exists) with a descriptive error.
- FR‑6: The tool shall return results that include the effective style and axis settings after successful operations to enable verification.
- FR‑7: The tool shall maintain backward compatibility for existing flows where style is omitted (no style change applied implicitly).
- FR‑8: The tool’s error responses shall be JSON with success=false and clear messages; no exceptions for business errors.

## Success Criteria

- SC‑1: 100% of invalid style inputs are rejected before execution with a clear, single‑sentence message.
- SC‑2: 100% of valid style inputs result in the intended style applied; responses reflect the applied style.
- SC‑3: 100% of axis operations specify a clear target; ambiguous requests are rejected with guidance.
- SC‑4: Documentation and prompts list the complete style catalog and axis targeting semantics.
- SC‑5: In guided tests, users can set or change chart style in a single attempt ≥95% of the time.

## Assumptions

- Built‑in styles are a fixed, enumerable set exposed to users as a catalog of values.
- Axis targeting is expressed as a single parameter combining lane and dimension (PrimaryValue, SecondaryValue, PrimaryCategory, SecondaryCategory) for unambiguity.
- Existing clients that do not pass a style are not impacted.

## Dependencies

- Existing excel_chart actions for create/update must accept the new inputs and surface validation.
- Public documentation must be updated where excel_chart usage is described.

## Risks & Mitigations

- Risk: Users rely on numeric style IDs outside the catalog.
  - Mitigation: Provide a migration note and error with explicit acceptable values.
- Risk: Axis targeting differs across chart types.
  - Mitigation: Validate presence of targeted axis on the specific chart and error with guidance when unsupported.

## Out of Scope

- Managing or importing custom chart templates (.crtx).
- Auto‑creating secondary axes when missing.

## Decisions

- Style input: Catalog only; raw numeric IDs are not accepted.
- Style catalog scope: Built‑in styles only (no .crtx templates).
- Axis targeting: Single combined parameter (PrimaryValue, SecondaryValue, PrimaryCategory, SecondaryCategory).
