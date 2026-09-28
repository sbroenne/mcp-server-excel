Attribute VB_Name = "ExcelMcpHelper"
Option Explicit

Private Const HELPER_VERSION As String = "1.3.0"
Private Const PROTOCOL_VERSION As Long = 1
Private Const MAX_PAYLOAD_BYTES As Long = 262144
Private Const MAX_SAFE_ERROR_DETAIL_CHARS As Long = 512
Private Const STANDARD_MODULE_TYPE As Long = 1
Private Const PROJECT_LOCKED As Long = 1
Private mSafeErrorDetail As String

Private Type QueryLoadSnapshot
    HasWorksheetLoad As Boolean
    SheetName As String
    CellAddress As String
End Type

Public Function ExcelMcpDispatch(ByVal requestJson As String) As String
    mSafeErrorDetail = vbNullString
    Dim requestId As String
    On Error GoTo DispatchError

    If Utf8ByteCount(requestJson) > MAX_PAYLOAD_BYTES Then
        ExcelMcpDispatch = ErrorResponse("", "InvalidInput", "request_too_large", _
            "The helper request exceeds the 262144-byte UTF-8 transport limit.")
        Exit Function
    End If

    ValidateRequestEnvelope requestJson
    requestId = JsonRequiredString(requestJson, "requestId")
    If Not IsCanonicalRequestId(requestId) Then
        Err.Raise vbObjectError + 7000, "ExcelMcpHelper", "invalid_request_id"
    End If

    Dim version As Long
    version = JsonRequiredLong(requestJson, "version")
    If version <> PROTOCOL_VERSION Then
        Err.Raise vbObjectError + 7001, "ExcelMcpHelper", "unsupported_version"
    End If

    Dim workbookPath As String
    workbookPath = JsonRequiredString(requestJson, "workbookPath")
    Dim target As Workbook
    Set target = WorkbookByExactFullName(workbookPath)

    Dim action As String
    action = JsonRequiredString(requestJson, "action")
    If StrComp(action, "helper.capabilities", vbBinaryCompare) <> 0 And _
       StrComp(CStr(target.FullName), CStr(ThisWorkbook.FullName), vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 7004, "ExcelMcpHelper", "helper_target_forbidden"
    End If
    Dim argumentsJson As String
    argumentsJson = JsonRequiredObject(requestJson, "arguments")

    Dim resultJson As String
    Select Case action
        Case "helper.capabilities"
            ValidateArgumentKeys argumentsJson, ""
            resultJson = HelperCapabilities(target)
        Case "helper.inspect-engines"
            ValidateArgumentKeys argumentsJson, ""
            resultJson = InspectOptionalEngines(target)
        Case "powerquery.list"
            ValidateArgumentKeys argumentsJson, ""
            resultJson = PowerQueryList(target)
        Case "powerquery.view"
            ValidateArgumentKeys argumentsJson, "name"
            resultJson = PowerQueryView(target, JsonRequiredString(argumentsJson, "name"))
        Case "powerquery.create"
            ValidateArgumentKeys argumentsJson, "name,formula,destination,sheetName,cellAddress"
            resultJson = PowerQueryCreate(target, _
                JsonRequiredString(argumentsJson, "name"), _
                JsonRequiredString(argumentsJson, "formula"), _
                JsonRequiredString(argumentsJson, "destination"), _
                JsonOptionalString(argumentsJson, "sheetName"), _
                JsonOptionalString(argumentsJson, "cellAddress"))
        Case "powerquery.update"
            ValidateArgumentKeys argumentsJson, "name,formula,refresh"
            resultJson = PowerQueryUpdate(target, _
                JsonRequiredString(argumentsJson, "name"), _
                JsonRequiredString(argumentsJson, "formula"), _
                JsonRequiredBoolean(argumentsJson, "refresh"))
        Case "powerquery.rename"
            ValidateArgumentKeys argumentsJson, "name,newName"
            resultJson = PowerQueryRename(target, _
                JsonRequiredString(argumentsJson, "name"), _
                JsonRequiredString(argumentsJson, "newName"))
        Case "powerquery.delete"
            ValidateArgumentKeys argumentsJson, "name,deleteConnection"
            resultJson = PowerQueryDelete(target, _
                JsonRequiredString(argumentsJson, "name"), _
                JsonRequiredBoolean(argumentsJson, "deleteConnection"))
        Case "powerquery.refresh"
            ValidateArgumentKeys argumentsJson, "name"
            resultJson = PowerQueryRefresh(target, JsonRequiredString(argumentsJson, "name"))
        Case "powerquery.refresh-all"
            ValidateArgumentKeys argumentsJson, ""
            resultJson = PowerQueryRefreshAll(target)
        Case "powerquery.load-to"
            ValidateArgumentKeys argumentsJson, "name,destination,sheetName,cellAddress"
            resultJson = PowerQueryLoadTo(target, _
                JsonRequiredString(argumentsJson, "name"), _
                JsonRequiredString(argumentsJson, "destination"), _
                JsonOptionalString(argumentsJson, "sheetName"), _
                JsonOptionalString(argumentsJson, "cellAddress"))
        Case "powerquery.unload"
            ValidateArgumentKeys argumentsJson, "name"
            resultJson = PowerQueryUnload(target, JsonRequiredString(argumentsJson, "name"))
        Case "powerquery.evaluate"
            ValidateArgumentKeys argumentsJson, "formula"
            resultJson = PowerQueryEvaluate(target, _
                JsonRequiredString(argumentsJson, "formula"), requestId)
        Case "analysis.create-scenario"
            ValidateArgumentKeys argumentsJson, _
                "sheetName,scenarioName,changingCells,values,comment,locked,hidden"
            resultJson = AnalysisCreateScenario(target, _
                JsonRequiredString(argumentsJson, "sheetName"), _
                JsonRequiredString(argumentsJson, "scenarioName"), _
                JsonRequiredString(argumentsJson, "changingCells"), _
                JsonRequiredArrayValues(argumentsJson, "values"), _
                JsonOptionalString(argumentsJson, "comment"), _
                JsonRequiredBoolean(argumentsJson, "locked"), _
                JsonRequiredBoolean(argumentsJson, "hidden"))
        Case "analysis.show-scenario"
            ValidateArgumentKeys argumentsJson, "sheetName,scenarioName"
            resultJson = AnalysisShowScenario(target, _
                JsonRequiredString(argumentsJson, "sheetName"), _
                JsonRequiredString(argumentsJson, "scenarioName"))
        Case "vba.list"
            ValidateArgumentKeys argumentsJson, ""
            resultJson = VbaList(target)
        Case "vba.view"
            ValidateArgumentKeys argumentsJson, "moduleName"
            resultJson = VbaView(target, JsonRequiredString(argumentsJson, "moduleName"))
        Case "vba.import"
            ValidateArgumentKeys argumentsJson, "moduleName,source"
            resultJson = VbaImport(target, _
                JsonRequiredString(argumentsJson, "moduleName"), _
                JsonRequiredString(argumentsJson, "source"))
        Case "vba.update"
            ValidateArgumentKeys argumentsJson, "moduleName,source"
            resultJson = VbaUpdate(target, _
                JsonRequiredString(argumentsJson, "moduleName"), _
                JsonRequiredString(argumentsJson, "source"))
        Case "vba.delete"
            ValidateArgumentKeys argumentsJson, "moduleName"
            resultJson = VbaDelete(target, JsonRequiredString(argumentsJson, "moduleName"))
        Case "vba.run"
            ValidateArgumentKeys argumentsJson, "procedureName,parameters"
            resultJson = VbaRun(target, _
                JsonRequiredString(argumentsJson, "procedureName"), _
                JsonRequiredArrayValues(argumentsJson, "parameters"))
        Case Else
            Err.Raise vbObjectError + 7002, "ExcelMcpHelper", "unsupported_action"
    End Select

    ExcelMcpDispatch = SuccessResponse(requestId, resultJson)
    If Utf8ByteCount(ExcelMcpDispatch) > MAX_PAYLOAD_BYTES Then
        ExcelMcpDispatch = ErrorResponse(requestId, "InvalidInput", "response_too_large", _
            "The helper response exceeds the 262144-byte UTF-8 transport limit.")
    End If
    Exit Function

DispatchError:
    Dim errorCode As String
    Dim errorCategory As String
    ClassifyError Err.Number, Err.Description, errorCategory, errorCode
    ExcelMcpDispatch = ErrorResponse(requestId, errorCategory, errorCode, _
        SafeErrorMessage(errorCode))
End Function

Private Function HelperCapabilities(ByVal target As Workbook) As String
    Dim queryReady As Boolean
    Dim projectReady As Boolean
    Dim ignoredCount As Long

    On Error Resume Next
    ignoredCount = target.Queries.Count
    queryReady = (Err.Number = 0)
    Err.Clear
    ignoredCount = target.VBProject.VBComponents.Count
    projectReady = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    Dim output As String
    output = "{" & _
        """helperVersion"":" & JsonQuote(HELPER_VERSION) & "," & _
        """protocolVersion"":" & CStr(PROTOCOL_VERSION) & "," & _
        """staticAvailability"":{" & _
            """queriesApi"":true," & _
            """queryTableApi"":true," & _
            """scenarioApi"":true," & _
            """vbProjectApi"":true," & _
            """codeModuleApi"":true},"
    output = output & """engineCapabilities"":{" & _
            """xmlMapsApi"":null," & _
            """rangeXPathApi"":null," & _
            """workbookModelApi"":null," & _
            """dataModelConnectionApi"":null},"
    output = output & """supportedActions"":[" & _
            """helper.capabilities"",""helper.inspect-engines""," & _
            """powerquery.list"",""powerquery.view"",""powerquery.create""," & _
            """powerquery.update"",""powerquery.rename"",""powerquery.delete""," & _
            """powerquery.refresh"",""powerquery.refresh-all""," & _
            """powerquery.load-to"",""powerquery.unload"",""powerquery.evaluate""," & _
            """analysis.create-scenario"",""analysis.show-scenario""," & _
            """vba.list"",""vba.view"",""vba.import"",""vba.update"",""vba.delete"",""vba.run""],"
    output = output & """trustReadiness"":{" & _
            """powerQueryReadable"":" & JsonBoolean(queryReady) & "," & _
            """vbaProjectReadable"":" & JsonBoolean(projectReady) & "},"
    output = output & """provenMethods"":{" & _
            """powerQueryList"":false," & _
            """powerQueryCreate"":false," & _
            """powerQueryUpdate"":false," & _
            """powerQueryRename"":false," & _
            """powerQueryDelete"":false," & _
            """powerQueryRefresh"":false," & _
            """powerQueryRefreshAll"":false," & _
            """powerQueryLoadTo"":false," & _
            """powerQueryUnload"":false," & _
            """powerQueryEvaluate"":false," & _
            """xmlXPathRead"":false," & _
            """dataModelRead"":false," & _
            """scenarioCreateShow"":false," & _
            """vbaListView"":false," & _
            """vbaMutation"":false," & _
            """vbaRun"":false}}"
    HelperCapabilities = output
End Function

Private Function InspectOptionalEngines(ByVal target As Workbook) As String
    InspectOptionalEngines = "{""xmlMaps"":" & InspectXmlMaps(target) & _
        ",""workbookModel"":" & InspectWorkbookModel(target) & "}"
End Function

Private Function InspectXmlMaps(ByVal target As Workbook) As String
    On Error GoTo ProbeFailed
    Dim xmlMaps As Object
    Set xmlMaps = CallByName(target, "XmlMaps", VbGet)
    If xmlMaps Is Nothing Then
        InspectXmlMaps = EngineObservationJson( _
            "unknown", True, Null, "no_collection_object")
        Exit Function
    End If
    Dim mapCount As Long
    mapCount = CLng(CallByName(xmlMaps, "Count", VbGet))
    InspectXmlMaps = EngineObservationJson( _
        "accessible", True, mapCount, "api_access_only")
    Exit Function

ProbeFailed:
    Dim failureNumber As Long
    failureNumber = Err.Number
    Err.Clear
    If failureNumber = 438 Then
        InspectXmlMaps = EngineObservationJson( _
            "unavailable", False, Null, "api_not_exposed")
    Else
        InspectXmlMaps = EngineObservationJson( _
            "error", False, Null, "probe_failed")
    End If
End Function

Private Function InspectWorkbookModel(ByVal target As Workbook) As String
    On Error GoTo ProbeFailed
    Dim model As Object
    Set model = CallByName(target, "Model", VbGet)
    If model Is Nothing Then
        InspectWorkbookModel = EngineObservationJson( _
            "unknown", True, Null, "no_model_object_observed")
        Exit Function
    End If
    Dim modelTables As Object
    Set modelTables = CallByName(model, "ModelTables", VbGet)
    If modelTables Is Nothing Then
        InspectWorkbookModel = EngineObservationJson( _
            "unknown", True, Null, "no_model_objects_observed")
        Exit Function
    End If
    Dim tableCount As Long
    tableCount = CLng(CallByName(modelTables, "Count", VbGet))
    If tableCount = 0 Then
        InspectWorkbookModel = EngineObservationJson( _
            "unknown", True, tableCount, "no_model_objects_observed")
    Else
        InspectWorkbookModel = EngineObservationJson( _
            "accessible", True, tableCount, "api_access_only")
    End If
    Exit Function

ProbeFailed:
    Dim failureNumber As Long
    failureNumber = Err.Number
    Err.Clear
    If failureNumber = 438 Then
        InspectWorkbookModel = EngineObservationJson( _
            "unavailable", False, Null, "api_not_exposed")
    Else
        InspectWorkbookModel = EngineObservationJson( _
            "error", False, Null, "probe_failed")
    End If
End Function

Private Function EngineObservationJson( _
    ByVal status As String, _
    ByVal apiAccessible As Boolean, _
    ByVal objectCount As Variant, _
    ByVal reasonCode As String) As String
    Dim countJson As String
    If IsNull(objectCount) Then
        countJson = "null"
    Else
        countJson = CStr(CLng(objectCount))
    End If
    EngineObservationJson = "{""status"":" & JsonQuote(status) & _
        ",""apiAccessible"":" & JsonBoolean(apiAccessible) & _
        ",""objectCount"":" & countJson & _
        ",""reasonCode"":" & JsonQuote(reasonCode) & "}"
End Function

Private Function PowerQueryList(ByVal target As Workbook) As String
    Dim output As String
    output = "{""queries"":["
    Dim index As Long
    For index = 1 To target.Queries.Count
        If index > 1 Then output = output & ","
        output = output & "{""name"":" & JsonQuote(CStr(target.Queries(index).Name)) & "}"
    Next index
    PowerQueryList = output & "]}"
End Function

Private Function PowerQueryView(ByVal target As Workbook, ByVal queryName As String) As String
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    PowerQueryView = "{""name"":" & JsonQuote(CStr(query.Name)) & _
        ",""formula"":" & JsonQuote(CStr(query.Formula)) & "}"
End Function

Private Function PowerQueryCreate( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal mCode As String, _
    ByVal destination As String, _
    ByVal sheetName As Variant, _
    ByVal cellAddress As Variant) As String
    queryName = Trim$(queryName)
    If Len(queryName) = 0 Then
        Err.Raise vbObjectError + 7015, "ExcelMcpHelper", "invalid_query_name"
    End If
    If QueryExists(target, queryName) Then
        Err.Raise vbObjectError + 7010, "ExcelMcpHelper", "query_conflict"
    End If
    destination = NormalizeQueryDestination(destination)

    Dim query As Object
    Dim queryCreated As Boolean
    On Error GoTo CreateFailed
    Set query = target.Queries.Add(Name:=queryName, Formula:=mCode)
    queryCreated = True
    Dim canonicalName As String
    canonicalName = CStr(query.Name)
    If destination = "load-to-table" Then
        Dim resolvedSheet As String
        Dim resolvedCell As String
        resolvedSheet = OptionalTextOrDefault(sheetName, queryName)
        resolvedCell = OptionalTextOrDefault(cellAddress, "A1")
        LoadQueryToWorksheet target, canonicalName, resolvedSheet, resolvedCell
    End If
    PowerQueryCreate = "{}"
    Exit Function

CreateFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim cleanupNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    If failureDescription = "rollback_failed" Then
        Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
    End If
    On Error Resume Next
    Err.Clear
    If queryCreated Then
        DeleteExactQueryLoads target, canonicalName
        If Err.Number <> 0 Then cleanupNumber = Err.Number
        Err.Clear
        If Not query Is Nothing Then query.Delete
        If Err.Number <> 0 And cleanupNumber = 0 Then cleanupNumber = Err.Number
    End If
    On Error GoTo 0
    If cleanupNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function PowerQueryUpdate( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal mCode As String, _
    ByVal refresh As Boolean) As String
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    Dim previousFormula As String
    previousFormula = CStr(query.Formula)
    On Error GoTo UpdateFailed
    query.Formula = mCode
    If refresh Then
        Dim ignoredResult As String
        ignoredResult = PowerQueryRefresh(target, CStr(query.Name))
    End If
    PowerQueryUpdate = "{}"
    Exit Function

UpdateFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    Err.Clear
    query.Formula = previousFormula
    Dim rollbackNumber As Long
    rollbackNumber = Err.Number
    Err.Clear
    If rollbackNumber = 0 And refresh Then
        ignoredResult = PowerQueryRefresh(target, CStr(query.Name))
        rollbackNumber = Err.Number
    End If
    On Error GoTo 0
    If rollbackNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function PowerQueryRename( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal newName As String) As String
    queryName = Trim$(queryName)
    newName = Trim$(newName)
    If Len(newName) = 0 Then
        Err.Raise vbObjectError + 7015, "ExcelMcpHelper", "invalid_query_name"
    End If
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    If StrComp(CStr(query.Name), newName, vbBinaryCompare) = 0 Then
        PowerQueryRename = "{}"
        Exit Function
    End If
    If QueryNameOwnedByOther(target, newName, CStr(query.Name)) Then
        Err.Raise vbObjectError + 7010, "ExcelMcpHelper", "query_conflict"
    End If
    query.Name = newName
    PowerQueryRename = "{}"
End Function

Private Function PowerQueryDelete( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal deleteConnection As Boolean) As String
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    Dim canonicalName As String
    canonicalName = CStr(query.Name)
    Dim previousFormula As String
    previousFormula = CStr(query.Formula)
    Dim previousLoad As QueryLoadSnapshot
    If deleteConnection Then CaptureQueryLoad target, canonicalName, previousLoad
    On Error GoTo DeleteFailed
    If deleteConnection Then DeleteExactQueryLoads target, canonicalName
    query.Delete
    PowerQueryDelete = "{}"
    Exit Function

DeleteFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim rollbackNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    Err.Clear
    If Not QueryExists(target, canonicalName) Then
        Set query = target.Queries.Add(Name:=canonicalName, Formula:=previousFormula)
        If Err.Number <> 0 Then rollbackNumber = Err.Number
    End If
    Err.Clear
    If deleteConnection And previousLoad.HasWorksheetLoad Then
        DeleteExactQueryLoads target, canonicalName
        Err.Clear
        LoadQueryToWorksheet target, canonicalName, _
            previousLoad.SheetName, previousLoad.CellAddress
        If Err.Number <> 0 And rollbackNumber = 0 Then rollbackNumber = Err.Number
    End If
    On Error GoTo 0
    If rollbackNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function PowerQueryRefresh(ByVal target As Workbook, ByVal queryName As String) As String
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    Dim loadedSheet As String
    Dim queryTable As Object
    Set queryTable = ExactQueryTable(target, CStr(query.Name), loadedSheet)
    If queryTable Is Nothing Then
        If ExactQueryConnectionExists(target, CStr(query.Name)) Then
            Err.Raise vbObjectError + 7016, "ExcelMcpHelper", "query_destination_unsupported"
        End If
        Err.Raise vbObjectError + 7012, "ExcelMcpHelper", "query_load_not_found"
    End If

    On Error GoTo RefreshFailed
    If Not CBool(queryTable.Refresh(False)) Then
        Err.Raise vbObjectError + 7013, "ExcelMcpHelper", "query_refresh_failed"
    End If
    PowerQueryRefresh = "{""queryName"":" & JsonQuote(CStr(query.Name)) & _
        ",""hasErrors"":false,""errorMessages"":[]," & _
        """refreshTime"":" & JsonQuote(IsoTimestamp(Now)) & _
        ",""isConnectionOnly"":false,""loadedToSheet"":" & JsonQuote(loadedSheet) & "}"
    Exit Function

RefreshFailed:
    Err.Raise vbObjectError + 7013, "ExcelMcpHelper", "query_refresh_failed"
End Function

Private Function PowerQueryRefreshAll(ByVal target As Workbook) As String
    Dim failureCount As Long
    Dim failureDetails As String
    Dim index As Long
    For index = 1 To target.Queries.Count
        On Error Resume Next
        Dim ignoredResult As String
        Dim currentName As String
        currentName = CStr(target.Queries(index).Name)
        ignoredResult = PowerQueryRefresh(target, currentName)
        If Err.Number <> 0 Then
            failureCount = failureCount + 1
            AppendRefreshFailure failureDetails, currentName, Err.Description
        End If
        Err.Clear
        On Error GoTo 0
    Next index
    If failureCount > 0 Then
        mSafeErrorDetail = failureDetails
        Err.Raise vbObjectError + 7017, "ExcelMcpHelper", "query_refresh_all_failed"
    End If
    PowerQueryRefreshAll = "{}"
End Function

Private Function PowerQueryLoadTo( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal destination As String, _
    ByVal sheetName As Variant, _
    ByVal cellAddress As Variant) As String
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    destination = NormalizeQueryDestination(destination)

    Dim previousLoad As QueryLoadSnapshot
    CaptureQueryLoad target, CStr(query.Name), previousLoad
    On Error GoTo TransitionFailed
    DeleteExactQueryLoads target, CStr(query.Name)
    If destination = "load-to-table" Then
        LoadQueryToWorksheet target, CStr(query.Name), _
            OptionalTextOrDefault(sheetName, CStr(query.Name)), _
            OptionalTextOrDefault(cellAddress, "A1")
    End If
    PowerQueryLoadTo = "{}"
    Exit Function

TransitionFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim rollbackNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    If failureDescription = "rollback_failed" Then
        Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
    End If
    On Error Resume Next
    Err.Clear
    DeleteExactQueryLoads target, CStr(query.Name)
    If Err.Number <> 0 Then rollbackNumber = Err.Number
    Err.Clear
    If previousLoad.HasWorksheetLoad Then
        LoadQueryToWorksheet target, CStr(query.Name), previousLoad.SheetName, previousLoad.CellAddress
        If Err.Number <> 0 And rollbackNumber = 0 Then rollbackNumber = Err.Number
    End If
    On Error GoTo 0
    If rollbackNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function PowerQueryUnload(ByVal target As Workbook, ByVal queryName As String) As String
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    Dim previousLoad As QueryLoadSnapshot
    CaptureQueryLoad target, CStr(query.Name), previousLoad
    On Error GoTo UnloadFailed
    DeleteExactQueryLoads target, CStr(query.Name)
    If ExactQueryTableCount(target, CStr(query.Name)) <> 0 Or _
       ExactQueryConnectionExists(target, CStr(query.Name)) Then
        Err.Raise vbObjectError + 7018, "ExcelMcpHelper", "query_unload_incomplete"
    End If
    PowerQueryUnload = "{}"
    Exit Function

UnloadFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim rollbackNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    Err.Clear
    If previousLoad.HasWorksheetLoad Then
        DeleteExactQueryLoads target, CStr(query.Name)
        Err.Clear
        LoadQueryToWorksheet target, CStr(query.Name), previousLoad.SheetName, previousLoad.CellAddress
        rollbackNumber = Err.Number
    End If
    On Error GoTo 0
    If rollbackNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function PowerQueryEvaluate( _
    ByVal target As Workbook, _
    ByVal mCode As String, _
    ByVal requestId As String) As String
    If Len(Trim$(mCode)) = 0 Then
        Err.Raise vbObjectError + 7036, "ExcelMcpHelper", "invalid_query_formula"
    End If
    Dim tempName As String
    tempName = "__excelmcp_eval_" & Left$(requestId, 8)
    Dim existingSheet As Worksheet
    Set existingSheet = WorksheetByName(target, tempName)
    If QueryExists(target, tempName) Or Not existingSheet Is Nothing Or _
       ExactQueryConnectionExists(target, tempName) Or _
       ExactQueryTableCount(target, tempName) <> 0 Then
        Err.Raise vbObjectError + 7037, "ExcelMcpHelper", "temporary_name_conflict"
    End If

    Dim query As Object
    Dim queryCreated As Boolean
    Dim resultJson As String
    On Error GoTo EvaluateFailed
    Set query = target.Queries.Add(Name:=tempName, Formula:=mCode)
    queryCreated = True
    LoadQueryToWorksheet target, tempName, tempName, "A1"
    resultJson = QueryTableValuesJson(target, tempName)
    CleanupTemporaryQuery target, tempName, query
    PowerQueryEvaluate = resultJson
    Exit Function

EvaluateFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim cleanupNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    If failureDescription = "rollback_failed" Then
        Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
    End If
    On Error Resume Next
    Err.Clear
    If queryCreated Then
        CleanupTemporaryQuery target, tempName, query
        cleanupNumber = Err.Number
    End If
    On Error GoTo 0
    If cleanupNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function QueryTableValuesJson(ByVal target As Workbook, ByVal queryName As String) As String
    Dim loadedSheet As String
    Dim queryTable As Object
    Set queryTable = ExactQueryTable(target, queryName, loadedSheet)
    If queryTable Is Nothing Then
        Err.Raise vbObjectError + 7012, "ExcelMcpHelper", "query_load_not_found"
    End If
    Dim listObject As ListObject
    Set listObject = queryTable.ResultRange.ListObject
    Dim columnCount As Long
    columnCount = CLng(listObject.ListColumns.Count)
    Dim output As String
    output = "{""columns"":["
    Dim columnIndex As Long
    For columnIndex = 1 To columnCount
        If columnIndex > 1 Then output = output & ","
        output = output & JsonQuote(CStr(listObject.ListColumns(columnIndex).Name))
    Next columnIndex
    output = output & "],""rows"":["

    Dim rowCount As Long
    Dim dataBodyRange As Range
    Set dataBodyRange = listObject.DataBodyRange
    If Not dataBodyRange Is Nothing Then
        rowCount = CLng(dataBodyRange.Rows.Count)
        Dim rowIndex As Long
        For rowIndex = 1 To rowCount
            If rowIndex > 1 Then output = output & ","
            output = output & "["
            For columnIndex = 1 To columnCount
                If columnIndex > 1 Then output = output & ","
                output = output & JsonCellValue( _
                    dataBodyRange.Cells(rowIndex, columnIndex).Value2)
            Next columnIndex
            output = output & "]"
        Next rowIndex
    End If
    QueryTableValuesJson = output & "],""rowCount"":" & CStr(rowCount) & _
        ",""columnCount"":" & CStr(columnCount) & "}"
End Function

Private Sub CleanupTemporaryQuery( _
    ByVal target As Workbook, _
    ByVal tempName As String, _
    ByVal query As Object)
    Dim cleanupNumber As Long
    On Error Resume Next
    Err.Clear
    DeleteExactQueryLoads target, tempName
    If Err.Number <> 0 Then cleanupNumber = Err.Number
    Err.Clear
    Dim sheet As Worksheet
    Set sheet = WorksheetByName(target, tempName)
    If Not sheet Is Nothing Then DeleteWorksheetWithoutPrompt sheet
    If Err.Number <> 0 And cleanupNumber = 0 Then cleanupNumber = Err.Number
    Err.Clear
    If Not query Is Nothing Then query.Delete
    If Err.Number <> 0 And cleanupNumber = 0 Then cleanupNumber = Err.Number
    On Error GoTo 0
    If cleanupNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
End Sub

Private Function JsonCellValue(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then
        JsonCellValue = "null"
    ElseIf VarType(value) = vbBoolean Then
        JsonCellValue = JsonBoolean(CBool(value))
    Else
        Select Case VarType(value)
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                JsonCellValue = Trim$(Str$(CDbl(value)))
            Case vbDate
                JsonCellValue = JsonQuote(Format$(CDate(value), "yyyy-mm-dd\Thh:nn:ss"))
            Case Else
                JsonCellValue = JsonQuote(CStr(value))
        End Select
    End If
End Function

Private Sub DeleteExactQueryLoads(ByVal target As Workbook, ByVal queryName As String)
    Dim sheet As Worksheet
    For Each sheet In target.Worksheets
        Dim index As Long
        For index = sheet.ListObjects.Count To 1 Step -1
            On Error Resume Next
            Dim matchesQuery As Boolean
            matchesQuery = ConnectionMatchesQuery( _
                sheet.ListObjects(index).QueryTable.WorkbookConnection, queryName)
            If Err.Number = 0 And matchesQuery Then
                Err.Clear
                sheet.ListObjects(index).Delete
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Err.Raise vbObjectError + 7014, "ExcelMcpHelper", "query_load_delete_failed"
                End If
            End If
            Err.Clear
            On Error GoTo 0
        Next index
    Next sheet

    For index = target.Connections.Count To 1 Step -1
        If ConnectionMatchesQuery(target.Connections(index), queryName) Then
            target.Connections(index).Delete
        End If
    Next index
End Sub

Private Sub LoadQueryToWorksheet( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal sheetName As String, _
    ByVal cellAddress As String)
    sheetName = Trim$(sheetName)
    cellAddress = Trim$(cellAddress)
    If Len(sheetName) = 0 Or Len(cellAddress) = 0 Then
        Err.Raise vbObjectError + 7034, "ExcelMcpHelper", "invalid_load_target"
    End If
    If ExactQueryTableCount(target, queryName) <> 0 Or _
       ExactQueryConnectionExists(target, queryName) Then
        Err.Raise vbObjectError + 7035, "ExcelMcpHelper", "query_destination_conflict"
    End If

    On Error GoTo LoadFailed
    Dim sheet As Worksheet
    Dim createdSheet As Boolean
    Set sheet = WorksheetByName(target, sheetName)
    If sheet Is Nothing Then
        Set sheet = target.Worksheets.Add
        createdSheet = True
        sheet.Name = sheetName
    End If

    Dim destinationRange As Range
    Set destinationRange = sheet.Range(cellAddress)
    If destinationRange.Cells.CountLarge <> 1 Then
        Err.Raise vbObjectError + 7034, "ExcelMcpHelper", "invalid_load_target"
    End If
    If Not IsEmpty(destinationRange.Value2) Then
        Err.Raise vbObjectError + 7035, "ExcelMcpHelper", "query_destination_conflict"
    End If
    Dim existingTable As ListObject
    For Each existingTable In sheet.ListObjects
        Dim overlap As Range
        Set overlap = Application.Intersect(destinationRange, existingTable.Range)
        If Not overlap Is Nothing Then
            Err.Raise vbObjectError + 7035, "ExcelMcpHelper", "query_destination_conflict"
        End If
    Next existingTable

    Dim connectionName As String
    connectionName = "Query - " & queryName
    Dim connectionString As String
    connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;" & _
        "Data Source=$Workbook$;Location=" & queryName
    Dim commandText As String
    commandText = "SELECT * FROM [" & queryName & "]"

    Dim connection As WorkbookConnection
    Set connection = target.Connections.Add2( _
        Name:=connectionName, _
        Description:="Connection to the '" & queryName & "' query in the workbook.", _
        ConnectionString:=connectionString, _
        CommandText:=commandText, _
        lCmdtype:=2, _
        CreateModelConnection:=False, _
        ImportRelationships:=False)
    Dim listObject As ListObject
    Dim loadStarted As Boolean
    Set listObject = sheet.ListObjects.Add( _
        SourceType:=0, _
        Source:=connection, _
        XlListObjectHasHeaders:=1, _
        Destination:=destinationRange)
    loadStarted = True
    Dim queryTable As QueryTable
    Set queryTable = listObject.QueryTable
    queryTable.CommandType = 2
    queryTable.CommandText = commandText
    queryTable.AdjustColumnWidth = True
    queryTable.PreserveFormatting = True
    queryTable.BackgroundQuery = False
    queryTable.RefreshStyle = 1
    queryTable.PreserveColumnInfo = False
    If Not ConnectionMatchesQuery(queryTable.WorkbookConnection, queryName) Then
        Err.Raise vbObjectError + 7035, "ExcelMcpHelper", "query_destination_conflict"
    End If
    If Not CBool(queryTable.Refresh(False)) Then
        Err.Raise vbObjectError + 7013, "ExcelMcpHelper", "query_refresh_failed"
    End If
    Exit Sub

LoadFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim cleanupNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    Err.Clear
    DeleteExactQueryLoads target, queryName
    If Err.Number <> 0 Then cleanupNumber = Err.Number
    Err.Clear
    If createdSheet Then DeleteWorksheetWithoutPrompt sheet
    If Err.Number <> 0 And cleanupNumber = 0 Then cleanupNumber = Err.Number
    If loadStarted And Not createdSheet And cleanupNumber = 0 Then
        cleanupNumber = vbObjectError + 7025
    End If
    On Error GoTo 0
    If cleanupNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Sub

Private Sub CaptureQueryLoad( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByRef snapshot As QueryLoadSnapshot)
    Dim loadCount As Long
    loadCount = ExactQueryTableCount(target, queryName)
    If loadCount > 1 Then
        Err.Raise vbObjectError + 7016, "ExcelMcpHelper", "query_destination_unsupported"
    End If
    If loadCount = 0 Then
        If ExactQueryConnectionExists(target, queryName) Then
            Err.Raise vbObjectError + 7016, "ExcelMcpHelper", "query_destination_unsupported"
        End If
        Exit Sub
    End If

    Dim loadedSheet As String
    Dim queryTable As Object
    Set queryTable = ExactQueryTable(target, queryName, loadedSheet)
    snapshot.HasWorksheetLoad = True
    snapshot.SheetName = loadedSheet
    snapshot.CellAddress = CStr(queryTable.ResultRange.Cells(1, 1).Address(False, False))
End Sub

Private Function ExactQueryTable( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByRef loadedSheet As String) As Object
    Dim found As Boolean
    Dim sheet As Worksheet
    For Each sheet In target.Worksheets
        Dim listObject As ListObject
        For Each listObject In sheet.ListObjects
            On Error Resume Next
            Dim readSucceeded As Boolean
            Dim matchesQuery As Boolean
            matchesQuery = ConnectionMatchesQuery( _
                listObject.QueryTable.WorkbookConnection, queryName)
            readSucceeded = (Err.Number = 0)
            Err.Clear
            On Error GoTo 0
            If readSucceeded And matchesQuery Then
                If found Then
                    Err.Raise vbObjectError + 7016, _
                        "ExcelMcpHelper", "query_destination_unsupported"
                End If
                Set ExactQueryTable = listObject.QueryTable
                loadedSheet = CStr(sheet.Name)
                found = True
            End If
        Next listObject
    Next sheet
End Function

Private Function ExactQueryTableCount(ByVal target As Workbook, ByVal queryName As String) As Long
    Dim sheet As Worksheet
    For Each sheet In target.Worksheets
        Dim listObject As ListObject
        For Each listObject In sheet.ListObjects
            On Error Resume Next
            Dim matchesQuery As Boolean
            matchesQuery = ConnectionMatchesQuery( _
                listObject.QueryTable.WorkbookConnection, queryName)
            If Err.Number = 0 And matchesQuery Then
                ExactQueryTableCount = ExactQueryTableCount + 1
            End If
            Err.Clear
            On Error GoTo 0
        Next listObject
    Next sheet
End Function

Private Function ExactQueryConnectionExists( _
    ByVal target As Workbook, _
    ByVal queryName As String) As Boolean
    Dim index As Long
    For index = 1 To target.Connections.Count
        If ConnectionMatchesQuery(target.Connections(index), queryName) Then
            ExactQueryConnectionExists = True
            Exit Function
        End If
    Next index
End Function

Private Function ConnectionMatchesQuery( _
    ByVal connection As WorkbookConnection, _
    ByVal queryName As String) As Boolean
    Dim location As String
    If TryGetMashupLocation(connection, location) Then
        ConnectionMatchesQuery = _
            (StrComp(location, queryName, vbTextCompare) = 0)
    End If
End Function

Private Function TryGetMashupLocation( _
    ByVal connection As WorkbookConnection, _
    ByRef location As String) As Boolean
    On Error GoTo NotMashup
    Dim connectionString As String
    connectionString = CStr(connection.OLEDBConnection.Connection)
    Dim provider As String
    If Not TryGetConnectionProperty(connectionString, "Provider", provider) Then Exit Function
    If StrComp(provider, "Microsoft.Mashup.OleDb.1", vbTextCompare) <> 0 Then Exit Function
    If Not TryGetConnectionProperty(connectionString, "Location", location) Then Exit Function
    TryGetMashupLocation = (Len(location) > 0)
    Exit Function

NotMashup:
    Err.Clear
End Function

Private Function TryGetConnectionProperty( _
    ByVal connectionString As String, _
    ByVal propertyName As String, _
    ByRef propertyValue As String) As Boolean
    Dim segmentStart As Long
    segmentStart = 1
    Dim quoteCharacter As String
    Dim index As Long
    index = 1
    Do While index <= Len(connectionString) + 1
        Dim atEnd As Boolean
        atEnd = (index > Len(connectionString))
        Dim character As String
        If atEnd Then
            character = ";"
        Else
            character = Mid$(connectionString, index, 1)
        End If

        If Len(quoteCharacter) > 0 Then
            If character = quoteCharacter Then
                If index < Len(connectionString) And _
                   Mid$(connectionString, index + 1, 1) = quoteCharacter Then
                    index = index + 1
                Else
                    quoteCharacter = vbNullString
                End If
            End If
        ElseIf character = """" Or character = "'" Then
            quoteCharacter = character
        ElseIf character = ";" Then
            Dim segment As String
            segment = Mid$(connectionString, segmentStart, index - segmentStart)
            Dim equalsIndex As Long
            equalsIndex = InStr(1, segment, "=", vbBinaryCompare)
            If equalsIndex > 0 Then
                Dim key As String
                key = Trim$(Left$(segment, equalsIndex - 1))
                If StrComp(key, propertyName, vbTextCompare) = 0 Then
                    Dim rawValue As String
                    rawValue = Trim$(Mid$(segment, equalsIndex + 1))
                    propertyValue = DecodeConnectionPropertyValue(rawValue)
                    TryGetConnectionProperty = True
                    Exit Function
                End If
            End If
            segmentStart = index + 1
        End If
        index = index + 1
    Loop
End Function

Private Function DecodeConnectionPropertyValue(ByVal rawValue As String) As String
    If Len(rawValue) >= 2 Then
        Dim quoteCharacter As String
        quoteCharacter = Left$(rawValue, 1)
        If (quoteCharacter = """" Or quoteCharacter = "'") And _
           Right$(rawValue, 1) = quoteCharacter Then
            Dim innerValue As String
            innerValue = Mid$(rawValue, 2, Len(rawValue) - 2)
            DecodeConnectionPropertyValue = _
                Replace(innerValue, quoteCharacter & quoteCharacter, quoteCharacter)
            Exit Function
        End If
    End If
    DecodeConnectionPropertyValue = rawValue
End Function

Private Function WorksheetByName(ByVal target As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet
    For Each sheet In target.Worksheets
        If StrComp(CStr(sheet.Name), sheetName, vbTextCompare) = 0 Then
            Set WorksheetByName = sheet
            Exit Function
        End If
    Next sheet
End Function

Private Sub DeleteWorksheetWithoutPrompt(ByVal sheet As Worksheet)
    Dim previousAlerts As Boolean
    previousAlerts = Application.DisplayAlerts
    On Error GoTo RestoreAlerts
    Application.DisplayAlerts = False
    sheet.Delete
RestoreAlerts:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description
    Application.DisplayAlerts = previousAlerts
    If failureNumber <> 0 Then
        Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
    End If
End Sub

Private Function NormalizeQueryDestination(ByVal destination As String) As String
    destination = LCase$(Trim$(destination))
    Select Case destination
        Case "load-to-table", "connection-only"
            NormalizeQueryDestination = destination
        Case Else
            Err.Raise vbObjectError + 7016, _
                "ExcelMcpHelper", "query_destination_unsupported"
    End Select
End Function

Private Function OptionalTextOrDefault( _
    ByVal value As Variant, _
    ByVal defaultValue As String) As String
    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
        OptionalTextOrDefault = defaultValue
    Else
        OptionalTextOrDefault = Trim$(CStr(value))
    End If
End Function

Private Function IsoTimestamp(ByVal value As Date) As String
    IsoTimestamp = Format$(value, "yyyy-mm-dd\THH:nn:ss")
End Function

Private Function AnalysisCreateScenario( _
    ByVal target As Workbook, _
    ByVal sheetName As String, _
    ByVal scenarioName As String, _
    ByVal changingCells As String, _
    ByVal values As Variant, _
    ByVal comment As Variant, _
    ByVal locked As Boolean, _
    ByVal hidden As Boolean) As String
    Dim sheet As Worksheet
    Set sheet = WorksheetByExactName(target, sheetName)
    If ScenarioExists(sheet, scenarioName) Then
        Err.Raise vbObjectError + 7040, "ExcelMcpHelper", "scenario_conflict"
    End If

    Dim changingRange As Range
    Set changingRange = sheet.Range(changingCells)
    Dim cellCount As Long
    cellCount = CLng(changingRange.CountLarge)
    If cellCount < 1 Or cellCount > 32 Then
        Err.Raise vbObjectError + 7041, "ExcelMcpHelper", "scenario_cell_count"
    End If
    If ArrayLength(values) <> cellCount Then
        Err.Raise vbObjectError + 7042, "ExcelMcpHelper", "scenario_value_count"
    End If

    Dim scenarios As Scenarios
    Set scenarios = sheet.Scenarios
    scenarios.Add Name:=scenarioName, ChangingCells:=changingRange, Values:=values, _
        Comment:=comment, Locked:=locked, Hidden:=hidden
    AnalysisCreateScenario = "{""message"":" & _
        JsonQuote("Scenario '" & scenarioName & "' created on '" & sheetName & "'.") & "}"
End Function

Private Function AnalysisShowScenario( _
    ByVal target As Workbook, _
    ByVal sheetName As String, _
    ByVal scenarioName As String) As String
    Dim sheet As Worksheet
    Set sheet = WorksheetByExactName(target, sheetName)
    Dim scenario As Object
    Set scenario = ScenarioByExactName(sheet, scenarioName)
    scenario.Show
    AnalysisShowScenario = "{""message"":" & _
        JsonQuote("Scenario '" & scenarioName & "' shown on '" & sheetName & "'.") & "}"
End Function

Private Function WorksheetByExactName(ByVal target As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet
    For Each sheet In target.Worksheets
        If StrComp(CStr(sheet.Name), sheetName, vbBinaryCompare) = 0 Then
            Set WorksheetByExactName = sheet
            Exit Function
        End If
    Next sheet
    Err.Raise vbObjectError + 7043, "ExcelMcpHelper", "worksheet_not_found"
End Function

Private Function ScenarioByExactName(ByVal sheet As Worksheet, ByVal scenarioName As String) As Object
    Dim scenarios As Object
    Set scenarios = sheet.Scenarios
    Dim index As Long
    For index = 1 To scenarios.Count
        If StrComp(CStr(scenarios(index).Name), scenarioName, vbBinaryCompare) = 0 Then
            Set ScenarioByExactName = scenarios(index)
            Exit Function
        End If
    Next index
    Err.Raise vbObjectError + 7044, "ExcelMcpHelper", "scenario_not_found"
End Function

Private Function ScenarioExists(ByVal sheet As Worksheet, ByVal scenarioName As String) As Boolean
    Dim scenarios As Object
    Set scenarios = sheet.Scenarios
    Dim index As Long
    For index = 1 To scenarios.Count
        If StrComp(CStr(scenarios(index).Name), scenarioName, vbTextCompare) = 0 Then
            ScenarioExists = True
            Exit Function
        End If
    Next index
End Function

Private Function ArrayLength(ByVal values As Variant) As Long
    On Error GoTo EmptyArray
    ArrayLength = UBound(values) - LBound(values) + 1
    Exit Function
EmptyArray:
    ArrayLength = 0
End Function

Private Function VbaList(ByVal target As Workbook) As String
    Dim project As Object
    Set project = target.VBProject
    Dim output As String
    output = "{""modules"":["
    Dim index As Long
    For index = 1 To project.VBComponents.Count
        Dim component As Object
        Set component = project.VBComponents(index)
        If index > 1 Then output = output & ","
        output = output & "{""name"":" & JsonQuote(CStr(component.Name)) & _
            ",""type"":" & CStr(CLng(component.Type)) & _
            ",""lineCount"":" & CStr(CLng(component.CodeModule.CountOfLines)) & _
            ",""procedures"":" & VbaProceduresJson(component.CodeModule) & "}"
    Next index
    VbaList = output & "]}"
End Function

Private Function VbaView(ByVal target As Workbook, ByVal moduleName As String) As String
    Dim component As Object
    Set component = ComponentByExactName(target, moduleName)
    Dim count As Long
    count = CLng(component.CodeModule.CountOfLines)
    Dim source As String
    If count > 0 Then source = CStr(component.CodeModule.Lines(1, count))
    VbaView = "{""moduleName"":" & JsonQuote(CStr(component.Name)) & _
        ",""moduleType"":" & CStr(CLng(component.Type)) & _
        ",""lineCount"":" & CStr(count) & _
        ",""source"":" & JsonQuote(source) & _
        ",""procedures"":" & VbaProceduresJson(component.CodeModule) & "}"
End Function

Private Function VbaImport( _
    ByVal target As Workbook, _
    ByVal moduleName As String, _
    ByVal source As String) As String
    EnsureMutableProject target
    EnsureValidModuleName moduleName
    If ComponentExists(target, moduleName) Then
        Err.Raise vbObjectError + 7020, "ExcelMcpHelper", "module_conflict"
    End If
    Dim component As Object
    On Error GoTo ImportFailed
    Set component = target.VBProject.VBComponents.Add(STANDARD_MODULE_TYPE)
    component.Name = moduleName
    component.CodeModule.AddFromString source
    VbaImport = "{}"
    Exit Function

ImportFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim cleanupNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    Err.Clear
    If Not component Is Nothing Then target.VBProject.VBComponents.Remove component
    cleanupNumber = Err.Number
    On Error GoTo 0
    If cleanupNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function VbaUpdate( _
    ByVal target As Workbook, _
    ByVal moduleName As String, _
    ByVal source As String) As String
    EnsureMutableProject target
    Dim component As Object
    Set component = ComponentByExactName(target, moduleName)
    EnsureStandardModule component
    Dim lineCount As Long
    lineCount = CLng(component.CodeModule.CountOfLines)
    Dim previousSource As String
    If lineCount > 0 Then previousSource = CStr(component.CodeModule.Lines(1, lineCount))
    On Error GoTo UpdateFailed
    If lineCount > 0 Then component.CodeModule.DeleteLines 1, lineCount
    component.CodeModule.AddFromString source
    VbaUpdate = "{}"
    Exit Function

UpdateFailed:
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim rollbackNumber As Long
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    Err.Clear
    lineCount = CLng(component.CodeModule.CountOfLines)
    If Err.Number <> 0 Then rollbackNumber = Err.Number
    Err.Clear
    If lineCount > 0 Then component.CodeModule.DeleteLines 1, lineCount
    If Err.Number <> 0 And rollbackNumber = 0 Then rollbackNumber = Err.Number
    Err.Clear
    If Len(previousSource) > 0 Then component.CodeModule.AddFromString previousSource
    If Err.Number <> 0 And rollbackNumber = 0 Then rollbackNumber = Err.Number
    On Error GoTo 0
    If rollbackNumber <> 0 Then
        Err.Raise vbObjectError + 7025, "ExcelMcpHelper", "rollback_failed"
    End If
    Err.Raise failureNumber, "ExcelMcpHelper", failureDescription
End Function

Private Function VbaDelete(ByVal target As Workbook, ByVal moduleName As String) As String
    EnsureMutableProject target
    Dim component As Object
    Set component = ComponentByExactName(target, moduleName)
    EnsureStandardModule component
    target.VBProject.VBComponents.Remove component
    VbaDelete = "{}"
End Function

Private Function VbaRun( _
    ByVal target As Workbook, _
    ByVal procedureName As String, _
    ByVal parameters As Variant) As String
    Dim separator As Long
    separator = InStr(1, procedureName, ".", vbBinaryCompare)
    If separator < 2 Or separator <> InStrRev(procedureName, ".", -1, vbBinaryCompare) Or _
            separator = Len(procedureName) Then
        Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_procedure_name"
    End If
    If InStr(1, procedureName, " ", vbBinaryCompare) > 0 Or _
            InStr(1, procedureName, vbTab, vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_procedure_name"
    End If
    EnsureValidProcedureIdentity _
        Left$(procedureName, separator - 1), _
        Mid$(procedureName, separator + 1)

    Dim count As Long
    count = ArrayLength(parameters)
    If count > 30 Then
        Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "too_many_parameters"
    End If
    Dim index As Long
    For index = 0 To count - 1
        If VarType(parameters(index)) <> vbString Then
            Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_parameter"
        End If
    Next index

    Dim qualifiedName As String
    qualifiedName = "'" & Replace$(target.Name, "'", "''") & "'!" & procedureName
    RunVbaProcedure qualifiedName, parameters, count
    VbaRun = "{}"
End Function

Private Sub EnsureValidProcedureIdentity( _
    ByVal moduleName As String, _
    ByVal memberName As String)
    EnsureValidModuleName moduleName
    If Len(memberName) < 1 Or Len(memberName) > 255 Then
        Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_procedure_name"
    End If
    Dim index As Long
    For index = 1 To Len(memberName)
        Dim character As String
        character = Mid$(memberName, index, 1)
        If index = 1 Then
            If Not ((character >= "A" And character <= "Z") Or _
                    (character >= "a" And character <= "z")) Then
                Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_procedure_name"
            End If
        ElseIf Not ((character >= "A" And character <= "Z") Or _
                (character >= "a" And character <= "z") Or _
                (character >= "0" And character <= "9") Or character = "_") Then
            Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_procedure_name"
        End If
    Next index
End Sub

Private Sub RunVbaProcedure( _
    ByVal qualifiedName As String, _
    ByVal parameters As Variant, _
    ByVal count As Long)
    Select Case count
        Case 0: Application.Run qualifiedName
        Case 1: Application.Run qualifiedName, parameters(0)
        Case 2: Application.Run qualifiedName, parameters(0), parameters(1)
        Case 3: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2)
        Case 4: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3)
        Case 5: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4)
        Case 6: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5)
        Case 7: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6)
        Case 8: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7)
        Case 9: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8)
        Case 10: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9)
        Case 11: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10)
        Case 12: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11)
        Case 13: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12)
        Case 14: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13)
        Case 15: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14)
        Case 16: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15)
        Case 17: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16)
        Case 18: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17)
        Case 19: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18)
        Case 20: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19)
        Case 21: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20)
        Case 22: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21)
        Case 23: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22)
        Case 24: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23)
        Case 25: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), parameters(24)
        Case 26: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), parameters(24), parameters(25)
        Case 27: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), parameters(24), parameters(25), parameters(26)
        Case 28: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), parameters(24), parameters(25), parameters(26), parameters(27)
        Case 29: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), parameters(24), parameters(25), parameters(26), parameters(27), parameters(28)
        Case 30: Application.Run qualifiedName, parameters(0), parameters(1), parameters(2), parameters(3), parameters(4), parameters(5), parameters(6), parameters(7), parameters(8), parameters(9), parameters(10), parameters(11), parameters(12), parameters(13), parameters(14), parameters(15), parameters(16), parameters(17), parameters(18), parameters(19), parameters(20), parameters(21), parameters(22), parameters(23), parameters(24), parameters(25), parameters(26), parameters(27), parameters(28), parameters(29)
        Case Else
            Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "too_many_parameters"
    End Select
End Sub

Private Function VbaProceduresJson(ByVal codeModule As Object) As String
    Dim output As String
    output = "["
    Dim emitted As Long
    Dim lineNumber As Long
    For lineNumber = 1 To CLng(codeModule.CountOfLines)
        Dim procedureName As String
        procedureName = VbaProcedureName(CStr(codeModule.Lines(lineNumber, 1)))
        If Len(procedureName) > 0 Then
            If emitted > 0 Then output = output & ","
            output = output & JsonQuote(procedureName)
            emitted = emitted + 1
        End If
    Next lineNumber
    VbaProceduresJson = output & "]"
End Function

Private Function VbaProcedureName(ByVal codeLine As String) As String
    Dim value As String
    value = LTrim$(codeLine)
    Dim prefixes As Variant
    prefixes = Array( _
        "Public Function ", _
        "Public Sub ", _
        "Private Function ", _
        "Private Sub ", _
        "Function ", _
        "Sub ")
    Dim index As Long
    For index = LBound(prefixes) To UBound(prefixes)
        Dim prefix As String
        prefix = CStr(prefixes(index))
        If StrComp(Left$(value, Len(prefix)), prefix, vbBinaryCompare) = 0 Then
            value = Mid$(value, Len(prefix) + 1)
            Dim parenthesis As Long
            Dim whitespace As Long
            parenthesis = InStr(1, value, "(", vbBinaryCompare)
            whitespace = InStr(1, value, " ", vbBinaryCompare)
            If parenthesis > 0 And (whitespace = 0 Or parenthesis < whitespace) Then
                VbaProcedureName = Left$(value, parenthesis - 1)
            ElseIf whitespace > 0 Then
                VbaProcedureName = Left$(value, whitespace - 1)
            Else
                VbaProcedureName = value
            End If
            Exit Function
        End If
    Next index
End Function

Private Sub EnsureMutableProject(ByVal target As Workbook)
    If CBool(target.VBASigned) Then
        Err.Raise vbObjectError + 7021, "ExcelMcpHelper", "signed_project"
    End If
    If CLng(target.VBProject.Protection) = PROJECT_LOCKED Then
        Err.Raise vbObjectError + 7022, "ExcelMcpHelper", "locked_project"
    End If
End Sub

Private Sub EnsureStandardModule(ByVal component As Object)
    If CLng(component.Type) <> STANDARD_MODULE_TYPE Then
        Err.Raise vbObjectError + 7023, "ExcelMcpHelper", "component_type"
    End If
End Sub

Private Sub EnsureValidModuleName(ByVal moduleName As String)
    If Len(moduleName) < 1 Or Len(moduleName) > 31 Then
        Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_module_name"
    End If
    Dim firstCharacter As String
    firstCharacter = Left$(moduleName, 1)
    If Not ((firstCharacter >= "A" And firstCharacter <= "Z") Or _
            (firstCharacter >= "a" And firstCharacter <= "z")) Then
        Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_module_name"
    End If
    Dim index As Long
    For index = 2 To Len(moduleName)
        Dim character As String
        character = Mid$(moduleName, index, 1)
        If Not ((character >= "A" And character <= "Z") Or _
                (character >= "a" And character <= "z") Or _
                (character >= "0" And character <= "9") Or character = "_") Then
            Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "invalid_module_name"
        End If
    Next index
End Sub

Private Function QueryByExactName(ByVal target As Workbook, ByVal queryName As String) As Object
    queryName = Trim$(queryName)
    Dim index As Long
    For index = 1 To target.Queries.Count
        If StrComp(CStr(target.Queries(index).Name), queryName, vbTextCompare) = 0 Then
            Set QueryByExactName = target.Queries(index)
            Exit Function
        End If
    Next index
    Err.Raise vbObjectError + 7011, "ExcelMcpHelper", "query_not_found"
End Function

Private Function QueryExists(ByVal target As Workbook, ByVal queryName As String) As Boolean
    queryName = Trim$(queryName)
    Dim index As Long
    For index = 1 To target.Queries.Count
        If StrComp(CStr(target.Queries(index).Name), queryName, vbTextCompare) = 0 Then
            QueryExists = True
            Exit Function
        End If
    Next index
End Function

Private Function QueryNameOwnedByOther( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal currentName As String) As Boolean
    Dim index As Long
    For index = 1 To target.Queries.Count
        Dim candidateName As String
        candidateName = CStr(target.Queries(index).Name)
        If StrComp(candidateName, queryName, vbTextCompare) = 0 And _
           StrComp(candidateName, currentName, vbBinaryCompare) <> 0 Then
            QueryNameOwnedByOther = True
            Exit Function
        End If
    Next index
End Function

Private Function ComponentByExactName(ByVal target As Workbook, ByVal moduleName As String) As Object
    Dim components As Object
    Set components = target.VBProject.VBComponents
    Dim index As Long
    For index = 1 To components.Count
        If StrComp(CStr(components(index).Name), moduleName, vbBinaryCompare) = 0 Then
            Set ComponentByExactName = components(index)
            Exit Function
        End If
    Next index
    Err.Raise vbObjectError + 7024, "ExcelMcpHelper", "module_not_found"
End Function

Private Function ComponentExists(ByVal target As Workbook, ByVal moduleName As String) As Boolean
    Dim components As Object
    Set components = target.VBProject.VBComponents
    Dim index As Long
    For index = 1 To components.Count
        If StrComp(CStr(components(index).Name), moduleName, vbBinaryCompare) = 0 Then
            ComponentExists = True
            Exit Function
        End If
    Next index
End Function

Private Function WorkbookByExactFullName(ByVal workbookPath As String) As Workbook
    Dim candidate As Workbook
    For Each candidate In Application.Workbooks
        If StrComp(CStr(candidate.FullName), workbookPath, vbBinaryCompare) = 0 Then
            Set WorkbookByExactFullName = candidate
            Exit Function
        End If
    Next candidate
    Err.Raise vbObjectError + 7003, "ExcelMcpHelper", "workbook_not_found"
End Function

Private Sub ValidateRequestEnvelope(ByVal json As String)
    ValidateArgumentKeys json, "version,requestId,workbookPath,action,arguments"
End Sub

Private Sub ValidateArgumentKeys(ByVal json As String, ByVal expectedCsv As String)
    Dim keys As Collection
    Set keys = JsonObjectKeys(json)
    Dim expectedCount As Long
    If Len(expectedCsv) > 0 Then expectedCount = UBound(Split(expectedCsv, ",")) + 1
    If keys.Count <> expectedCount Then
        Err.Raise vbObjectError + 7030, "ExcelMcpHelper", "invalid_properties"
    End If

    Dim index As Long
    For index = 1 To keys.Count
        If Not CsvContainsExact(expectedCsv, CStr(keys(index))) Then
            Err.Raise vbObjectError + 7030, "ExcelMcpHelper", "invalid_properties"
        End If
        Dim later As Long
        For later = index + 1 To keys.Count
            If StrComp(CStr(keys(index)), CStr(keys(later)), vbBinaryCompare) = 0 Then
                Err.Raise vbObjectError + 7031, "ExcelMcpHelper", "duplicate_property"
            End If
        Next later
    Next index
End Sub

Private Function CsvContainsExact(ByVal csv As String, ByVal value As String) As Boolean
    If Len(csv) = 0 Then Exit Function
    Dim item As Variant
    For Each item In Split(csv, ",")
        If StrComp(CStr(item), value, vbBinaryCompare) = 0 Then
            CsvContainsExact = True
            Exit Function
        End If
    Next item
End Function

Private Function JsonObjectKeys(ByVal json As String) As Collection
    Dim keys As New Collection
    Dim position As Long
    position = SkipWhitespace(json, 1)
    If Mid$(json, position, 1) <> "{" Then Err.Raise vbObjectError + 7032
    position = SkipWhitespace(json, position + 1)
    If Mid$(json, position, 1) = "}" Then
        Set JsonObjectKeys = keys
        Exit Function
    End If

    Do
        Dim key As String
        key = ParseJsonString(json, position)
        keys.Add key
        position = SkipWhitespace(json, position)
        If Mid$(json, position, 1) <> ":" Then Err.Raise vbObjectError + 7032
        position = SkipJsonValue(json, position + 1)
        position = SkipWhitespace(json, position)
        If Mid$(json, position, 1) = "}" Then Exit Do
        If Mid$(json, position, 1) <> "," Then Err.Raise vbObjectError + 7032
        position = SkipWhitespace(json, position + 1)
    Loop
    position = SkipWhitespace(json, position + 1)
    If position <= Len(json) Then Err.Raise vbObjectError + 7032
    Set JsonObjectKeys = keys
End Function

Private Function JsonRequiredString(ByVal json As String, ByVal propertyName As String) As String
    Dim raw As String
    raw = JsonRequiredRaw(json, propertyName)
    Dim position As Long
    position = 1
    JsonRequiredString = ParseJsonString(raw, position)
    If SkipWhitespace(raw, position) <= Len(raw) Then Err.Raise vbObjectError + 7032
End Function

Private Function JsonRequiredLong(ByVal json As String, ByVal propertyName As String) As Long
    Dim raw As String
    raw = JsonRequiredRaw(json, propertyName)
    If Not IsJsonUnsignedLong(raw) Then Err.Raise vbObjectError + 7032
    JsonRequiredLong = CLng(raw)
End Function

Private Function IsJsonUnsignedLong(ByVal value As String) As Boolean
    If Len(value) = 0 Or Len(value) > 10 Then Exit Function
    If Len(value) > 1 And Left$(value, 1) = "0" Then Exit Function
    Dim index As Long
    For index = 1 To Len(value)
        Dim character As String
        character = Mid$(value, index, 1)
        If character < "0" Or character > "9" Then Exit Function
    Next index
    If Len(value) = 10 And StrComp(value, "2147483647", vbBinaryCompare) > 0 Then Exit Function
    IsJsonUnsignedLong = True
End Function

Private Function JsonRequiredBoolean(ByVal json As String, ByVal propertyName As String) As Boolean
    Dim raw As String
    raw = JsonRequiredRaw(json, propertyName)
    If StrComp(raw, "true", vbBinaryCompare) = 0 Then
        JsonRequiredBoolean = True
    ElseIf StrComp(raw, "false", vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7032
    End If
End Function

Private Function JsonRequiredObject(ByVal json As String, ByVal propertyName As String) As String
    Dim raw As String
    raw = JsonRequiredRaw(json, propertyName)
    If Left$(raw, 1) <> "{" Then Err.Raise vbObjectError + 7032
    JsonRequiredObject = raw
End Function

Private Function JsonOptionalString(ByVal json As String, ByVal propertyName As String) As Variant
    Dim raw As String
    raw = JsonRequiredRaw(json, propertyName)
    If StrComp(raw, "null", vbBinaryCompare) = 0 Then
        JsonOptionalString = Empty
        Exit Function
    End If
    Dim position As Long
    position = 1
    JsonOptionalString = ParseJsonString(raw, position)
    If SkipWhitespace(raw, position) <= Len(raw) Then Err.Raise vbObjectError + 7032
End Function

Private Function JsonRequiredArrayValues(ByVal json As String, ByVal propertyName As String) As Variant
    Dim raw As String
    raw = JsonRequiredRaw(json, propertyName)
    Dim position As Long
    position = SkipWhitespace(raw, 1)
    If Mid$(raw, position, 1) <> "[" Then Err.Raise vbObjectError + 7032
    position = SkipWhitespace(raw, position + 1)

    Dim items As New Collection
    If Mid$(raw, position, 1) <> "]" Then
        Do
            Dim valueStart As Long
            valueStart = position
            Dim valueEnd As Long
            valueEnd = SkipJsonValue(raw, valueStart)
            items.Add JsonScalarValue(Mid$(raw, valueStart, valueEnd - valueStart))
            position = SkipWhitespace(raw, valueEnd)
            If Mid$(raw, position, 1) = "]" Then Exit Do
            If Mid$(raw, position, 1) <> "," Then Err.Raise vbObjectError + 7032
            position = SkipWhitespace(raw, position + 1)
        Loop
    End If
    position = SkipWhitespace(raw, position + 1)
    If position <= Len(raw) Then Err.Raise vbObjectError + 7032

    Dim values() As Variant
    If items.Count = 0 Then
        JsonRequiredArrayValues = Array()
        Exit Function
    End If
    ReDim values(0 To items.Count - 1)
    Dim index As Long
    For index = 1 To items.Count
        values(index - 1) = items(index)
    Next index
    JsonRequiredArrayValues = values
End Function

Private Function JsonScalarValue(ByVal raw As String) As Variant
    raw = Trim$(raw)
    If Left$(raw, 1) = """" Then
        Dim position As Long
        position = 1
        JsonScalarValue = ParseJsonString(raw, position)
        If SkipWhitespace(raw, position) <= Len(raw) Then Err.Raise vbObjectError + 7032
    ElseIf StrComp(raw, "true", vbBinaryCompare) = 0 Then
        JsonScalarValue = True
    ElseIf StrComp(raw, "false", vbBinaryCompare) = 0 Then
        JsonScalarValue = False
    ElseIf StrComp(raw, "null", vbBinaryCompare) = 0 Then
        JsonScalarValue = Empty
    ElseIf IsJsonNumber(raw) Then
        JsonScalarValue = Val(raw)
    Else
        Err.Raise vbObjectError + 7032
    End If
End Function

Private Function IsJsonNumber(ByVal value As String) As Boolean
    If Len(value) = 0 Then Exit Function
    Dim position As Long
    position = 1
    If Mid$(value, position, 1) = "-" Then position = position + 1
    If position > Len(value) Then Exit Function
    If Mid$(value, position, 1) = "0" Then
        position = position + 1
    ElseIf Mid$(value, position, 1) >= "1" And Mid$(value, position, 1) <= "9" Then
        Do While position <= Len(value) And Mid$(value, position, 1) >= "0" And Mid$(value, position, 1) <= "9"
            position = position + 1
        Loop
    Else
        Exit Function
    End If
    If position <= Len(value) And Mid$(value, position, 1) = "." Then
        position = position + 1
        Dim fractionStart As Long
        fractionStart = position
        Do While position <= Len(value) And Mid$(value, position, 1) >= "0" And Mid$(value, position, 1) <= "9"
            position = position + 1
        Loop
        If position = fractionStart Then Exit Function
    End If
    If position <= Len(value) And (Mid$(value, position, 1) = "e" Or Mid$(value, position, 1) = "E") Then
        position = position + 1
        If position <= Len(value) And (Mid$(value, position, 1) = "+" Or Mid$(value, position, 1) = "-") Then
            position = position + 1
        End If
        Dim exponentStart As Long
        exponentStart = position
        Do While position <= Len(value) And Mid$(value, position, 1) >= "0" And Mid$(value, position, 1) <= "9"
            position = position + 1
        Loop
        If position = exponentStart Then Exit Function
    End If
    IsJsonNumber = (position > Len(value))
End Function

Private Function JsonRequiredRaw(ByVal json As String, ByVal propertyName As String) As String
    Dim position As Long
    position = SkipWhitespace(json, 1)
    If Mid$(json, position, 1) <> "{" Then Err.Raise vbObjectError + 7032
    position = SkipWhitespace(json, position + 1)
    Do While position <= Len(json) And Mid$(json, position, 1) <> "}"
        Dim key As String
        key = ParseJsonString(json, position)
        position = SkipWhitespace(json, position)
        If Mid$(json, position, 1) <> ":" Then Err.Raise vbObjectError + 7032
        Dim valueStart As Long
        valueStart = SkipWhitespace(json, position + 1)
        Dim valueEnd As Long
        valueEnd = SkipJsonValue(json, valueStart)
        If StrComp(key, propertyName, vbBinaryCompare) = 0 Then
            JsonRequiredRaw = Mid$(json, valueStart, valueEnd - valueStart)
            Exit Function
        End If
        position = SkipWhitespace(json, valueEnd)
        If Mid$(json, position, 1) = "," Then position = SkipWhitespace(json, position + 1)
    Loop
    Err.Raise vbObjectError + 7033, "ExcelMcpHelper", "missing_property"
End Function

Private Function ParseJsonString(ByVal json As String, ByRef position As Long) As String
    position = SkipWhitespace(json, position)
    If Mid$(json, position, 1) <> """" Then Err.Raise vbObjectError + 7032
    position = position + 1
    Dim output As String
    Do While position <= Len(json)
        Dim character As String
        character = Mid$(json, position, 1)
        If character = """" Then
            position = position + 1
            ParseJsonString = output
            Exit Function
        End If
        If character = "\" Then
            position = position + 1
            character = Mid$(json, position, 1)
            Select Case character
                Case """", "\", "/": output = output & character
                Case "b": output = output & Chr$(8)
                Case "f": output = output & Chr$(12)
                Case "n": output = output & vbLf
                Case "r": output = output & vbCr
                Case "t": output = output & vbTab
                Case "u"
                    Dim hexValue As String
                    hexValue = Mid$(json, position + 1, 4)
                    If Len(hexValue) <> 4 Or Not IsHex4(hexValue) Then Err.Raise vbObjectError + 7032
                    output = output & ChrW$(CLng("&H" & hexValue))
                    position = position + 4
                Case Else: Err.Raise vbObjectError + 7032
            End Select
        Else
            If AscW(character) >= 0 And AscW(character) < 32 Then Err.Raise vbObjectError + 7032
            output = output & character
        End If
        position = position + 1
    Loop
    Err.Raise vbObjectError + 7032
End Function

Private Function SkipJsonValue(ByVal json As String, ByVal position As Long) As Long
    position = SkipWhitespace(json, position)
    Dim character As String
    character = Mid$(json, position, 1)
    If character = """" Then
        Dim ignored As String
        ignored = ParseJsonString(json, position)
        SkipJsonValue = position
        Exit Function
    End If
    If character = "{" Or character = "[" Then
        Dim opening As String
        Dim closing As String
        opening = character
        closing = IIf(character = "{", "}", "]")
        Dim depth As Long
        depth = 1
        position = position + 1
        Do While position <= Len(json) And depth > 0
            character = Mid$(json, position, 1)
            If character = """" Then
                ignored = ParseJsonString(json, position)
            Else
                If character = opening Then depth = depth + 1
                If character = closing Then depth = depth - 1
                position = position + 1
            End If
        Loop
        If depth <> 0 Then Err.Raise vbObjectError + 7032
        SkipJsonValue = position
        Exit Function
    End If
    Do While position <= Len(json)
        character = Mid$(json, position, 1)
        If character = "," Or character = "}" Or character = "]" Then Exit Do
        position = position + 1
    Loop
    SkipJsonValue = position
End Function

Private Function SkipWhitespace(ByVal text As String, ByVal position As Long) As Long
    Do While position <= Len(text)
        Dim character As String
        character = Mid$(text, position, 1)
        If character <> " " And character <> vbTab And character <> vbCr And character <> vbLf Then Exit Do
        position = position + 1
    Loop
    SkipWhitespace = position
End Function

Private Function JsonQuote(ByVal value As String) As String
    Dim output As String
    output = """"
    Dim index As Long
    For index = 1 To Len(value)
        Dim code As Long
        code = AscW(Mid$(value, index, 1))
        Select Case code
            Case 34: output = output & "\"""
            Case 92: output = output & "\\"
            Case 8: output = output & "\b"
            Case 9: output = output & "\t"
            Case 10: output = output & "\n"
            Case 12: output = output & "\f"
            Case 13: output = output & "\r"
            Case 0 To 31: output = output & "\u" & Right$("000" & Hex$(code), 4)
            Case Else: output = output & Mid$(value, index, 1)
        End Select
    Next index
    JsonQuote = output & """"
End Function

Private Function JsonBoolean(ByVal value As Boolean) As String
    JsonBoolean = IIf(value, "true", "false")
End Function

Private Function SuccessResponse(ByVal requestId As String, ByVal resultJson As String) As String
    SuccessResponse = "{""version"":" & CStr(PROTOCOL_VERSION) & _
        ",""requestId"":" & JsonQuote(requestId) & _
        ",""success"":true,""result"":" & resultJson & ",""error"":null}"
End Function

Private Function ErrorResponse( _
    ByVal requestId As String, _
    ByVal category As String, _
    ByVal code As String, _
    ByVal message As String) As String
    ErrorResponse = "{""version"":" & CStr(PROTOCOL_VERSION) & _
        ",""requestId"":" & JsonQuote(requestId) & _
        ",""success"":false,""result"":null,""error"":{" & _
        """category"":" & JsonQuote(category) & _
        ",""code"":" & JsonQuote(code) & _
        ",""message"":" & JsonQuote(message) & "}}"
End Function

Private Sub ClassifyError( _
    ByVal number As Long, _
    ByVal description As String, _
    ByRef category As String, _
    ByRef code As String)
    If number = vbObjectError + 7032 Then
        category = "InvalidInput"
        code = "invalid_json"
        Exit Sub
    End If
    code = description
    Select Case description
        Case "workbook_not_found", "query_not_found", "module_not_found", "query_load_not_found", _
             "worksheet_not_found", "scenario_not_found"
            category = "NotFound"
        Case "query_conflict", "module_conflict", "scenario_conflict", _
             "query_destination_conflict", "temporary_name_conflict"
            category = "Conflict"
        Case "signed_project", "locked_project", "helper_target_forbidden"
            category = "Permissions"
        Case "rollback_failed"
            category = "RecoveryRequired"
        Case "query_destination_unsupported"
            category = "PlatformNotSupported"
        Case "component_type", "invalid_request_id", "unsupported_version", _
             "unsupported_action", "invalid_properties", "duplicate_property", _
             "missing_property", "scenario_cell_count", "scenario_value_count", _
             "invalid_module_name", "invalid_query_name", "invalid_load_target", _
             "invalid_query_formula"
            category = "InvalidInput"
        Case "query_refresh_failed", "query_refresh_all_failed", _
             "query_unload_incomplete", "query_load_delete_failed"
            category = "ComInterop"
        Case Else
            category = "ComInterop"
            code = "helper_failure"
    End Select
End Sub

Private Function SafeErrorMessage(ByVal code As String) As String
    Select Case code
        Case "workbook_not_found": SafeErrorMessage = "The exact target workbook is not open."
        Case "query_not_found": SafeErrorMessage = "The exact Power Query was not found."
        Case "module_not_found": SafeErrorMessage = "The exact VBA component was not found."
        Case "worksheet_not_found": SafeErrorMessage = "The exact worksheet was not found."
        Case "scenario_not_found": SafeErrorMessage = "The exact scenario was not found."
        Case "query_load_not_found": SafeErrorMessage = "No exact worksheet load was found for the Power Query."
        Case "query_conflict": SafeErrorMessage = "A Power Query with that exact name already exists."
        Case "query_destination_conflict": SafeErrorMessage = "The exact Power Query load destination is already occupied."
        Case "query_destination_unsupported": SafeErrorMessage = "That Power Query destination is not supported by the verified helper tier."
        Case "query_refresh_failed": SafeErrorMessage = "Excel could not complete the synchronous Power Query refresh."
        Case "query_refresh_all_failed"
            SafeErrorMessage = "One or more Power Queries did not complete synchronous refresh."
            If Len(mSafeErrorDetail) > 0 Then
                SafeErrorMessage = SafeErrorMessage & " Failures: " & mSafeErrorDetail
            End If
        Case "query_unload_incomplete": SafeErrorMessage = "Excel did not remove every exact Power Query worksheet load."
        Case "query_load_delete_failed": SafeErrorMessage = "Excel could not remove an exact Power Query load artifact."
        Case "invalid_load_target": SafeErrorMessage = "The Power Query worksheet destination is invalid."
        Case "invalid_query_formula": SafeErrorMessage = "Power Query M code is required."
        Case "temporary_name_conflict": SafeErrorMessage = "A temporary Power Query evaluation name is already in use."
        Case "module_conflict": SafeErrorMessage = "A VBA component with that exact name already exists."
        Case "scenario_conflict": SafeErrorMessage = "A scenario with that name already exists."
        Case "scenario_cell_count": SafeErrorMessage = "Scenario changing cells must contain between 1 and 32 cells."
        Case "scenario_value_count": SafeErrorMessage = "Scenario values must match the changing-cell count."
        Case "invalid_module_name": SafeErrorMessage = "The VBA module name is not a valid standard-module identifier."
        Case "invalid_query_name": SafeErrorMessage = "The Power Query name must not be empty."
        Case "invalid_json": SafeErrorMessage = "The helper request contains invalid JSON."
        Case "helper_target_forbidden": SafeErrorMessage = "The helper add-in cannot be used as an operation target."
        Case "rollback_failed": SafeErrorMessage = "Workbook state could not be restored after a failed mutation; close without saving."
        Case "signed_project": SafeErrorMessage = "The VBA project is signed and was not modified."
        Case "locked_project": SafeErrorMessage = "The VBA project is protected and was not modified."
        Case "component_type": SafeErrorMessage = "Only standard VBA modules can be updated or deleted."
        Case "unsupported_action": SafeErrorMessage = "The requested helper action is not allowed."
        Case "unsupported_version": SafeErrorMessage = "The helper protocol version is not supported."
        Case Else: SafeErrorMessage = "The helper rejected the request or Excel could not complete it."
    End Select
End Function

Private Sub AppendRefreshFailure( _
    ByRef details As String, _
    ByVal queryName As String, _
    ByVal failureCode As String)
    Select Case failureCode
        Case "query_load_not_found", "query_destination_unsupported", "query_refresh_failed"
        Case Else
            failureCode = "helper_failure"
    End Select

    Dim entry As String
    entry = queryName & ":" & failureCode
    If Len(details) > 0 Then entry = ", " & entry
    Dim remaining As Long
    remaining = MAX_SAFE_ERROR_DETAIL_CHARS - Len(details)
    If remaining <= 0 Then Exit Sub
    If Len(entry) > remaining Then entry = Left$(entry, remaining)
    details = details & entry
End Sub

Private Function IsCanonicalRequestId(ByVal value As String) As Boolean
    If Len(value) <> 32 Then Exit Function
    Dim index As Long
    For index = 1 To Len(value)
        Dim character As String
        character = Mid$(value, index, 1)
        If InStr(1, "0123456789abcdef", character, vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsCanonicalRequestId = True
End Function

Private Function IsHex4(ByVal value As String) As Boolean
    If Len(value) <> 4 Then Exit Function
    Dim index As Long
    For index = 1 To 4
        If InStr(1, "0123456789abcdefABCDEF", Mid$(value, index, 1), vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsHex4 = True
End Function

Private Function Utf8ByteCount(ByVal value As String) As Long
    Dim count As Long
    Dim index As Long
    For index = 1 To Len(value)
        Dim code As Long
        code = AscW(Mid$(value, index, 1))
        If code < 0 Then code = code + 65536
        If code <= &H7F Then
            count = count + 1
        ElseIf code <= &H7FF Then
            count = count + 2
        ElseIf code >= &HD800 And code <= &HDBFF And index < Len(value) Then
            Dim nextCode As Long
            nextCode = AscW(Mid$(value, index + 1, 1))
            If nextCode < 0 Then nextCode = nextCode + 65536
            If nextCode >= &HDC00 And nextCode <= &HDFFF Then
                count = count + 4
                index = index + 1
            Else
                count = count + 3
            End If
        Else
            count = count + 3
        End If
    Next index
    Utf8ByteCount = count
End Function
