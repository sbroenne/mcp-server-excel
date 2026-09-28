Attribute VB_Name = "ExcelMcpHelper"
Option Explicit

Private Const HELPER_VERSION As String = "1.0.1"
Private Const PROTOCOL_VERSION As Long = 1
Private Const MAX_PAYLOAD_BYTES As Long = 262144
Private Const STANDARD_MODULE_TYPE As Long = 1
Private Const PROJECT_LOCKED As Long = 1

Public Function ExcelMcpDispatch(ByVal requestJson As String) As String
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
        Case "powerquery.list"
            ValidateArgumentKeys argumentsJson, ""
            resultJson = PowerQueryList(target)
        Case "powerquery.view"
            ValidateArgumentKeys argumentsJson, "name"
            resultJson = PowerQueryView(target, JsonRequiredString(argumentsJson, "name"))
        Case "powerquery.create"
            ValidateArgumentKeys argumentsJson, "name,formula"
            resultJson = PowerQueryCreate(target, _
                JsonRequiredString(argumentsJson, "name"), _
                JsonRequiredString(argumentsJson, "formula"))
        Case "powerquery.update"
            ValidateArgumentKeys argumentsJson, "name,formula"
            resultJson = PowerQueryUpdate(target, _
                JsonRequiredString(argumentsJson, "name"), _
                JsonRequiredString(argumentsJson, "formula"))
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
            """helper.capabilities""," & _
            """powerquery.list"",""powerquery.view"",""powerquery.create""," & _
            """powerquery.update"",""powerquery.rename"",""powerquery.delete""," & _
            """analysis.create-scenario"",""analysis.show-scenario""," & _
            """vba.list"",""vba.view"",""vba.import"",""vba.update"",""vba.delete""],"
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
            """vbaMutation"":false}}"
    HelperCapabilities = output
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
    ByVal mCode As String) As String
    queryName = Trim$(queryName)
    If Len(queryName) = 0 Then
        Err.Raise vbObjectError + 7015, "ExcelMcpHelper", "invalid_query_name"
    End If
    If QueryExists(target, queryName) Then
        Err.Raise vbObjectError + 7010, "ExcelMcpHelper", "query_conflict"
    End If
    target.Queries.Add Name:=queryName, Formula:=mCode
    PowerQueryCreate = "{}"
End Function

Private Function PowerQueryUpdate( _
    ByVal target As Workbook, _
    ByVal queryName As String, _
    ByVal mCode As String) As String
    Dim query As Object
    Set query = QueryByExactName(target, queryName)
    query.Formula = mCode
    PowerQueryUpdate = "{}"
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
    If deleteConnection Then DeleteExactQueryLoads target, CStr(query.Name)
    query.Delete
    PowerQueryDelete = "{}"
End Function

Private Sub DeleteExactQueryLoads(ByVal target As Workbook, ByVal queryName As String)
    Dim expectedConnectionName As String
    expectedConnectionName = "Query - " & queryName
    Dim sheet As Worksheet
    For Each sheet In target.Worksheets
        Dim index As Long
        For index = sheet.ListObjects.Count To 1 Step -1
            On Error Resume Next
            Dim connectionName As String
            connectionName = CStr(sheet.ListObjects(index).QueryTable.WorkbookConnection.Name)
            If Err.Number = 0 And _
               StrComp(connectionName, expectedConnectionName, vbTextCompare) = 0 Then
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
        If StrComp(CStr(target.Connections(index).Name), expectedConnectionName, vbTextCompare) = 0 Then
            target.Connections(index).Delete
        End If
    Next index
End Sub

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
            ",""lineCount"":" & CStr(CLng(component.CodeModule.CountOfLines)) & "}"
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
        ",""source"":" & JsonQuote(source) & "}"
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
        Case "query_conflict", "module_conflict", "scenario_conflict"
            category = "Conflict"
        Case "signed_project", "locked_project", "helper_target_forbidden"
            category = "Permissions"
        Case "rollback_failed"
            category = "RecoveryRequired"
        Case "component_type", "invalid_request_id", "unsupported_version", _
             "unsupported_action", "invalid_properties", "duplicate_property", _
             "missing_property", "scenario_cell_count", "scenario_value_count", _
             "invalid_module_name", "invalid_query_name"
            category = "InvalidInput"
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
        Case "module_conflict": SafeErrorMessage = "A VBA component with that exact name already exists."
        Case "scenario_conflict": SafeErrorMessage = "A scenario with that name already exists."
        Case "scenario_cell_count": SafeErrorMessage = "Scenario changing cells must contain between 1 and 32 cells."
        Case "scenario_value_count": SafeErrorMessage = "Scenario values must match the changing-cell count."
        Case "invalid_module_name": SafeErrorMessage = "The VBA module name is not a valid standard-module identifier."
        Case "invalid_query_name": SafeErrorMessage = "The Power Query name must not be empty."
        Case "invalid_json": SafeErrorMessage = "The helper request contains invalid JSON."
        Case "helper_target_forbidden": SafeErrorMessage = "The helper add-in cannot be used as an operation target."
        Case "rollback_failed": SafeErrorMessage = "The VBA project could not be restored after a failed mutation; close without saving."
        Case "signed_project": SafeErrorMessage = "The VBA project is signed and was not modified."
        Case "locked_project": SafeErrorMessage = "The VBA project is protected and was not modified."
        Case "component_type": SafeErrorMessage = "Only standard VBA modules can be updated or deleted."
        Case "unsupported_action": SafeErrorMessage = "The requested helper action is not allowed."
        Case "unsupported_version": SafeErrorMessage = "The helper protocol version is not supported."
        Case Else: SafeErrorMessage = "The helper rejected the request or Excel could not complete it."
    End Select
End Function

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
