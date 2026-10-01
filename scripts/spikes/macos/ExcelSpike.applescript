use framework "Foundation"
use scripting additions

on jsonResult(keys, values)
    set payload to current application's NSDictionary's dictionaryWithObjects:values forKeys:keys
    set {jsonData, jsonError} to current application's NSJSONSerialization's dataWithJSONObject:payload options:0 |error|:(reference)
    if jsonData is missing value then error (jsonError's localizedDescription() as text)
    return (current application's NSString's alloc()'s initWithData:jsonData encoding:(current application's NSUTF8StringEncoding)) as text
end jsonResult

on findWorkbook(filePath)
    set hfsPath to (POSIX file filePath) as text
    tell application "Microsoft Excel"
        repeat with workbookIndex from 1 to count of workbooks
            set candidate to workbook workbookIndex
            set candidatePath to full name of candidate
            if candidatePath is filePath or candidatePath is hfsPath then return candidate
        end repeat
    end tell
    error "Spike workbook is not open." number 9001
end findWorkbook

on populateWorkbook(w)
    tell application "Microsoft Excel"
        set s to worksheet 1 of w
        set name of s to "Spike Data"
        set value of range "A1:B3" of s to {{"Item", "Amount"}, {"Alpha", 10}, {"Beta", 20}}
        set formula of range "C1" of s to "=SUM(B2:B3)"
        set number format of range "B2:B3" of s to "0.00"
        set value of range "D1" of s to "quote \" slash \\ newline" & linefeed & character id 937
        calculate range "C1" of s
    end tell
end populateWorkbook

on run argv
    return my performProbe(argv)
end run

on invokeProbe(actionName, filePath)
    return my performProbe({actionName, filePath})
end invokeProbe

on performProbe(argv)
    if (count argv) is not 2 then error "Expected action and disposable workbook path." number 9002
    set actionName to item 1 of argv
    set filePath to item 2 of argv
    with timeout of 30 seconds
        tell application "Microsoft Excel"
            if actionName is "version" then
                return my jsonResult({"version"}, {version})
            else if actionName is "create" then
                set w to make new workbook
                set stage to "populate"
                try
                    my populateWorkbook(w)
                    set stage to "save"
                    save workbook as w filename filePath file format Excel XML file format
                on error errorText number errorNumber
                    try
                        close w saving no
                    on error cleanupText number cleanupNumber
                        error errorText & "; workbook cleanup failed: " & cleanupText number cleanupNumber
                    end try
                    error stage & ": " & errorText number errorNumber
                end try
                return my jsonResult({"success", "errorMessage"}, {true, ""})
            else if actionName is "open" then
                open workbook workbook file name filePath update links do not update links read only false add to mru false
                set w to my findWorkbook(filePath)
                return my jsonResult({"success", "errorMessage"}, {true, ""})
            else if actionName is "wait-open" then
                repeat 100 times
                    try
                        set w to my findWorkbook(filePath)
                        return my jsonResult({"success", "errorMessage"}, {true, ""})
                    on error errorText number errorNumber
                        if errorNumber is not 9001 then error errorText number errorNumber
                        delay 0.1
                    end try
                end repeat
                error "LaunchServices did not open the fixture within ten seconds." number 9007
            else if actionName is "prepare-existing" then
                set w to my findWorkbook(filePath)
                my populateWorkbook(w)
                close w saving yes
                return my jsonResult({"success", "errorMessage"}, {true, ""})
            else if actionName is "cleanup" then
                set hfsPath to (POSIX file filePath) as text
                repeat with workbookIndex from (count of workbooks) to 1 by -1
                    set candidate to workbook workbookIndex
                    set candidatePath to full name of candidate
                    if candidatePath is filePath or candidatePath is hfsPath then
                        close candidate saving no
                        exit repeat
                    end if
                end repeat
                try
                    set remainingWorkbook to my findWorkbook(filePath)
                on error errorText number errorNumber
                    if errorNumber is not 9001 then error errorText number errorNumber
                    return my jsonResult({"success", "errorMessage"}, {true, ""})
                end try
                error "Spike workbook remained open after cleanup." number 9005
            end if

            set w to my findWorkbook(filePath)
            set s to worksheet "Spike Data" of w
            if actionName is "read" then
                set matrix to value of range "A1:B3" of s
                set scalarMatrix to {{value of range "C1" of s}}
                set formulaMatrix to {{formula of range "C1" of s}}
                set specialText to value of range "D1" of s
                set cellFormat to number format of range "B2" of s
                return my jsonResult({"values", "calculated", "formulas", "text", "numberFormat"}, {matrix, scalarMatrix, formulaMatrix, specialText, cellFormat})
            else if actionName is "bulk" then
                set matrix to {}
                repeat with rowIndex from 1 to 1000
                    set rowValues to {}
                    repeat with columnIndex from 1 to 10
                        set end of rowValues to rowIndex * columnIndex
                    end repeat
                    set end of matrix to rowValues
                end repeat
                set value of range "F2:O1001" of s to matrix
                set formula of range "P1" of s to "=SUM(O2:O1001)"
                calculate range "P1" of s
                set actualMatrix to value of range "F2:O1001" of s
                set calculatedTotal to value of range "P1" of s
                return my jsonResult({"values", "total"}, {actualMatrix, calculatedTotal})
            else if actionName is "missing-sheet" then
                if not (exists worksheet "Missing Spike Sheet" of w) then
                    error "Worksheet does not exist." number 9006
                end if
                set unused to value of range "A1" of worksheet "Missing Spike Sheet" of w
                error "Missing worksheet unexpectedly succeeded." number 9003
            else if actionName is "discard" then
                set value of range "B2" of s to 999
                close w saving no
                return my jsonResult({"success", "errorMessage"}, {true, ""})
            else if actionName is "close" then
                close w saving no
                return my jsonResult({"success", "errorMessage"}, {true, ""})
            end if
            error "Unknown spike action." number 9004
        end tell
    end timeout
end performProbe
