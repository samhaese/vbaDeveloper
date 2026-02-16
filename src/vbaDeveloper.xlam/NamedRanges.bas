Attribute VB_Name = "NamedRanges"
Option Explicit

Private Const NAMED_RANGES_FILE_NAME As String = "NamedRanges.csv"

Private Enum columns
    name = 0
    RefersTo
    Comments
End Enum

Function Sanitize(s As String) As String
    Dim result As String
    result = Replace(s, "\", "\\")
    result = Replace(result, Chr(10), "\n")
    result = Replace(result, Chr(13), "\r")
    Sanitize = result
End Function

Function Desanitize(s As String) As String
    Dim result As String
    result = Replace(s, "\r", Chr(13))
    result = Replace(result, "\n", Chr(10))
    result = Replace(result, "\\", "\")
    Desanitize = result
End Function

Function qualifyText(s As String, Optional delim As String = ",", Optional qualifier As String = """") As String
    Dim result As String
    result = Replace(s, qualifier, qualifier & qualifier)
    If InStr(result, delim) <> 0 Then
        result = qualifier & result & qualifier
    End If
    qualifyText = result
End Function

Function splitTextWithQualifiers(s As String, Optional delim As String = ",", Optional qualifier As String = """") As Variant
    Dim i As Integer
    Dim coll As Collection
    Dim accumulator As String
    Dim inQuote As Boolean
    Dim char As String
    
    Set coll = New Collection
    accumulator = ""
    i = 1
    While i <= Len(s)
        char = Mid(s, i, 1)
        If char = qualifier Then
            If i < Len(s) And Mid(s, i + 1, 1) = qualifier Then
                i = i + 1
                accumulator = accumulator & qualifier
            Else
                inQuote = Not inQuote
            End If
        ElseIf char = delim Then
            If inQuote Then
                accumulator = accumulator & char
            Else
                coll.Add accumulator
                accumulator = ""
            End If
        Else
            accumulator = accumulator & char
        End If
        i = i + 1
    Wend
    coll.Add accumulator
    
    Dim result() As Variant
    ReDim result(0 To coll.Count - 1)
    For i = 1 To coll.Count
        result(i - 1) = coll.Item(i)
    Next i
    
    splitTextWithQualifiers = result
    
End Function

' Import named ranges from csv file
' Existing ranges with the same identifier will be replaced.
Public Sub importNamedRanges(wb As Workbook)
    Dim importDir As String
    importDir = Build.getSourceDir(wb.FullName, createIfNotExists:=False)
    If importDir = "" Then
        Debug.Print "No import directory for workbook " & wb.name & ", skipping"
        Exit Sub
    End If

    Dim fileName As String
    fileName = importDir & NAMED_RANGES_FILE_NAME
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FileExists(fileName) Then
        Dim inStream As TextStream
        Set inStream = FSO.OpenTextFile(fileName, ForReading, Create:=False)
        Dim line As String
        Do Until inStream.AtEndOfStream
            line = inStream.ReadLine
            importName wb, line
        Loop
        inStream.Close
    End If
End Sub


Private Sub importName(wb As Workbook, line As String)
    Dim parts As Variant
    parts = splitTextWithQualifiers(Desanitize(line), ",")
    Dim rangeName As String, rangeAddress As String, comment As String
    rangeName = parts(columns.name)
    rangeAddress = parts(columns.RefersTo)
    comment = parts(columns.Comments)

    ' Existing namedRanges don't need to be removed first.
    ' wb.Names.Add will automatically replace or add the given namedRange.
    wb.Names.Add(rangeName, rangeAddress).comment = comment
End Sub


'Export named ranges to csv file
Public Sub exportNamedRanges(wb As Workbook)
    Dim exportDir As String
    exportDir = Build.getSourceDir(wb.FullName, createIfNotExists:=True)
    Dim fileName As String
    fileName = exportDir & NAMED_RANGES_FILE_NAME

    Dim lines As Collection
    Set lines = New Collection
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If hasValidRange(aName) Then
            lines.Add Sanitize(qualifyText(aName.name) & "," & qualifyText(aName.RefersTo) & "," & qualifyText(aName.comment))
        End If
    Next
    If lines.Count > 0 Then
        'We have some names to export
        Debug.Print "writing to  " & fileName

        Dim FSO As New Scripting.FileSystemObject
        Dim outStream As TextStream
        Set outStream = FSO.CreateTextFile(fileName, overwrite:=True, unicode:=False)
        On Error GoTo closeStream
        Dim line As Variant
        For Each line In lines
            outStream.WriteLine line
        Next line
closeStream:
        outStream.Close
    End If
End Sub


Private Function hasValidRange(aName As name) As Boolean
    On Error GoTo no
    hasValidRange = False
    Dim aRange As Range
    Set aRange = aName.RefersToRange
    hasValidRange = True
no:
End Function


' Clean up all named ranges that don't refer to a valid range.
' This sub is not used by the import and export functions.
' It is provided only for convenience and can be run manually.
Public Sub removeInvalidNamedRanges(wb As Workbook)
    Dim aName As name
    Dim t As Variant
    For Each t In wb.Names
        Set aName = t
        If Not hasValidRange(aName) Then
            aName.Delete
        End If
    Next
End Sub

