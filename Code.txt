Public Function ExportCSV()
    Dim ws As Worksheet
    Dim path As String: path = ActiveWorkbook.path & "\"
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        If (StrComp(ws.Name, "tool", vbTextCompare) = 0) Then
            GoTo ForEnd
        End If
        
        Dim csvFilename As String: csvFilename = path & "csv_" & ws.Name & ".csv"
        
       ' ws.Copy
       ' ActiveWorkbook.SaveAs fileName:=csvFilename, FileFormat:=xlCSV, CreateBackup:=False
        WriteCSV ws, csvFilename
       ' ActiveWorkbook.Close False
        
ForEnd:
    Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Function

Public Function WriteCSV(wkb As Worksheet, fileName As String)
    If IsNull(wkb) Then
        End
    End If

    If fileName = "False" Then
        End
    End If

    On Error GoTo eh
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    
    Dim fileStream
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Charset = "UTF-8"
    fileStream.Type = adTypeText
    fileStream.Open
    
    Dim lastRow As Long
    With wkb
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    For r = 1 To lastRow
        s = ""
        c = 1
        While Not IsEmpty(wkb.Cells(r, c).Value)
            s = s & wkb.Cells(r, c).Value & ","
            c = c + 1
        Wend
        fileStream.WriteText s, 1
    Next r
    
    fileStream.Position = 3
    
    Dim saveStream: Set saveStream = CreateObject("ADODB.Stream")

    With saveStream
      .Type = 1
      .Open
      fileStream.CopyTo saveStream
      .SaveToFile fileName, adSaveCreateOverWrite
    End With

    saveStream.Flush
    saveStream.Close
    
    fileStream.Flush
    fileStream.Close
    
    MsgBox "sucessfully saved " & fileName
    
eh:
   
End Function
