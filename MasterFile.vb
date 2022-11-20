Sub RunRoutine()
    'turn off screen Updating and alerts
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Merge_CSV_Files
    Format_TCSS_Report

    'turn on screen Updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub Format_TCSS_Report()
    Dim LastRow As Long
    Dim LastCol As Long
    Dim NewCol As Long
    Dim data_sheet As Worksheet
    Set data_sheet = ThisWorkbook.Worksheets("TCSS Report")
    
    'Find last row
    LastRow = data_sheet.UsedRange.Rows.Count
          
    'Insert New Columns
    data_sheet.Range("A1").EntireColumn.Insert
    data_sheet.Range("A1").Value = "SPWO"
    data_sheet.Range("C1").EntireColumn.Insert
    data_sheet.Range("C1").Value = "IMC"
    data_sheet.Range("G1").EntireColumn.Insert
    data_sheet.Range("G1").Value = "State"
    
    'Add formulas
    data_sheet.Range("A2:A" & LastRow).Formula = "=D2&""/""&TEXT(E2,""000"")"
    data_sheet.Range("C2:C" & LastRow).Formula = "=RIGHT(LEFT(D2,5),3)"
    data_sheet.Range("G2:G" & LastRow).Formula = "=MID(H2,FIND(""("",H2)+1,FIND("")"",H2)-FIND(""("",H2)-1)"
    
    'Find last column and specify next column
    LastCol = 21
    NewCol = LastCol + 1
    
    'Fix up Requirement Date Column
    For x = 2 To LastRow
        data_sheet.Range(Cells(x, 14), Cells(x, 14)) = CDate(Replace(data_sheet.Range(Cells(x, 14), Cells(x, 14)), "Quote Required By ", ""))
    Next
    
    'Add additional columns
    data_sheet.Cells(1, NewCol) = "MAXPC"
    data_sheet.Range(Cells(2, NewCol), Cells(LastRow, NewCol)).Formula = "=MAX(O2:S2)"
    NewCol = NewCol + 1
    
    data_sheet.Cells(1, NewCol) = "TTD"
    data_sheet.Range(Cells(2, NewCol), Cells(LastRow, NewCol)).Formula = "=N2-AA2"
    NewCol = NewCol + 1
    
    data_sheet.Cells(1, NewCol) = "TBD"
    data_sheet.Range(Cells(2, NewCol), Cells(LastRow, NewCol)).Formula = "=V2-N2"
    NewCol = NewCol + 1
    
    data_sheet.Cells(1, NewCol) = "CIMS Status"
    data_sheet.Range(Cells(2, NewCol), Cells(LastRow, NewCol)).Formula = "=XLOOKUP(A2,'CIMS ML'!A:A,'CIMS ML'!C:C)"
    NewCol = NewCol + 1
    
    data_sheet.Cells(1, NewCol) = "Designer"
    data_sheet.Range(Cells(2, NewCol), Cells(LastRow, NewCol)).Formula = "=XLOOKUP(A2,'CIMS ML'!A:A,'CIMS ML'!D:D)"
    NewCol = NewCol + 1
    
    data_sheet.Cells(1, NewCol) = "RegDate"
    data_sheet.Range(Cells(2, NewCol), Cells(LastRow, NewCol)).Formula = "=XLOOKUP(A2,'CIMS ML'!A:A,'CIMS ML'!B:B)"
    NewCol = NewCol + 1
    
    data_sheet.Cells(1, NewCol) = "WorkType"
    data_sheet.Range(Cells(2, NewCol), Cells(LastRow, NewCol)).Formula = "=XLOOKUP(A2,'CIMS ML'!A:A,'CIMS ML'!E:E)"
    NewCol = NewCol + 1
    
    'Format Columns
    data_sheet.Columns(23).NumberFormat = "#,##0"
    data_sheet.Columns(24).NumberFormat = "#,##0"
    data_sheet.Columns(27).NumberFormat = "dd/mm/yyyy"
    
    'Hide Columns4,5,6,12,13,16,17,18,19,20
    data_sheet.Columns(4).Hidden = True
    data_sheet.Columns(5).Hidden = True
    data_sheet.Columns(6).Hidden = True
    data_sheet.Columns(12).Hidden = True
    data_sheet.Columns(13).Hidden = True
    data_sheet.Columns(16).Hidden = True
    data_sheet.Columns(17).Hidden = True
    data_sheet.Columns(18).Hidden = True
    data_sheet.Columns(19).Hidden = True
    data_sheet.Columns(20).Hidden = True
           
    'Colour Columns
    'Green (10092441) Columns A C G H K N
    Columns("A").Interior.Color = 10092441
    Columns("C").Interior.Color = 10092441
    Columns("G").Interior.Color = 10092441
    Columns("H").Interior.Color = 10092441
    Columns("K").Interior.Color = 10092441
    Columns("N").Interior.Color = 10092441
    'Pink(16764159) Columns Y Z AA AB
    Columns("Y").Interior.Color = 16764159
    Columns("Z").Interior.Color = 16764159
    Columns("AA").Interior.Color = 16764159
    Columns("AB").Interior.Color = 16764159
    
    'Colour Cells by State
    'QLD (16737945)
    'WA/SA (13434879)
    'VIC(5287936)
    'TAS(3506772)
    'NSW(15773696)
    'NT (16764159)
    On Error Resume Next
    For x = 2 To LastRow
        If data_sheet.Range(Cells(x, 7), Cells(x, 7)).Value = "NSW" Then
            data_sheet.Range(Cells(x, 7), Cells(x, 7)).Interior.Color = 15773696
        ElseIf data_sheet.Range(Cells(x, 7), Cells(x, 7)).Value = "NT" Then
            data_sheet.Range(Cells(x, 7), Cells(x, 7)).Interior.Color = 16764159
        ElseIf data_sheet.Range(Cells(x, 7), Cells(x, 7)).Value = "QLD" Then
            data_sheet.Range(Cells(x, 7), Cells(x, 7)).Interior.Color = 16737945
        ElseIf data_sheet.Range(Cells(x, 7), Cells(x, 7)).Value = "SA" Then
            data_sheet.Range(Cells(x, 7), Cells(x, 7)).Interior.Color = 13434879
        ElseIf data_sheet.Range(Cells(x, 7), Cells(x, 7)).Value = "TAS" Then
            data_sheet.Range(Cells(x, 7), Cells(x, 7)).Interior.Color = 3506772
        ElseIf data_sheet.Range(Cells(x, 7), Cells(x, 7)).Value = "VIC" Then
            data_sheet.Range(Cells(x, 7), Cells(x, 7)).Interior.Color = 5287936
        ElseIf data_sheet.Range(Cells(x, 7), Cells(x, 7)).Value = "WA" Then
            data_sheet.Range(Cells(x, 7), Cells(x, 7)).Interior.Color = 13434879
        End If
    Next
    On Error GoTo 0
    
    'Set Column sizes of Dates 14, 15, 21, 22, 27
    data_sheet.Columns(1).ColumnWidth = 13
    data_sheet.Columns(2).ColumnWidth = 8.5
    data_sheet.Columns(3).ColumnWidth = 3.5
    data_sheet.Columns(8).ColumnWidth = 35
    data_sheet.Columns(14).ColumnWidth = 12
    data_sheet.Columns(15).ColumnWidth = 12
    data_sheet.Columns(21).ColumnWidth = 12
    data_sheet.Columns(22).ColumnWidth = 12
    data_sheet.Columns(27).ColumnWidth = 12
    data_sheet.Columns(23).ColumnWidth = 12

    'Filter by released, returned and request for resubmit
    data_sheet.Range(Cells(1, 1), Cells(data_sheet.UsedRange.Rows.Count, data_sheet.UsedRange.Columns.Count)).AutoFilter Field:=11, Criteria1:=Array("Released", "Returned", "Request to Re-submit"), Operator:=xlFilterValues

End Sub

'Merge all CSV files that are in the same
Private Sub Merge_CSV_Files()
   
    
    Dim target_workbook As Workbook
    Dim data_sheet As Worksheet
    Dim folder_path As String, my_file As String
    Dim LastRow As Long
    Dim FirstBook As Integer
    Set data_sheet = ThisWorkbook.Worksheets("TCSS Report")
    
    folder_path = GetWorkbookPath & "\"
    
    my_file = Dir(folder_path & "*.csv")
    
    '// Step 1: Clear worksheet
    
    If data_sheet.FilterMode Then data_sheet.ShowAllData
    data_sheet.Cells.EntireColumn.Hidden = False
    
    If my_file = vbNullString Then
        MsgBox "CSV files not found.", vbInformation
    Else:
        data_sheet.Cells.Clear
    End If
    
    FirstBook = 0
    
    '// Step 2: Iterate CSV Files
    Do While my_file <> vbNullString
        Set target_workbook = Workbooks.Open(folder_path & my_file, Local:=True)
            
        LastRow = data_sheet.Cells(Rows.Count, "A").End(xlUp).Row
        
        If FirstBook > 0 Then
            target_workbook.Worksheets(1).Rows(1).Delete
        End If
        
        target_workbook.Worksheets(1).UsedRange.Copy
        data_sheet.Cells(LastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
        target_workbook.Close False
        
        Set target_workbook = Nothing
        FirstBook = FirstBook + 1
        my_file = Dir()
    Loop

    '// Step 3: Clean up
    data_sheet.Rows(1).Delete
    data_sheet.UsedRange.WrapText = False
    
    'Set format for date columns 12, 13, 14, 15, 16, 18
    data_sheet.Columns(12).NumberFormat = "dd/mm/yyyy"
    data_sheet.Columns(13).NumberFormat = "dd/mm/yyyy"
    data_sheet.Columns(14).NumberFormat = "dd/mm/yyyy"
    data_sheet.Columns(15).NumberFormat = "dd/mm/yyyy"
    data_sheet.Columns(16).NumberFormat = "dd/mm/yyyy"
    data_sheet.Columns(18).NumberFormat = "dd/mm/yyyy"
    
    Set data_sheet = Nothing
    

    
End Sub

Function GetWorkbookPath(Optional wb As Workbook)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Purpose:  Returns a workbook's physical path, even when they are saved in
    '           synced OneDrive Personal, OneDrive Business or Microsoft Teams folders.
    '           If no value is provided for wb, it's set to ThisWorkbook object instead.
    ' Author:   Ricardo Gerbaudo
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    GetWorkbookPath = wb.Path
    
    If InStr(1, wb.Path, "https://") <> 0 Then
        
        Const HKEY_CURRENT_USER = &H80000001
        Dim objRegistryProvider As Object
        Dim strRegistryPath As String
        Dim arrSubKeys()
        Dim strSubKey As Variant
        Dim strUrlNamespace As String
        Dim strMountPoint As String
        Dim strLocalPath As String
        Dim strRemainderPath As String
        Dim strLibraryType As String
    
        Set objRegistryProvider = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
        strRegistryPath = "SOFTWARE\SyncEngines\Providers\OneDrive"
        objRegistryProvider.EnumKey HKEY_CURRENT_USER, strRegistryPath, arrSubKeys
        
        For Each strSubKey In arrSubKeys
            objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "UrlNamespace", strUrlNamespace
            If InStr(1, wb.Path, strUrlNamespace) <> 0 Or InStr(1, strUrlNamespace, wb.Path) <> 0 Then
                objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "MountPoint", strMountPoint
                objRegistryProvider.GetStringValue HKEY_CURRENT_USER, strRegistryPath & "\" & strSubKey & "\", "LibraryType", strLibraryType
                
                If InStr(1, wb.Path, strUrlNamespace) <> 0 Then
                    strRemainderPath = Replace(wb.Path, strUrlNamespace, vbNullString)
                Else
                    GetWorkbookPath = strMountPoint
                    Exit Function
                End If
                
                'If OneDrive Personal, skips the GUID part of the URL to match with physical path
                If InStr(1, strUrlNamespace, "https://d.docs.live.net") <> 0 Then
                    If InStr(2, strRemainderPath, "/") = 0 Then
                        strRemainderPath = vbNullString
                    Else
                        strRemainderPath = Mid(strRemainderPath, InStr(2, strRemainderPath, "/"))
                    End If
                End If
                
                'If OneDrive Business, adds extra slash at the start of string to match the pattern
                strRemainderPath = IIf(InStr(1, strUrlNamespace, "my.sharepoint.com") <> 0, "/", vbNullString) & strRemainderPath
                
                strLocalPath = ""
                
                If (InStr(1, strRemainderPath, "/")) <> 0 Then
                    strLocalPath = Mid(strRemainderPath, InStr(1, strRemainderPath, "/"))
                    strLocalPath = Replace(strLocalPath, "/", "\")
                End If
                
                strLocalPath = strMountPoint & strLocalPath
                GetWorkbookPath = strLocalPath
                If Dir(GetWorkbookPath & "\" & wb.Name) <> "" Then Exit Function
            End If
        Next
    End If
    
End Function





