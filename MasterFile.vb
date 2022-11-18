Sub Merge_CSV_Files()
   'turn off screen Updating and alerts
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim target_workbook As Workbook
    Dim data_sheet As Worksheet
    Dim folder_path As String, my_file As String
    Dim LastRow As Long
    Dim FirstBook As Integer
    Set data_sheet = ThisWorkbook.Worksheets("Master")
    
    folder_path = GetWorkbookPath & "\"
    
    my_file = Dir(folder_path & "*.csv")
    
    '// Step 1: Clear worksheet
    If my_file = vbNullString Then
        MsgBox "CSV files not found.", vbInformation
    Else:
        data_sheet.Cells.ClearContents
    End If
    
    FirstBook = 0
    
    '// Step 2: Iterate CSV Files
    Do While my_file <> vbNullString
        Set target_workbook = Workbooks.Open(folder_path & my_file)
            
        LastRow = data_sheet.Cells(Rows.Count, "A").End(xlUp).Row
        
        If FirstBook > 0 Then
            target_workbook.Worksheets(1).Rows(1).Delete
        End If
        
        target_workbook.Worksheets(1).Range("A1").CurrentRegion.Copy
        data_sheet.Cells(LastRow + 1, "A").PasteSpecial Paste:=xlPasteValues
        target_workbook.Close False
        
        Set target_workbook = Nothing
        FirstBook = FirstBook + 1
        my_file = Dir()
    Loop

    '// Step 3: Clean up
    data_sheet.Rows(1).Delete
    data_sheet.UsedRange.WrapText = False
    Set data_sheet = Nothing
    
    'turn on screen Updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
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



