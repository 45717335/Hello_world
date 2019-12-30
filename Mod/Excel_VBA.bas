Attribute VB_Name = "Excel_VBA"
Option Explicit

Attribute VB_Name = "Excel_VBA"
Option Explicit

Function open_wb(ByRef wb As Workbook, ByVal flfp As String, Optional read_only_nomacro As Boolean = False) As Boolean
    '==========================================================
    'Open File(*.xls*):  Microsoft Excel
    '==========================================================
    open_wb = False

    Dim i As Integer
    Dim fln, flp As String
    fln = Right(flfp, Len(flfp) - InStrRev(flfp, "\"))
    flp = Left(flfp, Len(flfp) - Len(fln))
    Dim temp_b As Boolean
    temp_b = False
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = fln Then
            temp_b = True
            Set wb = Workbooks(i)
            Exit For
        End If
    Next
    If temp_b = False Then
        If Dir(flp & fln) <> "" Then

            On Error GoTo Error1:

            If read_only_nomacro = False Then
                Set wb = Workbooks.Open(flp & fln)
            Else
                Application.AutomationSecurity = msoAutomationSecurityForceDisable
                Set wb = Workbooks.Open(flp & fln, False, True)
                Application.AutomationSecurity = msoAutomationSecurityLow
            End If


            temp_b = True
        End If
    End If
    open_wb = temp_b
    Exit Function
Error1:
    MsgBox "open_wb function:" + Err.Description
    Err.Clear
    Exit Function

End Function

Function ws_exist(ByRef wb As Workbook, ByVal wsn As String) As Boolean
    '==========================================================
    'Check ws Exist
    '==========================================================
    On Error GoTo ErrorHand
    ws_exist = True
    Dim ws As Worksheet
    Set ws = wb.Worksheets(wsn)
    Exit Function
ErrorHand:
    ws_exist = False
End Function

Function get_ws(ByRef wb As Workbook, ByVal wsname As String) As Worksheet
    On Error GoTo ErrorHand
    Dim i As Integer
    Dim havewsT As Boolean
    havewsT = False
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name = wsname Then
            Set get_ws = wb.Worksheets(i)
            havewsT = True
        End If
    Next
    If havewsT = False Then
        wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count)).Name = wsname
        Set get_ws = wb.Worksheets(wsname)
    End If
    Exit Function
ErrorHand:
    If Err.Number <> 0 Then MsgBox "get_ws function: " + Err.Description
    Err.Clear
End Function




Function open_wb2(ByRef wb As Workbook, ByVal flfp As String) As Boolean
    '==========================================================
    '在新窗口中打开 workbook
    '==========================================================
    open_wb2 = False

    Dim app As Object
    Set app = CreateObject("Excel.application")
    app.Visible = True


    Dim i As Integer
    Dim fln, flp As String
    fln = Right(flfp, Len(flfp) - InStrRev(flfp, "\"))
    flp = Left(flfp, Len(flfp) - Len(fln))
    Dim temp_b As Boolean
    temp_b = False
    For i = 1 To app.Workbooks.Count
        If app.Workbooks(i).Name = fln Then
            temp_b = True
            Set wb = app.Workbooks(i)
            Exit For
        End If
    Next
    If temp_b = False Then
        If Dir(flp & fln) <> "" Then

            On Error GoTo Error1:
            Set wb = app.Workbooks.Open(flp & fln)

            temp_b = True
        End If
    End If
    open_wb2 = temp_b
    Exit Function
Error1:
    MsgBox "open_wb2 function:" + Err.Description
    Err.Clear
    Exit Function

End Function


Function Close_wb2(ByRef wb As Workbook) As Boolean
    '==========================================================
    '在新窗口中打开 workbook
    '==========================================================
    On Error GoTo ErrorHand
    Dim app As Object
    Set app = wb.Application
    If wb.Application.Workbooks.Count = 1 Then
        wb.Close
        app.Quit
        Set app = Nothing
    End If
    Exit Function
ErrorHand:
    MsgBox "Close_wb2 function:" + Err.Description
    Err.Clear
End Function



Function GetColName(ByVal intCol As Long) As String
    '列号转列名
    If InStr(CStr(Application.Version), "11") > 0 And intCol >= 1 And intCol <= 256 Then
        GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)
    ElseIf InStr(CStr(Application.Version), "12") > 0 And intCol >= 1 And intCol <= 16384 Then
        GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)

    ElseIf InStr(CStr(Application.Version), "14") > 0 And intCol >= 1 And intCol <= 16384 Then
        GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)

    Else

        GetColName = "Error"
    End If
End Function



