Attribute VB_Name = "Excel_VBA"
Option Explicit

Function open_wb(ByRef wb As Workbook, ByVal flfp As String) As Boolean
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
Set wb = Workbooks.Open(flp & fln)

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

Function add_comm(ByVal comm_s As String, ws1 As Worksheet, ByVal h_i As Integer, ByVal l_i As Integer, ByVal visiable As Boolean) As Boolean
On Error GoTo ErrorHand
If ws1.Cells(h_i, l_i).Comment Is Nothing Then
    ws1.Cells(h_i, l_i).AddComment
End If
ws1.Cells(h_i, l_i).Comment.Text Text:=comm_s
ws1.Cells(h_i, l_i).Comment.Visible = visiable
Exit Function
ErrorHand:
If Err.Number <> 0 Then MsgBox "get_ws function: " + Err.Description
Err.Clear
End Function


Function open_wb2(ByRef wb As Workbook, ByVal flfp As String) As Boolean
'==========================================================
'���´����д� workbook
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
'���´����д� workbook
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



Function get_bomlastrow(ws As Worksheet) As Integer
'��ȡbom�����һ��
On Error GoTo ErrorHand
Dim i As Integer
Dim i_lastrow As Integer
i_lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
If i_lastrow > 300 Then i_lastrow = 300
get_bomlastrow = i_lastrow
Do While Len(ws.Range("A" & get_bomlastrow) & ws.Range("B" & get_bomlastrow) & ws.Range("C" & get_bomlastrow)) = 0
get_bomlastrow = get_bomlastrow - 1
Loop
Exit Function
ErrorHand:
MsgBox "get_bomlastrow:" + Err.Description
Err.Clear

End Function
Function Str_TO_Num(in_s As String, ByRef out_i As Integer) As Boolean
'�����������ַ���ת����
On Error GoTo ErrorHand
Str_TO_Num = True
out_i = CInt(in_s)
Exit Function
ErrorHand:
Str_TO_Num = False
'MsgBox "Str_TO_Num:" + Err.Description
Err.Clear
End Function
Function Sort_BOM(ws As Worksheet, Optional start_r As Integer = 11, Optional key_col As String = "A") As Boolean
'�������������
On Error GoTo ErrorHand
Sort_BOM = False
ws.Activate

Dim i_lastrow As Integer
i_lastrow = get_bomlastrow(ws)
Dim temp_s2 As String
temp_s2 = ws.UsedRange.Address
temp_s2 = Replace(temp_s2, "$A$1", "A" & start_r)
temp_s2 = Left(temp_s2, InStrRev(temp_s2, "$"))
temp_s2 = temp_s2 & i_lastrow
Dim temp_s As String
temp_s = key_col & start_r & ":" & key_col & i_lastrow
ws.Sort.SortFields.Clear
ws.Sort.SortFields.Add Key:=Range _
        (temp_s), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ws.Sort
        .SetRange Range(temp_s2)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Sort_BOM = True
Exit Function
ErrorHand:
MsgBox "Sort_BOM function:" + Err.Description
Err.Clear
End Function




Function BOM_LIST_HEAD(ws As Worksheet) As Boolean

ws.Range("A1") = "FLFP_BOM"
ws.Range("B1") = "TKID"
ws.Columns("B:B").ColumnWidth = 17

ws.Range("C1") = "SIZE"
ws.Range("D1") = "DATE"
ws.Range("E1") = "FLN"
ws.Range("F1") = "CUSTID"
'G1,���������������ԣ������г��ֵ�BOM�����ͼֽ����Ӧ������BOM����BOM�г���
'G1=YES,G1=NO
ws.Range("G1") = "USED"

'20150609_xuefeng.gao@thyssenkrupp.com ����
'H1,��־ �ڼ�ҳ_����ҳ��
'H1=1_1
'H1=1_2,H1=2_2
'H1=1_3,H1=2_3,H1=3_3
'H1=DUPLICATE
'H1=NOT_UNIQUE
ws.Range("H1") = "SHEETS_NUM"

'20150623 ����3�����ڵ�����ת���ͻ���ʽ
'�ֱ��ǣ�CUST_STATUS,CUST_FDN,CUST_FLN
ws.Range("I1") = "CUST_STATUS"
ws.Range("J1") = "CUST_FDN"
ws.Range("K1") = "CUST_FLN"


    
    
'20150624 ����һ�����ڴ�ſ�תΪ�ͻ���ʽ���м�Ʒ
ws.Range("L1") = "TRANS_INPUT"
ws.Range("M1") = "TRANS_OUTPUT"

'20150710,����N������Ҫ�� CATIA2D���ļ�ȫ·��
ws.Range("N1") = "FLFP_DRAWING"

ws.Range("O1") = "OP_NUM"
ws.Range("P1") = "STATION_NAME"







End Function





Function BOM_LIST_Add(ws As Worksheet, fdn As String) As Boolean
'���������ڽ�ָ���ļ������BOM�������ָ����������
'A1=FLFP_BOM;   B1=TKID;    C1D1E1=SIZE;DATE;FLN;
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fd As Object
Set fd = fso.getfolder(fdn)
Dim fl As Object
Dim sfd As Object
Dim i As Integer
Dim cur_row As Integer
For Each fl In fd.Files
If fl.Name Like "?.?????.???.ST.??*.xls*" Then
cur_row = append_ws(ws, "A", fl.path)
ws.Range("E" & cur_row) = fl.Name
ws.Range("C" & cur_row) = fl.Size
ws.Range("B" & cur_row) = Get_FLN_TKID(fl.Name)
End If
Next fl
For Each sfd In fd.subfolders
BOM_LIST_Add ws, sfd.path
Next sfd
End Function


Function TKID_UNIQUE_CHECK(ws As Worksheet, Optional FLFP_COL As Integer = 1, Optional TKID_COL As Integer = 2, Optional SHEETS_NUM_COL As Integer = 8) As Boolean
'������������֤һ�ű��������ļ���Ωһ��
'��һ����������֤������ FLFP_COL,TKID_COL,�ı�ͷ���Ƿ�ֱ��� "FLFP*","TKID*"
'�ڶ�������������֤������ws�����Ƿ��б�ͷ��"SIZE*","DATE*","FLN*"
'�����������ա�TKID����С�����������
'���Ĳ����ֱ���д SHEETS_NUM_COL �������еĿ���
'���1��1_1,��ʾ������ĳ��TKIDֻ��һ���ļ������Ӧ
'���2��"DUPLICATE",����ĳ��TKID��Ӧ����ļ���������ļ���ǰ��SIZE��DATE��FLN��ȫ��ͬ��˵���Ǹ�����־Ϊ��DUPLICATE��
'���3����ȥ"DUPLICATE"�����һ���ļ����������1��д
'���4����ȥ"DUPLICATE"���ж���1�����ļ��������Щ�ļ�����ͬһ����Ŀ¼�£���־Ϊ"NOT_UNIQUE"
'���5����ȥ"DUPLICATE"���ж���1�����ļ��������Щ�ļ���ͬһ����Ŀ¼�£����α�־"1_2,2_2"����"1_3,2_3,3_3"����...

'��һ��
If ws.Cells(1, FLFP_COL) Like "FLFP*" And ws.Cells(1, TKID_COL) Like "TKID*" Then
Else
TKID_UNIQUE_CHECK = False
MsgBox "�޷�����Ωһ�Լ�飬��Ϊ��ָ���в����� �ļ�ȫ·�������߲�����TKID"
Exit Function
End If

'�ڶ���
Dim SIZE_COL As Integer, DATE_COL As Integer, FLN_COL As Integer
Dim i As Integer
For i = 1 To ws.UsedRange.Columns.Count
If ws.Cells(1, i) = "SIZE" Then
SIZE_COL = i
ElseIf ws.Cells(1, i) = "DATE" Then
DATE_COL = i
ElseIf ws.Cells(1, i) = "FLN" Then
FLN_COL = i
Else
End If
Next

'������
sort_ws ws, GetColName(TKID_COL) & "1", 2

'���Ĳ�
Dim j As Long, j_lastrow As Long
Dim k As Long
j_lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
For j = 2 To j_lastrow
k = 1
Do While ws.Cells(j, TKID_COL) = ws.Cells(j + k, TKID_COL)
If k <= j_lastrow Then
k = k + 1
End If
Loop

'���1
If k = 1 Then
ws.Cells(j, SHEETS_NUM_COL) = "1_1"
Else

'���2
Dim numofdup As Integer
numofdup = 0
Dim l As Integer
l = 0
Dim m As Integer
m = 0
For l = 0 To k - 1
For m = l + 1 To k
If ws.Cells(j + l, SHEETS_NUM_COL) <> "DUPLICATE" Then
'������ƣ���С������޸����ڶ�һ��˵���Ǹ���
If ws.Cells(j + l, SIZE_COL) = ws.Cells(j + m, SIZE_COL) And ws.Cells(j + l, DATE_COL) = ws.Cells(j + m, DATE_COL) And ws.Cells(j + l, FLN_COL) = ws.Cells(j + m, FLN_COL) Then
ws.Cells(j + m, SHEETS_NUM_COL) = "DUPLICATE"
numofdup = numofdup + 1
End If
'������ƣ���С������޸����ڶ�һ��˵���Ǹ���
End If
Next
Next

'�����
If 1 = k - numofdup Then
ws.Cells(j, SHEETS_NUM_COL) = "1_1"
Else

'���4:
Dim unique_b As Boolean
unique_b = True
Dim n As Integer
Dim sfdn As String
sfdn = Left(ws.Cells(j, FLFP_COL), InStrRev(ws.Cells(j, FLFP_COL), "\"))
For n = 1 To k - 1
If ws.Cells(j + n, SHEETS_NUM_COL) <> "DUPLICATE" Then
If sfdn <> Left(ws.Cells(j + n, FLFP_COL), InStrRev(ws.Cells(j + n, FLFP_COL), "\")) Then
unique_b = False
Exit For
End If
End If

Next

If unique_b = False Then
    For n = 0 To k - 1
    If ws.Cells(j + n, SHEETS_NUM_COL) <> "DUPLICATE" Then
    ws.Cells(j + n, SHEETS_NUM_COL) = "NOT_UNIQUE"
    End If
    Next
Else

'�����
    Dim total_sheet As Integer
    total_sheet = k - numofdup
    Dim cur_sheet As Integer
    cur_sheet = 0
    For n = 0 To k
    If ws.Cells(j + n, SHEETS_NUM_COL) <> "DUPLICATE" Then
    cur_sheet = cur_sheet + 1
    ws.Cells(j + n, SHEETS_NUM_COL) = cur_sheet & "_" & total_sheet
    End If
    Next
End If
End If
End If
j = j + k - 1

Next
End Function



Function sort_ws(ws As Worksheet, key_rgn As String, Optional start_row As Integer = 2)
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range(key_rgn), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        '.SetRange ws.UsedRange
        .SetRange Range(Replace(ws.UsedRange.Address, "$A$1", "$A$" & start_row))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Function



Function WS_ROW_DEL(ws As Worksheet, col_name As String, TPF_str As String, Optional row_start As Long = 2) As Boolean
'2015 06 23 �������豸ͼֽ�������ֲ����

'ɾ�� ָ���������У�ָ������ ����ָ��ͨ�������
If Len(TPF_str) <= 1 Then
Exit Function
End If


Dim i As Long
Dim i_lastrow As Long
i_lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
Dim del_row As Boolean
Dim j As Long
Dim TPF_Arr() As String
Dim ub_i As Integer
TPF_Arr = Split(TPF_str, Chr(10))
ub_i = UBound(TPF_Arr)
For i = row_start To i_lastrow
del_row = False
For j = 0 To ub_i
If InStr(ws.Range(col_name & i), TPF_Arr(j)) > 0 Then
del_row = True
Exit For
End If
Next
If del_row Then
ws.Rows(i).Delete
i = i - 1
i_lastrow = i_lastrow - 1
End If
Next
End Function

Function append_ws(ByRef ws As Worksheet, ByVal a As String, ByVal A_val) As Integer
append_ws = 0
Dim lastrow As Integer
lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
ws.Range(a & lastrow + 1) = A_val
append_ws = lastrow + 1

End Function
Function Get_FLN_TKID(fln As String) As String
Get_FLN_TKID = ""

If fln Like "?_?????_???_??_??*" Then
fln = Replace(fln, "_", ".")
End If

If fln Like "?.?????.???.??.??*" Then
Get_FLN_TKID = Left(fln, 17)
If Left(Get_FLN_TKID, 2) = "k." Then
Get_FLN_TKID = Replace(Get_FLN_TKID, "k.", "K.")
ElseIf Left(Get_FLN_TKID, 2) = "d." Then
Get_FLN_TKID = Replace(Get_FLN_TKID, "d.", "D.")
End If
End If


End Function

Function GetColName(ByVal intCol As Long) As String
'�к�ת����
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
