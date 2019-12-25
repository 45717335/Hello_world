VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Fso As Object

Public Function folderexists(fdn As String) As Boolean
folderexists = False
If Fso.folderexists(fdn) Then
folderexists = True
End If
End Function

Public Function FileExists(flfp As String) As Boolean
FileExists = False
If Fso.FileExists(flfp) Then
FileExists = True
End If
End Function

   
Public Function PickFolder() As String

   
        '** ʹ��FileDialog������ѡ���ļ���
        Dim fd As FileDialog
        Dim strPath As String
       
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
       
        '** ��ʾѡ���ļ��жԻ���
        If fd.Show = -1 Then        '** �û�ѡ�����ļ���
            strPath = fd.SelectedItems(1)
        Else
            strPath = ""
        End If
        Set fd = Nothing
       PickFolder = strPath
End Function



   
   
Private Sub Class_Initialize()
    '������󴴽�����
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
End Sub
Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set Fso = Nothing
End Sub


  
Public Function CreateFolder(ByVal strPath)
On Error Resume Next
Dim arrPath, pathNum, i, tmpPath, currPath
strPath = Replace(strPath, "\", "/")
arrPath = Split(strPath, "/")
pathNum = UBound(arrPath)
'Set fso = CreateObject("Scripting.FileSystemObject")
If Err <> 0 Then
 MsgBox "ERROR"
End If
For i = 0 To pathNum
   If i = 0 Then
    tmpPath = arrPath(0) & "/"
   Else
    tmpPath = tmpPath & arrPath(i) & "/"
   End If
   currPath = Left(tmpPath, Len(tmpPath) - 1)
   If Not Fso.folderexists(currPath) Then Fso.CreateFolder currPath
Next

End Function

Public Function Datelastmodify(flfp As String) As Date
If folderexists(flfp) Then
Datelastmodify = Fso.getfolder(flfp).DateLastModified
ElseIf FileExists(flfp) Then
Datelastmodify = Fso.getfile(flfp).DateLastModified
Else
Datelastmodify = CDate("1900-1-1")
End If
End Function

Public Function Userlastmodify(flfp As String) As String
 Userlastmodify = ""
On Error GoTo Err:

 If flfp Like "*.xls*" Then
 Dim wb As Workbook
 If open_wb(wb, flfp) Then
 
 
    
'Dim p As DocumentProperty

 '   For Each p In wb.BuiltinDocumentProperties
'
 '       If Len(p.Value) > 0 Then
'
 '           Debug.Print p.name & "_" & p.Value
'
 '       Else
'
 '           Debug.Print p.name
'
  '      End If
 '
'Next

Userlastmodify = wb.BuiltinDocumentProperties.Item("Last author").Value

 
 
 wb.Saved = True
 wb.Close
 End If
 
 End If
 Exit Function
 
Err:
 MsgBox "FSO ERR:Userlastmodify" & Err.Number
 
End Function


Public Function copy_file(ByVal flfp_from As String, ByVal flfp_to As String) As Boolean
On Error GoTo ErrorHand


copy_file = False
Dim fdn As String
Dim fln As String
Dim FDN_FROM As String

Dim fln_to As String


fdn = Left(flfp_to, InStrRev(flfp_to, "\"))
FDN_FROM = Left(flfp_from, InStrRev(flfp_from, "\"))
fln = Right(flfp_from, Len(flfp_from) - Len(FDN_FROM))

If Right(flfp_to, 1) = "\" Then
fln_to = fln
Else
fln_to = ""
End If

CreateFolder fdn

'Dim fso As Object
'Set fso = CreateObject("Scripting.FileSystemObject")
'Dim fd As Object
'Set fd = FSO.GetFolder(FDN_BOM)
          
If Fso.FileExists(flfp_from) = True Then
        Fso.copyfile flfp_from, flfp_to & fln_to, True
        copy_file = True
End If
Exit Function
ErrorHand:
If Err.Number = 70 Then
Application.Wait (Now() + TimeValue("00:00:10"))

If FileExists(flfp_to) = False Then
copy_file flfp_from, flfp_to
End If

Else
MsgBox Err.Description & "copy_file"
End If
End Function

Public Function deletefile(str1 As String) As Boolean
On Error GoTo ErrorHand
deletefile = True
Fso.deletefile str1
Exit Function
ErrorHand:
deletefile = False
End Function

'-----------Sub GetFiles----------Copyright@�ٶȲ���ȥ�ȸ� QQ80871835 2014/3/10---------------------------
'���� :'����pathĿ¼,���������ļ��������ļ���������,fullname������ѡ�Ƿ񷵻�����·��
'���� :path      string  -�ļ���·��
'      Fullname  Boolean -�Ƿ񷵻�����·��,Ĭ��Ϊtrue��������
'      ftype     string  -�����ļ������ļ���,fΪ�ļ�,foΪ�ļ���
'--------------------------------------------------------------------------------------------------
Public Function GetFiles(path$, Optional Fullname As Boolean = True, Optional ftype_f_fo = "f", Optional digui As Boolean = True)
    Dim dic As Object, i&
    Set dic = CreateObject("Scripting.Dictionary")
    Dim Fso As Object, Folder As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = Fso.getfolder(path)
    i = 1
    Call aGetFile(Folder, dic, ftype_f_fo, digui)
    If Fullname Then
        GetFiles = dic.keys '�����ļ���
    Else
        GetFiles = dic.items '��������·�����ļ���
    End If
    Set Folder = Nothing
    Set Fso = Nothing
    'MsgBox "ok"
End Function
Private Function aGetFile(ByVal Folder As Object, dic, Optional ftype = "f", Optional digui As Boolean = True)
    Dim SubFolder As Object
    Dim File As Object
    If ftype = "fo" Then
        For Each SubFolder In Folder.SubFolders
            dic.Add SubFolder.path, SubFolder.name
            If digui Then Call aGetFile(SubFolder, dic, ftype)   '�ݹ�������ļ���
        Next
    Else '�����ļ�
        For Each File In Folder.Files    '�����ļ�
        '��� file.path ���Ѿ� ���� �ļ����� �����ٴ������ļ���
        
            'dic.Add File.path & "" & File.Name, File.Name
            
            If InStr(File.path, File.name) = Len(File.path) - Len(File.name) + 1 Then
            dic.Add File.path & "", File.name
            Else
            dic.Add File.path & "" & File.name, File.name
            End If
            
            
        Next
        For Each SubFolder In Folder.SubFolders
            If digui Then Call aGetFile(SubFolder, dic, ftype)     '�ݹ�������ļ���
        Next
    End If
End Function

Function get_flndatesize(flfp As String) As String
On Error GoTo ErrorHand
Dim fobj As Object
Set fobj = Fso.getfile(flfp)
get_flndatesize = fobj.name & "_" & Format(fobj.DateLastModified, "YYYYMMDDHHMMSS") & "_" & fobj.Size
Exit Function
ErrorHand:
get_flndatesize = ""
End Function