'******************************************************
' �h���b�O&�h���b�v�����t�H���_���ɂ���eml�t�@�C����
' �������X�g������Excel�ɏo�͂���X�N���v�g
' 
' 2014/02/20 @kinuasa
'******************************************************
 
Option Explicit
 
Dim Args
 
Set Args = WScript.Arguments
'�p�����[�^���`�F�b�N
If Args.Count < 1 Then
  WScript.Echo "���X�N���v�g�Ƀt�H���_���h���b�O&�h���b�v���ď��������s���Ă��������B"
  WScript.Quit
End If
 
'�t�H���_����
With CreateObject("Scripting.FileSystemObject")
  If .FolderExists(Args(0)) = False Then
    WScript.Echo "�t�H���_��������܂���B" & vbCrLf & "���邢�̓t�H���_�ł͂���܂���B"
    WScript.Quit
  End If
End With
 
'eml�t�@�C���̗L���`�F�b�N
If IsExistsParticularFile(Args(0), "eml") = False Then
  WScript.Echo "�w�肵���t�H���_����eml�t�@�C����������܂���ł����B"
  WScript.Quit
End If
 
ListEmlFiles Args(0)
WScript.Echo "�������I�����܂����B"
 
Private Sub ListEmlFiles(ByVal FolderPath)
'�w�肵���t�H���_����eml�t�@�C���̏������X�g��(Excel)
  Dim exApp
  Dim exWb
  Dim exWs
  Dim msg
  Dim f
  Dim i
   
  Set exApp = CreateObject("Excel.Application")
  exApp.Visible = True
  Set exWb = exApp.Workbooks.Add
  Set exWs = exWb.Worksheets(1)
  i = 2 '������
 
  '���o��
  exWs.Cells(1, 1).Value = "No."
  exWs.Cells(1, 2).Value = "�t�@�C����"
  exWs.Cells(1, 3).Value = "����"
  exWs.Cells(1, 4).Value = "�{��"
  exWs.Cells(1, 5).Value = "���M��"
  exWs.Cells(1, 6).Value = "����"
  exWs.Cells(1, 7).Value = "�b�b"
  exWs.Cells(1, 8).Value = "�a�b�b"
  exWs.Cells(1, 9).Value = "���M����"
  exWs.Cells(1, 10).Value = "��M����"
  exWs.Cells(1, 11).Value = "�Y�t�t�@�C����"
   
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(FolderPath).Files
      Select Case LCase(.GetExtensionName(f))
        'eml�t�@�C���̂ݏ���
        Case "eml"
          Set msg = GetMessage(f.Path)
          exWs.Cells(i, 1).Value = i - 1
          exWs.Cells(i, 2).Value = f.Name
          exWs.Cells(i, 3).Value = msg.Subject
          exWs.Cells(i, 4).Value = msg.TextBody
          exWs.Cells(i, 5).Value = msg.From
          exWs.Cells(i, 6).Value = msg.To
          exWs.Cells(i, 7).Value = msg.CC
          exWs.Cells(i, 8).Value = msg.BCC
          exWs.Cells(i, 9).Value = msg.SentOn
          exWs.Cells(i, 10).Value = msg.ReceivedTime
          exWs.Cells(i, 11).Value = msg.Attachments.Count
          Set msg = Nothing
          i = i + 1
      End Select
    Next
  End With
  exWs.Range(exWs.Rows(2), exWs.Rows(i - 1)).WrapText = False
End Sub
 
Private Function GetMessage(ByVal FilePath)
'eml�t�@�C������Message�擾
  Dim stm
  Dim msg
   
  Set stm = CreateObject("ADODB.Stream")
  Set msg = CreateObject("CDO.Message")
  stm.Open
  stm.LoadFromFile FilePath
  msg.DataSource.OpenObject stm, "_Stream"
  stm.Close
  Set GetMessage = msg
End Function
 
Private Function IsExistsParticularFile(ByVal FolderPath, ByVal FileExtension)
'�w�肵���t�H���_���ɓ���̊g���q�̃t�@�C�������邩�𒲂ׂ�
  Dim ret
  Dim f
   
  ret = False '������
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(FolderPath).Files
      Select Case LCase(.GetExtensionName(f))
        Case LCase(FileExtension)
          ret = True
          Exit For
      End Select
    Next
  End With
  IsExistsParticularFile = ret
End Function