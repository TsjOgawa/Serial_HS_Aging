'******************************************************
' ドラッグ&ドロップしたフォルダ内にあるemlファイルの
' 情報をリスト化してExcelに出力するスクリプト
' 
' 2014/02/20 @kinuasa
'******************************************************
 
Option Explicit
 
Dim Args
 
Set Args = WScript.Arguments
'パラメータ数チェック
If Args.Count < 1 Then
  WScript.Echo "当スクリプトにフォルダをドラッグ&ドロップして処理を実行してください。"
  WScript.Quit
End If
 
'フォルダ判別
With CreateObject("Scripting.FileSystemObject")
  If .FolderExists(Args(0)) = False Then
    WScript.Echo "フォルダが見つかりません。" & vbCrLf & "あるいはフォルダではありません。"
    WScript.Quit
  End If
End With
 
'emlファイルの有無チェック
If IsExistsParticularFile(Args(0), "eml") = False Then
  WScript.Echo "指定したフォルダ内にemlファイルが見つかりませんでした。"
  WScript.Quit
End If
 
ListEmlFiles Args(0)
WScript.Echo "処理が終了しました。"
 
Private Sub ListEmlFiles(ByVal FolderPath)
'指定したフォルダ内のemlファイルの情報をリスト化(Excel)
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
  i = 2 '初期化
 
  '見出し
  exWs.Cells(1, 1).Value = "No."
  exWs.Cells(1, 2).Value = "ファイル名"
  exWs.Cells(1, 3).Value = "件名"
  exWs.Cells(1, 4).Value = "本文"
  exWs.Cells(1, 5).Value = "送信者"
  exWs.Cells(1, 6).Value = "宛先"
  exWs.Cells(1, 7).Value = "ＣＣ"
  exWs.Cells(1, 8).Value = "ＢＣＣ"
  exWs.Cells(1, 9).Value = "送信日時"
  exWs.Cells(1, 10).Value = "受信日時"
  exWs.Cells(1, 11).Value = "添付ファイル数"
   
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(FolderPath).Files
      Select Case LCase(.GetExtensionName(f))
        'emlファイルのみ処理
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
'emlファイルからMessage取得
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
'指定したフォルダ内に特定の拡張子のファイルがあるかを調べる
  Dim ret
  Dim f
   
  ret = False '初期化
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