Attribute VB_Name = "AccessUpdater"
Option Compare Database


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Const PATH = ""
Const Dest = Array( _
  "", _
  "", _
  "")
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/



Public Sub AccessUpdater()

  'Accessファイルの有無確認
  If Dir(PATH) = "" Then
    Debug.Print "指定のファイルが存在しない" & _
                "------------------------------------------------------\n\n" & _
                PATH & "\n\n" & _
                "------------------------------------------------------\n"
    Exit Sub
  End If
  
    
  'Accessを開く
  Dim objAcs As Object: Set objAcs = CreateObject("Access.Applicaion")
  objAcs.OpenCurrentDatabase PATH
    
    
  '名前をつけて保存(ファイル形式を変更、保存場所)
  Dim ii As Integer
  For ii = LBound(Dest, 1) To UBound(Dest, 1)
    objAcs.CurrentDb.excute "SaveAsText acDatabase, '" & Dest(ii) & "'"
  Next ii


  'Accessを終了
  objscs.CloseCurrentDatabase
  objscs.Quit
  Set objAcs = Nothing

End Sub
