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

  'Access�t�@�C���̗L���m�F
  If Dir(PATH) = "" Then
    Debug.Print "�w��̃t�@�C�������݂��Ȃ�" & _
                "------------------------------------------------------\n\n" & _
                PATH & "\n\n" & _
                "------------------------------------------------------\n"
    Exit Sub
  End If
  
    
  'Access���J��
  Dim objAcs As Object: Set objAcs = CreateObject("Access.Applicaion")
  objAcs.OpenCurrentDatabase PATH
    
    
  '���O�����ĕۑ�(�t�@�C���`����ύX�A�ۑ��ꏊ)
  Dim ii As Integer
  For ii = LBound(Dest, 1) To UBound(Dest, 1)
    objAcs.CurrentDb.excute "SaveAsText acDatabase, '" & Dest(ii) & "'"
  Next ii


  'Access���I��
  objscs.CloseCurrentDatabase
  objscs.Quit
  Set objAcs = Nothing

End Sub
