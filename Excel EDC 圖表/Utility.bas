Attribute VB_Name = "Utility"
Option Explicit

Public goFSO            As New FileSystemObject '���B�z��
Private Const msTitle   As String = "Module1."

Public Sub LogMsg(sMsg As String)
'����Log��
On Error GoTo Err

    Dim sFileDir As String
    Dim sFileName As String
    Dim sTitle As String
    
    sTitle = msTitle & "LogMsg"
    sFileDir = App.Path & "\LogMsg\"
    
    If Not goFSO.FolderExists(sFileDir) Then MkDir sFileDir
    
    sFileName = App.EXEName & Formate$(Now, "yyyymmdd") & ".log"
    
    Open sFileDir & sFileName For Append As #1
        Print #1, "[" & Format$(Now, "yyyy/mm/dd hh:mm:ss") & "]" & sMsg
    Close #1
    
    Debug.Print sMsg
    
    Exit Sub
Err:
    MsgBox Err.Number & "-" & Err.Description, , sTitle
End Sub

Public Function PDBDataBase(sSQL As String) As Collection
'��Ʈw����
End Function

Public Function Q(str As String) As String
    Q = "'" & str & "'"
End Function
