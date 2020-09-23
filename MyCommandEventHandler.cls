VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyCommandEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    On Error Resume Next
    handled = True
    CancelDefault = True

    Dim S As String, sa() As String
    S = CommandBarControl.OnAction
    sa = Split(S, ".")
    If UBound(sa) < 1 Then Exit Sub
    
    Dim cm As CodeModule
    Set cm = gVBE.VBProjects(sa(0)).VBComponents(sa(1)).CodeModule
    
    cm.CodePane.Show
    

'    For Each c In p.VBComponents
'        c.CodeModule.CodePane.Show
End Sub
