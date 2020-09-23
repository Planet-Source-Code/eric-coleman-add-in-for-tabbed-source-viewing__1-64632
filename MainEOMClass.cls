VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainEOMClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "VBgamer.com Tabbed Source Viewer"
Option Explicit

'Public WithEvents PrjHandler  As VBProjectsEvents          'projects event handler
Public WithEvents CmpHandler  As VBComponentsEvents        'components event handler
Attribute CmpHandler.VB_VarHelpID = -1
'Public WithEvents CtlHandler  As VBControlsEvents          'controls event handler
'Public WithEvents FCHandler As FileControlEvents

Public WithEvents MenuHandlerLeft As CommandBarEvents
Attribute MenuHandlerLeft.VB_VarHelpID = -1
Public WithEvents MenuHandlerRight As CommandBarEvents
Attribute MenuHandlerRight.VB_VarHelpID = -1

Dim EvtHandlers As Collection


Implements IDTExtensibility


Const guidMYTOOL$ = "{4A204572-6963-2043-6F6C-656D616E2031}"


Private Function GetProjectOfComponent(ByRef vbc As VBIDE.VBComponent) As VBIDE.VBProject
    Dim p As VBProject
    Dim c As VBComponent
    Set GetProjectOfComponent = Nothing
    For Each p In gVBE.VBProjects
        For Each c In p.VBComponents
            If c Is vbc Then
                Set GetProjectOfComponent = p
                Exit Function
            End If
        Next
    Next
End Function

Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
'    Debug.Print "CmpHandler", "Activated"
End Sub
Private Sub CmpHandler_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
'
    'For Each p In gVBE.VBProjects
    '    For Each c In p.VBComponents
    Dim p As VBProject
    Set p = GetProjectOfComponent(VBComponent)
    If Not Nothing Is p Then
        AddButton p, VBComponent
    End If
End Sub
Private Sub CmpHandler_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
'
End Sub
Private Sub CmpHandler_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    Debug.Print "CmpHandler", "Removed"
    Debug.Print "Is removed ", VBComponent.Name
 '   vbcomponent.

    Dim ocbcTemp As Office.CommandBarControl
    Dim s As String
    Dim p As VBProject
    
    Set p = GetProjectOfComponent(VBComponent)
    If Not Nothing Is p Then
        s = p.Name & "." & VBComponent.Name
        For Each ocbcTemp In gVBE.CommandBars("VBgamer").Controls
            If ocbcTemp.OnAction = s Then ocbcTemp.Delete: Exit For
        Next
    End If

       
End Sub
Private Sub CmpHandler_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    RefreshToolbar
End Sub
Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
'    Debug.Print "CMPHandler", "Item Selected"
End Sub



Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'' Comment to prevent procedure from being
'' deleted on compilation.
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
'
'MsgBox "Add-in is now connected"

Set gVBE = VBInst
Set EvtHandlers = New Collection

'Dim docTabOrderObject As Object  'user doc instance
' The guidMYTOOL$ constant is the unique registry
' identifier for your add-in.


 'gVBInstance.Windows.CreateToolWindow(gVBInstance.Addins(1), "TabOrder.docTabOrder", LoadResString(10), guidMYTOOL$, gdocTabOrder)

'/////////////////////////////////////
'Set gWindow = gVBE.Windows.CreateToolWindow(gVBE.Addins("VBgamerSourceBin.MainEOMClass"), "VBgamerSourceBin.udSourceControl", "Source Control", guidMYTOOL$, gDocWindow)
'    gWindow.Visible = True
'//////////////////////
  

'Set gWindow = gVBE.Windows.CreateToolWindow(gVBE.Addins("VBgamerSourceBin.udSourceControl"), "TabOrder.docTabOrder", _
LoadResString(10), guidMYTOOL$, docTabOrderObject)

'    Set Me.PrjHandler = gVBE.Events.VBProjectsEvents
    Set Me.CmpHandler = gVBE.Events.VBComponentsEvents(Nothing)
    

    RefreshToolbar
    


'Set Me.MenuHandler = gVBE.Events.CommandBarEvents(cb)
Exit Sub
errConnection:
Debug.Print "Error", Err.Number, Err.Description
Err.Clear
Resume Next
    
End Sub

Private Sub RefreshToolbar()
On Error Resume Next

Dim cb As CommandBar
Set cb = gVBE.CommandBars("VBgamer")

If Err.Number <> 0 Then
    Set cb = gVBE.CommandBars.Add("VBgamer", 1, , True)
End If

On Error GoTo 0

cb.Visible = True
    
Dim ocbcTemp As Office.CommandBarControl
For Each ocbcTemp In gVBE.CommandBars("VBgamer").Controls
    ocbcTemp.Delete
Next
    


Set EvtHandlers = Nothing 'easy way to clear a collection.
Set EvtHandlers = New Collection


Dim mcbMenuCommandBar         As Office.CommandBarControl  'command bar object


AddControlButtons

    Dim p As VBProject
    Dim c As VBComponent
    Dim ct As VBControl

'    Dim cp As CodePane
'    Dim sc As String
'    Dim sp As String
'    Dim L As Long
Dim m As Member
    For Each p In gVBE.VBProjects
        For Each c In p.VBComponents
            'c.CodeModule.CodePane.Show
                
            AddButton p, c
            
        Next
    Next

End Sub

Public Sub AddControlButtons()
    Dim oPic As StdPicture
    Dim prp As Property
    Dim mcls As MyCommandEventHandler
    
    
    gCountControlButtons = 2  'increase this number if you want to add other buttons.

    Set oPic = LoadResPicture(110, vbResBitmap) 'unknown
    Set mcbMenuCommandBar = gVBE.CommandBars("VBgamer").Controls.Add(1)
    mcbMenuCommandBar.Caption = "Scroll Left"
    mcbMenuCommandBar.BeginGroup = False
    mcbMenuCommandBar.Style = msoButtonIcon
    mcbMenuCommandBar.ToolTipText = "Scroll Left"
    CopyBitmapAsButtonFace oPic, &HFF
    mcbMenuCommandBar.PasteFace
    Set MenuHandlerLeft = gVBE.Events.CommandBarEvents(mcbMenuCommandBar)

    Set oPic = LoadResPicture(111, vbResBitmap) 'unknown
    Set mcbMenuCommandBar = gVBE.CommandBars("VBgamer").Controls.Add(1)
    mcbMenuCommandBar.Caption = "Scroll Right"
    mcbMenuCommandBar.BeginGroup = False
    mcbMenuCommandBar.Style = msoButtonIcon
    mcbMenuCommandBar.ToolTipText = "Scroll Right"
    CopyBitmapAsButtonFace oPic, &HFF
    mcbMenuCommandBar.PasteFace
    Set MenuHandlerRight = gVBE.Events.CommandBarEvents(mcbMenuCommandBar)
End Sub

Private Sub AddButton(ByRef p As VBProject, ByRef c As VBComponent)

    Dim blnSkip As Boolean
    Dim oPic As StdPicture
    Dim prp As Property
    Dim mcls As MyCommandEventHandler
    
            Select Case c.Type
                Case vbext_ComponentType.vbext_ct_ActiveXDesigner
                    Set oPic = LoadResPicture(109, vbResBitmap) 'unknown
                Case vbext_ComponentType.vbext_ct_ClassModule
                    Set oPic = LoadResPicture(101, vbResBitmap)
                Case vbext_ComponentType.vbext_ct_DocObject
                    Set oPic = LoadResPicture(108, vbResBitmap)
                Case vbext_ComponentType.vbext_ct_MSForm
                    Set oPic = LoadResPicture(109, vbResBitmap) 'unknown
                Case vbext_ComponentType.vbext_ct_PropPage
                    Set oPic = LoadResPicture(106, vbResBitmap)
                Case vbext_ComponentType.vbext_ct_RelatedDocument   'nothing
                    blnSkip = True
                Case vbext_ComponentType.vbext_ct_ResFile           'nothing.
                    blnSkip = True
                Case vbext_ComponentType.vbext_ct_StdModule
                    Set oPic = LoadResPicture(105, vbResBitmap)
                Case vbext_ComponentType.vbext_ct_UserControl
                    Set oPic = LoadResPicture(107, vbResBitmap)
                Case vbext_ComponentType.vbext_ct_VBForm
                    Set prp = c.Properties("MDIChild")
                    If prp.Value = True Then
                        Set oPic = LoadResPicture(103, vbResBitmap)
                    Else 'normal form.
                        Set oPic = LoadResPicture(102, vbResBitmap)
                    End If
                Case vbext_ComponentType.vbext_ct_VBMDIForm
                    Set oPic = LoadResPicture(104, vbResBitmap)
                Case Else
                    blnSkip = True
            End Select
            

            If blnSkip = False Then
                Set mcbMenuCommandBar = gVBE.CommandBars("VBgamer").Controls.Add(1)
                mcbMenuCommandBar.Caption = c.Name
                mcbMenuCommandBar.BeginGroup = True
                mcbMenuCommandBar.Style = msoButtonCaption Or msoButtonIcon
                mcbMenuCommandBar.OnAction = p.Name & "." & c.Name
                mcbMenuCommandBar.ToolTipText = p.Name & "." & c.Name
                'Set MnuEvt = New CommandBarEvents
                CopyBitmapAsButtonFace oPic, &HFF
                ' Paste the icon on the button.
                mcbMenuCommandBar.PasteFace
                Set mcls = New MyCommandEventHandler
                Set mcls.MenuHandler = gVBE.Events.CommandBarEvents(mcbMenuCommandBar)
                EvtHandlers.Add mcls
            End If

End Sub


Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)

On Error Resume Next
gVBE.CommandBars("VBgamer").Delete
Set EvtHandlers = Nothing

Set gVBE = Nothing  'must be last line

'MsgBox "Add-in is now disconnected"
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'' Comment to prevent procedure from being
'' deleted on compilation.
End Sub

Private Sub MenuHandlerLeft_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    'shift visible buttons left
    Dim ocbcTemp As Office.CommandBarControl
    Dim s As String
    Dim p As VBProject
    Dim buttonCount As Long
    Dim ndx As Long
        
    gButtonVisibleIndex = gButtonVisibleIndex - 1
    buttonCount = gVBE.CommandBars("VBgamer").Controls.Count '- gCountControlButtons
    If gButtonVisibleIndex > buttonCount Then gButtonVisibleIndex = gCountControlButtons
    If gButtonVisibleIndex < gCountControlButtons Then gButtonVisibleIndex = gCountControlButtons
        
        
    For ndx = gCountControlButtons + 1 To gButtonVisibleIndex
        gVBE.CommandBars("VBgamer").Controls(ndx).Visible = False
    Next
    For ndx = gButtonVisibleIndex + 1 To buttonCount
        gVBE.CommandBars("VBgamer").Controls(ndx).Visible = True
    Next

End Sub

Private Sub MenuHandlerRight_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    'shift visible buttons right
    'shift visible buttons left
    Dim ocbcTemp As Office.CommandBarControl
    Dim s As String
    Dim p As VBProject
    Dim buttonCount As Long
    Dim ndx As Long
    
    If gButtonVisibleIndex < gCountControlButtons Then gButtonVisibleIndex = gCountControlButtons
    gButtonVisibleIndex = gButtonVisibleIndex + 1
    buttonCount = gVBE.CommandBars("VBgamer").Controls.Count '- gCountControlButtons
    If gButtonVisibleIndex >= buttonCount Then gButtonVisibleIndex = buttonCount - 1
        

    For ndx = gCountControlButtons + 1 To gButtonVisibleIndex
        gVBE.CommandBars("VBgamer").Controls(ndx).Visible = False
    Next
    For ndx = gButtonVisibleIndex + 1 To buttonCount
        gVBE.CommandBars("VBgamer").Controls(ndx).Visible = True
    Next
End Sub
'
'Private Sub PrjHandler_ItemActivated(ByVal VBProject As VBIDE.VBProject)
'    Debug.Print "PrjHandler", "Item Activated"
'End Sub
'Private Sub PrjHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
'    Debug.Print "PrjHandler", "Item Added"
'End Sub
'
'Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
'    Debug.Print "PrjHandler", "Item Removed"
'End Sub
'
'Private Sub PrjHandler_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
'    Debug.Print "PrjHandler", "Item Renamed"
'End Sub
