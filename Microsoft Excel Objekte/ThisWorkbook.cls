VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    If ActiveWorkbook Is Nothing Then CreateCommandBar
End Sub

Sub CreateCommandBar()
    CreateMenuItem "Convert table to LaTeX", "Conversion.LaTeX"
    CreateMenuItem "Convert all stored tables to LaTeX", "Conversion.LaTeXAllToFiles"
End Sub

Private Sub CreateMenuItem(ByVal Caption As String, ByVal Action As String)
    Dim ctl As CommandBarControl
    Dim i As Long
    Dim ControlCollection As New Collection
    Dim myMenubar As CommandBar, toolsMenu As CommandBarPopup, newMenuItem As CommandBarControl, newButton As CommandBarControl
    
    'First Create Menu Item
    Set myMenubar = Application.CommandBars.ActiveMenuBar
    Set toolsMenu = myMenubar.Controls(6)
    Set newMenuItem = myMenubar.FindControl(Tag:=Action, recursive:=True)
    If Not newMenuItem Is Nothing Then newMenuItem.Delete
    Set newMenuItem = toolsMenu.Controls.Add(Type:=msoControlButton, Before:=3)
    newMenuItem.Tag = Action
    
    ' Versions before Office 2007 only (=> ribbons!)
    If CLng(Split(Application.Version, ".")(0)) < 12 Then
        'Now create tool bar
        Dim myToolBar As CommandBar
        On Error Resume Next
        Set myToolBar = Application.CommandBars(Action)
        On Error GoTo 0
        If myToolBar Is Nothing Then
            Set myToolBar = Application.CommandBars.Add(Name:=Action)
        End If
        If myToolBar.Controls.Count > 0 Then myToolBar.Controls(1).Delete
        
        myToolBar.Position = msoBarTop
        myToolBar.Visible = True
        Set newButton = myToolBar.Controls.Add(msoControlButton)
    End If
      
    If Not newButton Is Nothing Then
        ControlCollection.Add newButton
    End If
    If Not newMenuItem Is Nothing Then
        ControlCollection.Add newMenuItem
    End If
    
    For Each ctl In ControlCollection
        ctl.OnAction = Action
        ctl.FaceId = 107
        ctl.TooltipText = Caption
        ctl.Caption = Caption
    Next
End Sub
