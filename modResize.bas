Attribute VB_Name = "modResize"
Option Explicit

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Public Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type
Public swFormControlPositions() As ControlPositionType
Public gblFormControlPositions() As ControlPositionType
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : Stretch
' Author    : Ellis Dee
' Date      : 29/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Stretch(pic As PictureBox, pstrFile As String)
   On Error GoTo Stretch_Error

    Screen.MousePointer = vbHourglass
    'LockWindowUpdate pic.Parent.hwnd
    If Dir(pstrFile) <> "" Then
        pic.Picture = LoadPicture(pstrFile)
        With pic
            .AutoRedraw = True
            .PaintPicture .Picture, 0, 0, .ScaleWidth, .ScaleHeight
            Set .Picture = .Image
            .Refresh
            .AutoRedraw = False
        End With
    Else
        pic.Picture = LoadPicture
    End If
    'LockWindowUpdate 0
    Screen.MousePointer = vbNormal

   On Error GoTo 0
   Exit Sub

Stretch_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Stretch of Module modResize"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ResizeControls
' Author    : adapted from Rod Stephens @ vb-helper.com
' Date      : 16/04/2021
' Purpose   : Arrange the controls for a new size.
'---------------------------------------------------------------------------------------
'
Public Sub resizeControls(ByRef thisForm As Form, ByRef m_ControlPositions() As ControlPositionType, ByVal m_FormWid As Double, ByVal m_FormHgt As Double, ByVal formFontSize As Single)
    Dim i As Integer: i = 0
    Dim Ctrl As Control
    Dim x_scale As Single: x_scale = 0
    Dim y_scale As Single: y_scale = 0
    Dim fileToLoad As String: fileToLoad = vbNullString
    
    On Error GoTo ResizeControls_Error

    ' Get the form's current scale factors.
    x_scale = thisForm.ScaleWidth / m_FormWid
    y_scale = thisForm.ScaleHeight / m_FormHgt
    
    gblResizeRatio = x_scale

    ' Position the controls.
    i = 1

    For Each Ctrl In thisForm.Controls
        With m_ControlPositions(i)
            If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is VScrollBar) Or (TypeOf Ctrl Is HScrollBar) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is Image) Or (TypeOf Ctrl Is PictureBox) Or (TypeOf Ctrl Is Slider) Then

                If (TypeOf Ctrl Is Image) Then

                    Ctrl.Stretch = True
                    Ctrl.Left = x_scale * .Left
                    Ctrl.Top = y_scale * .Top
                    Ctrl.Height = y_scale * .Height
                    Ctrl.Width = x_scale * .Width

                    Ctrl.Refresh
                ElseIf (TypeOf Ctrl Is PictureBox) Then

                    Ctrl.Left = x_scale * .Left
                    Ctrl.Top = y_scale * .Top
                    Ctrl.Height = y_scale * .Height
                    Ctrl.Width = x_scale * .Width
                    
                    If Ctrl.Name = "picIcon" Then
                        If Ctrl.Index = 0 Then fileToLoad = App.Path & "\general.gif"
                        If Ctrl.Index = 1 Then fileToLoad = App.Path & "\icons.gif"
                        If Ctrl.Index = 2 Then fileToLoad = App.Path & "\behaviour.gif"
                        If Ctrl.Index = 3 Then fileToLoad = App.Path & "\style.gif"
                        If Ctrl.Index = 4 Then fileToLoad = App.Path & "\position.gif"
                        If Ctrl.Index = 5 Then fileToLoad = App.Path & "\about.gif"
                            
                        Ctrl.PaintPicture LoadPicture(fileToLoad), 0, 0, Ctrl.Width, Ctrl.Height
                        Stretch Ctrl, fileToLoad
                    End If
                                  
                    If Ctrl.Name = "picIconPressed" Then
                        If Ctrl.Index = 0 Then fileToLoad = App.Path & "\generalHighlighted.gif"
                        If Ctrl.Index = 1 Then fileToLoad = App.Path & "\iconsHighlighted.gif"
                        If Ctrl.Index = 2 Then fileToLoad = App.Path & "\behaviourHighlighted.gif"
                        If Ctrl.Index = 3 Then fileToLoad = App.Path & "\styleHighlighted.gif"
                        If Ctrl.Index = 4 Then fileToLoad = App.Path & "\positionHighlighted.gif"
                        If Ctrl.Index = 5 Then fileToLoad = App.Path & "\aboutHighlighted.gif"
                            
                        Ctrl.PaintPicture LoadPicture(fileToLoad), 0, 0, Ctrl.Width, Ctrl.Height
                        Stretch Ctrl, fileToLoad
                    End If
                                
                Else
                    Ctrl.Left = x_scale * .Left
                    Ctrl.Top = y_scale * .Top
                    Ctrl.Width = x_scale * .Width
                    If Not (TypeOf Ctrl Is ComboBox) Then
                        ' Cannot change height of ComboBoxes.
                        Ctrl.Height = y_scale * .Height
                    End If
                    On Error Resume Next
                    
                    Ctrl.Font.Size = y_scale * formFontSize
                    
                    ' when resized, a combobox automatically highlights in blue, this removes that
                    If TypeOf Ctrl Is ComboBox Then
                        Ctrl.SelLength = 0
                    End If
                    
                    If Ctrl.Name = "sliGenRunAppInterval" Then
                        Ctrl.Visible = False
                        Ctrl.Visible = True
                        Ctrl.Refresh
                    End If
                
                    Ctrl.Refresh
                    On Error GoTo 0
                End If
            End If
        End With
        i = i + 1
    Next Ctrl
    
'   Dim W: W = thisForm.ScaleX(thisForm.ScaleWidth, thisForm.ScaleMode, vbTwips)
'   Dim H: H = thisForm.ScaleY(thisForm.ScaleHeight, thisForm.ScaleMode, vbTwips)
''
   On Error GoTo 0
   Exit Sub

ResizeControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ResizeControls of Form modResize"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : saveControlSizes
' Author    : Rod Stephens vb-helper.com
' Date      : 16/04/2021
' Purpose   : Resize controls to fit when a form resizes
'             Save the form's and controls' dimensions.
' Credit    : Rod Stephens vb-helper.com
'---------------------------------------------------------------------------------------
'
Public Sub saveControlSizes(ByVal thisForm As Form, ByRef m_ControlPositions() As ControlPositionType, ByRef m_FormWid As Long, ByRef m_FormHgt As Long)
    Dim i As Integer: i = 0
    Dim Ctrl As Control

    ' Save the controls' positions and sizes.
    On Error GoTo saveControlSizes_Error

    ReDim m_ControlPositions(1 To thisForm.Controls.Count)
    i = 1
    For Each Ctrl In thisForm.Controls
        With m_ControlPositions(i)
        
            If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is VScrollBar) Or (TypeOf Ctrl Is HScrollBar) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is Image) Or (TypeOf Ctrl Is PictureBox) Or (TypeOf Ctrl Is Slider) Then
                .Left = Ctrl.Left
                .Top = Ctrl.Top
                .Width = Ctrl.Width
                .Height = Ctrl.Height
                On Error Resume Next ' cater for any controls that do not have a font property that may cause an error
                .FontSize = Ctrl.Font.Size
                On Error GoTo 0
            End If
        End With
        i = i + 1
    Next Ctrl

    ' Save the form's size.
    m_FormWid = thisForm.ScaleWidth
    m_FormHgt = thisForm.ScaleHeight

   On Error GoTo 0
   Exit Sub

saveControlSizes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveControlSizes of Form modResize"
End Sub
