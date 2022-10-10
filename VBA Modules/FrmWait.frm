VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmWait 
   Caption         =   "Please wait..."
   ClientHeight    =   420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4065
   OleObjectBlob   =   "FrmWait.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private calcSetting As Long

Private Declare Function GetWindowLong _
                       Lib "user32" Alias "GetWindowLongA" ( _
                       ByVal hWnd As Long, _
                       ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong _
                       Lib "user32" Alias "SetWindowLongA" ( _
                       ByVal hWnd As Long, _
                       ByVal nIndex As Long, _
                       ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar _
                       Lib "user32" ( _
                       ByVal hWnd As Long) As Long
Private Declare Function FindWindowA _
                       Lib "user32" (ByVal lpClassName As String, _
                       ByVal lpWindowName As String) As Long


Private Sub UserForm_Initialize()
  HideTitleBar Me
  
  ' ############
  'PURPOSE: Position userform to center of Excel Window (important for dual monitor compatibility)
  'SOURCE: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
  
  'Start Userform Centered inside Excel Screen (for dual monitors)
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
  Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
  ' ############

End Sub

Sub SetText(s As String)
  ' Disabling automatic calculation for performance reasons
  Application.Calculation = xlCalculationManual
  
  ' Disabling Screen updates for performance reasons
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False

  If s <> Me.lblTextSplash.Caption Then
    ' Prevent Excel from freezing up during Macro execution
    ' (Slight performance impact, but feels more responsive)
    DoEvents
  End If

  Me.Show
  If s <> "" Then
    Me.lblTextSplash.Caption = s
    Me.Repaint
  Else
    Me.Remove
  End If
End Sub

Sub Remove()
  Application.Calculation = xlCalculationAutomatic
  
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  
  Unload Me
End Sub


Private Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub
