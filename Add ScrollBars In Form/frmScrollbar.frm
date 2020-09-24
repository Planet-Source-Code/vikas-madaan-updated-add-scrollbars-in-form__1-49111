VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScrollbar 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HS 
      Height          =   255
      LargeChange     =   525
      Left            =   0
      SmallChange     =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3960
      Width           =   6255
   End
   Begin VB.VScrollBar VS 
      Height          =   3855
      LargeChange     =   525
      Left            =   6120
      SmallChange     =   225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4035
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   3975
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   3015
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   7525
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   3375
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5880
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   3840
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'PROGRAM    :   Add ScrollBars In Form
'AUTHOR     :   Vikas Madaan
'  __         __        ___      ___
'  \ \       / /       |   \    /   |
'   \ \     / /        | |\ \  / /| |
'    \ \   / /         | | \ \/ / | |
'     \ \_/ /    __    | |  \__/  | |
'      \___/    (__)   |_|        |_|
'
'DATE       :   October 09, 2003.
'
'COMMENTS   :   This Code is to show U how to Add ScrollBars in
'           the Form So the Controls in the Form Moves Up & Down
'           or Left & Right According to the Scrollbars.
'           It is very useful when UR Controls Exceed the
'           Width or Height of the Form.
'           No API Functions are Used Simple ScrollBar Controls.
'           If you need support or to give suggestions to improve,
'           you can email me at vikasmadaan25@hotmail.com
'           or thru yahoo messenger vikasmadaan25@yahoo.com
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

'IMPORTANT
'Note : After inserting all the Controls to the Form
'Select Both Scrollbars and StatusBar using Mouse and Ctrl
'Button Now Right Click the Mouse Button &
'Select "Bring to Front" Option so they Are on Top of
'all other Controls

Option Explicit
Dim VSLVal As Integer 'Last Value of Vertical Scrollbar
Dim HSLVal As Integer 'Last Value of Horizontal Scrollbar
Dim HMax As Integer 'Max Value for Horizontal Scrollbar
Dim VMax As Integer 'Max Value for Vertical Scrollbar

Private Sub Form_Load()
'Set the Max Values That the Scrollbar can have
'This value must be greater than the value of Last Control
'The following is used to set manualy
'HMax = 9525
'VMax = 7525
'or
'the following is used to find automaticaly
On Error Resume Next
Dim Cntrl As Control

For Each Cntrl In Me.Controls
 If (Not TypeOf Cntrl Is HScrollBar) And _
    (Not TypeOf Cntrl Is VScrollBar) And _
    (Not TypeOf Cntrl Is StatusBar) Then
  If (Cntrl.Top + Cntrl.Height + 125) > VMax Then
   VMax = Cntrl.Top + Cntrl.Height + 125
  End If
  If (Cntrl.Left + Cntrl.Width + 125) > HMax Then
   HMax = Cntrl.Left + Cntrl.Width + 125
  End If
 End If
Next
On Error GoTo 0

'Set the Default Values for ScrollBar
With HS
.Left = 0
.Top = Me.ScaleHeight - .Height
SBar.Height = .Height
.Width = Me.ScaleWidth - VS.Width
.Max = HMax - .Width
If .Max < 0 Then .Max = 0
End With

With VS
.Height = Me.ScaleHeight - HS.Height
.Left = Me.ScaleWidth - .Width
.Top = 0
.Max = VMax - .Height
If .Max < 0 Then .Max = 0
End With

'If Max Value <= 0 then Disable ScrollBar
If HS.Max <= 0 Then HS.Enabled = False
If VS.Max <= 0 Then VS.Enabled = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
 Exit Sub 'Exit if Minimized
ElseIf Me.Height < 2525 Then
 Me.Height = 2525 'Default Height
ElseIf Me.Width < 2525 Then
 Me.Width = 2525 'Default Width
End If

'Set the New Values for ScrollBar
With HS
.Left = 0
.Top = Me.ScaleHeight - .Height
.Width = Me.ScaleWidth - VS.Width
.Max = HMax - .Width
End With

With VS
.Height = Me.ScaleHeight - HS.Height
.Left = Me.ScaleWidth - .Width
.Top = 0
.Max = VMax - .Height
End With

'If Max Value <= 0 then Disable ScrollBar
If HS.Max <= 0 Then
 HS.Enabled = False
 HS.Max = 0
Else
 HS.Enabled = True
End If
If VS.Max <= 0 Then
 VS.Enabled = False
 VS.Max = 0
Else
 VS.Enabled = True
End If
End Sub

'Change Values if the Value of Horizontal ScrollBar Changed
Private Sub HSPositionChanged()
On Error Resume Next
Dim Val As Integer, Cntrl As Control
Val = HSLVal - HS.Value
For Each Cntrl In Me.Controls
 If TypeOf Cntrl Is HScrollBar Or _
    TypeOf Cntrl Is VScrollBar Or _
    TypeOf Cntrl Is Menu Or _
    Cntrl.Container.Name <> Me.Name Then
 'Do Nothing
 ElseIf TypeOf Cntrl Is Line Then
  Cntrl.X1 = Cntrl.X1 + Val
  Cntrl.X2 = Cntrl.X2 + Val
 Else
  Cntrl.Left = Cntrl.Left + Val
 End If
Next
HSLVal = HS.Value
End Sub

'Change Values if the Value of Vertical ScrollBar Changed
Private Sub VSPositionChanged()
On Error Resume Next
Dim Val As Integer, Cntrl As Control
Val = VSLVal - VS.Value
For Each Cntrl In Me.Controls
 If TypeOf Cntrl Is HScrollBar Or _
    TypeOf Cntrl Is VScrollBar Or _
    TypeOf Cntrl Is Menu Or _
    Cntrl.Container.Name <> Me.Name Then
 'Do Nothing
 ElseIf TypeOf Cntrl Is Line Then
  Cntrl.Y1 = Cntrl.Y1 + Val
  Cntrl.Y2 = Cntrl.Y2 + Val
 Else
  Cntrl.Top = Cntrl.Top + Val
 End If
Next
VSLVal = VS.Value
End Sub

Private Sub HS_Change()
HSPositionChanged
End Sub

Private Sub HS_Scroll()
HSPositionChanged
End Sub

Private Sub vs_Change()
VSPositionChanged
End Sub

Private Sub vs_Scroll()
VSPositionChanged
End Sub

