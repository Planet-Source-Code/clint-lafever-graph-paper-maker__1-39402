VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Grid Paper Maker"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2.219
   ScaleMode       =   5  'Inch
   ScaleWidth      =   3.25
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSOLID 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solid"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.HScrollBar scrHORIZONTAL 
      Height          =   255
      LargeChange     =   5
      Left            =   0
      Max             =   12
      Min             =   1
      TabIndex        =   0
      Top             =   0
      Value           =   1
      Width           =   1215
   End
   Begin VB.Label lblCOLOR 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSOLID_Click()
    If Me.chkSOLID.Value = 1 Then
        Me.DrawStyle = 0
    Else
        Me.DrawStyle = 2
    End If
    DrawGrid
End Sub

Private Sub Form_DblClick()
    DrawGrid
    Dim x As Double, y As Double
    Printer.DrawStyle = Me.DrawStyle
    Printer.ScaleMode = vbInches
    While x < Printer.ScaleWidth
        Printer.Line (x, 0)-(x, Printer.ScaleHeight), Me.lblCOLOR.BackColor
        x = x + (0.25 * Me.scrHORIZONTAL.Value)
    Wend
    While y < Printer.ScaleHeight
        Printer.Line (0, y)-(Printer.ScaleWidth, y), Me.lblCOLOR.BackColor
        y = y + (0.25 * Me.scrHORIZONTAL.Value)
    Wend
    Printer.DrawStyle = 0
    Printer.EndDoc
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever - [lafeverc@saic.com]
' Date: September,30 2002 @ 15:42:51
'------------------------------------------------------------
Private Sub DrawGrid()
    On Error GoTo ErrorDrawGrid
    Dim x As Double, y As Double
    y = 0
    x = 0
    Me.ScaleMode = vbInches
    Me.Cls
    While x < Me.ScaleWidth
        Me.Line (x, 0)-(x, Me.ScaleHeight), Me.lblCOLOR.BackColor
        x = x + (0.25 * Me.scrHORIZONTAL.Value)
    Wend
    While y < Me.ScaleHeight
        Me.Line (0, y)-(Me.ScaleWidth, y), Me.lblCOLOR.BackColor
        y = y + (0.25 * Me.scrHORIZONTAL.Value)
    Wend
    Me.CurrentX = Me.chkSOLID.Left + Me.chkSOLID.Width + Me.lblCOLOR.Width
    Me.CurrentY = Me.chkSOLID.Top
    Me.ForeColor = vbBlack
    Me.Print 0.25 * Me.scrHORIZONTAL.Value & """"
    Me.CurrentX = Me.lblCOLOR.Left
    Me.CurrentY = Me.CurrentY + Me.lblCOLOR.Height
    Me.Print "Double click grid to print."
    Exit Sub
ErrorDrawGrid:
    MsgBox Err & ":Error in call to DrawGrid()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub

Private Sub Form_Load()
    Me.BackColor = vbWhite
    Me.ForeColor = vbCyan
End Sub

Private Sub Form_Resize()
Me.scrHORIZONTAL.Move Me.scrHORIZONTAL.Left, 0, Me.ScaleWidth - Me.scrHORIZONTAL.Left
DrawGrid
End Sub

Private Sub lblCOLOR_Click()
    Dim obj As CDLG
    Dim c As Long
    c = Me.lblCOLOR.BackColor
    Set obj = New CDLG
    obj.VBChooseColor c, , True, , Me.hwnd
    Me.lblCOLOR.BackColor = c
    DrawGrid
End Sub

Private Sub scrHORIZONTAL_Change()
    DrawGrid
End Sub

Private Sub scrHORIZONTAL_Scroll()
    DrawGrid
End Sub
