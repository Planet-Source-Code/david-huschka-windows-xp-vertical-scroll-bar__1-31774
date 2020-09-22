VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin Project1.XPVScroll XPVScroll1 
      Align           =   4  'Align Right
      Height          =   3735
      Left            =   4770
      TabIndex        =   2
      ToolTipText     =   "Hello"
      Top             =   0
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   6588
      LargeChange     =   50
      Max             =   100
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3750
      LargeChange     =   50
      Left            =   1725
      Max             =   100
      TabIndex        =   0
      Top             =   15
      Width           =   240
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3885
      TabIndex        =   3
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   570
      TabIndex        =   1
      Top             =   1275
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

    VScroll1.Height = Me.ScaleHeight
End Sub

Private Sub VScroll1_Change()
    Label1 = VScroll1.Value
    XPVScroll1.Value = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    Label1 = VScroll1.Value
    XPVScroll1.Value = VScroll1.Value
End Sub

Private Sub XPVScroll1_Scroll()
    Label2 = XPVScroll1.Value
    VScroll1.Value = XPVScroll1.Value
End Sub

Private Sub XPVScroll1_Change()
    Label2 = XPVScroll1.Value
    VScroll1.Value = XPVScroll1.Value
End Sub
