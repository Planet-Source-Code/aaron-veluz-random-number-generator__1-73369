VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Number Generator"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4695
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   4695
      Begin VB.TextBox txtUB 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtLB 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblUB 
         BackStyle       =   0  'Transparent
         Caption         =   "Upper Bound:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblLB 
         BackStyle       =   0  'Transparent
         Caption         =   "Lower Bound:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label lblOutputx 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim LB, UB As Long
If txtLB.Text = "" Then
    MsgBox "Enter value for Lower Bound", vbOKOnly + vbInformation, "Missing Value"
    txtLB.SetFocus
ElseIf txtUB.Text = "" Then
    MsgBox "Enter value for Upper Bound", vbOKOnly + vbInformation, "Missing Value"
    txtUB.SetFocus
Else
    LB = Val(txtLB.Text)
    UB = Val(txtUB.Text)
    Randomize
    'lblOutputx.Caption = Int(Rnd * (UB + LB)) '+ LB
    lblOutputx.Caption = Int(Rnd * (UB + 1 - LB) + LB)
End If
End Sub

Private Sub Form_Load()
lblOutputx.Caption = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim msg As String
msg = MsgBox("Are you sure you want to exit the application?", vbYesNo + vbInformation, "Exit Application")
If msg = vbNo Then
    Cancel = 1
Else
    End
End If

End Sub

Private Sub txtLB_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Beep

End Sub

Private Sub txtUB_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0: Beep

End Sub
