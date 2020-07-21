VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pros As Integer
Dim porX As Integer

Private Sub Command1_Click()

If Me.ProgressBar1.Value = 100 Then
   pros = -1
   If PORS = -1 Then
   pros = 0
   End If
   Me.ProgressBar1.Value = pros
   Else
   Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
End If


pros = Me.ProgressBar1.Value
porX = pros
Me.Label1.Caption = "&" & pros & "%"

End Sub

Private Sub Command2_Click()
Me.ProgressBar1.Value = porX
Me.Label1.Caption = "&" & porX & "%"
End Sub

Private Sub Command3_Click()
Me.ProgressBar1.Value = 0
pros = Me.ProgressBar1.Value
Me.Label1.Caption = "&" & pros & "%"
End Sub

Private Sub Form_Load()
Me.Caption = "ProgresBar*"
Me.Command1.Caption = "&Incrementar"
Me.Command2.Caption = "&Obtener"
Me.Command3.Caption = "&Borrar"
End Sub
