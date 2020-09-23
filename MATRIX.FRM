VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4344
   ClientLeft      =   2232
   ClientTop       =   1584
   ClientWidth     =   5796
   Height          =   4764
   Left            =   2184
   LinkTopic       =   "Form1"
   ScaleHeight     =   4344
   ScaleWidth      =   5796
   Top             =   1212
   Width           =   5892
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   444
      Left            =   4260
      TabIndex        =   0
      Top             =   372
      Width           =   552
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim M As Object
Set M = CreateObject("Maths.Matrix")
Dim A(1 To 2, 1 To 2)
A(1, 1) = 1                      '  1       3
A(1, 2) = 3                      '   2      10
A(2, 1) = 2
A(2, 2) = 10
Print M.MatrixDet(A)

End Sub

Private Sub Form_Load()

End Sub


