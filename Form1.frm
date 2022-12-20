VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "OK"
      Height          =   795
      Left            =   3060
      TabIndex        =   2
      Top             =   5550
      Width           =   1035
   End
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2310
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4560
      Width           =   2505
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2310
      TabIndex        =   0
      Top             =   3420
      Width           =   2505
   End
   Begin VB.Image imgLogin 
      Height          =   2340
      Left            =   1740
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   420
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cursor As ADODB.Recordset
Dim dono  As String
Dim pets, raca As String

Private Sub cmdLogin_Click()

'If UCase(Trim(txtUsuario.Text)) = "ADMIN" And UCase(Trim(txtSenha.Text)) = "123" Then
    Form2.Show
    Me.Hide
'ElseIf (UCase(Trim(txtUsuario.Text)) = "ADMIN") Then
'    MsgBox "Senha incorreta! ", vbExclamation, " Erro de Login!"
'Else
'    MsgBox "Usuário não existe! ", vbExclamation, " Erro de Login!"
'End If


End Sub
