VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form TelaPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cli�nica Veterin�ria Tot�"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11925
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Atualizar"
      Height          =   795
      Left            =   7540
      TabIndex        =   55
      Top             =   6870
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   795
      Left            =   2565
      TabIndex        =   54
      Top             =   6870
      Width           =   1215
   End
   Begin VB.CommandButton cmdTeste 
      Caption         =   "Teste"
      Height          =   465
      Left            =   9030
      TabIndex        =   52
      Top             =   7020
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   795
      Left            =   3810
      TabIndex        =   51
      Top             =   6870
      Width           =   1215
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   795
      Left            =   5055
      TabIndex        =   50
      Top             =   6870
      Width           =   1215
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "Cadastrar"
      Height          =   795
      Left            =   6305
      TabIndex        =   49
      Top             =   6870
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6465
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEndereco"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCpf"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblNomeCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCliente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNomeCliente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtEndereco"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCpf"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdNovoPet"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Pet"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtHistorico1"
      Tab(1).Control(1)=   "txtNascimento"
      Tab(1).Control(2)=   "txtSexo"
      Tab(1).Control(3)=   "txtRaca"
      Tab(1).Control(4)=   "txtNome"
      Tab(1).Control(5)=   "txtNas"
      Tab(1).Control(6)=   "Label3"
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(8)=   "Label1"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Pet 2"
      TabPicture(2)   =   "Form2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtHistorico2"
      Tab(2).Control(1)=   "txtNome2"
      Tab(2).Control(2)=   "txtRaca2"
      Tab(2).Control(3)=   "txtSexo2"
      Tab(2).Control(4)=   "txtNascimento2"
      Tab(2).Control(5)=   "Label7"
      Tab(2).Control(6)=   "Label6"
      Tab(2).Control(7)=   "Label5"
      Tab(2).Control(8)=   "Label4"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Pet 3"
      TabPicture(3)   =   "Form2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtHistorico3"
      Tab(3).Control(1)=   "txtNome3"
      Tab(3).Control(2)=   "txtRaca3"
      Tab(3).Control(3)=   "txtSexo3"
      Tab(3).Control(4)=   "txtNascimento3"
      Tab(3).Control(5)=   "Label11"
      Tab(3).Control(6)=   "Label10"
      Tab(3).Control(7)=   "Label9"
      Tab(3).Control(8)=   "Label8"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Pet 4"
      TabPicture(4)   =   "Form2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtHistorico4"
      Tab(4).Control(1)=   "txtNome4"
      Tab(4).Control(2)=   "txtRaca4"
      Tab(4).Control(3)=   "txtSexo4"
      Tab(4).Control(4)=   "txtNascimento4"
      Tab(4).Control(5)=   "Label15"
      Tab(4).Control(6)=   "Label14"
      Tab(4).Control(7)=   "Label13"
      Tab(4).Control(8)=   "Label12"
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "Pet 5"
      TabPicture(5)   =   "Form2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtHistorico5"
      Tab(5).Control(1)=   "txtNome5"
      Tab(5).Control(2)=   "txtRaca5"
      Tab(5).Control(3)=   "txtSexo5"
      Tab(5).Control(4)=   "txtNascimento5"
      Tab(5).Control(5)=   "Label19"
      Tab(5).Control(6)=   "Label18"
      Tab(5).Control(7)=   "Label17"
      Tab(5).Control(8)=   "Label16"
      Tab(5).ControlCount=   9
      Begin VB.TextBox txtHistorico5 
         Height          =   3375
         Left            =   -67290
         TabIndex        =   60
         Top             =   1530
         Width           =   3525
      End
      Begin VB.TextBox txtHistorico4 
         Height          =   3375
         Left            =   -67290
         TabIndex        =   59
         Top             =   1530
         Width           =   3525
      End
      Begin VB.TextBox txtHistorico3 
         Height          =   3375
         Left            =   -67290
         TabIndex        =   58
         Top             =   1530
         Width           =   3525
      End
      Begin VB.TextBox txtHistorico2 
         Height          =   3375
         Left            =   -67290
         TabIndex        =   57
         Top             =   1530
         Width           =   3525
      End
      Begin VB.TextBox txtHistorico1 
         Height          =   3375
         Left            =   -67290
         TabIndex        =   56
         Top             =   1530
         Width           =   3525
      End
      Begin VB.CommandButton cmdNovoPet 
         Caption         =   "Novo Pet"
         Height          =   525
         Left            =   4980
         TabIndex        =   53
         Top             =   5220
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSMask.MaskEdBox txtNascimento 
         Height          =   375
         Left            =   -71010
         TabIndex        =   48
         Top             =   4000
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNome2 
         Height          =   375
         Left            =   -71010
         TabIndex        =   43
         Top             =   1600
         Width           =   3075
      End
      Begin VB.TextBox txtRaca2 
         Height          =   375
         Left            =   -71010
         TabIndex        =   42
         Top             =   2400
         Width           =   3075
      End
      Begin VB.TextBox txtSexo2 
         Height          =   375
         Left            =   -71010
         TabIndex        =   41
         Top             =   3200
         Width           =   3075
      End
      Begin VB.TextBox txtNascimento2 
         Height          =   375
         Left            =   -71010
         TabIndex        =   40
         Top             =   4000
         Width           =   1400
      End
      Begin VB.TextBox txtNome5 
         Height          =   375
         Left            =   -71010
         TabIndex        =   35
         Top             =   1600
         Width           =   3075
      End
      Begin VB.TextBox txtRaca5 
         Height          =   375
         Left            =   -71010
         TabIndex        =   34
         Top             =   2400
         Width           =   3075
      End
      Begin VB.TextBox txtSexo5 
         Height          =   375
         Left            =   -71010
         TabIndex        =   33
         Top             =   3200
         Width           =   3075
      End
      Begin VB.TextBox txtNascimento5 
         Height          =   375
         Left            =   -71010
         TabIndex        =   32
         Top             =   4000
         Width           =   1400
      End
      Begin VB.TextBox txtNome4 
         Height          =   375
         Left            =   -71010
         TabIndex        =   27
         Top             =   1600
         Width           =   3075
      End
      Begin VB.TextBox txtRaca4 
         Height          =   375
         Left            =   -71010
         TabIndex        =   26
         Top             =   2400
         Width           =   3075
      End
      Begin VB.TextBox txtSexo4 
         Height          =   375
         Left            =   -71010
         TabIndex        =   25
         Top             =   3200
         Width           =   3075
      End
      Begin VB.TextBox txtNascimento4 
         Height          =   375
         Left            =   -71010
         TabIndex        =   24
         Top             =   4000
         Width           =   1400
      End
      Begin VB.TextBox txtNome3 
         Height          =   375
         Left            =   -71010
         TabIndex        =   19
         Top             =   1600
         Width           =   3075
      End
      Begin VB.TextBox txtRaca3 
         Height          =   375
         Left            =   -71010
         TabIndex        =   18
         Top             =   2400
         Width           =   3075
      End
      Begin VB.TextBox txtSexo3 
         Height          =   375
         Left            =   -71010
         TabIndex        =   17
         Top             =   3200
         Width           =   3075
      End
      Begin VB.TextBox txtNascimento3 
         Height          =   375
         Left            =   -71010
         TabIndex        =   16
         Top             =   4000
         Width           =   1400
      End
      Begin VB.TextBox txtSexo 
         Height          =   375
         Left            =   -71010
         TabIndex        =   11
         Top             =   3200
         Width           =   3075
      End
      Begin VB.TextBox txtRaca 
         Height          =   375
         Left            =   -71010
         TabIndex        =   10
         Top             =   2400
         Width           =   3075
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   -71010
         TabIndex        =   9
         Top             =   1600
         Width           =   3075
      End
      Begin VB.TextBox txtCpf 
         Height          =   345
         Left            =   3990
         MaxLength       =   11
         TabIndex        =   8
         Top             =   2400
         Width           =   1995
      End
      Begin VB.TextBox txtEndereco 
         Height          =   345
         Left            =   3990
         TabIndex        =   6
         Top             =   4000
         Width           =   4145
      End
      Begin VB.TextBox txtNomeCliente 
         Height          =   345
         Left            =   3990
         TabIndex        =   5
         Top             =   3200
         Width           =   4145
      End
      Begin VB.TextBox txtCliente 
         Height          =   345
         Left            =   3990
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1600
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Nome"
         Height          =   255
         Left            =   -72150
         TabIndex        =   47
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Ra�a"
         Height          =   225
         Left            =   -72150
         TabIndex        =   46
         Top             =   2445
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Sexo"
         Height          =   225
         Left            =   -72150
         TabIndex        =   45
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   -72150
         TabIndex        =   44
         Top             =   4050
         Width           =   915
      End
      Begin VB.Label Label19 
         Caption         =   "Nome"
         Height          =   255
         Left            =   -72150
         TabIndex        =   39
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label Label18 
         Caption         =   "Ra�a"
         Height          =   225
         Left            =   -72150
         TabIndex        =   38
         Top             =   2445
         Width           =   915
      End
      Begin VB.Label Label17 
         Caption         =   "Sexo"
         Height          =   225
         Left            =   -72150
         TabIndex        =   37
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   -72150
         TabIndex        =   36
         Top             =   4050
         Width           =   915
      End
      Begin VB.Label Label15 
         Caption         =   "Nome"
         Height          =   255
         Left            =   -72150
         TabIndex        =   31
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label Label14 
         Caption         =   "Ra�a"
         Height          =   225
         Left            =   -72150
         TabIndex        =   30
         Top             =   2445
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "Sexo"
         Height          =   225
         Left            =   -72150
         TabIndex        =   29
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   -72150
         TabIndex        =   28
         Top             =   4050
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "Nome"
         Height          =   255
         Left            =   -72150
         TabIndex        =   23
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Ra�a"
         Height          =   225
         Left            =   -72150
         TabIndex        =   22
         Top             =   2445
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "Sexo"
         Height          =   225
         Left            =   -72150
         TabIndex        =   21
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   -72150
         TabIndex        =   20
         Top             =   4050
         Width           =   915
      End
      Begin VB.Label txtNas 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   -72150
         TabIndex        =   15
         Top             =   4050
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Sexo"
         Height          =   225
         Left            =   -72150
         TabIndex        =   14
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Ra�a"
         Height          =   225
         Left            =   -72150
         TabIndex        =   13
         Top             =   2445
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         Height          =   255
         Left            =   -72150
         TabIndex        =   12
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label lblNomeCliente 
         Caption         =   "Nome"
         Height          =   195
         Left            =   2850
         TabIndex        =   7
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label lblCpf 
         Caption         =   "CPF"
         Height          =   195
         Left            =   2850
         TabIndex        =   4
         Top             =   2445
         Width           =   915
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endere�o"
         Height          =   195
         Left            =   2850
         TabIndex        =   3
         Top             =   4050
         Width           =   915
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   2850
         TabIndex        =   2
         Top             =   1650
         Visible         =   0   'False
         Width           =   915
      End
   End
End
Attribute VB_Name = "TelaPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strQuery As String
Dim conn As Object
Dim cursor As ADODB.Recordset
Dim novoPet As Boolean
Dim clienteId, petId As Integer


Private Sub cmdAtualizar_Click()
    Atualizar
End Sub

Private Sub cmdCadastrar_Click()
    If novoPet Then
    
        strQuery = "SELECT ID FROM Clientes "
        strQuery = strQuery & "WHERE Cpf LIKE '%" & txtCpf.Text & "%'"
        
        Set cursor = New ADODB.Recordset
        
        cursor.Open strQuery, AbreConn
        
        txtCliente.Text = cursor.Fields(0).Value
        
        cursor.MoveNext
        
        cursor.Close
        
        strQuery = "SELECT * FROM pets p "
        strQuery = strQuery & "JOIN petsdonos pd  "
        strQuery = strQuery & "ON pd.pet = p.id  "
        strQuery = strQuery & "JOIN clientes c  "
        strQuery = strQuery & "ON c.id = pd.dono  "
        strQuery = strQuery & "WHERE c.id = " & txtCliente.Text
        
        cursor.Open strQuery, AbreConn
        j = 0
        Do Until cursor.EOF
            j = j + 1
            cursor.MoveNext
        Loop
        
        cursor.Close
                             
        If i < 5 Then
            CadastrarNovoPet txtCliente.Text
        Else
            MsgBox "N�mero m�ximo de usu�rios j� cadastrado"
            Exit Sub
        End If
        
    Else
        Cadastrar
    End If
End Sub

Private Sub cmdConsultar_Click()
     
    Dim pets() As String
    Dim raca() As String
    Dim sexo() As String
    Dim historico() As String
    Dim nascimento() As String
    Dim looping As Boolean
        
    looping = False
    cmdNovoPet.Visible = True
    
    strQuery = "SELECT Id FROM Clientes "
    strQuery = strQuery & " WHERE Nome LIKE '%" & UCase(Trim(txtNomeCliente.Text)) & "%' "
    
    Set cursor = New ADODB.Recordset
    cursor.Open strQuery, AbreConn
    
    If cursor.EOF Then
        cmdLimpar_Click
        MsgBox "Nenhum usu�rio cadastrado com esse nome!", vbExclamation, " Alerta!"
        Exit Sub
    End If
    
    clienteId = cursor.Fields(0).Value
    
    cursor.MoveNext
    
    If Not cursor.EOF Then
        cmdLimpar_Click
        MsgBox "Mais de um usu�rio cadastrado, especifique melhor o nome!", vbExclamation, " Alerta!"
        Exit Sub
    End If
    
    cursor.Close
    
    If Trim(txtNomeCliente.Text) = "" And looping Then
        MsgBox "Digite um c�digo de cliente v�lido!", vbExclamation, " Alerta!"
    Else

        strQuery = "SELECT p.nome, "
        strQuery = strQuery & "p.raca, "
        strQuery = strQuery & "p.sexo, "
        strQuery = strQuery & "p.datanascimento, "
        strQuery = strQuery & "p.historico, "
        strQuery = strQuery & "c.nome AS NomeDono, "
        strQuery = strQuery & "c.endereco, "
        strQuery = strQuery & "c.cpf "
        strQuery = strQuery & "FROM pets p "
        strQuery = strQuery & "JOIN petsdonos pd "
        strQuery = strQuery & "ON pd.pet = p.id "
        strQuery = strQuery & "JOIN clientes c "
        strQuery = strQuery & "ON c.id = pd.dono "
        strQuery = strQuery & "WHERE c.nome LIKE '%" & UCase(Trim(txtNomeCliente.Text)) & "%' "
                
        cursor.Open strQuery, AbreConn
                
        If cursor.EOF Then
            
            cursor.Close
            txtNomeCliente.Text = ""
            cmdLimpar_Click
            
            cmdConsultar.Enabled = True
            cmdCadastrar.Enabled = True
            cmdLimpar.Enabled = False
            
            MsgBox "Cliente n�o existe!", vbExclamation, " Alerta!"
            
            looping = True
            
            ' cursor.Open strQuery, AbreConn

        End If
          
        If Not cursor.EOF Then
            txtNomeCliente.Text = cursor.Fields("NomeDono").Value
            txtEndereco.Text = cursor.Fields("endereco").Value
            txtCpf.Text = cursor.Fields("cpf").Value
            
           ' txtCliente.Enabled = False
            
        End If
        i = 0
        Do Until cursor.EOF
        
            i = i + 1
            
            ReDim Preserve pets(1 To i)
            ReDim Preserve raca(1 To i)
            ReDim Preserve historico(1 To i)
            ReDim Preserve nascimento(1 To i)
            ReDim Preserve sexo(1 To i)
            
            Select Case i
            
                Case 1
                
                    pets(i) = cursor.Fields(0).Value
                    raca(i) = cursor.Fields(1).Value
                    sexo(i) = cursor.Fields(2).Value
                    historico(i) = cursor.Fields("historico").Value
                    
                    nascimento(i) = CStr(cursor.Fields(3).Value)
                    
                    txtNome.Text = pets(i)
                    txtRaca.Text = raca(i)
                    txtSexo.Text = sexo(i)
                    
                    txtHistorico1.Text = historico(i)
                    txtNascimento.Text = nascimento(i)
                    
                 Case 2
                    
                    SSTab1.TabEnabled(2) = True
                    SSTab1.TabVisible(2) = True
                    
                    pets(i) = cursor.Fields(0).Value
                    raca(i) = cursor.Fields(1).Value
                    sexo(i) = cursor.Fields(2).Value
                    historico(i) = cursor.Fields("historico").Value
                    
                    nascimento(i) = CStr(cursor.Fields(3).Value)
                    
                    txtNome2.Text = pets(i)
                    txtRaca2.Text = raca(i)
                    txtSexo2.Text = sexo(i)
                    
                    txtHistorico2.Text = historico(i)
                    txtNascimento2.Text = nascimento(i)
                    
                Case 3
                    
                    SSTab1.TabEnabled(3) = True
                    SSTab1.TabVisible(3) = True
                    
                    pets(i) = cursor.Fields(0).Value
                    raca(i) = cursor.Fields(1).Value
                    sexo(i) = cursor.Fields(2).Value
                    historico(i) = cursor.Fields("historico").Value
                    
                    nascimento(i) = CStr(cursor.Fields(3).Value)
                    
                    txtNome3.Text = pets(i)
                    txtRaca3.Text = raca(i)
                    txtSexo3.Text = sexo(i)
                    
                    txtHistorico3.Text = historico(i)
                    txtNascimento3.Text = nascimento(i)
                    
                Case 4
                    
                    SSTab1.TabEnabled(4) = True
                    SSTab1.TabVisible(4) = True
                    
                    pets(i) = cursor.Fields(0).Value
                    raca(i) = cursor.Fields(1).Value
                    sexo(i) = cursor.Fields(2).Value
                    historico(i) = cursor.Fields("historico").Value
                    
                    nascimento(i) = CStr(cursor.Fields(3).Value)
                    
                    txtNome4.Text = pets(i)
                    txtRaca4.Text = raca(i)
                    txtSexo4.Text = sexo(i)
                    
                    txtHistorico4.Text = historico(i)
                    txtNascimento4.Text = nascimento(i)
                    
                Case 5
                    
                    SSTab1.TabEnabled(5) = True
                    SSTab1.TabVisible(5) = True
                    
                    pets(i) = cursor.Fields(0).Value
                    raca(i) = cursor.Fields(1).Value
                    sexo(i) = cursor.Fields(2).Value
                    historico(i) = cursor.Fields("historico").Value
                    
                    nascimento(i) = CStr(cursor.Fields(3).Value)
                    
                    txtNome5.Text = pets(i)
                    txtRaca5.Text = raca(i)
                    txtSexo5.Text = sexo(i)
                    
                    txtHistorico5.Text = historico(i)
                    txtNascimento5.Text = nascimento(i)
                    
                Case Else
                
                    MsgBox "Fora do case"
                    
            End Select
            
            cursor.MoveNext
            
            
            
        Loop
        
        cursor.Close
        
        If Not looping Then
            cmdConsultar.Enabled = False
            cmdCadastrar.Enabled = False
            cmdLimpar.Enabled = True
            
        End If
    End If
    
    cmdExcluir.Enabled = True
    cmdAtualizar.Enabled = True
    
    strQuery = "SELECT ID FROM Pets "
    strQuery = strQuery & "WHERE Nome = '" & txtNome.Text & "' "
    strQuery = strQuery & "AND Raca = '" & txtRaca.Text & "' "
    
    cursor.Open strQuery, AbreConn

    petId = cursor.Fields(0).Value

    cursor.Close
    
End Sub

'Excluir

Private Sub cmdExcluir_Click()

    'Confirma��o
    Dim res
    res = Confirma("Deseja realmente excluir o cliente?")
    
    If res = 0 Then
        Exit Sub
    End If
    
    ' id clientes
    Dim idCliente As Integer
    Dim idPet() As Integer
    
    strQuery = "SELECT ID FROM Clientes "
    strQuery = strQuery & "WHERE Cpf LIKE '%" & Trim(txtCpf.Text) & "%'"
    
    Set cursor = New ADODB.Recordset
    
    cursor.Open strQuery, AbreConn
    idCliente = cursor.Fields(0).Value
    
    cursor.Close
    
    strQuery = "SELECT pets.id FROM pets "
    strQuery = strQuery & "JOIN petsdonos pd  "
    strQuery = strQuery & "ON pets.id  =  pd.pet "
    strQuery = strQuery & "WHERE  pd.dono = " & idCliente
    
    cursor.Open strQuery, AbreConn
    ' guardar id pets
    ind = 0
    
    Do Until cursor.EOF
    
        ind = ind + 1
        ReDim Preserve idPet(1 To ind)
        idPet(ind) = cursor.Fields(0).Value
        
    cursor.MoveNext
    Loop
    
    cursor.Close
    
    strQuery = "DELETE PetsDonos WHERE Dono = " & idCliente

    Set conn = AbreConn
    conn.Execute strQuery
    conn.Close
    
    For ix = 1 To UBound(idPet)
        
        strQuery = "DELETE Pets WHERE Id = " & idPet(ix)
        Set conn = AbreConn
        conn.Execute strQuery
        conn.Close

        
    Next ix
    
    strQuery = "DELETE Clientes WHERE Id = " & idCliente
    Set conn = AbreConn
    conn.Execute strQuery
    conn.Close


MsgBox "Cliente exclu�do com sucesso!", vbExclamation, " Alerta!"

cmdLimpar_Click

End Sub

Private Sub cmdLimpar_Click()

    txtCliente.Text = ""
    txtCpf.Text = ""
    txtNomeCliente.Text = ""
    txtEndereco.Text = ""
    
    txtNome.Text = ""
    txtRaca.Text = ""
    txtSexo.Text = ""
    txtHistorico1.Text = ""
    txtNascimento.Text = "__/__/____"
    
    
    For i = 2 To 5
    
        SSTab1.Tab = 1

        If InStr(SSTab1.Caption, "Pet") = 1 Then
            SSTab1.TabVisible(i) = False
        End If
        
    Next i
    
    SSTab1.Tab = 0
    cmdConsultar.Enabled = True
    cmdCadastrar.Enabled = True
    cmdLimpar.Enabled = False
    cmdExcluir.Enabled = False
    cmdAtualizar.Enabled = False
    cmdNovoPet.Visible = False
    
End Sub

Private Sub cmdNovoPet_Click()
  
    SSTab1.Tab = 1
    cmdConsultar.Enabled = False
    cmdCadastrar.Enabled = True
    cmdLimpar.Enabled = True
        
    txtNome.Text = ""
    txtRaca.Text = ""
    txtSexo.Text = ""
    txtNascimento.Text = "__/__/____"
    
    novoPet = True
    
End Sub

'Bot�o pra testes
Private Sub cmdTeste_Click()

    Dim res
    res = Confirma("Deseja realmente excluir o cliente?")
    MsgBox res
    
End Sub

Private Sub Form_Activate()
    txtNomeCliente.SetFocus
    novoPet = False
End Sub

Private Sub Form_Load()
    cmdLimpar_Click
End Sub

'Faz o cadastro do cliente e um pet
Sub Cadastrar()
    
    Dim cadPet As Boolean
    Dim cliente As cliente
    Dim pet As pet
    Dim conn As Object
    Dim dia, mes, ano, strData As String
    
    If txtNome.Text = "" Then
        cadPet = False
    End If
    
    Set pet = New pet
    Set cliente = New cliente
    
    txtCliente.Text = ""
    
    cliente.nome = UCase(Trim(txtNomeCliente.Text))
    cliente.cpf = UCase(Trim(txtCpf.Text))
    cliente.endereco = UCase(Trim(txtEndereco.Text))
    
    strQuery = "INSERT INTO clientes (nome, endereco,cpf,ativo) "
    strQuery = strQuery & "VALUES('" & cliente.nome & "','" & cliente.endereco & " ',' " & cliente.cpf & " ',1) "
    
    Set conn = AbreConn
    conn.Execute strQuery
    
    strQuery = "SELECT ID FROM Clientes "
    strQuery = strQuery & "WHERE Nome = '" & cliente.nome & "'"
    
    Set cursor = New ADODB.Recordset
    
    cursor.Open strQuery, AbreConn
    
    txtCliente.Text = cursor.Fields(0).Value
    
    cursor.Close
    
    If Not cadPet Then

        pet.nome = UCase(Trim(txtNome.Text))
        pet.raca = UCase(Trim(txtRaca.Text))
        pet.sexo = UCase(Trim(txtSexo.Text))
        pet.historico = UCase(Trim(txtHistorico1.Text))
        
        dia = Mid(UCase(Trim(txtNascimento.Text)), 1, 2)
        mes = Mid(UCase(Trim(txtNascimento.Text)), 4, 2)
        ano = Mid(UCase(Trim(txtNascimento.Text)), 7, 4)
        
        strData = ano & "-" & mes & "-" & dia

        pet.nascimento = UCase(Trim(txtNascimento.Text))

'        strQuery = "INSERT INTO pets (nome, raca,sexo,datanascimento) "
'        strQuery = strQuery & "VALUES('" & pet.nome & "','" & pet.raca & " ',' " & pet.sexo & " ',' " & strData & "') "
        
        strQuery = "INSERT INTO pets (nome, raca,sexo,datanascimento,historico) "
        strQuery = strQuery & "VALUES('" & pet.nome & "','" & pet.raca
        strQuery = strQuery & " ',' " & pet.sexo & " ',' " & strData & "','" & pet.historico & "') "

        conn.Execute strQuery

        strQuery = "SELECT ID FROM Pets "
        strQuery = strQuery & "WHERE Nome = '" & pet.nome & "' "
        strQuery = strQuery & "AND Raca = '" & pet.raca & "' "
        strQuery = strQuery & "AND DataNascimento = '" & strData & "' "

        cursor.Open strQuery, AbreConn

        petId = cursor.Fields(0).Value

        cursor.Close
        
        strQuery = "INSERT INTO PetsDonos(Dono, Pet) "
        strQuery = strQuery & "VALUES(" & txtCliente.Text & " , " & petId & ") "
        
        conn.Execute strQuery

    End If

cmdLimpar_Click
    
End Sub

'Adicionar pet
Sub CadastrarNovoPet(ByVal id As Integer)
    
    Dim pet As pet
    Dim conn As Object
    Dim petId As Integer
    Dim dia, mes, ano, strData As String
    
    Set conn = AbreConn
    
    Set pet = New pet
    
    pet.nome = UCase(Trim(txtNome.Text))
    pet.raca = UCase(Trim(txtRaca.Text))
    pet.sexo = UCase(Trim(txtSexo.Text))
    pet.historico = Trim(txtHistorico1.Text)
    
    dia = Mid(UCase(Trim(txtNascimento.Text)), 1, 2)
    mes = Mid(UCase(Trim(txtNascimento.Text)), 4, 2)
    ano = Mid(UCase(Trim(txtNascimento.Text)), 7, 4)

    strData = ano & "-" & mes & "-" & dia

    pet.nascimento = UCase(Trim(txtNascimento.Text))


    strQuery = "INSERT INTO pets (nome, raca,sexo,datanascimento,historico) "
    strQuery = strQuery & "VALUES('" & pet.nome & "','" & pet.raca
    strQuery = strQuery & " ',' " & pet.sexo & " ',' " & strData & "','" & pet.historico & "') "

    conn.Execute strQuery

    strQuery = "SELECT ID FROM Pets "
    strQuery = strQuery & "WHERE Nome = '" & pet.nome & "' "
    strQuery = strQuery & "AND Raca = '" & pet.raca & "' "
    strQuery = strQuery & "AND DataNascimento = '" & strData & "' "
    
    cursor.Open strQuery, AbreConn

    petId = cursor.Fields(0).Value

    cursor.Close

    strQuery = "INSERT INTO PetsDonos(Dono, Pet) "
    strQuery = strQuery & "VALUES(" & id & " , " & petId & ") "

    conn.Execute strQuery
    
    cmdLimpar_Click
    
End Sub

Function Confirma(strMensagem) As Integer

    Confirma = False

    Const MB_OK = 0, MB_OKCANCEL = 1
    Const MB_YESNOCANCEL = 3, MB_YESNO = 4
    Const MB_ICONSTOP = 16, MB_ICONQUESTION = 32
    Const MB_ICONEXCLAMATION = 48, MB_ICONINFORMATION = 64
    Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7

    Dim DgDef, msg, response, Title

    Title = "Confirma��o"

    DgDef = vbYesNo + vbQuestion + MB_DEFBUTTON2

    resposta = MsgBox(strMensagem, DgDef, Title)
    
    If resposta = IDYES Then
        Confirma = True
    End If

End Function

Sub Atualizar()

    Dim atualPet1, atualPet2, atualPet3, atualPet4, atualPet5 As Boolean
    Dim cliente As cliente
    Dim pet, pet2, pet3, pet4, pet5 As pet
    Dim conn As Object
    Dim dia, mes, ano, strData As String
    
    If txtNome.Text = "" Then
        atualPet1 = False
    End If
    If txtNome2.Text = "" Then
        atualPet2 = False
    End If
    If txtNome3.Text = "" Then
        atualPet3 = False
    End If
    If txtNome4.Text = "" Then
        atualPet4 = False
    End If
    If txtNome5.Text = "" Then
        atualPet5 = False
    End If

    
    Set pet = New pet
    Set cliente = New cliente
        
    cliente.nome = UCase(Trim(txtNomeCliente.Text))
    cliente.cpf = UCase(Trim(txtCpf.Text))
    cliente.endereco = UCase(Trim(txtEndereco.Text))
    
    strQuery = "UPDATE clientes "
    strQuery = strQuery & "SET nome = '" & cliente.nome & "', "
    strQuery = strQuery & "endereco = '" & cliente.endereco & "', "
    strQuery = strQuery & "cpf = '" & cliente.cpf & "' "
    strQuery = strQuery & "WHERE ID =" & clienteId
    
    Set conn = AbreConn
    conn.Execute strQuery
    
    pet.nome = UCase(Trim(txtNome.Text))
    pet.raca = UCase(Trim(txtRaca.Text))
    pet.sexo = UCase(Trim(txtSexo.Text))
    pet.historico = UCase(Trim(txtHistorico1.Text))
    
    dia = Mid(UCase(Trim(txtNascimento.Text)), 1, 2)
    mes = Mid(UCase(Trim(txtNascimento.Text)), 4, 2)
    ano = Mid(UCase(Trim(txtNascimento.Text)), 7, 4)
    
    strData = ano & "-" & mes & "-" & dia

    pet.nascimento = UCase(Trim(txtNascimento.Text))
    
        
    strQuery = "UPDATE Pets "
    strQuery = strQuery & "SET nome = '" & pet.nome & "', "
    strQuery = strQuery & "raca = '" & pet.raca & "', "
    strQuery = strQuery & "sexo = '" & pet.sexo & "', "
    strQuery = strQuery & "datanascimento = '" & strData & "', "
    strQuery = strQuery & "historico = '" & pet.historico & "' "
    strQuery = strQuery & "WHERE ID =" & petId
    
    conn.Execute strQuery
    
    MsgBox "OK"
    
    cmdLimpar_Click
    
End Sub


