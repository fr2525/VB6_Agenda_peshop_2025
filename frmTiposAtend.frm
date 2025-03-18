VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServicos 
   Caption         =   "Serviços"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstServicos 
      Height          =   3075
      Left            =   180
      TabIndex        =   12
      Top             =   1260
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5424
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtDuracao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   4
      Top             =   5700
      Width           =   750
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      MaxLength       =   14
      TabIndex        =   3
      Top             =   5700
      Width           =   1860
   End
   Begin VB.TextBox txtServico 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      MaxLength       =   50
      TabIndex        =   2
      Top             =   4800
      Width           =   6270
   End
   Begin Threed.SSCommand cmd_Adicionar 
      Height          =   675
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   " &Novo"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Picture         =   "frmTiposAtend.frx":0000
   End
   Begin Threed.SSCommand cmd_Limpar 
      Height          =   675
      Left            =   1425
      TabIndex        =   5
      Top             =   240
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   " &Limpar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Picture         =   "frmTiposAtend.frx":015A
   End
   Begin Threed.SSCommand cmd_Gravar 
      Height          =   675
      Left            =   2865
      TabIndex        =   6
      Top             =   240
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "&Gravar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BevelWidth      =   1
      Picture         =   "frmTiposAtend.frx":02B4
   End
   Begin Threed.SSCommand cmd_Sair 
      Height          =   675
      Left            =   5550
      TabIndex        =   7
      Top             =   240
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Picture         =   "frmTiposAtend.frx":040E
   End
   Begin Threed.SSCommand cmd_Excluir 
      Height          =   675
      Left            =   4260
      TabIndex        =   9
      Top             =   240
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BevelWidth      =   1
      Picture         =   "frmTiposAtend.frx":0728
   End
   Begin VB.Label Label3 
      Caption         =   "minutos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   6000
      TabIndex        =   1
      Top             =   5730
      Width           =   510
   End
   Begin VB.Label Label2 
      Caption         =   "Duração :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5160
      TabIndex        =   11
      Top             =   5280
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "Valor :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   270
      TabIndex        =   10
      Top             =   5340
      Width           =   840
   End
   Begin VB.Label lbl_Animal 
      Caption         =   "Descrição :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   255
      TabIndex        =   8
      Top             =   4470
      Width           =   1380
   End
End
Attribute VB_Name = "frmServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer
Private Sub Nomes_Colunas()
    With LstServicos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Código", 0, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Descrição", 4000, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Valor", 1400, lvwColumnRight
        .ColumnHeaders.Add 4, , "Duração", 900, lvwColumnRight
    End With
End Sub

Private Sub Dados_Colunas()
    
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT ID,DESCRICAO,valor,tempo_est FROM TAB_servicos ORDER BY DESCRICAO"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, CnnLocal, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            With LstServicos
                .ListItems.Add X, , Rstemp!ID
                
                If Not IsNull(Rstemp!DESCRICAO) Then
                    .ListItems(X).SubItems(1) = Rstemp!DESCRICAO
                Else
                    .ListItems(X).SubItems(1) = ""
                End If
                If Not IsNull(Rstemp!VALOR) Then
                    .ListItems(X).SubItems(2) = Format(Rstemp!VALOR, "###,##0.00")
                Else
                      .ListItems.Add(X).SubItems(2) = "0.00"
                End If
                If Not IsNull(Rstemp!TEMPO_EST) Then
                    .ListItems(X).SubItems(3) = Format(Rstemp!TEMPO_EST, "000")
                Else
                    .ListItems.Add(X).SubItems(3) = "00"
                End If
            End With
            Rstemp.MoveNext
        Next
    Else
        MsgBox "Sem registros", vbOKOnly
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub

Private Sub cmd_Adicionar_Click()
    txtServico.Enabled = True
    txtDuracao.Enabled = True
    txtValor.Enabled = True
    txtServico.SetFocus
    txtServico.Text = ""
    txtDuracao.Text = ""
    txtValor.Text = ""
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = False
    iTipoOperacao = 1

End Sub

Private Sub cmd_Excluir_Click()
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Descrição de serviço inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir o Serviço: " & Chr(13) & Chr(10) & _
                            Trim(LstServicos.SelectedItem.ListSubItems.Item(1)), vbYesNo) = vbYes Then
        If fExcluir_Servico() Then
            cmd_Adicionar.Enabled = True
            cmd_Excluir.Enabled = False
            cmd_Gravar.Enabled = False
            LstServicos.ListItems.Clear
            Call Dados_Colunas
            If LstServicos.ListItems.Count > 0 Then
                LstServicos.ListItems(1).Selected = True
                txtServico.Text = Trim(LstServicos.SelectedItem.ListSubItems.Item(1))
            End If
        Else
            MsgBox "Erro ao excluir o Serviço: " & Err.Description
        End If
    End If

End Sub

Private Sub cmd_Gravar_Click()
    
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
        MsgBox "Descrição de serviço inválida. Favor corrigir", vbOKOnly
        txtServico.SetFocus
        Exit Sub
    End If
    
    If Val(txtValor.Text) = 0 Then
        If MsgBox("Campo Valor do serviço não está preenchido. " & Chr(13) & Chr(10) & "Deseja continuar e gravar assim mesmo? ", vbYesNo) = vbNo Then
            txtValor.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtDuracao.Text) = 0 Then
        If MsgBox("Campo Tempo de duração do serviço não está preenchido. " & Chr(13) & Chr(10) & "Deseja continuar e gravar assim mesmo? ", vbYesNo) = vbNo Then
            txtDuracao.SetFocus
            Exit Sub
        End If
    End If
    
    If fGravar_Servico() Then
        cmd_Adicionar.Enabled = True
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        'cmd_Excluir.Enabled = true
        LstServicos.ListItems.Clear
        Call Dados_Colunas
        LstServicos.ListItems(1).Selected = True
        txtServico.Text = Trim(LstServicos.SelectedItem.ListSubItems.Item(1))
        'Call cmd_Limpar_Click
    Else
        MsgBox "Erro ao incluir o tipo de PET: " & Err.Description
    End If

End Sub

Private Sub cmd_Limpar_Click()
    txtServico.Text = ""
    txtValor.Text = "0.00"
    txtDuracao.Text = ""
    'txtServico.SetFocus
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True

End Sub

Private Sub cmd_Sair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Nomes_Colunas
    Call Dados_Colunas
    'lstServicos.ListItems = 1
    If LstServicos.ListItems.Count > 0 Then
        txtServico.Text = Trim(LstServicos.SelectedItem.ListSubItems.Item(1))
    End If
End Sub

Private Sub lstservicos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtServico.Text = Trim(LstServicos.SelectedItem.ListSubItems.Item(1))
    txtServico.Enabled = True
    txtValor.Text = Format(LstServicos.SelectedItem.ListSubItems.Item(2), "###,##0.00")
    txtValor.Enabled = True
    txtDuracao.Text = LstServicos.SelectedItem.ListSubItems.Item(3)
    txtDuracao.Enabled = True
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = True
    cmd_Limpar.Enabled = True
    iTipoOperacao = 2
End Sub

Private Sub lstservicos_KeyPress(KeyAscii As Integer)
    If LstServicos.ListItems.Count > 0 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtDuracao_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Tipo de Animal inválido. Favor corrigir", vbOKOnly
       txtServico.SetFocus
    Else
       cmd_Gravar.SetFocus
    End If
End If

End Sub

Private Function fGravar_Servico()
    
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Descrição do serviço inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Function
    End If
    
    fGravar_Servico = True
    
    On Error GoTo Erro_fGravar_Servico
    
    'ID,DESCRICAO,valor,tempo_est
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_servicos (DESCRICAO, VALOR, TEMPO_EST, OPERADOR, DT_ATUALIZA)"
        strSql = strSql + " VALUES( '" & UCase(txtServico.Text) & "',"
        strSql = strSql + Replace(txtValor.Text, ",", ".") & "," & txtDuracao.Text & ",'"
        strSql = strSql + sysNomeAcesso & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_servicos SET DESCRICAO = '" & UCase(txtServico.Text) & _
                                          "',VALOR =   " & Replace(txtValor.Text, ",", ".") & _
                                          ",tempo_est = " & txtDuracao.Text & _
                                          ",OPERADOR = '" & sysNomeAcesso & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE ID = '" & LstServicos.SelectedItem.Text & "'"
                                          
    End If
    CnnLocal.Execute strSql
    Exit Function
    
Erro_fGravar_Servico:
    fGravar_Servico = False
End Function

Private Function fExcluir_Servico()
    
    fExcluir_Servico = True
    
    On Error GoTo Erro_fExcluir_Servico
    
    strSql = "DELETE from tab_servicos WHERE ID = '" & LstServicos.SelectedItem.Text & "'"
    CnnLocal.Execute strSql
    Exit Function
Erro_fExcluir_Servico:
    fExcluir_Servico = False
End Function

Private Sub txtServico_GotFocus()
     Call SelText(txtServico)
End Sub

Private Sub txtServico_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        If Len(Trim(txtServico.Text)) = 0 Then
            MsgBox "Obrigatório Informar Descrição do Serviço.", vbInformation, "Aviso"
            txtServico.SetFocus
            Exit Sub
        End If

        SendKeys "{tab}"
    End If
End Sub

Private Sub txtServico_LostFocus()
    txtValor.Text = Format(0, "###,##0.00")
End Sub


Private Sub txtValor_GotFocus()
Call SelText(txtValor)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        If Len(Trim(txtValor.Text)) = 0 Then
            MsgBox "Obrigatório Informar o Valor do Serviço.", vbInformation, "Aviso"
            txtValor.SetFocus
            Exit Sub
        End If

        SendKeys "{tab}"
    End If

End Sub
