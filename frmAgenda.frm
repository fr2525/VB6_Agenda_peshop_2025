VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAgenda 
   Caption         =   "Agenda Pet Shop"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   LinkTopic       =   "Form6"
   ScaleHeight     =   7065
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   270
      TabIndex        =   20
      Top             =   360
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   -2147483647
      CalendarTitleBackColor=   -2147483632
      CalendarTitleForeColor=   16776960
      CalendarTrailingForeColor=   128
      Format          =   124190720
      CurrentDate     =   42902
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalhe"
      Height          =   5805
      Left            =   5640
      TabIndex        =   6
      Top             =   960
      Width           =   6225
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2310
         Width           =   3780
      End
      Begin VB.ComboBox cmbServicos 
         Height          =   315
         Left            =   3840
         TabIndex        =   26
         Text            =   "Servicos"
         Top             =   3120
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox cmbDonos 
         Height          =   315
         Left            =   3120
         TabIndex        =   25
         Text            =   "Donos"
         Top             =   5130
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox cmbPets 
         Height          =   315
         Left            =   30
         TabIndex        =   24
         Text            =   "Pets"
         Top             =   5310
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox CmbHorario 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAgenda.frx":0000
         Left            =   1440
         List            =   "frmAgenda.frx":0002
         TabIndex        =   22
         Text            =   "00:00"
         Top             =   390
         Width           =   1125
      End
      Begin VB.TextBox txtAnimal 
         Alignment       =   2  'Center
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   13
         Top             =   900
         Width           =   3780
      End
      Begin VB.TextBox txtDono 
         Alignment       =   2  'Center
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1350
         Width           =   3780
      End
      Begin VB.TextBox txtTipoAtend 
         Alignment       =   2  'Center
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1800
         Width           =   3780
      End
      Begin VB.TextBox txtObeserva 
         Alignment       =   2  'Center
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
         Height          =   1680
         Left            =   1440
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   3870
         Width           =   3780
      End
      Begin VB.TextBox txtValor 
         Alignment       =   2  'Center
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   9
         Top             =   3270
         Width           =   1410
      End
      Begin VB.OptionButton OptSim 
         Caption         =   "Sim"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2790
         Width           =   645
      End
      Begin VB.OptionButton OptNao 
         Caption         =   "Não"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   2790
         Width           =   705
      End
      Begin Threed.SSCommand cmdNovoPet 
         Height          =   360
         Left            =   5370
         TabIndex        =   23
         Top             =   900
         Visible         =   0   'False
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   635
         _StockProps     =   78
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
         Picture         =   "frmAgenda.frx":0004
      End
      Begin VB.Label Label2 
         Caption         =   "Especial :"
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
         Left            =   360
         TabIndex        =   29
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Horário :"
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
         Left            =   420
         TabIndex        =   21
         Top             =   390
         Width           =   900
      End
      Begin VB.Label lbl_Animal 
         Caption         =   "Pet :"
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
         Left            =   810
         TabIndex        =   19
         Top             =   900
         Width           =   600
      End
      Begin VB.Label lbl_Dono 
         Caption         =   "Dono :"
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
         Left            =   660
         TabIndex        =   18
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label lbl_TipoAtend 
         Caption         =   "Serviço :"
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
         Left            =   420
         TabIndex        =   17
         Top             =   1830
         Width           =   1140
      End
      Begin VB.Label lbl_Atendido 
         Alignment       =   2  'Center
         Caption         =   "Atendido :"
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
         Height          =   300
         Left            =   180
         TabIndex        =   16
         Top             =   2850
         Width           =   1275
      End
      Begin VB.Label lbl_Obseerv 
         Alignment       =   2  'Center
         Caption         =   "Observ :"
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
         Left            =   300
         TabIndex        =   15
         Top             =   3870
         Width           =   1170
      End
      Begin VB.Label lbl_Valor 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   450
         TabIndex        =   14
         Top             =   3270
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista"
      Height          =   5805
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   5205
      Begin MSComctlLib.ListView List_atendimentos 
         Height          =   5415
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   9551
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin Threed.SSCommand cmd_Adicionar 
      Height          =   675
      Left            =   7740
      TabIndex        =   0
      Top             =   180
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
      Picture         =   "frmAgenda.frx":015E
   End
   Begin Threed.SSCommand cmd_Limpar 
      Height          =   675
      Left            =   8805
      TabIndex        =   1
      Top             =   180
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
      Picture         =   "frmAgenda.frx":02B8
   End
   Begin Threed.SSCommand cmd_Gravar 
      Height          =   675
      Left            =   9825
      TabIndex        =   2
      Top             =   180
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
      Picture         =   "frmAgenda.frx":0412
   End
   Begin Threed.SSCommand cmd_Sair 
      Height          =   675
      Left            =   10860
      TabIndex        =   3
      Top             =   180
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
      Picture         =   "frmAgenda.frx":056C
   End
   Begin Threed.SSCommand cmd_Serviços 
      Height          =   675
      Left            =   6720
      TabIndex        =   27
      Top             =   180
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "ser&Viços"
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
      Picture         =   "frmAgenda.frx":0886
   End
   Begin Threed.SSCommand cmd_Tipos 
      Height          =   675
      Left            =   5670
      TabIndex        =   28
      Top             =   180
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   " &Tipos"
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
      Picture         =   "frmAgenda.frx":09E0
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aLista_Horario(48) As String

Private Sub Carrega_Colunas_Atendimentos()
    With List_atendimentos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Horário", 700, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Pet", 1900, lvwColumnLeft
        '.ColumnHeaders.Add 3, , "Dono", 1900, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Atendimento", 1900, lvwColumnLeft
        '.ColumnHeaders.Add 5, , "Duração", 800, lvwColumnLeft
        '.ColumnHeaders.Add 6, , "Atendido", 850, lvwColumnLeft
        '.ColumnHeaders.Add 7, , "Observações", 4400, lvwColumnLeft
        '.ColumnHeaders.Add 8, , "Valor", 1200, lvwColumnRight
    End With
End Sub

Private Sub MontaAtendimentos(pData As Date)
    
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT A.*, B.* FROM TAB_ATENDIMENTOS A , TAB_ANIMAIS B, TAB_TIPOS_AN C, "
    strSql = strSql & " TAB_SERVICOS D"
    strSql = strSql & " WHERE A.DT_ATEND = " & pData
    'strSql = strSql & " AND A.DATA_PED >= '" & Format(mskDe.Text, "mm/dd/yyyy") & "'"
    'strSql = strSql & " AND A.DATA_PED <= '" & Format(mskAte.Text, "mm/dd/yyyy") & "'"
   
   i = 1 'initialize counter untuk show data

'query = "SELECT * FROM users"

row = sqlite_get_table(DBz, strSql, minfo) ' query database
numrows = number_of_rows_from_last_call ' bilangan rows data yang di select
 
    
'    Set Rstemp = New ADODB.Recordset
'    Rstemp.Open strSql, CnnLocal, 1, 2
'    If Rstemp.RecordCount <> 0 Then
'        Rstemp.MoveLast
'        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
    If numrows > 0 Then
        For x = 1 To numrows
            If Not IsNull(row(x, 0)) Then
                List_atendimentos.ListItems.Add x, , Format(row(x, 0), "DD/MM/YYYY")
            Else
                List_atendimentos.ListItems.Add x, , ""
            End If
            If Not IsNull(Rstemp(0)) Then
                List_atendimentos.ListItems(x).SubItems(1) = Rstemp(0)
            Else
                List_atendimentos.ListItems(x).SubItems(1) = ""
            End If
            If Not IsNull(Rstemp!RAZAO_SOCIAL) Then
                List_atendimentos.ListItems(x).SubItems(2) = UCase(Rstemp!RAZAO_SOCIAL)
            Else
                  List_atendimentos.ListItems.Add(x).SubItems(2) = "Fornecedor não Encontrado...!"
            End If
            If Not IsNull(Rstemp!VALOR_TOTAL) Then
                List_atendimentos.ListItems(x).SubItems(3) = Format(Rstemp!VALOR_TOTAL, "0.00")
            Else
                List_atendimentos.ListItems.Add(x).SubItems(3) = ""
            End If
            
            Rstemp.MoveNext
        Next
        
'        For X = 1 To Rstemp.RecordCount
'            If Not IsNull(Rstemp!DATA_PED) Then
'                List_atendimentos.ListItems.Add X, , Format(Rstemp!DATA_PED, "DD/MM/YYYY")
'            Else
'                List_atendimentos.ListItems.Add X, , ""
'            End If
'            If Not IsNull(Rstemp(0)) Then
'                List_atendimentos.ListItems(X).SubItems(1) = Rstemp(0)
'            Else
'                List_atendimentos.ListItems(X).SubItems(1) = ""
'            End If
'            If Not IsNull(Rstemp!RAZAO_SOCIAL) Then
'                List_atendimentos.ListItems(X).SubItems(2) = UCase(Rstemp!RAZAO_SOCIAL)
'            Else
'                  List_atendimentos.ListItems.Add(X).SubItems(2) = "Fornecedor não Encontrado...!"
'            End If
'            If Not IsNull(Rstemp!VALOR_TOTAL) Then
'                List_atendimentos.ListItems(X).SubItems(3) = Format(Rstemp!VALOR_TOTAL, "0.00")
'            Else
'                List_atendimentos.ListItems.Add(X).SubItems(3) = ""
'            End If
'
'            Rstemp.MoveNext
'        Next
        
        
        List_atendimentos.SetFocus
       
    Else
        'MsgBox "Sem Atendimentos para a data selecionada", vbOKOnly
        'fmeListaPedidos.Visible = False
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub

Private Sub cmd_Sair_Click()
    Call closeDB
    Unload Me
End Sub

Private Sub cmd_Serviços_Click()
    frmServicos.Show vbModal
End Sub

Private Sub cmd_Tipos_Click()
    frmTipos.Show vbModal
End Sub

Private Sub DTPicker1_Change()
   Print DTPicker1.Value
   Call MontaAtendimentos(DTPicker1.Value)
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim sHora As String
   sHora = "07:00"   ' Estabelecemos um horario inicial que depopis pode ser parametrizado
   
   aLista_Horario(0) = sHora
   For i = 0 To 25   ' Vai até as 20:00 - Podemos ver parametrização depois
     sHora = DateAdd("n", 30, CDate(sHora))
     aLista_Horario(i) = Mid(sHora, 1, 5)
     CmbHorario.AddItem (aLista_Horario(i))
   Next
   
   Call Carrega_Colunas_Atendimentos
   Call MontaAtendimentos(Date)
   
End Sub

