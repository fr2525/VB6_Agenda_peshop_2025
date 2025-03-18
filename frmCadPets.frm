VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmCadPets 
   Caption         =   "Cadastro de Pets"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCuidEspec 
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
      Height          =   780
      Left            =   7410
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3060
      Width           =   4110
   End
   Begin VB.TextBox TxtObserv 
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
      Height          =   780
      Left            =   7380
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4170
      Width           =   4110
   End
   Begin VB.TextBox txtDtNasc 
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
      Left            =   7410
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2520
      Width           =   1170
   End
   Begin VB.ComboBox cmbTipos 
      Height          =   315
      Left            =   9450
      TabIndex        =   9
      Text            =   "Tipos"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.ComboBox cmbDonos 
      Height          =   315
      Left            =   7410
      TabIndex        =   6
      Text            =   "Donos"
      Top             =   1350
      Visible         =   0   'False
      Width           =   4110
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
      Left            =   7410
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1950
      Width           =   4110
   End
   Begin Threed.SSCommand cmd_Adicionar 
      Height          =   675
      Left            =   240
      TabIndex        =   0
      Top             =   390
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
      Picture         =   "frmCadPets.frx":0000
   End
   Begin Threed.SSCommand cmd_Limpar 
      Height          =   675
      Left            =   1245
      TabIndex        =   1
      Top             =   390
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
      Picture         =   "frmCadPets.frx":015A
   End
   Begin Threed.SSCommand cmd_Gravar 
      Height          =   675
      Left            =   2235
      TabIndex        =   2
      Top             =   390
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
      Picture         =   "frmCadPets.frx":02B4
   End
   Begin Threed.SSCommand cmd_Sair 
      Height          =   675
      Left            =   4170
      TabIndex        =   3
      Top             =   390
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
      Picture         =   "frmCadPets.frx":040E
   End
   Begin Threed.SSCommand cmd_Excluir 
      Height          =   675
      Left            =   3225
      TabIndex        =   16
      Top             =   390
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
      Picture         =   "frmCadPets.frx":0728
   End
   Begin VB.Label Label5 
      Caption         =   "Cuidados Especiais :"
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
      Height          =   480
      Left            =   6165
      TabIndex        =   14
      Top             =   3090
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Observ. :"
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
      Left            =   6300
      TabIndex        =   12
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Dt.Nasc.  :"
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
      Left            =   6195
      TabIndex        =   10
      Top             =   2550
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo :"
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
      Left            =   8760
      TabIndex        =   8
      Top             =   2550
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Proprietário :"
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
      Left            =   5940
      TabIndex        =   7
      Top             =   1380
      Width           =   1425
   End
   Begin VB.Label lbl_Animal 
      Caption         =   "Pet  :"
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
      Left            =   6765
      TabIndex        =   5
      Top             =   1920
      Width           =   600
   End
End
Attribute VB_Name = "frmCadPets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
