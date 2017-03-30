VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Demo de uso da NFe_Util 2G em VB - 2010-12-12"
   ClientHeight    =   8085
   ClientLeft      =   5310
   ClientTop       =   2040
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10335
   Begin VB.CommandButton Command15 
      Caption         =   "Gerar PDF do DANFE"
      Height          =   375
      Left            =   8160
      TabIndex        =   32
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Carta de Corre��o"
      Height          =   375
      Left            =   8160
      TabIndex        =   31
      Top             =   6480
      Width           =   2055
   End
   Begin VB.ComboBox cbVersao 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "DemoVB.frx":0000
      Left            =   4440
      List            =   "DemoVB.frx":0010
      TabIndex        =   29
      Text            =   "2.00"
      Top             =   45
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Monta procNFe"
      Height          =   375
      Left            =   8130
      TabIndex        =   28
      Top             =   3900
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Converte Txt em NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   27
      Top             =   3525
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Inutiliza Nro de NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   26
      Top             =   3135
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Cancelamento de NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   25
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Consulta Staus NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   24
      Top             =   2370
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Protocolo de Autoriza��o"
      Height          =   675
      Left            =   8160
      TabIndex        =   21
      Top             =   5760
      Width           =   2055
      Begin VB.TextBox txtProAutoUso 
         Height          =   285
         Left            =   90
         TabIndex        =   22
         Top             =   270
         Width           =   1830
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nro Recibo Lote"
      Height          =   645
      Left            =   8160
      TabIndex        =   20
      Top             =   5040
      Width           =   2055
      Begin VB.TextBox txtNroRecibo 
         Height          =   285
         Left            =   105
         TabIndex        =   23
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Busca resultado  NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   19
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Envia uma NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   18
      Top             =   4290
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Validar NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   17
      Top             =   840
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Retorno / Mensagem de Retorno  (msgRetWS)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   105
      TabIndex        =   14
      Top             =   4200
      Width           =   7800
      Begin VB.TextBox txtRetorno 
         Height          =   3315
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   270
         Width           =   7515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Entrada / Mensagem Enviada  (msgDados) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   105
      TabIndex        =   13
      Top             =   1065
      Width           =   7800
      Begin VB.TextBox txtEntrada 
         Height          =   2790
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   225
         Width           =   7440
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consulta Status WS"
      Height          =   375
      Left            =   8130
      TabIndex        =   12
      Top             =   1605
      Width           =   2055
   End
   Begin VB.ComboBox cbAmb 
      Height          =   315
      ItemData        =   "DemoVB.frx":002C
      Left            =   6480
      List            =   "DemoVB.frx":0036
      TabIndex        =   11
      Text            =   "Homologa��o"
      Top             =   45
      Width           =   1455
   End
   Begin VB.ComboBox cbUF 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "DemoVB.frx":0051
      Left            =   2880
      List            =   "DemoVB.frx":00A0
      TabIndex        =   10
      Text            =   "SP"
      Top             =   45
      Width           =   735
   End
   Begin VB.ComboBox cbWS 
      Height          =   315
      ItemData        =   "DemoVB.frx":0108
      Left            =   1080
      List            =   "DemoVB.frx":013C
      TabIndex        =   9
      Text            =   "SP"
      Top             =   45
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Certificado Digital"
      Height          =   375
      Left            =   8130
      TabIndex        =   5
      Top             =   75
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Certificado Digital "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   105
      TabIndex        =   3
      Top             =   420
      Width           =   7815
      Begin VB.TextBox txtCertificado 
         Height          =   285
         Left            =   135
         TabIndex        =   4
         Text            =   "CN=M R M KATO ASAKURA - EPP:69621187915, OU=AC CAIXA PJ-1 V1, OU=Caixa Economica Federal, O=ICP-Brasil, C=BR"
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Vers�o da DLL"
      Height          =   375
      Left            =   8130
      TabIndex        =   2
      Top             =   1995
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar XML da NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   1
      Top             =   1230
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Assinar NF-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   0
      Top             =   465
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Vers�o:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   30
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Ambiente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "UF:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Sigla WS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public versaoDLL As String

Private Sub Command1_Click()
'
'
'  Exemplo de gera��o de uma NF-e com as funcionalidades oferecidas pela DLL
'
'  A NF-e � formada por diversas grupos de tags e as grupos obrigat�rios s�o:
'
'  NFe
'    +-----infNFe
'    |          +-----+--ide  (identifica��o da NF-e)
'    |                |
'    |                +--emit (identifica��o do emitente)
'    |                |
'    |                +--dest (identifica��o do destinat�rio)
'    |                |
'    |                +--det
'    |                |    +-------+--prod (detalhe do produto)
'    |                |            |
'    |                |            +--imposto
'    |                |                     +-------+----+----+--ICMS   (informa��es do ICMS)
'    |                |                             |    |    |
'    |                |                             |    |    +--IPI    (informa��es do IPI)
'    |                |                             |    |    |
'    |                |                             |    |    +--II     (informa��es do II)
'    |                |                             |    |
'    |                |                             |    +-------ISS    (informa��es do ISS)
'    |                |                             |
'    |                |                             +--PIS    (informa��es do PIS)
'    |                |                             |
'    |                |                             +--COFINS (informa��es do COFINS)
'    |                +--det
'    |                |    +-------+--prod (detalhe do produto)
'    |                |            |
'    |                |            +--imposto
'    |                |                     +-------+----+----+--ICMS   (informa��es do ICMS)
'    |                |                             |    |    |
'    |                |                             |    |    +--IPI    (informa��es do IPI)
'    |                |                             |    |    |
'    |                |                             |    |    +--II     (informa��es do II)
'    |                |                             |    |
'    |                |                             |    +-------ISS    (informa��es do ISS)
'    |                |                             |
'    |                |                             |
'    |                |                             +--PIS    (informa��es do PIS)
'    |                |                             |
'    |                |                             +--COFINS (informa��es do COFINS)
'    |                +--det
'    |                |    +-------+--prod (detalhe do produto)
'    |                |            |
'    |                |            +--imposto
'    |                |                     +-------+----+----+--ICMS   (informa��es do ICMS)
'    |                |                             |    |    |
'    |                |                             |    |    +--IPI    (informa��es do IPI)
'    |                |                             |    |    |
'    |                |                             |    |    +--II     (informa��es do II)
'    |                |                             |    |
'    |                |                             |    +-------ISS    (informa��es do ISS)
'    |                |                             |
'    |                |                             |
'    |                |                             +--PIS    (informa��es do PIS)
'    |                |                             |
'    |                |                             +--COFINS (informa��es do COFINS)
'    |                |
'    |                +--total (total da NF-e)
'    |                |
'    |                +--transp (Informa��es do Transporte)
'    |                |
'    |                +--infAdic (informa��es adicionais)
'    |
'    +-----Signature  (assinatura digital XML)
'
'  Como este demo tem efeitos meramente did�tico, vamos criar uma NF-e com os campos m�nimos
'  o usu�rio dever� informar os demais campos se necess�rio, a p�gina 85 do Manual de integra��o tem diagrama simplificado da NF-e
'
'  IMPORTANTE: O desenvolvedor deve ter familiaridade com os nomes dos campos da NF-e, sendo altamente recomendada a correta
'              compress�o do leiuate da NF-e e das regras de preenchimento dos respectivos campos.
'
'  A NF-e � uma estrutra de �rvore com o elemento raiz chamada NF-e que tem diversos "galhos/folhas"
'  Para criar a NF-e com a DLL, o usu�rio deve come�ar a criar os itens das extrDim emidadas, ou seja os itens mais internos.
'  Assim, uma boa ordem de cria��o dos grupos seria:
'
'  1. criar o grupo de informa��es do emitente (emit);
'  2. criar o grupo de informa��es de identifica��o da NF-e (ide);
'  3. criar o grupo de informa��es do destinat�rio (dest);
'  4.1 criar o detalhe do produto (prod);
'  4.2 criar o detalhe do ICMS (ICMS);
'  4.3 criar o detalhe do PIS (PIS);
'  4.4 criar o detalhe do COFINS (COFINS);
'  4.5 criar o detalhe do imposto (imposto), consolidar ICMS, PIS e COFINS;
'  4.6 criar o detalhe do item (det), consolidar prod e imposto;
'  5. criar o grupo de Dim total da NF-e (total);
'  6. criar o grupo de informa��es do transporte (transp);
'  7. criar o grupo de informa��es adicionais (infAdic);
'  8. criar o grupo de informa��es da NF-e (infNFe), consolidando ide, emit, dest, det, total e transp
'  9. criar o grupo da NF-e
'
'
'   DECLARA��O DAS VARI�VEIS
'======Identifica��o do documento=======
'
Dim ide As String
Dim ide_cUF As Long
Dim ide_cNF As String               ' o tamanho do campo foi reduzido para 8 d�gitos na vers�o 2.00
Dim ide_natOp As String
Dim ide_indPag As Long
Dim ide_mode As Long
Dim ide_serie As Long
Dim ide_nNF As Long
Dim ide_dEmi As Date
Dim ide_dSaiEnt As Date
Dim ide_tpNF As Long
Dim ide_cMunFG As String
Dim ide_tpImp As Long
Dim ide_cDV As Long
Dim ide_tpAmb As Long
Dim ide_finNFe As Long
Dim ide_tpEmis As Long
Dim ide_procEmi As Long
Dim ide_verProc As String
Dim ide_NFref As String             ' string com os documentos fiscais referenciados
'
' novos campos do vers�o 2.00
'
Dim ide_hSaiEnt As String
Dim ide_dhCont As Date
Dim ide_xJust As String
'
'======  Dados do  Dim emitente==========
'
Dim emi As String
Dim emi_CNPJ As String
Dim emi_CPF As String
Dim emi_xNome As String
Dim emi_xFant As String
Dim emi_xLgr As String
Dim emi_nro As String
Dim emi_xCpl As String
Dim emi_xBairro As String
Dim emi_cMun As String
Dim emi_xMun As String
Dim emi_UF As String
Dim emi_CEP As String
Dim emi_cPais As String
Dim emi_xPais As String
Dim emi_fone As String
Dim emi_IE As String
Dim emi_IEST As String
Dim emi_IM As String
Dim emi_CNAE As String
'
'  campos novos da vers�o 2.00
'
Dim emi_CRT As String
'
'======  Dados do Dim destinat�rio==========
'
Dim dest As String
Dim dest_CNPJ As String
Dim dest_CPF As String
Dim dest_xNome As String
Dim dest_xFant As String
Dim dest_xLgr As String
Dim dest_nro As String
Dim dest_xCpl As String
Dim dest_xBairro As String
Dim dest_cMun As String
Dim dest_xMun As String
Dim dest_UF As String
Dim dest_CEP As String
Dim dest_cPais As String
Dim dest_xPais As String
Dim dest_fone As String
Dim dest_IE As String
Dim dest_IESUF As String
'
'  campos novos da vers�o 2.00
'
Dim dest_email As String
'
'====== Valor Dim total da NF-e ============
'
Dim total As String
Dim totICMS As String
Dim totICMS_vBC As Currency
Dim totICMS_vICMS As Currency
Dim totICMS_vBCST As Currency
Dim totICMS_vST As Currency
Dim totICMS_vProd As Currency
Dim totICMS_vFrete As Currency
Dim totICMS_vSeg As Currency
Dim totICMS_vDesc As Currency
Dim totICMS_vII As Currency
Dim totICMS_vIPI As Currency
Dim totICMS_vPIS As Currency
Dim totICMS_vCOFINS As Currency
Dim totICMS_vOutro As Currency
Dim totICMS_vNF As Currency
'
'======  Dados do Transportador=========
'
Dim transp As String
Dim transpModFrete As String
'
'======  Informa��es Adicionais==========
'
Dim infAdic As String
Dim infAdic_infAdiFisco As String
Dim infAdic_infCPL As String
'
'======== Nota Fiscal ==================
'
Dim NFe As String
Dim ChaveNFe As String
Dim versao As String
Dim id As String
Dim Retirada As String
Dim Entrega As String
Dim Compra As String
Dim Exporta As String
Dim Cobr As String
Dim Detalhes As String
Dim Imposto As String
Dim infAdic_infAdi_infCpl As String
'
'   Grupo novo da vers�o 2.00 do leiaute - aquisi��o de cana
'
Dim Cana As String
'
'======== Detalhe do Produto ==================
'
Dim Prod As String
Dim Prod_cProd As String
Dim Prod_cEAN As String
Dim Prod_xProd As String
Dim Prod_NCM As String
Dim Prod_ExTIPI As String
'
'   Campo eliminado do leiaute na vers�o 2.00
'
' Dim Prod_genero As Long
Dim Prod_CFOP As Long
Dim Prod_uCOM As String
Dim Prod_qCom As Double
Dim Prod_vUnCom As Double
Dim Prod_vProd  As Currency
Dim Prod_cEANTrib As String
Dim Prod_uTrib As String
Dim Prod_qTrib As Double
Dim Prod_vUnTrib As Double
Dim Prod_vFrete As Currency
Dim Prod_vSeguro As Currency
Dim Prod_vDesc As Currency
Dim Prod_DI As String
Dim Prod_DetEspecifico As String
Dim Prod_infAdProd As String
'
'   Campo novo da vers�o 2.00 do leiatue
'
Dim Prod_vOutro As Currency
Dim Prod_indTot As Long
Dim Prod_nItemPed As Long
Dim Prod_xPed As String
'
'=========dados do ICMS===========
'
Dim icms As String
Dim icms_orig As String
Dim icms_CST As String
Dim icms_modBC As Long
Dim icms_pRedBC As Double
Dim icms_vBC As Currency
Dim icms_pICMS As Double
Dim icms_vICMS As Currency
Dim icms_modBCST As Long
Dim icms_pmVAST As Double
Dim icms_pRedBCST As Double
Dim icms_vBCST As Currency
Dim icms_pICMSST As Double
Dim icms_vICMSST As Currency
'
'  campos novos da vers�o 2.00
'
Dim icms_vBCSTRet As Currency
Dim icms_vICMSSTRet As Currency
Dim icms_motDesICMS As Long
Dim icms_pBCOp As Double
Dim icms_UFST As String
Dim icms_pCredSN As Double
Dim icms_vCredICMSSN As Currency
Dim icms_vICMSSTDest As Currency
Dim icms_vBCICMSSTDest As Currency
'
'=========dados do PIS=============
'
Dim pis As String
Dim pis_CST As String
Dim pis_vBC As Currency
Dim pis_pPIS As Double
Dim pis_vPIS As Currency
Dim pis_qBCProd As Currency
Dim pis_vAliqProd As Double
'
'========dados do COFINS============
'
Dim cofins As String
Dim cofins_CST As String
Dim cofins_vBC As Currency
Dim cofins_pCOFINS As Double
Dim cofins_vCOFINS As Currency
Dim cofins_qBCProd As Currency
Dim cofins_vAliqProd As Double
'
'
'====== instancia DLL==================
'
'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

Retirada = ""
Entrega = ""
Compra = ""
Exporta = ""
Cobr = ""
Detalhes = ""
Cana = ""
'
'         cria��o dos grupos
'
'===================grupo de identifica��o do emitente (grupo B do Manual de integra��o - p�gina 113-)=======================
'
'        <>&" s�o caracteres reservados do XML e devem ser evitados ou substitu�dos
'        por &lt; &gy; &amp; &quot;
'
'        Vale ressaltar que algumas aplica��es das UF devem mostrar DIAS &amp; DIAS TENTANDO S/A,
'        pois n�o entedem &amp; como &, assim talvez seja melhor substituir o & por e.
'
emi_CNPJ = "99999999000191"                 ' CNPJ do emitente sem m�scara de formata��o
emi_CPF = ""                                ' CPF do emitente, uso exclusivo do Fisco
emi_xNome = "DIAS e DIAS TENTANDO S/A"      ' Raz�o social do emitente, evitar caracteres acentuados e &
emi_xFant = "DDT"                           ' Nome fantasia
emi_xLgr = "AV PRINCIPAL"                   ' logradouro
emi_nro = "S/N"                             ' n�mero, informar S/N quano inexistente para erro de Schema XML
emi_xCpl = "10 andar"                       ' complemento do endere�o, o conte�do pode ser omitido
emi_xBairro = "CENTRO"                      ' bairro
emi_cMun = "3550308"                        ' c�digo do munic�pio (vide p�gina 171 do manual), deve ser compat�vel com a UF
emi_xMun = "SAO PAULO"                      ' nome do munic�pio
emi_UF = "SP"                               ' sigla da UF
emi_CEP = "01300000"                        ' CEP - sem m�scara
emi_cPais = "1058"                          ' c�digo do pais - deve fixo em 1058 - Brasil
emi_xPais = "Brasil"                        ' nome do pais (Brasil ou BRASIL)

emi_fone = "1133221234"                     ' n�mero do telefone sem m�scara, o tamanho foi aumentado para 14 d�gitos no vers�o 2.00

emi_IE = "123456789011"                     ' Inscri��o Estadual do emitente sem m�scara
emi_IEST = ""                               ' informar a IE ST - Inscri��o Estadual como Substituto Tribut�rio, sem formata��o ou m�scara, quando praticar alguma opera��o como substituto tribut�rio
emi_IM = ""                                 ' informar a Inscri��o Municipal, sem formata��o ou m�scara, quando emitir NF conjugada (presta��o de servi�o com fornecimento de pe�as)
emi_CNAE = ""                               ' informar o CNAE Fiscal, este campo deve ser informado em conjunto com o campo IM e vice-versa, a informa��o de um e omiss�o do outro resulta em falha de Schema XML

emi_CRT = "3"                               ' informar o C�digo de Regime Tribut�rio - CRT, valores v�lidos: 1 - Simples Nacional; 2 - Simples Nacional - excesso de sublimite de receita bruta; 3 - Regime Normal
'
'       gera grupo emi
'
emi = objNFeUtil.emitente2G(emi_CNPJ, emi_CPF, emi_xNome, emi_xFant, emi_xLgr, emi_nro, emi_xCpl, emi_xBairro, emi_cMun, emi_xMun, emi_UF, emi_CEP, emi_cPais, emi_xPais, emi_fone, emi_IE, emi_IEST, emi_IM, emi_CNAE, emi_CRT)

'MsgBox "Grupo do emitente " + emi, vbInformation, "Resultado"

'
'========grupo de identifica��o da NF-e - grupo B do Manual de integra��o - p�ginas 108-
'
'        http://www.flexdocs.com.br/guiaNFe/gerarNFe.ide.identificador2G.html
'
'
ide_cUF = 35                    ' c�digo da UF - tabela do IBGE: 35 - SP, 43 - RS, etc. (vide p�gina 171 do manual)
ide_natOp = "Venda"             ' naturez da opera��o
ide_indPag = 0                  ' 0=pagamento � vista
ide_mode = 55                   ' modelo da nota fiscal eletronica
ide_serie = 0                   ' s�rie �nica = 0
ide_nNF = 1                     ' n�mero da NF-e
ide_dEmi = #11/28/2008#         ' data de emiss�o
ide_dSaiEnt = #12:00:00 AM#     ' data em branco = 30/12/1899
ide_tpNF = 1                    ' n�mero da nota fiscal de sa�da
ide_cMunFG = 3550308            ' c�digo do munic�pio do IBGE de ocorr�ncia do FG do ICMS (vide p�gina 171 do manual)
ide_tpImp = 1                   ' orienta��o da impress�o 1-retrato/2-paisagem
ide_tpAmb = 2                   ' ambiente de envio da NF-e 1-produ��o / 2 - homologa��o
ide_finNFe = 1                  ' finalidade da emiss�o da NF-e 1- NF-e normal
ide_tpEmis = 1                  ' forma de emiss�o da NF-e 1- normal, 2 - conting�ncia FS, 3 - conting�ncia SCAN, etc.
ide_procEmi = 0                 ' identifica��o do processo de emiss�o da NF-e 0 - aplica��o do contribuinte
ide_verProc = "NFe_Util_v1.4"   ' identifica��o da ves�o do processo de emiss�o
ide_NFref = ""                  ' NF referenciada, deve ser informado para nota fiscal complementar, devolu��o, etc. - http://www.flexdocs.com.br/guiaNFe/gerarNFe.ref.html

'
' novos campos do vers�o 2G
'
ide_hSaiEnt = ""                ' hora da sa�da
ide_dhCont = #12:00:00 AM#      ' data e hora de entrada em conting�ncia - informar quanto tpEmis diferente de 1, informe #12:00:00 AM# para deixar vazio em VB
ide_xJust = ""                  ' informar a justificativa de entrada em conting�ncia, deve ser informado sempre que tpEmis for diferente de 1.
'
'     gera a chave de acesso da NF-e
'
'     utilizar a fun��o criaChaveNFe para gerar a chave de acesso, c�digo da NF-e e DV
'
'=========vari�veis de trabalho
'
'
Dim Resultado As Long

Dim cUF, ano, mes, modelo, serie, tpEmis, numero, codigoSeguranca As String
Dim msgResultado As String
Dim cNF As String
Dim cDV As String
cUF = Trim(Str(ide_cUF))
ano = Format(ide_dEmi, "YY")
mes = Format(ide_dEmi, "mm")
modelo = Trim(Str(ide_mode))
serie = Trim(Str(ide_serie))
numero = Trim(Str(ide_nNF))
'
'  par�metro novo da vers�o 2.00
'
tpEmis = Trim(Str(ide_tpEmis))
'
msgResultado = ""
codigoSeguranca = "segredo"
cNF = ""
cDV = ""
ChaveNFe = ""

'
'                                 o retorno do CriaChaveNFe foi alterado - vide: http://www.flexdocs.com.br/guiaNFe/funcao.utilidades.criachaveNFe2G.html
'

If objNFeUtil.CriaChaveNFe2G(cUF, ano, mes, emi_CNPJ, modelo, serie, numero, tpEmis, codigoSeguranca, msgResultado, cNF, cDV, ChaveNFe) <> 5601 Then
   MsgBox "Ocorreu um erro ao gerar a chave de acesso " + msgResultado, vbInformation, "Resultado"
End If

ide_cNF = Val(cNF)                     ' c�digo num�rico que comp�e a chave de acesso, o tamanho do campo foi reduzido para 8 d�gitos na vers�o 2.00
ide_cDV = Val(cDV)                     ' DV da chave de acesso da NF-e
'
'   gera grupo ide
'
ide = objNFeUtil.identificador2G(ide_cUF, ide_cNF, ide_natOp, ide_indPag, ide_mode, ide_serie, ide_nNF, ide_dEmi, ide_dSaiEnt, ide_hSaiEnt, ide_tpNF, ide_cMunFG, ide_NFref, ide_tpImp, ide_tpEmis, ide_cDV, ide_tpAmb, ide_finNFe, ide_procEmi, ide_verProc, ide_dhCont, ide_xJust)
'
'MsgBox "Grupo de identifica��o " + ide, vbInformation, "Resultado"
'
'
'================grupo de identifica��o do destinatario (grupo E do Manual de integra��o - p�ginas 116-)=======================
'
'        <>&" s�o caracteres reservados do XML e devem ser evitados ou substitu�dos
'        por &lt; &gy; &amp; &quot;
'
'        Vale ressaltar que algumas aplica��es das UF devem mostrar DIAS &amp; DIAS TENTANDO S/A,
'        pois n�o entedem &amp; como &, assim talvez seja melhor substituir o & por e.
'
dest_CNPJ = "00000000000191"                 ' CNPJ do destinatario sem m�scara de formata��o
dest_CPF = ""                                ' CPF do destinatario, uso exclusivo do Fisco
dest_xNome = "Banco do Brasil S/A"           ' Raz�o social do destinatario, evitar caracteres acentuados e &
dest_xFant = "BB"                            ' Nome fantasia
dest_xLgr = "Rua Libero Badaro"              ' logradouro
dest_nro = "280"                             ' n�mero, informar S/N quano inexistente para erro de Schema XML
dest_xCpl = "10 andar"                       ' complemento do endere�o, o conte�do pode ser omitido
dest_xBairro = "CENTRO"                      ' bairro
dest_cMun = "3550308"                        ' c�digo do munic�pio (vide p�gina 171 do manual), deve ser compat�vel com a UF
dest_xMun = "SAO PAULO"                      ' nome do munic�pio
dest_UF = "SP"                               ' sigla da UF
dest_CEP = "01315000"                        ' CEP - sem m�scara
dest_cPais = "1058"                          ' c�digo do pais - deve fixo em 1058 - Brasil
dest_xPais = "Brasil"                        ' nome do pais (Brasil ou BRASIL)
dest_fone = "1133221234"                     ' n�mero do telefone sem m�scara, o tamanho do campo foi aumentado para 16 d�gitos na vers�o 2.00
dest_IE = "123456789011"                     ' Inscri��o Estadual do destinatario sem m�scara
dest_IESUF = ""                              ' Inscri��o SUFRAMA
'
' novos campos do vers�o 2.00
'
dest_email = "destinatario@empresa.com.br"   ' Infomrmar o e-mail do destinat�rio
'
'   gera grupo do destinat�rio - vide: http://www.flexdocs.com.br/guiaNFe/gerarNFe.des.destinatario2G.html
'
dest = objNFeUtil.destinatario2G(dest_CNPJ, dest_CPF, dest_xNome, dest_xLgr, dest_nro, dest_xCpl, dest_xBairro, dest_cMun, dest_xMun, dest_UF, dest_CEP, dest_cPais, dest_xPais, dest_fone, dest_IE, dest_IESUF, dest_email)
'MsgBox "Grupo de destinat�rio " + dest, vbInformation, "Resultado"
'
'           INICIALIZA��O
'
'       A DLL n�o acumula os valores do itens
'
'====== Valor total da NF-e ============
'
totICMS_vBC = 0
totICMS_vICMS = 0
totICMS_vBCST = 0
totICMS_vST = 0
totICMS_vProd = 0
totICMS_vFrete = 0
totICMS_vSeg = 0
totICMS_vDesc = 0
totICMS_vII = 0
totICMS_vIPI = 0
totICMS_vPIS = 0
totICMS_vCOFINS = 0
totICMS_vOutro = 0
totICMS_vNF = 0
'
'
'================grupo de detalhe do produto (grupo I01 do Manual de integra��o - p�ginas 120-)=======================
'
'                http://www.flexdocs.com.br/guiaNFe/gerarNFe.detalhe.pro.produto2G.html
'
Prod_cProd = "001152"                       ' c�digo do produto
Prod_cEAN = "7897844200115"                 ' c�digo EAN (0, 8,12, 13 ou 14 caracteres), o conte�do pode ser omitido se o produto n�o tiver EAN
Prod_xProd = "Cola Especial para EPS"       ' c�digo do produto, espa�os em branco consecutivos ou no in�cio ou fim do campo podem gerar erro de Schema XML, al�m de caracteres reservados do XML <>&"'
'
'   campo com nova regra de preenchimento
'
Prod_NCM = "35"                             ' c�digo NCM, informar o C�digo NCM com 8 d�gitos - http://www.mdic.gov.br/sitio/interna/interna.php?area=5&menu=1095#I;
                                            ' informar a posi��o do cap�tulo do NCM (as duas primeiras posi��es do NCM) quando a opera��o n�o for de com�rcio exterior (importa��o/ exporta��o) ou o produto n�o seja tributado pelo IPI;
                                            ' se for servi�os, informar 00
Prod_ExTIPI = ""                            ' ExTipi, especializa��o do c�digo NCM, informar apenas se existir e o NCM completo for informado
'
'  campo exclu�do da NF-e
'
'Prod_genero = 0                            ' informar as duas primeiras posi��es do NCM
Prod_CFOP = "5102"                          ' CFOP do opera��o, causa erro de XML se informado um c�digo inexistente
Prod_uCOM = "UN"                            ' unidade de comercializa��o
Prod_qCom = "10"                            ' quantidade de comercializa��o
Prod_vUnCom = "1"                           ' valor unit�rio de comercializa��o, campo de mera demonstra��o deve ser o resultado da divis�o do vProd / qCom
Prod_vProd = 10                             ' valor do total do item
Prod_cEANTrib = "7897844200115"             ' c�digo EAN (0, 8,12, 13 ou 14 caracteres), o conte�do pode ser omitido se n�o tiver EAN, em geral � o mesmo c�digo do EAN de comercializa��o
Prod_uTrib = "UN"                           ' unidade de tributa��o, na maioria dos casos � id�ntico  ao vUnCom, pode diferente nos casos de produtos sujeitos a ST em que a unidade de pauta � diferente da unidade de comercializa��o
                                            ' Ex. unidade de comercializa��o = 1 pack de lata de cerveja => unidade de tributa��o = 1 lata (pre�o de pauta)
Prod_qTrib = "10"                           ' quantidade de comercializa��o
Prod_vUnTrib = "1"                          ' valor unit�rio de tributa��o, campo de mera demonstra��o deve ser o resultado da divis�o do vProd / qTrib
Prod_vFrete = 0                             ' valor do frete, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vSeguro = 0                            ' valor do seguro, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vDesc = 0                              ' valor do desconto concedido
Prod_DI = ""                                ' dados da importa��o, informar apenas no caso de NF de entrada (importa��o)
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.di.html
Prod_DetEspecifico = ""                     ' dados espec�ficos, informar para medicamento, ve�culos novos, armamentos e combust�veis.
                                            ' veicProd - http://www.flexdocs.com.br/guiaNFe/gerarNFe.vei.html
                                            ' med - http://www.flexdocs.com.br/guiaNFe/gerarNFe.med.html
                                            ' arma - http://www.flexdocs.com.br/guiaNFe/gerarNFe.arm.html
                                            ' comb - http://www.flexdocs.com.br/guiaNFe/gerarNFe.com.html
                                            '
Prod_infAdProd = ""                         ' informa��es adicionais do produto
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.detalhe.html
Prod_indTot = 1                             ' indicador de totaliza��o do valor do produto

Prod_xPed = ""                              ' n�mero do pedido de compra
Prod_nItemPed = 0                           ' n�mero do item do pedido
'
'   campo novo da vers�o 2.00 do leiaute
'
Prod_vOutro = 0
'
'   gera grupo do destinat�rio
'
Prod = objNFeUtil.produto2G(Prod_cProd, Prod_cEAN, Prod_xProd, Prod_NCM, Prod_ExTIPI, Prod_CFOP, Prod_uCOM, Prod_qCom, Prod_vUnCom, Prod_vProd, Prod_cEANTrib, Prod_uTrib, Prod_qTrib, Prod_vUnTrib, Prod_vFrete, Prod_vSeguro, Prod_vDesc, Prod_vOutro, Prod_indTot, Prod_DI, Prod_DetEspecifico, Prod_xPed, Prod_nItemPed)

'MsgBox "Grupo de produto " + prod, vbInformation, "Resultado"
'
'
'=========dados do ICMS (grupo N01 do Manual de integra��o - p�ginas 128-)=====================
'
icms_orig = "0"                             ' Tabela A - origem da mercadoria 0=nacional
icms_CST = "00"                             ' Tabela B - CST=00-tributa��o normal
icms_modBC = 3                              ' modalidade de determina��o da BC = 3-valor da opera��o
icms_pRedBC = 0                             ' percentual de redu��o da BC
icms_vBC = 10                               ' valor da BC do ICMS = vProd + vFrete + vSeguro + vOutro
icms_pICMS = 18                             ' al�quota do ICMS
icms_vICMS = 1.8                            ' valor do ICMS
icms_modBCST = 0                            ' modalidade de determina��o da BC ICMS ST
icms_pmVAST = 0                             ' percentual de valor de margem e valor adicionado
icms_pRedBCST = 0                           ' percentual de redu��o da BC do ICMS ST
icms_vBCST = 0                              ' BC do ICMS ST
icms_pICMSST = 0                            ' percentual do ICMSST
icms_vICMSST = 0                            ' valor do ICMS ST devido
'
'   Campos novos da vers�o 2.00
'
icms_vBCSTRet = 0                           ' informa��o do ICMS retindo anteriormente por ST
icms_vICMSSTRet = 0                         ' estes campos devem ser informado somente no caso do CST = 60 ou CSOSN = 500
'
icms_motDesICMS = 0                         ' motivo de desonera��o do ICMS, s� deve ser informado no caso de CST = 40 (isen��o condicional)
'
icms_pBCOp = 0                              ' campos para uso nos casos de ICMSPart/ICMSST
icms_UFST = ""                              '
icms_vICMSSTDest = 0                        '
icms_vBCICMSSTDest = 0                      '
'
icms_pCredSN = 0                            ' campos exclusivos para emissor optante do Simples Nacional CSOSN= 101, 201 e 900
icms_vCredICMSSN = 0                        ' n�o esquecer de informar o CRT=1

'
'   gera grupo do ICMS
'

icms = objNFeUtil.icms2G(icms_orig, icms_CST, icms_modBC, icms_pRedBC, icms_vBC, icms_pICMS, icms_vICMS, icms_modBCST, icms_pmVAST, icms_pRedBCST, icms_vBCST, icms_pICMSST, icms_vICMSST, icms_vBCSTRet, icms_vICMSSTRet, icms_vBCICMSSTDest, icms_vICMSSTDest, icms_motDesICMS, icms_pBCOp, icms_UFST, icms_pCredSN, icms_vCredICMSSN)

'MsgBox "Grupo de Tributos/ICMS " + icms, vbInformation, "Resultado"

'
'=========dados do PIS (grupo Q do Manual de Integra��o - p�ginas 145) =============
'
pis_CST = "01"
pis_vBC = 10
pis_pPIS = 0.0165
pis_vPIS = 0.16
pis_qBCProd = 0
pis_vAliqProd = 0
'
'   gera grupo do PIS
'
pis = objNFeUtil.pis(pis_CST, pis_vBC, pis_pPIS, pis_vPIS, pis_qBCProd, pis_vAliqProd)

'MsgBox "Grupo de Tributos/PIS " + PIS, vbInformation, "Resultado"

'
'========dados do COFINS (grupo s do Manual de Integra��o - p�ginas 147) ============
'
cofins_CST = "01"
cofins_vBC = 10
cofins_pCOFINS = 0.03
cofins_vCOFINS = 0.3
cofins_qBCProd = 0
cofins_vAliqProd = 0
'
'   gera grupo do PIS
'
cofins = objNFeUtil.cofins(cofins_CST, cofins_vBC, cofins_pCOFINS, cofins_vCOFINS, cofins_qBCProd, cofins_vAliqProd)

'MsgBox "Grupo de Tributos/COFINS " + COFINS, vbInformation, "Resultado"

'
'========dados do IMPOSTO (grupo M do Manual de Integra��o - p�ginas 128) ============
'
Imposto = objNFeUtil.imposto2G(icms, "", "", pis, "", cofins, "", "")
'
'   atualiza��o de total
'
totICMS_vBC = totICMS_vBC + icms_vBC
totICMS_vICMS = totICMS_vICMS + icms_vICMS
totICMS_vBCST = totICMS_vBCST + icms_vBCST
totICMS_vST = totICMS_vST + icms_vICMSST
totICMS_vProd = totICMS_vProd + Prod_vProd
totICMS_vFrete = totICMS_vFrete + Prod_vFrete
totICMS_vSeg = totICMS_vSeg + Prod_vSeguro
totICMS_vDesc = totICMS_vDesc + Prod_vDesc

'
'  vOutro
'
totICMS_vOutro = totICMS_vOutro + Prod_vOutro
'
'
'MsgBox "Grupo de Tributos/COFINS " + COFINS, vbInformation, "Resultado"
'
'========dados do ITEM do detalhe (grupo H do Manual de Integra��o - p�ginas 120-) ============
'
'  item 1
'
Detalhes = objNFeUtil.detalhe(1, Prod, Imposto, Prod_infAdProd)
'MsgBox "Grupo de detalhe do Item " + det, vbInformation, "Resultado"
'
'
'================grupo de detalhe do produto (grupo I01 do Manual de integra��o - p�ginas 120-)=======================
'
'                   exemplo do segundo item ST
'
Prod_cProd = "002871"                       ' c�digo do produto
Prod_cEAN = "7896045512321"                 ' c�digo EAN (0, 8,12, 13 ou 14 caracteres), o conte�do pode ser omitido se n�o tiver EAN
Prod_xProd = "Cerveja da boa"       ' c�digo do produto, espa�os em branco consecutivos ou no in�cio ou fim do campo podem gerar erro de Schema XML, al�m de caracteres reservados do XML <>&"'
'
'   campo com nova regra de preenchimento
'
Prod_NCM = "22"                             ' c�digo NCM, informar o C�digo NCM com 8 d�gitos - http://www.mdic.gov.br/sitio/interna/interna.php?area=5&menu=1095#I;
                                            ' informar a posi��o do cap�tulo do NCM (as duas primeiras posi��es do NCM) quando a opera��o n�o for de com�rcio exterior (importa��o/ exporta��o) ou o produto n�o seja tributado pelo IPI;
                                            ' se for servi�os, informar 00
Prod_ExTIPI = ""                            ' ExTipi, especializa��o do c�digo NCM, informar apenas se existir e o NCM completo for informado
'
'  campo exclu�do da NF-e
'
'Prod_genero = 0                            ' informar as duas primeiras posi��es do NCMProd_CFOP = "5403"                          ' CFOP do opera��o, causa erro de XML se informado um c�digo inexistente
Prod_uCOM = "PAC12"                            ' unidade de comercializa��o
Prod_qCom = 10                              ' quantidade de comercializa��o
Prod_vUnCom = 10                            ' valor unit�rio de comercializa��o, campo de mera demonstra��o deve ser o resultado da divis�o do vProd / qCom
Prod_vProd = 100                            ' valor do total do item
Prod_cEANTrib = "7896045512317"             ' c�digo EAN (0, 8,12, 13 ou 14 caracteres), o conte�do pode ser omitido se n�o tiver EAN, em geral � o mesmo c�digo do EAN de comercializa��o
Prod_uTrib = "LATA"                         ' unidade de tributa��o, na maioria dos casos � id�ntico  ao vUnCom, pode diferente nos casos de produtos sujeitos a ST em que a unidade de pauta � diferente da unidade de comercializa��o
                                            ' Ex. unidade de comercializa��o = 1 pack de lata de cerveja => unidade de tributa��o = 1 lata (pre�o de pauta)
Prod_qTrib = 120                            ' quantidade de comercializa��o
Prod_vUnTrib = 0.8333                       ' valor unit�rio de tributa��o, campo de mera demonstra��o deve ser o resultado da divis�o do vProd / qTrib
Prod_vFrete = 0                             ' valor do frete, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vSeguro = 0                            ' valor do seguro, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vDesc = 0                              ' valor do desconto concedido
Prod_DI = ""                                ' dados da importa��o, informar apenas no caso de NF de entrada (importa��o)
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.di.html
Prod_DetEspecifico = ""                     ' dados espec�ficos, informar para medicamento, ve�culos novos, armamentos e combust�veis.
                                            ' veicProd - http://www.flexdocs.com.br/guiaNFe/gerarNFe.vei.html
                                            ' med - http://www.flexdocs.com.br/guiaNFe/gerarNFe.med.html
                                            ' arma - http://www.flexdocs.com.br/guiaNFe/gerarNFe.arm.html
                                            ' comb - http://www.flexdocs.com.br/guiaNFe/gerarNFe.com.html
                                            '
Prod_infAdProd = ""                         ' informa��es adicionais do produto
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.detalhe.html
'
'   campo novo da vers�o 2.00 do leiaute
'
Prod_vOutro = 0
Prod_xPed = ""                              ' n�mero do pedido de compra
Prod_nItemPed = 0                           ' n�mero do item do pedido
'
'   gera grupo do destinat�rio
'
Prod = objNFeUtil.produto2G(Prod_cProd, Prod_cEAN, Prod_xProd, Prod_NCM, Prod_ExTIPI, Prod_CFOP, Prod_uCOM, Prod_qCom, Prod_vUnCom, Prod_vProd, Prod_cEANTrib, Prod_uTrib, Prod_qTrib, Prod_vUnTrib, Prod_vFrete, Prod_vSeguro, Prod_vDesc, Prod_vOutro, Prod_indTot, Prod_DI, Prod_DetEspecifico, Prod_xPed, Prod_nItemPed)

'MsgBox "Grupo de produto " + prod, vbInformation, "Resultado"
'
'
'=========dados do ICMS (grupo N01 do Manual de integra��o - p�ginas 128)=====================
'
icms_orig = "0"                             ' Tabela A - origem da mercadoria 0=nacional
icms_CST = "10"                             ' Tabela B - CST=10-tRIBUTADA E COM COBRANCA POR ST
icms_modBC = 3                              ' modalidade de determina��o da BC = 3-valor da opera��o
icms_pRedBC = 0                             ' percentual de redu��o da BC
icms_vBC = 100                              ' valor da BC do ICMS = vProd + vFrete + vSeguro + vOutro
icms_pICMS = 18                             ' al�quota do ICMS
icms_vICMS = 18                             ' valor do ICMS
icms_modBCST = 5                            ' modalidade de determina��o da BC ICMS ST
icms_pmVAST = 0                             ' percentual de valor de margem e valor adicionado
icms_pRedBCST = 0                           ' percentual de redu��o da BC do ICMS ST
icms_vBCST = 180                            ' BC do ICMS ST
icms_pICMSST = 18                           ' percentual do ICMSST
icms_vICMSST = 14.4                         ' valor do ICMS ST devido
'
'   Campos novos da vers�o 2.00
'
icms_vBCSTRet = 0                           ' informa��o do ICMS retindo anteriormente por ST
icms_vICMSSTRet = 0                         ' estes campos devem ser informado somente no caso do CST = 60 ou CSOSN = 500
'
icms_motDesICMS = 0                         ' motivo de desonera��o do ICMS, s� deve ser informado no caso de CST = 40 (isen��o condicional)
'
icms_pBCOp = 0                              ' campos para uso nos casos de ICMSPart/ICMSST
icms_UFST = ""                              '
icms_vICMSSTDest = 0                        '
icms_vBCICMSSTDest = 0                      '
'
icms_pCredSN = 0                            ' campos exclusivos para emissor optante do Simples Nacional CSOSN= 101, 201 e 900
icms_vCredICMSSN = 0                        ' n�o esquecer de informar o CRT=1

'
'   gera grupo do ICMS
'
 icms = objNFeUtil.icms2G(icms_orig, icms_CST, icms_modBC, icms_pRedBC, icms_vBC, icms_pICMS, icms_vICMS, icms_modBCST, icms_pmVAST, icms_pRedBCST, icms_vBCST, icms_pICMSST, icms_vICMSST, icms_vBCSTRet, icms_vICMSSTRet, icms_vBCICMSSTDest, icms_vICMSSTDest, icms_motDesICMS, icms_pBCOp, icms_UFST, icms_pCredSN, icms_vCredICMSSN)
 
'MsgBox "Grupo de Tributos/ICMS " + icms, vbInformation, "Resultado"

'
'=========dados do PIS (grupo Q do Manual de Integra��o - p�ginas 145) =============
'
pis_CST = "01"
pis_vBC = 100
pis_pPIS = 0.0165
pis_vPIS = 1.65
pis_qBCProd = 0
pis_vAliqProd = 0
'
'   gera grupo do PIS
'
pis = objNFeUtil.pis(pis_CST, pis_vBC, pis_pPIS, pis_vPIS, pis_qBCProd, pis_vAliqProd)

'MsgBox "Grupo de Tributos/PIS " + PIS, vbInformation, "Resultado"

'
'========dados do COFINS (grupo s do Manual de Integra��o - p�ginas 147) ============
'
cofins_CST = "01"
cofins_vBC = 100
cofins_pCOFINS = 0.03
cofins_vCOFINS = 3
cofins_qBCProd = 0
cofins_vAliqProd = 0
'
'   gera grupo do PIS
'
cofins = objNFeUtil.cofins(cofins_CST, cofins_vBC, cofins_pCOFINS, cofins_vCOFINS, cofins_qBCProd, cofins_vAliqProd)

'MsgBox "Grupo de Tributos/COFINS " + COFINS, vbInformation, "Resultado"

'
'========dados do IMPOSTO (grupo M do Manual de Integra��o - p�ginas 128) ============
'
Imposto = objNFeUtil.imposto2G(icms, "", "", pis, "", cofins, "", "")
'MsgBox "Grupo de Tributos/COFINS " + COFINS, vbInformation, "Resultado"
'
'   atualiza��o de total
'
totICMS_vBC = totICMS_vBC + icms_vBC
totICMS_vICMS = totICMS_vICMS + icms_vICMS
totICMS_vBCST = totICMS_vBCST + icms_vBCST
totICMS_vST = totICMS_vST + icms_vICMSST
totICMS_vProd = totICMS_vProd + Prod_vProd
totICMS_vFrete = totICMS_vFrete + Prod_vFrete
totICMS_vSeg = totICMS_vSeg + Prod_vSeguro
totICMS_vDesc = totICMS_vDesc + Prod_vDesc
'
'  vOutro
'
totICMS_vOutro = totICMS_vOutro + Prod_vOutro
'
'========dados do ITEM do detalhe (grupo H do Manual de Integra��o - p�ginas 120) ============
'
'  item 2
'
Detalhes = Detalhes + objNFeUtil.detalhe(2, Prod, Imposto, Prod_infAdProd)
'MsgBox "Grupo de detalhe do Item 2 " + detalhes, vbInformation, "Resultado"

'
'========dados do total do nota
'
'   totalizar o valor da NF
'
totICMS_vNF = totICMS_vProd + totICMS_vFrete + totICMS_vSeg - totICMS_vDesc + totICMS_vST  ' verificar outros acr�scimos e descontos
'
'
'
totICMS = objNFeUtil.totalICMS(totICMS_vBC, totICMS_vICMS, totICMS_vBCST, totICMS_vST, totICMS_vProd, totICMS_vFrete, totICMS_vSeg, totICMS_vDesc, totICMS_vII, totICMS_vIPI, totICMS_vPIS, totICMS_vCOFINS, totICMS_vOutro, totICMS_vNF)

total = objNFeUtil.total(totICMS, "", "")       ' total da NF-e sem os valors de ISSQN e RetTributos

'MsgBox "Grupo de total " + total, vbInformation, "Resultado"

'
'============dados do transportador
'
transpModFrete = "0"        ' responsabilidade do frete 0-emitente, 1-destinat�rio
transp = objNFeUtil.transportador(transpModFrete, "", "", "", "", "")

'
'==============informa��es adcionais
'
infAdic_infAdiFisco = "NF-eletronica.com/blog - visite o blog da NF-e"
infAdic_infAdi_infCpl = "NF-e emitida em aplicacao DEMO em VB 6.0"
infAdic = objNFeUtil.infAdic(infAdic_infAdiFisco, infAdic_infAdi_infCpl, "", "", "")
'
'=============consolida a NF-e=================
'
versao = "2.00"
NFe = objNFeUtil.NFe2G(versao, ChaveNFe, ide, emi, "", dest, Retirada, Entrega, Detalhes, total, transp, Cobr, infAdic, Exporta, Compra, Cana)

txtRetorno.Text = NFe

MsgBox "Examine o c�digo fonte da aplica��o para compreender a l�gica de uso das funcionalidades de cria��o do XML da NF-e ", vbInformation, "Informa��o"


Set objNFeUtil = Nothing
'
End Sub

Private Sub Command10_Click()
'
'  Cancelamento da NF-e
'
'  Esta funcionaliade deve ser utilizada para cancelar
'  uma NF-e autorizada e ainda n�o tenha ocorrido o fato
'  gerador (circula��o da mercadoria).
'  Ex. falta de mercadoria, diverg�ncia de quantidade, pre�o, etc.
'  desist�ncia do comprador, etc.
'
'  veja detalhes da funcionalidade em: http://www.flexdocs.com.br/guiaNFe/WS.canc.cancelaNF2G.html
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaUF As String
Dim siglaWS As String
Dim certificado As String
'
'  As vari�veis do proxy devem ser informadas se necess�rio
'
'  proxy deve ser informado com o endere�o da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
Dim licenca As String
'
Dim ambiente As Integer
'
' define as vari�veis que passam/recebem informa��es importantes
'
Dim ChaveNFe As String          ' chave da NF-e objeto de cancelamento
Dim ProtAutNFe As String        ' protocolo de autoriza��o de uso
Dim Justificativa As String     ' justificativa de cancelamento
'
'  par�metros novos
'
Dim procCancNFe As String       ' estrturura XML que cont�m o pedido de cancelamento e a homologa��o do cancelamento,
                                ' que deve ser mantido pelo emissor e distribu�do ao destinat�rio.
Dim nProtocoloCanc As String    ' n�mero do protocolo de homomologa��o de cancelamento devolvido pela SEFA
Dim dProtocoloCanc As String    ' data e hora de homologa��o do cancelamento
Dim versao As String            'utilizado para escolha da vers�o do WS


'
'
'  IMPORTANTE: todas as vari�veis utilizadas como par�metro da DLL devem ser inicializadas
'
'
proxy = ""
usuario = ""
senha = ""
licenca = ""
msgDados = ""
msgRetWS = ""
msgResultado = ""
procCancNFe = ""
nProtocoloCanc = ""
dProtocoloCanc = ""

certificado = txtCertificado.Text
              ' informar com o assunto da certificado digital
              ' Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"

siglaWS = cbWS.Text ' se a UF utilizar SEFAZ Virtual, informar SVRS (Ex. RJ, SC, etc.) ou SVAN (Ex. ES, RN, etc.)
 
txtEntrada.Text = ""
txtRetorno.Text = ""

  

ChaveNFe = InputBox("Informe a Chave de Acesso da NF-e objeto de cancelamento", "Cancelamento de NF-e")


If ChaveNFe = "" Then '
        MsgBox "Necess�rio informar a chave de acesso da NF-e para cancelamento da NF-e.", vbCritical, "Aten��o:"
            Exit Sub
End If

ProtAutNFe = InputBox("Informe o n�mero do protocolo da autoriza��o de uso da NF-e objeto de cancelamento", "Cancelamento de NF-e")


If ProtAutNFe = "" Then '
        MsgBox "Necess�rio informar o n�mero do protocolo da autoriza��o de uso da NF-e para cancelamento da NF-e.", vbCritical, "Aten��o:"
            Exit Sub
End If

Justificativa = InputBox("Informe a Justificativa de cancelamento", "Cancelamento de NF-e")


If Len(Justificativa) < 15 Then '
        MsgBox "Necess�rio informar a justificativa com no m�nimo 15 caracteres", vbCritical, "Aten��o:"
            Exit Sub
End If

'
' par�metro novo - utilizado para escolha da vers�o do WS
'
versao = cbVersao.Text ' vers�o do Web Service, a vers�o anterior � 1.07, a vers�o nova � 2.00

If cbAmb.Text = "Produ��o" Then
   ambiente = 1
Else
   ambiente = 2
End If

Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.canc.cancelaNF2G.html

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'  trecho para instanciar a DLL em early binding
'  necessario fazer o referece da DLL
'
'Dim objNFeUtil As NFe_Util_2G.Util
'
'Set objNFeUtil = New NFe_Util_2G.Util
'
'
Screen.MousePointer = vbHourglass    ' ampulheta
'
'
procCancNFe = objNFeUtil.CancelaNF2G(siglaWS, ambiente, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, ChaveNFe, ProtAutNFe, Justificativa, nProtocoloCanc, dProtocoloCanc, proxy, usuario, senha, licenca)
'
'
Screen.MousePointer = vbDefault ' normal
'
' mostra mensagem XML enviada e a mensagem de retorno do WS
'
txtEntrada.Text = msgDados          ' string com a mensagem XML enviado ao WS

txtRetorno.Text = msgRetWS          ' string com a mensagem XML da resposta do WS

If cStat = 101 Then
                                      
   MsgBox msgResultado & Chr(13) & Chr(13) + "Protocolo de homologa��o de cancelamento: " + nProtocoloCanc + Chr(13) & Chr(13) + "Data e hora de homologa��o de cancelamento: " + dProtocoloCanc + Chr(13) & Chr(13) + "Grave o procCancNFe : " + procCancNFe, vbInformation, "Aten��o: Cancelamento da NF-e"
   
   '
   ' grave o procCancNFe, pois o XML deve ser mantido pelo emissor, al�m de ser distribu�do para o destinat�rio tamb�m.
   '

Else

    MsgBox msgResultado & Chr(13) & Chr(13), vbError, "Aten��o: Cancelamento da NF-e Falhou"

End If

End Sub

Private Sub Command11_Click()

'
'  Inutiliza N�mero de NF-e
'
'  A funcionalidade deve ser utilizada para inutilizar um
'  n�mero de NF-e que n�o vai ser utilizada (atribu�da) a
'  NF-e, por salto de numera��o, rejei��o de NF-e, etc.
'
'  veja os detalhes da chamada em: http://www.flexdocs.com.br/guiaNFe/WS.canc.inutilizaNro2G.html
'
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaUF As String
Dim siglaWS As String
Dim certificado As String
'
'  As vari�veis do proxy devem ser informadas se necess�rio
'
'  proxy deve ser informado com o endere�o da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
Dim licenca As String
'
Dim ambiente As Integer
'
' define as vari�veis que passam/recebem informa��es importantes
'
Dim cUF As String               ' c�digo da UF do solicitante - Tabela IBGE
Dim ano As String               ' ano de inutilizal��o da numera��o
Dim CNPJ As String              ' CNPJ do emitente
Dim modelo As String            ' modelo da NF-e (sempre 55)
Dim serie As String             ' serie da NF-e (sem zeros a esquerda)
Dim nInicial As String          ' n�mero inicial da faixa a ser inutilizada (sem zeros a esquerda)
Dim nFinal As String            ' n�mero final da faixa a ser inutilizada (sem zeros a esquerda)
                                ' Observa��es
                                ' s� � permitida a inutiliza��o de at� 1000 n�meros por vez
                                ' se a inutiliza��o for de um �nico n�mero nInicial e nFinal devem
                                ' ser iguais
Dim Justificativa As String     ' justificativa de cancelamento
'
'  par�metros novos
'
Dim procInutNFe As String       ' estrturura XML que cont�m o pedido de inutiliza��o e a homologa��o da inutiliza��o,
                                ' que deve ser mantido pelo emissor.
Dim nProtocoloInut As String    ' n�mero do protocolo de homomologa��o de Inutiliza��o de numera��p devolvido pela SEFA
Dim dProtocoloInut As String    ' data e hora de homologa��o da Inutiliza��o de numera��p
Dim versao As String            'utilizado para escolha da vers�o do WS
'
'
'  IMPORTANTE: todas as vari�veis utilizadas como par�metro da DLL devem ser inicializadas
'
'
proxy = ""
usuario = ""
senha = ""
licenca = ""
msgDados = ""
msgRetWS = ""
msgResultado = ""

procInutNFe = ""
nProtocoloInut = ""
dProtocoloInut = ""

certificado = txtCertificado.Text
              ' informar com o assunto da certificado digital
              ' Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"

siglaWS = cbWS.Text ' se a UF utilizar SEFAZ Virtual, informar SVRS (Ex. RJ, SC, etc.) ou SVAN (Ex. ES, RN, etc.)

siglaUF = cbUF.Text ' pega a sigla da UF
 
txtEntrada.Text = ""
txtRetorno.Text = ""

'
'  Solicita os par�metros da inutiliza��o
'

' o modelo da NF-e � sempre fixo em 55
modelo = "55"
'
' converte a sigla da UF no c�digo da UF da tabela do IBGE
'
cUF = Mid$("11 12 13 14 15 16 17 21 22 23 24 25 26 27 28 29 31 32 33 35 41 42 43 50 51 52 53", InStr(1, "RO AC AM RO PA AP TO MA PI CE RN PB PE AL SE BA MG ES RJ SP PR SC RS MS MT GO DF", siglaUF, 1), 2)
'
'
CNPJ = InputBox("Informe o CNPJ do emissor", "Inutiliza��o de NF-e")
If CNPJ = "" Then '
        MsgBox "Necess�rio informar o CNPJ do emissor", vbCritical, "Aten��o:"
            Exit Sub
End If

ano = InputBox("Informe o Ano (AA) da numera��o que ser� inutilizado", "Inutiliza��o de NF-e")
If ano = "" Then '
        MsgBox "Necess�rio informar o ano de inutiliza��o da numera��o.", vbCritical, "Aten��o:"
            Exit Sub
End If

serie = InputBox("Informe a s�rie da numera��o que ser� inutilizado", "Inutiliza��o de NF-e")
If serie = "" Then ' o certo � verificar se � um n�mero da faixa 0-999, sem zeros a esquerda
        MsgBox "Necess�rio informar a s�rie de inutiliza��o da numera��o.", vbCritical, "Aten��o:"
            Exit Sub
End If

nInicial = InputBox("Informe o n�mero inicial da numera��o que ser� inutilizado", "Inutiliza��o de NF-e")
If nInicial = "" Then ' o certo � verificar se � um n�mero da faixa 1-9999999999, sem zeros a esquerda
        MsgBox "Necess�rio informar o n�mero inicial de inutiliza��o da numera��o.", vbCritical, "Aten��o:"
            Exit Sub
End If

nFinal = InputBox("Informe o n�mero final da numera��o que ser� inutilizado", "Inutiliza��o de NF-e")
If nFinal = "" Then ' o certo � verificar se � um n�mero da faixa 1-9999999999, sem zeros a esquerda
        MsgBox "Necess�rio informar o n�mero inicial de inutiliza��o da numera��o.", vbCritical, "Aten��o:"
            Exit Sub
End If

Justificativa = InputBox("Informe a Justificativa da inutiliza��o com pelo menos 15 caracteres", "Inutiliza��o de NF-e")
If Len(Justificativa) < 15 Then '
        MsgBox "Necess�rio informar a justificativa com no m�nimo 15 caracteres", vbCritical, "Aten��o:"
            Exit Sub
End If

versao = cbVersao.Text ' vers�o do Web Service, a vers�o anterior � 1.07, a vers�o nova � 2.00

If cbAmb.Text = "Produ��o" Then
   ambiente = 1
Else
   ambiente = 2
End If

Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.canc.inutilizaNro2G.html


'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'  trecho para instanciar a DLL em early binding
'  necessario fazer o referece da DLL
'
'Dim objNFeUtil As NFe_Util_2G.Util
'
'Set objNFeUtil = New NFe_Util_2G.Util
'
'
'
'
Screen.MousePointer = vbHourglass    ' ampulheta
'
'
procInutNFe = objNFeUtil.InutilizaNroNF2G(siglaWS, ambiente, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, cUF, ano, CNPJ, modelo, serie, nInicial, nFinal, Justificativa, nProtocoloInut, dProtocoloInut, proxy, usuario, senha, licenca)
'
'
Screen.MousePointer = vbDefault ' normal
'
' mostra mensagem XML enviada e a mensagem de retorno do WS
'
txtEntrada.Text = msgDados          ' string com a mensagem XML enviado ao WS

txtRetorno.Text = msgRetWS          ' string com a mensagem XML da resposta do WS

If cStat = 102 Then
                                      
   MsgBox msgResultado & Chr(13) & Chr(13) + "Protocolo de homologa��o de Inutiliza��o de Numera��o: " + nProtocoloInut + Chr(13) & Chr(13) + "Data e hora de homologa��o de Inutiliza��o de Numera��o: " + dProtocoloInut + Chr(13) & Chr(13) + "Grave o procInutNFe : " + procInutNFe, vbInformation, "Aten��o: Inutiliza��o de Numera��o da NF-e"
   
   '
   ' grave o procInutNFe, pois o XML deve ser mantido pelo emissor.
   '

Else

    MsgBox msgResultado & Chr(13) & Chr(13), vbError, "Aten��o: Inutiliza��o de numera��o da NF-e Falhou"

End If


End Sub

Private Sub Command12_Click()

'
' Converte TXT em XML
'
' A convers�o � limitada a uma NF-e por vez.
'
' veja maiores detalhes da funcionalidade em: http://www.flexdocs.com.br/guiaNFe/TXT2XML.html
'
Dim txt As String
Dim geraChaveNFe As Integer
Dim codigoSeguranca As String
Dim txtNumerado As String
Dim XML As String
Dim msgResultado As String
Dim erros  As String
Dim qtdeErros As Long
'
'  carregar arquivo TXT na string TXT
'
On Error Resume Next

CommonDialog1.DialogTitle = "Escolha o arquivo TXT (vers�o selecionada: " + cbVersao.Text + ")"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Arquivo XML (*.txt)|*.txt|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo para valida��o de Schema XML." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0

Open CommonDialog1.FileName For Input As #1
txt = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = txt

geraChaveNFe = 1               ' gerar chave NF-e (op��o fixa neste exemplo)

qtdeErros = 0                   ' quantidade de erros
codigoSeguranca = "E segredo"   ' c�digo que ser� utilizado para gerar a chave, deve ser a mesma cliente
txtNumerado = ""                ' campo para receber o txt numerado, �til para localizar o erro
erros = ""                      ' relat�rio de erros encontrados
msgResultado = ""               ' literal da mensagem de resultado da chamada da fun��o

Dim Resultado As Long

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'
'
'
'
If cbVersao.Text = "1.10" Then

   Screen.MousePointer = vbHourglass    ' ponteiro ampulheta

   XML = objNFeUtil.Txt2XML(txt, geraChaveNFe, codigoSeguranca, txtNumerado, Resultado, erros, qtdeErros, msgResultado)
   
   Screen.MousePointer = vbDefault      ' ponteiro normal

Else
   If cbVersao.Text = "2.00" Then
   
       Screen.MousePointer = vbHourglass    ' ponteiro ampulheta
       
       XML = objNFeUtil.Txt2XML2G(txt, geraChaveNFe, codigoSeguranca, txtNumerado, Resultado, erros, qtdeErros, msgResultado)
   
       Screen.MousePointer = vbDefault      ' ponteiro normal
   
   Else
       
       MsgBox "A versao deve ser informado com 1.10 ou 2.00...", vbError, "Erro"
       '
       '
       Exit Sub
   
   End If
End If



'  tratar retorno
'
If Resultado = 6901 Then
 
       txtRetorno.Text = XML        ' necess�rio gravar o XML gerado para utilizar posteriormente.

       MsgBox msgResultado + vbCrLf + "Grave o retorno em um arquivo para utiliza-lo.", vbInformation, "Informa��o"

Else

       txtRetorno.Text = "Quantidade de Erros: " + Str(qtdeErros) & vbCrLf & erros
       MsgBox "Processo de valida��o do XML falhou..." & vbCrLf & msgResultado, vbExclamation, "Aten��o"

End If


End Sub

Private Sub Command13_Click()

'
' Monta procNFe
'
' A montagem � limitada a uma NF-e por vez.
' Entradas:
'   NFeAssinada: NF-e em formato XML, deve estar assinada
'   nomeCertificado: Nome do titular do certificado a ser utlizado na conex�o SSL
'    Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"
'   proxy ,usuario e senha: deve ser informado nos casos em que � necess�rio o uso de proxy
'   https://proxyserver:port'; // verificar com o cliente qual � o endere�o do servidor proxy e a porta https, a porta padr�o do https � 443, assim ter�amos algo do
'   tipo 'http://192.168.15.1:443'
'
' Retorno:
'

' msgResultado: literal da mensagem de resultado da chamada da fun��o
'
Dim siglaWS As String
Dim msgResultado As String
Dim NFeAssinada  As String
Dim procNFe As String
Dim retCancNFe As String
Dim nomeCertificado As String
Dim protocolo As String
Dim proxy As String
Dim usuario As String
Dim senha As String
Dim Resultado As Long

'
'  carregar arquivo a NFe na string NFeAssinada
'
On Error Resume Next

CommonDialog1.DialogTitle = "Escolha o arquivo da NF-e"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo para valida��o de Schema XML." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0

Open CommonDialog1.FileName For Input As #1
NFeAssinada = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = NFeAssinada

nomeCertificado = txtCertificado.Text

siglaWS = cbWS.Text             ' necess�rio apenas para a vers�o 2.00
retCancNFe = ""                 ' necess�rio apenas para a vers�o 2.00

proxy = ""                      ' preencher estes campos somente em caso de exist�ncia de proxy na rede
usuario = ""
senha = ""

protocolo = ""                  ' n�mero do protocolo + dd/mm/aa HH:MM:SS

msgResultado = ""               ' literal da mensagem de resultado da chamada da fun��o

Resultado = 0

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")


If cbVersao.Text = "1.10" Then
   '
   Screen.MousePointer = vbHourglass    ' ponteiro ampulheta
   '
   '
   procNFe = objNFeUtil.CriaProcNFe(NFeAssinada, protocolo, Resultado, nomeCertificado, msgResultado, proxy, usuario, senha)
   '
   '
   Screen.MousePointer = vbDefault      ' ponteiro normal
   '
   '  tratar retorno
   '
   If (Resultado = 6201 Or Resultado = 6216 Or Resultado = 6217) Then

       txtRetorno.Text = procNFe

       MsgBox msgResultado, vbInformation, "Informa��o"

   Else

       txtRetorno.Text = ""
       MsgBox "Processo de montagem procNFe falhou..." & vbCrLf & msgResultado, vbExclamation, "Aten��o"

   End If
Else
   If cbVersao.Text = "2.00" Then
      '
      Screen.MousePointer = vbHourglass    ' ponteiro ampulheta
      '
      '
      procNFe = objNFeUtil.CriaProcNFe2G(siglaWS, NFeAssinada, protocolo, retCancNFe, Resultado, nomeCertificado, msgResultado, proxy, usuario, senha)
      '
      '
      Screen.MousePointer = vbDefault      ' ponteiro normal
      '
      '  tratar retorno
      '
      If (Resultado = 6201 Or Resultado = 6216 Or Resultado = 6217) Then

         txtRetorno.Text = procNFe

         MsgBox msgResultado, vbInformation, "Informa��o"

      Else

         txtRetorno.Text = ""
         MsgBox "Processo de montagem procNFe falhou..." & vbCrLf & msgResultado, vbExclamation, "Aten��o"

      End If
    Else

         MsgBox "Vers�o da NF-e selecionada inv�lida, diferente de 1.10 e 2.00..." & vbCrLf & msgResultado, vbError, "Aten��o"
    End If
End If
'
'  liberar DLL
'
Set objNFeUtil = Nothing

End Sub

Private Sub Command14_Click()
     Form2.Show vbModal
End Sub

Private Sub Command15_Click()
     Form3.Show vbModal
End Sub

Private Sub Command2_Click()
'
' detalhes da funcionalidade dispon�vel em: http://www.flexdocs.com.br/guiaNFe/funcao.assinatura.assinar.html
'

Dim XMLString As String
Dim RefUri As String
Dim nomeCertificado As String
Dim XMLAssinado As String
Dim msgResultado As String
'
'
'  IMPORTANTE: todas as vari�veis utilizadas como par�metro da DLL devem ser inicializadas
'
'

On Error Resume Next
CommonDialog1.DialogTitle = "Escolha o arquivo XML"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo para assinatura." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0

Open CommonDialog1.FileName For Input As #1
XMLString = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = XMLString

RefUri = "infNFe"           ' indica a tag a ser assinada
msgResultado = ""
'nomeCertificado = "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"
nomeCertificado = txtCertificado.Text

Dim Resultado As Long

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

Screen.MousePointer = vbHourglass    ' ponteiro ampulheta
'
'  houve altera��o nos par�metros de retorno, agora o XMLAssinado � devolvido pela funcionalidade
'
XMLAssinado = objNFeUtil.Assinar(XMLString, RefUri, nomeCertificado, Resultado, msgResultado)
'
'
Screen.MousePointer = vbDefault      ' ponteiro normal
'
'  tratar retorno
'
If Resultado = 5300 Then

txtRetorno.Text = XMLAssinado
MsgBox msgResultado, vbInformation, "Informa��o"

Else

MsgBox "Processo de assinatura falhou..." & vbCrLf & msgResultado, vbExclamation, "Aten��o"

End If
'
'  liberar DLL
'
Set objNFeUtil = Nothing
End Sub

Private Sub Command3_Click()
'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
MsgBox "Vers�o em uso: " & objNFeUtil.versao, vbInformation, "Informa��o"

'
'  liberar DLL
'
Set objNFeUtil = Nothing


End Sub

Private Sub Command4_Click()
'
' Exemplo para escolher um certificado digital do reposit�rio de certificados digitais do usu�rio corrente do
' Windows
'
' Importante ressaltar que n�o � necess�rio executar esta funcioanlidade antes das chamadas da DLL, ofere�a esta funcionalidade apenas
' para a escolha do certificado digital que ser� utilizada na configura��o da aplica��o.
'
' Tamb�m vale observrar que existe uma funcionaliade que retorna a data de fim da validado do certificado digital que � mais interessante de ser utilizada
' http://www.flexdocs.com.br/guiaNFe/funcao.certificado.propriedade.html
'
' veja detalhes do uso da funcionalidade em http://www.flexdocs.com.br/guiaNFe/funcao.certificado.pegar.html
'
Dim Resultado As Long
Dim msgResultado As String
Dim certificado As String
certificado = ""
msgResultado = ""
'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
' pega certificado
'
' o texto que retorna no campo Certificado ser� utilizada para identificar
' o certificado digital em uso para as demais chamadas que necessitam de
' um certificado digital
'
Resultado = objNFeUtil.PegaNomeCertificado(certificado, msgResultado)
If Resultado < 5403 Then
   If InStr(1, certificado, "Associacao", vbTextCompare) > 0 Then
      MsgBox "O certificado digital da Associa��o n�o � um certificado v�lido para consumir os WS da NF-e! Procure adquirir um certificado digital v�lido para prosseguir com os testes...", vbInformation, "Resultado"
   End If
   txtCertificado.Text = certificado     ' escolhido um certificado digital
Else
   MsgBox msgResultado, vbInformation, "Resultado"
End If

'
' libera classe
'
Set objNFeUtil = Nothing
End Sub

Private Sub Command5_Click()
'
' ConsultaStatus2G: Consulta Situa��o do Web Service de Recep��o de NF-e
'
'
'
' declara��o das vari�veis que ser�o utilizadas na passagem de par�metros da DLL
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaUF As String
Dim siglaWS As String
Dim certificado As String
'
' As vari�veis do proxy devem ser informadas se necess�rio
'
' proxy deve ser informado com o endere�o da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
'
Dim ambiente As Integer
'
' par�metro novo - utilizado para escolha da vers�o do WS
'
Dim versao As String
'
' IMPORTANTE: todas as vari�veis utilizadas como par�metro da DLL devem ser inicializadas
'
'
proxy = ""
usuario = ""
senha = ""
msgDados = ""
msgRetWS = ""
'
' prepara vari�veis
'
certificado = txtCertificado.Text
siglaWS = cbWS.Text
siglaUF = cbUF.Text
'
' par�metro novo - utilizado para escolha da vers�o do WS
'
versao = cbVersao.Text ' vers�o do Web Service, a vers�o anterior � 1.07, a vers�o nova � 2.00

If cbAmb.Text = "Produ��o" Then
   ambiente = 1
Else
   ambiente = 2
End If
   
txtEntrada.Text = ""
txtRetorno.Text = ""

Dim cStat As Long
'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'
'
'
Screen.MousePointer = vbHourglass    ' ampulheta
'
'
cStat = objNFeUtil.ConsultaStatus2G(siglaWS, siglaUF, ambiente, certificado, versao, msgDados, msgRetWS, msgResultado, proxy, usuario, senha)
'
'
Screen.MousePointer = vbDefault ' normal
'
' mostra mensagem XML enviada e a mensagem de retorno do WS
'
txtEntrada.Text = msgDados
txtRetorno.Text = msgRetWS
MsgBox msgResultado + Chr(13) + Chr(13) + msgRetWS, vbInformation, "Resultado da Consulta Status do Servi�o"
'
' libera classe
'
Set objNFeUtil = Nothing
End Sub

Private Sub Command6_Click()
'
' ValidaXML:  Valida Schema XML
'
' Veja detalhes da funcionalidade em: http://www.flexdocs.com.br/guiaNFe/funcao.XML.validar.html
'
' ********IMPORTANTE O tipoXML da vers�o 2G � 19 ***************
'
Dim XML As String
Dim msgResultado As String
Dim erroXML  As String
Dim tipoXML  As Long
Dim qtdeErros As Long
'
'  carregar arquivo XML na string
'
On Error Resume Next

CommonDialog1.DialogTitle = "Escolha o arquivo XML"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo para valida��o de Schema XML." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0

Open CommonDialog1.FileName For Input As #1
XML = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = XML

tipoXML = 19    ' validar NF-e (op��o fixa para validar NF-e da vers�o 2.00 neste exemplo, se for validar a vers�o 1.10 da NF-e informe 1)
qtdeErros = 0   ' quantidade de erros, se o XML n�o estiver assinado vai ocorrer um erro
erroXML = ""
msgResultado = ""

Dim Resultado As Long

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
Screen.MousePointer = vbHourglass    ' ponteiro ampulheta
'
'
Resultado = objNFeUtil.ValidaXML(XML, tipoXML, msgResultado, qtdeErros, erroXML)
'
'
Screen.MousePointer = vbDefault      ' ponteiro normal
'
'  tratar retorno
'
If (Resultado = 5501) Then

txtRetorno.Text = ""
MsgBox msgResultado, vbInformation, "Informa��o"

ElseIf (Resultado = 5506) Then

txtRetorno.Text = ""
MsgBox "XML da NF-e sem assinatura...", vbInformation, "Informa��o"

Else

txtRetorno.Text = "Quantidade de Erros: " + Str(qtdeErros) & vbCrLf & erroXML
MsgBox "Processo de valida��o do XML falhou..." & vbCrLf & msgResultado, vbExclamation, "Aten��o"

End If
'
'  liberar DLL
'
Set objNFeUtil = Nothing



End Sub

Private Sub Command7_Click()

' EnviaNFe2G: Envio de uma �nica NF-e
'
' para mais detalhes da funcionalidade acesse: http://www.flexdocs.com.br/guiaNFe/WS.NFe.enviaNFe2G.html
'
'
' declara��o das vari�veis que ser�o utilizadas na passagem de par�metros da DLL
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaWS As String
Dim certificado As String

'
' As vari�veis do proxy devem ser informadas se necess�rio
'
' proxy deve ser informado com o endere�o da url : porta, ex: 192.168.15.1:443
'
Dim proxy   As String
Dim usuario As String
Dim senha As String
'
' licenca - deve ser informado com a chave da licen�a de uso para acessar os WS de produ��o
'
Dim licenca As String
'
' NFe         - informar com a NF-e a ser transmitida, n�o � necess�rio validar nem assinar
'
' NFeAssinada - � devolvido pela DLL se a chamada for realizada com sucesso.
'
' NroRecibo   - � devolvido pela DLL se a NF-e for transmitida corretamente,
'               este n�mero � necess�rio para buscar o resultado do processamento da NF-e
'               ========== o NroRecibo n�o indica que a NF-e foi autorizada ==============
Dim NFe As String
Dim NFeAssinada As String
Dim nroRecibo As String
Dim cStat As Long
Dim versao As String
Dim dhRecibo As String
Dim tMed As String
Dim x As Boolean
'
' IMPORTANTE: todas as vari�veis utilizadas como par�metro da DLL devem ser inicializadas
'
'
Dim nomeArquivo As String

proxy = ""
usuario = ""
senha = ""
msgDados = ""
msgRetWS = ""
licenca = ""
NFe = ""
NFeAssinada = ""
nroRecibo = ""
dhRecibo = ""
tMed = ""
'
' prepara vari�veis
'
certificado = txtCertificado.Text
siglaWS = cbWS.Text
versao = cbVersao.Text

'
' nesta chamada da DLL pega o ambiente que est� informado na NF-e
' assim tomar cuidado para n�o enviar a NF-e para o ambiente de produ��o
'
txtEntrada.Text = ""
txtRetorno.Text = ""

'
'  carregar arquivo XML na string NFe
'
On Error Resume Next

CommonDialog1.DialogTitle = "Escolha da NF-e (sem assinatura)"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo XML da NF-e para transmiss�o." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0
'
'  carrega o conte�do do arquivo xml em NFe
'
Open CommonDialog1.FileName For Input As #1
NFe = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = NFe

nomeArquivo = Replace(CommonDialog1.FileName, ".", "_assinado.")

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'
'
Screen.MousePointer = vbHourglass    ' ampulheta
'
'
Do

  NFeAssinada = objNFeUtil.EnviaNFe2G(siglaWS, NFe, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, nroRecibo, dhRecibo, tMed, proxy, usuario, senha, licenca)
  '
  '  se cStat = 105 - Lote em processamento, significa que a SEFAZ ainda n�o conseguiu processar a NF-e e a aplica��o deve persistir at� que o cStat seja diferent de 105
  '
  If cStat <> 105 Then
     Exit Do
  End If
Loop
'
'
Screen.MousePointer = vbDefault ' normal
'
' mostra mensagem XML enviada e a mensagem de retorno do WS
'
txtEntrada.Text = msgDados          ' string com a mensagem XML enviado ao WS

txtRetorno.Text = msgRetWS          ' string com a mensagem XML da resposta do WS
                                      
txtNroRecibo.Text = nroRecibo       ' n�mero do recibo do lote entregue
                                    ' este n�mero � necess�rio para fazer a consulta
                                    ' do resultado do processamento do lote enviado
                                    
' an�lise do retorno da chamada da fun��o enviaNFeSCAN
'
' resultado:
'
' 5000-7000 ------> falha na chamda da fun��o, vide a mensagem de erro e corrija o problema - http://www.flexdocs.com.br/guiaNFe/WS.NFe.enviaNFe2G.html
'
' 103 -------> SUCESSO! - LOTE RECEBIDO pelo WS, guarde o NroRecibo e efetua a busca do resultado do processamento
'
'              IMPORTANTE, ter o n�mero do recibo n�o significa que a NF-e est� autorizada, ainda � necess�rio buscar o resultado do processamento
'
' 108/109 ---> WS com problemas, sem condi��es de receber lotes
' 2xx -------> FALHA - existe algum problema com o lote enviado, verifique o c�digo do erro e corrigir o erro
'
Select Case cStat

       Case 5000 To 7003                ' problemas no envio da DLL para o WS
        
        If (MsgBox("Descri��o do erro: " & Chr(13) & Chr(13) + msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Aten��o: O envio da NF-e falhou...")) = vbYes Then
        
            x = Salva_Log("EnviaNFe2G", msgResultado, versao, msgDados, msgRetWS, NFe)
            
        End If
       
       Case 103
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "A NF-e foi enviada para o Web Service desejado, busque o resultado processamento do lote, informando o n�mero do recibo de entrega: " + nroRecibo, vbInformation, "Aten��o: NF-e enviada!"
       
       x = Salva_NF(NFeAssinada, nomeArquivo)    ' Salva NF-e assinada
       
       Case 108, 109
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "O Web Service com problemas de recep��o, adote o processo de conting�ncia se a emiss�o da NF-e for urgente." + nroRecibo, vbInformation, "Aten��o: WS com problemas na recep��o!"
       
       Case Else                                ' problemas na recep��o do WS

        If (MsgBox("Descri��o do erro: " & Chr(13) & Chr(13) & cStat & " - " & msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Aten��o: O WS rejeitou a NF-e (lote) ...")) = vbYes Then
        
            x = Salva_Log("EnviaNFe2G", msgResultado, versao, msgDados, msgRetWS, NFe)
            
        End If


End Select


'
' libera classe
'
Set objNFeUtil = Nothing
End Sub

Private Sub Command8_Click()
'
'   buscaNFe2G - funcionalidade para buscar o resultado do processamento da NF-e enviada
'   com uso do enviaNFe2G, necess�rio informar a NF-e assinada.
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaWS As String
Dim certificado As String
Dim cStat As Long
Dim x As Boolean
'
'  As vari�veis do proxy devem ser informadas se necess�rio
'
'  proxy deve ser informado com o endere�o da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
'
Dim tpAmbiente As Long
'
' chave da licenca de uso da DLL
'
Dim licenca As String
'
' define as vari�veis que passam/recebem informa��es importantes
'
Dim NFeAssinada As String
Dim nroRecibo As String
Dim procNFe As String       ' procNFe -> NF-e + protocolo de autoriza��o de uso da NF-e, deve ser mantido em arquivo e distribu�do ao destinat�rio.
'
' par�metros novos
'
Dim versao As String        ' vers�o do leiaute (1.10 ou 2.00), serve para escolher o WS da SEFAZ
Dim nroProtocolo As String  ' n�mero do protocolo de autoriza��o de uso da NF-e, este n�mero � necess�rio para cancelar a NF-e
Dim dhProtocolo As String   ' data e hora de autoriza��o de uso da NF-e
Dim cMsg As String          ' c�digo da mensagem da SEFAZ, a SEFAZ pode utiliza-lo como canal de comunica��o com o emissor
Dim xMsg As String          ' literal da mensagem da SEFAZ
'
'
'  IMPORTANTE: todas as vari�veis utilizadas como par�metro da DLL devem ser inicializadas
'
'
proxy = ""
usuario = ""
senha = ""
msgDados = ""
msgRetWS = ""
msgResultado = ""
procNFe = ""
nroProtocolo = ""
dhProtocolo = ""
cMsg = ""
xMsg = ""



certificado = txtCertificado.Text
              ' informar com o assunto da certificado digital
              ' Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"

siglaWS = cbWS.Text ' se a UF utilizar SEFAZ Virtual, informar SVRS (Ex. RJ, SC, etc.) ou SVAN (Ex. ES, RN, etc.)


versao = cbVersao.Text      ' vers�o da NF-e: 1.10 ou 2.00
  
txtEntrada.Text = ""
txtRetorno.Text = ""


' NFeAssinada = ' informar com a NFeAssinada que retornou do enviaNFeSCAN

'
'  carregar arquivo XML da NF-e assinada na string NFeAssinada, a NF-e assinada
'  � necess�rio para montar o procNFe
'
On Error Resume Next

CommonDialog1.DialogTitle = "Informe a NF-e Assinada (para montar o procNFe)"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo XML da NF-e para transmiss�o." & vbCrLf & _
            Err.Description
        Exit Sub
End If

On Error GoTo 0
'
'  carrega o conte�do do arquivo xml em NFe
'
Open CommonDialog1.FileName For Input As #1
NFeAssinada = Input$(LOF(1), 1)
Close #1

' nroRecibo =   ' informar com o nroRecibo que retorno do enviaNFeSCAN

If txtNroRecibo.Text = "" Then
        MsgBox "Necess�rio informar o n�mero do recibo para buscar o resultado do processamento!", vbCritical, "Aten��o:"
            Exit Sub
End If

nroRecibo = txtNroRecibo.Text

If cbAmb.Text = "Produ��o" Then
   tpAmbiente = 1
Else
   tpAmbiente = 2
End If

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
Screen.MousePointer = vbHourglass    ' ampulheta
'
'

Do

   procNFe = objNFeUtil.BuscaNFe2G(siglaWS, tpAmbiente, NFeAssinada, nroRecibo, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, nroProtocolo, dhProtocolo, cMsg, xMsg, proxy, usuario, senha, licenca)
   
   '
   '  se cStat = 105 - Lote em processamento, significa que a SEFAZ ainda n�o conseguiu processar a NF-e e a aplica��o deve persistir at� que o cStat seja diferent de 105
   '
   If cStat <> 105 Then
     Exit Do
   End If

Loop

'
Screen.MousePointer = vbDefault ' normal
'
' mostra mensagem XML enviada e a mensagem de retorno do WS
'
txtEntrada.Text = msgDados          ' string com a mensagem XML enviado ao WS

txtRetorno.Text = msgRetWS          ' string com a mensagem XML da resposta do WS
                                      
                                    
'
'   tratar o resultado da chamada:
'
'
'           WS chamada com sucesso
'
'           105 � lote em processamento -> tentar novamente
'           106 � lote n�o localizado   -> tentar enviar o lote novamente ou verificar se o nroRecibo est� correto
'           100 � NF-e autorizada       -> OK
'           2xx � motivo de rejei��o do WS -> erro na elabora��o da NF-e, verificar o c�digo de erro e corrigir a NF-e
Select Case cStat

       Case 5001 To 6423                ' erro da DLL
        
        If (MsgBox("Descri��o do erro: " & Chr(13) & Chr(13) + msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Aten��o: O envio da NF-e falhou...")) = vbYes Then
        
            x = Salva_Log("BuscaNFe2G", msgResultado, versao, msgDados, msgRetWS, NFeAssinada)
            
        End If
       
       Case 100    ' NF-e autorizada
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "A NF-e autorizada, guarde o n�mero do protocolo de autoriza��o de uso (" + nroRecibo + ") e respectivo procNFe (NF-e + autoriza��o de uso).", vbInformation, "Aten��o: NF-e autorizada!"
       
       x = Salva_NF_UTF8(procNFe, nroRecibo & "-nfeproc.xml")    ' Salva NF-e assinada
       
       txtProAutoUso.Text = nroProtocolo   ' n�mero do potocolo de autorizacao de uso
                                           ' este n�mero � necess�rio para um eventual cancelamento da NF-e.
       
       Case 101    ' NF-e denegada
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "A NF-e foi denegada, guarde o n�mero do protocolo de denega��o de uso (" + nroRecibo + ") e respectivo procNFe (NF-e + denega��o de uso).", vbInformation, "Aten��o: NF-e denegada!"
       
       x = Salva_NF_UTF8(procNFe, nroRecibo & "-nfeproc.xml")    ' Salva NF-e assinada
       
       txtProAutoUso.Text = nroProtocolo   ' n�mero do potocolo de denegacao de uso
              
       Case 106    ' Lote n�o localizado, verifique se o n�mero do recibdo est� correto
       
       MsgBox msgResultado, vbInformation, "Aten��o: Lote n�o localizado"
       
       Case 108, 109
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "O Web Service est� com problemas de recep��o, adote o processo de conting�ncia se a emiss�o da NF-e for urgente." + nroRecibo, vbInformation, "Aten��o: WS com problemas na recep��o!"
       
       Case Else

        If (MsgBox("Descri��o do erro: " & Chr(13) & Chr(13) & cStat & " - " & msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Aten��o: O WS rejeitou a NF-e (lote) ...")) = vbYes Then
        
            x = Salva_Log("BuscaNFe2G", msgResultado, versao, msgDados, msgRetWS, NFeAssinada)
                        
        End If


End Select

' libera classe
'
Set objNFeUtil = Nothing
End Sub

Private Sub Command9_Click()
'
'  Consulta Status da NF-e
'
'
Dim msgDados As String
'Dim msgCabec As String       n�o � mais necess�rio
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaUF As String
Dim siglaWS As String
Dim certificado As String
'
'  As vari�veis do proxy devem ser informadas se necess�rio
'
'  proxy deve ser informado com o endere�o da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
'
Dim ambiente As Integer
'
' par�metro novo - utilizado para escolha da vers�o do WS
'
Dim versao As String
'
' define as vari�veis que passam/recebem informa��es importantes
'
Dim ChaveNFe As String
'
'
'  IMPORTANTE: todas as vari�veis utilizadas como par�metro da DLL devem ser inicializadas
'
'
proxy = ""
usuario = ""
senha = ""
msgDados = ""
'msgCabec = ""  n�o � mais necess�rio
msgRetWS = ""
msgResultado = ""

certificado = txtCertificado.Text
              ' informar com o assunto da certificado digital
              ' Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"

siglaWS = cbWS.Text ' se a UF utilizar SEFAZ Virtual, informar SVRS (Ex. RJ, SC, etc.) ou SVAN (Ex. ES, RN, etc.)

versao = cbVersao.Text ' vers�o do Web Service, a vers�o anterior � 1.07, a vers�o nova � 2.00

txtEntrada.Text = ""
txtRetorno.Text = ""
 
ChaveNFe = InputBox("Informe a Chave de Acesso da NF-e", "Consulta Status NF-e")


If ChaveNFe = "" Then '
        MsgBox "Necess�rio informar a chave de acesso da NF-e para consultar o Status da NF-e.", vbCritical, "Aten��o:"
            Exit Sub
End If

If cbAmb.Text = "Produ��o" Then
   ambiente = 1
Else
   ambiente = 2
End If

Dim Resultado As Long

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'  trecho para instanciar a DLL em early binding
'  necessario fazer o referece da DLL
'
'Dim objNFeUtil As NFe_Util_2G.Util
'
'Set objNFeUtil = New NFe_Util_2G.Util
'
'
'
'
Screen.MousePointer = vbHourglass    ' ampulheta
'
'
Resultado = objNFeUtil.ConsultaNF2G(siglaWS, ambiente, certificado, versao, msgDados, msgRetWS, msgResultado, ChaveNFe, proxy, usuario, senha)
'
'
Screen.MousePointer = vbDefault ' normal
'
' mostra mensagem XML enviada e a mensagem de retorno do WS
'
txtEntrada.Text = msgDados          ' string com a mensagem XML enviado ao WS

txtRetorno.Text = msgRetWS          ' string com a mensagem XML da resposta do WS
                                      
MsgBox msgResultado & Chr(13) & Chr(13), vbInformation, "Aten��o: Consulta Status da NF-e"

End Sub

Private Sub Form_Load()

'
'   Exemplo para obter vers�o da DLL em uso
'
'
' instancia classe
'
On Error GoTo InexisteDLL
'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'  trecho para instanciar a DLL em early binding
'  necessario fazer o referece da DLL
'
'Dim objNFeUtil As NFe_Util_2G.Util
'
'Set objNFeUtil = New NFe_Util_2G.Util
'
'
'
' obtem vers�o
'

' MsgBox "A vers�o da DLL �: " + objNFeUtil.versao, vbInformation, "Resultado"

versaoDLL = objNFeUtil.versao

Form1.Caption = "Demo VB 6.0 da " & versaoDLL

MsgBox "Aplica��o Demo em VB 6.0 para demostrar o uso da: " & Chr(13) & Chr(13) & _
        versaoDLL & Chr(13) & Chr(13) & _
       "Para detalhes da forma correta de instala��o da DLL e uso em VB 6.0 " & Chr(13) & _
       "queira examinar o GUIA DE USO DA DLL, dispon�vel em: " & Chr(13) & _
       "http://www.flexdocs.com.br/guiaNFe" & Chr(13) & Chr(13) & _
       "www.flexdocs.com.br (c) 2008-2011 - todos os direitos reservados", vbInformation, "Informa��o da aplica��o"

'
' libera classe
'
Set objNFeUtil = Nothing

Exit Sub

InexisteDLL:

If (Err.Number = -2147024894) Then
   MsgBox "Erro n�mero : " & Str$(Err.Number) & " (80070002/80040002) " & Err.Description & Chr(13) & Chr(13) & "A DLL NFe_Util e as pastas URL e Schemas devem ser copiadas para a pasta: " + App.Path & Chr(13) & Chr(13) & "Em modo depura��o tamb�m � necess�rio que estes arquivos existam no pasta do VB98", vbCritical, "Erro"
ElseIf (Err.Number = -2146234304) Then
   MsgBox "Erro n�mero : " & Str$(Err.Number) & " (80131040/80231040) " & Err.Description & Chr(13) & "A vers�o da DLL existente na pasta da aplica��o � diferente da registrada no equipamento...", vbCritical, "Erro"
Else
   MsgBox "Erro n�mero : " & Str$(Err.Number) & " " & Err.Description & Chr(13) & "Anote o c�digo de erro e verifique se existe solu��o na FAQ.", vbCritical, "Erro"

End If

End Sub

Private Function Salva_Log(ByVal Funcionalidade As String, ByVal msgResultado As String, ByVal msgCabec As String, ByVal msgDados As String, ByVal msgRetWS As String, NFe As String) As Boolean

On Error Resume Next

Salva_Log = True
CommonDialog1.DialogTitle = "Informe o arquivo para gravar o log de erro"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = "LogErro" & Funcionalidade & Format$(Now, "_yyyy-mm-dd_hh.mm") & ".txt"
CommonDialog1.Filter = "Arquivo TXT (*.txt)|*.txt|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowSave

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   Salva_Log = False
   Exit Function

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo de log para grava��o." & vbCrLf & _
            Err.Description
         Salva_Log = False
        Exit Function
End If
On Error GoTo 0

Open CommonDialog1.FileName For Output As #1
Print #1, "LOG DE ERRO da chamada: " & Funcionalidade
Print #1, "----------------------------------------------"
Print #1, "1.Data do incidente:  "; Now
Print #1, "----------------------------------------------"
Print #1, "2.Status de retorno da fun��o:"
Print #1, "----------------------------------------------"
Print #1, msgResultado
Print #1, "----------------------------------------------"
Print #1, "3.XML da NF-e:"
Print #1, "----------------------------------------------"
Print #1, NFe
Print #1, "----------------------------------------------"
Print #1, "4.�rea de Dados:"
Print #1, "----------------------------------------------"
Print #1, UTF8_Encode(msgDados)
Print #1, "----------------------------------------------"
Print #1, "5.�rea de Retorno do WS:"
Print #1, "----------------------------------------------"
If msgRetWS = "" Then
Print #1, "***SEM RETORNO***"
Else
Print #1, UTF8_Encode(msgRetWS)
End If
Print #1, "6.Vers�o da DLL em uso:"
Print #1, "----------------------------------------------"
Print #1, versaoDLL
Close #1


End Function

Private Function Salva_NF(ByVal NFe As String, ByVal Nome As String) As Boolean

On Error Resume Next

Salva_NF = True
CommonDialog1.DialogTitle = "Informe o nome do arquivo para gravar a NF-e"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = Nome
CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowSave

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   Salva_NF = False
   Exit Function

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o nome do arquivo XML da NF-e para grava��o." & vbCrLf & _
            Err.Description
         Salva_NF = False
        Exit Function
End If
On Error GoTo 0

Open CommonDialog1.FileName For Output As #1
Print #1, NFe
Close #1


End Function

Private Function Salva_NF_UTF8(ByVal NFe As String, ByVal Nome As String) As Boolean

Dim Salva_NF As Boolean


On Error Resume Next

Salva_NF = True
CommonDialog1.DialogTitle = "Informe o nome do arquivo para gravar a NF-e"
CommonDialog1.InitDir = App.Path
CommonDialog1.FileName = Nome
CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml|Qualquer arquivo (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
CommonDialog1.CancelError = True
CommonDialog1.ShowSave

If Err.Number = cdlCancel Then 'cancelado pelo usu�rio
   Salva_NF = False
   Exit Function

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o nome do arquivo XML da NF-e para grava��o." & vbCrLf & _
            Err.Description
         Salva_NF = False
        Exit Function
End If
On Error GoTo 0

Open CommonDialog1.FileName For Output As #1
Print #1, UTF8_Encode(NFe)
Close #1


End Function

'
'  Converte a string para codifica��o UTF-8
'
'  Este processo evita problemas de leitura via browser
'  e principalmente no visualizador da RFB
'
Private Function UTF8_Encode(ByVal sStr As String)
    Dim l As Long, lChar As Integer, sUtf8 As String
    For l = 1 To Len(sStr)
        lChar = AscW(Mid(sStr, l, 1))
        If lChar < 128 Then
            sUtf8 = sUtf8 + Mid(sStr, l, 1)
        ElseIf ((lChar > 127) And (lChar < 2048)) Then
            sUtf8 = sUtf8 + Chr(((lChar \ 64) Or 192))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
        Else
            sUtf8 = sUtf8 + Chr(((lChar \ 144) Or 234))
            sUtf8 = sUtf8 + Chr((((lChar \ 64) And 63) Or 128))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
        End If
    Next l
    UTF8_Encode = sUtf8
End Function
