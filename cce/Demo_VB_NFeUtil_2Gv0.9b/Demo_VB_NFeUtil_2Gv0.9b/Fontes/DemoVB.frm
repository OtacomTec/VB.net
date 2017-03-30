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
      Caption         =   "Carta de Correção"
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
      Caption         =   "Protocolo de Autorização"
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
      Text            =   "Homologação"
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
      Caption         =   "Versão da DLL"
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
      Caption         =   "Versão:"
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
'  Exemplo de geração de uma NF-e com as funcionalidades oferecidas pela DLL
'
'  A NF-e é formada por diversas grupos de tags e as grupos obrigatórios são:
'
'  NFe
'    +-----infNFe
'    |          +-----+--ide  (identificação da NF-e)
'    |                |
'    |                +--emit (identificação do emitente)
'    |                |
'    |                +--dest (identificação do destinatário)
'    |                |
'    |                +--det
'    |                |    +-------+--prod (detalhe do produto)
'    |                |            |
'    |                |            +--imposto
'    |                |                     +-------+----+----+--ICMS   (informações do ICMS)
'    |                |                             |    |    |
'    |                |                             |    |    +--IPI    (informações do IPI)
'    |                |                             |    |    |
'    |                |                             |    |    +--II     (informações do II)
'    |                |                             |    |
'    |                |                             |    +-------ISS    (informações do ISS)
'    |                |                             |
'    |                |                             +--PIS    (informações do PIS)
'    |                |                             |
'    |                |                             +--COFINS (informações do COFINS)
'    |                +--det
'    |                |    +-------+--prod (detalhe do produto)
'    |                |            |
'    |                |            +--imposto
'    |                |                     +-------+----+----+--ICMS   (informações do ICMS)
'    |                |                             |    |    |
'    |                |                             |    |    +--IPI    (informações do IPI)
'    |                |                             |    |    |
'    |                |                             |    |    +--II     (informações do II)
'    |                |                             |    |
'    |                |                             |    +-------ISS    (informações do ISS)
'    |                |                             |
'    |                |                             |
'    |                |                             +--PIS    (informações do PIS)
'    |                |                             |
'    |                |                             +--COFINS (informações do COFINS)
'    |                +--det
'    |                |    +-------+--prod (detalhe do produto)
'    |                |            |
'    |                |            +--imposto
'    |                |                     +-------+----+----+--ICMS   (informações do ICMS)
'    |                |                             |    |    |
'    |                |                             |    |    +--IPI    (informações do IPI)
'    |                |                             |    |    |
'    |                |                             |    |    +--II     (informações do II)
'    |                |                             |    |
'    |                |                             |    +-------ISS    (informações do ISS)
'    |                |                             |
'    |                |                             |
'    |                |                             +--PIS    (informações do PIS)
'    |                |                             |
'    |                |                             +--COFINS (informações do COFINS)
'    |                |
'    |                +--total (total da NF-e)
'    |                |
'    |                +--transp (Informações do Transporte)
'    |                |
'    |                +--infAdic (informações adicionais)
'    |
'    +-----Signature  (assinatura digital XML)
'
'  Como este demo tem efeitos meramente didático, vamos criar uma NF-e com os campos mínimos
'  o usuário deverá informar os demais campos se necessário, a página 85 do Manual de integração tem diagrama simplificado da NF-e
'
'  IMPORTANTE: O desenvolvedor deve ter familiaridade com os nomes dos campos da NF-e, sendo altamente recomendada a correta
'              compressão do leiuate da NF-e e das regras de preenchimento dos respectivos campos.
'
'  A NF-e é uma estrutra de árvore com o elemento raiz chamada NF-e que tem diversos "galhos/folhas"
'  Para criar a NF-e com a DLL, o usuário deve começar a criar os itens das extrDim emidadas, ou seja os itens mais internos.
'  Assim, uma boa ordem de criação dos grupos seria:
'
'  1. criar o grupo de informações do emitente (emit);
'  2. criar o grupo de informações de identificação da NF-e (ide);
'  3. criar o grupo de informações do destinatário (dest);
'  4.1 criar o detalhe do produto (prod);
'  4.2 criar o detalhe do ICMS (ICMS);
'  4.3 criar o detalhe do PIS (PIS);
'  4.4 criar o detalhe do COFINS (COFINS);
'  4.5 criar o detalhe do imposto (imposto), consolidar ICMS, PIS e COFINS;
'  4.6 criar o detalhe do item (det), consolidar prod e imposto;
'  5. criar o grupo de Dim total da NF-e (total);
'  6. criar o grupo de informações do transporte (transp);
'  7. criar o grupo de informações adicionais (infAdic);
'  8. criar o grupo de informações da NF-e (infNFe), consolidando ide, emit, dest, det, total e transp
'  9. criar o grupo da NF-e
'
'
'   DECLARAÇÃO DAS VARIÁVEIS
'======Identificação do documento=======
'
Dim ide As String
Dim ide_cUF As Long
Dim ide_cNF As String               ' o tamanho do campo foi reduzido para 8 dígitos na versão 2.00
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
' novos campos do versão 2.00
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
'  campos novos da versão 2.00
'
Dim emi_CRT As String
'
'======  Dados do Dim destinatário==========
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
'  campos novos da versão 2.00
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
'======  Informações Adicionais==========
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
'   Grupo novo da versão 2.00 do leiaute - aquisição de cana
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
'   Campo eliminado do leiaute na versão 2.00
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
'   Campo novo da versão 2.00 do leiatue
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
'  campos novos da versão 2.00
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
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
'         criação dos grupos
'
'===================grupo de identificação do emitente (grupo B do Manual de integração - página 113-)=======================
'
'        <>&" são caracteres reservados do XML e devem ser evitados ou substituídos
'        por &lt; &gy; &amp; &quot;
'
'        Vale ressaltar que algumas aplicações das UF devem mostrar DIAS &amp; DIAS TENTANDO S/A,
'        pois não entedem &amp; como &, assim talvez seja melhor substituir o & por e.
'
emi_CNPJ = "99999999000191"                 ' CNPJ do emitente sem máscara de formatação
emi_CPF = ""                                ' CPF do emitente, uso exclusivo do Fisco
emi_xNome = "DIAS e DIAS TENTANDO S/A"      ' Razão social do emitente, evitar caracteres acentuados e &
emi_xFant = "DDT"                           ' Nome fantasia
emi_xLgr = "AV PRINCIPAL"                   ' logradouro
emi_nro = "S/N"                             ' número, informar S/N quano inexistente para erro de Schema XML
emi_xCpl = "10 andar"                       ' complemento do endereço, o conteúdo pode ser omitido
emi_xBairro = "CENTRO"                      ' bairro
emi_cMun = "3550308"                        ' código do município (vide página 171 do manual), deve ser compatível com a UF
emi_xMun = "SAO PAULO"                      ' nome do município
emi_UF = "SP"                               ' sigla da UF
emi_CEP = "01300000"                        ' CEP - sem máscara
emi_cPais = "1058"                          ' código do pais - deve fixo em 1058 - Brasil
emi_xPais = "Brasil"                        ' nome do pais (Brasil ou BRASIL)

emi_fone = "1133221234"                     ' número do telefone sem máscara, o tamanho foi aumentado para 14 dígitos no versão 2.00

emi_IE = "123456789011"                     ' Inscrição Estadual do emitente sem máscara
emi_IEST = ""                               ' informar a IE ST - Inscrição Estadual como Substituto Tributário, sem formatação ou máscara, quando praticar alguma operação como substituto tributário
emi_IM = ""                                 ' informar a Inscrição Municipal, sem formatação ou máscara, quando emitir NF conjugada (prestação de serviço com fornecimento de peças)
emi_CNAE = ""                               ' informar o CNAE Fiscal, este campo deve ser informado em conjunto com o campo IM e vice-versa, a informação de um e omissão do outro resulta em falha de Schema XML

emi_CRT = "3"                               ' informar o Código de Regime Tributário - CRT, valores válidos: 1 - Simples Nacional; 2 - Simples Nacional - excesso de sublimite de receita bruta; 3 - Regime Normal
'
'       gera grupo emi
'
emi = objNFeUtil.emitente2G(emi_CNPJ, emi_CPF, emi_xNome, emi_xFant, emi_xLgr, emi_nro, emi_xCpl, emi_xBairro, emi_cMun, emi_xMun, emi_UF, emi_CEP, emi_cPais, emi_xPais, emi_fone, emi_IE, emi_IEST, emi_IM, emi_CNAE, emi_CRT)

'MsgBox "Grupo do emitente " + emi, vbInformation, "Resultado"

'
'========grupo de identificação da NF-e - grupo B do Manual de integração - páginas 108-
'
'        http://www.flexdocs.com.br/guiaNFe/gerarNFe.ide.identificador2G.html
'
'
ide_cUF = 35                    ' código da UF - tabela do IBGE: 35 - SP, 43 - RS, etc. (vide página 171 do manual)
ide_natOp = "Venda"             ' naturez da operação
ide_indPag = 0                  ' 0=pagamento à vista
ide_mode = 55                   ' modelo da nota fiscal eletronica
ide_serie = 0                   ' série única = 0
ide_nNF = 1                     ' número da NF-e
ide_dEmi = #11/28/2008#         ' data de emissão
ide_dSaiEnt = #12:00:00 AM#     ' data em branco = 30/12/1899
ide_tpNF = 1                    ' número da nota fiscal de saída
ide_cMunFG = 3550308            ' código do município do IBGE de ocorrência do FG do ICMS (vide página 171 do manual)
ide_tpImp = 1                   ' orientação da impressão 1-retrato/2-paisagem
ide_tpAmb = 2                   ' ambiente de envio da NF-e 1-produção / 2 - homologação
ide_finNFe = 1                  ' finalidade da emissão da NF-e 1- NF-e normal
ide_tpEmis = 1                  ' forma de emissão da NF-e 1- normal, 2 - contingência FS, 3 - contingência SCAN, etc.
ide_procEmi = 0                 ' identificação do processo de emissão da NF-e 0 - aplicação do contribuinte
ide_verProc = "NFe_Util_v1.4"   ' identificação da vesão do processo de emissão
ide_NFref = ""                  ' NF referenciada, deve ser informado para nota fiscal complementar, devolução, etc. - http://www.flexdocs.com.br/guiaNFe/gerarNFe.ref.html

'
' novos campos do versão 2G
'
ide_hSaiEnt = ""                ' hora da saída
ide_dhCont = #12:00:00 AM#      ' data e hora de entrada em contingência - informar quanto tpEmis diferente de 1, informe #12:00:00 AM# para deixar vazio em VB
ide_xJust = ""                  ' informar a justificativa de entrada em contingência, deve ser informado sempre que tpEmis for diferente de 1.
'
'     gera a chave de acesso da NF-e
'
'     utilizar a função criaChaveNFe para gerar a chave de acesso, código da NF-e e DV
'
'=========variáveis de trabalho
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
'  parâmetro novo da versão 2.00
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

ide_cNF = Val(cNF)                     ' código numérico que compõe a chave de acesso, o tamanho do campo foi reduzido para 8 dígitos na versão 2.00
ide_cDV = Val(cDV)                     ' DV da chave de acesso da NF-e
'
'   gera grupo ide
'
ide = objNFeUtil.identificador2G(ide_cUF, ide_cNF, ide_natOp, ide_indPag, ide_mode, ide_serie, ide_nNF, ide_dEmi, ide_dSaiEnt, ide_hSaiEnt, ide_tpNF, ide_cMunFG, ide_NFref, ide_tpImp, ide_tpEmis, ide_cDV, ide_tpAmb, ide_finNFe, ide_procEmi, ide_verProc, ide_dhCont, ide_xJust)
'
'MsgBox "Grupo de identificação " + ide, vbInformation, "Resultado"
'
'
'================grupo de identificação do destinatario (grupo E do Manual de integração - páginas 116-)=======================
'
'        <>&" são caracteres reservados do XML e devem ser evitados ou substituídos
'        por &lt; &gy; &amp; &quot;
'
'        Vale ressaltar que algumas aplicações das UF devem mostrar DIAS &amp; DIAS TENTANDO S/A,
'        pois não entedem &amp; como &, assim talvez seja melhor substituir o & por e.
'
dest_CNPJ = "00000000000191"                 ' CNPJ do destinatario sem máscara de formatação
dest_CPF = ""                                ' CPF do destinatario, uso exclusivo do Fisco
dest_xNome = "Banco do Brasil S/A"           ' Razão social do destinatario, evitar caracteres acentuados e &
dest_xFant = "BB"                            ' Nome fantasia
dest_xLgr = "Rua Libero Badaro"              ' logradouro
dest_nro = "280"                             ' número, informar S/N quano inexistente para erro de Schema XML
dest_xCpl = "10 andar"                       ' complemento do endereço, o conteúdo pode ser omitido
dest_xBairro = "CENTRO"                      ' bairro
dest_cMun = "3550308"                        ' código do município (vide página 171 do manual), deve ser compatível com a UF
dest_xMun = "SAO PAULO"                      ' nome do município
dest_UF = "SP"                               ' sigla da UF
dest_CEP = "01315000"                        ' CEP - sem máscara
dest_cPais = "1058"                          ' código do pais - deve fixo em 1058 - Brasil
dest_xPais = "Brasil"                        ' nome do pais (Brasil ou BRASIL)
dest_fone = "1133221234"                     ' número do telefone sem máscara, o tamanho do campo foi aumentado para 16 dígitos na versão 2.00
dest_IE = "123456789011"                     ' Inscrição Estadual do destinatario sem máscara
dest_IESUF = ""                              ' Inscrição SUFRAMA
'
' novos campos do versão 2.00
'
dest_email = "destinatario@empresa.com.br"   ' Infomrmar o e-mail do destinatário
'
'   gera grupo do destinatário - vide: http://www.flexdocs.com.br/guiaNFe/gerarNFe.des.destinatario2G.html
'
dest = objNFeUtil.destinatario2G(dest_CNPJ, dest_CPF, dest_xNome, dest_xLgr, dest_nro, dest_xCpl, dest_xBairro, dest_cMun, dest_xMun, dest_UF, dest_CEP, dest_cPais, dest_xPais, dest_fone, dest_IE, dest_IESUF, dest_email)
'MsgBox "Grupo de destinatário " + dest, vbInformation, "Resultado"
'
'           INICIALIZAÇÃO
'
'       A DLL não acumula os valores do itens
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
'================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 120-)=======================
'
'                http://www.flexdocs.com.br/guiaNFe/gerarNFe.detalhe.pro.produto2G.html
'
Prod_cProd = "001152"                       ' código do produto
Prod_cEAN = "7897844200115"                 ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se o produto não tiver EAN
Prod_xProd = "Cola Especial para EPS"       ' código do produto, espaços em branco consecutivos ou no início ou fim do campo podem gerar erro de Schema XML, além de caracteres reservados do XML <>&"'
'
'   campo com nova regra de preenchimento
'
Prod_NCM = "35"                             ' código NCM, informar o Código NCM com 8 dígitos - http://www.mdic.gov.br/sitio/interna/interna.php?area=5&menu=1095#I;
                                            ' informar a posição do capítulo do NCM (as duas primeiras posições do NCM) quando a operação não for de comércio exterior (importação/ exportação) ou o produto não seja tributado pelo IPI;
                                            ' se for serviços, informar 00
Prod_ExTIPI = ""                            ' ExTipi, especialização do código NCM, informar apenas se existir e o NCM completo for informado
'
'  campo excluído da NF-e
'
'Prod_genero = 0                            ' informar as duas primeiras posições do NCM
Prod_CFOP = "5102"                          ' CFOP do operação, causa erro de XML se informado um código inexistente
Prod_uCOM = "UN"                            ' unidade de comercialização
Prod_qCom = "10"                            ' quantidade de comercialização
Prod_vUnCom = "1"                           ' valor unitário de comercialização, campo de mera demonstração deve ser o resultado da divisão do vProd / qCom
Prod_vProd = 10                             ' valor do total do item
Prod_cEANTrib = "7897844200115"             ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN, em geral é o mesmo código do EAN de comercialização
Prod_uTrib = "UN"                           ' unidade de tributação, na maioria dos casos é idêntico  ao vUnCom, pode diferente nos casos de produtos sujeitos a ST em que a unidade de pauta é diferente da unidade de comercialização
                                            ' Ex. unidade de comercialização = 1 pack de lata de cerveja => unidade de tributação = 1 lata (preço de pauta)
Prod_qTrib = "10"                           ' quantidade de comercialização
Prod_vUnTrib = "1"                          ' valor unitário de tributação, campo de mera demonstração deve ser o resultado da divisão do vProd / qTrib
Prod_vFrete = 0                             ' valor do frete, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vSeguro = 0                            ' valor do seguro, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vDesc = 0                              ' valor do desconto concedido
Prod_DI = ""                                ' dados da importação, informar apenas no caso de NF de entrada (importação)
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.di.html
Prod_DetEspecifico = ""                     ' dados específicos, informar para medicamento, veículos novos, armamentos e combustíveis.
                                            ' veicProd - http://www.flexdocs.com.br/guiaNFe/gerarNFe.vei.html
                                            ' med - http://www.flexdocs.com.br/guiaNFe/gerarNFe.med.html
                                            ' arma - http://www.flexdocs.com.br/guiaNFe/gerarNFe.arm.html
                                            ' comb - http://www.flexdocs.com.br/guiaNFe/gerarNFe.com.html
                                            '
Prod_infAdProd = ""                         ' informações adicionais do produto
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.detalhe.html
Prod_indTot = 1                             ' indicador de totalização do valor do produto

Prod_xPed = ""                              ' número do pedido de compra
Prod_nItemPed = 0                           ' número do item do pedido
'
'   campo novo da versão 2.00 do leiaute
'
Prod_vOutro = 0
'
'   gera grupo do destinatário
'
Prod = objNFeUtil.produto2G(Prod_cProd, Prod_cEAN, Prod_xProd, Prod_NCM, Prod_ExTIPI, Prod_CFOP, Prod_uCOM, Prod_qCom, Prod_vUnCom, Prod_vProd, Prod_cEANTrib, Prod_uTrib, Prod_qTrib, Prod_vUnTrib, Prod_vFrete, Prod_vSeguro, Prod_vDesc, Prod_vOutro, Prod_indTot, Prod_DI, Prod_DetEspecifico, Prod_xPed, Prod_nItemPed)

'MsgBox "Grupo de produto " + prod, vbInformation, "Resultado"
'
'
'=========dados do ICMS (grupo N01 do Manual de integração - páginas 128-)=====================
'
icms_orig = "0"                             ' Tabela A - origem da mercadoria 0=nacional
icms_CST = "00"                             ' Tabela B - CST=00-tributação normal
icms_modBC = 3                              ' modalidade de determinação da BC = 3-valor da operação
icms_pRedBC = 0                             ' percentual de redução da BC
icms_vBC = 10                               ' valor da BC do ICMS = vProd + vFrete + vSeguro + vOutro
icms_pICMS = 18                             ' alíquota do ICMS
icms_vICMS = 1.8                            ' valor do ICMS
icms_modBCST = 0                            ' modalidade de determinação da BC ICMS ST
icms_pmVAST = 0                             ' percentual de valor de margem e valor adicionado
icms_pRedBCST = 0                           ' percentual de redução da BC do ICMS ST
icms_vBCST = 0                              ' BC do ICMS ST
icms_pICMSST = 0                            ' percentual do ICMSST
icms_vICMSST = 0                            ' valor do ICMS ST devido
'
'   Campos novos da versão 2.00
'
icms_vBCSTRet = 0                           ' informação do ICMS retindo anteriormente por ST
icms_vICMSSTRet = 0                         ' estes campos devem ser informado somente no caso do CST = 60 ou CSOSN = 500
'
icms_motDesICMS = 0                         ' motivo de desoneração do ICMS, só deve ser informado no caso de CST = 40 (isenção condicional)
'
icms_pBCOp = 0                              ' campos para uso nos casos de ICMSPart/ICMSST
icms_UFST = ""                              '
icms_vICMSSTDest = 0                        '
icms_vBCICMSSTDest = 0                      '
'
icms_pCredSN = 0                            ' campos exclusivos para emissor optante do Simples Nacional CSOSN= 101, 201 e 900
icms_vCredICMSSN = 0                        ' não esquecer de informar o CRT=1

'
'   gera grupo do ICMS
'

icms = objNFeUtil.icms2G(icms_orig, icms_CST, icms_modBC, icms_pRedBC, icms_vBC, icms_pICMS, icms_vICMS, icms_modBCST, icms_pmVAST, icms_pRedBCST, icms_vBCST, icms_pICMSST, icms_vICMSST, icms_vBCSTRet, icms_vICMSSTRet, icms_vBCICMSSTDest, icms_vICMSSTDest, icms_motDesICMS, icms_pBCOp, icms_UFST, icms_pCredSN, icms_vCredICMSSN)

'MsgBox "Grupo de Tributos/ICMS " + icms, vbInformation, "Resultado"

'
'=========dados do PIS (grupo Q do Manual de Integração - páginas 145) =============
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
'========dados do COFINS (grupo s do Manual de Integração - páginas 147) ============
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
'========dados do IMPOSTO (grupo M do Manual de Integração - páginas 128) ============
'
Imposto = objNFeUtil.imposto2G(icms, "", "", pis, "", cofins, "", "")
'
'   atualização de total
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
'========dados do ITEM do detalhe (grupo H do Manual de Integração - páginas 120-) ============
'
'  item 1
'
Detalhes = objNFeUtil.detalhe(1, Prod, Imposto, Prod_infAdProd)
'MsgBox "Grupo de detalhe do Item " + det, vbInformation, "Resultado"
'
'
'================grupo de detalhe do produto (grupo I01 do Manual de integração - páginas 120-)=======================
'
'                   exemplo do segundo item ST
'
Prod_cProd = "002871"                       ' código do produto
Prod_cEAN = "7896045512321"                 ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN
Prod_xProd = "Cerveja da boa"       ' código do produto, espaços em branco consecutivos ou no início ou fim do campo podem gerar erro de Schema XML, além de caracteres reservados do XML <>&"'
'
'   campo com nova regra de preenchimento
'
Prod_NCM = "22"                             ' código NCM, informar o Código NCM com 8 dígitos - http://www.mdic.gov.br/sitio/interna/interna.php?area=5&menu=1095#I;
                                            ' informar a posição do capítulo do NCM (as duas primeiras posições do NCM) quando a operação não for de comércio exterior (importação/ exportação) ou o produto não seja tributado pelo IPI;
                                            ' se for serviços, informar 00
Prod_ExTIPI = ""                            ' ExTipi, especialização do código NCM, informar apenas se existir e o NCM completo for informado
'
'  campo excluído da NF-e
'
'Prod_genero = 0                            ' informar as duas primeiras posições do NCMProd_CFOP = "5403"                          ' CFOP do operação, causa erro de XML se informado um código inexistente
Prod_uCOM = "PAC12"                            ' unidade de comercialização
Prod_qCom = 10                              ' quantidade de comercialização
Prod_vUnCom = 10                            ' valor unitário de comercialização, campo de mera demonstração deve ser o resultado da divisão do vProd / qCom
Prod_vProd = 100                            ' valor do total do item
Prod_cEANTrib = "7896045512317"             ' código EAN (0, 8,12, 13 ou 14 caracteres), o conteúdo pode ser omitido se não tiver EAN, em geral é o mesmo código do EAN de comercialização
Prod_uTrib = "LATA"                         ' unidade de tributação, na maioria dos casos é idêntico  ao vUnCom, pode diferente nos casos de produtos sujeitos a ST em que a unidade de pauta é diferente da unidade de comercialização
                                            ' Ex. unidade de comercialização = 1 pack de lata de cerveja => unidade de tributação = 1 lata (preço de pauta)
Prod_qTrib = 120                            ' quantidade de comercialização
Prod_vUnTrib = 0.8333                       ' valor unitário de tributação, campo de mera demonstração deve ser o resultado da divisão do vProd / qTrib
Prod_vFrete = 0                             ' valor do frete, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vSeguro = 0                            ' valor do seguro, se cobrado do cliente deve ser rateado entre os itens de produto
Prod_vDesc = 0                              ' valor do desconto concedido
Prod_DI = ""                                ' dados da importação, informar apenas no caso de NF de entrada (importação)
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.di.html
Prod_DetEspecifico = ""                     ' dados específicos, informar para medicamento, veículos novos, armamentos e combustíveis.
                                            ' veicProd - http://www.flexdocs.com.br/guiaNFe/gerarNFe.vei.html
                                            ' med - http://www.flexdocs.com.br/guiaNFe/gerarNFe.med.html
                                            ' arma - http://www.flexdocs.com.br/guiaNFe/gerarNFe.arm.html
                                            ' comb - http://www.flexdocs.com.br/guiaNFe/gerarNFe.com.html
                                            '
Prod_infAdProd = ""                         ' informações adicionais do produto
                                            ' http://www.flexdocs.com.br/guiaNFe/gerarNFe.detalhe.html
'
'   campo novo da versão 2.00 do leiaute
'
Prod_vOutro = 0
Prod_xPed = ""                              ' número do pedido de compra
Prod_nItemPed = 0                           ' número do item do pedido
'
'   gera grupo do destinatário
'
Prod = objNFeUtil.produto2G(Prod_cProd, Prod_cEAN, Prod_xProd, Prod_NCM, Prod_ExTIPI, Prod_CFOP, Prod_uCOM, Prod_qCom, Prod_vUnCom, Prod_vProd, Prod_cEANTrib, Prod_uTrib, Prod_qTrib, Prod_vUnTrib, Prod_vFrete, Prod_vSeguro, Prod_vDesc, Prod_vOutro, Prod_indTot, Prod_DI, Prod_DetEspecifico, Prod_xPed, Prod_nItemPed)

'MsgBox "Grupo de produto " + prod, vbInformation, "Resultado"
'
'
'=========dados do ICMS (grupo N01 do Manual de integração - páginas 128)=====================
'
icms_orig = "0"                             ' Tabela A - origem da mercadoria 0=nacional
icms_CST = "10"                             ' Tabela B - CST=10-tRIBUTADA E COM COBRANCA POR ST
icms_modBC = 3                              ' modalidade de determinação da BC = 3-valor da operação
icms_pRedBC = 0                             ' percentual de redução da BC
icms_vBC = 100                              ' valor da BC do ICMS = vProd + vFrete + vSeguro + vOutro
icms_pICMS = 18                             ' alíquota do ICMS
icms_vICMS = 18                             ' valor do ICMS
icms_modBCST = 5                            ' modalidade de determinação da BC ICMS ST
icms_pmVAST = 0                             ' percentual de valor de margem e valor adicionado
icms_pRedBCST = 0                           ' percentual de redução da BC do ICMS ST
icms_vBCST = 180                            ' BC do ICMS ST
icms_pICMSST = 18                           ' percentual do ICMSST
icms_vICMSST = 14.4                         ' valor do ICMS ST devido
'
'   Campos novos da versão 2.00
'
icms_vBCSTRet = 0                           ' informação do ICMS retindo anteriormente por ST
icms_vICMSSTRet = 0                         ' estes campos devem ser informado somente no caso do CST = 60 ou CSOSN = 500
'
icms_motDesICMS = 0                         ' motivo de desoneração do ICMS, só deve ser informado no caso de CST = 40 (isenção condicional)
'
icms_pBCOp = 0                              ' campos para uso nos casos de ICMSPart/ICMSST
icms_UFST = ""                              '
icms_vICMSSTDest = 0                        '
icms_vBCICMSSTDest = 0                      '
'
icms_pCredSN = 0                            ' campos exclusivos para emissor optante do Simples Nacional CSOSN= 101, 201 e 900
icms_vCredICMSSN = 0                        ' não esquecer de informar o CRT=1

'
'   gera grupo do ICMS
'
 icms = objNFeUtil.icms2G(icms_orig, icms_CST, icms_modBC, icms_pRedBC, icms_vBC, icms_pICMS, icms_vICMS, icms_modBCST, icms_pmVAST, icms_pRedBCST, icms_vBCST, icms_pICMSST, icms_vICMSST, icms_vBCSTRet, icms_vICMSSTRet, icms_vBCICMSSTDest, icms_vICMSSTDest, icms_motDesICMS, icms_pBCOp, icms_UFST, icms_pCredSN, icms_vCredICMSSN)
 
'MsgBox "Grupo de Tributos/ICMS " + icms, vbInformation, "Resultado"

'
'=========dados do PIS (grupo Q do Manual de Integração - páginas 145) =============
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
'========dados do COFINS (grupo s do Manual de Integração - páginas 147) ============
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
'========dados do IMPOSTO (grupo M do Manual de Integração - páginas 128) ============
'
Imposto = objNFeUtil.imposto2G(icms, "", "", pis, "", cofins, "", "")
'MsgBox "Grupo de Tributos/COFINS " + COFINS, vbInformation, "Resultado"
'
'   atualização de total
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
'========dados do ITEM do detalhe (grupo H do Manual de Integração - páginas 120) ============
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
totICMS_vNF = totICMS_vProd + totICMS_vFrete + totICMS_vSeg - totICMS_vDesc + totICMS_vST  ' verificar outros acréscimos e descontos
'
'
'
totICMS = objNFeUtil.totalICMS(totICMS_vBC, totICMS_vICMS, totICMS_vBCST, totICMS_vST, totICMS_vProd, totICMS_vFrete, totICMS_vSeg, totICMS_vDesc, totICMS_vII, totICMS_vIPI, totICMS_vPIS, totICMS_vCOFINS, totICMS_vOutro, totICMS_vNF)

total = objNFeUtil.total(totICMS, "", "")       ' total da NF-e sem os valors de ISSQN e RetTributos

'MsgBox "Grupo de total " + total, vbInformation, "Resultado"

'
'============dados do transportador
'
transpModFrete = "0"        ' responsabilidade do frete 0-emitente, 1-destinatário
transp = objNFeUtil.transportador(transpModFrete, "", "", "", "", "")

'
'==============informações adcionais
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

MsgBox "Examine o código fonte da aplicação para compreender a lógica de uso das funcionalidades de criação do XML da NF-e ", vbInformation, "Informação"


Set objNFeUtil = Nothing
'
End Sub

Private Sub Command10_Click()
'
'  Cancelamento da NF-e
'
'  Esta funcionaliade deve ser utilizada para cancelar
'  uma NF-e autorizada e ainda não tenha ocorrido o fato
'  gerador (circulação da mercadoria).
'  Ex. falta de mercadoria, divergência de quantidade, preço, etc.
'  desistência do comprador, etc.
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
'  As variáveis do proxy devem ser informadas se necessário
'
'  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
Dim licenca As String
'
Dim ambiente As Integer
'
' define as variáveis que passam/recebem informações importantes
'
Dim ChaveNFe As String          ' chave da NF-e objeto de cancelamento
Dim ProtAutNFe As String        ' protocolo de autorização de uso
Dim Justificativa As String     ' justificativa de cancelamento
'
'  parâmetros novos
'
Dim procCancNFe As String       ' estrturura XML que contém o pedido de cancelamento e a homologação do cancelamento,
                                ' que deve ser mantido pelo emissor e distribuído ao destinatário.
Dim nProtocoloCanc As String    ' número do protocolo de homomologação de cancelamento devolvido pela SEFA
Dim dProtocoloCanc As String    ' data e hora de homologação do cancelamento
Dim versao As String            'utilizado para escolha da versão do WS


'
'
'  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
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
        MsgBox "Necessário informar a chave de acesso da NF-e para cancelamento da NF-e.", vbCritical, "Atenção:"
            Exit Sub
End If

ProtAutNFe = InputBox("Informe o número do protocolo da autorização de uso da NF-e objeto de cancelamento", "Cancelamento de NF-e")


If ProtAutNFe = "" Then '
        MsgBox "Necessário informar o número do protocolo da autorização de uso da NF-e para cancelamento da NF-e.", vbCritical, "Atenção:"
            Exit Sub
End If

Justificativa = InputBox("Informe a Justificativa de cancelamento", "Cancelamento de NF-e")


If Len(Justificativa) < 15 Then '
        MsgBox "Necessário informar a justificativa com no mínimo 15 caracteres", vbCritical, "Atenção:"
            Exit Sub
End If

'
' parâmetro novo - utilizado para escolha da versão do WS
'
versao = cbVersao.Text ' versão do Web Service, a versão anterior é 1.07, a versão nova é 2.00

If cbAmb.Text = "Produção" Then
   ambiente = 1
Else
   ambiente = 2
End If

Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.canc.cancelaNF2G.html

'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
                                      
   MsgBox msgResultado & Chr(13) & Chr(13) + "Protocolo de homologação de cancelamento: " + nProtocoloCanc + Chr(13) & Chr(13) + "Data e hora de homologação de cancelamento: " + dProtocoloCanc + Chr(13) & Chr(13) + "Grave o procCancNFe : " + procCancNFe, vbInformation, "Atenção: Cancelamento da NF-e"
   
   '
   ' grave o procCancNFe, pois o XML deve ser mantido pelo emissor, além de ser distribuído para o destinatário também.
   '

Else

    MsgBox msgResultado & Chr(13) & Chr(13), vbError, "Atenção: Cancelamento da NF-e Falhou"

End If

End Sub

Private Sub Command11_Click()

'
'  Inutiliza Número de NF-e
'
'  A funcionalidade deve ser utilizada para inutilizar um
'  número de NF-e que não vai ser utilizada (atribuída) a
'  NF-e, por salto de numeração, rejeição de NF-e, etc.
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
'  As variáveis do proxy devem ser informadas se necessário
'
'  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
Dim licenca As String
'
Dim ambiente As Integer
'
' define as variáveis que passam/recebem informações importantes
'
Dim cUF As String               ' código da UF do solicitante - Tabela IBGE
Dim ano As String               ' ano de inutilizalção da numeração
Dim CNPJ As String              ' CNPJ do emitente
Dim modelo As String            ' modelo da NF-e (sempre 55)
Dim serie As String             ' serie da NF-e (sem zeros a esquerda)
Dim nInicial As String          ' número inicial da faixa a ser inutilizada (sem zeros a esquerda)
Dim nFinal As String            ' número final da faixa a ser inutilizada (sem zeros a esquerda)
                                ' Observações
                                ' só é permitida a inutilização de até 1000 números por vez
                                ' se a inutilização for de um único número nInicial e nFinal devem
                                ' ser iguais
Dim Justificativa As String     ' justificativa de cancelamento
'
'  parâmetros novos
'
Dim procInutNFe As String       ' estrturura XML que contém o pedido de inutilização e a homologação da inutilização,
                                ' que deve ser mantido pelo emissor.
Dim nProtocoloInut As String    ' número do protocolo de homomologação de Inutilização de numeraçãp devolvido pela SEFA
Dim dProtocoloInut As String    ' data e hora de homologação da Inutilização de numeraçãp
Dim versao As String            'utilizado para escolha da versão do WS
'
'
'  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
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
'  Solicita os parâmetros da inutilização
'

' o modelo da NF-e é sempre fixo em 55
modelo = "55"
'
' converte a sigla da UF no código da UF da tabela do IBGE
'
cUF = Mid$("11 12 13 14 15 16 17 21 22 23 24 25 26 27 28 29 31 32 33 35 41 42 43 50 51 52 53", InStr(1, "RO AC AM RO PA AP TO MA PI CE RN PB PE AL SE BA MG ES RJ SP PR SC RS MS MT GO DF", siglaUF, 1), 2)
'
'
CNPJ = InputBox("Informe o CNPJ do emissor", "Inutilização de NF-e")
If CNPJ = "" Then '
        MsgBox "Necessário informar o CNPJ do emissor", vbCritical, "Atenção:"
            Exit Sub
End If

ano = InputBox("Informe o Ano (AA) da numeração que será inutilizado", "Inutilização de NF-e")
If ano = "" Then '
        MsgBox "Necessário informar o ano de inutilização da numeração.", vbCritical, "Atenção:"
            Exit Sub
End If

serie = InputBox("Informe a série da numeração que será inutilizado", "Inutilização de NF-e")
If serie = "" Then ' o certo é verificar se é um número da faixa 0-999, sem zeros a esquerda
        MsgBox "Necessário informar a série de inutilização da numeração.", vbCritical, "Atenção:"
            Exit Sub
End If

nInicial = InputBox("Informe o número inicial da numeração que será inutilizado", "Inutilização de NF-e")
If nInicial = "" Then ' o certo é verificar se é um número da faixa 1-9999999999, sem zeros a esquerda
        MsgBox "Necessário informar o número inicial de inutilização da numeração.", vbCritical, "Atenção:"
            Exit Sub
End If

nFinal = InputBox("Informe o número final da numeração que será inutilizado", "Inutilização de NF-e")
If nFinal = "" Then ' o certo é verificar se é um número da faixa 1-9999999999, sem zeros a esquerda
        MsgBox "Necessário informar o número inicial de inutilização da numeração.", vbCritical, "Atenção:"
            Exit Sub
End If

Justificativa = InputBox("Informe a Justificativa da inutilização com pelo menos 15 caracteres", "Inutilização de NF-e")
If Len(Justificativa) < 15 Then '
        MsgBox "Necessário informar a justificativa com no mínimo 15 caracteres", vbCritical, "Atenção:"
            Exit Sub
End If

versao = cbVersao.Text ' versão do Web Service, a versão anterior é 1.07, a versão nova é 2.00

If cbAmb.Text = "Produção" Then
   ambiente = 1
Else
   ambiente = 2
End If

Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.canc.inutilizaNro2G.html


'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
                                      
   MsgBox msgResultado & Chr(13) & Chr(13) + "Protocolo de homologação de Inutilização de Numeração: " + nProtocoloInut + Chr(13) & Chr(13) + "Data e hora de homologação de Inutilização de Numeração: " + dProtocoloInut + Chr(13) & Chr(13) + "Grave o procInutNFe : " + procInutNFe, vbInformation, "Atenção: Inutilização de Numeração da NF-e"
   
   '
   ' grave o procInutNFe, pois o XML deve ser mantido pelo emissor.
   '

Else

    MsgBox msgResultado & Chr(13) & Chr(13), vbError, "Atenção: Inutilização de numeração da NF-e Falhou"

End If


End Sub

Private Sub Command12_Click()

'
' Converte TXT em XML
'
' A conversão é limitada a uma NF-e por vez.
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

CommonDialog1.DialogTitle = "Escolha o arquivo TXT (versão selecionada: " + cbVersao.Text + ")"
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo para validação de Schema XML." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0

Open CommonDialog1.FileName For Input As #1
txt = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = txt

geraChaveNFe = 1               ' gerar chave NF-e (opção fixa neste exemplo)

qtdeErros = 0                   ' quantidade de erros
codigoSeguranca = "E segredo"   ' código que será utilizado para gerar a chave, deve ser a mesma cliente
txtNumerado = ""                ' campo para receber o txt numerado, útil para localizar o erro
erros = ""                      ' relatório de erros encontrados
msgResultado = ""               ' literal da mensagem de resultado da chamada da função

Dim Resultado As Long

'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
 
       txtRetorno.Text = XML        ' necessário gravar o XML gerado para utilizar posteriormente.

       MsgBox msgResultado + vbCrLf + "Grave o retorno em um arquivo para utiliza-lo.", vbInformation, "Informação"

Else

       txtRetorno.Text = "Quantidade de Erros: " + Str(qtdeErros) & vbCrLf & erros
       MsgBox "Processo de validação do XML falhou..." & vbCrLf & msgResultado, vbExclamation, "Atenção"

End If


End Sub

Private Sub Command13_Click()

'
' Monta procNFe
'
' A montagem é limitada a uma NF-e por vez.
' Entradas:
'   NFeAssinada: NF-e em formato XML, deve estar assinada
'   nomeCertificado: Nome do titular do certificado a ser utlizado na conexão SSL
'    Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"
'   proxy ,usuario e senha: deve ser informado nos casos em que é necessário o uso de proxy
'   https://proxyserver:port'; // verificar com o cliente qual é o endereço do servidor proxy e a porta https, a porta padrão do https é 443, assim teríamos algo do
'   tipo 'http://192.168.15.1:443'
'
' Retorno:
'

' msgResultado: literal da mensagem de resultado da chamada da função
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo para validação de Schema XML." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0

Open CommonDialog1.FileName For Input As #1
NFeAssinada = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = NFeAssinada

nomeCertificado = txtCertificado.Text

siglaWS = cbWS.Text             ' necessário apenas para a versão 2.00
retCancNFe = ""                 ' necessário apenas para a versão 2.00

proxy = ""                      ' preencher estes campos somente em caso de existência de proxy na rede
usuario = ""
senha = ""

protocolo = ""                  ' número do protocolo + dd/mm/aa HH:MM:SS

msgResultado = ""               ' literal da mensagem de resultado da chamada da função

Resultado = 0

'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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

       MsgBox msgResultado, vbInformation, "Informação"

   Else

       txtRetorno.Text = ""
       MsgBox "Processo de montagem procNFe falhou..." & vbCrLf & msgResultado, vbExclamation, "Atenção"

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

         MsgBox msgResultado, vbInformation, "Informação"

      Else

         txtRetorno.Text = ""
         MsgBox "Processo de montagem procNFe falhou..." & vbCrLf & msgResultado, vbExclamation, "Atenção"

      End If
    Else

         MsgBox "Versão da NF-e selecionada inválida, diferente de 1.10 e 2.00..." & vbCrLf & msgResultado, vbError, "Atenção"
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
' detalhes da funcionalidade disponível em: http://www.flexdocs.com.br/guiaNFe/funcao.assinatura.assinar.html
'

Dim XMLString As String
Dim RefUri As String
Dim nomeCertificado As String
Dim XMLAssinado As String
Dim msgResultado As String
'
'
'  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   
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
' não é necessário fazer o reference da DLL
' o intelisense não funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

Screen.MousePointer = vbHourglass    ' ponteiro ampulheta
'
'  houve alteração nos parâmetros de retorno, agora o XMLAssinado é devolvido pela funcionalidade
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
MsgBox msgResultado, vbInformation, "Informação"

Else

MsgBox "Processo de assinatura falhou..." & vbCrLf & msgResultado, vbExclamation, "Atenção"

End If
'
'  liberar DLL
'
Set objNFeUtil = Nothing
End Sub

Private Sub Command3_Click()
'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
MsgBox "Versão em uso: " & objNFeUtil.versao, vbInformation, "Informação"

'
'  liberar DLL
'
Set objNFeUtil = Nothing


End Sub

Private Sub Command4_Click()
'
' Exemplo para escolher um certificado digital do repositório de certificados digitais do usuário corrente do
' Windows
'
' Importante ressaltar que não é necessário executar esta funcioanlidade antes das chamadas da DLL, ofereça esta funcionalidade apenas
' para a escolha do certificado digital que será utilizada na configuração da aplicação.
'
' Também vale observrar que existe uma funcionaliade que retorna a data de fim da validado do certificado digital que é mais interessante de ser utilizada
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
' não é necessário fazer o reference da DLL
' o intelisense não funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
' pega certificado
'
' o texto que retorna no campo Certificado será utilizada para identificar
' o certificado digital em uso para as demais chamadas que necessitam de
' um certificado digital
'
Resultado = objNFeUtil.PegaNomeCertificado(certificado, msgResultado)
If Resultado < 5403 Then
   If InStr(1, certificado, "Associacao", vbTextCompare) > 0 Then
      MsgBox "O certificado digital da Associação não é um certificado válido para consumir os WS da NF-e! Procure adquirir um certificado digital válido para prosseguir com os testes...", vbInformation, "Resultado"
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
' ConsultaStatus2G: Consulta Situação do Web Service de Recepção de NF-e
'
'
'
' declaração das variáveis que serão utilizadas na passagem de parâmetros da DLL
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaUF As String
Dim siglaWS As String
Dim certificado As String
'
' As variáveis do proxy devem ser informadas se necessário
'
' proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
'
Dim ambiente As Integer
'
' parâmetro novo - utilizado para escolha da versão do WS
'
Dim versao As String
'
' IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
'
'
proxy = ""
usuario = ""
senha = ""
msgDados = ""
msgRetWS = ""
'
' prepara variáveis
'
certificado = txtCertificado.Text
siglaWS = cbWS.Text
siglaUF = cbUF.Text
'
' parâmetro novo - utilizado para escolha da versão do WS
'
versao = cbVersao.Text ' versão do Web Service, a versão anterior é 1.07, a versão nova é 2.00

If cbAmb.Text = "Produção" Then
   ambiente = 1
Else
   ambiente = 2
End If
   
txtEntrada.Text = ""
txtRetorno.Text = ""

Dim cStat As Long
'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
MsgBox msgResultado + Chr(13) + Chr(13) + msgRetWS, vbInformation, "Resultado da Consulta Status do Serviço"
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
' ********IMPORTANTE O tipoXML da versão 2G é 19 ***************
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo para validação de Schema XML." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0

Open CommonDialog1.FileName For Input As #1
XML = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = XML

tipoXML = 19    ' validar NF-e (opção fixa para validar NF-e da versão 2.00 neste exemplo, se for validar a versão 1.10 da NF-e informe 1)
qtdeErros = 0   ' quantidade de erros, se o XML não estiver assinado vai ocorrer um erro
erroXML = ""
msgResultado = ""

Dim Resultado As Long

'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
MsgBox msgResultado, vbInformation, "Informação"

ElseIf (Resultado = 5506) Then

txtRetorno.Text = ""
MsgBox "XML da NF-e sem assinatura...", vbInformation, "Informação"

Else

txtRetorno.Text = "Quantidade de Erros: " + Str(qtdeErros) & vbCrLf & erroXML
MsgBox "Processo de validação do XML falhou..." & vbCrLf & msgResultado, vbExclamation, "Atenção"

End If
'
'  liberar DLL
'
Set objNFeUtil = Nothing



End Sub

Private Sub Command7_Click()

' EnviaNFe2G: Envio de uma única NF-e
'
' para mais detalhes da funcionalidade acesse: http://www.flexdocs.com.br/guiaNFe/WS.NFe.enviaNFe2G.html
'
'
' declaração das variáveis que serão utilizadas na passagem de parâmetros da DLL
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaWS As String
Dim certificado As String

'
' As variáveis do proxy devem ser informadas se necessário
'
' proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
'
Dim proxy   As String
Dim usuario As String
Dim senha As String
'
' licenca - deve ser informado com a chave da licença de uso para acessar os WS de produção
'
Dim licenca As String
'
' NFe         - informar com a NF-e a ser transmitida, não é necessário validar nem assinar
'
' NFeAssinada - é devolvido pela DLL se a chamada for realizada com sucesso.
'
' NroRecibo   - é devolvido pela DLL se a NF-e for transmitida corretamente,
'               este número é necessário para buscar o resultado do processamento da NF-e
'               ========== o NroRecibo não indica que a NF-e foi autorizada ==============
Dim NFe As String
Dim NFeAssinada As String
Dim nroRecibo As String
Dim cStat As Long
Dim versao As String
Dim dhRecibo As String
Dim tMed As String
Dim x As Boolean
'
' IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
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
' prepara variáveis
'
certificado = txtCertificado.Text
siglaWS = cbWS.Text
versao = cbVersao.Text

'
' nesta chamada da DLL pega o ambiente que está informado na NF-e
' assim tomar cuidado para não enviar a NF-e para o ambiente de produção
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo XML da NF-e para transmissão." & vbCrLf & _
            Err.Description
        Exit Sub
End If
On Error GoTo 0
'
'  carrega o conteúdo do arquivo xml em NFe
'
Open CommonDialog1.FileName For Input As #1
NFe = Input$(LOF(1), 1)
Close #1

txtEntrada.Text = NFe

nomeArquivo = Replace(CommonDialog1.FileName, ".", "_assinado.")

'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
  '  se cStat = 105 - Lote em processamento, significa que a SEFAZ ainda não conseguiu processar a NF-e e a aplicação deve persistir até que o cStat seja diferent de 105
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
                                      
txtNroRecibo.Text = nroRecibo       ' número do recibo do lote entregue
                                    ' este número é necessário para fazer a consulta
                                    ' do resultado do processamento do lote enviado
                                    
' análise do retorno da chamada da função enviaNFeSCAN
'
' resultado:
'
' 5000-7000 ------> falha na chamda da função, vide a mensagem de erro e corrija o problema - http://www.flexdocs.com.br/guiaNFe/WS.NFe.enviaNFe2G.html
'
' 103 -------> SUCESSO! - LOTE RECEBIDO pelo WS, guarde o NroRecibo e efetua a busca do resultado do processamento
'
'              IMPORTANTE, ter o número do recibo não significa que a NF-e está autorizada, ainda é necessário buscar o resultado do processamento
'
' 108/109 ---> WS com problemas, sem condições de receber lotes
' 2xx -------> FALHA - existe algum problema com o lote enviado, verifique o código do erro e corrigir o erro
'
Select Case cStat

       Case 5000 To 7003                ' problemas no envio da DLL para o WS
        
        If (MsgBox("Descrição do erro: " & Chr(13) & Chr(13) + msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Atenção: O envio da NF-e falhou...")) = vbYes Then
        
            x = Salva_Log("EnviaNFe2G", msgResultado, versao, msgDados, msgRetWS, NFe)
            
        End If
       
       Case 103
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "A NF-e foi enviada para o Web Service desejado, busque o resultado processamento do lote, informando o número do recibo de entrega: " + nroRecibo, vbInformation, "Atenção: NF-e enviada!"
       
       x = Salva_NF(NFeAssinada, nomeArquivo)    ' Salva NF-e assinada
       
       Case 108, 109
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "O Web Service com problemas de recepção, adote o processo de contingência se a emissão da NF-e for urgente." + nroRecibo, vbInformation, "Atenção: WS com problemas na recepção!"
       
       Case Else                                ' problemas na recepção do WS

        If (MsgBox("Descrição do erro: " & Chr(13) & Chr(13) & cStat & " - " & msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Atenção: O WS rejeitou a NF-e (lote) ...")) = vbYes Then
        
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
'   com uso do enviaNFe2G, necessário informar a NF-e assinada.
'
Dim msgDados As String
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaWS As String
Dim certificado As String
Dim cStat As Long
Dim x As Boolean
'
'  As variáveis do proxy devem ser informadas se necessário
'
'  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
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
' define as variáveis que passam/recebem informações importantes
'
Dim NFeAssinada As String
Dim nroRecibo As String
Dim procNFe As String       ' procNFe -> NF-e + protocolo de autorização de uso da NF-e, deve ser mantido em arquivo e distribuído ao destinatário.
'
' parâmetros novos
'
Dim versao As String        ' versão do leiaute (1.10 ou 2.00), serve para escolher o WS da SEFAZ
Dim nroProtocolo As String  ' número do protocolo de autorização de uso da NF-e, este número é necessário para cancelar a NF-e
Dim dhProtocolo As String   ' data e hora de autorização de uso da NF-e
Dim cMsg As String          ' código da mensagem da SEFAZ, a SEFAZ pode utiliza-lo como canal de comunicação com o emissor
Dim xMsg As String          ' literal da mensagem da SEFAZ
'
'
'  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
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


versao = cbVersao.Text      ' versão da NF-e: 1.10 ou 2.00
  
txtEntrada.Text = ""
txtRetorno.Text = ""


' NFeAssinada = ' informar com a NFeAssinada que retornou do enviaNFeSCAN

'
'  carregar arquivo XML da NF-e assinada na string NFeAssinada, a NF-e assinada
'  é necessário para montar o procNFe
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   
   Exit Sub

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo XML da NF-e para transmissão." & vbCrLf & _
            Err.Description
        Exit Sub
End If

On Error GoTo 0
'
'  carrega o conteúdo do arquivo xml em NFe
'
Open CommonDialog1.FileName For Input As #1
NFeAssinada = Input$(LOF(1), 1)
Close #1

' nroRecibo =   ' informar com o nroRecibo que retorno do enviaNFeSCAN

If txtNroRecibo.Text = "" Then
        MsgBox "Necessário informar o número do recibo para buscar o resultado do processamento!", vbCritical, "Atenção:"
            Exit Sub
End If

nroRecibo = txtNroRecibo.Text

If cbAmb.Text = "Produção" Then
   tpAmbiente = 1
Else
   tpAmbiente = 2
End If

'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
   '  se cStat = 105 - Lote em processamento, significa que a SEFAZ ainda não conseguiu processar a NF-e e a aplicação deve persistir até que o cStat seja diferent de 105
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
'           105 – lote em processamento -> tentar novamente
'           106 – lote não localizado   -> tentar enviar o lote novamente ou verificar se o nroRecibo está correto
'           100 – NF-e autorizada       -> OK
'           2xx – motivo de rejeição do WS -> erro na elaboração da NF-e, verificar o código de erro e corrigir a NF-e
Select Case cStat

       Case 5001 To 6423                ' erro da DLL
        
        If (MsgBox("Descrição do erro: " & Chr(13) & Chr(13) + msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Atenção: O envio da NF-e falhou...")) = vbYes Then
        
            x = Salva_Log("BuscaNFe2G", msgResultado, versao, msgDados, msgRetWS, NFeAssinada)
            
        End If
       
       Case 100    ' NF-e autorizada
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "A NF-e autorizada, guarde o número do protocolo de autorização de uso (" + nroRecibo + ") e respectivo procNFe (NF-e + autorização de uso).", vbInformation, "Atenção: NF-e autorizada!"
       
       x = Salva_NF_UTF8(procNFe, nroRecibo & "-nfeproc.xml")    ' Salva NF-e assinada
       
       txtProAutoUso.Text = nroProtocolo   ' número do potocolo de autorizacao de uso
                                           ' este número é necessário para um eventual cancelamento da NF-e.
       
       Case 101    ' NF-e denegada
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "A NF-e foi denegada, guarde o número do protocolo de denegação de uso (" + nroRecibo + ") e respectivo procNFe (NF-e + denegação de uso).", vbInformation, "Atenção: NF-e denegada!"
       
       x = Salva_NF_UTF8(procNFe, nroRecibo & "-nfeproc.xml")    ' Salva NF-e assinada
       
       txtProAutoUso.Text = nroProtocolo   ' número do potocolo de denegacao de uso
              
       Case 106    ' Lote não localizado, verifique se o número do recibdo está correto
       
       MsgBox msgResultado, vbInformation, "Atenção: Lote não localizado"
       
       Case 108, 109
       
       MsgBox msgResultado & Chr(13) & Chr(13) + "O Web Service está com problemas de recepção, adote o processo de contingência se a emissão da NF-e for urgente." + nroRecibo, vbInformation, "Atenção: WS com problemas na recepção!"
       
       Case Else

        If (MsgBox("Descrição do erro: " & Chr(13) & Chr(13) & cStat & " - " & msgResultado & Chr(13) & Chr(13) + "Deseja gravar o log de erro?", vbCritical + vbYesNo, "Atenção: O WS rejeitou a NF-e (lote) ...")) = vbYes Then
        
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
'Dim msgCabec As String       não é mais necessário
Dim msgRetWS As String
Dim msgResultado As String
Dim siglaUF As String
Dim siglaWS As String
Dim certificado As String
'
'  As variáveis do proxy devem ser informadas se necessário
'
'  proxy deve ser informado com o endereço da url : porta, ex: 192.168.15.1:443
'
Dim proxy As String
Dim usuario As String
Dim senha As String
'
Dim ambiente As Integer
'
' parâmetro novo - utilizado para escolha da versão do WS
'
Dim versao As String
'
' define as variáveis que passam/recebem informações importantes
'
Dim ChaveNFe As String
'
'
'  IMPORTANTE: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
'
'
proxy = ""
usuario = ""
senha = ""
msgDados = ""
'msgCabec = ""  não é mais necessário
msgRetWS = ""
msgResultado = ""

certificado = txtCertificado.Text
              ' informar com o assunto da certificado digital
              ' Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"

siglaWS = cbWS.Text ' se a UF utilizar SEFAZ Virtual, informar SVRS (Ex. RJ, SC, etc.) ou SVAN (Ex. ES, RN, etc.)

versao = cbVersao.Text ' versão do Web Service, a versão anterior é 1.07, a versão nova é 2.00

txtEntrada.Text = ""
txtRetorno.Text = ""
 
ChaveNFe = InputBox("Informe a Chave de Acesso da NF-e", "Consulta Status NF-e")


If ChaveNFe = "" Then '
        MsgBox "Necessário informar a chave de acesso da NF-e para consultar o Status da NF-e.", vbCritical, "Atenção:"
            Exit Sub
End If

If cbAmb.Text = "Produção" Then
   ambiente = 1
Else
   ambiente = 2
End If

Dim Resultado As Long

'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
                                      
MsgBox msgResultado & Chr(13) & Chr(13), vbInformation, "Atenção: Consulta Status da NF-e"

End Sub

Private Sub Form_Load()

'
'   Exemplo para obter versão da DLL em uso
'
'
' instancia classe
'
On Error GoTo InexisteDLL
'
' referenciando a DLL em late binding
' não é necessário fazer o reference da DLL
' o intelisense não funciona
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
' obtem versão
'

' MsgBox "A versão da DLL é: " + objNFeUtil.versao, vbInformation, "Resultado"

versaoDLL = objNFeUtil.versao

Form1.Caption = "Demo VB 6.0 da " & versaoDLL

MsgBox "Aplicação Demo em VB 6.0 para demostrar o uso da: " & Chr(13) & Chr(13) & _
        versaoDLL & Chr(13) & Chr(13) & _
       "Para detalhes da forma correta de instalação da DLL e uso em VB 6.0 " & Chr(13) & _
       "queira examinar o GUIA DE USO DA DLL, disponível em: " & Chr(13) & _
       "http://www.flexdocs.com.br/guiaNFe" & Chr(13) & Chr(13) & _
       "www.flexdocs.com.br (c) 2008-2011 - todos os direitos reservados", vbInformation, "Informação da aplicação"

'
' libera classe
'
Set objNFeUtil = Nothing

Exit Sub

InexisteDLL:

If (Err.Number = -2147024894) Then
   MsgBox "Erro número : " & Str$(Err.Number) & " (80070002/80040002) " & Err.Description & Chr(13) & Chr(13) & "A DLL NFe_Util e as pastas URL e Schemas devem ser copiadas para a pasta: " + App.Path & Chr(13) & Chr(13) & "Em modo depuração também é necessário que estes arquivos existam no pasta do VB98", vbCritical, "Erro"
ElseIf (Err.Number = -2146234304) Then
   MsgBox "Erro número : " & Str$(Err.Number) & " (80131040/80231040) " & Err.Description & Chr(13) & "A versão da DLL existente na pasta da aplicação é diferente da registrada no equipamento...", vbCritical, "Erro"
Else
   MsgBox "Erro número : " & Str$(Err.Number) & " " & Err.Description & Chr(13) & "Anote o código de erro e verifique se existe solução na FAQ.", vbCritical, "Erro"

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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   Salva_Log = False
   Exit Function

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o arquivo de log para gravação." & vbCrLf & _
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
Print #1, "2.Status de retorno da função:"
Print #1, "----------------------------------------------"
Print #1, msgResultado
Print #1, "----------------------------------------------"
Print #1, "3.XML da NF-e:"
Print #1, "----------------------------------------------"
Print #1, NFe
Print #1, "----------------------------------------------"
Print #1, "4.Área de Dados:"
Print #1, "----------------------------------------------"
Print #1, UTF8_Encode(msgDados)
Print #1, "----------------------------------------------"
Print #1, "5.Área de Retorno do WS:"
Print #1, "----------------------------------------------"
If msgRetWS = "" Then
Print #1, "***SEM RETORNO***"
Else
Print #1, UTF8_Encode(msgRetWS)
End If
Print #1, "6.Versão da DLL em uso:"
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   Salva_NF = False
   Exit Function

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o nome do arquivo XML da NF-e para gravação." & vbCrLf & _
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

If Err.Number = cdlCancel Then 'cancelado pelo usuário
   Salva_NF = False
   Exit Function

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o nome do arquivo XML da NF-e para gravação." & vbCrLf & _
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
'  Converte a string para codificação UTF-8
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
