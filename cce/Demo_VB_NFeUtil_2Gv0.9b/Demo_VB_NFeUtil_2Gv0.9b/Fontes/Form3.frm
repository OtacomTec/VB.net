VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Geração do PDF do DANFE"
   ClientHeight    =   6855
   ClientLeft      =   5895
   ClientTop       =   2760
   ClientWidth     =   10620
   LinkTopic       =   "Form3"
   ScaleHeight     =   6855
   ScaleWidth      =   10620
   Begin VB.CommandButton cmdVisualizaDANFE 
      Caption         =   "Visualizar DANFE"
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame Frame6 
      Caption         =   "[ Número e data de registro do DPEC ]"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   8055
      Begin VB.TextBox txtDPEC 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "[ Separador de Item ]"
      Height          =   735
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Width           =   4215
      Begin VB.ComboBox cbSeparadorItem 
         Height          =   315
         ItemData        =   "Form3.frx":0000
         Left            =   120
         List            =   "Form3.frx":0010
         TabIndex        =   13
         Text            =   "Linha"
         Top             =   240
         Width           =   3165
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " [Posição do Quadro do Recibo]"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3735
      Begin VB.ComboBox cbQuadroRecibo 
         Height          =   315
         ItemData        =   "Form3.frx":0037
         Left            =   120
         List            =   "Form3.frx":0041
         TabIndex        =   11
         Text            =   "Superior"
         Top             =   240
         Width           =   3165
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " [Quadros/Colunas do DANFE]"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   8055
      Begin VB.CheckBox chkColDesconto 
         Caption         =   "Coluna Desconto"
         Height          =   200
         Left            =   6240
         TabIndex        =   9
         Top             =   270
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkQdISSQN 
         Caption         =   "Quadro ISSQN"
         Height          =   200
         Left            =   4320
         TabIndex        =   8
         Top             =   270
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkQdFatura 
         Caption         =   "Quadro Fatura/Duplicata"
         Height          =   200
         Left            =   1920
         TabIndex        =   7
         Top             =   270
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkQdEmitente 
         Caption         =   "Quadro Emitente"
         Height          =   200
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " [Origem dos dados do emissor]"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8055
      Begin VB.TextBox txtOrigemEmissor 
         BackColor       =   &H80000003&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   7815
      End
      Begin VB.ComboBox cbDescEventoAcentuado 
         Height          =   315
         ItemData        =   "Form3.frx":0059
         Left            =   120
         List            =   "Form3.frx":0063
         TabIndex        =   3
         Text            =   "Informações do XML"
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "XML da NF-e ou procNF-e  (identado para melhor visualização) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   8025
      Begin VB.TextBox txtNFe 
         Height          =   3030
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   225
         Width           =   7800
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVisualizaDANFE_Click()
'
' declaração das variáveis que serão utilizadas na passagem de parâmetros da DLL
'
Dim XML As String                 ' informar o XML da NF-e da versão 2.00
Dim OrigDadosEmissor As String    ' origem dos dados do emissor
Dim quadroRecibo As String        ' posicão de impressão do quadroRecibo [S]uperior [I]nferior
Dim quadroFatura As String        ' indicador de impressão do quadro de Fatura/Duplicata
Dim quadroISSQN As String         ' indicador de impressão do quadro ISSQN
Dim DPEC As String                ' data e número do registro do DPEC
Dim separadorItem As String       ' indicador do separador de item que será utilizado quando o item ocupar mais de
                                  ' uma linha: [L]inha, [T]racejado, espaço em [B]ranco, e [Z]ebrado.
Dim gravaPDF As String            ' serve para indicar o nome do arquivo que será gravado, se o conteúdo for omitido
                                  ' o PDF será visualizado na tela;
                                  ' informe [NFeId.PDF] para gravar um arquivo com a chave da NF-e;
                                  ' Se informado o literal [IMPRIMIR=n], a DLL irá enviar o PDF para impressora padrão,
                                  ' o n pode varia de 1 a 5.
Dim cResultado As Long            ' código deretorno da chamada da DLL
Dim msgResultado As String        ' literal com resultado da chamada da DLL

Dim XMLIdentado As String       ' XML identado
Dim cResultado2 As String       ' para uso no identaXML
Dim msgResultado2 As String     ' para uso no identaXML


'
'
'  Importante: todas as variáveis utilizadas como parâmetro da DLL devem ser inicializadas
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
'   Carrega o conteúdo do nome do arquivo em XMLString
'
Open CommonDialog1.FileName For Input As #1
XML = Input$(LOF(1), #1)
Close #1
'
OrigDadosEmissor = ""           ' origem dos dados do emissor no XML, possibilidades:
                                '  sem conteúdo -> os dados do emissor serão obtidos do XML;
                                '  nome arquivo -> a imagem informada irá ocupar todo o quadro dados do emitente;
                                '  literal [SEM DADOS EMITENTE] -> nenhum dado será impresso no quado dados do emitente;
quadroRecibo = "S"              ' quadro do recibo no topo
quadroFatura = "S"              ' imprimir o quadro de Fatura/Duplicatas
quadroISSQN = "S"               ' imprimir o quadro de ISSQN
DPEC = ""                       ' informar quando a NF-e tiver sido emitido em contingência DPEC
separadorItem = "T"             ' traço para separar o item quando o item ocupar mais de uma linha
gravaPDF = ""                   ' informar o nome do arquivo do PDF.
                                ' Se informado o literal [NFeId.PDF] a DLL irá gravar o PDF identificado com o chave de acesso da NF-e.
                                ' A omissão do nome gera uma visualização em tela.
                                '-----------------------------------------------------
                                'NOVO***NOVO***NOVO***NOVO
                                ' Parâmetro gravaPDF, valores válidos:
                                ' nomeArquivo -> grava PDF com nomeArquivo se existir apenas o nomeArquivo no parâmetro;
                                ' [NFeId.PDF] -> grava arquivo com nome = chave de acesso da NF-e;
                                ' [SEM COLUNA DESCONTO] -> não gera a coluna de desconto;
                                ' [RODAPE=texto do rodape] -> imprime o "texto do rodape" informado no RODAPE;
                                ' [PASTA=] -> indica a pasta de gravação do PDF;
                                ' [VISUALIZAR] -> indica visualização da PDF;
                                ' [ARQUIVO=nomeArquivo] -> grava o PDF com o nome indicado;
                                ' [COM FATURA] -> indica que os dados da fatura devem ser impressos em informações adicionais;
                                ' [MENSAGEM=texto da mensagem] -> imprime o "texto da mensagem" informado no corpo do DANFE;
                                '

cResultado = 0
msgResultado = ""
'
' instancia a DLL - late binding
'
Dim objNFeUtil As Object
'
Set objNFeUtil = CreateObject("NFe_Util_2G.util")
'
' chama DLL
'
'XML = objNFeUtil.LeArquivoANSI(CommonDialog1.FileName, cResultado2, msgResultado2)

XMLIdentado = objNFeUtil.IdentaXML(XML, cResultado2, msgResultado2)

If cResultado2 = 7310 Then
  
  txtNFe.Text = XMLIdentado

Else

  txtNFe.Text = XML

End If

'
'
'
cResultado = objNFeUtil.geraPdfDANFE(XML, OrigDadosEmissor, quadroRecibo, quadroFatura, quadroISSQN, DPEC, separadorItem, gravaPDF, msgResultado)

'
'  tratar retorno
'

If cResultado < 7902 Then            ' sucesso, conversão OK

MsgBox msgResultado, vbInformation, "Informação"
 
Else
'

MsgBox "Processo de geração do PDF falhou..." & vbCrLf & msgResultado, vbExclamation, "Atenção"
 
End If
'
'  liberar DLL
'
Set objNFeUtil = Nothing
End Sub
