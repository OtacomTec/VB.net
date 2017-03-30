VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta de Corre��o Eletr�nica"
   ClientHeight    =   8940
   ClientLeft      =   5235
   ClientTop       =   1815
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton certificadoButton 
      Caption         =   "Certificado Digital"
      Height          =   495
      Left            =   8520
      TabIndex        =   26
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame7 
      Caption         =   "N�mero da CC-e: "
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
      Left            =   6600
      TabIndex        =   24
      Top             =   1080
      Width           =   1695
      Begin VB.ComboBox cb_nroCCe 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form2.frx":0000
         Left            =   120
         List            =   "Form2.frx":0040
         TabIndex        =   25
         Text            =   "1"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Data/Hora CC-e"
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
      Left            =   4560
      TabIndex        =   22
      Top             =   1080
      Width           =   1935
      Begin VB.TextBox txtdhEvento 
         Height          =   285
         Left            =   135
         TabIndex        =   23
         Text            =   "2011-12-18 00:00:00"
         Top             =   240
         Width           =   1680
      End
   End
   Begin VB.ComboBox cb_descEventoAcentuado 
      Height          =   315
      ItemData        =   "Form2.frx":008B
      Left            =   6600
      List            =   "Form2.frx":0095
      TabIndex        =   20
      Text            =   "S"
      Top             =   1800
      Width           =   615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Texto da Corre��o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   18
      Top             =   2160
      Width           =   8280
      Begin VB.TextBox txtCorrecao 
         Height          =   1110
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   225
         Width           =   8040
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Chave da NF-e"
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
      Left            =   0
      TabIndex        =   16
      Top             =   1080
      Width           =   4455
      Begin VB.TextBox txtChaveNFe 
         Height          =   285
         Left            =   135
         TabIndex        =   17
         Top             =   240
         Width           =   4200
      End
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
      Left            =   0
      TabIndex        =   10
      Top             =   375
      Width           =   8295
      Begin VB.TextBox txtCertificado 
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Text            =   "CN=M R M KATO ASAKURA - EPP:69621187915, OU=AC CAIXA PJ-1 V1, OU=Caixa Economica Federal, O=ICP-Brasil, C=BR"
         Top             =   240
         Width           =   8040
      End
   End
   Begin VB.ComboBox cbWS 
      Height          =   315
      ItemData        =   "Form2.frx":009F
      Left            =   975
      List            =   "Form2.frx":00D3
      TabIndex        =   9
      Text            =   "SP"
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox cbUF 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Form2.frx":011D
      Left            =   2775
      List            =   "Form2.frx":016C
      TabIndex        =   8
      Text            =   "SP"
      Top             =   0
      Width           =   735
   End
   Begin VB.ComboBox cbAmb 
      Height          =   315
      ItemData        =   "Form2.frx":01D4
      Left            =   6855
      List            =   "Form2.frx":01DE
      TabIndex        =   7
      Text            =   "Homologa��o"
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   " Entrada / Mensagem Enviada  (msgDados)   (identado para melhor visualiza��o) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Width           =   8280
      Begin VB.TextBox txtEntrada 
         Height          =   2190
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   225
         Width           =   8040
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Retorno / Mensagem de Retorno  (msgRetWS)  (identado para melhor visualiza��o) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   8280
      Begin VB.TextBox txtRetorno 
         Height          =   2235
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   270
         Width           =   7995
      End
   End
   Begin VB.ComboBox cbVersao 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Form2.frx":01F9
      Left            =   4695
      List            =   "Form2.frx":0200
      TabIndex        =   2
      Text            =   "1.00"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton ExitlButton 
      Caption         =   "Sair"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton CCeButton 
      Caption         =   "Enviar CC-e"
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Indicador de acentua��o da descri��o do Evento e das condi��es de uso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   1875
      Width           =   6615
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
      Left            =   15
      TabIndex        =   15
      Top             =   75
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
      Left            =   2415
      TabIndex        =   14
      Top             =   75
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
      Left            =   5895
      TabIndex        =   13
      Top             =   75
      Width           =   1455
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
      Left            =   3975
      TabIndex        =   12
      Top             =   75
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'  Exemplo de uso da carta de corre��o eletr�nica da NF-e, para maiores detalhes veja o guia de uso da DLL: http://www.flexdocs.com.br/guiaNFe/WS.evento.CCe.html
'
'

Private Sub certificadoButton_Click()
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


Private Sub ExitlButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
'
' inicializa data do evento, a data do evento deve ser menor que a data e hora do servidor da SEFAZ e maior que a data de autoriza��o da NF-e
' verifique se a data do equipamento est� sincronizada com o Servidor da SEFAZ, n�o pode nunca estar adiantada.
'
txtdhEvento.Text = Format$(Now, "yyyy-mm-dd HH:mm:ss")
End Sub
Private Sub CCeButton_Click()
'
'  Carta de Corre��o eletr�nica
'
'  Exemplo de uso da funcionalidade de carta de corre��o eletr�nica
'
'  veja detalhes da funcionalidade em: http://www.flexdocs.com.br/guiaNFe/WS.evento.CCe.html
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
' define as vari�veis que passam informa��es para a DLL
'
Dim versao As String            ' utilizado para escolha da vers�o do WS, informar "1.00"
Dim ChaveNFe As String          ' chave da NF-e objeto de carta de corre��o eletr�nica
Dim Correcao  As String         ' texto da corre��o - string com at� 1000 caracteres
Dim dhCorrecao As String        ' data e hora da corre��o
Dim nCCe As Long                ' n�mero da carta de corre��o, deve ser um n�mero sequencial iniciado em 1, o valor m�ximo � 20
Dim descEventoAcentuado As Long ' indicardor de acentua��o da descri��o do evento e das condi��es de uso, deve ser informado com 0-n�o/1-sim
                                  ' indicar com 0 para as UF que n�o aceitam acento como � o caso do MT
                                  ' IMPORTANTE: o controle da acentua��o do texto da corre��o � da aplica��o do usu�rio, este indicador serve
                                  ' apenas para que a DLL informe os campos descEvento e xCondUso sem acentua��.
'
'  par�metros que devolvem informa��es
'
Dim procCCe As String           ' estrturura XML que cont�m a carta de corre��o eletr�nica e registro do evento da carta de corre��o eletr�nica,
                                ' que deve ser mantido pelo emissor e distribu�do ao destinat�rio.
Dim nProtocoloCCe  As String    ' n�mero do protocolo de  registro do evento da carta de corre��o eletr�nica devolvido pela SEFA
Dim dProtocoloCCe  As String    ' data e hora de  registro do evento da carta de corre��o eletr�nica

Dim resposta As Integer         ' retorno do msgBox
Dim sucesso As Boolean          ' retorno da grava��o do log, n�o utilizado

Dim cResultado2 As String       ' para uso no identaXML
Dim msgResultado2 As String     ' para uso no identaXML


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
procCCe = ""
nProtocoloCCe = ""
dProtocoloCCe = ""

certificado = txtCertificado.Text
              ' informar com o assunto da certificado digital
              ' Ex.: "CN=NFe - Associacao NF-e:99999090910270, C=BR, L=PORTO ALEGRE, O=Teste Projeto NFe RS, OU=Teste Projeto NFe RS, S=RS"

siglaWS = cbWS.Text ' se a UF utilizar SEFAZ Virtual, informar SVRS (Ex. RJ, SC, etc.) ou SVAN (Ex. ES, RN, etc.)
 
txtEntrada.Text = ""
txtRetorno.Text = ""

ChaveNFe = Trim(txtChaveNFe.Text)       ' elimina espa�os em branco no in�cio e fim do texto

If Len(ChaveNFe) <> 44 Then '
        MsgBox "Necess�rio informar a chave de acesso da NF-e objeto de carta de Corre��o com 44 d�gitos", vbCritical, "Aten��o: Erro no preenchimento da chave de acesso da NF-e"
            Exit Sub
End If

Correcao = Trim(txtCorrecao.Text)       ' elimina espa�os em branco no in�cio e fim do texto

If (InStr(txtCorrecao, Chr(13)) > 0) Or (InStr(txtCorrecao, Chr(10)) > 0) Then

        '
        ' para evitar este erro, o usu�rio pode substituir o chr(13) e chr(10) por espa�o ou outro caractere como o ;
        '
        
        MsgBox "O texto da corre��o n�o deve ter quebra de linha", vbCritical, "Aten��o: Erro no preenchimento do texto da corre��o"
            Exit Sub
End If


If Len(txtCorrecao) < 15 Then '
        MsgBox "Necess�rio informar o texto da corre��o com no m�nimo 15 caracteres", vbCritical, "Aten��o: Erro no preenchimento do texto da corre��o"
            Exit Sub
End If

If Len(txtCorrecao) > 1000 Then '
        MsgBox "O texto da corre��o deve ter at� 1000 caracteres", vbCritical, "Aten��o: Erro no preenchimento do texto da corre��o"
            Exit Sub
End If

'
' estamos utilizando os seguintes par�metro fixo na demonstra��o para facilitar o processo
'
versao = "1.00"                     ' vers�o do leiaute da carta de corre��o
dhCorrecao = Str(DateTime.Now)      ' data e hora da corre��o
                                    ' *** Aten��o ***
                                    ' se a data e hora for superior � data do Servidor da SEFAZ, ocorrer� a rejei��o: 578 - Rejei��o: A data do evento n�o pode ser maior que a data do processamento que volta em dhRegEvento no XML de retorno do WS
                                    ' se a data e hora for inferior � data de autoriza��o da NF-e, ocorrer� a rejei��o: 577 - Rejei��o: A data do evento n�o pode ser menor que a data de emiss�o da NF-e
                                    '
nCCe = 1                            ' n�mero da carta de corre��o, deve ser um n�mero sequencial iniciado em 1, o valor m�ximo � 20
descEventoAcentuado = 0             ' indicardor de acentua��o da descri��o do evento e das condi��es de uso, deve ser informado com 0-n�o/1-sim


If cbAmb.Text = "Produ��o" Then
   ambiente = 1
Else
   ambiente = 2
End If

Dim cStat As Long   ' status da chamada, veja os valores em http://www.flexdocs.com.br/guiaNFe/WS.evento.CCe.html

'
' referenciando a DLL em late binding
' n�o � necess�rio fazer o reference da DLL
' o intelisense n�o funciona
'
Dim objNFeUtil As Object

Set objNFeUtil = CreateObject("NFe_util_2G.util")

'
'
Screen.MousePointer = vbHourglass    ' ampulheta
'
'
procCCe = objNFeUtil.EnviaCCe2G(siglaWS, ambiente, certificado, versao, msgDados, msgRetWS, cStat, msgResultado, ChaveNFe, txtCorrecao, descEventoAcentuado, nCCe, dhCorrecao, nProtocoloCCe, dProtocoloCCe, proxy, usuario, senha, licenca)
'
'
Screen.MousePointer = vbDefault ' normal
'
' mostra mensagem XML enviada e a mensagem de retorno do WS
'
'
' identa XML para faciliar a visualiza��o
'
txtEntrada.Text = objNFeUtil.IdentaXML(msgDados, cResultado2, msgResultado2)          ' string com a mensagem XML enviado ao WS

txtRetorno.Text = objNFeUtil.IdentaXML(msgRetWS, cResultado2, msgResultado2)          ' string com a mensagem XML da resposta do WS

If cStat = 135 Then
                                      
   '
   ' grave o CCe, pois o XML deve ser mantido pelo emissor, al�m de ser distribu�do para o destinat�rio tamb�m.
   '
   resposta = MsgBox(msgResultado & Chr(13) & Chr(13) + "Protocolo de registro do evento : " + nProtocoloCCe + Chr(13) & Chr(13) + "Data e hora de registro evento: " + dProtocoloCCe + Chr(13) & Chr(13) + "Deseja gravar o procCCe?", vbInformation + vbYesNo, "Aten��o: Carta de Corre��o eletr�nica da NF-e")

   If resposta = vbYes Then
      sucesso = Salva_CCe(procCCe, nProtocoloCCe)
   End If
Else

    resposta = MsgBox(msgResultado & Chr(13) & Chr(13) + "O envio da Carta de Corre��o eletr�nica Falhou, deseja gravar arquivo de log? ", vbCritical + vbYesNo, "Aten��o: Carta de Corre��o eletr�nica da NF-e")
   
   If resposta = vbYes Then
      sucesso = Salva_Log("EnviaCCe2G", msgResultado, msgDados, msgRetWS)
   End If

End If
End Sub

Private Function Salva_Log(ByVal Funcionalidade As String, ByVal msgResultado As String, ByVal msgDados As String, ByVal msgRetWS As String) As Boolean

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
Print #1, "3.�rea de Dados:"
Print #1, "----------------------------------------------"
Print #1, UTF8_Encode(msgDados)
Print #1, "----------------------------------------------"
Print #1, "4.�rea de Retorno do WS:"
Print #1, "----------------------------------------------"
If msgRetWS = "" Then
Print #1, "***SEM RETORNO***"
Else
Print #1, UTF8_Encode(msgRetWS)
End If
Print #1, "5.Vers�o da DLL em uso:"
Print #1, "----------------------------------------------"
Print #1, Form1.versaoDLL
Close #1


End Function

Private Function Salva_CCe(ByVal CCe As String, ByVal Nome As String) As Boolean

On Error Resume Next

Salva_CCe = True
CommonDialog1.DialogTitle = "Informe o nome do arquivo para gravar a CC-e"
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
   Salva_CCe = False
   Exit Function

ElseIf Err.Number <> 0 Then ' erro desconhecido
        MsgBox "Erro: " & Format$(Err.Number) & _
            " ao selecionar o nome do arquivo XML da CC-e para grava��o." & vbCrLf & _
            Err.Description
         Salva_CCe = False
        Exit Function
End If
On Error GoTo 0

Open CommonDialog1.FileName For Output As #1
Print #1, UTF8_Encode(CCe)                      ' tratamento para que o XML seja aberto pelo Internet Explorer
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

