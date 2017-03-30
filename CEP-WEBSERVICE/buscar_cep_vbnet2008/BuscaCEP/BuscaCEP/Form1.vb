Imports System.Xml

Public Class Form1


    Private Sub txtCEP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCEP.TextChanged

        Dim strCEP As String
        strCEP = txtCEP.Text
        Dim strResult As String

        If strCEP.Length = 8 Then

            lbMsg.Text = "Consultando http://www.buscarcep.com.br/, aguarde..."
            Application.DoEvents()

            Dim filename As String = "http://www.buscarcep.com.br/?cep=" & strCEP & "&formato=xml"
            Dim reader As New XmlTextReader(filename)
            Dim strTempName, strTempValue As String
            reader.MoveToContent()

            Do While reader.Read
                strTempName = reader.Name
                If reader.NodeType = XmlNodeType.Element Then
                    reader.Read()
                    strTempValue = reader.Value
                    Select Case strTempName
                        Case "tipo_logradouro"
                            txtTipoLogradouro.Text = strTempValue
                        Case "logradouro"
                            txtLogradouro.Text = strTempValue
                        Case "bairro"
                            txtBairro.Text = strTempValue
                        Case "cidade"
                            txtCidade.Text = strTempValue
                        Case "uf"
                            Select Case strTempValue
                                Case "AC"
                                    cmbUF.SelectedItem = "Acre"
                                Case "AL"
                                    cmbUF.SelectedItem = "Alagoas"
                                Case "AP"
                                    cmbUF.SelectedItem = "Amapá"
                                Case "AM"
                                    cmbUF.SelectedItem = "Amazonas"
                                Case "BA"
                                    cmbUF.SelectedItem = "Bahia"
                                Case "CE"
                                    cmbUF.SelectedItem = "Ceará"
                                Case "DF"
                                    cmbUF.SelectedItem = "Distrito Federal"
                                Case "ES"
                                    cmbUF.SelectedItem = "Espírito Santo"
                                Case "GO"
                                    cmbUF.SelectedItem = "Goiás"
                                Case "MA"
                                    cmbUF.SelectedItem = "Maranhão"
                                Case "MT"
                                    cmbUF.SelectedItem = "Mato Grosso"
                                Case "MS"
                                    cmbUF.SelectedItem = "Mato Grosso do Sul"
                                Case "MG"
                                    cmbUF.SelectedItem = "Minas Gerais"
                                Case "PA"
                                    cmbUF.SelectedItem = "Pará"
                                Case "PB"
                                    cmbUF.SelectedItem = "Paraíba"
                                Case "PR"
                                    cmbUF.SelectedItem = "Paraná"
                                Case "PE"
                                    cmbUF.SelectedItem = "Pernambuco"
                                Case "PI"
                                    cmbUF.SelectedItem = "Piauí"
                                Case "RJ"
                                    cmbUF.SelectedItem = "Rio de Janeiro"
                                Case "RN"
                                    cmbUF.SelectedItem = "Rio Grande do Norte"
                                Case "RS"
                                    cmbUF.SelectedItem = "Rio Grande do Sul"
                                Case "RO"
                                    cmbUF.SelectedItem = "Rondônia"
                                Case "RR"
                                    cmbUF.SelectedItem = "Roraima"
                                Case "SC"
                                    cmbUF.SelectedItem = "Santa Catarina"
                                Case "SP"
                                    cmbUF.SelectedItem = "São Paulo"
                                Case "SE"
                                    cmbUF.SelectedItem = "Sergipe"
                                Case "TO"
                                    cmbUF.SelectedItem = "Tocantins"
                            End Select
                        Case "resultado"
                            If strTempValue = "1" Then
                                ' CEP OK
                            Else
                                txtTipoLogradouro.Text = ""
                                txtLogradouro.Text = ""
                                txtBairro.Text = ""
                                txtCidade.Text = ""
                                If strTempValue = "-1" Then
                                    MsgBox("CEP não encontrado.")
                                ElseIf strTempValue = "-2" Then
                                    MsgBox("Formato de CEP inválido.")
                                ElseIf strTempValue = "-3" Then
                                    MsgBox("Busca de CEP congestionada. " & vbCrLf & "Aguarde alguns segundos e tente novamente.")
                                Else
                                    MsgBox("Erro na busca de CEP.")
                                End If
                            End If
                    End Select
                End If
            Loop

            lbMsg.Text = "Busca de CEP"
            Application.DoEvents()
        End If

    End Sub

    Private Sub cmbUF_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbUF.SelectedIndexChanged

    End Sub
End Class
