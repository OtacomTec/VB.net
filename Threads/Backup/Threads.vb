Module macoratti
    Public Sub prenchelista(ByVal lstbox As ListBox, ByVal id As Integer)
        Dim j As Integer = 1
        While True
            lstbox.Items.Add(" Thread " & id & " : " & CStr(j))
            j += 1
        End While
    End Sub
End Module
