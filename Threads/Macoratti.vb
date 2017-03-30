Imports System.threading
Public Class Macorati

    Public lstbox As ListBox
    Public id As Integer

    Public Sub prenchelista()
        Dim j As Integer = 1
        While True
            lstbox.Items.Add(" Thread " & id & " : " & CStr(j))
            j += 1
            Thread.CurrentThread.Sleep(1000)
        End While
    End Sub
End Class
Public Class cAreaQuadrado

    Public lado As Double
    Public Area As Double

    Public Event ThreadCompleta(ByVal Area As Double)

    Public Sub CalculaArea()
        SyncLock GetType(cAreaQuadrado)
            Area = lado * lado
        End SyncLock
        'RaiseEvent ThreadCompleta(Area)
    End Sub

End Class

