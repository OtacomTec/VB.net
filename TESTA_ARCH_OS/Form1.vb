Imports System.Runtime.InteropServices

Public Class Form1

    <DllImport("kernel32.dll", SetLastError:=True, CallingConvention:=CallingConvention.Winapi)> _
       Public Shared Function IsWow64Process(<[In]()> ByVal hProcess As IntPtr, <Out()> ByRef lpSystemInfo As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean

    End Function
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim booVersao64 As Boolean = Tipo_Windows2()

        If booVersao64 = True Then
            MsgBox("64-bit")
        Else
            MsgBox("32-bit")
        End If

    End Sub

    Public Function Tipo_Windows2() As Boolean

        If IntPtr.Size = 8 OrElse (IntPtr.Size = 4 AndAlso Is32BitProcessOn64BitProcessor()) Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Function Is32BitProcessOn64BitProcessor() As Boolean

        Dim retVal As Boolean

        IsWow64Process(Process.GetCurrentProcess().Handle, retVal)

        Return retVal

    End Function

End Class
