Imports System.IO
Imports System.Net

Public Class excel
    Private Sub initExcel_Click(sender As Object, e As EventArgs) Handles initExcel.Click
        Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If My.Computer.Network.IsAvailable Then
            Descargar()
            Process.Start("EXCEL.EXE", Path + "/examenes-practicas/Practica.xlsx")
            Me.Hide()
            Instrucciones1.Width = Screen.PrimaryScreen.Bounds.Width
            Instrucciones1.Show()
        Else
            MsgBox("Computer is not connected.")
        End If
    End Sub

    Public Function Descargar()
        Dim remoteUri As String = "https://capacitacioninformatica.com/uploads/practicas/efac44645f6d17ec2ad7007627d44e2b.xlsx"
        Dim fileName As String = "Practica.xlsx"
        Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim client = New WebClient()
        If (Not System.IO.Directory.Exists(Path + "/examenes-practicas/")) Then
            System.IO.Directory.CreateDirectory(Path + "/examenes-practicas/")
        End If
        client.DownloadFile(remoteUri, Path + "/examenes-practicas/" + fileName)
    End Function
End Class

