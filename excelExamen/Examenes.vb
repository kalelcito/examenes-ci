Imports System.IO
Imports System.Net

Public Class excel
    Private Sub initExcel_Click(sender As Object, e As EventArgs) Handles initExcel.Click
        Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If My.Computer.Network.IsAvailable Then
            Descargar()
            Me.Hide()
            Instrucciones1.Width = Screen.PrimaryScreen.Bounds.Width / 2
            Instrucciones1.Height = Screen.PrimaryScreen.Bounds.Height / 4.5
            Instrucciones1.Show()
        Else
            MsgBox("Computadora Sin Conexión a Internet.")
        End If
    End Sub

    Public Function Descargar()
        Dim remoteUri As String = "https://capacitacioninformatica.com/uploads/examenes/practica-1.xlsx"
        Dim fileName As String = "Examen-Excel.xlsx"
        Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim client = New WebClient()
        If (Not System.IO.Directory.Exists(Path + "/examen/")) Then
            System.IO.Directory.CreateDirectory(Path + "/examen/")
        End If
        client.DownloadFile(remoteUri, Path + "/examen/" + fileName)
    End Function
End Class

