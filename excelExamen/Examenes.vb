Imports System.IO
Imports System.Net

Public Class excel
    Private Sub initExcel_Click(sender As Object, e As EventArgs) Handles initExcel.Click
        Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If My.Computer.Network.IsAvailable Then
            Descargar(1)
            Me.Hide()
            Instrucciones1.Width = Screen.PrimaryScreen.Bounds.Width / 1.8
            Instrucciones1.Height = Screen.PrimaryScreen.Bounds.Height / 4
            Instrucciones1.Show()
        Else
            MsgBox("Computadora Sin Conexión a Internet. Para realizar este Examen requiere conexión a Internet.")
        End If
    End Sub

    Private Sub initWord_Click(sender As Object, e As EventArgs) Handles initWord.Click
        Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If My.Computer.Network.IsAvailable Then
            Descargar(2)
            Me.Hide()
            Instrucciones2.Width = Screen.PrimaryScreen.Bounds.Width / 1.8
            Instrucciones2.Height = Screen.PrimaryScreen.Bounds.Height / 4
            Instrucciones2.Show()
        Else
            MsgBox("Computadora Sin Conexión a Internet. Para realizar este Examen requiere conexión a Internet.")
        End If
    End Sub

    Public Function Descargar(i As Integer)
        Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim client = New WebClient()
        If i = 1 Then
            Dim remoteUri As String = "https://capacitacioninformatica.com/uploads/examenes/practica-1.xlsx"
            Dim fileName As String = "Examen-Excel.xlsx"
            If (Not System.IO.Directory.Exists(Path + "/examen/")) Then
                System.IO.Directory.CreateDirectory(Path + "/examen/")
            End If
            client.DownloadFile(remoteUri, Path + "/examen/" + fileName)
        ElseIf i = 2 Then
            Dim remoteUri As String = "https://capacitacioninformatica.com/uploads/examenes/practica-1.docx"
            Dim fileName As String = "Examen-Word.docx"
            If (Not System.IO.Directory.Exists(Path + "/examen/")) Then
                System.IO.Directory.CreateDirectory(Path + "/examen/")
            End If
            client.DownloadFile(remoteUri, Path + "/examen/" + fileName)
        End If
    End Function
End Class

