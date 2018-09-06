Imports System.IO

Public Class Instrucciones1
    Dim excel As New Microsoft.Office.Interop.Excel.Application
    Dim wb As Microsoft.Office.Interop.Excel.Workbook
    Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

    Private Sub Instrucciones1_Load(sender As Object, e As EventArgs) Handles Me.Load
        wb = excel.Workbooks.Open(Path + "\examen\Examen-Excel.xlsx")
        excel.Visible = True
        excel.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized
        wb.Activate()
        Me.TopMost = True
        Dim x As Integer = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2
        Dim y As Integer = Screen.PrimaryScreen.WorkingArea.Height - Me.Height
        Me.Location = New Point(x, y)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim punto1 = 0
        sheet = wb.ActiveSheet
        For i As Integer = 4 To 10
            If (sheet.Range("E" + i).Formula) Then
                punto1 = punto1 + 1
            End If
        Next
        MessageBox.Show("Califiacion: " + Format(punto1 / 7, "0.0"))
    End Sub

    Function Test(ByRef sName As String) As Boolean
        Dim fs As FileStream
        Try
            fs = File.Open(sName, FileMode.Open, FileAccess.Read, FileShare.None)
            Test = False
        Catch ex As Exception
            Test = True
        End Try
    End Function

    Private Sub Instrucciones1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("¿Deseas cerrar el examen?", "Salir", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Dim bExist As Boolean
            bExist = Test("Examen-Excel.xlsx")
            If bExist = True Then
                Application.ExitThread()
                wb.Close(False)
                excel.Quit()
                System.IO.Directory.Delete(Path + "\examen\", True)
            End If
        Else
            e.Cancel = True
        End If
    End Sub
End Class