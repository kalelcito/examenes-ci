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
        sheet = wb.ActiveSheet
        Dim punto1 = 0
        Dim punto2 = 0
        Dim f As String
        Dim c As String

        For i As Integer = 4 To 10
            f = sheet.Cells(i, 5).Formula
            c = "=C" & i & "*D" & i
            If String.Compare(f, c) = 0 Then
                punto1 = punto1 + 1
            End If
        Next
        For i As Integer = 4 To 10
            f = sheet.Cells(i, 5).Style.NumberFormat
            If (f.Contains("#,##0.00")) Then
                punto2 = punto2 + 1
            End If
        Next
        MessageBox.Show("RESULTADOS" + Environment.NewLine + Environment.NewLine +
                        "Punto 1: " + Format((punto1 / 7) * 10, "0.0") + Environment.NewLine +
                        "Punto 2: " + Format((punto2 / 7) * 10, "0.0") + Environment.NewLine +
                        "Punto 3: ")
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