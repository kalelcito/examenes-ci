Imports System.IO

Public Class Instrucciones1
    Private Sub Instrucciones1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("¿Deseas cerrar el examen?", "Salir", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Dim bExist As Boolean
            bExist = Test("Examen-Excel.xlsx")
            If bExist = True Then
            End If
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub Instrucciones1_Load(sender As Object, e As EventArgs) Handles Me.Load
        MessageBox.Show("abrió")

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
End Class