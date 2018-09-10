Imports System.IO

Public Class Instrucciones2
    Dim word As New Microsoft.Office.Interop.Word.Application
    Dim doc As Microsoft.Office.Interop.Word.Document
    Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Dim check = 0
    Dim nav = 0
    Dim preguntas = New String() {"1.- Al título ""Las 7 maravillas"", aplica un espaciado expandido a 2 puntos.",
        "2.- En el primer párrafo del texto aplica un interlineado exacto a 20 puntos.",
        "3.- En el último párrafo del documento aplica una sangría izquierda a 4 cm y una derecha a 10 cm.",
        "4.- A la imagen contenida en el texto, aplica un ajuste de texto Estrecho.",
        "5.- Mueve la imagen a la izquierda del párrafo ""La gran pirámide de Guiza"".",
        "6.- En el párrafo que empieza con el texto *La Estatua de Zeus en Olimpia… crea dos columnas.",
        "7.- Quita el formato en el párrafo ""El coloso de rodas"".",
        "8.- Agrega a las propiedades de Excel, tu nombre en el campo del autor.",
        "9.- En las opciones de Excel, agrega a las palabras clave el texto ""Primer examen parcial"".",
        "10.- Centra el párrafo que inicia con el texto ""El hecho de que cinco ……..""."
    }

    Private Sub Instrucciones2_Load(sender As Object, e As EventArgs) Handles Me.Load
        doc = word.Documents.Open(Path + "\examen\Examen-Word.docx")
        word.Visible = True
        word.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMaximize
        doc.Activate()
        Me.TopMost = True
        Dim x As Integer = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) \ 2
        Dim y As Integer = Screen.PrimaryScreen.WorkingArea.Height - Me.Height
        Me.Location = New Point(x, y)
    End Sub
    Private Sub Button3_Click() Handles Button3.Click
        Dim puntos = 0
        Dim x As Integer

        'Pregunta 1
        x = doc.Content.Paragraphs(1).Range.Font.Spacing
        If x = 2 Then
            puntos = puntos + 1
        End If

        'Pregunta 2
        doc.Content.Paragraphs(3).Range.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly = 20
        x = doc.Content.Paragraphs(3).Range.ParagraphFormat.LineSpacing
        MessageBox.Show(x)


        MessageBox.Show("Calificación: " + puntos.ToString + "/10")
    End Sub

    Private Sub Button1_Click() Handles Button1.Click
        If nav = 9 Then
            nav = 0
        Else
            nav = nav + 1
        End If
        Pregunta(nav)
    End Sub
    Private Sub Button2_Click() Handles Button2.Click
        If nav = 0 Then
            nav = 9
        Else
            nav = nav - 1
        End If
        Pregunta(nav)
    End Sub
    Function Pregunta(i As Integer)
        Label1.Text = preguntas(i)
    End Function

    Function Test(ByRef sName As String) As Boolean
        Dim fs As FileStream
        Try
            fs = File.Open(sName, FileMode.Open, FileAccess.Read, FileShare.None)
            Test = False
        Catch ex As Exception
            Test = True
        End Try
    End Function

    Private Sub Cerrar_Click() Handles Cerrar.Click
        Dim e As System.Windows.Forms.FormClosingEventArgs
        Dim sender As Object
        Instrucciones2_FormClosing(sender, e)
    End Sub

    Private Sub Instrucciones2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If check = 0 Then
            If MessageBox.Show("¿Deseas cerrar el examen? Tus resultados no se guardarán.", "Salir del Examen", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                Dim bExist As Boolean
                bExist = Test("Examen-Word.docx")
                If bExist = True Then
                    Application.ExitThread()
                    doc.Close(False)
                    word.Quit()
                    System.IO.Directory.Delete(Path + "\examen\", True)
                End If
            End If
        ElseIf check = 1 Then
            Dim bExist As Boolean
            bExist = Test("Examen-Word.docx")
            If bExist = True Then
                Application.ExitThread()
                doc.Close(False)
                word.Quit()
                System.IO.Directory.Delete(Path + "\examen\", True)
            End If
        End If
    End Sub

    Private Sub Cerrar_Click(sender As Object, e As EventArgs) Handles Cerrar.Click

    End Sub
End Class