Imports System.IO

Public Class Instrucciones1
    Dim excel As New Microsoft.Office.Interop.Excel.Application
    Dim wb As Microsoft.Office.Interop.Excel.Workbook
    Dim Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    Dim nav = 0
    Dim check = 0
    Dim preguntas = New String() {"1.- Calcular el importe ""Cantidad * Precio"".",
        "2.- Colocar formato de millares a la columna de Precio para que las cantidades tengan 2 decimales y separador de miles.",
        "3.- Calcular el precio con descuento quitando el 25% del importe, y colocar el nuevo importe con el descuento.",
        "4.- Calcular el Impuesto Trasladado, calculando el 18% del precio con descuento.",
        "5.- Calcular el Impuesto del Aeropuerto calculando el 6% de Precio del producto.",
        "6.- Calcular el Precio final del producto Tomando en cuenta el Precio con descuento, menos el impuesto trasladado y el Impuesto Aeropuerto.",
        "7.- En la fila 11 agrega una fila para calcular los totales de cada columna  (C a la I).",
        "8.- Coloca el estilo de Celda""Enfasis 6"" a las  celdas de  los títulos de la fila 3 y a los totales de la fila 11.",
        "9.- Agrega bordes de color ""Verde, Enfasis 6"" a la tabla del Rango A3:I10.",
        "10.- Inserta una columna antes del precio final, reduce el ancho de la columna a 3. Agregale un color de fondo."
    }

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

    Private Sub Button2_Click() Handles Button2.Click
        If nav = 9 Then
            nav = 0
        Else
            nav = nav + 1
        End If
        Pregunta(nav)
    End Sub
    Private Sub Button3_Click() Handles Button3.Click
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet
        sheet = wb.ActiveSheet
        Dim punto1 = 0
        Dim punto2 = 0
        Dim punto3 = 0
        Dim punto4 = 0
        Dim punto5 = 0
        Dim punto6 = 0
        Dim punto7 = 0
        Dim punto8 = 0
        Dim punto9 = 0
        Dim punto10 = 0
        Dim ban = 0
        Dim f As String
        Dim c As String
        Dim v As Double
        Dim x As Double
        Dim y As Double
        Dim z As Double
        Dim t(7) As Double

        'For i As Integer = 4 To 10
        '    f = sheet.Cells(i, 5).formula
        '    c = "=c" & i & "*d" & i
        '    If String.Compare(f, c) = 0 Then
        '        punto1 = punto1 + 1
        '    End If
        'Next

        'Pregunta 1
        For i As Integer = 4 To 10
            If sheet.Cells(i, 5).Value IsNot Nothing Then
                v = Double.Parse(sheet.Cells(i, 5).Value)
                x = Double.Parse(sheet.Cells(i, 3).Value)
                y = Double.Parse(sheet.Cells(i, 4).Value)
                If v = (x * y) Then
                    punto1 = punto1 + 1
                End If
            Else
                ban = 1
            End If
        Next

        'Pregunta 2
        For i As Integer = 4 To 10
            If sheet.Cells(i, 4).Value IsNot Nothing Then
                f = sheet.Cells(i, 4).Style.NumberFormat
                If (f.Contains("#,##0.00")) Then
                    punto2 = punto2 + 1
                End If
            Else
                ban = 1
            End If
        Next

        'Pregunta 3
        For i As Integer = 4 To 10
            If sheet.Cells(i, 6).Value IsNot Nothing Then
                v = Double.Parse(sheet.Cells(i, 6).Value)
                x = Double.Parse(sheet.Cells(i, 5).Value)
                If v = (x * 0.75) Then
                    punto3 = punto3 + 1
                End If
            Else
                ban = 1
            End If
        Next

        'Pregunta 4
        For i As Integer = 4 To 10
            If sheet.Cells(i, 7).Value IsNot Nothing Then
                v = sheet.Cells(i, 7).Value
                x = sheet.Cells(i, 6).Value
                If v = (x * 0.18) Then
                    punto4 = punto4 + 1
                End If
            Else
                ban = 1
            End If
        Next

        'Pregunta 5
        For i As Integer = 4 To 10
            If sheet.Cells(i, 8).Value IsNot Nothing Then
                v = sheet.Cells(i, 8).Value
                x = sheet.Cells(i, 4).Value
                If v = (x * 0.06) Then
                    punto5 = punto5 + 1
                End If
            Else
                ban = 1
            End If
        Next

        'Pregunta 6
        Dim m = 0
        For i As Integer = 4 To 11
            If sheet.Cells(i, 10).Value IsNot Nothing Then
                v = sheet.Cells(i, 10).Value
                x = sheet.Cells(i, 6).Value
                y = sheet.Cells(i, 7).Value
                z = sheet.Cells(i, 8).Value
                If v = (x - y - z) Then
                    punto6 = punto6 + 1
                    t(m) = v
                    m = m + 1
                End If
            Else
                ban = 1
            End If
        Next

        'Pregunta 7
        If sheet.Cells(11, 3).Value IsNot Nothing Then
            x = 0
            v = sheet.Cells(11, 3).Value
            For i As Integer = 4 To 10
                x = x + sheet.Cells(i, 3).Value
            Next
            If v = x Then
                punto7 = punto7 + 1
            End If
        Else
            ban = 1
        End If

        If sheet.Cells(11, 4).Value IsNot Nothing Then
            x = 0
            v = sheet.Cells(11, 4).Value
            For i As Integer = 4 To 10
                x = x + sheet.Cells(i, 4).Value
            Next
            If v = x Then
                punto7 = punto7 + 1
            End If
        Else
            ban = 1
        End If

        If sheet.Cells(11, 5).Value IsNot Nothing Then
            x = 0
            v = sheet.Cells(11, 5).Value
            For i As Integer = 4 To 10
                x = x + sheet.Cells(i, 5).Value
            Next
            If v = x Then
                punto7 = punto7 + 1
            End If
        Else
            ban = 1
        End If

        If sheet.Cells(11, 6).Value IsNot Nothing Then
            x = 0
            v = sheet.Cells(11, 6).Value
            For i As Integer = 4 To 10
                x = x + sheet.Cells(i, 6).Value
            Next
            If v = x Then
                punto7 = punto7 + 1
            End If
        Else
            ban = 1
        End If

        If sheet.Cells(11, 7).Value IsNot Nothing Then
            x = 0
            v = sheet.Cells(11, 7).Value
            For i As Integer = 4 To 10
                x = x + sheet.Cells(i, 7).Value
            Next
            If v = x Then
                punto7 = punto7 + 1
            End If
        Else
            ban = 1
        End If

        If sheet.Cells(11, 8).Value IsNot Nothing Then
            x = 0
            v = sheet.Cells(11, 8).Value
            For i As Integer = 4 To 10
                x = x + sheet.Cells(i, 8).Value
            Next
            If v = x Then
                punto7 = punto7 + 1
            End If
        Else
            ban = 1
        End If

        If sheet.Cells(11, 10).Value IsNot Nothing Then
            x = 0
            v = sheet.Cells(11, 10).Value
            For i As Integer = 4 To 10
                x = x + sheet.Cells(i, 10).Value
            Next
            If v = x Then
                punto7 = punto7 + 1
            End If
        Else
            ban = 1
        End If

        'Pregunta 8
        For i As Integer = 1 To 10
            f = sheet.Cells(3, i).Font.ColorIndex
            c = sheet.Cells(3, i).Interior.ColorIndex
            If f = 2 And c = 50 Then
                punto8 = punto8 + 1
            End If
        Next

        For i As Integer = 3 To 10
            f = sheet.Cells(11, i).Font.ColorIndex
            c = sheet.Cells(11, i).Interior.ColorIndex
            If f = 2 And c = 50 Then
                punto8 = punto8 + 1
            End If
        Next

        'Pregunta 9
        For i As Integer = 3 To 10
            For j As Integer = 1 To 10
                f = sheet.Cells(i, j).Borders.ColorIndex
                If f = 50 Then
                    punto9 = punto9 + 1
                End If
            Next
        Next

        'Pregunta 10
        f = sheet.Cells(3, 10).Text
        If f.Equals("Precio final ") Then
            punto10 = punto10 + 1
        End If
        m = 0
        For j As Integer = 4 To 11
            v = sheet.Cells(j, 10).Value
            If v = t(m) Then
                punto10 = punto10 + 1
            End If
            m = m + 1
        Next
        f = sheet.Cells(3, 9).ColumnWidth
        c = sheet.Cells(4, 9).Interior.ColorIndex
        If f = 3 And c <> -4142 Then
            punto10 = punto10 + 1
        End If


        'Revisión Final
        If ban = 1 Then
            MsgBox("Tienes preguntas sin resolver. Revisa tu examen!",, "Error!")
        Else
            'MsgBox("Punto 1: " + Format((punto1 / 7) * 10, "0.0") + Environment.NewLine +
            '            "Punto 2: " + Format((punto2 / 7) * 10, "0.0") + Environment.NewLine +
            '            "Punto 3: " + Format((punto3 / 7) * 10, "0.0") + Environment.NewLine +
            '            "Punto 4: " + Format((punto4 / 7) * 10, "0.0") + Environment.NewLine +
            '            "Punto 5: " + Format((punto5 / 7) * 10, "0.0") + Environment.NewLine +
            '            "Punto 6: " + Format((punto6 / 8) * 10, "0.0") + Environment.NewLine +
            '            "Punto 7: " + Format((punto7 / 7) * 10, "0.0") + Environment.NewLine +
            '            "Punto 8: " + Format((punto8 / 18) * 10, "0.0") + Environment.NewLine +
            '            "Punto 9: " + Format((punto9 / 80) * 10, "0.0") + Environment.NewLine +
            '            "Punto 10: " + Format((punto10 / 10) * 10, "0.0") + Environment.NewLine,, "RESULTADOS")

            'Dim nombre As String
            'Try
            '    Do While String.Compare(nombre, "")
            '        nombre = InputBox("Ingresa tu Nombre.", "Datos Personales - CI")
            '    Loop
            'Catch
            '    MsgBox("Ingresa tu Nombre.")
            'End Try

            'Dim matricula As String
            'Try
            '    Do While String.Compare(matricula, "")
            '        matricula = InputBox("Ingresa tu Matricula.", "Datos Personales - CI")
            '    Loop
            'Catch
            '    MsgBox("Ingresa tu Matricula.")
            'End Try
            Dim nombre As String
            Dim matricula As String
            nombre = InputBox("Ingresa tu Nombre.", "Datos Personales - CI")
            matricula = InputBox("Ingresa tu Matricula.", "Datos Personales - CI")

            Do While nombre = ""
                nombre = InputBox("Ingresa tu Nombre.", "Datos Personales - CI")
            Loop

            Do While matricula = ""
                matricula = InputBox("Ingresa tu Matricula.", "Datos Personales - CI")
            Loop

            Dim webAddress As String = "https://capacitacioninformatica.com/examenes/" + nombre + "/" + matricula + "/" + Format((punto1 / 7) * 10, "0.0") + "/" + Format((punto2 / 7) * 10, "0.0") + "/" + Format((punto3 / 7) * 10, "0.0") + "/" + Format((punto4 / 7) * 10, "0.0") + "/" + Format((punto5 / 7) * 10, "0.0") + "/" + Format((punto6 / 8) * 10, "0.0") + "/" + Format((punto7 / 7) * 10, "0.0") + "/" + Format((punto8 / 18) * 10, "0.0") + "/" + Format((punto9 / 80) * 10, "0.0") + "/" + Format((punto10 / 10) * 10, "0.0")
            Process.Start(webAddress)
            check = 1
            Cerrar_Click()
        End If
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

    Private Sub Cerrar_Click() Handles Cerrar.Click
        Dim e As System.Windows.Forms.FormClosingEventArgs
        Dim sender As Object
        Instrucciones1_FormClosing(sender, e)
    End Sub
    Private Sub Instrucciones1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If check = 0 Then
            If MessageBox.Show("¿Deseas cerrar el examen? Tus resultados no se guardarán.", "Salir del Examen", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                Dim bExist As Boolean
                bExist = Test("Examen-Excel.xlsx")
                If bExist = True Then
                    Application.ExitThread()
                    wb.Close(False)
                    excel.Quit()
                    System.IO.Directory.Delete(Path + "\examen\", True)
                End If
            End If
        ElseIf check = 1 Then
            Dim bExist As Boolean
            bExist = Test("Examen-Excel.xlsx")
            If bExist = True Then
                Application.ExitThread()
                wb.Close(False)
                excel.Quit()
                System.IO.Directory.Delete(Path + "\examen\", True)
            End If
        End If
    End Sub
End Class