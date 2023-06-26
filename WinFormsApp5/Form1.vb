Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar

Public Class Form1
    Dim alf As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim dt As New DataTable ' — создание таблицы данных
    Dim ds As New DataSet ' — создание набора данных
    Dim column As Integer = 0
    Dim cicle As Integer = 0
    Dim matrixM As String = ""
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label1.Text = "Размерность матриц:"
        Button1.Text = "Добавить столбец"
        Button2.Text = "Добавить строку"
        Button3.Text = "Заполнить матрицы"
        Button6.Text = "Добавить 25*25"
        Button8.Text = "Сохранить"
        If IO.File.Exists("C:\Users\Shadrina Maria\Desktop\tabl.xml") = False Then
            ' Если XML-файла НЕТ:
            DataGridView1.DataSource = dt
            'Добавить объект dt в DataSet:
            ds.Tables.Add(dt)
        Else 'Если XML-файл ЕCТЬ:
            ds.ReadXml("C:\Users\Shadrina Maria\Desktop\tabl.xml")
            DataGridView1.DataMember = "Элемент"
            DataGridView1.DataSource = ds
        End If
    End Sub
    Private Sub Write(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        dt.TableName = "Название таблицы"
        ds.WriteXml("C:\Users\Shadrina Maria\Desktop\tabl.xml")
    End Sub
    Private Sub AddColumn(sender As Object, e As EventArgs) Handles Button1.Click
        If column = 26 Then
            cicle += 1
            column = 0
        Else
            If cicle = 0 Then
                dt.Columns.Add(alf(column))
            Else
                dt.Columns.Add(alf(column) + Str(cicle))
            End If
            column = column + 1
        End If
    End Sub
    Private Sub AddRow(sender As Object, e As EventArgs) Handles Button2.Click
        dt.Rows.Add()
    End Sub

    Private Sub Eve(sender As Object, e As EventArgs) Handles Button3.Click
        Dim nM1 = CInt(TextBox1.Text)
        Dim nM2 = CInt(TextBox2.Text)
        Dim nM3 = CInt(TextBox3.Text)

        DataGridView1.Rows(0).Cells(0).Value = "Матрица A:"
        FillRandom(0, nM1 - 1, 1, nM1, nM1)
        Dim maxA = MaxNormMatrix(0, nM1 - 1, 1, nM1, nM1)
        DataGridView1.Rows(0).Cells(1).Value = "Максимум = " + Str(maxA)

        DataGridView1.Rows(nM1 + 1).Cells(0).Value = "Матрица B:"
        FillRandom(0, nM2 - 1, nM1 + 2, nM1 + nM2 + 1, nM2)
        Dim maxB = MaxNormMatrix(0, nM2 - 1, nM1 + 2, nM1 + nM2 + 1, nM2)
        DataGridView1.Rows(nM1 + 1).Cells(1).Value = "Максимум = " + Str(maxB)

        DataGridView1.Rows(nM1 + nM2 + 2).Cells(0).Value = "Матрица C:"
        FillRandom(0, nM3 - 1, nM1 + nM2 + 3, nM1 + nM2 + nM3 + 2, nM3)
        Dim maxC = MaxNormMatrix(0, nM3 - 1, nM1 + nM2 + 3, nM1 + nM2 + nM3 + 2, nM3)
        DataGridView1.Rows(nM1 + nM2 + 2).Cells(1).Value = "Максимум = " + Str(maxC)



        If maxA <= maxB And maxA <= maxC Then
            matrixM = MatrixStr(0, nM1 - 1, 1, nM1)
        ElseIf maxB <= maxA And maxB <= maxC Then
            matrixM = MatrixStr(0, nM2 - 1, nM1 + 2, nM1 + nM2 + 1)
        ElseIf maxC <= maxA And maxC <= maxB Then
            matrixM = MatrixStr(0, nM3 - 1, nM1 + nM2 + 3, nM1 + nM2 + nM3 + 2)
        End If

        PrintDialog1.Document = PrintDocument1
        If (PrintDialog1.ShowDialog() = DialogResult.OK) Then
            PrintDocument1.Print()
        End If
    End Sub
    Private Sub FillRandom(col1 As Integer, col2 As Integer, row1 As Integer, row2 As Integer, stepAbs As Integer)
        Randomize()
        For i = row1 To row2
            For j = col1 To col2
                Dim randomValue = CInt(-200 * Rnd() + 100)
                DataGridView1.Rows(i).Cells(j).Value = randomValue
                DataGridView1.Rows(i).Cells(j + stepAbs).Value = Math.Abs(randomValue)
                DataGridView1.Rows(i).Cells(j + stepAbs).Style.BackColor = Color.LightGray
            Next j
        Next i
    End Sub
    Private Function MaxNormMatrix(col1 As Integer, col2 As Integer, row1 As Integer, row2 As Integer, stepAbs As Integer) As Integer
        Dim maxCount = DataGridView1.Rows(row1).Cells(col1 + stepAbs).Value
        For i = row1 To row2
            For j = col1 To col2
                Dim n = DataGridView1.Rows(i).Cells(j + stepAbs).Value
                If n > maxCount Then
                    maxCount = n
                End If
            Next j
        Next i
        'Закрашивание клетки с максимумом матрицы
        For i = row1 To row2
            For j = col1 To col2
                If DataGridView1.Rows(i).Cells(j + stepAbs).Value = maxCount Then
                    DataGridView1.Rows(i).Cells(j + stepAbs).Style.BackColor = Color.LightGreen
                End If
            Next j
        Next i
        Return maxCount
    End Function

    Private Function MatrixStr(col1 As Integer, col2 As Integer, row1 As Integer, row2 As Integer) As String
        For i = row1 To row2
            For j = col1 To col2
                matrixM += Str(DataGridView1.Rows(i).Cells(j).Value) + " "
            Next j
            matrixM += vbCrLf
        Next i
        Return matrixM
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button6.Click
        For i = 0 To 25
            dt.Rows.Add()
            If column = 26 Then
                cicle += 1
                column = 0
            Else
                If cicle = 0 Then
                    dt.Columns.Add(alf(column))
                Else
                    dt.Columns.Add(alf(column) + Str(cicle))
                End If
                column = column + 1
            End If
        Next i
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim printFont = New Font("Times New Roman", 14)
        e.Graphics.DrawString(matrixM, printFont, Brushes.Black, e.MarginBounds, New StringFormat())
        e.HasMorePages = False
    End Sub
End Class
