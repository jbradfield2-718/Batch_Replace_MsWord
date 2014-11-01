Imports System
Imports System.IO

Public Class Form1
    Public excelfile As String
    Public wordfile As String
    Public numRows As Short
    Public numColumns As Short
    Public initial_change_flag As Byte
    Public find_vars(1) As String
    Public replace_vars(1, 1) As String
    Public filenames(1) As String

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        excelfile = OpenFileDialog1.FileName()

        If initial_change_flag = 0 Then
            initial_change_flag = 1
            RichTextBox1.Text = excelfile & " Successfully added as excel File!"
        Else
            RichTextBox1.AppendText(vbCrLf & excelfile & " Successfully added as excel File!")
        End If

    End Sub

    Private Sub excelButton_Click(sender As Object, e As EventArgs) Handles excelButton.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        initial_change_flag = 0
        RichTextBox1.ReadOnly = True
    End Sub

    Private Sub TextBoxNumRows_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub wordButton_Click(sender As Object, e As EventArgs) Handles wordButton.Click
        OpenFileDialog2.ShowDialog()
    End Sub

    Private Sub OpenFileDialog2_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        wordfile = OpenFileDialog2.FileName()
        If initial_change_flag = 0 Then
            initial_change_flag = 1
            RichTextBox1.Text = wordfile & " Successfully added as word template File!"
        Else
            RichTextBox1.AppendText(vbCrLf & wordfile & " Successfully added as word template File!")
        End If
    End Sub

    Private Sub runscriptButton_Click(sender As Object, e As EventArgs) Handles runscriptButton.Click
        Dim start_time, end_time As Double
        start_time = Microsoft.VisualBasic.DateAndTime.Timer

        RichTextBox1.Text = ""


        Read_Excel(excelfile)




        Const wdReplaceAll = 2
        Dim objWord As Object
        Dim objDoc As Object
        Dim objSelection As Object
        Dim path As String = Directory.GetCurrentDirectory()

        For i As Integer = 0 To (numRows - 2)
            objWord = CreateObject("Word.Application")
            objWord.Visible = False

            Try
                objDoc = objWord.Documents.Open(wordfile)
            Catch ex As Exception
                objWord.Quit()
                RichTextBox1.AppendText("Problem with Word file.  Invalid input.  Try again!" & vbLf)
                wordfile = ""
                Return
            End Try

            objSelection = objWord.Selection

            For j As Integer = 0 To numColumns - 1
                If replace_vars(i, j) = Nothing Then
                    replace_vars(i, j) = ""
                End If
                objSelection.Find.Text = find_vars(j)
                objSelection.Find.Forward = True
                objSelection.Find.MatchWholeWord = True
                objSelection.Find.MatchCase = True
                objSelection.Find.Replacement.Text = replace_vars(i, j)
                objSelection.Find.Execute(, , , , , , , , , , wdReplaceAll)
                RichTextBox1.AppendText("Searching for " & find_vars(j) & vbLf)
                RichTextBox1.AppendText("Replacing with " & replace_vars(i, j) & vbLf)
                Scroll_Down()
            Next j

            Try
                objDoc.SaveAs(path & "\" & filenames(i))
                RichTextBox1.AppendText("Created new file " & filenames(i) & "!" & vbLf)
                RichTextBox1.AppendText("Saved in directory " & path & "\" & filenames(i) & vbLf & vbLf)
                Scroll_Down()
                objDoc = Nothing
                objWord.Quit()
            Catch ex As Exception
                objWord.Quit()
                Exit For
            End Try

        Next i

        excelfile = ""
        wordfile = ""
        end_time = Microsoft.VisualBasic.DateAndTime.Timer - start_time

        RichTextBox1.AppendText(vbLf & "The total program runtime was " & end_time.ToString("G3") & " seconds.")
        Scroll_Down()
    End Sub

    Private Sub Read_Excel(file)
        Dim intRow As Integer
        Dim objExcel As Object
        Dim objWorkbook As Object
        Dim file_index As Byte = 0

        Try
            objExcel = CreateObject("Excel.Application")
            objWorkbook = objExcel.Workbooks.Open _
                (file)
        Catch ex As Exception
            objExcel.Quit()
            RichTextBox1.AppendText("Problem with input Excel file.  Invalid input.  Try again!" & vbLf)
            excelfile = ""
            Return
        End Try

        Try
            Get_Columns(objExcel)
            Get_Rows(objExcel)

        Catch ex As Exception
            RichTextBox1.Text = "Input excel file error, try again!"
            Return
        End Try

        Try
            ReDim find_vars(numColumns - 1)
            ReDim replace_vars(numRows - 1, numColumns - 1)
            ReDim filenames(numRows - 1)

        Catch ex As Exception
            ReDim find_vars(1)
            ReDim replace_vars(1, numColumns)
            ReDim filenames(1)
        End Try


        intRow = -1

        Do Until CStr(objExcel.Cells(intRow + 2, 1).Value) = ""

            For index As Integer = 0 To numColumns - 1

                If intRow = -1 Then
                    find_vars(index) = CStr(objExcel.Cells(intRow + 2, index + 1).Value)

                Else
                    replace_vars(intRow, index) = CStr(objExcel.Cells(intRow + 2, index + 1).Value)

                End If

            Next index

            If intRow <> -1 Then

                filenames(file_index) = CStr(objExcel.Cells(intRow + 2, numColumns + 1).Value)        ' Store filename in extra column
                file_index += 1

            End If

            intRow = intRow + 1

        Loop

        objExcel.Quit()

    End Sub

    Private Sub Scroll_Down()
        RichTextBox1.SelectionStart = RichTextBox1.Text.Length
        RichTextBox1.ScrollToCaret()
    End Sub

    Private Sub Get_Columns(objExcel)
        numColumns = 0
        Do Until CStr(objExcel.Cells(1, numColumns + 1).Value) = ""
            numColumns += 1
        Loop
    End Sub

    Private Sub Get_Rows(objExcel)
        numRows = 0
        Do Until CStr(objExcel.Cells(numRows + 1, 1).Value) = ""
            numRows += 1
        Loop
    End Sub

End Class

