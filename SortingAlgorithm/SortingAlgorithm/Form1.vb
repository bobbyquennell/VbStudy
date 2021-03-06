Imports SortingAlgorithm

Public Class Form1
    Dim Algorithm As Integer
    Private a(), ArrayLength As Integer
    Private Const DEFAULT_LENGTH_INDEX As Integer = 2 '20'
    Private Const DEFAULT_ALGORITHM_INDEX As Integer = 0 'Bubble'
    Private mySort As New Sorting
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Default Settings:
        ComboBox1.SelectedIndex = DEFAULT_ALGORITHM_INDEX
        ComboBox2.SelectedIndex = DEFAULT_LENGTH_INDEX
        TextBox1.ReadOnly = True

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If CheckBox1.CheckState = CheckState.Checked Then
                Dim Len As Integer
                Len = ConvertStringToValue(TextBox1.Text, a)
                TextBox3.Text = ""
                mySort.SetBufferLen(Len)
                mySort.FillBuffer(a, Len)
            End If
            mySort.Sorting()
            ' Display the result
            mySort.BufferDumping(TextBox2.Text)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub GeneratRandomArray_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GeneratRandomArray.Click
        '生成0～100的随机整数，填充到buffer中等待排序
        Dim i As Integer
        TextBox3.Text = ""
        Select Case ComboBox2.SelectedIndex
            Case 0
                ArrayLength = 6
                ReDim a(5)
            Case 1
                ArrayLength = 10
                ReDim a(9)
            Case 2
                ArrayLength = 20
                ReDim a(19)
        End Select
        For i = 0 To ArrayLength - 1
            a(i) = Int(Rnd() * 100) + 1
            TextBox3.Text = TextBox3.Text + Str(a(i))
        Next
        mySort.FillBuffer(a, ArrayLength)
    End Sub
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Sort App 1.4 Created by Bobby 2013-May-23")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Select Case ComboBox1.GetItemText(ComboBox1.SelectedItem)
            Case "Bubble Sort"
                Algorithm = 1
                mySort.m_SelectedAlgorithm = Sorting.SelectedAlgorithm.Bubble
            Case "Selection Sort"
                Algorithm = 2
                mySort.m_SelectedAlgorithm = Sorting.SelectedAlgorithm.Selection
            Case "Merge Sort"
                Algorithm = 3
                mySort.m_SelectedAlgorithm = Sorting.SelectedAlgorithm.Merge
        End Select
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

        Select Case ComboBox2.GetItemText(ComboBox2.SelectedItem)
            Case "6"
                ArrayLength = 6
                ReDim a(5)
            Case "10"
                ArrayLength = 10
                ReDim a(9)
            Case "20"
                ArrayLength = 20
                ReDim a(19)
        End Select
        mySort.SetBufferLen(ArrayLength)
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            TextBox1.ReadOnly = False
            GeneratRandomArray.Enabled = False
            ComboBox2.Enabled = False
        Else
            TextBox1.ReadOnly = True
            GeneratRandomArray.Enabled = True
            ComboBox2.Enabled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim intArraySize As Integer
        Try
            If Not TextBox1.Text = Nothing Then
                intArraySize = CountArrayLength(TextBox1.Text)
                'display the results
                Label7.Text = " Array size: " & intArraySize
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function CountArrayLength(ByVal text As String) As Integer

        Dim Length As Integer = 0

        If text = Nothing Then
            Throw New ArgumentException("empty array")
        End If

        FormatText(text)
        If Not text = Nothing Then
            Dim count() As String = text.Split(" ")
            Length = count.Length
        End If
        Return Length

    End Function
    Private Function ConvertStringToValue(ByVal text As String, ByRef intArrayElements() As Integer) As Integer
        FormatText(text)
        If text = Nothing Then
            Throw New ArgumentException("empty input, please feed me some numbers splited with a space or a ','")
        End If
        Dim i, intLength As Integer, strArrayElements() As String = text.Split(" ")
        intLength = strArrayElements.Length
        ReDim intArrayElements(intLength - 1)
        For i = 0 To intLength - 1
            If IsNumeric(strArrayElements(i)) Then
                intArrayElements(i) = Val(strArrayElements(i))
            Else
                Throw New ArgumentException("invalid input, plsease input numbers splited with a space or a ','")
            End If
        Next
        Return intLength
    End Function
    Private Sub FormatText(ByRef text As String)
        text = Trim(text)
        text = Replace(text, ",", " ")
        text = Replace(text, "    ", " ")
        text = Replace(text, "   ", " ")
        text = Replace(text, "  ", " ")
    End Sub

    Private Sub ClrInputBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClrInputBox.Click
        TextBox1.Clear()
    End Sub
End Class