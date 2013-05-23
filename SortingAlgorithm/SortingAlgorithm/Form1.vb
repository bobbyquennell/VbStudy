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
        If CheckBox1.CheckState = CheckState.Checked Then
            Dim Len As Integer
            Len = ConvertStringArrayToValueArray(TextBox1.Text, a)
            If Len = 0 Then
                MessageBox.Show("Invalid value,please input numbers splited with a space or a ',' ")
            Else
                TextBox3.Text = ""
                mySort.SetBufferLen(Len)
                mySort.FillBuffer(a, Len)
            End If
        End If
        If mySort.Sorting() = True Then
            ' Display the result
            mySort.BufferDumping(TextBox2.Text)
        End If
    End Sub

    Private Sub GeneratRandomArray_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GeneratRandomArray.Click
        '给数组赋值，0～100中的随机整数
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
        MsgBox("Sort App 1.3 Created by Bobby 2013-May-22")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        'MessageBox.Show(ComboBox1.GetItemText(ComboBox1.SelectedItem))

        Select Case ComboBox1.GetItemText(ComboBox1.SelectedItem)
            Case "Bubble Sort"
                Algorithm = 1
                mySort.m_SelectedAlgorithm = Sorting.SelectedAlgorithm.Bubble
                'MsgBox("bubble!!!")
            Case "Selection Sort"
                Algorithm = 2
                mySort.m_SelectedAlgorithm = Sorting.SelectedAlgorithm.Selection
                'MsgBox("selection!!!")
            Case "Merge Sort"
                Algorithm = 3
                mySort.m_SelectedAlgorithm = Sorting.SelectedAlgorithm.Merge
                'MsgBox("merge!!!")
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
        Else
            TextBox1.ReadOnly = True
            GeneratRandomArray.Enabled = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim intArraySize As Integer = CountArrayLength(TextBox1.Text)
        'display the results
        Label7.Text = " Array size: " & intArraySize
    End Sub

    Private Function CountArrayLength(ByVal text As String)
        If text = Nothing Then
            MessageBox.Show(" error: NULL data to be sorted ")
        End If
        Try
            text = Trim(text)
            text = Replace(text, ",", " ")
            text = Replace(text, "   ", " ")
            text = Replace(text, "  ", " ")

        Catch

        End Try
        Dim count() As String = text.Split(" ")
        Return count.Length
        'Return text.Length
    End Function
    Private Function ConvertStringArrayToValueArray(ByVal text As String, ByRef intArrayElements() As Integer) As Integer
        text = Trim(text)
        text = Replace(text, ",", " ")
        text = Replace(text, "    ", " ")
        text = Replace(text, "   ", " ")
        text = Replace(text, "  ", " ")
        Dim i, intLength As Integer, strArrayElements() As String = text.Split(" ")
        intLength = strArrayElements.Length
        ReDim intArrayElements(intLength - 1)
        For i = 0 To intLength - 1
            If IsNumeric(strArrayElements(i)) Then
                intArrayElements(i) = Val(strArrayElements(i))
            Else
                Return 0
            End If
        Next
        Return intLength
    End Function
End Class