Imports SortingAlgorithm

Public Class Form1
    Dim Algorithm As Integer
    Private a(), ArrayLength As Integer
    Private mySort As New Sorting
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Default Settings:
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 2
        TextBox1.ReadOnly = True

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If CheckBox1.CheckState = CheckState.Checked Then
            Dim IsConvertSuccess As Boolean
            IsConvertSuccess = ConvertStringArrayToValueArray(TextBox1.Text, a)
            If IsConvertSuccess = False Then
                TextBox1.Text = ""
            Else
                TextBox3.Text = ""
                mySort.FillBuffer(a, ArrayLength)
                mySort.BufferDumping(TextBox3.Text)

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
        MsgBox("Sort App 1.2 Created by Bobby 2013-May-21")
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
        Dim strArrayElements() As String = text.Split(" ")
        Return strArrayElements.Length
        'Return text.Length
    End Function
    Private Function ConvertStringArrayToValueArray(ByVal text As String, ByRef intArrayElements() As Integer) As Boolean
        Dim i, intLength As Integer, strArrayElements() As String = text.Split(" ")
        intLength = strArrayElements.Length
        ReDim intArrayElements(intLength - 1)
        For i = 0 To intLength - 1
            If IsNumeric(strArrayElements(i)) Then
                intArrayElements(i) = Val(strArrayElements(i))
            Else
                MessageBox.Show("Invalid value,please input numbers splited with a space ")
                Return False
            End If
        Next
        ArrayLength = intLength
        Return True
    End Function
End Class