
Public Class Form1
    Dim Algorithm As Integer, IsArrayFilled As Boolean
    Private a(), ArrayLength As Integer
    'Private i, arayInput(9) As Integer
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Default Settings:
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 2
        'ArrayLength = 20 'default length: 20
        'Algorithm = 1  ' defaul algorithm:Bubble
        IsArrayFilled = False
        TextBox1.ReadOnly = True

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If CheckBox1.CheckState = CheckState.Checked Then
            Dim IsConvertSuccess As Boolean, i As Integer
            IsConvertSuccess = ConvertStringArrayToValueArray(TextBox1.Text, a)
            If IsConvertSuccess = False Then
                TextBox1.Text = ""
                IsArrayFilled = False
            Else
                IsArrayFilled = True
                TextBox3.Text = ""
                For i = 0 To ArrayLength - 1
                    TextBox3.Text = TextBox3.Text & " " & a(i)
                Next
            End If
        End If
 

        '排序执行
        If IsArrayFilled = True Then
            Dim k, n, i, j, t As Integer
            Randomize()
            n = ArrayLength
            TextBox2.Text = ""
            Select Case Algorithm
                Case 1 'default algorithm: bubble sorting
                    '冒泡排序法
                    For i = 0 To n - 1
                        For j = n - 1 To i + 1 Step -1
                            If a(j - 1) > a(j) Then    '相邻元素比较
                                t = a(j)
                                a(j) = a(j - 1)
                                a(j - 1) = t
                            End If
                        Next
                        '将排序结果显示在TextBox上
                        TextBox2.Text = TextBox2.Text + Str(a(i))
                    Next
                    'TextBox2.Text = TextBox2.Text + Str(a(i))
                Case 2 'Selection sorting
                    '选择排序法
                    For i = 0 To n - 1
                        k = i
                        For j = i + 1 To n - 1
                            If a(k) > a(j) Then k = j '找出最小值的下标
                        Next
                        '交换数组元素，使最小的元素排在第一位
                        t = a(k) : a(k) = a(i) : a(i) = t
                        '将排序结果显示在TextBox上
                        TextBox2.Text = TextBox2.Text + Str(a(i))
                    Next
                    'TextBox2.Text = TextBox2.Text + Str(a(i))
                Case 3 ' Merge sorting
                    MsgBox("TBD :P")

            End Select

        Else
            MsgBox("给我数据先 Feed me some number first :P")

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
                IsArrayFilled = False
            Case 1
                ArrayLength = 10
                ReDim a(9)
                IsArrayFilled = False
            Case 2
                ArrayLength = 20
                ReDim a(19)
                IsArrayFilled = False
        End Select
        For i = 0 To ArrayLength - 1
            a(i) = Int(Rnd() * 100) + 1
            TextBox3.Text = TextBox3.Text + Str(a(i))
        Next
        IsArrayFilled = True
    End Sub
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Sort App 1.1 Created by Bobby 2013-May-20")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        'MessageBox.Show(ComboBox1.GetItemText(ComboBox1.SelectedItem))

        Select Case ComboBox1.GetItemText(ComboBox1.SelectedItem)
            Case "Bubble Sort"
                Algorithm = 1
                'MsgBox("bubble!!!")
            Case "Selection Sort"
                Algorithm = 2
                'MsgBox("selection!!!")
            Case "Merge Sort"
                Algorithm = 3
                'MsgBox("merge!!!")
        End Select
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

        Select Case ComboBox2.GetItemText(ComboBox2.SelectedItem)
            Case "6"
                ArrayLength = 6
                ReDim a(5)
                IsArrayFilled = False
                'MsgBox("6!!!")
            Case "10"
                ArrayLength = 10
                ReDim a(9)
                IsArrayFilled = False
                'MsgBox("10!!!")
            Case "20"
                ArrayLength = 20
                ReDim a(19)
                IsArrayFilled = False
                'MsgBox("20!!!")
        End Select
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            TextBox1.ReadOnly = False
            GeneratRandomArray.Enabled = False
            IsArrayFilled = False

        Else
            TextBox1.ReadOnly = True
            GeneratRandomArray.Enabled = True
            IsArrayFilled = False

        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim intArraySize As Integer = CountArrayLength(TextBox1.Text)
        'display the results
        Label7.Text = " Array size: " & intArraySize
    End Sub
    ''' <summary>
    ''' Count the characters in a block of text
    ''' </summary>
    ''' param name: text the tring
    ''' <remarks></remarks>
    Private Function CountArrayLength(ByVal text As String)
        Dim strArrayElements() As String = text.Split(" ")
        Return strArrayElements.Length
        'Return text.Length
    End Function
    Private Function ConvertStringArrayToValueArray(ByVal text As String, ByRef intArrayElements() as Integer) As Boolean
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