Imports System.IO
Imports Microsoft.Win32


Public Class Sorting

#Region " Constants"
    Private Const DEFAULT_LENGTH As Integer = 20
#End Region
#Region "Enumerations"
    Private Enum BufferState As Integer
        UnInitialized = 0
        Initialized = 1
    End Enum
    Public Enum SelectedAlgorithm As Integer
        Bubble = 0
        Selection = 1
        Merge = 2
    End Enum
#End Region
#Region "Member Variables"
    Protected m_Buffer() As Integer
    Protected m_BufferLength As Integer
    Private m_BufferState As BufferState
    Public m_SelectedAlgorithm As SelectedAlgorithm
    Private m_Text As String
#End Region
#Region "Properties"

#End Region
#Region "Private Methods"

#End Region
#Region "Protected Methods"

#End Region
#Region "Public Methods"
    Public Sub New()
        'create a new obj 
        m_BufferLength = DEFAULT_LENGTH
        ReDim m_Buffer(m_BufferLength - 1)
        m_BufferState = BufferState.UnInitialized
        m_SelectedAlgorithm = SelectedAlgorithm.Bubble
        'm_Text = ""
    End Sub
    Public Function SetBufferLen(ByVal len As Integer) As Boolean
        If len > 0 Then
            m_BufferLength = len
            ReDim m_Buffer(m_BufferLength - 1)
            m_BufferState = BufferState.UnInitialized
            Return True
        Else
            Return False
        End If
    End Function
    Public Sub FillBuffer(ByVal array() As Integer, ByVal len As Integer)
        Dim i As Integer
        If len = m_BufferLength Then
            For i = 0 To m_BufferLength - 1
                m_Buffer(i) = array(i)
            Next
            m_BufferState = BufferState.Initialized
        Else
            MessageBox.Show("Warning: data length is invalid")
            m_BufferState = BufferState.UnInitialized
        End If
    End Sub

    Public Function Sorting() As Boolean
        'check buffer state
        If m_BufferState = BufferState.Initialized Then
            Dim k, n, i, j, t As Integer
            n = m_BufferLength
            m_Text = ""
            Select Case m_SelectedAlgorithm
                Case SelectedAlgorithm.Bubble
                    '冒泡排序法
                    For i = 0 To n - 1
                        For j = n - 1 To i + 1 Step -1
                            If m_Buffer(j - 1) > m_Buffer(j) Then    '相邻元素比较
                                t = m_Buffer(j)
                                m_Buffer(j) = m_Buffer(j - 1)
                                m_Buffer(j - 1) = t
                            End If
                        Next
                        '将排序结果显示在TextBox上
                        m_Text = m_Text + Str(m_Buffer(i))
                    Next
                Case SelectedAlgorithm.Selection
                    '选择排序法
                    For i = 0 To n - 1
                        k = i
                        For j = i + 1 To n - 1
                            If m_Buffer(k) > m_Buffer(j) Then k = j '找出最小值的下标
                        Next
                        '交换数组元素，使最小的元素排在第一位
                        t = m_Buffer(k) : m_Buffer(k) = m_Buffer(i) : m_Buffer(i) = t
                        '将排序结果显示在TextBox上
                        m_Text = m_Text + Str(m_Buffer(i))
                    Next
                Case SelectedAlgorithm.Merge
                    MsgBox("TBD sdfsd :P")

            End Select
            Return True
        Else
            MessageBox.Show("Warning: No data to be sorted")
            Return False

        End If

    End Function

    Public Sub BufferDumping(ByRef text As String)
        'Dim i As Integer
        text = ""
        'For i = 0 To m_BufferLength - 1
        'text = text + Str(m_Buffer(i))
        'Next
        text = m_Text
    End Sub


#End Region
#Region "Internal Class"

#End Region

End Class
