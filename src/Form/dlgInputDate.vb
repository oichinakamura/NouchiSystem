Public Class dlgInputDate
    Public Sub New(ByVal sMessage As String, ByVal dtDefault As DateTime, Optional ByVal sTitle As String = "")
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        Me.Label1.Text = sMessage
        If sTitle.Length > 0 Then
            Me.Text = sTitle
        End If
        DateTimePicker1.Value = dtDefault

    End Sub

    Public Property ResultDate As DateTime
        Set(value As DateTime)

        End Set
        Get
            Return DateTimePicker1.Value
        End Get
    End Property

    Private Sub OK_Button_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
        Me.DateTimePicker1.Format = DateTimePickerFormat.Long
    End Sub
End Class