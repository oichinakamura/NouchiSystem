Imports System.Windows.Forms

Public Class dlgInputBWDate

    Public RestultDate As DateTime
    Public Restult年数 As Integer = 0

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If DateTimePicker1.Value <> DateTimePicker2.Value Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            RestultDate = DateTimePicker2.Value

            Dim DT2 As DateTime = DateAdd(DateInterval.Day, 1, DateTimePicker2.Value)


            Restult年数 = DateDiff(DateInterval.Year, DateTimePicker1.Value, DT2)

            Me.Close()
        Else
            MsgBox("正しく期間が設定されていません", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgInputBWDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New(ByVal pStartDate As DateTime)
        InitializeComponent()
        DateTimePicker1.Value = pStartDate
        DateTimePicker2.Value = pStartDate
    End Sub

    Public Property StartDate() As DateTime
        Get
            Return DateTimePicker1.Value
        End Get
        Set(value As DateTime)
            DateTimePicker1.Value = value
        End Set
    End Property
    Public Property EndDate() As DateTime
        Get
            Return DateTimePicker2.Value
        End Get
        Set(value As DateTime)
            DateTimePicker2.Value = value
        End Set
    End Property


    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        AddValue()
    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged
        AddValue()
    End Sub

    Private Sub NumericUpDown3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown3.ValueChanged
        AddValue()
    End Sub
    Private Sub AddValue()

        Dim D1 As DateTime = DateAdd(DateInterval.Year, NumericUpDown1.Value, DateTimePicker1.Value)

        D1 = DateAdd(DateInterval.Month, NumericUpDown2.Value, D1)
        D1 = DateAdd(DateInterval.Day, NumericUpDown3.Value - 1, D1)
        DateTimePicker2.Value = D1
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Label7.Text = 和暦Format(DateTimePicker1.Value)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        Label8.Text = 和暦Format(DateTimePicker2.Value)
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            NumericUpDown1.Value = 3
            NumericUpDown2.Value = 0
            NumericUpDown3.Value = 0
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            NumericUpDown1.Value = 5
            NumericUpDown2.Value = 0
            NumericUpDown3.Value = 0
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            NumericUpDown1.Value = 10
            NumericUpDown2.Value = 0
            NumericUpDown3.Value = 0
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            NumericUpDown1.Value = 20
            NumericUpDown2.Value = 0
            NumericUpDown3.Value = 0
        End If
    End Sub
End Class
