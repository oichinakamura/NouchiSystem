Imports System.Windows.Forms

Public Class dlgInputForm
    Private WithEvents OK_Button As New ToolStripButton()
    Private WithEvents Cancel_Button As New ToolStripButton()
    Public List As New Dictionary(Of String, InputParam)
    Private pMessLabel As New Label
    Private WithEvents mvarUpsize As New ToolStripButton()


    Private Sub OK_Button_Click(sender As Object, e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgInputForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New(sTitle As String, sMessage As String, ParamArray pList() As InputParam)
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        Me.Text = sTitle
        OK_Button.Text = "OK"
        Cancel_Button.Text = "キャンセル"

        ToolStrip1.Items.Add(Cancel_Button)
        ToolStrip1.Items.Add(OK_Button)

        AddHandler OK_Button.Click, AddressOf OK_Button_Click
        AddHandler Cancel_Button.Click, AddressOf Cancel_Button_Click

        mvarUpsize.Text = "上"
        ToolStrip1.Items.Add(mvarUpsize)


        OK_Button.Alignment = ToolStripItemAlignment.Right
        Cancel_Button.Alignment = ToolStripItemAlignment.Right

        pMessLabel.AutoSize = False
        pMessLabel.BorderStyle = BorderStyle.FixedSingle
        pMessLabel.BackColor = Color.White
        pMessLabel.Text = sMessage
        pMessLabel.Dock = DockStyle.Fill
        SplitContainer1.Panel1.Controls.Add(pMessLabel)

        Dim pTop As Integer = 0
        For Each pItem As InputParam In pList
            Dim pLabel As New Label()
            List.Add(pItem.DataFieldName, pItem)
            pLabel.AutoSize = False
            pLabel.Size = New Size(Panel1.ClientSize.Width - 1, 25)
            pLabel.BackColor = Color.LightCyan
            pLabel.BorderStyle = BorderStyle.FixedSingle
            pLabel.TextAlign = ContentAlignment.MiddleCenter
            pLabel.Text = pItem.HeaderName
            Panel1.Controls.Add(pLabel)
            Dim pCtrl As Control = Nothing
            Select Case pItem.DataType.Name
                Case "Int32", "Decimal"
                    pCtrl = New NumericUpDownEx
                    With CType(pCtrl, NumericUpDownEx)
                        .Left = 0
                        .Top = pTop
                        .Margin = New Padding(0, 0, 0, 0)
                        .TextAlign = HorizontalAlignment.Right
                        .DataBindings.Add(New Binding("Value", pItem, "Value"))
                        .Maximum = 1000
                    End With
                Case "DateTime"
                    pCtrl = New DateTimePicker和
                    With CType(pCtrl, DateTimePicker和)
                        .Left = 0
                        .Top = pTop
                        .Margin = New Padding(0, 0, 0, 0)
                        .DataBindings.Add(New Binding("Value", pItem, "Value"))
                    End With
                Case Else
                    Stop
            End Select

            Panel2.Controls.Add(pCtrl)
            pLabel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            pLabel.Location = New Point(0, pTop)
            pCtrl.Width = Panel2.ClientSize.Width - 1
            pCtrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            If pLabel.Height > pCtrl.Height Then
                pCtrl.Height = pLabel.Height
                pTop += pLabel.Height
            Else
                pLabel.Height = pCtrl.Height
                pTop += pCtrl.Height

            End If
        Next

    End Sub

    Private Sub mvarUpsize_Click(sender As Object, e As System.EventArgs) Handles mvarUpsize.Click
        Me.Font = New Font(Me.Font.FontFamily, Me.Font.Size + 1)

        For Each pCtrl As Control In Me.Controls
            pCtrl.Font = New Font(Me.Font.FontFamily, Me.Font.Size)
        Next
    End Sub


End Class

Public Class NumericUpDownEx
    Inherits NumericUpDown

    Public Sub New()

    End Sub

    Private Sub NumericUpDownEx_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.ValueChanged
        'Me.Validate()
        'Me.ToString()
    End Sub

    Private Sub NumericUpDownEx_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.TextChanged

    End Sub
End Class

Public Class DateTimePicker和
    Inherits DateTimePicker

    Public Sub New()
        Me.Format = DateTimePickerFormat.Custom
    End Sub

    Private Sub DateTimePicker和_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        nMouseClick += -(e.Button = Windows.Forms.MouseButtons.Left)
        If nMouseClick = 2 Then
            Dim Dt As String = InputBox("日付を入力してください", "直接入力", Me.Value.ToShortDateString)

            If IsDate(Dt) Then
                Me.Value = CDate(Dt)
            End If
            nMouseClick = 0
        End If
    End Sub

    Private nMouseClick As Integer = 0
    Private Sub DateTimePicker和_MouseEnter(sender As Object, e As System.EventArgs) Handles Me.MouseEnter
        nMouseClick = 0
    End Sub

    Private Sub DateTimePicker和_MouseLeave(sender As Object, e As System.EventArgs) Handles Me.MouseLeave
        nMouseClick = 0
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.ValueChanged
        Dim calendar As New System.Globalization.JapaneseCalendar()
        Dim culture As New System.Globalization.CultureInfo("ja-JP")
        culture.DateTimeFormat.Calendar = calendar

        Dim target As DateTime = Me.Value
        Dim result As String = target.ToString("gggyy年M月d日", culture)
        If InStr(result, "平成") > 0 AndAlso CInt(target.ToString("yyyyMMdd")) >= 20190501 Then
            target = Me.Value.AddYears(-30)
            result = target.ToString("gg yy", culture).Replace("平成", "令和") + "年MM月dd日"
            Me.CustomFormat = result
        Else
            Me.CustomFormat = Me.Value.ToString("gg yy", culture) + "年MM月dd日"
        End If
    End Sub



    Protected Overrides Sub OnMouseDoubleClick(e As System.Windows.Forms.MouseEventArgs)
        Dim Dt As String = InputBox("日付を入力してください", "直接入力", Me.Value.ToString)
        If IsDate(Dt) Then
            Me.Value = CDate(Dt)
        End If
        MyBase.OnMouseDoubleClick(e)
    End Sub
End Class

Public Class InputParam
    Public DataFieldName As String = ""
    Public HeaderName As String = ""
    Public GroupName As String = ""
    Public Property Value As Object

    Public DataType As System.Type
    Public ControlType As Control
    Public NullDefault As Object
    Public SelectSource As DataView

    Public Sub New(sField As String, sHeader As String, pValue As Object, pType As System.Type)
        DataFieldName = sField
        HeaderName = sHeader
        Value = pValue
        DataType = pType
    End Sub

    Public Overrides Function ToString() As String
        Return Value.ToString()
    End Function


End Class
