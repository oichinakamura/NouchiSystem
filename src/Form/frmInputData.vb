Imports System.Windows.Forms
Imports System.ComponentModel

Public Class frmInputData
    Private WithEvents btnOK As New ToolStripButton
    Private WithEvents btnCancel As New ToolStripButton
    Private WithEvents pGrid As System.Windows.Forms.PropertyGrid
    Private WithEvents pObject As CInputDate

    Public RetObj As CInputDate

    Public Sub New(ByVal sTitle As String, ByVal sMessage As String, ByRef Obj As CInputDate)
        InitializeComponent()

        Me.Text = sTitle
        Me.RichTextBox1.Text = sMessage
        pObject = Obj
        RichTextBox1.ReadOnly = True
        StatusStrip1.Height = 30
        StatusStrip1.AutoSize = True

        StatusStrip1.LayoutStyle = ToolStripLayoutStyle.StackWithOverflow
        btnOK.Text = "OK"
        btnOK.Font = New Font(Me.Font.FontFamily, 12, FontStyle.Regular)
        btnOK.Alignment = ToolStripItemAlignment.Right
        btnOK.BackColor = Color.LightGray
        AddHandler btnOK.Click, AddressOf OK_Button_Click

        btnCancel.Text = "ｷｬﾝｾﾙ"
        btnCancel.Font = New Font(Me.Font.FontFamily, 12, FontStyle.Regular)
        btnCancel.Alignment = ToolStripItemAlignment.Right
        btnCancel.BackColor = Color.LightGray
        AddHandler btnCancel.Click, AddressOf Cancel_Button_Click

        StatusStrip1.Items.AddRange(New ToolStripItem() {btnCancel, btnOK})

        pGrid = New System.Windows.Forms.PropertyGrid
        pGrid.LineColor = Color.LightGray
        pGrid.Dock = DockStyle.Fill
        pGrid.ToolbarVisible = False
        pGrid.Update()
        pGrid.Font = New Font(Me.Font.FontFamily, 12, FontStyle.Regular)
        pGrid.PropertySort = PropertySort.Categorized
        pGrid.SelectedObject = pObject
        Me.SplitContainer1.Panel2.Controls.Add(pGrid)
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        RetObj = pObject
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub pGrid_PropertyValueChanged(ByVal s As Object, ByVal e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles pGrid.PropertyValueChanged
        btnOK.Enabled = pObject.DataValidate
    End Sub

    Private Sub pObject_ValueError(ByVal s As Object, ByVal sPropertyName As String, ByVal sMessage As String, ByVal NewValue As Object) Handles pObject.ValueError
        MsgBox(sMessage, , "入力エラー:" & sPropertyName)
    End Sub

    Private Sub frmInputData_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SplitContainer1_Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub
End Class

Public MustInherit Class InputDate
    Public MustOverride ReadOnly Property DataValidate As Boolean

    Public Event ValueError(ByVal s As Object, ByVal sPropertyName As String, ByVal sMessage As String, ByVal NewValue As Object)

    Public Sub ErrorEvent(ByVal pObject As Object, ByVal sPropeertyName As String, ByVal sMessage As String, ByVal NewValue As Object)
        RaiseEvent ValueError(pObject, sPropeertyName, sMessage, NewValue)
    End Sub

    <TypeConverter(GetType(C範囲入力ClassConverter))>
    Public Class C範囲入力
        Private _範囲開始 As Date = Now.Date
        Private _範囲終了 As Date = Now.Date

        <ReadOnlyAttribute(False)>
        Public Property 範囲開始() As Date
            Get
                Return _範囲開始
            End Get
            Set(ByVal Value As Date)
                _範囲開始 = Value
            End Set
        End Property

        <ReadOnlyAttribute(False)>
        Public Property 範囲終了() As Date
            Get
                Return _範囲終了
            End Get
            Set(ByVal Value As Date)
                _範囲終了 = Value
            End Set
        End Property

        <ReadOnlyAttribute(True)>
        Public ReadOnly Property 日数() As Integer
            Get
                Try
                    Dim pBit As TimeSpan = _範囲終了.Subtract(_範囲開始)
                    Return pBit.TotalDays + 1
                Catch ex As Exception
                    Return 0
                End Try
            End Get
        End Property

        Public Overrides Function ToString() As String
            If _範囲開始.Year = _範囲終了.Year AndAlso _範囲開始.Month = _範囲終了.Month AndAlso _範囲開始.Day = _範囲終了.Day Then
                Return Strings.Format(_範囲開始, "MM月dd日")
            ElseIf _範囲開始.Year = _範囲終了.Year AndAlso _範囲開始.Month = _範囲終了.Month Then
                Return Strings.Format(_範囲開始, "MM月dd日") & "～" & Strings.Format(_範囲終了, "dd日")
            ElseIf _範囲開始.Year = _範囲終了.Year Then
                Return Strings.Format(_範囲開始, "MM月dd日") & "～" & Strings.Format(_範囲終了, "MM月dd日")
            Else
                Return Strings.Format(_範囲開始, "yyyy年MM月dd日") & "," & Strings.Format(_範囲終了, "yyyy年MM月dd日")
            End If
        End Function
    End Class

    Public Class C範囲入力ClassConverter
        Inherits ExpandableObjectConverter

        Public Overloads Overrides Function CanConvertTo(ByVal context As ITypeDescriptorContext, ByVal destinationType As Type) As Boolean
            If destinationType Is GetType(C範囲入力) Then
                Return True
            End If
            Return MyBase.CanConvertTo(context, destinationType)
        End Function

        Public Overloads Overrides Function ConvertTo(ByVal context As ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As Type) As Object
            If destinationType Is GetType(String) And TypeOf value Is C範囲入力 Then
                Dim cc As C範囲入力 = CType(value, C範囲入力)
                Return cc.ToString
            End If
            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End Function

        Public Overloads Overrides Function CanConvertFrom(ByVal context As ITypeDescriptorContext, ByVal sourceType As Type) As Boolean
            If sourceType Is GetType(String) Then
                Return True
            End If
            Return MyBase.CanConvertFrom(context, sourceType)
        End Function

        Public Overloads Overrides Function ConvertFrom(ByVal context As ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object
            If TypeOf value Is String Then

                Try
                    Dim ss As String() = Split(value, "～")
                    Dim cc As New C範囲入力

                    Select Case UBound(ss)
                        Case 0

                    End Select

                    cc.範囲開始 = CDate(ss(0))
                    cc.範囲終了 = CDate(ss(1))
                    Return cc

                Catch ex As Exception
                    Return Nothing
                End Try
            End If
            Return MyBase.ConvertFrom(context, culture, value)
        End Function
    End Class
End Class

