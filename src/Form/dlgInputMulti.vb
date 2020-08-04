Imports System.Windows.Forms

Public Class dlgInputMulti
    Private mvarImputObj As InputObjectParam

    Private Sub OKBTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKBTN.Click
        If mvarImputObj.CheckValues Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()

        End If
    End Sub

    Private Sub CancelBTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelBTN.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New(ByVal DataSource As InputObjectParam, ByVal sTitle As String, ByVal sMess As String)
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        Me.Text = sTitle
        Me.Label1.Text = sMess
        mvarImputObj = DataSource
        OKBTN.Enabled = False


        CPropertyGridPlus1.SelectedObject = DataSource

    End Sub

    Private Sub CPropertyGridPlus1_PropertyValueChanged(ByVal s As Object, ByVal e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles CPropertyGridPlus1.PropertyValueChanged
        OKBTN.Enabled = mvarImputObj.CheckValues
    End Sub
End Class
