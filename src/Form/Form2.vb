Public Class Form2

    Private WithEvents mvarRow As HimTools2012.Data.DataRowEx
    Public Sub New(pRow As HimTools2012.Data.DataRowEx)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        mvarRow = pRow

        Dim pCtrl As CustomControlLIB.CustomViewPage = ElementHost1.Child
        'AddHandler pCtrl.OnButtonClick, AddressOf onB
        pCtrl.SetData(SysAD.InterfaceSetting, "個人Info", mvarRow)
    End Sub

    Private Sub mvarRow_DataChange(s As Object, e As System.Data.DataColumnChangeEventArgs) Handles mvarRow.DataChange

    End Sub

    Private Sub Form2_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class