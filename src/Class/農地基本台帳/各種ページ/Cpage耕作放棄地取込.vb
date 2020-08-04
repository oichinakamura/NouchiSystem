

Public Class Cpage耕作放棄地取込
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvar耕作放棄地用MDB As String = ""

    Public Sub New()
        MyBase.New(True)
        Me.Name = "耕作放棄地取込み"
        Me.Text = "耕作放棄地取込み"

        mvar耕作放棄地用MDB = SysAD.GetXMLProperty("耕作放棄地取込み", "DBFilePath", "")
        If Not IO.File.Exists(mvar耕作放棄地用MDB) Then
            With New OpenFileDialog
                .Filter = "農政*.MDB|農政*.MDB"
                .Title = "農政情報.MDB"
                If .ShowDialog = DialogResult.OK Then
                    Dim pTable As New DataTable
                    Dim pLDB As New HimTools2012.Data.CLocalDataEngine("")
                    pLDB.LocalPath = .FileName

                End If
            End With
        End If

    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

   
End Class
