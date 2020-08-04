

''' <summary>
''' いちき串木野市(\\10.5.8.13\システム配置\農地台帳システム\いちき串木野市\)
''' </summary>
''' <remarks></remarks>
Public Class Cいちき串木野市
    Inherits C市町村別

    Public Sub New()
        MyBase.New("いちき串木野市")
    End Sub
    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", ImageKey.閲覧検索, "閲覧・検索", AddressOf 農家一覧)

            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [S_システムデータ] WHERE [Key] Is Not Null")
            pTBL.PrimaryKey = New DataColumn() {pTBL.Columns("Key")}
            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "集計一覧", "印刷", AddressOf sub利用権終期管理)

            MyBase.InitMenu(pMain)
        End With

    End Sub

    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function

    Public Overrides Sub InitLocalData()

    End Sub


    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
End Class
