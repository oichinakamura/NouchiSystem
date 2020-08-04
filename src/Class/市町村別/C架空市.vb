

Public Class C架空市
    Inherits C市町村別
    Public Sub New()
        MyBase.New("架空市")
    End Sub

    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "印刷", "印刷", AddressOf sub利用権終期管理)


            .ListView.ItemAdd("固定資産比較", "固定資産比較", "他システム連携", "操作", AddressOf 固定資産比較)
            .ListView.ItemAdd("農地重複", "農地重複", "他システム連携", "操作", AddressOf sub農地重複)
            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("自然解約の実行", "自然解約の実行", ImageKey.作業, "操作", AddressOf ClickMenu)

            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            MyBase.InitMenu(pMain)
        End With

    End Sub
    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overrides Sub InitLocalData()

    End Sub
End Class
