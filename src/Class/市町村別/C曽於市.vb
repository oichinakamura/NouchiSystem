

Public Class C曽於市
    Inherits C市町村別
    Public Sub New()
        MyBase.New("曽於市")
    End Sub

    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            .ListView.ItemAdd("フェーズ２利用状況調査・意向調査情報", "フェーズ２利用状況調査・意向調査情報", ImageKey.他システム連携, "他システム連携", AddressOf ClickMenu)
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
