
''' <summary>
''' 枕崎市農業委員会
''' 
''' </summary>
''' <remarks>\\192.168.1.1\mapdrive\農家台帳\システム更新\枕崎市\</remarks>
Public Class C枕崎市
    Inherits C市町村別

    Public Sub New()
        MyBase.New("枕崎市")
    End Sub

    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function

    Public Overrides Sub InitLocalData()
        SysAD.SystemInfo.ユーザー.n権利 = 1
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].農振法区分 = IIf([農業振興地域]=0,2,IIf([農業振興地域]=2,3,[農業振興地域])) WHERE ((([D:農地Info].農業振興地域) Is Not Null));")
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].都市計画法区分 = IIf([都市計画法]=2,1,IIf([都市計画法]=3,2,IIf([都市計画法]=4,1,IIf([都市計画法]=5,3,[都市計画法])))) WHERE ((([D:農地Info].都市計画法) Is Not Null));")

    End Sub
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)

            .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("固定読み込み", "固定読み込み", "メンテナンス", "設定", AddressOf ClickMenu)
            .ListView.ItemAdd("耕作放棄地情報との関連付け", "耕作放棄地情報との関連付け", "メンテナンス", "設定", AddressOf ClickMenu)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "集計一覧", "印刷", AddressOf sub利用権終期管理)

            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)

            MyBase.InitMenu(pMain)
        End With
    End Sub

End Class
