

Public Module MenuGroup
    Public Const grp印刷 = "印刷"
End Module

Public Class CMainPage
    Inherits HimTools2012.SystemWindows.CMainPageSK

    Public Sub New()
        MyBase.New(False, False, "Main", "メイン")
        SysAD.ImageKeyAlias.Add("作業", "申請")

        mvarListView.Groups.Add("閲覧・検索", "閲覧・検索").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add(MenuGroup.grp印刷, MenuGroup.grp印刷).HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("選挙関連", "選挙関連").HeaderAlignment = HorizontalAlignment.Left

        mvarListView.Groups.Add("操作", "操作>>").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("他システム連携", "他システム連携").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("設定", "設定").HeaderAlignment = HorizontalAlignment.Left
    End Sub

    Public Sub EndProg()
        End
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

End Class
