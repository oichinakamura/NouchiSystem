
Imports System.CodeDom.Compiler
Imports System.Reflection
Imports System.Text

Public Class CPage共通一覧集計
    Inherits HimTools2012.SystemWindows.CMainPageSK
    Public Sub New()
        MyBase.New(True, False, "SUB共通一覧集計", "共通一覧集計")


        mvarListView.Groups.Add("閲覧・検索", "閲覧・検索").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add(MenuGroup.grp印刷, MenuGroup.grp印刷).HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("選挙関連", "選挙関連").HeaderAlignment = HorizontalAlignment.Left

        mvarListView.Groups.Add("操作", "操作>>").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("他システム連携", "他システム連携").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("設定", "設定").HeaderAlignment = HorizontalAlignment.Left

        With Me
            .ListView.ItemAdd("H26各種集計", "H26各種集計", "集計一覧", MenuGroup.grp印刷, AddressOf ClickMenu)
            .ListView.ItemAdd("戻る", "戻る", "作業", "操作", AddressOf ClickMenu)
        End With
    End Sub
    Public Sub ClickMenu(s As Object, ByVal e As EventArgs)
        Select Case CType(s, ListViewItem).Text
            Case "H26各種集計"
                If Not SysAD.MainForm.MainTabCtrl.ExistPage("H26各種集計") Then
                    SysAD.MainForm.MainTabCtrl.AddPage(New CH26各種集計)
                End If
            Case "戻る"
                If SysAD.MainForm.MainTabCtrl.ExistPage("Main") Then
                    SysAD.MainForm.MainTabCtrl.TabPages.Remove(Me)
                    Me.Dispose()
                End If
        End Select
    End Sub
End Class
