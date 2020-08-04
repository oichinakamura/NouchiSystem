

''' <summary>
''' \\ibm-mapserver\システム発行\農地台帳\長島町\
''' \\10.5.1.14\システム発行\農地台帳\長島町\　（旧）RKKによる変更があり↑のパスに変更
''' </summary>
''' <remarks></remarks>
Public Class C長島町
    Inherits C市町村別

    Public Sub New()
        MyBase.New("長島町")
    End Sub

    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農家検索", ImageKey.閲覧検索, "閲覧・検索", AddressOf 農家一覧)
            '.ListView.ItemAdd("固定読み込み", "固定読み込み",ImageKey.作業, "操作", AddressOf Sub固定読み込み)
            .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", ImageKey.集計一覧, "印刷", AddressOf sub利用権終期管理)

            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", ImageKey.印刷, "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("利用意向調査出力", "利用意向調査出力", ImageKey.印刷, "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("非農地通知", "非農地通知", "閲覧・検索", "閲覧・検索", AddressOf ClickMenu)

            MyBase.InitMenu(pMain)
        End With
    End Sub
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Sub 申請10a当金額変換()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("申請10a当金額変換") Then
            SysAD.MainForm.MainTabCtrl.AddPage(New CTabPage申請10a当金額変換())
        End If
    End Sub

    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function

    Public Overrides Sub InitLocalData()

    End Sub
    Private Function QtTextTrim(ByVal sText As String) As String
        If sText.StartsWith(Chr(34)) Then
            sText = Strings.Mid(sText, 2)
        End If
        If sText.EndsWith(Chr(34)) Then
            sText = Strings.Left(sText, Len(sText) - 1)
        End If

        Return Trim(sText)
    End Function

    Private Function GetDateStr(ByVal St As String) As String
        Return St.Substring(0, 4) & "/" & St.Substring(4, 2) & "/" & St.Substring(6, 2)
    End Function

    Private Function GetDateStrR(ByVal St As String) As String
        Return Val(St.Substring(4, 2)) & "/" & Val(St.Substring(6, 2)) & "/" & St.Substring(0, 4)
    End Function

 
End Class
