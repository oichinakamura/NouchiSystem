'20160531霧島

Imports System.ComponentModel
Imports HimTools2012.controls.PropertyGridSupport
Imports HimTools2012.TypeConverterCustom

Public MustInherit Class C市町村別
    Inherits HimTools2012.System管理.CMunicipality

    Public Sub New(ByVal s市町村名 As String)
        MyBase.New(s市町村名)
    End Sub

    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.ItemAdd("共通一覧集計", "共通一覧集計", ImageKey.集計一覧, "印刷", AddressOf sub共通一覧集計)
            .ListView.ItemAdd("相続未登記農地等に係る実態調査", "相続未登記農地等に係る実態調査", ImageKey.集計一覧, "印刷", AddressOf sub相続未登記)
            .ListView.ItemAdd("相続未登記地区別調査", "相続未登記地区別調査", ImageKey.集計一覧, "印刷", AddressOf sub相続未登記地区別)
            .ListView.ItemAdd("農家別自小作法令別集計", "農家別自小作法令別耕作面積集計", ImageKey.集計一覧, "印刷", AddressOf sub農家別自小作法令別集計)
            .ListView.ItemAdd("公開用CSV出力", "公開用CSV出力", ImageKey.他システム連携, "他システム連携", AddressOf subCSV出力)
            .ListView.ItemAdd("公開用CSV出力(農地一覧)", "公開用CSV出力(農地一覧)", ImageKey.他システム連携, "他システム連携", AddressOf subCSV出力_カスタム)
            .ListView.ItemAdd("フェーズ2移行用CSV出力", "フェーズ2移行用CSV出力", ImageKey.他システム連携, "他システム連携", AddressOf subCSV出力2)

            '.ListView.ItemAdd("フェーズ２エラーチェック", "フェーズ２エラーチェック", ImageKey.他システム連携, "他システム連携", AddressOf subF2ErrorCheck)
            .ListView.ItemAdd("標準フォーマット出力", "標準フォーマット出力", ImageKey.他システム連携, "他システム連携", AddressOf subフェーズ2標準フォーマット出力)
            .ListView.ItemAdd("農地権利移動・借賃等調査データ出力", "農地権利移動・借賃等調査データ出力", ImageKey.集計一覧, "印刷", AddressOf sub権利移動借賃データ出力)
            .ListView.ItemAdd("事務処理状況出力", "事務処理状況出力", ImageKey.集計一覧, "印刷", AddressOf sub事務処理状況出力)
            .ListView.ItemAdd("利用意向調査出力", "利用意向調査出力", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("農地10a当金額変換", "農地10a当金額変換", ImageKey.作業, "操作", AddressOf 農地10a当金額変換)
            .ListView.ItemAdd("アップデート", "アップデート", ImageKey.メンテナンス, "設定", AddressOf UpdateSystem)
            .ListView.ItemAdd("検索フリガナの初期化", "検索フリガナの初期化", ImageKey.メンテナンス, "設定", AddressOf sub検索フリガナ修正)
            .ListView.ItemAdd("再起動", "再起動", ImageKey.終了, "操作", AddressOf 再起動)
            .ListView.ItemAdd("終了", "終了", ImageKey.終了, "操作", AddressOf EndProgram)
            .ListView.ItemAdd("印刷様式管理", "印刷様式管理", ImageKey.メンテナンス, "設定", Sub(s, e) SysAD.ShowFolder(SysAD.CustomReportFolder(SysAD.市町村.市町村名)))
            .ListView.ItemAdd("調査日一括設定", "調査日一括設定", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("利用意向調査実施農地状況一覧", "利用意向調査実施農地状況一覧", ImageKey.集計一覧, "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("全農地CSV出力", "全農地CSV出力", ImageKey.他システム連携, "操作", AddressOf sub全農地CSV出力)

            If Not SysAD.IsClickOnceDeployed Then
                .ListView.ItemAdd("各種グリッド設定", "各種グリッド設定", ImageKey.メンテナンス, "設定", AddressOf 各種グリッド設定)
                .ListView.ItemAdd("クラス作成", "クラス作成", ImageKey.メンテナンス, "設定", AddressOf クラス作成)
                .ListView.ItemAdd("共通マスタ管理", "共通マスタ管理", ImageKey.メンテナンス, "設定", AddressOf 共通マスタ管理)
            End If

        End With
    End Sub

    <TypeConverter(GetType(PropertyOrderConverter))>
    Public Class TestC
        Inherits HimTools2012.InputSupport.CInputSupport

        Public Sub New()
            MyBase.New(Nothing)
        End Sub

        <Browsable(True)>
        Public Property Text As String
        Public Int As Integer
    End Class

    Public MustOverride ReadOnly Property 旧農振都市計画使用 As Boolean

    Public Sub 共通マスタ管理()
        SysAD.MainForm.MainTabCtrl.ExistPage("CommonMASTER管理", True, GetType(CommonMASTER管理))
    End Sub

    Public Sub UpdateSystem()
        If SysAD.UpdateSystem() Then
            MsgBox("終了します。プログラムを再起動してください。")
            End
        End If
    End Sub

    Public Sub 再起動()
        Application.Restart()
    End Sub
    Public MustOverride Function Get選挙世帯一覧() As DataTable

    Public Sub 固定資産比較()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("固定資産比較", True) Then
            SysAD.MainForm.MainTabCtrl.AddPage(New C固定資産比較(-12))
        End If
    End Sub

    Private Sub クラス作成()
        SysAD.MainForm.MainTabCtrl.ExistPage("クラス作成", True, GetType(CTabPageClassGenerator))
    End Sub
    Private Sub 各種グリッド設定()
        SysAD.MainForm.MainTabCtrl.ExistPage("固定資産比較", True, GetType(CTabPage各種グリッド設定))
    End Sub
    Public Sub sub10a当たりの平均賃借料集計()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("10a当たりの平均賃借料集計", True) Then
            Dim pGrid As New HimTools2012.controls.DataGridViewWithDataView()
            Dim pPage As New CNList農地台帳("10a当たりの平均賃借料集計", "10a当たりの平均賃借料集計", True)
            SysAD.MainForm.MainTabCtrl.AddPage(pPage)
            pPage.GView.SetDataView(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Year([小作開始年月日]) AS 設定年, [D:農地Info].大字ID, V_大字.名称 AS 大字名, Avg([D:農地Info].[10a賃借料]) AS 10a賃借料の平均, Sum([D:農地Info].[10a賃借料]) AS 10a賃借料の合計, Count([D:農地Info].ID) AS 件数 FROM [D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID WHERE ((([D:農地Info].[10a賃借料])>0) AND (([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作形態)=1) AND (([D:農地Info].自小作別)>0) AND (([D:農地Info].小作開始年月日) Is Not Null)) GROUP BY Year([小作開始年月日]), [D:農地Info].大字ID, V_大字.名称 ORDER BY Year([小作開始年月日]), [D:農地Info].大字ID;"), "", "")
        End If
    End Sub

    Public Sub sub共通一覧集計()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("SUB共通一覧集計", True) Then
            SysAD.MainForm.MainTabCtrl.AddPage(New CPage共通一覧集計())
        End If
    End Sub

    Public Sub sub農家別自小作法令別集計()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("農家別自小作法令別集計", True) Then
            SysAD.MainForm.MainTabCtrl.AddPage(New C農家別自小作法令別集計)
        End If
    End Sub

    Public Sub sub利用権終期管理()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("利用権終期管理") Then
            SysAD.MainForm.MainTabCtrl.AddPage(New C利用権終期管理())
        End If
    End Sub

    Public Sub sub農地重複()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("重複農地") Then
            SysAD.MainForm.MainTabCtrl.AddPage(New CTabPage重複農地)
        End If
    End Sub

    Public Enum 地目Type
        指定なし = -1
        田地目 = 1
        畑地目 = 2
        農地地目 = 3
        その他地目 = 0
    End Enum

    Public Overridable ReadOnly Property 市町村別登記地目CD(ByVal nType As 地目Type) As Integer()
        Get
            Return Make市町村別登記地目コード(nType)
        End Get
    End Property

    Protected ReadOnly Property Make市町村別登記地目コード(ByVal nType As 地目Type) As Integer()
        Get
            Dim pArr As New List(Of Integer)
            Select Case nType
                Case 地目Type.田地目
                    For Each pRow As DataRow In App農地基本台帳.TBL地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "田", "介在田", ""
                                pArr.Add(pRow.Item("ID"))
                            Case "塩田"
                            Case Else
                                If InStr(pRow.Item("名称").ToString, "田") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                End If
                        End Select
                    Next
                    Return pArr.ToArray
                Case 地目Type.畑地目
                    For Each pRow As DataRow In App農地基本台帳.TBL地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "畑", "介在畑", ""
                                pArr.Add(pRow.Item("ID"))
                            Case Else
                                If InStr(pRow.Item("名称").ToString, "畑") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                End If
                        End Select
                    Next
                    Return pArr.ToArray
                Case 地目Type.農地地目
                    For Each pRow As DataRow In App農地基本台帳.TBL地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "田", "介在田", ""
                                pArr.Add(pRow.Item("ID"))
                            Case "畑", "介在畑", ""
                                pArr.Add(pRow.Item("ID"))
                            Case "塩田"
                            Case Else
                                If InStr(pRow.Item("名称").ToString, "田") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                ElseIf InStr(pRow.Item("名称").ToString, "畑") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                End If
                        End Select
                    Next
                    Return pArr.ToArray
                Case Else
                    For Each pRow As DataRow In App農地基本台帳.TBL地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "田", "介在田", ""
                            Case "畑", "介在畑", ""
                            Case Else
                                If InStr(pRow.Item("名称").ToString, "田") > 0 Then

                                ElseIf InStr(pRow.Item("名称").ToString, "畑") > 0 Then
                                Else
                                    pArr.Add(pRow.Item("ID"))
                                End If
                        End Select
                    Next
                    Return pArr.ToArray
            End Select
        End Get
    End Property

    Public Overridable ReadOnly Property 市町村別現況地目CD(ByVal nType As 地目Type) As Integer()
        Get
            Return Make市町村別現況地目コード(nType)
        End Get
    End Property
    Protected ReadOnly Property Make市町村別現況地目コード(ByVal nType As 地目Type) As Integer()
        Get
            Dim pArr As New List(Of Integer)
            Select Case nType
                Case 地目Type.田地目
                    For Each pRow As DataRow In App農地基本台帳.TBL現況地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "田", "介在田", ""
                                pArr.Add(pRow.Item("ID"))
                            Case "塩田"
                            Case Else
                                If InStr(pRow.Item("名称").ToString, "田") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                End If
                        End Select
                    Next
                    Return pArr.ToArray
                Case 地目Type.畑地目
                    For Each pRow As DataRow In App農地基本台帳.TBL現況地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "畑", "介在畑", ""
                                pArr.Add(pRow.Item("ID"))
                            Case Else
                                If InStr(pRow.Item("名称").ToString, "畑") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                End If
                        End Select
                    Next
                    Return pArr.ToArray
                Case 地目Type.農地地目
                    For Each pRow As DataRow In App農地基本台帳.TBL現況地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "田", "介在田", ""
                                pArr.Add(pRow.Item("ID"))
                            Case "畑", "介在畑", ""
                                pArr.Add(pRow.Item("ID"))
                            Case "塩田"
                            Case Else
                                If InStr(pRow.Item("名称").ToString, "田") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                ElseIf InStr(pRow.Item("名称").ToString, "畑") > 0 Then
                                    pArr.Add(pRow.Item("ID"))
                                End If
                        End Select
                    Next
                    Return pArr.ToArray
                Case Else
                    Dim n地目 As New List(Of Integer)
                    For Each pRow As DataRow In App農地基本台帳.TBL現況地目.Rows
                        Select Case pRow.Item("名称").ToString
                            Case "田", "畑", "介在田", "介在畑", "宅地介在田", "採草地"
                            Case "宅地", "雑種地介在宅地", "塩田", "鉱泉地", "池沼", "山林", "宅地介在山林", "農地介在山林", "牧場", "附奇洲", "古墳", "養鶏場敷地", "原野", "ため池", "河川敷", "公有地", "学校敷地", "私道", "荒地", "溶岩", "公民館敷地", "公営住宅用地", "ゴルフ場", "鉄軌道", "雑種地", "墓地", "境内地", "運河用地", "水道用地", "用悪水路", "その他", "堤", "井溝", "保安林", "公衆用道路", "公園"
                                n地目.Add(pRow.Item("ID"))
                            Case Else
                                n地目.Add(pRow.Item("ID"))
                                Debug.Print("""" & pRow.Item("名称").ToString & """")
                                If Not SysAD.IsClickOnceDeployed Then
                                    Stop
                                End If
                        End Select
                    Next
                    Return n地目.ToArray
            End Select
        End Get
    End Property

    Public Overridable ReadOnly Property Get受付しめ月() As Integer
        Get
            Return 0
        End Get
    End Property

    Public Overridable ReadOnly Property Get総会締日() As Integer
        Get
            Return Val(SysAD.DB(sLRDB).DBProperty("申請締切", 15))
        End Get
    End Property

    Public Overridable ReadOnly Property Get総会次受付日() As Integer
        Get
            Return 16
        End Get
    End Property

    Public Sub sub相続未登記()
        SysAD.MainForm.MainTabCtrl.ExistPage("相続未登記農地等に係る実態調査", True, GetType(C相続未登記実態調査))
    End Sub

    Public Sub sub相続未登記地区別()
        SysAD.MainForm.MainTabCtrl.ExistPage("相続未登記地区別調査", True, GetType(C相続未登記地区別調査))
    End Sub

    Public Sub sub事務処理状況出力()
        'With New dlgInputBWDate(SysAD.GetXMLProperty("集計", "事務処理状況出力-開始", Now.Date))
        '    If .ShowDialog = DialogResult.OK Then
        '        Dim Start日付 As String = String.Format("#{0}/{1}/{2}#", .StartDate.Month, .StartDate.Day, .StartDate.Year)
        '        Dim End日付 As String = String.Format("#{0}/{1}/{2}#", .EndDate.Month, .EndDate.Day, .EndDate.Year)
        '        Dim 事務処理状況出力 = New CTabPage事務処理状況出力(Start日付, End日付)
        '    End If
        'End With
        SysAD.MainForm.MainTabCtrl.ExistPage("事務処理状況出力", True, GetType(CTabPage事務処理状況出力))
    End Sub

    Public Sub subフェーズ2標準フォーマット出力()
        SysAD.MainForm.MainTabCtrl.ExistPage("標準フォーマット出力", True, GetType(CTabPage標準フォーマット出力))
    End Sub

    Public Sub 農家一覧()
        If SysAD.page農家世帯 Is Nothing OrElse Not SysAD.MainForm.MainTabCtrl.TabPages.Contains(SysAD.page農家世帯) Then
            SysAD.page農家世帯 = New classPage農家世帯
            SysAD.MainForm.MainTabCtrl.TabPages.Add(SysAD.page農家世帯)
        ElseIf Not SysAD.MainForm.MainTabCtrl.TabPages.Contains(SysAD.page農家世帯) Then
            SysAD.MainForm.MainTabCtrl.TabPages.Add(SysAD.page農家世帯)
        End If
        SysAD.MainForm.MainTabCtrl.SelectedTab = SysAD.page農家世帯
    End Sub

    Public Sub 現地調査()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("現地調査表作成.0", True, GetType(CTabPage現地調査表作成)) Then
        End If
    End Sub
    Public Sub 農地一覧()
        If SysAD.page農家世帯 Is Nothing Then
            SysAD.page農家世帯 = New classPage農家世帯
            SysAD.MainForm.MainTabCtrl.TabPages.Add(SysAD.page農家世帯)
        End If

        SysAD.MainForm.MainTabCtrl.SelectedTab = SysAD.page農家世帯
    End Sub

    Public Shared Sub sub外字テスト()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_BASICALL]")

        For Each pRow As DataRow In pTBL.Rows
            Dim n As Integer = SJIS外字Check(pRow.Item("名称").ToString)
            If n > 0 Then
                Debug.Print(pRow.Item("Class").ToString & "[" & pRow.Item("ID").ToString & "]" & pRow.Item("名称").ToString)
            End If
        Next
    End Sub

    Public Shared Function SJIS外字Check(ByRef sText As String) As Integer
        For n As Integer = 1 To sText.Length
            Select Case sText.Substring(n - 1, 1)
                Case Else
                    Dim sHex As String = HimTools2012.StringF.Right("00000000" & Hex(Asc(sText.Substring(n - 1, 1))), 8)
                    Dim HCD As Byte = Val("&H" & HimTools2012.StringF.Mid(sHex, 5, 2))
                    Dim LCD As Byte = Val("&H" & HimTools2012.StringF.Right(sHex, 2))
                    Select Case HCD
                        Case &HFA
                            Select Case LCD
                                Case &H40 To &H49
                                Case &H55 To &H57
                                Case &H5C To &H7E
                                Case &H80 To &HFC
                                Case Else
                                    Return n
                            End Select
                        Case &HFB
                            Select Case LCD
                                Case &H40 To &H7E
                                Case &H80 To &HFC
                                Case Else
                                    Return n
                            End Select
                        Case &HFC
                            Select Case LCD
                                Case &H40 To &H4B

                                Case Else
                                    Return n
                            End Select
                            '0xeb40～ 0xeffc　
                            '0xf040～
                        Case &HEB, &HEC, &HED, &HEE
                            Return n
                        Case &HEF
                            If LCD <= &HFC Then
                                Return n
                            End If
                        Case &HF0 To &HFF
                            Return n
                        Case Else
                    End Select
            End Select
        Next
        Return 0
    End Function

    Public Sub ClickMenu(ByVal s As Object, ByVal e As EventArgs)
        Select Case CType(s, ListViewItem).Text
            Case "転用非農地整合性確認"
                If Not SysAD.MainForm.MainTabCtrl.ExistPage("転用非農地整合性確認") Then
                    SysAD.MainForm.MainTabCtrl.AddPage(New classpage転用非農地整合性確認)
                End If
            Case "自然解約の実行"
                If Not SysAD.MainForm.MainTabCtrl.ExistPage("自然解約の実行") Then
                    Dim sDate As String = InputBox("自然解約の期日を入力してください", "自然解約の実行", Now.Date.ToString.Replace(" 0:00:00", ""))
                    If Not IsDBNull(sDate) Then
                        SysAD.MainForm.MainTabCtrl.AddPage(New CTabPage自然解約実行(CDate(sDate)))
                    End If
                End If
            Case "重複農地検索"
                If Not SysAD.MainForm.MainTabCtrl.ExistPage("重複農地検索.0") Then
                    SysAD.MainForm.MainTabCtrl.AddPage(New CPage重複農地検索())
                End If
            Case "調査日一括設定"
                With New HimTools2012.PropertyGridDialog(New C調査日設定, " 調査日一括設定")

                    If .ShowDialog = DialogResult.OK Then
                        With CType(.ResultProperty, C調査日設定)
                            If IsDate(.調査日) AndAlso .調査日種類 = C調査日設定.enum調査日種類.利用状況調査 Then
                                If .設定条件 = C調査日設定.enum設定条件.全て Then
                                    With .調査日
                                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].利用状況調査日 = #{1}/{0}/{2}#;", .Day, .Month, .Year)
                                        MsgBox("設定しました。")
                                    End With
                                Else
                                    With .調査日
                                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].利用状況調査日 = #{1}/{0}/{2}# WHERE ((([D:農地Info].利用状況調査日) Is Null Or ([D:農地Info].利用状況調査日)<#1/1/1950#));", .Day, .Month, .Year)
                                        MsgBox("設定しました。")
                                    End With
                                End If
                            ElseIf IsDate(.調査日) AndAlso .調査日種類 = C調査日設定.enum調査日種類.利用意向調査 Then
                                If .設定条件 = C調査日設定.enum設定条件.全て Then
                                    With .調査日
                                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].利用意向調査日 = #{1}/{0}/{2}#;", .Day, .Month, .Year)
                                        MsgBox("設定しました。")
                                    End With
                                Else
                                    With .調査日
                                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].利用意向調査日 = #{1}/{0}/{2}# WHERE ((([D:農地Info].利用意向調査日) Is Null Or ([D:農地Info].利用意向調査日)<#1/1/1950#));", .Day, .Month, .Year)
                                        MsgBox("設定しました。")
                                    End With
                                End If

                            End If
                        End With
                    End If
                End With
            Case "利用意向調査出力" : SysAD.MainForm.MainTabCtrl.ExistPage("利用意向調査出力", True, GetType(CPage利用意向調査H26))
            Case "利用意向調査実施農地状況一覧" : SysAD.MainForm.MainTabCtrl.ExistPage("利用意向調査実施農地状況一覧", True, GetType(CTabPage利用意向調査実施農地状況一覧))
            Case "総会資料作成" : 総会資料作成()
            Case "住記―農家照合" : MsgBox("夜間に既に実行されています。")
            Case "非農地通知"
                If SysAD.市町村.市町村名 = "長島町" Then
                    SysAD.MainForm.MainTabCtrl.ExistPage("非農地通知", True, GetType(CTabPage非農地通知長島))
                ElseIf SysAD.市町村.市町村名 = "伊佐市" Then
                    SysAD.MainForm.MainTabCtrl.ExistPage("非農地通知", True, GetType(CTabPage非農地通知伊佐))
                Else
                    SysAD.MainForm.MainTabCtrl.ExistPage("非農地通知", True, GetType(CTabPage非農地通知))
                End If
            Case "農地権利移動・借賃等調査システム出力" : MsgBox("データベースの設定などを行ってください")
            Case "農家一覧" : SysAD.MainForm.MainTabCtrl.ExistPage("農家一覧", True, GetType(CTabPage農家一覧))
            Case "農地台帳一括印刷" : SysAD.MainForm.MainTabCtrl.ExistPage("農地台帳一括印刷", , GetType(CPage農地台帳一括印刷))
            Case "工事進捗状況一覧"
                If Not SysAD.MainForm.MainTabCtrl.ExistPage("工事進歩状況一覧表") Then
                    SysAD.MainForm.MainTabCtrl.AddPage(New CTabPage工事進歩状況一覧表)
                End If
            Case "耕作放棄地情報との関連付け"
                If Not SysAD.MainForm.MainTabCtrl.ExistPage("耕作放棄地取込み") Then
                    SysAD.MainForm.MainTabCtrl.AddPage(New Cpage耕作放棄地取込)
                End If
            Case "フェーズ２利用状況調査・意向調査情報"
                'SysAD.MainForm.MainTabCtrl.ExistPage("フェーズ２利用状況調査・意向調査情報", True, GetType(CTabPageF2利用状況意向調査))
                'Process.Start(System.IO.Path.Combine(SysAD.CustomReportFolder(SysAD.市町村.市町村名), "CSVConverter.exe"))

            Case "マニュアル" : SysAD.MainForm.MainTabCtrl.ExistPage("マニュアル", True, GetType(CManualPage))
            Case Else
                Select Case CType(s, ListViewItem).Name
                    Case "申請・許可名簿"
                        Dim X As New CPrint受付許可名簿作成
                        X.Dialog.StartProc(True, False)

                    Case Else
                        If Not SysAD.IsClickOnceDeployed Then
                            Stop
                        Else
                            MsgBox("データベースを正しく設定してください。")
                        End If

                End Select

        End Select
    End Sub

    Public Sub CSVto農地()
        SysAD.MainForm.MainTabCtrl.ExistPage("CSVTo農地", , GetType(CTabPageCSVTo農地))
    End Sub

    Public Sub 農地10a当金額変換()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("農地10a当金額変換") Then
            SysAD.MainForm.MainTabCtrl.AddPage(New CTabPage農地10a当金額変換())
        End If
    End Sub

    Public Sub sub農用地面積()
        If Not SysAD.MainForm.MainTabCtrl.TabPages.ContainsKey("農用地面積集計") Then
            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].大字ID, V_大字.名称, V_大字.大字, Sum([D:農地Info].田面積) AS 田面計, Sum([D:農地Info].畑面積) AS 畑面計, Sum([D:農地Info].田面積+[D:農地Info].畑面積) AS 農地計 FROM [D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID WHERE ((([D:農地Info].所在) Is Null) AND (([D:農地Info].農業振興地域)=1)) GROUP BY [D:農地Info].大字ID, V_大字.名称, V_大字.大字 ORDER BY [D:農地Info].大字ID;")

            Dim A1 As Integer = 0
            Dim A2 As Integer = 0

            For Each pRow As DataRow In pTable.Rows
                A1 += pRow.Item("田面計")
                A2 += pRow.Item("畑面計")
            Next
            Dim pNewRow As DataRow = pTable.NewRow
            pNewRow.Item("田面計") = A1
            pNewRow.Item("畑面計") = A2
            pNewRow.Item("農地計") = A1 + A2
            pTable.Rows.Add(pNewRow)

            Dim pPage As CViewToExcel = New CViewToExcel("農用地面積集計", "農用地面積集計", pTable, SysAD.OutputFolder, "農用地面積集計.xls")
            SysAD.MainForm.MainTabCtrl.TabPages.Add(pPage)
        ElseIf Not SysAD.MainForm.MainTabCtrl.TabPages.ContainsKey("農用地面積集計") Then
        End If
        SysAD.MainForm.MainTabCtrl.SelectedTab = SysAD.MainForm.MainTabCtrl.TabPages("農用地面積集計")
    End Sub

    Public Sub sub利用権設定面積()
        If Not SysAD.MainForm.MainTabCtrl.TabPages.ContainsKey("利用権設定面積") Then
            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT CStr([D:農地Info].小作形態) AS 小作形態, Sum([D:農地Info].田面積) AS 田面計, Sum([D:農地Info].畑面積) AS 畑面計, Sum([D:農地Info].樹園地) AS 樹園計, Sum([田面積]+[畑面積]+[樹園地]) AS 農地計 FROM [D:農地Info] WHERE ((([D:農地Info].自小作別)<>1) AND (([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作開始年月日)<#4/1/2013#)) GROUP BY [D:農地Info].小作形態 HAVING ((([D:農地Info].小作形態) Is Not Null));")
            Dim A1 As Integer = 0
            Dim A2 As Integer = 0
            Dim A3 As Integer = 0

            For Each pRow As DataRow In pTable.Rows
                A1 += pRow.Item("田面計")
                A2 += pRow.Item("畑面計")
                A3 += pRow.Item("樹園計")
                Select Case pRow.Item("小作形態")
                    Case "0" : pRow.Item("小作形態") = "不明"
                    Case "1" : pRow.Item("小作形態") = "賃貸借"
                    Case "2" : pRow.Item("小作形態") = "使用貸借"
                    Case Else
                        pRow.Item("小作形態") = "不明"
                End Select
            Next
            Dim pNewRow As DataRow = pTable.NewRow
            pNewRow.Item("田面計") = A1
            pNewRow.Item("畑面計") = A2
            pNewRow.Item("樹園計") = A3
            pNewRow.Item("農地計") = A1 + A2 + A3
            pTable.Rows.Add(pNewRow)


            Dim pPage As CViewToExcel = New CViewToExcel("利用権設定面積", "利用権設定面積", pTable, SysAD.OutputFolder, "利用権設定面積.xls")
            SysAD.MainForm.MainTabCtrl.TabPages.Add(pPage)
        ElseIf Not SysAD.MainForm.MainTabCtrl.TabPages.ContainsKey("利用権設定面積") Then
        End If
        SysAD.MainForm.MainTabCtrl.SelectedTab = SysAD.MainForm.MainTabCtrl.TabPages("利用権設定面積")
    End Sub

    Public Sub sub農業改善計画認定別経営面積()
        If Not SysAD.MainForm.MainTabCtrl.TabPages.ContainsKey("農業改善計画認定別経営面積") Then
            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, V_農業改善計画認定項目.名称, V_行政区.行政区, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].住所, Sum(IIf([自小作別]=0,[田面積],0)) AS 田自作, Sum(IIf([自小作別]=0,[畑面積],0)) AS 畑自作, Sum(IIf([自小作別]<>0,[田面積],0)) AS 田借受, Sum(IIf([自小作別]<>0,[畑面積],0)) AS 畑借受, Sum(V_農地.田面積) AS 田合計, Sum(V_農地.畑面積) AS 畑合計 FROM (V_行政区 INNER JOIN (V_農地 INNER JOIN [D:個人Info] ON V_農地.耕作者ID = [D:個人Info].ID) ON V_行政区.ID = [D:個人Info].行政区ID) INNER JOIN V_農業改善計画認定項目 ON [D:個人Info].農業改善計画認定 = V_農業改善計画認定項目.ID WHERE ((([D:個人Info].農業改善計画認定)>0) AND ((V_農地.農地状況)<20)) GROUP BY [D:個人Info].ID, V_農業改善計画認定項目.名称, V_行政区.行政区, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].住所;")

            Dim pPage As CViewToExcel = New CViewToExcel("農業改善計画認定別経営面積", "農業改善計画認定別経営面積", pTable, SysAD.OutputFolder, "農業改善計画認定別経営面積.xls")
            SysAD.MainForm.MainTabCtrl.TabPages.Add(pPage)
        ElseIf Not SysAD.MainForm.MainTabCtrl.TabPages.ContainsKey("農業改善計画認定別経営面積") Then
        End If
        SysAD.MainForm.MainTabCtrl.SelectedTab = SysAD.MainForm.MainTabCtrl.TabPages("農業改善計画認定別経営面積")
    End Sub
    Public Sub sub検索フリガナ修正()

        If MsgBox("検索フリガナを初期化します", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim p変換 As New C検索フリガナ修正
            p変換.Execute()

            My.Application.DoEvents()

            If MessageBox.Show("終了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) = DialogResult.OK Then
            End If
        End If

    End Sub

    Public Class C検索フリガナ修正
        Inherits HimTools2012.clsAccessor

        Public Sub New()
            Me.Start(True, True)
            SysAD.MainForm.Focus()
        End Sub

        Public Overrides Sub Execute()
            'SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID SET [D:世帯Info].世帯主ID = [D:個人Info].[ID] WHERE ((([D:世帯Info].世帯主ID)=0) AND (([D:個人Info].続柄1)=1) AND (([D:個人Info].住民区分)=0));")
            'SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:個人Info] ( ID, 世帯ID, 氏名, フリガナ, 住所, 住民区分, 続柄1, 続柄2, 続柄3, 性別, 生年月日 ) SELECT M_住民情報.ID, M_住民情報.世帯No, M_住民情報.氏名, M_住民情報.[フリガナ], M_住民情報.住所, M_住民情報.住民区分, M_住民情報.続柄, M_住民情報.続柄2, M_住民情報.続柄3, M_住民情報.性別, M_住民情報.生年月日 FROM (([D:個人Info] AS [D:個人Info_1] INNER JOIN [D:世帯Info] ON [D:個人Info_1].世帯ID = [D:世帯Info].ID) LEFT JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) INNER JOIN M_住民情報 ON [D:世帯Info].世帯主ID = M_住民情報.ID GROUP BY [D:個人Info_1].住民区分, [D:個人Info_1].選挙権の有無, [D:個人Info].ID, M_住民情報.ID, M_住民情報.世帯No, M_住民情報.氏名, M_住民情報.[フリガナ], M_住民情報.住所, M_住民情報.住民区分, M_住民情報.続柄, M_住民情報.続柄2, M_住民情報.続柄3, M_住民情報.性別, M_住民情報.生年月日 HAVING ((([D:個人Info_1].住民区分)=0) AND (([D:個人Info_1].選挙権の有無)=True) AND (([D:個人Info].ID) Is Null));")


            Me.Message = "データベースにアクセスします。"
            Dim pTBL個人 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("Select * From [D:個人Info]")

            Me.Message = "変更を開始します。"

            Dim sB As New System.Text.StringBuilder
            Me.Maximum = pTBL個人.Rows.Count
            Me.Value = 0

            For Each pRow As DataRow In pTBL個人.Rows
                Me.Value += 1
                If pRow.Item("検索フリガナ").ToString <> Replace(pRow.Item("フリガナ").ToString, " ", "") Then
                    sB.AppendLine(String.Format("UPDATE [D:個人INfo] SET [検索フリガナ]='{0}' WHERE [ID]={1}", Replace(pRow.Item("フリガナ").ToString, " ", ""), pRow.Item("ID")))
                End If
                If Me._Cancel Then
                    Exit For
                End If
                If sB.Length > 64 Then
                    Me.Message = String.Format("変更中({0}/{1})...", Me.Value, Me.Maximum)
                    My.Application.DoEvents()
                    SysAD.DB(sLRDB).ExecuteSQL(sB.ToString)
                    sB.Clear()
                End If
            Next
            If Not Me._Cancel AndAlso sB.Length > 0 Then
                SysAD.DB(sLRDB).ExecuteSQL(sB.ToString)
                sB.Clear()
            End If
        End Sub
    End Class

    Public Overridable Function Get地区情報(ByVal s住所 As String) As String
        Return "-"
    End Function

    Public Sub subCSV出力()
        If MsgBox("フェーズ1CSV出力を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim pデータ出力 As New CF1データ出力()
            My.Application.DoEvents()
            If MessageBox.Show("終了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) = DialogResult.OK Then
            End If
        End If
    End Sub

    Public Sub subCSV出力_カスタム()

        If Not SysAD.MainForm.MainTabCtrl.ExistPage("公開用CSV出力(農地一覧)") Then
            SysAD.MainForm.MainTabCtrl.AddPage(New CSV農地一覧.COutPutCSV農地一覧)
        End If
    End Sub

    Public Sub subCSV出力2()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("フェーズ２機能一覧", True) Then
            SysAD.MainForm.MainTabCtrl.AddPage(New CフェーズMainPage())
        End If
    End Sub

    Public Sub subF2ErrorCheck()
        'SysAD.MainForm.MainTabCtrl.ExistPage("フェーズ２エラーチェック", True, GetType(CF2ErrorCheck))
    End Sub

    Public Sub sub全農地CSV出力()
        If MsgBox("全農地CSV出力を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim pデータ出力 As New C全農地CSV出力()
            My.Application.DoEvents()
            If MessageBox.Show("終了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) = DialogResult.OK Then
            End If
        End If

    End Sub

    Public Sub subCSV農地利用状況調査()
        If MsgBox("農地利用状況調査CSV出力を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim pデータ出力 As New COutPutCSV農地利用状況調査()
            My.Application.DoEvents()
            If MessageBox.Show("終了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) = DialogResult.OK Then
            End If
        End If

    End Sub

    Public Sub sub権利移動借賃データ出力()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("権利移動借賃データ出力一覧", True) Then
            SysAD.MainForm.MainTabCtrl.AddPage(New C権利移動借賃MainPage())
        End If
    End Sub

    Public Sub sub農地期間満了の終了()
        Dim nID As Integer = 0
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].* AS D, [D:個人Info].氏名 FROM [D:農地Info] LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID WHERE ((([D:農地Info].自小作別) Is Not Null And ([D:農地Info].自小作別)<>0) AND (([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作終了年月日)<Now())) OR ((([D:農地Info].自小作別) Is Not Null And ([D:農地Info].自小作別)<>0) AND (([D:農地Info].小作地適用法)=1) AND (([D:農地Info].小作終了年月日)<Now()) AND (([D:農地Info].小作形態)=2));")
        'App農地基本台帳.TBL農地.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].自小作別) Is Not Null And ([D:農地Info].自小作別)<>0) AND (([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作終了年月日)<Now())) OR ((([D:農地Info].自小作別) Is Not Null And ([D:農地Info].自小作別)<>0) AND (([D:農地Info].小作地適用法)=1) AND (([D:農地Info].小作形態)=2) AND (([D:農地Info].小作終了年月日)<Now()));"))

        For Each pRow As DataRow In pTBL.Rows
            nID = pRow.Item("ID")
            If Val(pRow.Item("小作地適用法").ToString) = 1 Then
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 内容, 異動日, 更新日, 入力日, 関係者A, 関係者B, 異動事由 ) SELECT [D:農地Info].ID, [氏名] & 'への農地法使用貸借[' & Format$([D:農地Info].[小作開始年月日],'gggee\年mm\月dd\日') & ']～[' & Format$([D:農地Info].[小作終了年月日],'gggee\年mm\月dd\日') & ']を期間満了で終了。' AS 式1, [D:農地Info].小作終了年月日 AS 式2, Now() AS 式3, Now() AS 式4, [D:農地Info].所有者ID, [D:農地Info].借受人ID, 10201 AS 式5 FROM [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].借受人ID = [D:個人Info].ID WHERE ((([D:農地Info].ID)=" & nID & ") AND (([D:農地Info].自小作別)<>0));")
            Else
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 内容, 異動日, 更新日, 入力日, 関係者A, 関係者B, 異動事由 ) SELECT [D:農地Info].ID, [氏名] & 'への利用権設定[' & Format$([D:農地Info].[小作開始年月日],'gggee\年mm\月dd\日') & ']～[' & Format$([D:農地Info].[小作終了年月日],'gggee\年mm\月dd\日') & ']を期間満了で終了。' AS 式1, [D:農地Info].小作終了年月日 AS 式2, Now() AS 式3, Now() AS 式4, [D:農地Info].所有者ID, [D:農地Info].借受人ID, 10201 AS 式5 FROM [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].借受人ID = [D:個人Info].ID WHERE ((([D:農地Info].ID)=" & nID & ") AND (([D:農地Info].自小作別)<>0));")
            End If

            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:農地Info] AS [D:農地Info_1] ON ([D:農地Info].小作終了年月日 = [D:農地Info_1].小作終了年月日) AND ([D:農地Info].小作開始年月日 = [D:農地Info_1].小作開始年月日) AND ([D:農地Info].小作地適用法 = [D:農地Info_1].小作地適用法) AND ([D:農地Info].借受人ID = [D:農地Info_1].借受人ID) AND ([D:農地Info].自小作別 = [D:農地Info_1].自小作別) SET [D:農地Info_1].自小作別 = 0 WHERE ((([D:農地Info].ID)=" & nID & ") AND (([D:農地Info].自小作別)<>0));")
        Next
    End Sub
End Class

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C調査日設定
    Inherits HimTools2012.InputSupport.CInputSupport

    Public Sub New()
        MyBase.New(Nothing)
    End Sub

    Public Enum enum調査日種類
        利用状況調査 = 0
        利用意向調査 = 1
    End Enum

    Public Enum enum設定条件
        空白のみ = 0
        全て = 1
    End Enum


    Private mvar調査日種類 As enum調査日種類 = enum調査日種類.利用状況調査
    <PropertyOrderAttribute(0)> <Category("01 調査日種類")> <Description("調査日種類")>
    Public Property 調査日種類 As enum調査日種類
        Get
            Return mvar調査日種類
        End Get
        Set(value As enum調査日種類)
            mvar調査日種類 = value
        End Set
    End Property
    <PropertyOrderAttribute(1)> <Category("02 調査日")> <Description("調査日")>
    Public Property 調査日 As DateTime = Now.Date

    <PropertyOrderAttribute(2)> <Category("03 設定条件")> <Description("設定条件")>
    Public Property 設定条件 As enum設定条件 = 0


End Class

Public Class CTabPage農家一覧
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarInsideTabCtrl As HimTools2012.controls.TabControlBase
    Private WithEvents mvarMakeExcel As ToolStripButton
    Private WithEvents mvarGrid As HimTools2012.controls.DataGridViewWithDataView

    Public Sub New()
        MyBase.New(True, True, "農家一覧", "農家一覧")
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_行政区.名称 AS 自治会, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].住所, [D:個人Info].生年月日, [D:個人Info].郵便番号, [D:個人Info].電話番号, Count(V_農地.ID) AS 耕作農地数, Count(V_農地_1.ID) AS 所有農地数, Count([V_農地].[ID])-Count([V_農地_1].[ID]) AS 貸借貸付差分 FROM V_農地 AS V_農地_1 RIGHT JOIN (V_農地 RIGHT JOIN (([D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) INNER JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) ON V_農地.耕作世帯ID = [D:世帯Info].ID) ON V_農地_1.所有世帯ID = [D:世帯Info].ID WHERE ((([D:個人Info].住民区分)=0 Or ([D:個人Info].住民区分)=0)) GROUP BY V_行政区.名称, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].住所, [D:個人Info].生年月日, [D:個人Info].郵便番号, [D:個人Info].電話番号 HAVING (((Count(V_農地.ID))>0)) OR (((Count(V_農地_1.ID))>0)) ORDER BY [D:個人Info].[フリガナ];")
        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView(pTBL, "", "")
        Me.ControlPanel.Add(mvarGrid)
        mvarGrid.Createエクセル出力Ctrl(Me.ToolStrip)
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property
End Class
