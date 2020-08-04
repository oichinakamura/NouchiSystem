
Imports HimTools2012
Imports System.ComponentModel

Public Class C農地基本台帳 : Inherits HimTools2012.System管理.CApplication
    Private mvarMenu As ToolStripDropDownButton

    Private masterViews As New Dictionary(Of String, DataView)
    Private mvarListColumnDesign As CListColumnDesign

#Region "データテーブル"
    Private mvarTBL農地履歴 As DataTable

    Private mvarTBL世帯 As CTBL世帯
    Private mvarTBL個人 As CTBL個人
    Private mvarTBL申請 As CTBL申請
    Private mvarTBL農地 As CTBL農地

    Private mvarTBL転用農地 As CTBL転用農地
    Private mvarTBL削除農地 As CTBL削除農地
    Private mvarTBL固定情報 As CTBL固定情報
    Private mvarTBL世帯営農 As CTBL世帯営農

    Private mvarTBL削除個人 As CTBL削除個人

    Private mvarTBL土地系図 As DataTable
    Private mvarTBL土地履歴 As CTBL土地履歴
#End Region

    Public Sub New()
        MyBase.New()

        mvarDSet = New HimTools2012.Data.CDataSet("農地台帳")
        Dim pTableCast As New Dictionary(Of String, HimTools2012.Data.TableExCast)

        pTableCast.Add("M_固定情報", New HimTools2012.Data.TableExCast("M_固定情報", GetType(CTBL固定情報)))

        Dim strXSD As String = My.Resources.Resource1.DSet

        Try
            mvarDSet.ReadXmlSchema(New IO.StringReader(strXSD), pTableCast)
        Catch ex As Exception
            Stop
        End Try
    End Sub

#Region "キャッシュテーブル"
    Public ReadOnly Property TBL個人 As CTBL個人
        Get
            Return mvarTBL個人
        End Get
    End Property

    <Browsable(False)>
    Public Overrides ReadOnly Property MyNamespace As String
        Get
            Return sLRDB
        End Get
    End Property

    Public ReadOnly Property TBL申請 As CTBL申請
        Get
            Return mvarTBL申請
        End Get
    End Property

    Public ReadOnly Property TBL農地() As CTBL農地
        Get
            Return mvarTBL農地
        End Get
    End Property

    Public ReadOnly Property TBL世帯() As CTBL世帯
        Get
            Return mvarTBL世帯
        End Get
    End Property

    Public ReadOnly Property TBL削除農地() As CTBL削除農地
        Get
            If mvarTBL削除農地 Is Nothing Then
                mvarTBL削除農地 = New CTBL削除農地(DSet, DSet.Tables("D_削除農地"))
            End If
            Return mvarTBL削除農地
        End Get
    End Property

    Public ReadOnly Property TBL削除個人() As CTBL削除個人
        Get
            If mvarTBL削除個人 Is Nothing Then
                mvarTBL削除個人 = New CTBL削除個人(DSet, DSet.Tables("D_削除個人"))
            End If
            Return mvarTBL削除個人
        End Get
    End Property

    Public ReadOnly Property TBL土地履歴 As CTBL土地履歴
        Get

            Return mvarTBL土地履歴
        End Get
    End Property

    Public ReadOnly Property TBL転用農地() As CTBL転用農地
        Get
            If mvarTBL転用農地 Is Nothing Then
                mvarTBL転用農地 = New CTBL転用農地(DSet, DSet.Tables("D_転用農地"))
            End If
            Return mvarTBL転用農地
        End Get
    End Property
    Public ReadOnly Property TBL固定情報() As CTBL固定情報
        Get
            If mvarTBL固定情報 Is Nothing Then
                mvarTBL固定情報 = DSet.Tables("M_固定情報")
            End If
            Return mvarTBL固定情報
        End Get
    End Property

    Private mvar筆情報 As CTBL筆情報
    Public ReadOnly Property TBL筆情報() As CTBL筆情報
        Get
            If mvar筆情報 Is Nothing Then
                mvar筆情報 = New CTBL筆情報(DSet, SysAD.DB(s地図情報).GetTableBySqlSelect("SELECT * FROM [D:LotProperty] WHERE [ID]=0"))
                mvar筆情報.PrimaryKey = {mvar筆情報.Columns("ID")}
            End If
            Return mvar筆情報
        End Get
    End Property
    Private mvar点情報 As DataTable
    Public ReadOnly Property TBL点情報() As DataTable
        Get
            If mvar点情報 Is Nothing Then
                mvar点情報 = SysAD.DB(s地図情報).GetTableBySqlSelect("SELECT * FROM [D:XY] WHERE [ID]=0")
                mvar点情報.PrimaryKey = {mvar点情報.Columns("ID")}
            End If
            Return mvar点情報
        End Get
    End Property


    Public ReadOnly Property TBL世帯営農() As CTBL世帯営農
        Get
            Return mvarTBL世帯営農
        End Get
    End Property



    Public ReadOnly Property TBL土地系図 As DataTable
        Get
            Return mvarTBL土地系図
        End Get
    End Property

    Public ReadOnly Property TBL大字 As DataTable
        Get
            Return DSet.Tables("V_大字")
        End Get
    End Property
    Public ReadOnly Property TBL小字 As DataTable
        Get
            Return DSet.Tables("V_小字")
        End Get
    End Property
    Public ReadOnly Property TBL地目 As DataTable
        Get
            Return DSet.Tables("V_地目")
        End Get
    End Property
    Public ReadOnly Property TBL現況地目 As DataTable
        Get
            Return DSet.Tables("V_現況地目")
        End Get
    End Property


    Public ReadOnly Property TBL続柄 As DataTable
        Get
            Return DSet.Tables("V_続柄")
        End Get
    End Property

#End Region

    Public ReadOnly Property ListColumnDesign As CListColumnDesign
        Get
            If mvarListColumnDesign Is Nothing Then
                mvarListColumnDesign = New CListColumnDesign()
            End If
            Return mvarListColumnDesign
        End Get
    End Property

    Public Function DataTableCache(ByVal sTable As String, ByVal sWhere As String, ByVal sViewWhere As String, ByVal sSort As String, ByVal sKey As String, ByVal sIcon As String, ParamArray sPrimaryKey() As String) As DataView
        If DSet.Tables.Contains(sTable) Then
            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT *," & sKey & " AS [Key]," & sIcon & " AS [アイコン] FROM [" & sTable & "] WHERE " & sWhere)

            DSet.Tables(sTable).Merge(pTable, False, MissingSchemaAction.AddWithKey)
        Else
            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT *," & sKey & " AS [Key]," & sIcon & " AS [アイコン] FROM [" & sTable & "] WHERE " & sWhere)
            pTable.TableName = sTable

            pTable.PrimaryKey = New DataColumn() {pTable.Columns(sPrimaryKey(0))}
            DSet.Tables.Add(pTable)

        End If
        Return New DataView(DSet.Tables(sTable), sViewWhere, sSort, DataViewRowState.CurrentRows)
    End Function


    Public Sub Command(ByVal sCommand As String, ByVal ParamArray sParam() As String)
        Dim st市町村種別 As String = "" 'SysAD.DB(sLRDB).DBProperty("市町村種別")

        Select Case sCommand
            Case "同一地番の整合性" : sub同一地番の整合性()
            Case "同一地番一覧"

                SysAD.page農家世帯.中央Tab.AddPage(New Cカスタムリスト("同一地番一覧",
                        "同一地番一覧", "SELECT '農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, [D:個人Info].氏名, V_農地.登記簿面積, V_農地.実面積, V_農地.一部現況, V_農地.所在, V_農地.自小作別 FROM V_農地 INNER JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID WHERE (((V_農地.大字ID) In (SELECT [大字ID] FROM [V_農地] As Tmp GROUP BY [大字ID],[地番] HAVING Count(*)>1  And [地番] = [V_農地].[地番]))) ORDER BY V_農地.土地所在, V_農地.実面積 DESC;",
                        "土地所在;土地所在;氏名;氏名;@登記簿面積;登記簿面積;@実面積;実面積;一部現況;一部現況"
                    )
                )
            Case "住記検索フリガナ初期化"
                'Sub検索フリガナ作成()
                'Case "転用履歴の整合性"
                '    Dim Rs As NK97.RecordsetEx
                '    Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT D_申請.ID,D_申請.農地リスト FROM D_申請 WHERE (((D_申請.法令)=40) AND ((D_申請.状態)=2));", 0, , Me)
                '    Do Until Rs.EOF
                '        St = Replace(Replace(Rs.Value("農地リスト"), "転用農地.", ""), ";", ",")
                '        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( 更新日,LID, 異動事由, 内容, 申請ID ) SELECT Date() AS 式1, D_転用農地.ID, 10040 AS 式2, '4条転用許可' AS 式3, " & Rs.ID & " AS 式1 FROM D_転用農地 LEFT JOIN D_土地履歴 ON D_転用農地.ID = D_土地履歴.ID WHERE (((D_転用農地.ID) In (" & St & ")));")
                '        Rs.MoveNext()
                '    Loop
                '    Rs.CloseRs()

                '    Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT D_申請.ID,D_申請.農地リスト FROM D_申請 WHERE (((D_申請.法令)=50 Or (D_申請.法令)=51) AND ((D_申請.状態)=2));", 0, , Me)
                '    Do Until Rs.EOF
                '        St = Replace(Replace(Rs.Value("農地リスト"), "転用農地.", ""), ";", ",")
                '        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( 更新日,LID, 異動事由, 内容, 申請ID ) SELECT Date() AS 式1, D_転用農地.ID, 10050 AS 式2, '5条転用許可' AS 式3, " & Rs.ID & " AS 式1 FROM D_転用農地 LEFT JOIN D_土地履歴 ON D_転用農地.ID = D_土地履歴.ID WHERE (((D_転用農地.ID) In (" & St & ")));")

                '        Rs.MoveNext()
                '    Loop
                '    Rs.CloseRs()

                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE (D_土地履歴 AS D_土地履歴_1 INNER JOIN D_申請 ON D_土地履歴_1.申請ID = D_申請.ID) INNER JOIN D_転用農地 ON D_土地履歴_1.LID = D_転用農地.ID SET D_土地履歴_1.更新日 = [D_申請].[公告年月日], D_土地履歴_1.異動日 = [許可年月日], D_土地履歴_1.関係者A = [申請者A], D_土地履歴_1.関係者B = [申請者B] WHERE (((D_土地履歴_1.異動事由)=10050 Or (D_土地履歴_1.異動事由)=10040));")
                '    FncNet.CopyAbleMessage("終了")
                'Case "世帯数調べ"
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:世帯Info] SET [D:世帯Info].農地との関連 = False WHERE ((([D:世帯Info].農地との関連)=True));")
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN ([D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].ID = [D:個人Info].世帯ID) ON [D:農地Info].所有者ID = [D:個人Info].ID SET [D:世帯Info].農地との関連 = True WHERE ((([D:世帯Info].農地との関連)=False));")
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN ([D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].ID = [D:個人Info].世帯ID) ON [D:農地Info].借受人ID = [D:個人Info].ID SET [D:世帯Info].農地との関連 = True WHERE ((([D:世帯Info].農地との関連)=False));")
                '    St = SysAD.DB(sLRDB).DBProperty("市町村ID")
                '    If St = "" Then St = "-999"
                '    mvarPDW.SQLListview.SQLListviewCHead("SELECT 'X.' & IIf([市町村ID]=" & St & ",'在住','不在') AS [KEY],IIf([市町村ID]=" & St & ",'在住','不在') AS 地主の所在, Count([D:世帯Info].ID) AS 世帯数 FROM [D:世帯Info] WHERE ((([D:世帯Info].農地との関連) = True)) GROUP BY 'X.' & IIf([市町村ID]=" & St & ",'在住','不在'),IIf([市町村ID]=" & St & ",'在住','不在');", "地主の所在;地主の所在;@世帯数;世帯数", "世帯数調べ")
                'Case "世帯と農地の整合チェック" : sub農家世帯設定()
                'Case "貸し手・借り手同一者エラー農地一覧" : view農地List("[所有者ID]=[借受人ID] AND [自小作別]>0", "貸し手・借り手同一者エラー農地一覧") '/** 農地0001-0001-0002
                'Case "管理者＝所有者エラー農地一覧" : view農地List("[管理世帯ID]=[所有世帯ID] AND [所有者ID]=[管理者ID]", "管理者＝所有者エラー農地一覧") '/** 農地0001-0001-0002
                'Case "管理者＝所有者エラー農地解除"
                '    If MsgBox("管理者＝所有者である不正情報を解除しますか", vbYesNo) = vbYes Then SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].管理世帯ID = 0,[D:農地Info].管理者ID = 0 WHERE ([管理世帯ID]=[所有世帯ID] AND [所有者ID]=[管理者ID]);")
                'Case "貸し手・借り手同一者解約"
                '    If MsgBox("貸し手・借り手が同一である不正情報を解約しますか", vbYesNo) = vbYes Then SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].自小作別 = 0 WHERE ((([D:農地Info].自小作別)<>0) AND (([D:農地Info].所有者ID)=[借受人ID]));")
                'Case "住記情報の転送"
                '    If Fnc.InputText("重要な変更", "パスワードを入力してください", "", 0, vbIMEModeOff) = "Avail" Then
                '        mvarPDW.WaitMessage = "住記情報のバックアップ転送"
                '        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].[住記世帯番号]=[世帯ID], [D:個人Info].住記続柄1 = [続柄1], [D:個人Info].住記続柄2 = [続柄2], [D:個人Info].住記続柄3 = [続柄3], [D:個人Info].住記住所 = [住所];")
                '        mvarPDW.WaitMessage = ""
                '    End If
                'Case "所有／管理世帯の整合性" : SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID SET [D:農地Info].所有世帯ID = [世帯ID] WHERE ((([D:農地Info].所有世帯ID)=0 Or ([D:農地Info].所有世帯ID) Is Null) AND (([D:個人Info].世帯ID)<>0));", "更新中....")
                'Case "借受世帯の整合性" : SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].借受人ID = [D:個人Info].ID SET [D:農地Info].借受世帯ID = [世帯ID] WHERE ((([D:農地Info].借受世帯ID)=0 Or ([D:農地Info].借受世帯ID) Is Null) AND (([D:個人Info].世帯ID)<>0));", "更新中....")
                'Case "データベースの更新" : MsgBox("データベースが見つかりません", vbCritical)
                'Case "農地面積の整合性"
                '    mvarPDW.WaitMessage = "農地面積の整合中"
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].田面積 = 0 WHERE ((([D:農地Info].現況地目) Not In (" & SysAD.DB(sLRDB).DBProperty("田地目") & ")));")
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].畑面積 = 0 WHERE ((([D:農地Info].現況地目) Not In (" & SysAD.DB(sLRDB).DBProperty("畑地目") & ")));")
                '    MsgBox("終了しました.")
                '    mvarPDW.WaitMessage = ""
                'Case "ファイルメニュー"
                '    If SysAD.ApplicationDLL.User.n権利 = 0 Then
                '        CDataviewSK_DoCommand2 = ";"
                '    Else
                '        CDataviewSK_DoCommand2 = ">取込;基幹データの取込;<;>出力;全農地データの出力;<;-;>修正;フリガナ修正;住所修正;<" & IIf(Len(St), ";>エクスポート;" & St & "<", "") & ";"
                '    End If

                'Case "住記テキストの設定"
                '    St = Fnc.OpenFileDlg("住記テキストの指示", "テキストファイル|*.csv;*.txt", SysAD.DB(sLRDB).DBProperty("住記テキストの指示"))
                '    If FncNet.FileExists(St) Then SysAD.DB(sLRDB).DBProperty("住記テキストの指示") = St
                'Case "エラーリスト作成"
                '    Select Case FncNet.OptionSelect("所有者;登記面積;登記地目;地番不一致", "固定資産との突合せをします。")
                '        Case "所有者" : mvarPDW.SQLListview.SQLListviewCHead("SELECT '所有者が違う' AS [事由],'農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, V_農地.所有者ID, '農地台帳所有者：' & [D:個人Info].氏名 AS [台帳内容], M_固定情報.所有者ID, '固定情報所有者：' & M_住民情報.氏名 AS [固定内容] FROM ((V_農地 INNER JOIN M_固定情報 ON (V_農地.大字ID = M_固定情報.大字ID) AND (V_農地.地番 = M_固定情報.地番)) LEFT JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) INNER JOIN M_住民情報 ON M_固定情報.所有者ID = M_住民情報.ID WHERE (((V_農地.所有者ID)<>[M_固定情報].[所有者ID]));", "土地の所在;土地所在;事由;事由;農地台帳所有者;台帳内容;固定情報所有者;固定内容", "土地のエラーリスト(所有者)")
                '        Case "登記地目" : mvarPDW.SQLListview.SQLListviewCHead("SELECT '登記地目が違う' AS 事由, '農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, V_農地.登記簿地目, '農地台帳登記地目：' & [V_地目].[名称] AS 台帳内容, '固定情報登記地目：' & [V_地目2].[名称] AS 固定内容 FROM ((V_農地 INNER JOIN M_固定情報 ON (V_農地.地番 = M_固定情報.地番) AND (V_農地.大字ID = M_固定情報.大字ID)) INNER JOIN V_地目 AS V_地目2 ON M_固定情報.登記地目 = V_地目2.ID) INNER JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID WHERE (((V_農地.登記簿地目)<>[M_固定情報].[登記地目]));", "土地の所在;土地所在;事由;事由;農地台帳地目;台帳内容;固定情報地目;固定内容", "土地のエラーリスト(地目)")
                '        Case "登記面積" : mvarPDW.SQLListview.SQLListviewCHead("SELECT '登記面積が違う' AS 事由, '農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, '農地台帳登記面積：' & [登記簿面積] AS 台帳内容, '固定情報登記面積：' & [登記面積] AS 固定内容 FROM V_農地 INNER JOIN M_固定情報 ON (V_農地.地番 = M_固定情報.地番) AND (V_農地.大字ID = M_固定情報.大字ID) WHERE (((V_農地.登記簿面積)<>[M_固定情報].[登記面積]));", "土地の所在;土地所在;事由;事由;農地台帳登記面積;台帳内容;固定情報登記面積;固定内容", "土地のエラーリスト(登記面積)")
                '        Case "地番不一致" : mvarPDW.SQLListview.SQLListviewCHead("SELECT '同一地番無し' AS 事由, '農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, M_固定情報.ID, '農地台帳：' & [V_農地].[地番] AS 台帳内容, '固定情報：同一地番なし' AS 固定内容 FROM V_農地 LEFT JOIN M_固定情報 ON (V_農地.地番 = M_固定情報.地番) AND (V_農地.大字ID = M_固定情報.大字ID) WHERE (((M_固定情報.ID) Is Null));", "土地の所在;土地所在;事由;事由;台帳内容;台帳内容;固定内容;固定内容", "土地のエラーリスト(地番不一致)")
                '    End Select
                'Case "土地履歴と申請の不一致" : List土地履歴と申請の不一致()
                'Case "世帯自動生成"
                '    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:世帯Info] ( 世帯主ID, ID, 行政区ID ) " & _
                '    "SELECT [D:個人Info].ID, [D:個人Info].世帯ID, [D:個人Info].行政区ID FROM [D:個人Info] LEFT JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID " & _
                '    "WHERE ((([D:世帯Info].ID) Is Null) AND (([D:個人Info].世帯ID)>0) AND (([D:個人Info].続柄1)=" & SysAD.DB(sLRDB).DBProperty("世帯主続柄コード") & ") AND (([D:個人Info].住民区分)=" & SysAD.DB(sLRDB).DBProperty("記録住民コード") & "));")
                'Case "土地テキストの設定"
                '    St = Fnc.OpenFileDlg("土地テキストの指示", "テキストファイル|*.csv;*.txt;*.DAT", SysAD.DB(sLRDB).DBProperty("土地テキストの指示"))
                '    If FncNet.FileExists(St) Then
                '        SysAD.DB(sLRDB).DBProperty("土地テキストの指示") = St
                '    End If
                '    '            Case "世帯リスト": SysAD.MDIForm.ListViewCustom "農地基本台帳", "世帯リスト", ADApp.Common.AppData("世帯リスト")     '
                'Case "世帯員リスト" : SysAD.MDIForm.ListViewCustom("農地基本台帳", "世帯員リスト", ADApp.Common.AppData("世帯員リスト"))
                'Case "申請書受付中リスト" : SysAD.MDIForm.ListViewCustom("農地基本台帳", "申請受付中リスト表示", ADApp.Common.AppData("申請受付中リスト表示"))
                'Case "申請書審査中リスト" : SysAD.MDIForm.ListViewCustom("農地基本台帳", "申請審査中リスト表示", ADApp.Common.AppData("申請審査中リスト表示"))
                'Case "申請書許可済リスト" : SysAD.MDIForm.ListViewCustom("農地基本台帳", "申請許可済リスト表示", ADApp.Common.AppData("申請許可済リスト表示"))
                'Case "農地リスト"
                '    SysAD.MDIForm.ListViewCustom("農地基本台帳", "農地リスト", ADApp.Common.AppData("農地リスト"))
                'Case "全農地データの出力"
                '    Call NCSV()
                'Case "小字テーブルの正規化"
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] INNER JOIN M_BASICALL ON [D:農地Info].小字ID = M_BASICALL.ID SET M_BASICALL.ParentKey = '大字.' & [大字ID] WHERE (((M_BASICALL.Class)='小字'));")
                '    '/*************選挙関連******************/
                'Case "選挙権確定者" : mvarPDW.PrintGo(ADApp.ObjectMan.GetObject("選挙権確定者"))
                'Case "入場整理券" : mvarPDW.PrintGo(ADApp.ObjectMan.GetObject("選挙入場整理券"))
                '    '/****土地管理情報データベースの設定******/
                'Case "土地管理情報データベースの設定"
                'Case "市" : SysAD.DB(sLRDB).SetDirectData("S_システムデータ", "DATA", "市", , "市町村種別") : FolderChange()
                'Case "町" : SysAD.DB(sLRDB).SetDirectData("S_システムデータ", "DATA", "町", , "市町村種別") : FolderChange()
                'Case "村" : SysAD.DB(sLRDB).SetDirectData("S_システムデータ", "DATA", "村", , "市町村種別") : FolderChange()
                '    '/*************データビュー**************/
                'Case "開く" : SQLListフォルダ("[ParentKey]='" & DVProperty.Key & "'", "農家台帳")
                'Case "閉じる" : ADApp.DataviewCol.Remove(DVProperty.Key)


                'Case "市町村名の設定", "会長の設定", "農業委員会会長の設定" : ADApplicationDLL_DoCommand("プロパティ")
                'Case "町外農業者の追加" : CreateFarmer()
                'Case "オンラインヘルプの表示" : mvarPDW.ExplorerBar.Add(ADApp.ObjectMan.GetObject("オンラインヘルプ.0"))
                'Case "世帯番号の確認" : MsgBox("現バージョンでは対応できません")
                'Case "管理情報の表示" : ADApp.DataviewCol.Add(ObjectMan.GetObjectDVP("土地管理.0"))
                'Case "表示メニュー"
                '    St = "非農地申請フォルダ;>一覧の設定;世帯リスト;世帯員リスト;農地リスト;選挙入力リスト;" & _
                '        ">申請書;申請書受付中リスト;申請書審査中リスト;申請書許可済リスト;" & IIf(Val(SysAD.DB(sLRDB).DBProperty("申請書の入力画面")), "標準申請入力;@拡張申請入力", "@標準申請入力;拡張申請入力") & _
                '        ";-;" & IIf(Val(SysAD.DB(sLRDB).DBProperty("申請書の昇順・降順")), "申請書昇順;@申請書降順", "@申請書昇順;申請書降順") & ";<;"
                '    CDataviewSK_DoCommand2 = St
                'Case "申請書昇順" : SysAD.DB(sLRDB).DBProperty("申請書の昇順・降順") = 0
                'Case "申請書降順" : SysAD.DB(sLRDB).DBProperty("申請書の昇順・降順") = -1
                'Case "フリガナ修正" : SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].フリガナ = TRIM(StrConv([フリガナ],8));") : MsgBox("終了")
                'Case "農地の検索" : mvarPDW.ExplorerBar.Add(ADApp.ObjectMan.GetObject("農地検索.1"))
                'Case "農家一覧"
                '    sSelecet = "SELECT '農家.' & [D:世帯INFO].[ID] AS [Key], [D:個人INFO.氏名] AS [名称], 'House' AS Icon, 'House' AS SelectIcon, [D:個人INFO].[住所] AS 住所, [D:世帯INFO].農家番号,[D:個人Info].[フリガナ], M_BASICALL.名称 AS [専兼区分], M_BASICALL.Class"
                '    mvarPDW.SQLListview.SQLListviewCHead(sSelecet & " FROM ([D:世帯INFO] LEFT JOIN [D:個人INFO] ON [D:世帯INFO].世帯主ID = [D:個人INFO].ID) LEFT JOIN (SELECT * FROM [M_BASICALL] WHERE [Class]='専兼区分') AS [M_BASICALL]  ON [D:世帯INFO].兼業区分 = M_BASICALL.ID " & sParam, "%2000,世帯主;名称;@農家番号;農家番号;フリガナ;フリガナ;$専兼区分;専兼区分;%3500,住所;住所", "農家一覧")
                'Case "農地一覧" : view農地List(sParam, "農地一覧") '/** 農地0001-0001-0002
                'Case "死亡名義管理者不明農地一覧" : view農地List("((([D:個人Info].住民区分)=" & SysAD.DB(sLRDB).DBProperty("死亡住民コード") & ") AND ((V_農地.管理世帯ID) Is Null Or (V_農地.管理世帯ID)=0))", "死亡名義管理者不明農地一覧") '/** 農地0001-0001-0002
                'Case "不在地主名義管理者不明農地一覧"
                '    St = Fnc.InputText("不在地主名義管理者不明農地一覧", "住民区分を入力してください")
                '    If Len(St) Then view農地List("[D:個人Info].[住所] LIKE '%" & SysAD.DB(sLRDB).DBProperty("市町村名") & "%' AND ((([D:個人Info].住民区分) IN (" & St & ")) AND ((V_農地.管理世帯ID) Is Null Or (V_農地.管理世帯ID)=0))", "不在地主名義管理者不明農地一覧") '/** 農地0001-0001-0002
                'Case "世帯番号からの検索" : Find世帯With世帯番号(0)
                'Case "世帯番号からの営農情報検索" : Find世帯With世帯番号(1)
                'Case "町外農業者の作成" : CreateFarmer(-1)
                'Case "住民番号からの検索" : Find個人With住民番号()
                'Case "生年月日からの検索" : Find個人With生年月日()
                'Case "農地ＩＤからの検索" : Find農地With農地番号()
                'Case "一筆コードからの検索" : Find農地With固定番号()
                'Case "プロパティ" : mvarPDW.DataviewCol.Add(Me)
                'Case "検索メニュー" : CDataviewSK_DoCommand2 = "農家検索;農地の検索;-;世帯番号からの検索;世帯番号からの営農情報検索;-;住民番号からの検索;生年月日からの検索;該当世帯番号から住民の検索;-;農地ＩＤからの検索;一筆コードからの検索;"
                'Case "該当世帯番号から住民の検索"
                '    St = Fnc.InputText("世帯番号", "世帯番号を入力してください", , 1, 0)
                '    If Not Val(St) = 0 Then SQLList世帯員("[世帯ID]=" & Val(St), "住民番号=" & Val(St) & "の検索", True)
                'Case "農家検索", "農業者検索", "法人の検索" : mvarPDW.ExplorerBar.Add(ADApp.ObjectMan.GetObject("農家検索.1"))

                'Case "基本集計"
                '    If 市内外の調整() Then pPrint = ADApp.ObjectMan.GetObject("基本集計") : pPrint.Printview("")
                'Case "住民区分別集計" : sub住民区分別集計()

                'Case "集落集計参照" : DVProperty.Controls.Value("TX集落集計") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("集落集計定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX集落集計"))
                'Case "基本集計参照" : DVProperty.Controls.Value("TX基本集計") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("基本集計定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX基本集計"))
                'Case "住記テキストの読込開始" : MsgBox("現バージョンでは対応できません")
                'Case "利用権終期農地一覧参照" : DVProperty.Controls.Value("TX利用権終期農地一覧") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("利用権終期農地一覧定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX利用権終期農地一覧"))
                'Case "基幹DBファイル参照" : DVProperty.Controls.Value("TX基幹DBファイル") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("基幹DBファイルの設定", "*.MDB"), DVProperty.Controls.Value("TX基幹DBファイル"))
                'Case "意見書ファイル参照" : DVProperty.Controls.Value("TX意見書ファイル") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("意見書ファイルの設定", "*.XLS"), DVProperty.Controls.Value("TX意見書ファイル"))
                'Case "農業者検索ﾌｫｰﾑ参照" : DVProperty.Controls.Value("TX農業者検索ﾌｫｰﾑ") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("農業者検索ﾌｫｰﾑ定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX農業者検索ﾌｫｰﾑ"))

                'Case "５年以上貸借"
                '    sFrom = Fnc.InputText("基準日", "審査の基準日を入力してください", Date)
                '    If Len(sFrom) = 0 Then Exit Sub
                '    If Not IsDate(sFrom) Then Exit Sub

                '    St = "SELECT '農地.' & [V_農地].ID AS [Key],V_農地.小作終了年月日, V_農地.小作開始年月日, [D:個人Info].氏名 AS 借受者, V_農地.土地所在, [D:個人Info_1].氏名 AS 所有者, IIf([小作地適用法]=1,'農','基') AS 法令, IIf([小作形態]=1,'賃','使') AS 形態, DateDiff('yyyy',[小作開始年月日],[小作終了年月日]) AS 期間 " & _
                '    "FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.借受人ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.所有者ID = [D:個人Info_1].ID " & _
                '    "WHERE ([V_農地].自小作別 > 0) AND ([V_農地].小作終了年月日 > CDate('" & sFrom & "')) AND ([V_農地].小作開始年月日 <= CDate('" & DateSerial(Year(CDate(sFrom)) - 5, Month(CDate(sFrom)), Day(CDate(sFrom))) & "')) " & _
                '    "ORDER BY V_農地.小作終了年月日, [D:個人Info].氏名;"

                '    mvarPDW.SQLListview.SQLListviewCHead(St, "土地所在;土地所在;小作終了年月日;小作終了年月日;小作開始年月日;小作開始年月日;@期間(年);期間;借受者;借受者;所有者;所有者;法令;法令;形態;形態", "")
            Case "３年以上貸借"

                '    sFrom = Fnc.InputText("基準日", "審査の基準日を入力してください", Date)
                '    If Len(sFrom) = 0 Then Exit Sub
                '    If Not IsDate(sFrom) Then Exit Sub

                '    St = "SELECT '農地.' & [V_農地].ID AS [Key],V_農地.小作終了年月日, V_農地.小作開始年月日, [D:個人Info].氏名 AS 借受者, V_農地.土地所在, [D:個人Info_1].氏名 AS 所有者, IIf([小作地適用法]=1,'農','基') AS 法令, IIf([小作形態]=1,'賃','使') AS 形態, DateDiff('yyyy',[小作開始年月日],[小作終了年月日]) AS 期間 " & _
                '        "FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.借受人ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.所有者ID = [D:個人Info_1].ID " & _
                '        "WHERE ([V_農地].自小作別 > 0) AND ([V_農地].小作終了年月日 > CDate('" & sFrom & "')) AND ([V_農地].小作開始年月日 <= CDate('" & DateSerial(Year(CDate(sFrom)) - 3, Month(CDate(sFrom)), Day(CDate(sFrom))) & "')) " & _
                '        "ORDER BY V_農地.小作終了年月日, [D:個人Info].氏名;"

                '    mvarPDW.SQLListview.SQLListviewCHead(St, "土地所在;土地所在;小作終了年月日;小作終了年月日;小作開始年月日;小作開始年月日;@期間(年);期間;借受者;借受者;所有者;所有者;法令;法令;形態;形態", "")


                'Case "住所別集計" : sub住所別()
            Case ""
                '/****************************　取込関連 **************************/

                'Case "農地現況調査について" : mvarPDW.PrintGo(ADApp.ObjectMan.GetObject("農地現況調査"))
                'Case "集落別世帯数一覧" : subSQLList集落別世帯一覧()
                'Case "切り取り", "コピー", "すべて選択" : SysAD.MDIForm.DoEditMenu(sCmd)
                'Case "農地-遊休化", "農地-無断転用"
                '    sDate = Fnc.InputText("調査年月日", "調査年月日を入力してください", Date, 1, 2)
                '    If IsDate(sDate) And Len(sParam) Then
                '        Ar = Split(sParam, ";")
                '        For n = 0 To UBound(Ar)
                '            pDV = ADApp.ObjectMan.GetObject(CStr(Ar(n)))
                '            pDV.DoCommand2(Mid$(sCmd, InStr(sCmd, "-") + 1), sDate)
                '        Next
                '        mvarPDW.SQLListview.Refresh()
                '    End If
                'Case "台帳管理システムについて" : FncNet.About("市町村:" & SysAD.DB(sLRDB).DBProperty("市町村名") & vbCrLf & "DB:")
                'Case "農地-"
                'Case "農地-樹園地へ変換"
                '    Ar = Split(sParam, ";")
                '    For n = 0 To UBound(Ar)
                '        pDV = ADApp.ObjectMan.GetObject(CStr(Ar(n)))
                '        pDV.DoCommand2("樹園地設定")
                '    Next
                '    mvarPDW.SQLListview.Refresh()
                'Case "農地-同一条件の貸借設定"
                '    Dim Rg As NK97.RecordsetEx
                '    St = GetIDList(sParam, "農地")
                '    Ar = Split(St, ",")
                '    Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT Right$('0000000000' & [ID],10) & ':' & [土地所在] & '(' & [小作開始年月日] & '～' & [小作終了年月日] & ')' FROM [V_農地] WHERE [ID] IN (" & St & ")", , , Me)
                '    SS = Val(FncNet.OptionSelect(Rs.GetString(":", ";"), "基準になる農地を選択してください", ""))
                '    SysAD.DB(sLRDB).CloseRs(Rs)
                '    If Val(SS) Then
                '        Rg = SysAD.DB(sLRDB).GetRecordsetEx("SELECT * FROM [D:農地Info] WHERE [ID]=" & SS, , , Me)
                '        If Not Rg.IsNoRecord Then
                '            Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT * FROM [D:農地Info] WHERE [ID] IN (" & St & ")", , , Me)

                '            Do Until Rs.EOF
                '                If Rs.ID = Rg.ID Then
                '                Else
                '                    Rs.Copy(Rg, GetArray("自小作別", "借受世帯ID", "借受人ID", "小作地適用法", "小作形態", "小作開始年月日", "小作終了年月日", "小作料", "小作料単位"))
                '                End If
                '                Rs.MoveNext()
                '            Loop
                '            SysAD.DB(sLRDB).CloseRs(Rs)
                '        End If
                '        mvarPDW.SQLListview.Refresh()
                '        Rg.CloseRs()
                '    End If
                'Case "農地-終了履歴付き強制解約" ' In 農地情報DLL
                '    sDate = Fnc.InputText("終了年月日", "終了年月日を入力してください", Date, 1, 2)
                '    If IsDate(sDate) And Len(sParam) Then
                '        Ar = Split(sParam, ";")
                '        For n = 0 To UBound(Ar)
                '            ID = FncNet.GetKeyCode(Ar(n))
                '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].自小作別 = 0 WHERE [D:農地Info].[ID]=" & ID)
                '            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 異動事由, 異動日, 更新日, 内容 ) VALUES(" & ID & ", 10201 , #" & Format(CDate(sDate), "MM/DD/YYYY") & "# , Date() , '利用権満期終了');")
                '        Next
                '        mvarPDW.SQLListview.Refresh()
                '    End If
                'Case "終了時刻表示"
                '    St = Fnc.InputText("終了時刻設定", "本日の終了時刻を表示します" & vbCrLf & "例、17:30の場合　[17:30]")
                '    If IsDate(St) Then
                '        SysAD.DB(sLRDB).DBProperty("終了時刻") = DateSerial(Year(Now), Month(Now), Day(Now)) + TimeSerial(Hour(CDate(St)), Minute(CDate(St)), 0)
                '    Else
                '        MsgBox("時刻の設定が不正です", vbCritical)
                '    End If
                'Case "農地-解約履歴付き強制解約" ' In 農地情報DLL
                '    sDate = Fnc.InputText("終了年月日", "終了年月日を入力してください", Date, 1, 2)
                '    If IsDate(sDate) And Len(sParam) Then
                '        Ar = Split(sParam, ";")
                '        For n = 0 To UBound(Ar)
                '            ID = FncNet.GetKeyCode(Ar(n))
                '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].自小作別 = 0 WHERE [D:農地Info].[ID]=" & ID)
                '            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 異動事由, 異動日, 更新日, 内容 ) VALUES(" & ID & ", 10202 , #" & Format(CDate(sDate), "MM/DD/YYYY") & "# , Date() , '利用権解約終了');")
                '        Next
                '        mvarPDW.SQLListview.Refresh()
                '    End If
                'Case "ツールメニュー" : CDataviewSK_DoCommand2 = ADApp.Common.SetMenuCmd("ツール", "申請の整合性;>集計;集落別世帯数一覧;<;-;>選挙関連;選挙権の初期化;<;>エラーリスト作成;固定データの照合;<;支所番号の設定;-;終了時刻表示;-;転作データ作成;テストダイアログ;")
                'Case "申請の整合性" : sub申請の整合性()
                'Case "選挙権の初期化"
                '    St = SysAD.DB(sLRDB).DBProperty("選挙権クリア年月日")
                '    If IsDate(St) Then
                '        If MsgBox("選挙権の初期化を最後に行ったのは[" & St & "]ですが、初期化しますか", vbYesNo) = vbYes Then
                '            SysAD.DB(sLRDB).DBProperty("選挙権クリア年月日") = Date
                '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].選挙権の有無 = False, [D:個人Info].前年選挙権 = [選挙権の有無];")
                '        End If
                '    ElseIf MsgBox("選挙権を初期化します", vbYesNo) = vbNo Then
                '    Else
                '        SysAD.DB(sLRDB).DBProperty("選挙権クリア年月日") = Date
                '        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [D:個人Info].選挙権の有無 = False, [D:個人Info].前年選挙権 = [選挙権の有無];")
                '    End If

                'Case "３条受理通知書" : mvarPDW.PrintGo(ObjectMan.GetObject("受理通知書"), "3条")
                'Case "４条受理通知書" : mvarPDW.PrintGo(ObjectMan.GetObject("受理通知書"), "4条")
                'Case "５条受理通知書" : mvarPDW.PrintGo(ObjectMan.GetObject("受理通知書"), "5条")
                'Case "適格要件届出書" : mvarPDW.PrintGo(ObjectMan.GetObject("適格要件届出書"))

                'Case "３条届出書" : mvarPDW.PrintGo(ObjectMan.GetObject("３条届出書"))
                'Case "転用申請チェック表" : mvarPDW.PrintGo(ObjectMan.GetObject("転用申請チェック表"))
                'Case "営農計画書" : mvarPDW.PrintGo(ADApp.ObjectMan.GetObject("営農計画書"))
                'Case "奨励金交付申請書" : mvarPDW.PrintGo(ADApp.ObjectMan.GetObject("奨励金交付申請書"))
                'Case "奨励金総括表" : mvarPDW.PrintGo(ADApp.ObjectMan.GetObject("奨励金総括表"))
                'Case "集落別選挙人集計"
                '    St = "SELECT [D:個人Info].投票区, '行政区.' & [行政区ID] AS [Key], [V_行政区].ID AS 行政区ID, [M:投票区].選挙区, [V_行政区].行政区, Count([D:個人Info].ID) AS 人数 " & _
                '    "FROM ([M:投票区] INNER JOIN [D:個人Info] ON [M:投票区].ID = [D:個人Info].投票区) INNER JOIN [V_行政区] ON [D:個人Info].行政区ID = [V_行政区].ID " & _
                '    "GROUP BY [D:個人Info].投票区, '行政区.' & [行政区ID], [V_行政区].ID, [M:投票区].選挙区, [V_行政区].行政区, [D:個人Info].選挙権の有無 " & _
                '    "HAVING ((([D:個人Info].選挙権の有無)=True)) ORDER BY [D:個人Info].投票区, [V_行政区].ID;"
                '    mvarPDW.SQLListview.SQLListviewCHead(St, "投票区;選挙区;行政区;行政区;@人数;人数", "集落別選挙人数集計")
                'Case "選挙人リスト一覧"
                '    St = "SELECT '世帯員.' & [D:個人Info].ID AS [KEY], [D:個人Info].投票区, [M:投票区].選挙区 AS [投票区名], [V_行政区].ID AS 行政区ID, [V_行政区].行政区, [D:個人Info_1].フリガナ AS [ソート用], [V_続柄].ID AS 順序, [D:個人Info].フリガナ, [D:個人Info].氏名, [D:個人Info].住所, [V_続柄].続柄 AS 続柄A, [V_続柄_1].続柄 AS 続柄B, [D:個人Info].生年月日 " & _
                '    "FROM [D:個人Info] AS [D:個人Info_1] INNER JOIN (((([M:投票区] INNER JOIN [D:個人Info] ON [M:投票区].ID = [D:個人Info].投票区) INNER JOIN [V_続柄] ON [D:個人Info].続柄1 = [V_続柄].ID) INNER JOIN [V_続柄] AS [V_続柄_1] ON [D:個人Info].続柄2 = [V_続柄_1].ID) INNER JOIN [V_行政区] ON [D:個人Info].行政区ID = [V_行政区].ID) ON [D:個人Info_1].世帯ID = [D:個人Info].世帯ID " & _
                '    "WHERE ((([D:個人Info].選挙権の有無)=True) AND (([D:個人Info_1].続柄1)=1) AND (([D:個人Info_1].住民区分)=0)) " & _
                '    "ORDER BY [D:個人Info].投票区, [V_行政区].ID, [D:個人Info_1].フリガナ, [V_続柄].ID; "
                '    mvarPDW.SQLListview.SQLListviewCHead(St, "氏名;氏名;投票区名;投票区名;行政区;行政区;フリガナ;フリガナ;住所;住所;続柄1;続柄A;続柄2;続柄B;=生年月日;生年月日", "選挙人リスト一覧")
                'Case "証明願い"
                'Case "選挙連番設定"
                '    mvarPDW.WaitMessage = "面積クリア中"
                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:世帯Info] SET [D:世帯Info].総経営地 = 0") '面積クリア
                '    mvarPDW.WaitMessage = "面積当込み中"
                '    SysAD.DB(sLRDB).ExecuteSQL("DELETE wk_耕作世帯.世帯ID, wk_耕作世帯.耕作面積 FROM wk_耕作世帯")
                '    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO wk_耕作世帯 ( 世帯ID, 耕作面積 ) " & _
                '    "SELECT IIf([自小作別]=0,IIf([管理世帯ID]<>0,[管理世帯ID],[所有世帯ID]),[借受世帯ID]) AS 世帯ID, Sum(IIf(IsNull([田面積]),0,[田面積])+IIf(IsNull([畑面積]),0,[畑面積])+IIf(IsNull([樹園地]),0,[樹園地])) AS 耕作面積 " & _
                '    "FROM [D:農地Info] " & _
                '    "WHERE (((IIf([自小作別] = 0, IIf([管理世帯ID] <> 0, [管理世帯ID], [所有世帯ID]), [借受世帯ID])) Is Not Null) And (([D:農地Info].農地状況) < 20 Or ([D:農地Info].農地状況) Is Null)) " & _
                '    "GROUP BY IIf([自小作別]=0,IIf([管理世帯ID]<>0,[管理世帯ID],[所有世帯ID]),[借受世帯ID]);")

                '    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:世帯Info] INNER JOIN wk_耕作世帯 ON [D:世帯Info].ID = wk_耕作世帯.世帯ID SET [D:世帯Info].総経営地 = [wk_耕作世帯].[耕作面積]")
                '    mvarPDW.WaitMessage = ""
                'Case "標準申請入力" : SysAD.DB(sLRDB).DBProperty("申請書の入力画面") = 0
                'Case "拡張申請入力" : SysAD.DB(sLRDB).DBProperty("申請書の入力画面") = -1
                'Case "選挙CSVファイル参照" : DVProperty.Controls.Value("TX選挙CSVファイル") = Fnc.CanGetFileName(SysAD.MDIForm.SaveFileDlg("選挙CSVファイルの設定", "*.CSV"), DVProperty.Controls.Value("TX選挙CSVファイル"))
                'Case "FileExport"
                '    St = SysAD.DB(sLRDB).DBProperty("選挙CSVファイル")
                '    Rs = SysAD.DB(sLRDB).GetRecordsetEx(sParam, , , Me)
                '    Ts = Fnc.Fs.CreateTextFile(St, ForWriting, True)
                '    Ts.WriteS(Rs.GetString(",", vbCrLf))
                '    Ts.CloseFile()
                '    MsgBox("[" & St & "]の書込みが終了しました")
                '    SysAD.DB(sLRDB).CloseRs(Rs)
                'Case "支所番号の設定" : mod農家台帳.sub支所番号設定()
                'Case "基幹データの取込" : MsgBox("現バージョンでは対応できません")
                'Case "現地調査日程表", "現地調査担当委員" : mvarPDW.PrintGo(ADApp.ObjectMan.GetObject(sCmd), sCmd)
                'Case "現地調査管理"
                '    Dim p現地Frm As New frm現地調査日程表
                '    With Fnc.InputMulti("現地調査日程表", "入力をお願いします。", "今月総会番号;数値;" & Val(SysAD.DB(sLRDB).DBProperty("今月総会番号")), , True)
                '        If .IsALLEmpty Then
                '        Else
                '            SysAD.DB(sLRDB).DBProperty("今月総会番号") = Val(.Item("今月総会番号"))
                '            Select Case p現地Frm.SortData(Val(.Item("今月総会番号")))
                '                Case 1
                '                Case 2
                '            End Select
                '        End If
                '    End With
                'Case "現地調査報告書"
                '    St = Fnc.InputText("総会番号", "総会番号を入力してください", SysAD.DB(sLRDB).DBProperty("今月総会番号"), 0, vbIMEModeOff)
                '    If St <> "" Then
                '        mvarPDW.PrintGo(ObjectMan.GetObject("現地調査報告書"), St)
                '    End If
                'Case "農地-農用地内へ変更"
                '    If Len(sParam) Then
                '        Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
                '        If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を農用地内へ変更しますか", vbYesNo) Then
                '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 1 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));")
                '        End If
                '    End If
                'Case "農地-農振地内へ変更"
                '    If Len(sParam) Then
                '        Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
                '        If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を農振地へ変更しますか", vbYesNo) Then
                '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 0 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));")
                '        End If
                '    End If
                'Case "農地-農振除外"
                '    If Len(sParam) Then
                '        Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
                '        If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を農振除外しますか", vbYesNo) Then
                '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 2 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));")
                '        End If
                '    End If
                'Case "筆選択"
                'Case "農地-一括で転用農地へ（申請無し）"
                '    If Len(sParam) Then
                '        Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
                '        If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を転用農地にしますか", vbYesNo) Then
                '            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 2 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));")
                '            For n = 0 To UBound(Ar)
                '                If Len(Ar(n)) Then
                '                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 更新日, 内容 ) VALUES (" & Val(Ar(n)) & ",Date(),'職権による転用修正')")
                '                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_転用農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & Val(Ar(n)) & "));")
                '                    SysAD.DB(sLRDB).ExecuteSQL("DELETE [D:農地Info].ID FROM [D:農地Info] WHERE [D:農地Info].ID=" & Val(Ar(n)))
                '                End If
                '            Next
                '            mvarPDW.SQLListview.Refresh()
                '        End If
                '    End If
                'Case "転作データ作成"
                '    Dim sFName1 As String
                '    Dim sFName2 As String
                '    Dim sData As String
                '    Dim d個人 As New CDataDictionary

                '    Ar = ADApp.Common.MasterTable.ItemArray("大字")
                '    St = Replace(Join(Ar, ";"), ",", ":")

                '    St = Fnc.SelectMultiString(St, "大字を選択してください")
                '    If Len(St) Then
                '        sFName1 = Fnc.OpenFileDlg("転作用データ出力", "農地のファイル名を指定してください(*.csv)", CreateObject("WScript.Shell").SpecialFolders("DeskTop") & "\")
                '        sFName2 = Fnc.OpenFileDlg("転作用データ出力", "個人のファイル名を指定してください(*.csv)", CreateObject("WScript.Shell").SpecialFolders("DeskTop") & "\")

                '        Ar = Split(St, ";")
                '        For n = 0 To UBound(Ar)
                '            If Len(Ar(n)) Then
                '                Ar(n) = Left$(Ar(n), InStr(Ar(n), ":") - 1)
                '            End If
                '        Next
                '        St = Join(Ar, ",")
                '        If Right$(St, 1) = "," Then St = Left$(St, Len(St) - 1)

                '        If Len(sFName1) And sFName1 <> "No Select" Then
                '            Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT IIF([V_農地].管理者ID<>0,[V_農地].管理世帯ID,[V_農地].所有世帯ID) AS [所有世帯ID],[管理人ID] AS [所有者ID],IIF([自小作別]>0,[借受人ID],0) AS [借受者ID], [土地所在], 田面積 AS [面積] FROM [V_農地] WHERE [大字ID] IN (" & St & ") AND [田面積]>0 ORDER BY [大字ID],[小字ID],[地番]", 0, , Me)
                '            sData = Join(Rs.FildeHeaders, ",")
                '            If Not Rs.EOF Then
                '                sData = sData & vbCrLf & Rs.GetString(",")
                '            End If
                '            Rs.CloseRs()
                '            Ts = Fnc.Fs.OpenTextFile(sFName1, ForWriting, True)
                '            Ts.WriteS(sData)
                '            Ts.CloseFile()
                '        End If

                '        If Len(sFName2) And sFName2 <> "No Select" Then
                '            Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT [V_農地].管理人ID FROM [V_農地] WHERE [大字ID] IN (" & St & ") AND [田面積]>0 AND [管理人ID]<>0 GROUP BY [V_農地].管理人ID;", 0, , Me)
                '            Do Until Rs.EOF
                '                If Not d個人.Exists(Rs.Value("管理人ID")) Then
                '                    d個人.Add(Rs.Value("管理人ID"), Rs.Value("管理人ID"))
                '                End If
                '                Rs.MoveNext()
                '            Loop
                '            Rs.CloseRs()
                '            Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT [V_農地].借受人ID FROM [V_農地] WHERE [大字ID] IN (" & St & ") AND [自小作別]>0 AND [田面積]>0 AND [借受人ID]<>0 GROUP BY [V_農地].借受人ID;", 0, , Me)
                '            Do Until Rs.EOF
                '                If Not d個人.Exists(Rs.Value("借受人ID")) Then
                '                    d個人.Add(Rs.Value("借受人ID"), Rs.Value("借受人ID"))
                '                End If
                '                Rs.MoveNext()
                '            Loop
                '            Rs.CloseRs()

                '            Rs = SysAD.DB(sLRDB).GetRecordsetEx("SELECT [D:個人Info].行政区ID, [V_行政区].名称 AS [行政区], [D:個人Info].ID AS [個人ID], [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].生年月日 FROM [D:個人Info] LEFT JOIN [V_行政区] ON [D:個人Info].行政区ID = [V_行政区].ID ORDER BY [D:個人Info].行政区ID, [D:個人Info].[フリガナ], [D:個人Info].[ID];", 0, , Me)
                '            sData = Join(Rs.FildeHeaders, ",")
                '            Do Until Rs.EOF
                '                If d個人.Exists(Rs.Value("個人ID")) Then sData = sData & vbCrLf & Rs.GetLineString(GetArray("行政区ID", "行政区", "個人ID", "氏名", "住所", "生年月日"), ",")
                '                Rs.MoveNext()
                '            Loop
                '            Rs.CloseRs()
                '            Ts = Fnc.Fs.OpenTextFile(sFName2, ForWriting, True, TristateTrue)
                '            Ts.WriteLine(sData)
                '            Ts.CloseFile()
                '        End If
                '    End If
                'Case "選挙人名簿" : mvarPDW.PrintGo(New C選挙人名簿)
                'Case "関連者住所"
                '    CDataviewSK_DoCommand2 = fnc関連者一覧(sParam)
                'Case "ＳＱＬの実行"
                '    St = InputBox("SQL文を入力してください", "ＳＱＬの実行", "")
                '    If Len(St) Then
                '        SysAD.DB(sLRDB).Execute(St)
                '    End If
                'Case "非農地申請フォルダ"
                '    SysAD.SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [S_Folder]([KEY],ParentKey,表示順序,名称,Class,Icon,SelectIcon,Visible) VALUES('非農地受付中.0','受付中.0',10000,'非農地受付中','受付中','CloseFolderG','OpenFolderG',-1)")
                '    SysAD.SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [S_Folder]([KEY],ParentKey,表示順序,名称,Class,Icon,SelectIcon,Visible) VALUES('非農地審査中.0','審査中.0',10000,'非農地審査中','審査中','CloseFolderG','OpenFolderG',-1)")
                '    SysAD.SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [S_Folder]([KEY],ParentKey,表示順序,名称,Class,Icon,SelectIcon,Visible) VALUES('非農地承認済.0','許可済.0',10000,'非農地承認済','許可済','CloseFolderG','OpenFolderG',-1)")

                'Case Else
                '    If CheckBoolParams(sCmd, "小作料の表示;コードの表示;住民区分の表示;住民区分コードの表示;年齢の表示;選挙権の表示;農地確認の表示", "農地リストの小作料表示;世帯員リストのコード表示;世帯員リストの住民区分表示;世帯員リストの住民区分コード表示;世帯員リストの年齢表示;世帯員リストの選挙権表示;世帯リストの確認表示;農地リストの確認表示") Then
                '    ElseIf Len(SysAD.DB(sLRDB).DBProperty(sCmd)) Then
                '        Ar = Split(SysAD.DB(sLRDB).DBProperty(sCmd), ";")
                '        Select Case UBound(Ar)
                '            Case 0 : CDataviewSK_DoCommand2(CStr(Ar(0)))
                '            Case Else : CDataviewSK_DoCommand2(CStr(Ar(0)), CStr(Ar(1)))
                '        End Select
                '    Else
                '        Debug.Assert(CaseAssertPrint(sCmd))
                '    End If

        End Select
    End Sub
    Public Sub Set台帳Menu(ByVal pToolStrip As ToolStrip)
        'If SysAD.DB(sLRDB).DBPropertyEx("ヤミ小作を小作扱いする", vbBoolean, False) Then St = St & "^"
        'St = St & "ヤミ小作を小作扱いする"
        'If Not DVProperty.Body Is Nothing Then St = St & "-;閉じる;"

        mvarMenu = New ToolStripDropDownButton("農地基本台帳")
        With mvarMenu
            '世帯と農地の整合チェック
            AddHandler .DropDownItems.Add("同一地番一覧").Click, AddressOf ClickMenu
            Dim p農地一覧 As New ToolStripDropDownButton("農地一覧")
            mvarMenu.DropDownItems.Add(p農地一覧)
            With p農地一覧
                AddHandler .DropDownItems.Add("３年以上貸借").Click, AddressOf ClickMenu
                '５年以上貸借
                '死亡名義管理者不明農地一覧
                '同一地番一覧
                '不在地主名義管理者不明農地一覧
                '-
                '貸し手・借り手同一者エラー農地一覧
                '管理者＝所有者エラー農地一覧
            End With

            Dim p整合性 As New ToolStripDropDownButton("整合性")
            mvarMenu.DropDownItems.Add(p整合性)
            With p整合性
                '住記検索フリガナ初期化
                '住記の整合性
                '土地履歴の作成
                '-
                '小字テーブルの正規化
                '所有／管理世帯の整合性
                '借受世帯の整合性
                '転用履歴の整合性
                '-
                '世帯自動生成
                ';-
                ';同一地番の整合性;農地面積の整合性;土地履歴と申請の不一致
                '-
                ';データベースの更新;フォルダの整理;住記情報の転送
                ';-
                ';貸し手・借り手同一者解約
                ';管理者＝所有者エラー農地解除
                ';現況地目で面積設定
            End With



            'プロパティ;
        End With
        pToolStrip.Items.Add(mvarMenu)
    End Sub

    Public Overrides Sub ClickMenu(s As Object, e As System.EventArgs)
        Me.Command(s.text)
    End Sub

    Public Sub sub同一地番の整合性()
        Dim 大字ID As Long
        Dim s地番 As String
        Dim St As String
        Dim n As Long
        Dim n登記面積 As Decimal
        Dim n現況合計 As Decimal

        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(
            "SELECT [D:農地Info].ID, [D:農地Info].大字ID, [D:農地Info].地番, [D:農地Info].登記簿面積, [D:農地Info].所有者ID, [D:農地Info].実面積, [D:農地Info].一部現況, [D:農地Info].所在, [D:農地Info].自小作別 FROM [D:農地Info] WHERE ((([D:農地Info].大字ID) In (SELECT [大字ID] FROM [D:農地Info] As Tmp GROUP BY [大字ID],[地番] HAVING Count(*)>1  And [地番] = [D:農地Info].[地番])) AND (([D:農地Info].一部現況)=0 Or ([D:農地Info].一部現況) Is Null)) ORDER BY [D:農地Info].大字ID, [D:農地Info].地番, [D:農地Info].実面積 DESC;")

        大字ID = 0
        s地番 = ""
        n登記面積 = 0
        n現況合計 = 0

        St = ""
        For Each pRow As DataRow In pTBL.Rows
            If 大字ID = 0 And s地番 = "" Then
                大字ID = pRow.Item("大字ID")
                s地番 = pRow.Item("地番")
                n登記面積 = pRow.Item("登記簿面積")
                n現況合計 = pRow.Item("実面積")
            ElseIf 大字ID <> pRow.Item("大字ID") Or s地番 <> pRow.Item("地番") Or n登記面積 <> pRow.Item("登記簿面積") Then
                If n登記面積 * 10 >= n現況合計 Then
                    St = St & vbBack & 大字ID & ";" & s地番 & ";" & n登記面積
                End If

                大字ID = pRow.Item("大字ID")
                s地番 = pRow.Item("地番")
                n登記面積 = pRow.Item("登記簿面積")
                n現況合計 = pRow.Item("実面積")
            Else
                n現況合計 = n現況合計 + pRow.Item("実面積")
            End If

        Next
        If n登記面積 * 10 >= n現況合計 Then
            St = St & vbBack & 大字ID & ";" & s地番 & ";" & n登記面積
        End If

        St = St & vbBack
        n = 1
        大字ID = 0
        s地番 = ""
        For Each pRow As DataRow In pTBL.Rows
            If InStr(St, vbBack & pRow.Item("大字ID") & ";" & pRow.Item("地番") & ";" & pRow.Item("登記簿面積") & vbBack) Then
                If 大字ID = 0 And s地番 = "" Then
                    大字ID = pRow.Item("大字ID")
                    s地番 = pRow.Item("地番")
                    n登記面積 = pRow.Item("登記簿面積")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一部現況]=" & n & " WHERE [ID]=" & pRow.Item("ID"))
                    n = n + 1
                ElseIf 大字ID <> pRow.Item("大字ID") Or s地番 <> pRow.Item("地番") Or n登記面積 <> pRow.Item("登記簿面積") Then
                    n = 1
                    大字ID = pRow.Item("大字ID")
                    s地番 = pRow.Item("地番")
                    n登記面積 = pRow.Item("登記簿面積")
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一部現況]=" & n & " WHERE [ID]=" & pRow.Item("ID"))
                    n = n + 1
                Else
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一部現況]=" & n & " WHERE [ID]=" & pRow.Item("ID"))
                    n = n + 1
                End If
            End If
        Next
    End Sub


    Protected mvar更新履歴 As DataTable
    Dim ATime As DateTime = Now
    Dim nPt As Integer = 0

    Private Sub CheckTime(sErr As String)
        Dim s As TimeSpan = Now - ATime
        Debug.Print(sErr & ")" & Right("0000" & nPt, 5) & "=>" & s.Seconds.ToString & ":" & s.Milliseconds.ToString) : nPt += 1
    End Sub

    Private Sub SaveDSetSchema()
        If Not SysAD.IsClickOnceDeployed Then
            Dim pPath As New IO.DirectoryInfo(My.Application.Info.DirectoryPath)
            Do Until pPath.Name = "農地基本台帳"
                pPath = pPath.Parent
            Loop

            For Each pTarget As IO.DirectoryInfo In pPath.GetDirectories
                If pTarget.Name = "Resources" Then
                    Try
                        DSet.WriteXmlSchema(pTarget.FullName & "\DSet.xsd")
                        Exit For
                    Catch ex As Exception

                    End Try
                End If
            Next
        End If
    End Sub

    Public Sub InitSystem()
        With SysAD.DB(sLRDB)
            Dim sError As String = "000"

            Try
                ATime = Now
                sError = "000"
                CheckTime(sError)
                mvar更新履歴 = .AddPrimaryKey(.GetTableBySqlSelect("SELECT * FROM [S_システム更新履歴]"), "Key")
                mvar更新履歴.TableName = "S_システム更新履歴"
                CType(DSet.Tables("S_システム更新履歴"), HimTools2012.Data.DataTableEx).MergePlus(mvar更新履歴, False, MissingSchemaAction.AddWithKey)

                'If mvar更新履歴.Rows.Find("土地系図20150226") Is Nothing Then
                '    Dim sRet As String = .ExecuteSQL("CREATE TABLE [D_土地系図]([自ID] DECIMAL,[元ID] DECIMAL,[元土地所在] VARCHAR(255),PRIMARY KEY ([自ID],[元ID]))")
                '    If sRet = "" OrElse sRet = "OK" Then
                '        .ExecuteSQL("INSERT INTO S_システム更新履歴([Key],[Update]) VALUES('{0}',Now)", "土地系図20150226")
                '    End If
                'End If
                'If mvar更新履歴.Rows.Find("非農地通知判定20150611") Is Nothing Then
                '    Dim sRet As String = .ExecuteSQL("CREATE TABLE D_非農地通知判定([ID] LONG CONSTRAINT pkey PRIMARY KEY,[NID] LONG,[一筆コード] LONG,[発行番号] VARCHAR(64),[通知番号] LONG,[大字] VARCHAR(255),[小字] VARCHAR(255),[調査時地番] VARCHAR(255),[調査時登記地目] VARCHAR(255),[調査時現況地目] VARCHAR(255),[調査時面積] Decimal(18,3),[所有者ID] Decimal(18,0),[所有者氏名] VARCHAR(255),[所有者住所] VARCHAR(255),[所有者郵便番号] VARCHAR(255),[所有者住民区分] LONG,[納税義務者ID] Decimal(18,0),[納税義務者氏名] VARCHAR(255),[納税義務者住所] VARCHAR(255),[納税義務者郵便番号] VARCHAR(255),[納税義務者住民区分] LONG,[送付先ID] Decimal(18,0),[送付先氏名] VARCHAR(255),[送付先住所] VARCHAR(255),[送付先郵便番号] VARCHAR(255),[発行年月日] DATETIME)")

                '    If sRet = "" OrElse sRet = "OK" Then
                '        .ExecuteSQL("INSERT INTO S_システム更新履歴([Key],[Update]) VALUES('{0}',Now)", "非農地通知判定20150611")
                '    End If
                'End If

                'If mvar更新履歴.Rows.Find("D_公開用個人追加") Is Nothing Then
                '    SysAD.DB(sLRDB).ExecuteSQL("CREATE TABLE [D_公開用個人] ([AutoID] COUNTER NOT NULL CONSTRAINT pKey PRIMARY KEY,[PID] Long,[氏名] VARCHAR(255))")
                '    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('D_公開用個人追加',Now)")
                'End If

                'マスタ登録
                sError = "001"
                CheckTime(sError)
                DSet.Tables("M_BASICALL").Merge(.AddPrimaryKey(.GetTableBySqlSelect("SELECT * FROM M_BASICALL"), "ID", "Class"))

                CheckTime(sError)
                sError = "002"

                Dim TBLXML As New HimTools2012.Data.DataTableEx()
                TBLXML.LoadText(My.Resources.Resource1.M_BASICALL)
                DSet.Tables("M_BASICALL").Merge(TBLXML, False, MissingSchemaAction.AddWithKey)
                mvarDataMaster = New HimTools2012.System管理.CBASICALL(DSet.Tables("M_BASICALL"), sLRDB)

                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_土地履歴] WHERE [ID]=0")

                CheckTBL(pTBL, "数値結果", "LONG")
                CheckTBL(pTBL, "受付年月日", "DATETIME")
                CheckTBL(pTBL, "許可年月日", "DATETIME")
                CheckTBL(pTBL, "関連情報", "TEXT")
                CheckTBL(pTBL, "結果区分", "VARCHAR(255)")



                sError = "003"
                CheckTime(sError)


                sError = "009"
                CheckTime(sError)
                With mvarDataMaster
                    .InitMaster("性別", "V_性別", DSet)
                    .InitMaster("続柄", "V_続柄", DSet)
                    .InitMaster("行政区", "V_行政区", DSet)
                    .InitMaster("住民区分", "V_住民区分", DSet)
                    .InitMaster("大字", "V_大字", DSet)
                    .InitMaster("小字", "V_小字", DSet)
                    .InitMaster("地目", "V_地目", DSet)
                    .InitMaster("課税地目", "V_現況地目", DSet)
                    .InitMaster("農委地目", "V_農委地目", DSet)
                    .InitMaster("適用法令", "V_適用法令", DSet)
                    .InitMaster("農地区分", "申請時農地区分", DSet)
                    .InitMaster("農業委員", "農業委員", DSet)
                    .InitMaster("申請時農振区分", "申請時農振区分", DSet)
                    .InitMaster("都市計画区分", "申請時都計区分", DSet)
                    .InitMaster("小作形態", "V_小作形態", DSet)
                    .InitMaster("土地異動事由", "土地異動事由", DSet)
                    .InitMaster("農地状況", "V_農地状況", DSet)
                End With

                sError = "015"
                CheckTime(sError)
                mvarTBL個人 = New CTBL個人(DSet, DSet.Tables("D:個人Info"))

                sError = "016"
                CheckTime(sError)
                mvarTBL世帯 = New CTBL世帯(DSet, DSet.Tables("D:世帯Info"))



                sError = "024"
                CheckTime(sError)
                mvarTBL農地 = New CTBL農地(DSet, DSet.Tables("D:農地Info"))


                'If mvar更新履歴.Rows.Find("農地Info20181204") Is Nothing Then
                '    With SysAD.DB(sLRDB)
                '        .ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('農地Info20181204',Now)")
                '        .ExecuteSQL("ALTER TABLE [D:農地Info] ADD [農地所有内訳] LONG")
                '        .ExecuteSQL("ALTER TABLE [D_転用農地] ADD [農地所有内訳] LONG")
                '        .ExecuteSQL("ALTER TABLE [D_削除農地] ADD [農地所有内訳] LONG")
                '    End With
                'End If

                'If mvar更新履歴.Rows.Find("農家Info20181211") Is Nothing Then
                '    With SysAD.DB(sLRDB)
                '        .ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('農家Info20181211',Now)")
                '        .ExecuteSQL("ALTER TABLE [D:世帯Info] ADD [その他家畜内訳] TEXT")
                '        .ExecuteSQL("ALTER TABLE [D_削除世帯] ADD [その他家畜内訳] TEXT")
                '        .ExecuteSQL("ALTER TABLE [D:世帯Info] ADD [その他機具内訳] TEXT")
                '        .ExecuteSQL("ALTER TABLE [D_削除世帯] ADD [その他機具内訳] TEXT")
                '        .ExecuteSQL("ALTER TABLE [D:世帯Info] ADD [その他施設内訳] TEXT")
                '        .ExecuteSQL("ALTER TABLE [D_削除世帯] ADD [その他施設内訳] TEXT")
                '    End With
                'End If

                If mvar更新履歴.Rows.Find("農地Info20190710") Is Nothing Then
                    With SysAD.DB(sLRDB)
                        .ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('農地Info20190710',Now)")
                        If Not mvarTBL農地.Columns.Contains("本番") Then
                            .ExecuteSQL("ALTER TABLE [D:農地Info] ADD [本番] LONG")
                            .ExecuteSQL("ALTER TABLE [D_転用農地] ADD [本番] LONG")
                            .ExecuteSQL("ALTER TABLE [D_削除農地] ADD [本番] LONG")
                        End If
                        If Not mvarTBL農地.Columns.Contains("枝番]") Then
                            .ExecuteSQL("ALTER TABLE [D:農地Info] ADD [枝番] LONG")
                            .ExecuteSQL("ALTER TABLE [D_転用農地] ADD [枝番] LONG")
                            .ExecuteSQL("ALTER TABLE [D_削除農地] ADD [枝番] LONG")
                        End If
                    End With
                End If

                If mvar更新履歴.Rows.Find("農地Info20190724") Is Nothing Then
                    With SysAD.DB(sLRDB)
                        .ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('農地Info20190724',Now)")
                        If Not mvarTBL農地.Columns.Contains("区分地上権") Then
                            .ExecuteSQL("ALTER TABLE [D_申請] ADD [区分地上権] LONG")
                            .ExecuteSQL("ALTER TABLE [D_申請] ADD [区分地上権内容] MEMO")
                        End If
                    End With
                End If

                If mvar更新履歴.Rows.Find("農地Info20191030") Is Nothing Then
                    With SysAD.DB(sLRDB)
                        .ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('農地Info20191030',Now)")
                        If Not mvarTBL農地.Columns.Contains("人農地プラン区分") Then
                            .ExecuteSQL("ALTER TABLE [D:農地Info] ADD [人農地プラン区分] LONG")
                            .ExecuteSQL("ALTER TABLE [D_転用農地] ADD [人農地プラン区分] LONG")
                            .ExecuteSQL("ALTER TABLE [D_削除農地] ADD [人農地プラン区分] LONG")
                        End If
                        If Not mvarTBL農地.Columns.Contains("人農地プラン貸付内訳]") Then
                            .ExecuteSQL("ALTER TABLE [D:農地Info] ADD [人農地プラン貸付内訳] LONG")
                            .ExecuteSQL("ALTER TABLE [D_転用農地] ADD [人農地プラン貸付内訳] LONG")
                            .ExecuteSQL("ALTER TABLE [D_削除農地] ADD [人農地プラン貸付内訳] LONG")
                        End If
                    End With

                    Dim getTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_BASICALL] WHERE [Class]='人農地プラン区分'")
                    If getTBL.Rows.Count = 0 Then
                        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(0,'人農地プラン区分','-');")
                        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(1,'人農地プラン区分','売渡');")
                        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(2,'人農地プラン区分','貸付');")

                        getTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_BASICALL] WHERE [Class]='人農地プラン貸付内訳'")
                        If getTBL.Rows.Count = 0 Then
                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(0,'人農地プラン貸付内訳','-');")
                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(1,'人農地プラン貸付内訳','機構活用あり');")
                            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_BASICALL]([ID],Class,名称) VALUES(2,'人農地プラン貸付内訳','機構活用なし');")

                            MsgBox("マスターデータの更新を行いました。再起動してください。")
                            End
                        End If


                    End If
                End If

                If mvar更新履歴.Rows.Find("個人Info20191204") Is Nothing Then
                    With SysAD.DB(sLRDB)
                        .ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('個人Info20191204',Now)")
                        If Not mvarTBL個人.Columns.Contains("集積協力金の有無") Then
                            .ExecuteSQL("ALTER TABLE [D:個人Info] ADD [集積協力金の有無] LONG")
                            .ExecuteSQL("ALTER TABLE [D_削除個人] ADD [集積協力金の有無] LONG")
                        End If
                        If Not mvarTBL個人.Columns.Contains("転換協力金の有無") Then
                            .ExecuteSQL("ALTER TABLE [D:個人Info] ADD [転換協力金の有無] LONG")
                            .ExecuteSQL("ALTER TABLE [D_削除個人] ADD [転換協力金の有無] LONG")
                        End If
                        If Not mvarTBL個人.Columns.Contains("集積協力金開始時期") Then
                            .ExecuteSQL("ALTER TABLE [D:個人Info] ADD [集積協力金開始時期] VARCHAR(255)")
                            .ExecuteSQL("ALTER TABLE [D_削除個人] ADD [集積協力金開始時期] VARCHAR(255)")
                        End If
                        If Not mvarTBL個人.Columns.Contains("転換協力金開始時期") Then
                            .ExecuteSQL("ALTER TABLE [D:個人Info] ADD [転換協力金開始時期] VARCHAR(255)")
                            .ExecuteSQL("ALTER TABLE [D_削除個人] ADD [転換協力金開始時期] VARCHAR(255)")
                        End If
                    End With
                End If

                If mvar更新履歴.Rows.Find("個人Info20191205") Is Nothing Then
                    With SysAD.DB(sLRDB)
                        .ExecuteSQL("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('個人Info20191205',Now)")
                        .ExecuteSQL("INSERT INTO S_システムデータ ( [KEY], DATA ) VALUES ( '協力金開始時期', 'H26年Ⅲ期;H27年Ⅰ期;H27年Ⅱ期;H27年Ⅲ期;H27年Ⅳ期;H28年Ⅰ期;H28年Ⅱ期;H28年Ⅲ期;H28年Ⅳ期;H29年Ⅰ期;H29年Ⅱ期;H29年Ⅲ期;H29年Ⅴ期;H29年Ⅶ期;H30年Ⅰ期;H30年Ⅱ期;H30年Ⅲ期;H30年Ⅳ期;H30年Ⅴ期;H31年Ⅰ期;H31年Ⅱ期;H31年Ⅲ期;H31年Ⅳ期;H31Ⅴ期;R2年Ⅰ期;R2年Ⅱ期;R2年Ⅲ期;R2年Ⅳ期;R2年Ⅴ期;R3年Ⅰ期;R3年Ⅱ期;R3年Ⅲ期;R3年Ⅳ期;R3年Ⅴ期');")
                    End With
                End If

                mvarTBL土地履歴 = New CTBL土地履歴(DSet, DSet.Tables("D_土地履歴"))
                'SaveDSetSchema()
                sError = "029"
                CheckTime(sError)
                mvarTBL世帯営農 = New CTBL世帯営農(DSet, DSet.Tables("D_世帯営農"))

                sError = "030"
                CheckTime(sError)
                mvarTBL申請 = New CTBL申請(DSet, DSet.Tables("D_申請"))

                sError = "031"
                CheckTime(sError)

            Catch ex As Exception
                MsgBox(sError & ":" & ex.Message)
                End
            End Try

        End With
    End Sub

    Public Sub Set大字地番InList(ByRef p申請 As HimTools2012.Data.UpdateRow, ByVal sList As String)
        If SysAD.市町村.市町村名 = "日置市" Then Exit Sub
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].大字ID,[D:農地Info].小字ID, [D:農地Info].地番, [D:農地Info].ID FROM [D:農地Info] WHERE ((([D:農地Info].ID) In (" & Replace(Replace(sList, "農地.", ""), ";", ",") & "))) ORDER BY [D:農地Info].大字ID, [D:農地Info].地番;")
        If pTBL.Rows.Count = 0 Then
            p申請.SetValue("申請地大字ID", DBNull.Value)
            p申請.SetValue("申請地地番", DBNull.Value)
        Else
            With pTBL.Rows(0)
                If IsDBNull(.Item("小字ID")) Then
                    p申請.SetValue("申請地大字ID", .Item("大字ID") * 1000)
                ElseIf .Item("小字ID") = 0 Then
                    p申請.SetValue("申請地大字ID", .Item("大字ID") * 1000)
                Else
                    p申請.SetValue("申請地大字ID", .Item("小字ID"))
                End If

            End With
        End If
    End Sub
    Private Sub CheckTBL(ByRef pTBL As DataTable, sName As String, sType As String)
        If Not pTBL.Columns.Contains(sName) Then
            SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [D_土地履歴] ADD [{0}] {1};", sName, sType)
        End If

    End Sub

    Public Sub subOBJ削除(ByVal sTitle As String, ByVal sMessage As String, ByVal sSQL As String, ByVal pObj As HimTools2012.TargetSystem.CTargetObjWithView)
        If MsgBox(sMessage, vbYesNo, sTitle) = vbYes Then
            pObj.DoCommand("閉じる")
            SysAD.DB(sLRDB).ExecuteSQL(sSQL)
        End If
    End Sub


    '    Public Sub NCSV()
    '        Dim sSQL As String
    '        Dim St As String
    '        Dim sFile As String
    '        Dim Rs As RecordsetEx

    '        'On Error GoTo Ext:
    '        sFile = SysAD.MDIForm.SaveFileDlg("農地出力CSV", "*.CSV|*.CSV", SysAD.Functions.Fs.GetSpecialFolderSH("DeskTop"))
    '        If Not sFile = "No Select" Then
    '            sSQL = "SELECT V_農地.土地所在, V_地目.名称 AS 登記地目, V_現況地目.名称 AS 現況地目,V_農地.登記簿面積,V_農地.実面積,IIF([V_農地].[遊休化],'×','') AS 遊休化,IIF([V_農地].[無断転用],'×','') AS [無断転用],Choose(IIF(IsNull([農業振興地域]),0,[農業振興地域]+2),'-','農用地外','農用地内','農振地外') AS [農振区分],Choose(IIF(IsNull([都市計画法]),0,[都市計画法]+2),'-','都計外','都計内','用途地域内','調整区域内','市街化区域内','都市計画白地') AS [都市計画法区分],IIf([V_農地].[自小作別]=2,'移譲年金',IIf([V_農地].[自小作別]=1,'小作','-')) AS 借受状態, [D:個人Info].氏名 AS 管理者,[所有].ID AS [所有者コード],[所有].世帯ID AS [所有者世帯コード],所有.氏名 AS [所有者],所有.住所 AS [所有者住所], IIf([V_農地].[自小作別]>0,[T_借受者].[氏名],'') AS 借受者,IIf([V_農地].[自小作別]>0,[T_借受者].[ID],'') AS 借受者コード,IIf([V_農地].[自小作別]>0,[T_借受者].[世帯ID],'') AS 借受者世帯コード,IIf([V_農地].[自小作別]>0,[T_借受者].[住所],'') AS [借受者住所]," & _
    '            " IIf([V_農地].[自小作別]>0,[V_農地].[小作開始年月日]) AS 貸借開始日, IIf([V_農地].[自小作別]>0,[V_農地].[小作終了年月日]) AS 貸借終了日,[小作料S], IIf([V_農地].[共有持分分母]>0 AND [V_農地].[共有持分分子]>0,[V_農地].[共有持分分子] & '/' & [V_農地].[共有持分分母]) AS [持分]" & _
    '            ",[V_農地].[確認日時],[地目名] AS [農業委員会認定地目],[田面積],[畑面積],[樹園地],[ソート文字] FROM (((((V_農地 LEFT JOIN [D:個人Info] ON V_農地.管理人ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [T_借受者] ON V_農地.借受人ID = [T_借受者].ID) LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON V_農地.現況地目 = V_現況地目.ID) LEFT JOIN [V_農地状況] ON [V_農地].農地状況 = [V_農地状況].[ID]) LEFT JOIN [D:個人Info] AS 所有 ON V_農地.所有者ID = 所有.ID ORDER BY [V_農地].大字ID,[ソート文字] & IIf([一部現況]>0,'(' & [一部現況] & ')',0);"

    '            SysAD.MDIForm.Message = "農地データ作成中"
    '            Rs = SysAD.SysAD.DB(sLRDB).GetRecordsetEx(sSQL, 0)
    '            St = Join(Rs.FildeHeaders, ",") & vbCrLf & Rs.GetString(",", vbCrLf)
    '            SysAD.MDIForm.Message = "農地データ書き込み中"

    '            With SysAD.Func.FileSystem
    '                Dim Ts As NK97.CTextStream
    '                Ts = .CreateTextFile(sFile, True, True)
    '                Ts.WriteS(St)
    '                Ts.CloseFile()
    '            End With

    '            '            SysAD.Func.Fs.SaveTextFile sFile, St, True
    '            Rs.CloseRs()
    '            MsgBox("[" & sFile & "]に保存しました。")
    '            SysAD.MDIForm.Message = ""
    '        End If
    '        Exit Sub
    'Ext:
    '        Select Case Err.Number
    '            Case 70
    '                MsgBox("操作が不正です。既に開いているファイルに書き込みできません。", vbCritical)
    '            Case Else
    '                MsgBox("エラー:" & Err.Description, vbCritical)
    '        End Select
    '    End Sub

    Public Sub CreateFarmer(Optional ByVal nID As Long = -1, Optional ByVal nFamID As Long = 0, Optional ByVal nTownID As Long = -1, Optional ByVal n住民区分 As Long = 0, Optional ByVal sTitle As String = "住民の追加")
        Dim pClg As New 住民追加条件
        pClg.住民番号 = -1
        Dim pDlg As New dlgInputMulti(pClg, "住民追加", "追加する住民・法人の情報を入力してください")

        If pDlg.ShowDialog = DialogResult.OK AndAlso pClg.CheckValues Then
            If pClg.住民番号 = -1 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT MIn([ID])-1 as 最小値 FROM [D:個人Info]")
                pClg.住民番号 = pTBL.Rows(0).Item("最小値")
            End If

            Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(pClg.住民番号)
            If pRow IsNot Nothing Then
                MsgBox("既に存在します")
            Else
                nID = pClg.住民番号
                Dim sName As String = pClg.氏名
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:個人INFO]([ID],[氏名],[市町村ID],[世帯ID]) VALUES(" & nID & ",'" & sName & "'," & nTownID & "," & nFamID & ");")
                If Len(pClg.フリガナ) Then
                    SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人INFO] SET [フリガナ]='" & pClg.フリガナ & "' WHERE [ID]=" & nID)
                End If
                pRow = App農地基本台帳.TBL個人.FindRowByID(pClg.住民番号)

            End If
            CType(ObjectMan.GetObject("個人." & nID), HimTools2012.TargetSystem.CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
        End If

    End Sub


    Public Overrides Function GetProperty(sParam As String) As Object
        Return Nothing
    End Function



    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Return Nothing
    End Function


    Public Overrides ReadOnly Property アプリケーション名 As String
        Get
            Return "農地台帳"
        End Get
    End Property
End Class

'
'    Private Function CDataviewSK_PopupMenu(Optional ByVal sParam As String = "", Optional vParam As Variant) As String
'        Dim St As String
'        Dim st市町村種別 As String
'        Dim n As Long
'        
'        st市町村種別 = SystemDB.DBProperty("市町村種別")
'        

'        St = St & ";<;世帯と農地の整合チェック;-;" & _
'        ">農地一覧;３年以上貸借;５年以上貸借;死亡名義管理者不明農地一覧;同一地番一覧;不在地主名義管理者不明農地一覧;-;貸し手・借り手同一者エラー農地一覧;管理者＝所有者エラー農地一覧;<;" & _
'        ">整合性;住記検索フリガナ初期化;住記の整合性;土地履歴の作成;-;小字テーブルの正規化;所有／管理世帯の整合性;借受世帯の整合性;転用履歴の整合性;-;世帯自動生成;-;同一地番の整合性;農地面積の整合性;土地履歴と申請の不一致;-;データベースの更新;フォルダの整理;住記情報の転送;-;貸し手・借り手同一者解約;管理者＝所有者エラー農地解除;現況地目で面積設定;<;-;プロパティ;-"
'        
'        If SystemDB.DBPropertyEx("ヤミ小作を小作扱いする", vbBoolean, False) Then St = St & "^"
'        St = St & "ヤミ小作を小作扱いする"
'        If Not DVProperty.Body Is Nothing Then St = St & "-;閉じる;"
'        
'        St = SysAD.PopupMenu(St)
'        
'        Select Case St
'            Case "ヤミ小作を小作扱いする": SystemDB.DBProperty("ヤミ小作を小作扱いする") = Not SystemDB.DBPropertyEx("ヤミ小作を小作扱いする", vbBoolean, False)
'            Case "現況地目で面積設定"
'                Select Case SystemDB.DBProperty("集計は現況面積を優先")
'                    Case "-1"
'                        SystemDB.ExecuteSQL "UPDATE V_農地 SET V_農地.田面積 = IIf([実面積]>0,[実面積],[登記簿面積]), V_農地.畑面積 = 0, V_農地.樹園地 = 0, V_農地.採草放牧面積 = 0 WHERE [V_農地].[現況地目] In (" & SystemDB.DBProperty("田地目") & ");"
'                        SystemDB.ExecuteSQL "UPDATE V_農地 SET V_農地.田面積 = 0, V_農地.畑面積 = IIf([実面積]>0,[実面積],[登記簿面積]), V_農地.樹園地 = 0, V_農地.採草放牧面積 = 0 WHERE [V_農地].[現況地目] In (" & SystemDB.DBProperty("畑地目") & ");"
'                    Case Else
'                        SystemDB.ExecuteSQL "UPDATE V_農地 SET V_農地.田面積 = IIf([登記簿面積]>0,[登記簿面積],[実面積]), V_農地.畑面積 = 0, V_農地.樹園地 = 0, V_農地.採草放牧面積 = 0 WHERE [V_農地].[現況地目] In (" & SystemDB.DBProperty("田地目") & ");"
'                        SystemDB.ExecuteSQL "UPDATE V_農地 SET V_農地.田面積 = 0, V_農地.畑面積 = IIf([登記簿面積]>0,[登記簿面積],[実面積]), V_農地.樹園地 = 0, V_農地.採草放牧面積 = 0 WHERE [V_農地].[現況地目] In (" & SystemDB.DBProperty("畑地目") & ");"
'                End Select
'            Case "転用履歴の整合性"
'                Dim Rs As ADBasicLIB2.RecordsetEx
'                Set Rs = SystemDB.GetRecordsetEx("SELECT D_申請.ID,D_申請.農地リスト FROM D_申請 WHERE (((D_申請.法令)=40) AND ((D_申請.状態)=2));", 0, , Me)
'                Do Until Rs.EOF
'                    St = Replace(Replace(Rs.Value("農地リスト"), "転用農地.", ""), ";", ",")
'                    SystemDB.ExecuteSQL "INSERT INTO D_土地履歴 ( 更新日,LID, 異動事由, 内容, 申請ID ) SELECT Date() AS 式1, D_転用農地.ID, 10040 AS 式2, '4条転用許可' AS 式3, " & Rs.ID & " AS 式1 FROM D_転用農地 LEFT JOIN D_土地履歴 ON D_転用農地.ID = D_土地履歴.ID WHERE (((D_転用農地.ID) In (" & St & ")));"
'                    Rs.MoveNext
'                Loop
'                Rs.CloseRs
'                
'                Set Rs = SystemDB.GetRecordsetEx("SELECT D_申請.ID,D_申請.農地リスト FROM D_申請 WHERE (((D_申請.法令)=50 Or (D_申請.法令)=51) AND ((D_申請.状態)=2));", 0, , Me)
'                Do Until Rs.EOF
'                    St = Replace(Replace(Rs.Value("農地リスト"), "転用農地.", ""), ";", ",")
'                    SystemDB.ExecuteSQL "INSERT INTO D_土地履歴 ( 更新日,LID, 異動事由, 内容, 申請ID ) SELECT Date() AS 式1, D_転用農地.ID, 10050 AS 式2, '5条転用許可' AS 式3, " & Rs.ID & " AS 式1 FROM D_転用農地 LEFT JOIN D_土地履歴 ON D_転用農地.ID = D_土地履歴.ID WHERE (((D_転用農地.ID) In (" & St & ")));"
'                    
'                    Rs.MoveNext
'                Loop
'                Rs.CloseRs
'                
'                SystemDB.ExecuteSQL "UPDATE (D_土地履歴 AS D_土地履歴_1 INNER JOIN D_申請 ON D_土地履歴_1.申請ID = D_申請.ID) INNER JOIN D_転用農地 ON D_土地履歴_1.LID = D_転用農地.ID SET D_土地履歴_1.更新日 = [D_申請].[公告年月日], D_土地履歴_1.異動日 = [許可年月日], D_土地履歴_1.関係者A = [申請者A], D_土地履歴_1.関係者B = [申請者B] WHERE (((D_土地履歴_1.異動事由)=10050 Or (D_土地履歴_1.異動事由)=10040));"
'                SysAD.CopyAbleMessage "終了"
'            Case "世帯数調べ":
'                SystemDB.ExecuteSQL "UPDATE [D:世帯Info] SET [D:世帯Info].農地との関連 = False WHERE ((([D:世帯Info].農地との関連)=True));"
'                SystemDB.ExecuteSQL "UPDATE [D:農地Info] INNER JOIN ([D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].ID = [D:個人Info].世帯ID) ON [D:農地Info].所有者ID = [D:個人Info].ID SET [D:世帯Info].農地との関連 = True WHERE ((([D:世帯Info].農地との関連)=False));"
'                SystemDB.ExecuteSQL "UPDATE [D:農地Info] INNER JOIN ([D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].ID = [D:個人Info].世帯ID) ON [D:農地Info].借受人ID = [D:個人Info].ID SET [D:世帯Info].農地との関連 = True WHERE ((([D:世帯Info].農地との関連)=False));"
'                St = SystemDB.DBProperty("市町村ID")
'                If St = "" Then St = "-999"
'                mvarPDW.SQLListview.SQLListviewCHead "SELECT 'X.' & IIf([市町村ID]=" & St & ",'在住','不在') AS [KEY],IIf([市町村ID]=" & St & ",'在住','不在') AS 地主の所在, Count([D:世帯Info].ID) AS 世帯数 FROM [D:世帯Info] WHERE ((([D:世帯Info].農地との関連) = True)) GROUP BY 'X.' & IIf([市町村ID]=" & St & ",'在住','不在'),IIf([市町村ID]=" & St & ",'在住','不在');", "地主の所在;地主の所在;@世帯数;世帯数", "世帯数調べ"
'            Case "世帯と農地の整合チェック": sub農家世帯設定
'            Case "貸し手・借り手同一者エラー農地一覧": view農地List "[所有者ID]=[借受人ID] AND [自小作別]>0", "貸し手・借り手同一者エラー農地一覧" '/** 農地0001-0001-0002
'            Case "管理者＝所有者エラー農地一覧": view農地List "[管理世帯ID]=[所有世帯ID] AND [所有者ID]=[管理者ID]", "管理者＝所有者エラー農地一覧" '/** 農地0001-0001-0002
'            Case "管理者＝所有者エラー農地解除"
'                If MsgBox("管理者＝所有者である不正情報を解除しますか", vbYesNo) = vbYes Then SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].管理世帯ID = 0,[D:農地Info].管理者ID = 0 WHERE ([管理世帯ID]=[所有世帯ID] AND [所有者ID]=[管理者ID]);"
'            Case "貸し手・借り手同一者解約"
'                If MsgBox("貸し手・借り手が同一である不正情報を解約しますか", vbYesNo) = vbYes Then SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].自小作別 = 0 WHERE ((([D:農地Info].自小作別)<>0) AND (([D:農地Info].所有者ID)=[借受人ID]));"
'            Case "住記情報の転送"
'                If Fnc.InputText("重要な変更", "パスワードを入力してください", "", 0, vbIMEModeOff) = "Avail" Then
'                    mvarPDW.WaitMessage = "住記情報のバックアップ転送"
'                    SystemDB.ExecuteSQL "UPDATE [D:個人Info] SET [D:個人Info].[住記世帯番号]=[世帯ID], [D:個人Info].住記続柄1 = [続柄1], [D:個人Info].住記続柄2 = [続柄2], [D:個人Info].住記続柄3 = [続柄3], [D:個人Info].住記住所 = [住所];"
'                    mvarPDW.WaitMessage = ""
'                End If
'            Case "所有／管理世帯の整合性": SystemDB.ExecuteSQL "UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID SET [D:農地Info].所有世帯ID = [世帯ID] WHERE ((([D:農地Info].所有世帯ID)=0 Or ([D:農地Info].所有世帯ID) Is Null) AND (([D:個人Info].世帯ID)<>0));", "更新中...."
'            Case "借受世帯の整合性": SystemDB.ExecuteSQL "UPDATE [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].借受人ID = [D:個人Info].ID SET [D:農地Info].借受世帯ID = [世帯ID] WHERE ((([D:農地Info].借受世帯ID)=0 Or ([D:農地Info].借受世帯ID) Is Null) AND (([D:個人Info].世帯ID)<>0));", "更新中...."
'            Case "データベースの更新": MsgBox "データベースが見つかりません", vbCritical
'            Case "農地面積の整合性"
'                mvarPDW.WaitMessage = "農地面積の整合中"
'                SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].田面積 = 0 WHERE ((([D:農地Info].現況地目) Not In (" & SystemDB.DBProperty("田地目") & ")));"
'                SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].畑面積 = 0 WHERE ((([D:農地Info].現況地目) Not In (" & SystemDB.DBProperty("畑地目") & ")));"
'                MsgBox "終了しました."
'                mvarPDW.WaitMessage = ""
'            Case Else
'                ADApplicationDLL_DoCommand St
'        End Select
'    End Function
'    
'    Private Function CDataviewSK_DoCommand2(sCmd As String, Optional sParam As String = "") As Variant
'        
'        Select Case sCmd
'    '/**********アプリケーション関連**********/
'            Case "申請変換": Sub申請変換
'            Case "プログラムの終了": SysAD.Quit
'            Case "HELPメニュー": CDataviewSK_DoCommand2 = ""
'            Case "バージョン": CDataviewSK_DoCommand2 = App.Major & "." & App.Minor & "." & App.Revision
'            Case "住記検索フリガナ初期化": Sub検索フリガナ作成
'            Case "住記テキストの設定":
'                St = Fnc.OpenFileDlg("住記テキストの指示", "テキストファイル|*.csv;*.txt", SystemDB.DBProperty("住記テキストの指示"))
'                If Fnc.Fs.FileExists(St) Then SystemDB.DBProperty("住記テキストの指示") = St
'            Case "エラーリスト作成":
'                Select Case Fnc.OptionSelect("所有者;登記面積;登記地目;地番不一致", "固定資産との突合せをします。")
'                    Case "所有者": mvarPDW.SQLListview.SQLListviewCHead "SELECT '所有者が違う' AS [事由],'農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, V_農地.所有者ID, '農地台帳所有者：' & [D:個人Info].氏名 AS [台帳内容], M_固定情報.所有者ID, '固定情報所有者：' & M_住民情報.氏名 AS [固定内容] FROM ((V_農地 INNER JOIN M_固定情報 ON (V_農地.大字ID = M_固定情報.大字ID) AND (V_農地.地番 = M_固定情報.地番)) LEFT JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) INNER JOIN M_住民情報 ON M_固定情報.所有者ID = M_住民情報.ID WHERE (((V_農地.所有者ID)<>[M_固定情報].[所有者ID]));", "土地の所在;土地所在;事由;事由;農地台帳所有者;台帳内容;固定情報所有者;固定内容", "土地のエラーリスト(所有者)"
'                    Case "登記地目": mvarPDW.SQLListview.SQLListviewCHead "SELECT '登記地目が違う' AS 事由, '農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, V_農地.登記簿地目, '農地台帳登記地目：' & [V_地目].[名称] AS 台帳内容, '固定情報登記地目：' & [V_地目2].[名称] AS 固定内容 FROM ((V_農地 INNER JOIN M_固定情報 ON (V_農地.地番 = M_固定情報.地番) AND (V_農地.大字ID = M_固定情報.大字ID)) INNER JOIN V_地目 AS V_地目2 ON M_固定情報.登記地目 = V_地目2.ID) INNER JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID WHERE (((V_農地.登記簿地目)<>[M_固定情報].[登記地目]));", "土地の所在;土地所在;事由;事由;農地台帳地目;台帳内容;固定情報地目;固定内容", "土地のエラーリスト(地目)"
'                    Case "登記面積": mvarPDW.SQLListview.SQLListviewCHead "SELECT '登記面積が違う' AS 事由, '農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, '農地台帳登記面積：' & [登記簿面積] AS 台帳内容, '固定情報登記面積：' & [登記面積] AS 固定内容 FROM V_農地 INNER JOIN M_固定情報 ON (V_農地.地番 = M_固定情報.地番) AND (V_農地.大字ID = M_固定情報.大字ID) WHERE (((V_農地.登記簿面積)<>[M_固定情報].[登記面積]));", "土地の所在;土地所在;事由;事由;農地台帳登記面積;台帳内容;固定情報登記面積;固定内容", "土地のエラーリスト(登記面積)"
'                    Case "地番不一致": mvarPDW.SQLListview.SQLListviewCHead "SELECT '同一地番無し' AS 事由, '農地.' & [V_農地].[ID] AS [KEY], V_農地.土地所在, M_固定情報.ID, '農地台帳：' & [V_農地].[地番] AS 台帳内容, '固定情報：同一地番なし' AS 固定内容 FROM V_農地 LEFT JOIN M_固定情報 ON (V_農地.地番 = M_固定情報.地番) AND (V_農地.大字ID = M_固定情報.大字ID) WHERE (((M_固定情報.ID) Is Null));", "土地の所在;土地所在;事由;事由;台帳内容;台帳内容;固定内容;固定内容", "土地のエラーリスト(地番不一致)"
'                End Select
'            Case "土地管理データのデータベースへの出力": Sub土地管理
'            Case "土地履歴と申請の不一致": List土地履歴と申請の不一致
'            Case "世帯自動生成"
'                SystemDB.ExecuteSQL "INSERT INTO [D:世帯Info] ( 世帯主ID, ID, 行政区ID ) " & _
'                "SELECT [D:個人Info].ID, [D:個人Info].世帯ID, [D:個人Info].行政区ID FROM [D:個人Info] LEFT JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID " & _
'                "WHERE ((([D:世帯Info].ID) Is Null) AND (([D:個人Info].世帯ID)>0) AND (([D:個人Info].続柄1)=" & SystemDB.DBProperty("世帯主続柄コード") & ") AND (([D:個人Info].住民区分)=" & SystemDB.DBProperty("記録住民コード") & "));"
'            Case "土地テキストの設定":
'                St = Fnc.OpenFileDlg("土地テキストの指示", "テキストファイル|*.csv;*.txt;*.DAT", SystemDB.DBProperty("土地テキストの指示"))
'                If Fnc.Fs.FileExists(St) Then
'                    SystemDB.DBProperty("土地テキストの指示") = St
'                End If
''            Case "世帯リスト": SysAD.MDIForm.ListViewCustom "農地基本台帳", "世帯リスト", ADApp.Common.AppData("世帯リスト")     '
'            Case "世帯員リスト": SysAD.MDIForm.ListViewCustom "農地基本台帳", "世帯員リスト", ADApp.Common.AppData("世帯員リスト")
'            Case "申請書受付中リスト": SysAD.MDIForm.ListViewCustom "農地基本台帳", "申請受付中リスト表示", ADApp.Common.AppData("申請受付中リスト表示")
'            Case "申請書審査中リスト": SysAD.MDIForm.ListViewCustom "農地基本台帳", "申請審査中リスト表示", ADApp.Common.AppData("申請審査中リスト表示")
'            Case "申請書許可済リスト": SysAD.MDIForm.ListViewCustom "農地基本台帳", "申請許可済リスト表示", ADApp.Common.AppData("申請許可済リスト表示")
'            Case "農地リスト":
'                SysAD.MDIForm.ListViewCustom "農地基本台帳", "農地リスト", ADApp.Common.AppData("農地リスト")
'            Case "期間満了の抽出": Sub期間満了
'            Case "小字テーブルの正規化"
'                SystemDB.ExecuteSQL "UPDATE [D:農地Info] INNER JOIN M_BASICALL ON [D:農地Info].小字ID = M_BASICALL.ID SET M_BASICALL.ParentKey = '大字.' & [大字ID] WHERE (((M_BASICALL.Class)='小字'));"
'            Case "管理情報の作成":
'                Dim n都道府県ID As Long, n支庁郡ID As Long, n市町村ID As Long
'                
'                St = Year(Now) + (Month(Now) < 4)
'                'St = 2006
'                n都道府県ID = Val(SystemDB.DBProperty("都道府県ID"))
'                n支庁郡ID = Val(SystemDB.DBProperty("支庁郡ID"))
'                n市町村ID = Val(SystemDB.DBProperty("市町村ID"))
'
'                If n都道府県ID = 0 Then
'                    MsgBox "都道府県IDが設定されていません。", vbCritical
'                    mvarPDW.DataviewCol.Add Me
'                ElseIf n支庁郡ID = 0 Then
'                    MsgBox "支庁郡IDが設定されていません。", vbCritical
'                    mvarPDW.DataviewCol.Add Me
'                ElseIf n市町村ID = 0 Then
'                    MsgBox "市町村IDが設定されていません。", vbCritical
'                    mvarPDW.DataviewCol.Add Me
'                Else
'                    Dim nCnt As Long
'                    Dim MCnt As Long
'                    Dim XX As Variant
'                    nCnt = 0
'                    
'                    If MsgBox("土地管理データをクリアします。" & vbCrLf & "よろしいですか？", vbQuestion + vbYesNo) = vbYes Then
'                        SystemDB.Execute "DELETE D_土地管理.* FROM D_土地管理;"
'                        SystemDB.Execute "DELETE D_土地管一筆.* FROM D_土地管一筆;"
'                        SystemDB.Execute "DELETE D_土地管補助.* FROM D_土地管補助;"
'                    Else
'                        MsgBox "処理を中止します。", vbCritical
'                        Exit Function
'                    End If
'                    
'                    mvarPDW.WaitMessage = "土地管理情報作成"
'                    mvarPDW.WaitMessage = "管理情報の作成中"
'                    Set Rs = SystemDB.GetRecordsetEx("SELECT * FROM [D_申請] WHERE [許可年月日]>=#" & St & "/1/1# AND [許可年月日]<=#" & St & "/12/31# AND [状態]=2 ORDER BY [許可年月日]", 0, , Me)
'                    Rs.MoveFirst
'                    XX = Split(Rs.GetString(vbBack, vbCrLf), vbCrLf)
'                    MCnt = UBound(XX)
'                    
'                    Rs.MoveFirst
'                    Do Until Rs.EOF
'                        nCnt = nCnt + 1
'                        mvarPDW.WaitMessage = nCnt & "/" & MCnt
'                        Sub土地管作成 Rs.ID, 0, n都道府県ID, n支庁郡ID, n市町村ID
'                        Rs.MoveNext
'                    Loop
'                    mvarPDW.WaitMessage = ""
'                    Rs.CloseRs
'                End If
'    '/*************選挙関連******************/
'            Case "選挙権確定者": mvarPDW.PrintGo ADApp.ObjectMan.GetObject("選挙権確定者")
'            Case "入場整理券": mvarPDW.PrintGo ADApp.ObjectMan.GetObject("選挙入場整理券")
'    '/****土地管理情報データベースの設定******/
'            Case "メニュー": CDataviewSK_PopupMenu
'            Case "市町村名の設定", "会長の設定", "農業委員会会長の設定": ADApplicationDLL_DoCommand "プロパティ"
'            Case "町外農業者の追加": CreateFarmer
'            Case "オンラインヘルプの表示": mvarPDW.ExplorerBar.Add ADApp.ObjectMan.GetObject("オンラインヘルプ.0")
'            Case "世帯番号の確認": sub世帯番号の確認
'            Case "管理情報の表示": ADApp.DataviewCol.Add ObjectMan.GetObjectDVP("土地管理.0")
'            Case "表示メニュー":
'                St = "非農地申請フォルダ;>一覧の設定;世帯リスト;世帯員リスト;農地リスト;選挙入力リスト;" & _
'                    ">申請書;申請書受付中リスト;申請書審査中リスト;申請書許可済リスト;" & IIf(Val(SystemDB.DBProperty("申請書の入力画面")), "標準申請入力;@拡張申請入力", "@標準申請入力;拡張申請入力") & _
'                    ";-;" & IIf(Val(SystemDB.DBProperty("申請書の昇順・降順")), "申請書昇順;@申請書降順", "@申請書昇順;申請書降順") & ";<;"
'                CDataviewSK_DoCommand2 = St
'            Case "申請書昇順": SystemDB.DBProperty("申請書の昇順・降順") = 0
'            Case "申請書降順": SystemDB.DBProperty("申請書の昇順・降順") = -1
'            Case "フリガナ修正": SystemDB.ExecuteSQL "UPDATE [D:個人Info] SET [D:個人Info].フリガナ = TRIM(StrConv([フリガナ],8));": MsgBox "終了"
'            Case "農地の検索": mvarPDW.ExplorerBar.Add ADApp.ObjectMan.GetObject("農地検索.1")
'            Case "農家一覧"
'                sSelecet = "SELECT '農家.' & [D:世帯INFO].[ID] AS [Key], [D:個人INFO.氏名] AS [名称], 'House' AS Icon, 'House' AS SelectIcon, [D:個人INFO].[住所] AS 住所, [D:世帯INFO].農家番号,[D:個人Info].[フリガナ], M_BASICALL.名称 AS [専兼区分], M_BASICALL.Class"
'                mvarPDW.SQLListview.SQLListviewCHead sSelecet & " FROM ([D:世帯INFO] LEFT JOIN [D:個人INFO] ON [D:世帯INFO].世帯主ID = [D:個人INFO].ID) LEFT JOIN (SELECT * FROM [M_BASICALL] WHERE [Class]='専兼区分') AS [M_BASICALL]  ON [D:世帯INFO].兼業区分 = M_BASICALL.ID " & sParam, "%2000,世帯主;名称;@農家番号;農家番号;フリガナ;フリガナ;$専兼区分;専兼区分;%3500,住所;住所", "農家一覧"
'            Case "農地一覧": view農地List sParam, "農地一覧" '/** 農地0001-0001-0002
'            Case "死亡名義管理者不明農地一覧": view農地List "((([D:個人Info].住民区分)=" & SystemDB.DBProperty("死亡住民コード") & ") AND ((V_農地.管理世帯ID) Is Null Or (V_農地.管理世帯ID)=0))", "死亡名義管理者不明農地一覧" '/** 農地0001-0001-0002
'            Case "不在地主名義管理者不明農地一覧":
'                St = Fnc.InputText("不在地主名義管理者不明農地一覧", "住民区分を入力してください")
'                If Len(St) Then view農地List "[D:個人Info].[住所] LIKE '%" & SystemDB.DBProperty("市町村名") & "%' AND ((([D:個人Info].住民区分) IN (" & St & ")) AND ((V_農地.管理世帯ID) Is Null Or (V_農地.管理世帯ID)=0))", "不在地主名義管理者不明農地一覧" '/** 農地0001-0001-0002
'            Case "世帯番号からの検索": Find世帯With世帯番号 0
'            Case "世帯番号からの営農情報検索": Find世帯With世帯番号 1
'            Case "町外農業者の作成": CreateFarmer -1
'            Case "住民番号からの検索": Find個人With住民番号
'            Case "生年月日からの検索": Find個人With生年月日
'            Case "農地ＩＤからの検索": Find農地With農地番号
'            Case "一筆コードからの検索": Find農地With固定番号
'            Case "プロパティ": mvarPDW.DataviewCol.Add Me
'            Case "検索メニュー": CDataviewSK_DoCommand2 = "農家検索;農地の検索;-;世帯番号からの検索;世帯番号からの営農情報検索;-;住民番号からの検索;生年月日からの検索;該当世帯番号から住民の検索;-;農地ＩＤからの検索;一筆コードからの検索;"
'            Case "該当世帯番号から住民の検索"
'                St = Fnc.InputText("世帯番号", "世帯番号を入力してください", , 1, 0)
'                If Not Val(St) = 0 Then SQLList世帯員 "[世帯ID]=" & Val(St), "住民番号=" & Val(St) & "の検索", True
'            Case "農家検索", "農業者検索", "法人の検索": mvarPDW.ExplorerBar.Add ADApp.ObjectMan.GetObject("農家検索.1")

'            Case "基本集計":
'                If 市内外の調整() Then Set pPrint = ADApp.ObjectMan.GetObject("基本集計"): pPrint.Printview ""
'            Case "住民区分別集計": sub住民区分別集計

'            Case "集落集計参照": DVProperty.Controls.Value("TX集落集計") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("集落集計定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX集落集計"))
'            Case "基本集計参照": DVProperty.Controls.Value("TX基本集計") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("基本集計定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX基本集計"))
'            Case "住記テキストの読込開始": Load住記Text
'            Case "利用権終期農地一覧参照": DVProperty.Controls.Value("TX利用権終期農地一覧") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("利用権終期農地一覧定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX利用権終期農地一覧"))
'            Case "基幹DBファイル参照": DVProperty.Controls.Value("TX基幹DBファイル") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("基幹DBファイルの設定", "*.MDB"), DVProperty.Controls.Value("TX基幹DBファイル"))
'            Case "意見書ファイル参照": DVProperty.Controls.Value("TX意見書ファイル") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("意見書ファイルの設定", "*.XLS"), DVProperty.Controls.Value("TX意見書ファイル"))
'            Case "農業者検索ﾌｫｰﾑ参照": DVProperty.Controls.Value("TX農業者検索ﾌｫｰﾑ") = Fnc.CanGetFileName(SysAD.MDIForm.OpenFileDlg("農業者検索ﾌｫｰﾑ定義ファイルの設定", "*.Htm"), DVProperty.Controls.Value("TX農業者検索ﾌｫｰﾑ"))
'            Case "封筒印刷": mvarPDW.PrintGo ADApp.ObjectMan.GetObject("封筒印刷")
'            Case "５年以上貸借"
'                sFrom = Fnc.InputText("基準日", "審査の基準日を入力してください", Date)
'                If Len(sFrom) = 0 Then Exit Function
'                If Not IsDate(sFrom) Then Exit Function
'                    
'                St = "SELECT '農地.' & [V_農地].ID AS [Key],V_農地.小作終了年月日, V_農地.小作開始年月日, [D:個人Info].氏名 AS 借受者, V_農地.土地所在, [D:個人Info_1].氏名 AS 所有者, IIf([小作地適用法]=1,'農','基') AS 法令, IIf([小作形態]=1,'賃','使') AS 形態, DateDiff('yyyy',[小作開始年月日],[小作終了年月日]) AS 期間 " & _
'                "FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.借受人ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.所有者ID = [D:個人Info_1].ID " & _
'                "WHERE ([V_農地].自小作別 > 0) AND ([V_農地].小作終了年月日 > CDate('" & sFrom & "')) AND ([V_農地].小作開始年月日 <= CDate('" & DateSerial(Year(CDate(sFrom)) - 5, Month(CDate(sFrom)), Day(CDate(sFrom))) & "')) " & _
'                "ORDER BY V_農地.小作終了年月日, [D:個人Info].氏名;"
'                                            
'                mvarPDW.SQLListview.SQLListviewCHead St, "土地所在;土地所在;小作終了年月日;小作終了年月日;小作開始年月日;小作開始年月日;@期間(年);期間;借受者;借受者;所有者;所有者;法令;法令;形態;形態", ""
'            Case "３年以上貸借"
'                sFrom = Fnc.InputText("基準日", "審査の基準日を入力してください", Date)
'                If Len(sFrom) = 0 Then Exit Function
'                If Not IsDate(sFrom) Then Exit Function
'                    
'                St = "SELECT '農地.' & [V_農地].ID AS [Key],V_農地.小作終了年月日, V_農地.小作開始年月日, [D:個人Info].氏名 AS 借受者, V_農地.土地所在, [D:個人Info_1].氏名 AS 所有者, IIf([小作地適用法]=1,'農','基') AS 法令, IIf([小作形態]=1,'賃','使') AS 形態, DateDiff('yyyy',[小作開始年月日],[小作終了年月日]) AS 期間 " & _
'                    "FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.借受人ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.所有者ID = [D:個人Info_1].ID " & _
'                    "WHERE ([V_農地].自小作別 > 0) AND ([V_農地].小作終了年月日 > CDate('" & sFrom & "')) AND ([V_農地].小作開始年月日 <= CDate('" & DateSerial(Year(CDate(sFrom)) - 3, Month(CDate(sFrom)), Day(CDate(sFrom))) & "')) " & _
'                    "ORDER BY V_農地.小作終了年月日, [D:個人Info].氏名;"
'                                            
'                mvarPDW.SQLListview.SQLListviewCHead St, "土地所在;土地所在;小作終了年月日;小作終了年月日;小作開始年月日;小作開始年月日;@期間(年);期間;借受者;借受者;所有者;所有者;法令;法令;形態;形態", ""
'            Case "住記の整合性": mod住記設定.sub住記整合性

'    '/****************************　取込関連 **************************/
'            Case "集落別世帯数一覧": subSQLList集落別世帯一覧
'            Case "農地-遊休化", "農地-無断転用"
'                sDate = Fnc.InputText("調査年月日", "調査年月日を入力してください", Date, 1, 2)
'                If IsDate(sDate) And Len(sParam) Then
'                    Ar = Split(sParam, ";")
'                    For n = 0 To UBound(Ar)
'                        Set pDV = ADApp.ObjectMan.GetObject(CStr(Ar(n)))
'                        pDV.DoCommand2 MID$(sCmd, InStr(sCmd, "-") + 1), sDate
'                    Next
'                    mvarPDW.SQLListview.Refresh
'                End If
'            Case "台帳管理システムについて": SysAD.About "市町村:" & SystemDB.DBProperty("市町村名") & vbCrLf & "DB:"
'
'            Case "農地-樹園地へ変換"
'                Ar = Split(sParam, ";")
'                For n = 0 To UBound(Ar)
'                    Set pDV = ADApp.ObjectMan.GetObject(CStr(Ar(n)))
'                    pDV.DoCommand2 "樹園地設定"
'                Next
'                    mvarPDW.SQLListview.Refresh
'            Case "農地-同一条件の貸借設定"
'                Dim Rg As ADBasicLIB2.RecordsetEx
'                St = GetIDList(sParam, "農地")
'                Ar = Split(St, ",")
'                Set Rs = SystemDB.GetRecordsetEx("SELECT Right$('0000000000' & [ID],10) & ':' & [土地所在] & '(' & [小作開始年月日] & '～' & [小作終了年月日] & ')' FROM [V_農地] WHERE [ID] IN (" & St & ")", , , Me)
'                SS = Val(Fnc.OptionSelect(Rs.GetString(":", ";"), "基準になる農地を選択してください", ""))
'                SystemDB.CloseRs Rs
'                If Val(SS) Then
'                    Set Rg = SystemDB.GetRecordsetEx("SELECT * FROM [D:農地Info] WHERE [ID]=" & SS, , , Me)
'                    If Not Rg.IsNoRecord Then
'                        Set Rs = SystemDB.GetRecordsetEx("SELECT * FROM [D:農地Info] WHERE [ID] IN (" & St & ")", , , Me)
'                        
'                        Do Until Rs.EOF
'                        If Rs.ID = Rg.ID Then
'                        Else
'                            Rs.Copy Rg, GetArray("自小作別", "借受世帯ID", "借受人ID", "小作地適用法", "小作形態", "小作開始年月日", "小作終了年月日", "小作料", "小作料単位")
'                        End If
'                        Rs.MoveNext
'                        Loop
'                        SystemDB.CloseRs Rs
'                    End If
'                    mvarPDW.SQLListview.Refresh
'                    Rg.CloseRs
'                End If
'            Case "農地-終了履歴付き強制解約" ' In 農地情報DLL
'                sDate = Fnc.InputText("終了年月日", "終了年月日を入力してください", Date, 1, 2)
'                If IsDate(sDate) And Len(sParam) Then
'                    Ar = Split(sParam, ";")
'                    For n = 0 To UBound(Ar)
'                        ID = Fnc.GetKeyCode(Ar(n))
'                        SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].自小作別 = 0 WHERE [D:農地Info].[ID]=" & ID
'                        SystemDB.ExecuteSQL "INSERT INTO D_土地履歴 ( LID, 異動事由, 異動日, 更新日, 内容 ) VALUES(" & ID & ", 10201 , #" & Format(CDate(sDate), "MM/DD/YYYY") & "# , Date() , '利用権満期終了');"
'                    Next
'                    mvarPDW.SQLListview.Refresh
'                End If
'            Case "終了時刻表示":
'                St = Fnc.InputText("終了時刻設定", "本日の終了時刻を表示します" & vbCrLf & "例、17:30の場合　[17:30]")
'                If IsDate(St) Then
'                    SystemDB.DBProperty("終了時刻") = DateSerial(Year(Now), Month(Now), Day(Now)) + TimeSerial(Hour(CDate(St)), Minute(CDate(St)), 0)
'                Else
'                    MsgBox "時刻の設定が不正です", vbCritical
'                End If
'            Case "農地-解約履歴付き強制解約" ' In 農地情報DLL
'                sDate = Fnc.InputText("終了年月日", "終了年月日を入力してください", Date, 1, 2)
'                If IsDate(sDate) And Len(sParam) Then
'                    Ar = Split(sParam, ";")
'                    For n = 0 To UBound(Ar)
'                        ID = Fnc.GetKeyCode(Ar(n))
'                        SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].自小作別 = 0 WHERE [D:農地Info].[ID]=" & ID
'                        SystemDB.ExecuteSQL "INSERT INTO D_土地履歴 ( LID, 異動事由, 異動日, 更新日, 内容 ) VALUES(" & ID & ", 10202 , #" & Format(CDate(sDate), "MM/DD/YYYY") & "# , Date() , '利用権解約終了');"
'                    Next
'                    mvarPDW.SQLListview.Refresh
'                End If
'
'            Case "固定データの照合": mod固定データ照合
'            Case "ツールメニュー": CDataviewSK_DoCommand2 = ADApp.Common.SetMenuCmd("ツール", "申請の整合性;>集計;集落別世帯数一覧;<;-;>選挙関連;選挙権の初期化;<;>エラーリスト作成;固定データの照合;<;支所番号の設定;-;終了時刻表示;-;転作データ作成;テストダイアログ;")
'            Case "申請の整合性": sub申請の整合性
'            Case "選挙権の初期化":
'                St = SystemDB.DBProperty("選挙権クリア年月日")
'                If IsDate(St) Then
'                    If MsgBox("選挙権の初期化を最後に行ったのは[" & St & "]ですが、初期化しますか", vbYesNo) = vbYes Then
'                        SystemDB.DBProperty("選挙権クリア年月日") = Date
'                        SystemDB.ExecuteSQL "UPDATE [D:個人Info] SET [D:個人Info].選挙権の有無 = False, [D:個人Info].前年選挙権 = [選挙権の有無];"
'                    End If
'                ElseIf MsgBox("選挙権を初期化します", vbYesNo) = vbNo Then
'                Else
'                    SystemDB.DBProperty("選挙権クリア年月日") = Date
'                    SystemDB.ExecuteSQL "UPDATE [D:個人Info] SET [D:個人Info].選挙権の有無 = False, [D:個人Info].前年選挙権 = [選挙権の有無];"
'                End If
'            Case "買受適格証明願": mvarPDW.PrintGo ObjectMan.GetObject("買受適格証明願")
'            Case "３条受理通知書": mvarPDW.PrintGo ObjectMan.GetObject("受理通知書"), "3条"
'            Case "４条受理通知書": mvarPDW.PrintGo ObjectMan.GetObject("受理通知書"), "4条"
'            Case "５条受理通知書": mvarPDW.PrintGo ObjectMan.GetObject("受理通知書"), "5条"
'            Case "適格要件届出書": mvarPDW.PrintGo ObjectMan.GetObject("適格要件届出書")
'            
'            Case "３条届出書": mvarPDW.PrintGo ObjectMan.GetObject("３条届出書")
'            Case "転用申請チェック表": mvarPDW.PrintGo ObjectMan.GetObject("転用申請チェック表")
'            Case "営農計画書": mvarPDW.PrintGo ADApp.ObjectMan.GetObject("営農計画書")
'            Case "奨励金交付申請書": mvarPDW.PrintGo ADApp.ObjectMan.GetObject("奨励金交付申請書")

'            Case "集落別選挙人集計"
'                St = "SELECT [D:個人Info].投票区, '行政区.' & [行政区ID] AS [Key], [V_行政区].ID AS 行政区ID, [M:投票区].選挙区, [V_行政区].行政区, Count([D:個人Info].ID) AS 人数 " & _
'                "FROM ([M:投票区] INNER JOIN [D:個人Info] ON [M:投票区].ID = [D:個人Info].投票区) INNER JOIN [V_行政区] ON [D:個人Info].行政区ID = [V_行政区].ID " & _
'                "GROUP BY [D:個人Info].投票区, '行政区.' & [行政区ID], [V_行政区].ID, [M:投票区].選挙区, [V_行政区].行政区, [D:個人Info].選挙権の有無 " & _
'                "HAVING ((([D:個人Info].選挙権の有無)=True)) ORDER BY [D:個人Info].投票区, [V_行政区].ID;"
'                mvarPDW.SQLListview.SQLListviewCHead St, "投票区;選挙区;行政区;行政区;@人数;人数", "集落別選挙人数集計"
'            Case "選挙人リスト一覧"
'                St = "SELECT '世帯員.' & [D:個人Info].ID AS [KEY], [D:個人Info].投票区, [M:投票区].選挙区 AS [投票区名], [V_行政区].ID AS 行政区ID, [V_行政区].行政区, [D:個人Info_1].フリガナ AS [ソート用], [V_続柄].ID AS 順序, [D:個人Info].フリガナ, [D:個人Info].氏名, [D:個人Info].住所, [V_続柄].続柄 AS 続柄A, [V_続柄_1].続柄 AS 続柄B, [D:個人Info].生年月日 " & _
'                "FROM [D:個人Info] AS [D:個人Info_1] INNER JOIN (((([M:投票区] INNER JOIN [D:個人Info] ON [M:投票区].ID = [D:個人Info].投票区) INNER JOIN [V_続柄] ON [D:個人Info].続柄1 = [V_続柄].ID) INNER JOIN [V_続柄] AS [V_続柄_1] ON [D:個人Info].続柄2 = [V_続柄_1].ID) INNER JOIN [V_行政区] ON [D:個人Info].行政区ID = [V_行政区].ID) ON [D:個人Info_1].世帯ID = [D:個人Info].世帯ID " & _
'                "WHERE ((([D:個人Info].選挙権の有無)=True) AND (([D:個人Info_1].続柄1)=1) AND (([D:個人Info_1].住民区分)=0)) " & _
'                "ORDER BY [D:個人Info].投票区, [V_行政区].ID, [D:個人Info_1].フリガナ, [V_続柄].ID; "
'                mvarPDW.SQLListview.SQLListviewCHead St, "氏名;氏名;投票区名;投票区名;行政区;行政区;フリガナ;フリガナ;住所;住所;続柄1;続柄A;続柄2;続柄B;=生年月日;生年月日", "選挙人リスト一覧"
'            Case "証明願い"

'            Case "標準申請入力": SystemDB.DBProperty("申請書の入力画面") = 0
'            Case "拡張申請入力": SystemDB.DBProperty("申請書の入力画面") = -1
'            Case "選挙CSVファイル参照": DVProperty.Controls.Value("TX選挙CSVファイル") = Fnc.CanGetFileName(SysAD.MDIForm.SaveFileDlg("選挙CSVファイルの設定", "*.CSV"), DVProperty.Controls.Value("TX選挙CSVファイル"))
'            Case "FileExport"
'                St = SystemDB.DBProperty("選挙CSVファイル")
'                Set Rs = SystemDB.GetRecordsetEx(sParam, , , Me)
'                Set Ts = Fnc.Fs.CreateTextFile(St, ForWriting, True)
'                    Ts.WriteS Rs.GetString(",", vbCrLf)
'                Ts.CloseFile
'                MsgBox "[" & St & "]の書込みが終了しました"
'                SystemDB.CloseRs Rs
'            Case "支所番号の設定": mod農家台帳.sub支所番号設定
'            Case "基幹データの取込": mod住記設定.sub住記読込み
'            Case "現地調査日程表", "現地調査担当委員": mvarPDW.PrintGo ADApp.ObjectMan.GetObject(sCmd), sCmd
'            Case "現地調査管理"
'                Dim p現地Frm As New frm現地調査日程表
'                With Fnc.InputMulti("現地調査日程表", "入力をお願いします。", "今月総会番号;数値;" & Val(SystemDB.DBProperty("今月総会番号")), , True)
'                    If .IsALLEmpty Then
'                    Else
'                        SystemDB.DBProperty("今月総会番号") = Val(.Item("今月総会番号"))
'                        Select Case p現地Frm.SortData(Val(.Item("今月総会番号")))
'                            Case 1:
'                            Case 2:
'                        End Select
'                    End If
'                End With
'            Case "現地調査報告書":
'                St = Fnc.InputText("総会番号", "総会番号を入力してください", SystemDB.DBProperty("今月総会番号"), 0, vbIMEModeOff)
'                If St <> "" Then
'                    mvarPDW.PrintGo ObjectMan.GetObject("現地調査報告書"), St
'                End If
'            Case "農地-農用地内へ変更":
'                If Len(sParam) Then
'                    Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
'                    If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を農用地内へ変更しますか", vbYesNo) Then
'                        SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 1 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));"
'                    End If
'                End If
'            Case "農地-農振地内へ変更":
'                If Len(sParam) Then
'                    Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
'                    If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を農振地へ変更しますか", vbYesNo) Then
'                        SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 0 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));"
'                    End If
'                End If
'            Case "農地-農振除外":
'                If Len(sParam) Then
'                    Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
'                    If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を農振除外しますか", vbYesNo) Then
'                        SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 2 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));"
'                    End If
'                End If
'            Case "筆選択"
'            Case "農地-一括で転用農地へ（申請無し）"
'                If Len(sParam) Then
'                    Ar = Split(Replace(Replace(sParam, "農地.", ""), ";", ","), ",")
'                    If MsgBox("選択した" & UBound(Ar) + 1 & "筆の農地を転用農地にしますか", vbYesNo) Then
'                        SystemDB.ExecuteSQL "UPDATE [D:農地Info] SET [D:農地Info].農業振興地域 = 2 WHERE ((([D:農地Info].ID) In (" & Join(Ar, ",") & ")));"
'                        For n = 0 To UBound(Ar)
'                            If Len(Ar(n)) Then
'                                SystemDB.ExecuteSQL "INSERT INTO D_土地履歴 ( LID, 更新日, 内容 ) VALUES (" & Val(Ar(n)) & ",Date(),'職権による転用修正')"
'                                SystemDB.ExecuteSQL "INSERT INTO D_転用農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & Val(Ar(n)) & "));"
'                                SystemDB.ExecuteSQL "DELETE [D:農地Info].ID FROM [D:農地Info] WHERE [D:農地Info].ID=" & Val(Ar(n))
'                            End If
'                        Next
'                        mvarPDW.SQLListview.Refresh
'                    End If
'                End If
'            Case "転作データ作成":
'                Dim sFName1 As String
'                Dim sFName2 As String
'                Dim d個人 As New CDataDictionary
'                
'                Ar = ADApp.Common.MasterTable.ItemArray("大字")
'                St = Replace(Join(Ar, ";"), ",", ":")
'                
'                St = Fnc.SelectMultiString(St, "大字を選択してください")
'                If Len(St) Then
'                    sFName1 = Fnc.OpenFileDlg("転作用データ出力", "農地のファイル名を指定してください(*.csv)", CreateObject("WScript.Shell").SpecialFolders("DeskTop") & "\")
'                    sFName2 = Fnc.OpenFileDlg("転作用データ出力", "個人のファイル名を指定してください(*.csv)", CreateObject("WScript.Shell").SpecialFolders("DeskTop") & "\")
'                    
'                    Ar = Split(St, ";")
'                    For n = 0 To UBound(Ar)
'                        If Len(Ar(n)) Then
'                            Ar(n) = Left$(Ar(n), InStr(Ar(n), ":") - 1)
'                        End If
'                    Next
'                    St = Join(Ar, ",")
'                    If Right$(St, 1) = "," Then St = Left$(St, Len(St) - 1)
'                    
'                    If Len(sFName1) And sFName1 <> "No Select" Then
'                        Set Rs = SystemDB.GetRecordsetEx("SELECT IIF([V_農地].管理者ID<>0,[V_農地].管理世帯ID,[V_農地].所有世帯ID) AS [所有世帯ID],[管理人ID] AS [所有者ID],IIF([自小作別]>0,[借受人ID],0) AS [借受者ID], [土地所在], 田面積 AS [面積] FROM [V_農地] WHERE [大字ID] IN (" & St & ") AND [田面積]>0 ORDER BY [大字ID],[小字ID],[地番]", 0, , Me)
'                        sData = Join(Rs.FildeHeaders, ",")
'                        If Not Rs.EOF Then
'                            sData = sData & vbCrLf & Rs.GetString(",")
'                        End If
'                        Rs.CloseRs
'                        Set Ts = Fnc.Fs.OpenTextFile(sFName1, ForWriting, True)
'                        Ts.WriteS sData
'                        Ts.CloseFile
'                    End If
'                    
'                    If Len(sFName2) And sFName2 <> "No Select" Then
'                        Set Rs = SystemDB.GetRecordsetEx("SELECT [V_農地].管理人ID FROM [V_農地] WHERE [大字ID] IN (" & St & ") AND [田面積]>0 AND [管理人ID]<>0 GROUP BY [V_農地].管理人ID;", 0, , Me)
'                        Do Until Rs.EOF
'                            If Not d個人.Exists(Rs.Value("管理人ID")) Then
'                                d個人.Add Rs.Value("管理人ID"), Rs.Value("管理人ID")
'                            End If
'                            Rs.MoveNext
'                        Loop
'                        Rs.CloseRs
'                        Set Rs = SystemDB.GetRecordsetEx("SELECT [V_農地].借受人ID FROM [V_農地] WHERE [大字ID] IN (" & St & ") AND [自小作別]>0 AND [田面積]>0 AND [借受人ID]<>0 GROUP BY [V_農地].借受人ID;", 0, , Me)
'                        Do Until Rs.EOF
'                            If Not d個人.Exists(Rs.Value("借受人ID")) Then
'                                d個人.Add Rs.Value("借受人ID"), Rs.Value("借受人ID")
'                            End If
'                            Rs.MoveNext
'                        Loop
'                        Rs.CloseRs
'
'                        Set Rs = SystemDB.GetRecordsetEx("SELECT [D:個人Info].行政区ID, [V_行政区].名称 AS [行政区], [D:個人Info].ID AS [個人ID], [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].生年月日 FROM [D:個人Info] LEFT JOIN [V_行政区] ON [D:個人Info].行政区ID = [V_行政区].ID ORDER BY [D:個人Info].行政区ID, [D:個人Info].[フリガナ], [D:個人Info].[ID];", 0, , Me)
'                        sData = Join(Rs.FildeHeaders, ",")
'                        Do Until Rs.EOF
'                            If d個人.Exists(Rs.Value("個人ID")) Then sData = sData & vbCrLf & Rs.GetLineString(GetArray("行政区ID", "行政区", "個人ID", "氏名", "住所", "生年月日"), ",")
'                            Rs.MoveNext
'                        Loop
'                        Rs.CloseRs
'                        Set Ts = Fnc.Fs.OpenTextFile(sFName2, ForWriting, True, TristateTrue)
'                        Ts.WriteLine sData
'                        Ts.CloseFile
'                    End If
'                End If

'            Case Else
'                If CheckBoolParams(sCmd, "小作料の表示;コードの表示;住民区分の表示;住民区分コードの表示;年齢の表示;選挙権の表示;農地確認の表示", "農地リストの小作料表示;世帯員リストのコード表示;世帯員リストの住民区分表示;世帯員リストの住民区分コード表示;世帯員リストの年齢表示;世帯員リストの選挙権表示;世帯リストの確認表示;農地リストの確認表示") Then
'                ElseIf Len(SystemDB.DBProperty(sCmd)) Then
'                    Ar = Split(SystemDB.DBProperty(sCmd), ";")
'                    Select Case UBound(Ar)
'                        Case 0: CDataviewSK_DoCommand2 CStr(Ar(0))
'                        Case Else: CDataviewSK_DoCommand2 CStr(Ar(0)), CStr(Ar(1))
'                    End Select
'                Else
'                    Debug.Assert CaseAssertPrint(sCmd)
'                End If
'        End Select
'
'    End Function
