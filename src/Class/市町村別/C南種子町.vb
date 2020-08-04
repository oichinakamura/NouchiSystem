'20160411霧島

Public Class C南種子町
    Inherits C市町村別

    '\\10.81.8.8\MAPFolder\システム配置\南種子町\
    'MAPServer
    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function

    Public Overrides Sub InitLocalData()
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [S:作業者登録]")
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE S_システムデータ SET S_システムデータ.DATA = '10' WHERE (((S_システムデータ.KEY)='田地目'));")
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE S_システムデータ SET S_システムデータ.DATA = '20' WHERE (((S_システムデータ.KEY)='畑地目'));")

        ''TODO 2015/09/18作成、2015/09/30まで有効
        'Dim pView As New DataView(pTBL, "[登録者名]='羽生　幸一'", "", DataViewRowState.CurrentRows)
        'If pView.Count > 0 Then
        '    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [S:作業者登録] WHERE [ID]=1")
        '    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [S:作業者登録] WHERE [ID]=3")

        '    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [S:作業者登録]([ID],[登録者名],[PW],[権利フラグ]) VALUES(1,'古市　義朗','0466',1);")
        '    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [S:作業者登録]([ID],[登録者名],[PW],[権利フラグ]) VALUES(3,'園田　孝太郎','7939',1);")
        '    pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [S:作業者登録]")
        'End If

     

        With New dlgLoginForm(pTBL)
            If Not .ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try
                    End
                Catch ex As Exception

                End Try
            End If
        End With

        SysAD.SystemInfo.ユーザー.n権利 = 1
    End Sub

    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()
            .ListView.ItemAdd("農家検索", "農地・農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            .ListView.ItemAdd("利用権終期台帳", "利用権終期台帳", "印刷", "印刷", AddressOf sub利用権終期管理)
            .ListView.ItemAdd("総会資料作成", "総会資料作成", ImageKey.作業, "操作", AddressOf ClickMenu)
            .ListView.ItemAdd("受付公布簿", "受付公布簿", "印刷", "印刷", Sub(s, e) SysAD.MainForm.MainTabCtrl.ExistPage("受付公布簿", True, GetType(CTabPage受付公布簿)))
            .ListView.ItemAdd("所有権売買価格一覧表", "所有権売買価格一覧表", "印刷", "印刷", Sub(s, e) SysAD.MainForm.MainTabCtrl.ExistPage("所有権売買価格一覧表", True, GetType(CTabPage3条所有権売買価格一覧表出力)))
            .ListView.ItemAdd("現地調査表作成", "現地調査表作成", ImageKey.作業, "操作", AddressOf 現地調査)
            .ListView.ItemAdd("南種子町利用意向調査出力", "南種子町利用意向調査出力", "印刷", "印刷", AddressOf 意向調査H26)
            'If Now.Date < #12/5/2014# Then
            '    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [M_BASICALL] WHERE [ID]=100009 AND [Class]='土地異動事由'")
            '    If pTBL.Rows.Count = 0 Then
            '        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO M_BASICALL([ID],[Class],[名称]) VALUES(100009,'土地異動事由','所有者変更')")
            '    End If
            'End If
            '.ListView.ItemAdd("TST", "TST", "操作", "操作", AddressOf TSTG)
            .ListView.ItemAdd("フリガナ初期化", "フリガナ初期化", ImageKey.作業, "操作", AddressOf subフリガナ初期化)
            .ListView.ItemAdd("利用権等実績表出力", "利用権等実績表出力", "印刷", "印刷", AddressOf SUB利用権等実績)

            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)

            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [続柄1]=-4,[続柄2]=0,[続柄3]=0 WHERE [続柄1]=4")

            MyBase.InitMenu(pMain)
        End With
    End Sub

    Public Sub New()
        MyBase.New("南種子町")
    End Sub

    Private Sub subフリガナ初期化()
        SysAD.DB(sLRDB).ExecuteSQL("Update [D:個人Info] SET [検索フリガナ]=Null")

    End Sub

    Private Sub 意向調査H26()
        If Not SysAD.MainForm.MainTabCtrl.ExistPage("利用意向調査") Then
            SysAD.MainForm.MainTabCtrl.AddPage(New CPage利用意向調査H26南種子)
        End If
    End Sub

    Private Sub SUB利用権等実績()
        Dim mvar出力 As New CDataToExcel("利用権等実績.xls")

        My.Application.DoEvents()
    End Sub

    Private mvar利用権終期管理 As C利用権終期管理


End Class
