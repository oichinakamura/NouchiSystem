
Imports HimTools2012

Public Class CTabPage自然解約実行
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView
    Private mvarDT As Date
    Public Sub New(p期限 As DateTime)
        MyBase.new(True, True, "自然解約の実行", "自然解約の実行")
        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView
        mvarDT = p期限
        Dim sWHERE As String = String.Format("[自小作別]>0 and [小作地適用法]=2 and [小作終了年月日]<#{0}/{1}/{2}#", p期限.Month, p期限.Day, p期限.Year)
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT True AS 対象, * FROM [D:農地Info] WHERE " & sWHERE)
        App農地基本台帳.TBL農地.MergePlus(pTBL)

        mvarGrid.SetDataView(App農地基本台帳.TBL農地.Body, sWHERE, "")
        mvarGrid.AutoGenerateColumns = False
        mvarGrid.AllowUserToAddRows = False
        mvarGrid.AllowUserToDeleteRows = False

        With mvarGrid
            .AddColumnCheckBox("対象", "対象", "対象", enumReadOnly.bCanEdit, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("ID", "ID", "ID", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("大字", "大字", "大字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小字", "小字", "小字", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("地番", "地番", "地番", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("登記簿面積", "登記簿面積", "登記簿面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("実面積", "実面積", "実面積", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("登記簿地目名", "登記簿地目名", "登記簿地目名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("現況地目名", "現況地目名", "現況地目名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("所有者氏名", "所有者氏名", "所有者氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("自小作", "自小作", "自小作", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("借受人氏名", "借受人氏名", "借受人氏名", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("貸借始期", "貸借始期", "貸借始期", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("貸借終期", "小作終了年月日", "小作終了年月日", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料", "小作料", "小作料", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AddColumnText("小作料単位", "小作料単位", "小作料単位", enumReadOnly.bReadOnly, DataGridViewContentAlignment.MiddleRight)
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With
        Me.ControlPanel.Add(mvarGrid)

        AddHandler Me.ToolStrip.Items.Add("自然解約の実行").Click, AddressOf 実行
        AddHandler Me.ToolStrip.Items.Add("エクセル出力").Click, AddressOf mvarGrid.ToExcel
    End Sub

    Public Sub 実行()
        If MsgBox("処理を開始しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            With mvarDT
                For Each pRow As DataGridViewRow In mvarGrid.Rows
                    If Val(pRow.Cells("対象").Value.ToString) = True Then
                        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_土地履歴 ( LID, 異動日, 更新日, 入力日, 異動事由, 内容 ) SELECT [D:農地Info].ID, [D:農地Info].小作終了年月日, Date() AS 式1, Date() AS 式2, 10201 AS 式3, '[' & [氏名] & ']への貸借の期間満了による終了' AS 式4 FROM [D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].借受人ID = [D:個人Info].ID WHERE ((([D:農地Info].ID) In ({0})) AND (([D:農地Info].小作終了年月日)<=#{1}/{2}/{3}#) AND (([D:農地Info].自小作別)>0) AND (([D:農地Info].小作地適用法)=2));", pRow.Cells("ID").Value, .Month, .Day, .Year)
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [D:農地Info].自小作別 = 0 WHERE ((([D:農地Info].自小作別)>0) AND (([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作終了年月日)<=#{0}/{1}/{2}#) AND (([D:農地Info].ID) In ({3})));", .Month, .Day, .Year, pRow.Cells("ID").Value)
                    End If
                Next

                App農地基本台帳.TBL農地.Body.Clear()
            End With
            MsgBox("終了しました")
        End If
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

End Class
