
Public Class C相続未登記等農地調査
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private SContainer As New SplitContainer
    Private TSContainer1 As New ToolStripContainer
    Private TStrip1 As New ToolStrip
    Private TSLabel1 As New ToolStripLabel("様式Ⅰ")
    Private WithEvents TSBtnExcel1 As New ToolStripButton("Excel出力")
    Private mvarGrid1 As New HimTools2012.controls.DataGridViewWithDataView
    Private TSContainer2 As New ToolStripContainer
    Private TStrip2 As New ToolStrip
    Private TSLabel2 As New ToolStripLabel("様式Ⅱ")
    Private WithEvents TSBtnExcel2 As New ToolStripButton("Excel出力")
    Private mvarGrid2 As New HimTools2012.controls.DataGridViewWithDataView

    Private TBL相続未登記Info As DataTable
    Private TBL所有者Info As DataTable
    Private TBL重複所有者 As DataTable
    Private View相続未登記Info As DataView
    Private TBL相続Format1 As New DataTable
    Private TBL相続Format2 As New DataTable

    Private Ar死亡 As String() = {"死亡", "死亡者", "死亡所有者", "死亡所有", "死亡(世帯主)"}
    Private Ar共有 As String() = {"共有", "共有者", "共有名義"}
    Private Arその他 As String() = {"-", "市外住民", "市外居住者", "町外住民", "町外者", "住登外", "住登外（住）", "住登外（法人）", "住登外（共有者）", "住登外個人", "住登外法人", "未登録住民", "不明住民", "転出予定者", "転出", "転出者", "転出確定者", "転出確認", "転居出者", "職権消除者", "職権消除", "その他消除者", "消除（転出等）", "除票（住）", "その他除票者", "戸籍登載者", "戸籍転籍除籍者", "戸籍死亡者", "戸籍職権消除者", "戸籍その他消除者", "戸籍除籍者", "外国人", "外国人（消除者）", "登録外国人", "外国人転出予定者", "外国人転出確定者", "外国人死亡者", "外国人職権消除者", "外国人その他消除者", "市外外国人", "外国人死亡所有", "未登録外国人", "不明外国人", "学特等", "学特等（消除者）", "改徐票者", "喪失者", "住民喪失後住登外", "水道宛名非連動"}

    Private 全所有者数 As Integer = 0
    Private 全死亡人数 As Integer = 0
    Private 全複数人数 As Integer = 0
    Private 全その他人数 As Integer = 0

    Public Sub New()
        MyBase.New(True, True, "相続未登記等農地調査集計", "相続未登記等農地調査集計")

        Try
            Me.Controls.Add(SContainer)
            SContainer.Panel1.Controls.Add(TSContainer1)
            TSContainer1.TopToolStripPanel.Controls.Add(TStrip1)
            TStrip1.Items.AddRange({TSLabel1, New ToolStripSeparator, TSBtnExcel1})
            TSContainer1.ContentPanel.Controls.Add(mvarGrid1)
            SContainer.Panel2.Controls.Add(TSContainer2)
            TSContainer2.TopToolStripPanel.Controls.Add(TStrip2)
            TStrip2.Items.AddRange({TSLabel2, New ToolStripSeparator, TSBtnExcel2})
            TSContainer2.ContentPanel.Controls.Add(mvarGrid2)

            SContainer.Dock = DockStyle.Fill
            SContainer.Orientation = Orientation.Horizontal
            SContainer.SplitterDistance = 35

            TSContainer1.Dock = DockStyle.Fill
            mvarGrid1.Dock = DockStyle.Fill
            mvarGrid1.AllowUserToAddRows = False

            TSContainer2.Dock = DockStyle.Fill
            mvarGrid2.Dock = DockStyle.Fill
            mvarGrid2.AllowUserToAddRows = False

            TBL重複所有者 = New DataTable
            TBL重複所有者.Columns.Add("個人ID", GetType(Integer))
            TBL重複所有者.PrimaryKey = New DataColumn() {TBL重複所有者.Columns("個人ID")}

            SubMaster読み込み()
            SubFormat1処理()
            SubFormat2処理()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub SubMaster読み込み()
        TBL相続未登記Info = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT IIf([農振法区分]=0 Or [農振法区分] Is Null,IIf([農業振興地域]=0,2,IIf([農業振興地域]=2,3,[農業振興地域])),[農振法区分]) AS 農振法, IIf(InStr([氏名],'外')>0,IIf(Right([氏名],1)='名','共有',[V_住民区分].[名称]),[V_住民区分].[名称]) AS 名称区分, V_現況地目.名称 AS 地目, Sum([D:農地Info].登記簿面積) AS 登記簿面積の合計, Sum([D:農地Info].田面積) AS 田面積の合計, Sum([D:農地Info].畑面積) AS 畑面積の合計, Count([D:農地Info].ID) AS 筆数, [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].生年月日 " _
                                                              & "FROM (([D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) INNER JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID " _
                                                              & "GROUP BY IIf([農振法区分]=0 Or [農振法区分] Is Null,IIf([農業振興地域]=0,2,IIf([農業振興地域]=2,3,[農業振興地域])),[農振法区分]), IIf(InStr([氏名],'外')>0,IIf(Right([氏名],1)='名','共有',[V_住民区分].[名称]),[V_住民区分].[名称]), V_現況地目.名称, [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].生年月日 " _
                                                              & "HAVING (((IIf([農振法区分]=0 Or [農振法区分] Is Null,IIf([農業振興地域]=0,2,IIf([農業振興地域]=2,3,[農業振興地域])),[農振法区分]))>0 And (IIf([農振法区分]=0 Or [農振法区分] Is Null,IIf([農業振興地域]=0,2,IIf([農業振興地域]=2,3,[農業振興地域])),[農振法区分])) Is Not Null) AND ((V_現況地目.名称) In ('田','畑')));")

        TBL所有者Info = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT IIf(InStr([氏名],'外')>0,IIf(Right([氏名],1)='名','共有',[V_住民区分].[名称]),[V_住民区分].[名称]) AS 名称区分, V_現況地目.名称 AS 地目, Sum([D:農地Info].登記簿面積) AS 登記簿面積の合計, Sum([D:農地Info].田面積) AS 田面積の合計, Sum([D:農地Info].畑面積) AS 畑面積の合計, Count([D:農地Info].ID) AS 筆数, [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].生年月日 " _
                                                          & "FROM (([D:農地Info] INNER JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) INNER JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID " _
                                                          & "WHERE(((IIf([農振法区分] = 0 Or [農振法区分] Is Null, IIf([農業振興地域] = 0, 2, IIf([農業振興地域] = 2, 3, [農業振興地域])), [農振法区分])) > 0)) " _
                                                          & "GROUP BY IIf(InStr([氏名],'外')>0,IIf(Right([氏名],1)='名','共有',[V_住民区分].[名称]),[V_住民区分].[名称]), V_現況地目.名称, [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].生年月日 " _
                                                          & "HAVING (((V_現況地目.名称) In ('田','畑')));")
    End Sub

    ''' <summary>
    ''' 様式Ⅰの処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SubFormat1処理()
        With TBL相続Format1
            .Columns.Add("地目", GetType(String))
            .Columns.Add("農_全_筆数", GetType(Integer))
            .Columns.Add("農_死_筆数", GetType(Integer))
            .Columns.Add("農_複_筆数", GetType(Integer))
            .Columns.Add("農_他_筆数", GetType(Integer))
            .Columns.Add("農_合計_筆数", GetType(Integer), "[農_死_筆数]+[農_複_筆数]+[農_他_筆数]")
            .Columns.Add("農_合計_筆割合", GetType(Integer), "IIF([農_全_筆数]>0,[農_合計_筆数]/[農_全_筆数]*100,0)")
            .Columns.Add("農_全_面積", GetType(Decimal))
            .Columns.Add("農_死_面積", GetType(Decimal))
            .Columns.Add("農_複_面積", GetType(Decimal))
            .Columns.Add("農_他_面積", GetType(Decimal))
            .Columns.Add("農_合計_面積", GetType(Decimal), "[農_死_面積]+[農_複_面積]+[農_他_面積]")
            .Columns.Add("農_合計_面割合", GetType(Integer), "IIF([農_全_面積]>0,[農_合計_面積]/[農_全_面積]*100,0)")

            .Columns.Add("他_全_筆数", GetType(Integer))
            .Columns.Add("他_死_筆数", GetType(Integer))
            .Columns.Add("他_複_筆数", GetType(Integer))
            .Columns.Add("他_他_筆数", GetType(Integer))
            .Columns.Add("他_合計_筆数", GetType(Integer), "[他_死_筆数]+[他_複_筆数]+[他_他_筆数]")
            .Columns.Add("他_合計_筆割合", GetType(Integer), "IIF([他_全_筆数]>0,[他_合計_筆数]/[他_全_筆数]*100,0)")
            .Columns.Add("他_全_面積", GetType(Decimal))
            .Columns.Add("他_死_面積", GetType(Decimal))
            .Columns.Add("他_複_面積", GetType(Decimal))
            .Columns.Add("他_他_面積", GetType(Decimal))
            .Columns.Add("他_合計_面積", GetType(Decimal), "[他_死_面積]+[他_複_面積]+[他_他_面積]")
            .Columns.Add("他_合計_面割合", GetType(Integer), "IIF([他_全_面積]>0,[他_合計_面積]/[他_全_面積]*100,0)")

            .Columns.Add("外_全_筆数", GetType(Integer))
            .Columns.Add("外_死_筆数", GetType(Integer))
            .Columns.Add("外_複_筆数", GetType(Integer))
            .Columns.Add("外_他_筆数", GetType(Integer))
            .Columns.Add("外_合計_筆数", GetType(Integer), "[外_死_筆数]+[外_複_筆数]+[外_他_筆数]")
            .Columns.Add("外_合計_筆割合", GetType(Integer), "IIF([外_全_筆数]>0,[外_合計_筆数]/[外_全_筆数]*100,0)")
            .Columns.Add("外_全_面積", GetType(Decimal))
            .Columns.Add("外_死_面積", GetType(Decimal))
            .Columns.Add("外_複_面積", GetType(Decimal))
            .Columns.Add("外_他_面積", GetType(Decimal))
            .Columns.Add("外_合計_面積", GetType(Decimal), "[外_死_面積]+[外_複_面積]+[外_他_面積]")
            .Columns.Add("外_合計_面割合", GetType(Integer), "IIF([外_全_面積]>0,[外_合計_面積]/[外_全_面積]*100,0)")

            .Columns.Add("合計_全_所有者", GetType(Integer))
            .Columns.Add("合計_死_所有者", GetType(Integer))
            .Columns.Add("合計_複_所有者", GetType(Integer))
            .Columns.Add("合計_他_所有者", GetType(Integer))
            .Columns.Add("合計_合計_所有者", GetType(Integer), "[合計_死_所有者]+[合計_複_所有者]+[合計_他_所有者]")
            .Columns.Add("合計_合計_所割合", GetType(Integer), "IIF([合計_全_所有者]>0,[合計_合計_所有者]/[合計_全_所有者]*100,0)")
            .Columns.Add("合計_全_筆数", GetType(Integer), "[農_全_筆数]+[他_全_筆数]+[外_全_筆数]")
            .Columns.Add("合計_死_筆数", GetType(Integer), "[農_死_筆数]+[他_死_筆数]+[外_死_筆数]")
            .Columns.Add("合計_複_筆数", GetType(Integer), "[農_複_筆数]+[他_複_筆数]+[外_複_筆数]")
            .Columns.Add("合計_他_筆数", GetType(Integer), "[農_他_筆数]+[他_他_筆数]+[外_他_筆数]")
            .Columns.Add("合計_合計_筆数", GetType(Integer), "[合計_死_筆数]+[合計_複_筆数]+[合計_他_筆数]")
            .Columns.Add("合計_合計_筆割合", GetType(Integer), "IIF([合計_全_筆数]>0,[合計_合計_筆数]/[合計_全_筆数]*100,0)")
            .Columns.Add("合計_全_面積", GetType(Decimal), "[農_全_面積]+[他_全_面積]+[外_全_面積]")
            .Columns.Add("合計_死_面積", GetType(Decimal), "[農_死_面積]+[他_死_面積]+[外_死_面積]")
            .Columns.Add("合計_複_面積", GetType(Decimal), "[農_複_面積]+[他_複_面積]+[外_複_面積]")
            .Columns.Add("合計_他_面積", GetType(Decimal), "[農_他_面積]+[他_他_面積]+[外_他_面積]")
            .Columns.Add("合計_合計_面積", GetType(Decimal), "[合計_死_面積]+[合計_複_面積]+[合計_他_面積]")
            .Columns.Add("合計_合計_面割合", GetType(Integer), "IIF([合計_全_面積]>0,[合計_合計_面積]/[合計_全_面積]*100,0)")
        End With

        SetFormat1()
    End Sub
    Private Sub SetFormat1()
        Dim pRow1 As DataRow = TBL相続Format1.NewRow
        pRow1.Item("地目") = "田"
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[農振法]=1 And [地目]='田'", "", DataViewRowState.CurrentRows)
            Format1共通筆面積(pRow1, pRowV, 0)
        Next
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[農振法]=2 And [地目]='田'", "", DataViewRowState.CurrentRows)
            Format1共通筆面積(pRow1, pRowV, 1)
        Next
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[農振法]=3 And [地目]='田'", "", DataViewRowState.CurrentRows)
            Format1共通筆面積(pRow1, pRowV, 2)
        Next
        For Each pRowO As DataRowView In New DataView(TBL所有者Info, "[地目]='田'", "", DataViewRowState.CurrentRows)
            Format1共通所有者(pRow1, pRowO)
        Next
        AreaConv1(pRow1)
        TBL相続Format1.Rows.Add(pRow1)

        pRow1 = TBL相続Format1.NewRow
        pRow1.Item("地目") = "畑"
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[農振法]=1 And [地目]='畑'", "", DataViewRowState.CurrentRows)
            Format1共通筆面積(pRow1, pRowV, 0)
        Next
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[農振法]=2 And [地目]='畑'", "", DataViewRowState.CurrentRows)
            Format1共通筆面積(pRow1, pRowV, 1)
        Next
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[農振法]=3 And [地目]='畑'", "", DataViewRowState.CurrentRows)
            Format1共通筆面積(pRow1, pRowV, 2)
        Next
        For Each pRowO As DataRowView In New DataView(TBL所有者Info, "[地目]='畑'", "", DataViewRowState.CurrentRows)
            Format1共通所有者(pRow1, pRowO)
        Next
        AreaConv1(pRow1)
        TBL相続Format1.Rows.Add(pRow1)

        TBL重複所有者.Clear()

        pRow1 = TBL相続Format1.NewRow
        pRow1.Item("地目") = "計"
        For Each pRowT As DataRowView In New DataView(TBL相続Format1, "", "", DataViewRowState.CurrentRows)
            Format1合計処理(pRow1, pRowT)
        Next
        TBL相続Format1.Rows.Add(pRow1)

        mvarGrid1.SetDataView(TBL相続Format1, "", "")
    End Sub
    Private Sub Format1共通筆面積(ByRef pRow1 As DataRow, ByRef pRowV As DataRowView, ByVal pKey As Integer)
        Dim 農振区分 As String = ""
        Select Case pKey
            Case 0 : 農振区分 = "農"
            Case 1 : 農振区分 = "他"
            Case 2 : 農振区分 = "外"
            Case Else
        End Select

        If 0 <= Array.IndexOf(Ar死亡, pRowV.Item("名称区分")) Then
            pRow1.Item(農振区分 & "_死_筆数") = Val(pRow1.Item(農振区分 & "_死_筆数").ToString) + Val(pRowV.Item("筆数").ToString)
            pRow1.Item(農振区分 & "_全_筆数") = Val(pRow1.Item(農振区分 & "_全_筆数").ToString) + Val(pRowV.Item("筆数").ToString)
            pRow1.Item(農振区分 & "_死_面積") = Val(pRow1.Item(農振区分 & "_死_面積").ToString) + Val(pRowV.Item("登記簿面積の合計").ToString)
            pRow1.Item(農振区分 & "_全_面積") = Val(pRow1.Item(農振区分 & "_全_面積").ToString) + Val(pRowV.Item("登記簿面積の合計").ToString)
        ElseIf 0 <= Array.IndexOf(Ar共有, pRowV.Item("名称区分")) Then
            pRow1.Item(農振区分 & "_複_筆数") = Val(pRow1.Item(農振区分 & "_複_筆数").ToString) + Val(pRowV.Item("筆数").ToString)
            pRow1.Item(農振区分 & "_全_筆数") = Val(pRow1.Item(農振区分 & "_全_筆数").ToString) + Val(pRowV.Item("筆数").ToString)
            pRow1.Item(農振区分 & "_複_面積") = Val(pRow1.Item(農振区分 & "_複_面積").ToString) + Val(pRowV.Item("登記簿面積の合計").ToString)
            pRow1.Item(農振区分 & "_全_面積") = Val(pRow1.Item(農振区分 & "_全_面積").ToString) + Val(pRowV.Item("登記簿面積の合計").ToString)
        ElseIf 0 <= Array.IndexOf(Arその他, pRowV.Item("名称区分")) Then
            pRow1.Item(農振区分 & "_他_筆数") = Val(pRow1.Item(農振区分 & "_他_筆数").ToString) + Val(pRowV.Item("筆数").ToString)
            pRow1.Item(農振区分 & "_全_筆数") = Val(pRow1.Item(農振区分 & "_全_筆数").ToString) + Val(pRowV.Item("筆数").ToString)
            pRow1.Item(農振区分 & "_他_面積") = Val(pRow1.Item(農振区分 & "_他_面積").ToString) + Val(pRowV.Item("登記簿面積の合計").ToString)
            pRow1.Item(農振区分 & "_全_面積") = Val(pRow1.Item(農振区分 & "_全_面積").ToString) + Val(pRowV.Item("登記簿面積の合計").ToString)
        Else    'Case "住登者", "記録住民", "登記名義人", "市内住民", "法人", "財務", "医療", "医療機関", "住民票登載者"
            pRow1.Item(農振区分 & "_全_筆数") = Val(pRow1.Item(農振区分 & "_全_筆数").ToString) + Val(pRowV.Item("筆数").ToString)
            pRow1.Item(農振区分 & "_全_面積") = Val(pRow1.Item(農振区分 & "_全_面積").ToString) + Val(pRowV.Item("登記簿面積の合計").ToString)
        End If
    End Sub
    Private Sub Format1共通所有者(ByRef pRow1 As DataRow, ByRef pRowO As DataRowView)
        Dim FindRow As DataRow = Nothing
        If TBL重複所有者.Rows.Count > 0 Then
            FindRow = TBL重複所有者.Rows.Find(pRowO.Item("ID"))
        End If

        If FindRow IsNot Nothing Then
        Else
            Dim pRow As DataRow = TBL重複所有者.NewRow()
            pRow.Item("個人ID") = pRowO.Item("ID")
            TBL重複所有者.Rows.Add(pRow)

            If 0 <= Array.IndexOf(Ar死亡, pRowO.Item("名称区分")) Then
                全死亡人数 += 1
                全所有者数 += 1
            ElseIf 0 <= Array.IndexOf(Ar共有, pRowO.Item("名称区分")) Then
                全複数人数 += 1
                全所有者数 += 1
            ElseIf 0 <= Array.IndexOf(Arその他, pRowO.Item("名称区分")) Then
                全その他人数 += 1
                全所有者数 += 1
            Else    'Case "住登者", "記録住民", "登記名義人", "市内住民", "法人", "財務", "医療", "医療機関", "住民票登載者"
                全所有者数 += 1
            End If
        End If
    End Sub
    Private Sub AreaConv1(ByRef pRow1 As DataRow)
        pRow1.Item("農_死_筆数") = Val(pRow1.Item("農_死_筆数").ToString)
        pRow1.Item("農_複_筆数") = Val(pRow1.Item("農_複_筆数").ToString)
        pRow1.Item("農_他_筆数") = Val(pRow1.Item("農_他_筆数").ToString)
        pRow1.Item("農_全_筆数") = Val(pRow1.Item("農_全_筆数").ToString)
        pRow1.Item("他_死_筆数") = Val(pRow1.Item("他_死_筆数").ToString)
        pRow1.Item("他_複_筆数") = Val(pRow1.Item("他_複_筆数").ToString)
        pRow1.Item("他_他_筆数") = Val(pRow1.Item("他_他_筆数").ToString)
        pRow1.Item("他_全_筆数") = Val(pRow1.Item("他_全_筆数").ToString)
        pRow1.Item("外_死_筆数") = Val(pRow1.Item("外_死_筆数").ToString)
        pRow1.Item("外_複_筆数") = Val(pRow1.Item("外_複_筆数").ToString)
        pRow1.Item("外_他_筆数") = Val(pRow1.Item("外_他_筆数").ToString)
        pRow1.Item("外_全_筆数") = Val(pRow1.Item("外_全_筆数").ToString)

        pRow1.Item("農_死_面積") = Math.Round(Val(pRow1.Item("農_死_面積").ToString) / 10000, 1)
        pRow1.Item("農_複_面積") = Math.Round(Val(pRow1.Item("農_複_面積").ToString) / 10000, 1)
        pRow1.Item("農_他_面積") = Math.Round(Val(pRow1.Item("農_他_面積").ToString) / 10000, 1)
        pRow1.Item("農_全_面積") = Math.Round(Val(pRow1.Item("農_全_面積").ToString) / 10000, 1)
        pRow1.Item("他_死_面積") = Math.Round(Val(pRow1.Item("他_死_面積").ToString) / 10000, 1)
        pRow1.Item("他_複_面積") = Math.Round(Val(pRow1.Item("他_複_面積").ToString) / 10000, 1)
        pRow1.Item("他_他_面積") = Math.Round(Val(pRow1.Item("他_他_面積").ToString) / 10000, 1)
        pRow1.Item("他_全_面積") = Math.Round(Val(pRow1.Item("他_全_面積").ToString) / 10000, 1)
        pRow1.Item("外_死_面積") = Math.Round(Val(pRow1.Item("外_死_面積").ToString) / 10000, 1)
        pRow1.Item("外_複_面積") = Math.Round(Val(pRow1.Item("外_複_面積").ToString) / 10000, 1)
        pRow1.Item("外_他_面積") = Math.Round(Val(pRow1.Item("外_他_面積").ToString) / 10000, 1)
        pRow1.Item("外_全_面積") = Math.Round(Val(pRow1.Item("外_全_面積").ToString) / 10000, 1)
    End Sub

    Private Sub Format1合計処理(ByRef pRow1 As DataRow, ByRef pRowT As DataRowView)
        pRow1.Item("農_全_筆数") = Val(pRow1.Item("農_全_筆数").ToString) + Val(pRowT.Item("農_全_筆数").ToString)
        pRow1.Item("農_死_筆数") = Val(pRow1.Item("農_死_筆数").ToString) + Val(pRowT.Item("農_死_筆数").ToString)
        pRow1.Item("農_複_筆数") = Val(pRow1.Item("農_複_筆数").ToString) + Val(pRowT.Item("農_複_筆数").ToString)
        pRow1.Item("農_他_筆数") = Val(pRow1.Item("農_他_筆数").ToString) + Val(pRowT.Item("農_他_筆数").ToString)
        pRow1.Item("農_全_面積") = Val(pRow1.Item("農_全_面積").ToString) + Val(pRowT.Item("農_全_面積").ToString)
        pRow1.Item("農_死_面積") = Val(pRow1.Item("農_死_面積").ToString) + Val(pRowT.Item("農_死_面積").ToString)
        pRow1.Item("農_複_面積") = Val(pRow1.Item("農_複_面積").ToString) + Val(pRowT.Item("農_複_面積").ToString)
        pRow1.Item("農_他_面積") = Val(pRow1.Item("農_他_面積").ToString) + Val(pRowT.Item("農_他_面積").ToString)

        pRow1.Item("他_全_筆数") = Val(pRow1.Item("他_全_筆数").ToString) + Val(pRowT.Item("他_全_筆数").ToString)
        pRow1.Item("他_死_筆数") = Val(pRow1.Item("他_死_筆数").ToString) + Val(pRowT.Item("他_死_筆数").ToString)
        pRow1.Item("他_複_筆数") = Val(pRow1.Item("他_複_筆数").ToString) + Val(pRowT.Item("他_複_筆数").ToString)
        pRow1.Item("他_他_筆数") = Val(pRow1.Item("他_他_筆数").ToString) + Val(pRowT.Item("他_他_筆数").ToString)
        pRow1.Item("他_全_面積") = Val(pRow1.Item("他_全_面積").ToString) + Val(pRowT.Item("他_全_面積").ToString)
        pRow1.Item("他_死_面積") = Val(pRow1.Item("他_死_面積").ToString) + Val(pRowT.Item("他_死_面積").ToString)
        pRow1.Item("他_複_面積") = Val(pRow1.Item("他_複_面積").ToString) + Val(pRowT.Item("他_複_面積").ToString)
        pRow1.Item("他_他_面積") = Val(pRow1.Item("他_他_面積").ToString) + Val(pRowT.Item("他_他_面積").ToString)

        pRow1.Item("外_全_筆数") = Val(pRow1.Item("外_全_筆数").ToString) + Val(pRowT.Item("外_全_筆数").ToString)
        pRow1.Item("外_死_筆数") = Val(pRow1.Item("外_死_筆数").ToString) + Val(pRowT.Item("外_死_筆数").ToString)
        pRow1.Item("外_複_筆数") = Val(pRow1.Item("外_複_筆数").ToString) + Val(pRowT.Item("外_複_筆数").ToString)
        pRow1.Item("外_他_筆数") = Val(pRow1.Item("外_他_筆数").ToString) + Val(pRowT.Item("外_他_筆数").ToString)
        pRow1.Item("外_全_面積") = Val(pRow1.Item("外_全_面積").ToString) + Val(pRowT.Item("外_全_面積").ToString)
        pRow1.Item("外_死_面積") = Val(pRow1.Item("外_死_面積").ToString) + Val(pRowT.Item("外_死_面積").ToString)
        pRow1.Item("外_複_面積") = Val(pRow1.Item("外_複_面積").ToString) + Val(pRowT.Item("外_複_面積").ToString)
        pRow1.Item("外_他_面積") = Val(pRow1.Item("外_他_面積").ToString) + Val(pRowT.Item("外_他_面積").ToString)

        pRow1.Item("合計_全_所有者") = 全所有者数
        pRow1.Item("合計_死_所有者") = 全死亡人数
        pRow1.Item("合計_複_所有者") = 全複数人数
        pRow1.Item("合計_他_所有者") = 全その他人数
    End Sub

    ''' <summary>
    ''' 様式Ⅱの処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SubFormat2処理()
        With TBL相続Format2
            .Columns.Add("地目", GetType(String))
            .Columns.Add("全_所有者", GetType(String))
            .Columns.Add("未20年_所有者", GetType(String))
            .Columns.Add("超20年_所有者", GetType(String))
            .Columns.Add("不明_所有者", GetType(String))

            .Columns.Add("全_筆数", GetType(String))
            .Columns.Add("未20年_筆数", GetType(String))
            .Columns.Add("超20年_筆数", GetType(String))
            .Columns.Add("不明_筆数", GetType(String))

            .Columns.Add("全_面積", GetType(String))
            .Columns.Add("未20年_面積", GetType(String))
            .Columns.Add("超20年_面積", GetType(String))
            .Columns.Add("不明_面積", GetType(String))
        End With

        SetFormat2()
    End Sub
    Private Sub SetFormat2()
        Dim pRow2 As DataRow = TBL相続Format2.NewRow
        pRow2.Item("地目") = "田"
        Visible割合(pRow2, 0)
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[地目]='田'", "", DataViewRowState.CurrentRows)
            Format2共通筆面積(pRow2, pRowV)
        Next
        AreaConv2(pRow2)
        Visible割合(pRow2, 1)
        Visible割合(pRow2, 2)
        TBL相続Format2.Rows.Add(pRow2)

        pRow2 = TBL相続Format2.NewRow
        pRow2.Item("地目") = "畑"
        Visible割合(pRow2, 0)
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "[地目]='畑'", "", DataViewRowState.CurrentRows)
            Format2共通筆面積(pRow2, pRowV)
        Next
        AreaConv2(pRow2)
        Visible割合(pRow2, 1)
        Visible割合(pRow2, 2)
        TBL相続Format2.Rows.Add(pRow2)

        TBL重複所有者.Clear()

        pRow2 = TBL相続Format2.NewRow
        pRow2.Item("地目") = "計"
        For Each pRowO As DataRowView In New DataView(TBL所有者Info, "", "", DataViewRowState.CurrentRows)
            Format2共通所有者(pRow2, pRowO)
        Next
        Visible割合(pRow2, 0)
        For Each pRowV As DataRowView In New DataView(TBL相続未登記Info, "", "", DataViewRowState.CurrentRows)
            Format2共通筆面積(pRow2, pRowV)
        Next
        AreaConv2(pRow2)
        Visible割合(pRow2, 1)
        Visible割合(pRow2, 2)
        TBL相続Format2.Rows.Add(pRow2)

        mvarGrid2.SetDataView(TBL相続Format2, "", "")
    End Sub
    Private Sub Format2共通所有者(ByRef pRow2 As DataRow, ByRef pRowO As DataRowView)
        Dim FindRow As DataRow = Nothing
        If TBL重複所有者.Rows.Count > 0 Then
            FindRow = TBL重複所有者.Rows.Find(pRowO.Item("ID"))
        End If

        If FindRow IsNot Nothing Then
        Else
            Dim pRow As DataRow = TBL重複所有者.NewRow()
            pRow.Item("個人ID") = pRowO.Item("ID")
            TBL重複所有者.Rows.Add(pRow)


            If 0 <= Array.IndexOf(Ar死亡, pRowO.Item("名称区分")) Then
                pRow2.Item("全_所有者") = Val(pRow2.Item("全_所有者").ToString) + 1

                If IsDBNull(pRowO.Item("生年月日")) = False Then
                    pRow2.Item("不明_所有者") = Val(pRow2.Item("不明_所有者").ToString) + 1
                Else
                    pRow2.Item("不明_所有者") = Val(pRow2.Item("不明_所有者").ToString) + 1
                End If
            End If
        End If
    End Sub
    Private Sub Format2共通筆面積(ByRef pRow2 As DataRow, ByRef pRowV As DataRowView)


        If 0 <= Array.IndexOf(Ar死亡, pRowV.Item("名称区分")) Then
            pRow2.Item("全_筆数") = Val(pRow2.Item("全_筆数").ToString) + pRowV.Item("筆数")
            pRow2.Item("全_面積") = Val(pRow2.Item("全_面積").ToString) + pRowV.Item("登記簿面積の合計")

            pRow2.Item("不明_筆数") = Val(pRow2.Item("不明_筆数").ToString) + pRowV.Item("筆数")
            pRow2.Item("不明_面積") = Val(pRow2.Item("不明_面積").ToString) + pRowV.Item("登記簿面積の合計")
        End If
    End Sub
    Private Sub AreaConv2(ByRef pRow2 As DataRow)
        With pRow2
            .Item("全_面積") = Math.Round(Val(.Item("全_面積").ToString) / 10000, 1)
            .Item("超20年_面積") = Math.Round(Val(.Item("超20年_面積").ToString) / 10000, 1)
            .Item("未20年_面積") = Math.Round(Val(.Item("未20年_面積").ToString) / 10000, 1)
            .Item("不明_面積") = Math.Round(Val(.Item("不明_面積").ToString) / 10000, 1)
        End With
    End Sub
    Private Sub Visible割合(ByRef pRow2 As DataRow, ByVal pOption As Integer)
        Dim ConvType As String = ""
        Select Case pOption
            Case 0 : ConvType = "所有者"
            Case 1 : ConvType = "筆数"
            Case 2 : ConvType = "面積"
        End Select

        If Val(pRow2.Item("未20年_" & ConvType).ToString) > 0 Then
            pRow2.Item("未20年_" & ConvType) = pRow2.Item("未20年_" & ConvType) & "(" & Math.Round(pRow2.Item("未20年_" & ConvType) / pRow2.Item("全_" & ConvType) * 100) & "%)"
        End If
        If Val(pRow2.Item("超20年_" & ConvType).ToString) > 0 Then
            pRow2.Item("超20年_" & ConvType) = pRow2.Item("超20年_" & ConvType) & "(" & Math.Round(pRow2.Item("超20年_" & ConvType) / pRow2.Item("全_" & ConvType) * 100) & "%)"
        End If
        If Val(pRow2.Item("不明_" & ConvType).ToString) > 0 Then
            pRow2.Item("不明_" & ConvType) = pRow2.Item("不明_" & ConvType) & "(" & Math.Round(pRow2.Item("不明_" & ConvType) / pRow2.Item("全_" & ConvType) * 100) & "%)"
        End If
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

    Private Sub TSBtnExcel1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TSBtnExcel1.Click
        mvarGrid1.ToExcel()
    End Sub

    Private Sub TSBtnExcel2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TSBtnExcel2.Click
        mvarGrid2.ToExcel()
    End Sub
End Class
