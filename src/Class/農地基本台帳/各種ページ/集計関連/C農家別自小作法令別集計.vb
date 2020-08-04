

Public Class C農家別自小作法令別集計
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView

    Public Sub New()
        MyBase.New(True, True, "農家別自小作法令別集計", "農家別自小作法令別集計")
        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView
        Me.ControlPanel.Add(mvarGrid)
        mvarGrid.Createエクセル出力Ctrl(Me.ToolStrip)
        App農地基本台帳.ListColumnDesign.SetGridColumns(mvarGrid, "農家別自小作法令別集計")

        With New DataLoder
            .Dialog.StartProc(True, True)
            .TBL.Columns.Add("Key", GetType(String), "'農家.'+[世帯番号]")
            .TBL.Columns.Add("自作計", GetType(Decimal), "[自作田]+[自作畑]")
            .TBL.Columns.Add("農地法小作計", GetType(Decimal), "[農地法小作田]+[農地法小作畑]")
            .TBL.Columns.Add("利用権小作計", GetType(Decimal), "[利用権小作田]+[利用権小作畑]")
            .TBL.Columns.Add("経営面積田", GetType(Decimal), "[自作田]+[農地法小作田]+[利用権小作田]")
            .TBL.Columns.Add("経営面積畑", GetType(Decimal), "[自作畑]+[農地法小作畑]+[利用権小作畑]")
            .TBL.Columns.Add("経営面積", GetType(Decimal), "[経営面積田]+[経営面積畑]")

            mvarGrid.AutoGenerateColumns = False
            mvarGrid.SetDataView(.TBL, "", "")
        End With
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

    'Public Overrides Function ViewMenuDropDownOpening(ByRef pViewMenu As System.Windows.Forms.ToolStripMenuItem) As System.Windows.Forms.ToolStripMenuItem
    '    Return Nothing
    'End Function

    Private Class DataLoder
        Inherits HimTools2012.clsAccessor
        Public TBL As DataTable

        Public Sub New()
        End Sub

        Public Overrides Sub Execute()
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [{0}]=0 WHERE [{0}] Is Null", "小作地適用法")
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [{0}]=0 WHERE [{0}] Is Null", "田面積")
            SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [{0}]=0 WHERE [{0}] Is Null", "畑面積")

            'TBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:世帯Info].ID AS [世帯番号], [D:個人Info].氏名, [D:個人Info].住所, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=0,[V_農地].[田面積],0))) AS 自作田, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=0,[V_農地].[畑面積],0))) AS 自作畑, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=1,[V_農地].[田面積],0))) AS 農地法小作田, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=1,[V_農地].[畑面積],0))) AS 農地法小作畑, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=2,[V_農地].[田面積],0))) AS 利用権小作田, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=2,[V_農地].[畑面積],0))) AS 利用権小作畑 FROM (V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE (((V_農地.耕作世帯ID)<>0) AND ((V_農地.農地状況)<20)) GROUP BY [D:世帯Info].ID, [D:個人Info].氏名, [D:個人Info].住所;")
            TBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_行政区.名称 AS 集落別, IIf([農業改善計画認定]=1,'認定農業者',IIf([農業改善計画認定]=2,'担い手農家',IIf([農業改善計画認定]=3,'農業生産法人',IIf([農業改善計画認定]=4,'認定農業者＋担い手農家','なし')))) AS 認定区分別, [D:世帯Info].ID AS 世帯番号, [D:個人Info].氏名, [D:個人Info].住所, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=0,[V_農地].[田面積],0))) AS 自作田, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=0,[V_農地].[畑面積],0))) AS 自作畑, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=1,[V_農地].[田面積],0))) AS 農地法小作田, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=1,[V_農地].[畑面積],0))) AS 農地法小作畑, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=2,[V_農地].[田面積],0))) AS 利用権小作田, Sum(Int(IIf(IIf([自小作別]>0,1,0)*[小作地適用法]=2,[V_農地].[畑面積],0))) AS 利用権小作畑 FROM ((V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE (((V_農地.耕作世帯ID)<>0) AND ((V_農地.農地状況)<20)) GROUP BY V_行政区.名称, IIf([農業改善計画認定]=1,'認定農業者',IIf([農業改善計画認定]=2,'担い手農家',IIf([農業改善計画認定]=3,'農業生産法人',IIf([農業改善計画認定]=4,'認定農業者＋担い手農家','なし')))), [D:世帯Info].ID, [D:個人Info].氏名, [D:個人Info].住所, V_行政区.ID, [D:世帯Info].ID ORDER BY V_行政区.ID, [D:世帯Info].ID;")
        End Sub
    End Class

End Class

