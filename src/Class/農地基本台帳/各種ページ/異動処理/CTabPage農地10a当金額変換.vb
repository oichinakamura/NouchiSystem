

Public Class CTabPage農地10a当金額変換
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvarTable As DataTable
    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView

    Public Sub New()
        MyBase.New(True, True, "農地10a当金額変換", "農地10a当金額変換")
        mvarTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_農地.ID,V_農地.土地所在, '賃貸借' AS 形態, V_農地.小作料, V_農地.小作料単位, 0.0/1.0 AS [計算結果10a賃借料],'' AS [変換後単位],[V_農地].[10a賃借料],'' AS [処理区分] FROM V_農地 WHERE (((V_農地.小作形態)=1) AND ((V_農地.自小作別)>0));")
        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView

        For Each pRow As DataRow In mvarTable.Rows
            Dim s小作料 As String = pRow.Item("小作料").ToString
            If Not IsDBNull(pRow.Item("小作料")) Then
                If IsNumeric(s小作料) Then
                Else
                    pRow.Item("処理区分") = "換算失敗"
                End If
            End If

            Select Case pRow.Item("小作料単位").ToString.ToLower
                Case "円/10a"
                    If IsDBNull(pRow.Item("10a賃借料")) OrElse pRow.Item("10a賃借料") = 0 Then
                        pRow.Item("計算結果10a賃借料") = Val(Replace(s小作料, ",", ""))
                        pRow.Item("変換後単位") = "円/10a"
                        pRow.Item("処理区分") = "換算成功"
                    ElseIf Not pRow.Item("10a賃借料") = Val(Replace(s小作料, ",", "")) Then
                        pRow.Item("計算結果10a賃借料") = Val(Replace(s小作料, ",", ""))
                        pRow.Item("変換後単位") = "円/10a"
                        pRow.Item("処理区分") = "換算成功"
                    ElseIf pRow.Item("10a賃借料") = Val(Replace(s小作料, ",", "")) Then
                        pRow.Item("計算結果10a賃借料") = Val(Replace(s小作料, ",", ""))
                        pRow.Item("変換後単位") = "円/10a"
                        pRow.Item("処理区分") = "換算済み"
                    End If
                Case "万円/10a", "万円/10a当"
                    If IsDBNull(pRow.Item("10a賃借料")) OrElse pRow.Item("10a賃借料") = 0 Then
                        pRow.Item("計算結果10a賃借料") = Val(Replace(s小作料, ",", "")) * 10000
                        pRow.Item("変換後単位") = "円/10a"
                        pRow.Item("処理区分") = "換算成功"
                    ElseIf Not pRow.Item("10a賃借料") = Val(Replace(s小作料, ",", "")) * 10000 Then
                        pRow.Item("計算結果10a賃借料") = Val(Replace(s小作料, ",", "")) * 10000
                        pRow.Item("変換後単位") = "円/10a"
                        pRow.Item("処理区分") = "換算成功"
                    ElseIf pRow.Item("10a賃借料") = Val(Replace(s小作料, ",", "")) * 10000 Then
                        pRow.Item("計算結果10a賃借料") = Val(Replace(s小作料, ",", "")) * 10000
                        pRow.Item("変換後単位") = "円/10a"
                        pRow.Item("処理区分") = "換算済み"
                    End If
                Case "円"
                    pRow.Item("処理区分") = "換算失敗"
                Case "俵"
                    pRow.Item("処理区分") = "換算失敗"
                Case ""
                    pRow.Item("処理区分") = "換算失敗"
                Case Else
                    pRow.Item("処理区分") = "換算失敗"
            End Select
        Next
        AddHandler Me.ToolStrip.Items.Add("換算結果を適用").Click, AddressOf Set換算金額
        AddHandler Me.ToolStrip.Items.Add("エクセルへ").Click, AddressOf mvarGrid.ToExcel
        AddHandler Me.ToolStrip.Items.Add("換算済みを非表示").Click, AddressOf Sub換算済みを非表示

        Me.ControlPanel.Add(mvarGrid)
        mvarGrid.SetDataView(mvarTable, "", "")
    End Sub

    Public Sub Sub換算済みを非表示()
        If mvarGrid.RowFilter.Length > 0 Then
            mvarGrid.RowFilter = ""
        Else
            mvarGrid.RowFilter = "[処理区分] <> '換算済み'"
        End If
    End Sub

    Public Sub Set換算金額()
        If MsgBox("成功した換算結果を10a当の賃借料に設定しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            With New C繰返クエリ("10a賃借料金額換算")
                .Params.Add("データテーブル", mvarTable)
                .Dialog.StartProc(True, True)

                If .Dialog._objException Is Nothing = False Then
                    If .Dialog._objException.Message = "Cancel" Then
                        MsgBox("処理を中止しました。　", , "処理中止")
                        Exit Sub
                    Else
                        'Throw objDlg._objException
                    End If
                Else
                    MsgBox("終了しました")
                End If
            End With
        End If
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

End Class

