'20160406霧島

Imports HimTools2012.Excel.XMLSS2003

Public Class CPrint総会資料作成農用地利用集積計画貸借
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        Try
            SetNO(pSheet, True)
            申請者A(pSheet, pRow)
            調査委員(pSheet, pRow)

            Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))

            If Not n中間管理機構ID = 0 AndAlso Val(pRow.Item("経由法人ID").ToString) = n中間管理機構ID AndAlso SysAD.市町村.市町村名 <> "伊佐市" AndAlso SysAD.市町村.市町村名 <> "日置市" Then
                Dim p経由 As DataRow = App農地基本台帳.TBL個人.FindRowByID(n中間管理機構ID)
                pSheet.ValueReplace("{申請者Ｂ氏名}", p経由.Item("氏名").ToString)
                pSheet.ValueReplace("{申請者Ｂ住所}", p経由.Item("住所").ToString)
                pSheet.ValueReplace("{申請者Ｂ職業}", "")
                pSheet.ValueReplace("{申請者Ｂ年齢}", "")
                pSheet.ValueReplace("{申請者Ｂ年齢2}", "")
                pSheet.ValueReplace("{申請者Ｂ経営面積}", "")
                pSheet.ValueReplace("{申請者Ｂ申請理由}", "")
                pSheet.ValueReplace("{申請者Ｂ集落名}", "")
                pSheet.ValueReplace("{申請者Ｂ労力}", "")
                pSheet.ValueReplace("{申請者Ｂ認定}", "")
            Else
                Dim p受け手 As CObj個人 = ObjectMan.GetObject("個人." & pRow.Item("申請者B"))
                Select Case Val(p受け手.Row.Body.Item("農業改善計画認定").ToString)
                    Case 1 : pSheet.ValueReplace("{申請者Ｂ認定}", "認")
                    Case 2 : pSheet.ValueReplace("{申請者Ｂ認定}", "担")
                    Case 3 : pSheet.ValueReplace("{申請者Ｂ認定}", "法")
                    Case 4 : pSheet.ValueReplace("{申請者Ｂ認定}", "認・担")
                    Case 5 : pSheet.ValueReplace("{申請者Ｂ認定}", "認・法")
                    Case 6 : pSheet.ValueReplace("{申請者Ｂ認定}", "認新")
                    Case Else
                        pSheet.ValueReplace("{申請者Ｂ認定}", "")
                End Select

                pSheet.ValueReplace("{申請者Ｂ氏名}", pRow.Item("氏名B").ToString)
                pSheet.ValueReplace("{申請者Ｂ住所}", pRow.Item("住所B").ToString)
                pSheet.ValueReplace("{申請者Ｂ職業}", pRow.Item("職業B").ToString)
                pSheet.ValueReplace("{申請者Ｂ年齢}", IIf(Val(pRow.Item("年齢B").ToString) = 0, "-", pRow.Item("年齢B").ToString))
                pSheet.ValueReplace("{申請者Ｂ年齢2}", IIf(Val(pRow.Item("年齢B").ToString) = 0, "-", pRow.Item("年齢B").ToString) & "歳")
                pSheet.ValueReplace("{申請者Ｂ経営面積}", Val(pRow.Item("経営面積B").ToString).ToString("#,##0"))
                pSheet.ValueReplace("{申請者Ｂ経営面積ha}", Format(Val(pRow.Item("経営面積B").ToString) / 10000, "0.0"))    '少数第1位まで表示
                pSheet.ValueReplace("{申請者Ｂ申請理由}", pRow.Item("申請理由B").ToString)
                pSheet.ValueReplace("{申請者Ｂ集落名}", pRow.Item("集落B").ToString)

                pSheet.ValueReplace("{申請者Ｂ労力}", pRow.Item("稼動人数").ToString)
                pSheet.ValueReplace("{申請者Ｂ世帯員数}", pRow.Item("世帯員数B").ToString)
                If p受け手.農業改善計画認定 = enum農業改善計画認定.認定農業者 OrElse p受け手.農業改善計画認定 = enum農業改善計画認定.認定農業者_担い手農家 OrElse p受け手.農業改善計画認定 = enum農業改善計画認定.認定農業者_農業生産法人 Then
                    pSheet.ValueReplace("{認定農業者}", "○")
                Else
                    pSheet.ValueReplace("{認定農業者}", "")
                End If

                If Not n中間管理機構ID = 0 AndAlso Val(pRow.Item("経由法人ID").ToString) = n中間管理機構ID Then
                    Dim p経由 As DataRow = App農地基本台帳.TBL個人.FindRowByID(n中間管理機構ID)
                    pSheet.ValueReplace("{申請者C氏名}", p経由.Item("氏名").ToString)
                    pSheet.ValueReplace("{申請者C住所}", p経由.Item("住所").ToString)
                Else
                    pSheet.ValueReplace("{申請者C氏名}", "")
                    pSheet.ValueReplace("{申請者C住所}", "")
                End If

            End If

            pSheet.ValueReplace("{農用地}", "")
            pSheet.ValueReplace("{土地改良}", "")

            pSheet.ValueReplace("{受付年月日}", 和暦Format(pRow.Item("受付年月日"), "gyy.M.d"))
            If IsDBNull(pRow.Item("公告年月日")) Then
                pSheet.ValueReplace("{公告年月日D}", "")
            Else
                pSheet.ValueReplace("{公告年月日D}", 和暦Format(pRow.Item("公告年月日"), "gyy.M.d"))
            End If
            pSheet.ValueReplace("{利用権内容}", pRow.Item("利用権内容").ToString)
            pSheet.ValueReplace("{作物名}", pRow.Item("利用権内容").ToString)

            Dim s備考 As String = ""

            pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
            pSheet.ValueReplace("{受付補助記号}", pRow.Item("受付補助記号").ToString + pRow.Item("受付番号").ToString)

            Dim p総括Item As dt総括表 = Nothing
            Dim b再設定 As Boolean = False
            Try
                Dim dt始期 As Object = pRow.Item("始期")
                Dim dt終期 As Object = pRow.Item("終期")
                Dim n期間年 As Integer = 999

                If IsDBNull(pRow.Item("期間")) OrElse pRow.Item("期間") = 0 Then
                    If Not IsDBNull(dt始期) AndAlso Not IsDBNull(dt終期) Then
                        n期間年 = DateDiff(DateInterval.Year, dt始期, dt終期)
                    End If
                Else
                    n期間年 = pRow.Item("期間")
                End If

                Dim n期間月 As Integer = 0
                Dim dt公告日 As DateTime
                If IsDBNull(pRow.Item("公告年月日")) Then
                    dt公告日 = dt始期
                Else
                    dt公告日 = pRow.Item("公告年月日")
                End If

                If 総括Data.ContainsKey(dt総括表.GetKeyStr(n期間年, n期間月, dt公告日, dt始期)) Then
                    p総括Item = 総括Data.Item(dt総括表.GetKeyStr(n期間年, n期間月, dt公告日, dt始期))
                Else
                    p総括Item = New dt総括表
                    p総括Item.n期間年 = n期間年
                    p総括Item.n期間月 = n期間月
                    p総括Item.dt公告日 = dt公告日
                    p総括Item.dt始期 = dt始期
                    p総括Item.dt終期 = dt終期
                    総括Data.Add(p総括Item.ToString, p総括Item)
                End If

                Select Case pRow.Item("法令")
                    Case 61
                        ' 「利用権設定」は同じ人をカウントしない
                        Set人数(p総括Item.s貸し手, pRow.Item("氏名A").ToString)
                        Set人数(p総括Item.s借り手, pRow.Item("氏名B").ToString)

                        If Not IsDBNull(pRow.Item("再設定")) AndAlso pRow.Item("再設定") = True Then
                            b再設定 = True
                            Set人数(p総括Item.s貸し手再, pRow.Item("氏名A").ToString)
                            Set人数(p総括Item.s借り手再, pRow.Item("氏名B").ToString)
                        End If
                End Select
            Catch ex As Exception
                MsgBox("利用権総会資料作成：" & ex.Message)
            End Try

            複数土地設定(pSheet, pRow, p総括Item, b再設定)
            If p総括Item IsNot Nothing Then
                p総括Item.n貸し件数 += 1
                p総括Item.n借り件数 += 1
                If b再設定 Then
                    p総括Item.n貸し件数再 += 1
                    p総括Item.n借り件数再 += 1
                End If
            End If

            貸借共通(pSheet, pRow)

            pSheet.ValueReplace("{図頁}", "")
            pSheet.ValueReplace("{同意書}", "")
        Catch ex As Exception
            Stop
        End Try

    End Sub

    Private Sub Set人数(ByRef pDic As Dictionary(Of String, String), ByRef pStr2 As String)
        If Not pDic.ContainsKey(pStr2) Then
            pDic.Add(pStr2, pStr2)
        End If
    End Sub

    Public Sub Set総括Data(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByVal pData As dt総括表)
        SetNO(pSheet, True)
        pSheet.ValueReplace("{市町村名}", SysAD.市町村.市町村名)
        pSheet.ValueReplace("{公告年月日}", 和暦Format(pData.dt公告日, "gyy/M/d"))

        pSheet.ValueReplace("{始期}", 和暦Format(pData.dt始期, "gyy.M.d"))
        If Not IsDBNull(pData.dt終期) Then
            pSheet.ValueReplace("{終期}", 和暦Format(pData.dt終期, "gyy.M.d"))
        Else
            pSheet.ValueReplace("{終期}", "-")
        End If

        '
        If pData.n期間年 = 999 Then
            pSheet.ValueReplace("{存続期間}", "永年")
        Else
            pSheet.ValueReplace("{存続期間}", pData.n期間年 & "年" & IIf(pData.n期間月 > 0, pData.n期間月 & "月", "") & "間")
        End If

        pSheet.ValueReplace("-1000", pData.n面積_田.ToString("#,##0"))
        pSheet.ValueReplace("-1001", pData.n面積_畑.ToString("#,##0"))
        pSheet.ValueReplace("-1002", pData.n面積_樹.ToString("#,##0"))
        pSheet.ValueReplace("-1003", pData.n面積_他.ToString("#,##0"))
        pSheet.ValueReplace("-1004", pData.総面積.ToString("#,##0"))
        pSheet.ValueReplace("-1005", pData.n面積_再.ToString("#,##0"))
        pSheet.ValueReplace("-1006", pData.s貸し手.Count.ToString("#,##0"))
        pSheet.ValueReplace("-1007", pData.s貸し手再.Count.ToString("#,##0"))
        pSheet.ValueReplace("-1008", pData.s借り手.Count.ToString("#,##0"))
        pSheet.ValueReplace("-1009", pData.s借り手再.Count.ToString("#,##0"))

        pSheet.ValueReplace("-2005", pData.n貸し件数.ToString("#,##0"))
        pSheet.ValueReplace("-2006", pData.n貸し件数再.ToString("#,##0"))
        pSheet.ValueReplace("-2007", pData.n借り件数.ToString("#,##0"))
        pSheet.ValueReplace("-2008", pData.n借り件数再.ToString("#,##0"))
    End Sub

    Public Sub Set総括表(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByVal pTab As 申請Page, ByVal s処理名称 As String, ByVal pDataCreater As C総会資料Data作成)
        If Not InStr(pSheet.Table.InnerXML, "{No}") > 0 Then
            Exit Sub
        End If

        Dim pXMLLoopRows As New List(Of XMLSSRow)
        Dim InsetRow As Integer = 0


        Dim MergeDown As Integer = 0
        For Each pXRow As XMLSSRow In pSheet.Table.Rows.Items
            If InStr(pXRow.InnerXML, "{No}") > 0 Then
                For Each pCell As XMLSSCell In pXRow.Cells.Items
                    If pCell.MergeDown IsNot Nothing Then
                        If MergeDown < pCell.MergeDown.Value Then
                            MergeDown = pCell.MergeDown.Value
                        End If
                    End If
                Next

                pXMLLoopRows.Add(pXRow)
                If MergeDown = 0 Then
                    InsetRow += 1
                    Exit For
                End If
            ElseIf MergeDown > 0 Then
                pXMLLoopRows.Add(pXRow)
                MergeDown -= 1
                If MergeDown = 0 Then
                    InsetRow += 1
                    Exit For
                End If
            End If
            InsetRow += 1
        Next

        pDataCreater.Maximum = 総括Data.Count
        nLoop = -1
        Dim pList As New List(Of String)

        For Each sKey As String In 総括Data.Keys
            pList.Add(sKey)
        Next
        pList.Sort()

        Dim n面積_田 As Decimal = 0
        Dim n面積_畑 As Decimal = 0
        Dim n面積_樹 As Decimal = 0
        Dim n面積_他 As Decimal = 0
        Dim n総面積 As Decimal = 0
        Dim n面積_再 As Decimal = 0
        Dim n貸し手 As Integer = 0
        Dim n貸し手再 As Integer = 0
        Dim n受け手 As Integer = 0
        Dim n受け手再 As Integer = 0

        Dim n貸し件数 As Integer = 0
        Dim n貸し件数再 As Integer = 0
        Dim n受け件数 As Integer = 0
        Dim n受け件数再 As Integer = 0

        LoopRows = New XMLLoopRows(pSheet)
        For Each sKey As String In pList
            Dim pLine As dt総括表 = 総括Data(sKey)

            If nLoop = -1 Then

            Else
                For Each pXRow As XMLSSRow In LoopRows
                    Dim pCopyRow = pXRow.CopyRow

                    pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                    LoopRows.InsetRow += 1
                Next
            End If

            nLoop += 1
            pDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & 総括Data.Count & ")"
            pDataCreater.Value = nLoop

            n面積_田 += pLine.n面積_田
            n面積_畑 += pLine.n面積_畑
            n面積_樹 += pLine.n面積_樹
            n面積_他 += pLine.n面積_他
            n総面積 += pLine.総面積
            n面積_再 += pLine.n面積_再

            n貸し手 += pLine.s貸し手.Count
            n貸し手再 += pLine.s貸し手再.Count
            n受け手 += pLine.s借り手.Count
            n受け手再 += pLine.s借り手再.Count

            n貸し件数 += pLine.n貸し件数
            n貸し件数再 += pLine.n貸し件数再
            n受け件数 += pLine.n借り件数
            n受け件数再 += pLine.n借り件数再
            Me.Set総括Data(pSheet, pLine)
        Next

        With pSheet
            .ValueReplace("{総括田面積}", n面積_田.ToString("#,##0"))
            .ValueReplace("{総括畑面積}", n面積_畑.ToString("#,##0"))
            .ValueReplace("{総括樹面積}", n面積_樹.ToString("#,##0"))
            .ValueReplace("{総括他面積}", n面積_他.ToString("#,##0"))
            .ValueReplace("{総括計面積}", n総面積.ToString("#,##0"))
            .ValueReplace("{総括再面積}", n面積_再.ToString("#,##0"))

            .ValueReplace("{利用権出し手数}", n貸し手.ToString("#,##0"))
            .ValueReplace("{利用権出し手再数}", n貸し手再.ToString("#,##0"))
            .ValueReplace("{利用権受け手数}", n受け手.ToString("#,##0"))
            .ValueReplace("{利用権受け手再数}", n受け手再.ToString("#,##0"))
            .ValueReplace("{市町村名}", SysAD.市町村.市町村名)

            .ValueReplace("{利用権出し件数}", n貸し件数.ToString("#,##0"))
            .ValueReplace("{利用権出し件再数}", n貸し件数再.ToString("#,##0"))
            .ValueReplace("{利用権受け件数}", n受け件数.ToString("#,##0"))
            .ValueReplace("{利用権受け件再数}", n受け件数再.ToString("#,##0"))
        End With
    End Sub
End Class

