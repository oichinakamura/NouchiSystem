'20160406霧島

Imports HimTools2012.Excel.XMLSS2003

Public Class CPrint総会資料作成農用地利用集積計画所有権
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        '複数申請人A

        SetNO(pSheet, True)
        申請者A(pSheet, pRow)
        申請者B(pSheet, pRow)
        調査委員(pSheet, pRow)

        If Not IsDBNull(pRow.Item("公告年月日")) Then pSheet.ValueReplace("{公告年月日D}", 和暦Format(pRow.Item("公告年月日"), "gyy.M.d")) Else pSheet.ValueReplace("{公告年月日D}", "")
        pSheet.ValueReplace("{対価}", pRow.Item("小作料").ToString & pRow.Item("小作料単位").ToString)
        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
        pSheet.ValueReplace("{契約期間}", "")
        pSheet.ValueReplace("{同意書}", "")

        複数土地設定(pSheet, pRow, Nothing)

        pSheet.ValueReplace("{申請者Ｂ労力}", "")
        Select Case pRow.Item("法令")
            Case 30, 32
                pSheet.ValueReplace("{契約内容}", "所有権移転")
            Case 31, 33

                Select Case Val(pRow.Item("権利種類"))
                    Case 1 : pSheet.ValueReplace("{契約内容}", "賃借権")
                    Case 2 : pSheet.ValueReplace("{契約内容}", "使用貸借権")
                    Case Else
                        pSheet.ValueReplace("{契約内容}", "その他")
                End Select
        End Select

        pSheet.ValueReplace("{図頁}", "")
    End Sub

    Public Sub Set総括Data(ByRef pSheet As XMLSSWorkSheet, ByVal pData As dt総括表)
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

    Public Sub Set総括表(ByRef pSheet As XMLSSWorkSheet, ByVal pTab As 申請Page, ByVal s処理名称 As String, ByVal pDataCreater As C総会資料Data作成)
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
