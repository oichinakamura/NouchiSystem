'20160406霧島
Imports System.ComponentModel
Imports HimTools2012.Excel.XMLSS2003


Public Class dt総括表
    Public Enum enum地目
        田 = 1
        畑 = 2
        樹 = 3
        他 = 4
    End Enum

    Public dt公告日 As DateTime
    Public dt始期 As DateTime
    Public dt終期 As Object
    Public n期間年 As Integer = 0
    Public n期間月 As Integer = 0
    Public n面積_田 As Decimal = 0
    Public n面積_畑 As Decimal = 0
    Public n面積_樹 As Decimal = 0
    Public n面積_他 As Decimal = 0
    Public n面積_再 As Decimal = 0

    Public n貸し件数 As Integer = 0
    Public n貸し件数再 As Integer = 0
    Public n借り件数 As Integer = 0
    Public n借り件数再 As Integer = 0


    Public s貸し手 As New Dictionary(Of String, String)
    Public s貸し手再 As New Dictionary(Of String, String)
    Public s借り手 As New Dictionary(Of String, String)
    Public s借り手再 As New Dictionary(Of String, String)


    Public Shared Function GetKeyStr(ByVal p年間数 As Integer, p期間月 As Integer, p公告日 As DateTime, dt始期 As DateTime) As String
        Return String.Format("{0:D3}{1:D2}", p年間数, p期間月) & Strings.Format(p公告日, "yyyyMMdd") & Strings.Format(dt始期, "yyyyMMdd")
    End Function
    Public Overrides Function ToString() As String
        Return GetKeyStr(n期間年, n期間月, dt公告日, dt始期)
    End Function

    Public Sub Set面積(p地目 As enum地目, n面積 As Decimal, b再設定 As Boolean)
        Select Case p地目
            Case enum地目.田 : n面積_田 += n面積
            Case enum地目.畑 : n面積_畑 += n面積
            Case enum地目.樹 : n面積_樹 += n面積
            Case enum地目.他 : n面積_他 += n面積
        End Select

        If b再設定 Then
            n面積_再 += n面積
        End If


    End Sub

    Public Function 総面積() As Decimal
        Return Me.n面積_田 + Me.n面積_畑 + Me.n面積_樹 + Me.n面積_他
    End Function
End Class

Public MustInherit Class CPrint資料作成共通
    Protected nLoop As Integer = -1
    Public LoopRows As XMLLoopRows

    MustOverride Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView, Optional pObj As Object = Nothing)
    MustOverride Sub SetData(ByRef XMLSS As HimTools2012.Excel.XMLSS2003.CXMLSS2003, ByRef pTab As 申請Page, ByVal s処理名称 As String, ByRef pDataCreater As C総会資料Data作成)

    Public Overridable Function LoopSub(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal p作成 As CPrint資料作成共通, ByVal sFile As String, ParamArray pTabs() As Object) As Boolean
        Return False
    End Function


    Public Function SetNO(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet) As Integer

        pSheet.ValueReplace("{No}", (nLoop + 1).ToString)

        Return nLoop + 1
    End Function
End Class

Public MustInherit Class CPrint総会資料作成
    Inherits CPrint資料作成共通

    Protected mvar総合計 As New C筆明細と集計作成

    Protected 総括Data As New Dictionary(Of String, dt総括表)

    Public Sub New()

    End Sub

    Public Overloads Function SetNO(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByVal b行高調整 As Boolean) As Integer
        If b行高調整 Then
            pSheet.ValueReplace("{No}", "" & (nLoop + 1).ToString)
        Else
            pSheet.ValueReplace("{No}", (nLoop + 1).ToString)
        End If
        Return nLoop + 1
    End Function

    Public Sub Set複数行(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pTab As 申請Page, ByVal s処理名称 As String, ByRef pDataCreater As C総会資料Data作成, sWhere As String)
        If Not InStr(pSheet.Table.InnerXML, "{No}") > 0 Then
            Exit Sub
        End If
        LoopRows = New XMLLoopRows(pSheet)

        Dim pTable As DataTable = CType(pTab.List.DataSource, DataView).Table

        Dim pView As New DataView(pTable, CType(pTab.List.DataSource, DataView).RowFilter & " AND [選択]=True", "中間管理Flag,受付補助記号,受付番号", DataViewRowState.CurrentRows)

        pDataCreater.Maximum = pView.Count
        nLoop = -1

        For Each pRow As DataRowView In pView
            If pRow.Item("選択") Then

                If nLoop = -1 Then

                Else
                    For Each pXRow As XMLSSRow In LoopRows
                        Dim pCopyRow = pXRow.CopyRow

                        pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                        LoopRows.InsetRow += 1
                    Next
                End If

                nLoop += 1
                pDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
                pDataCreater.Value = nLoop

                Me.SetDataRow(pSheet, pRow)
            End If
        Next
    End Sub

    Protected Function 複数土地設定(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView, ByRef p総括Item As dt総括表, Optional b再設定 As Boolean = False) As C筆明細と集計作成
        'Try
        Dim sNList As String = pRow.Item("農地リスト").ToString
        Dim s登記地目 As String = ""
        Dim s現況地目 As String = ""
        Dim s農委地目 As String = ""
        Dim s自小作別 As String = ""
        Dim s持分 As String = ""
        Dim s農振区分 As String = ""

        Dim p案件内集計 As New C筆明細と集計作成

        Dim nCount As Integer = 0
        Dim sNID As String = ""

        Dim Ar As String() = Split(sNList, ";")
        For Each sKey As String In Ar

            If sKey.StartsWith("農地.") Then
                sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(3)
            ElseIf sKey.StartsWith("転用農地.") Then
                sNID &= IIf(sNID.Length > 0, ",", "") & sKey.Substring(5)
            Else
                Stop
            End If

        Next

        App農地基本台帳.TBL農地.FindRowBySQL("[ID] In (" & sNID & ")")
        Dim s管理者内訳 As String = ""
        Dim pView As New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
        For Each pRowV As DataRowView In pView
            If InStr("," & sNID & ",", "," & pRowV.Item("ID") & ",") Then
                sNID = Replace("," & sNID & ",", "," & pRowV.Item("ID") & ",", ",")
            End If
            nCount += 1


            s登記地目 = s登記地目 & p案件内集計.R & pRowV.Item("登記簿地目名").ToString
            s現況地目 = s現況地目 & p案件内集計.R & pRowV.Item("現況地目名").ToString
            s農委地目 = s農委地目 & p案件内集計.R & pRowV.Item("農委地目名").ToString
            s自小作別 = s自小作別 & p案件内集計.R & IIf(Val(pRowV.Item("自小作別").ToString) = 0, "自", "小")
            s持分 = s持分 & p案件内集計.R & IIf(Val(pRowV.Item("共有持分分子").ToString) > 0 And Val(pRowV.Item("共有持分分母").ToString) > 0, Val(pRowV.Item("共有持分分子").ToString) & "/" & Val(pRowV.Item("共有持分分母").ToString), "")
            s農振区分 = s農振区分 & p案件内集計.R & IIf(Val(pRowV.Item("農振法区分").ToString) = 0,
                                                    IIf(Val(pRowV.Item("農業振興地域").ToString) = 0, "他", IIf(Val(pRowV.Item("農業振興地域").ToString) = 1, "内", IIf(Val(pRowV.Item("農業振興地域").ToString) = 2, "外", "-"))),
                                                    IIf(Val(pRowV.Item("農振法区分").ToString) = 1, "内", IIf(Val(pRowV.Item("農振法区分").ToString) = 2, "他", IIf(Val(pRowV.Item("農振法区分").ToString) = 3, "外", "-"))))
            p案件内集計.Set筆情報(pRowV.Row, pRow.Row, p総括Item, b再設定)

            If Not IsDBNull(pRowV.Item("管理者ID")) Then
                Select Case Val(pRowV.Item("農地所有内訳").ToString)
                    Case 0, 2 : s管理者内訳 = "代理人"
                    Case 1 : s管理者内訳 = "管理人"
                End Select
            End If
        Next

        sNID = Replace(sNID, ",,", ",")
        Do Until Not sNID.StartsWith(",") AndAlso Not sNID.EndsWith(",")
            If sNID.StartsWith(",") Then sNID = Strings.Mid(sNID, 2)
            If sNID.EndsWith(",") Then sNID = Strings.Left(sNID, Len(sNID) - 1)
        Loop

        If Len(sNID) Then
            App農地基本台帳.TBL転用農地.FindRowBySQL("[ID] In (" & sNID & ")")
            Dim pViewT As New DataView(App農地基本台帳.TBL転用農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)
            For Each pRowV As DataRowView In pViewT
                If InStr("," & sNID & ",", "," & pRowV.Item("ID") & ",") Then
                    sNID = Replace("," & sNID & ",", "," & pRowV.Item("ID") & ",", ",")
                End If
                nCount += 1

                s登記地目 = s登記地目 & p案件内集計.R & pRowV.Item("登記簿地目名").ToString
                s現況地目 = s現況地目 & p案件内集計.R & pRowV.Item("現況地目名").ToString
                s農委地目 = s農委地目 & p案件内集計.R & pRowV.Item("農委地目名").ToString
                s自小作別 = s自小作別 & p案件内集計.R & IIf(Val(pRowV.Item("自小作別").ToString) = 0, "自", "小")
                s持分 = s持分 & p案件内集計.R & IIf(Val(pRowV.Item("共有持分分子").ToString) > 0 And Val(pRowV.Item("共有持分分母").ToString) > 0, Val(pRowV.Item("共有持分分子").ToString) & "/" & Val(pRowV.Item("共有持分分母").ToString), "")
                s農振区分 = s農振区分 & p案件内集計.R & IIf(Val(pRowV.Item("農振法区分").ToString) > 0,
                                                    IIf(Val(pRowV.Item("農業振興地域").ToString) = 0, "他", IIf(Val(pRowV.Item("農業振興地域").ToString) = 1, "内", IIf(Val(pRowV.Item("農業振興地域").ToString) = 2, "外", "-"))),
                                                    IIf(Val(pRowV.Item("農振法区分").ToString) = 1, "内", IIf(Val(pRowV.Item("農振法区分").ToString) = 2, "他", IIf(Val(pRowV.Item("農振法区分").ToString) = 3, "外", "-"))))
                p案件内集計.Set筆情報(pRowV.Row, pRow.Row, p総括Item, b再設定)
            Next

            pSheet.ValueReplace("{申請者Ａ貸付管理者}", "[" & s管理者内訳 & "]" & pRow.Item("代理人名").ToString)
        End If

        With p案件内集計
            pSheet.ValueReplace("{筆数計}", nCount)
            pSheet.ValueReplace("{残筆数}", nCount - 1)
            pSheet.ValueReplace("{土地の所在}", "" & .明細作成.To土地所在文字列("&#10;"))
            pSheet.ValueReplace("{土地の所在B}", "" & Replace地番(.明細作成.To土地所在文字列("&#10;")))
            pSheet.ValueReplace("{調査票土地の所在}", (Split(.明細作成.To土地所在文字列("&#10;"), "&#10;")(0)))
            pSheet.ValueReplace("{地目}", s登記地目)
            pSheet.ValueReplace("{登記地目}", s登記地目)
            pSheet.ValueReplace("{現況地目}", s現況地目)
            pSheet.ValueReplace("{農委地目}", s農委地目)
            pSheet.ValueReplace("{自小作別}", s自小作別)
            pSheet.ValueReplace("{持分}", s持分)
            pSheet.ValueReplace("{農振区分}", s農振区分)

            p案件内集計.Replace案件毎集計(pSheet)
            mvar総合計.Add合計(p案件内集計)

            pSheet.ValueReplace("{面積}", .明細作成.To面積文字列("&#10;"))
            pSheet.ValueReplace("{面積内}", .明細作成.To面積内文字列("&#10;"))
        End With
        'Catch ex As Exception
        '    Stop
        'End Try
        Return p案件内集計
    End Function

    Protected Class C筆明細と集計作成
        Public 田数計 As Integer = 0
        Public 畑数計 As Integer = 0
        Public 樹数計 As Integer = 0
        Public 他数計 As Integer = 0

        Public 田面計 As Decimal = 0
        Public 畑面計 As Decimal = 0
        Public 樹面計 As Decimal = 0
        Public 他面計 As Decimal = 0

        Public Is田内 As Boolean = False
        Public Is畑内 As Boolean = False
        Public Is樹内 As Boolean = False
        Public Is他内 As Boolean = False

        Public 田面計内 As Decimal = 0
        Public 畑面計内 As Decimal = 0
        Public 樹面計内 As Decimal = 0
        Public 他面計内 As Decimal = 0



        Public 明細作成 As New C明細作成
        Public R As String = ""

        Public Sub New()

        End Sub

        Public ReadOnly Property 総面積() As Decimal
            Get
                Return 田面計 + 畑面計 + 他面計
            End Get
        End Property

        Public Sub Set筆情報(ByVal pRow As DataRow, ByVal p申請Row As DataRow, ByVal p総括Item As dt総括表, ByVal b再設定 As Boolean)
            Dim n田面積 As Decimal = 0
            Dim n田面積内 As Decimal = 0
            Dim n畑面積 As Decimal = 0
            Dim n畑面積内 As Decimal = 0
            Dim n樹面積 As Decimal = 0
            Dim n樹面積内 As Decimal = 0
            Dim n他面積 As Decimal = 0
            Dim n他面積内 As Decimal = 0
            Dim n部分面積 As Decimal = 0

            '/*************
            Dim pTBL As New HimTools2012.Data.DataTableEx
            If Not IsDBNull(p申請Row.Item("パラメータリスト")) Then
                pTBL.LoadText(p申請Row.Item("パラメータリスト").ToString())
                If pTBL.Columns.Count = 0 Then
                    pTBL.Columns.Add(New DataColumn("Key", GetType(String)))
                End If
                If Not pTBL.Columns.Contains("申請部分面積") Then
                    pTBL.Columns.Add(New DataColumn("申請部分面積", GetType(Decimal)))
                End If
                pTBL.SetPrimaryKey("Key")
            End If

            Dim pParamRow = Nothing
            If pTBL IsNot Nothing AndAlso pTBL.PrimaryKey.Length > 0 Then
                pParamRow = pTBL.Rows.Find("農地." & pRow.Item("ID"))
                If pParamRow Is Nothing Then
                    pParamRow = pTBL.Rows.Find("転用農地." & pRow.Item("ID"))
                End If
            End If

            If pParamRow IsNot Nothing AndAlso pTBL.Columns.Contains("申請部分面積") AndAlso Not IsDBNull(pParamRow.Item("申請部分面積")) Then
                n部分面積 = pParamRow.Item("申請部分面積")
            Else
                n部分面積 = Val(pRow.Item("部分面積").ToString)
            End If
            '/************

            If pRow.Item("登記簿地目名").ToString.IndexOf("田") > -1 AndAlso pRow.Item("登記簿地目名").ToString <> "塩田" Then
                田数計 += 1
                n田面積 += Val(pRow.Item("登記簿面積").ToString)
                n田面積内 = IIf(n部分面積 > 0, n部分面積, Val(pRow.Item("登記簿面積").ToString))
                Is田内 = n部分面積 > 0　'20190605修正

                田面計 += n田面積
                田面計内 += n田面積内
            ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("畑") > -1 Then
                畑数計 += 1
                n畑面積 += Val(pRow.Item("登記簿面積").ToString)
                n畑面積内 = IIf(n部分面積 > 0, n部分面積, Val(pRow.Item("登記簿面積").ToString))
                Is畑内 = n部分面積 > 0　'20190605修正

                畑面計 += n畑面積
                畑面計内 += n畑面積内
            ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("樹") > -1 Then
                樹数計 += 1
                n樹面積 += Val(pRow.Item("登記簿面積").ToString)
                n樹面積内 = IIf(n部分面積 > 0, n部分面積, Val(pRow.Item("登記簿面積").ToString))
                Is樹内 = n部分面積 > 0　'20190605修正

                樹面計 += n樹面積
                樹面計内 += n樹面積内
            Else
                他数計 += 1
                n他面積 += Val(pRow.Item("登記簿面積").ToString)
                n他面積内 = IIf(n部分面積 > 0, n部分面積, Val(pRow.Item("登記簿面積").ToString))
                Is他内 = n部分面積 > 0　'20190605修正

                他面計 += n他面積
                他面計内 += n他面積内
            End If

            明細作成.Plus(pRow, Val(pRow.Item("登記簿面積").ToString), n部分面積, n田面積, n畑面積)
            R = "&#10;"

            If p総括Item IsNot Nothing Then
                If pRow.Item("登記簿地目名").ToString.IndexOf("田") > -1 Then
                    p総括Item.Set面積(dt総括表.enum地目.田, n田面積内, b再設定)
                ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("畑") > -1 Then
                    p総括Item.Set面積(dt総括表.enum地目.畑, n畑面積内, b再設定)
                ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("樹") > -1 Then
                    p総括Item.Set面積(dt総括表.enum地目.樹, n樹面積内, b再設定)
                Else
                    p総括Item.Set面積(dt総括表.enum地目.他, n他面積内, b再設定)
                End If
            End If

        End Sub

        Public Sub Replace案件毎集計(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)
            pSheet.ValueReplace("{田筆数計}", 田数計)
            pSheet.ValueReplace("{田面積計}", IIf(田面計内 > 0, Replace(田面計内.ToString("#,##0.00"), ".00", ""), ""))
            pSheet.ValueReplace("{田面積計内}", IIf(田面計 > 0, 田面計.ToString("#,##0"), "") & IIf(Is田内, "(内" & 田面計内.ToString("#,##0") & ")", ""))

            pSheet.ValueReplace("{畑筆数計}", 畑数計)
            pSheet.ValueReplace("{畑面積計}", IIf(畑面計内 > 0, Replace(畑面計内.ToString("#,##0"), ".00", ""), ""))
            pSheet.ValueReplace("{畑面積計内}", IIf(畑面計 > 0, 畑面計.ToString("#,##0"), "") & IIf(Is畑内, "(内" & 畑面計内.ToString("#,##0") & ")", ""))

            pSheet.ValueReplace("{樹筆数計}", 樹数計)
            pSheet.ValueReplace("{樹面積計}", IIf(樹面計内 > 0, Replace(樹面計内.ToString("#,##0"), ".00", ""), ""))
            pSheet.ValueReplace("{樹面積計内}", IIf(樹面計 > 0, 樹面計.ToString("#,##0"), "") & IIf(Is樹内, "(内" & 樹面計内.ToString("#,##0") & ")", ""))

            pSheet.ValueReplace("{他筆数計}", 他数計)
            pSheet.ValueReplace("{他面積計}", IIf(他面計内 > 0, Replace(他面計内.ToString("#,##0"), ".00", ""), ""))
            pSheet.ValueReplace("{他面積計内}", IIf(他面計 > 0, 他面計.ToString("#,##0"), "") & IIf(Is他内, "(内" & 他面計内.ToString("#,##0") & ")", ""))

            Dim 総計 As Decimal = 田面計 + 畑面計 + 樹面計 + 他面計
            Dim 総計内 As Decimal = 田面計内 + 畑面計内 + 樹面計内 + 他面計内
            pSheet.ValueReplace("{面積計}", Replace(総計内.ToString("#,##0.00"), ".00", ""))
            pSheet.ValueReplace("{面積計内}", 総計.ToString("#,##0") & IIf((Is田内 Or Is畑内 Or Is樹内 Or Is他内), "(内" & 総計内.ToString("#,##0") & ")", ""))
        End Sub

        Public Sub Replace明細総合計(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)
            Dim 総筆数 As Decimal = 田数計 + 畑数計 + 樹数計 + 他数計
            Dim 総面計 As Decimal = 田面計 + 畑面計 + 樹面計 + 他面計
            Dim 総面計内 As Decimal = 田面計内 + 畑面計内 + 樹面計内 + 他面計内

            With pSheet
                .ValueReplace("{田筆数総合計}", 田数計.ToString("#,##0"))
                .ValueReplace("{田面積総合計}", 田面計内.ToString("#,##0"))
                .ValueReplace("{田面積総合計内}", 田面計.ToString("#,##0") & IIf(Is田内, "(内" & 田面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{畑筆数総合計}", 畑数計.ToString("#,##0"))
                .ValueReplace("{畑面積総合計}", 畑面計内.ToString("#,##0"))
                .ValueReplace("{畑面積総合計内}", 畑面計.ToString("#,##0") & IIf(Is畑内, "(内" & 畑面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{樹筆数総合計}", 樹数計.ToString("#,##0"))
                .ValueReplace("{樹面積総合計}", 樹面計内.ToString("#,##0"))
                .ValueReplace("{樹面積総合計内}", 樹面計.ToString("#,##0") & IIf(Is樹内, "(内" & 樹面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{他筆数総合計}", 他数計.ToString("#,##0"))
                .ValueReplace("{他面積総合計}", 他面計内.ToString("#,##0"))
                .ValueReplace("{他面積総合計内}", 他面計.ToString("#,##0") & IIf(Is他内, "(内" & 他面計内.ToString("#,##0") & ")", ""))

                .ValueReplace("{筆数総合計}", 総筆数.ToString("#,##0"))
                .ValueReplace("{面積総合計}", 総面計内.ToString("#,##0"))
                .ValueReplace("{面積総合計内}", 総面計.ToString("#,##0") & IIf(Is田内 Or Is畑内 Or Is樹内 Or Is他内, "(内" & 総面計内.ToString("#,##0") & ")", ""))
            End With
        End Sub

        Public Sub Add合計(ByVal p案件毎 As C筆明細と集計作成)
            田数計 += p案件毎.田数計
            畑数計 += p案件毎.畑数計
            樹数計 += p案件毎.樹数計
            他数計 += p案件毎.他数計

            田面計 += p案件毎.田面計
            畑面計 += p案件毎.畑面計
            樹面計 += p案件毎.樹面計
            他面計 += p案件毎.他面計

            Is田内 = Is田内 Or p案件毎.Is田内
            Is畑内 = Is畑内 Or p案件毎.Is畑内
            Is樹内 = Is樹内 Or p案件毎.Is樹内
            Is他内 = Is他内 Or p案件毎.Is他内

            田面計内 += p案件毎.田面計内
            畑面計内 += p案件毎.畑面計内
            樹面計内 += p案件毎.樹面計内
            他面計内 += p案件毎.他面計内
        End Sub

        Public Class C明細作成
            Inherits List(Of 筆毎面積)

            Public Sub Plus(ByVal pRow As DataRow, ByVal p本地面積 As Decimal, ByVal p部分面積 As Decimal, ByVal p田面積 As Decimal, ByVal p畑面積 As Decimal)
                Dim s土地所在 As String
                If pRow.Item("小字").ToString = "" OrElse pRow.Item("小字").ToString = "-" Then
                    s土地所在 = IIf(pRow.Item("所在").ToString.Length > 0, pRow.Item("所在").ToString, pRow.Item("大字").ToString & pRow.Item("地番").ToString)
                Else
                    s土地所在 = IIf(pRow.Item("所在").ToString.Length > 0, pRow.Item("所在").ToString, pRow.Item("大字").ToString & IIf(pRow.Item("小字").ToString.Length > 0, "字", "") & pRow.Item("小字").ToString) & pRow.Item("地番").ToString
                End If

                Me.Add(New 筆毎面積(s土地所在, p本地面積, p部分面積, p田面積, p畑面積))
            End Sub

            Public Function To面積文字列(Optional bySp As String = ",") As String
                Dim sB As New System.Text.StringBuilder

                For Each pNum As 筆毎面積 In Me
                    If Fix(pNum.本地面積) = pNum.本地面積 Then
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0"))
                    Else
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0.00"))
                    End If
                Next

                Return sB.ToString
            End Function
            Public Function To面積内文字列(Optional bySp As String = ",") As String
                Dim sB As New System.Text.StringBuilder

                For Each pNum As 筆毎面積 In Me
                    If Fix(pNum.本地面積) = pNum.本地面積 Then
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0"))
                    Else
                        sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.本地面積.ToString("#,##0.00"))
                    End If
                    Dim pD As Decimal = pNum.部分面積
                    If pD > 0 Then
                        sB.Append("(内")
                        If Fix(pNum.部分面積) = pNum.部分面積 Then
                            sB.Append(pNum.部分面積.ToString("#,##0"))
                        Else
                            sB.Append(pNum.部分面積.ToString("#,##0.00"))
                        End If
                        sB.Append(")")
                    End If
                Next


                Return sB.ToString
            End Function

            Public Function To土地所在文字列(Optional bySp As String = ",") As String
                Dim sB As New System.Text.StringBuilder

                For Each pNum As 筆毎面積 In Me
                    sB.Append(IIf(sB.Length > 0, bySp, "") & pNum.土地所在)
                Next

                Return sB.ToString
            End Function

            Public Class 筆毎面積
                Public 土地所在 As String = ""
                Public 本地面積 As Decimal
                Public 部分面積 As Decimal
                Public 田面積 As Decimal
                Public 畑面積 As Decimal

                Public Sub New(ByVal s所在 As String, ByVal d本地面積 As Decimal, ByVal d部分面積 As Decimal, ByVal dt田面積 As Decimal, ByVal dt畑面積 As Decimal)
                    土地所在 = s所在
                    本地面積 = d本地面積
                    部分面積 = d部分面積
                    田面積 = dt田面積
                    畑面積 = dt畑面積
                End Sub

            End Class
        End Class

    End Class

    Public Sub 転用共通(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{資金計画}", pRow.Item("資金計画").ToString)
        If pRow.Item("数量").ToString.Length > 0 AndAlso pRow.Item("数量") > 0 Then
            pSheet.ValueReplace("{棟数}", pRow.Item("数量").ToString & "棟")
        Else
            pSheet.ValueReplace("{棟数}", "")
        End If
        If pRow.Item("数量").ToString.Length > 0 AndAlso pRow.Item("数量") > 0 Then
            pSheet.ValueReplace("{棟数n}", pRow.Item("数量").ToString & "棟")
        Else
            pSheet.ValueReplace("{棟数n}", " 0")
        End If

        pSheet.ValueReplace("{不許可例外}", pRow.Item("不許可例外").ToString)

        If pRow.Item("建築面積").ToString.Length > 0 AndAlso pRow.Item("建築面積") > 0 Then
            pSheet.ValueReplace("{建築面積}", pRow.Item("建築面積").ToString & "㎡")
        Else
            pSheet.ValueReplace("{建築面積}", "")
        End If
        If pRow.Item("用途").ToString.Length > 0 Then
            pSheet.ValueReplace("{用途}", pRow.Item("用途").ToString)
        Else
            pSheet.ValueReplace("{用途}", "")
        End If

        If pRow.Item("申請地目安").ToString.Length > 0 Then
            pSheet.ValueReplace("{申請地目安}", pRow.Item("申請地目安").ToString)
        Else
            pSheet.ValueReplace("{申請地目安}", "")
        End If


        '
        If Not IsDBNull(pRow.Item("同意書")) AndAlso pRow.Item("同意書") Then
            pSheet.ValueReplace("{同意書}", "同意書あり&#10;")
            pSheet.ValueReplace("{同意書○}", "○")
        Else
            pSheet.ValueReplace("{同意書}", "")
            pSheet.ValueReplace("{同意書○}", "")
        End If
        If Not IsDBNull(pRow.Item("始末書")) AndAlso pRow.Item("始末書") Then
            pSheet.ValueReplace("{始末書}", "始末書あり&#10;")
            pSheet.ValueReplace("{始末書○}", "○")
        Else
            pSheet.ValueReplace("{始末書}", "")
            pSheet.ValueReplace("{始末書○}", "")
        End If
        If Not IsDBNull(pRow.Item("被害防除計画書")) AndAlso pRow.Item("被害防除計画書") Then
            pSheet.ValueReplace("{被害防除計画書}", "被害防除計画あり&#10;")
            pSheet.ValueReplace("{被害防除計画書○}", "○")
        Else
            pSheet.ValueReplace("{被害防除計画書}", "")
            pSheet.ValueReplace("{被害防除計画書○}", "")
        End If


        If Not IsDBNull(pRow.Item("建ぺい率")) AndAlso pRow.Item("建ぺい率") <> 0 Then
            pSheet.ValueReplace("{建ぺい率}", pRow.Item("建ぺい率"))
        Else
            pSheet.ValueReplace("{建ぺい率}", "")
        End If
        Dim St1 As String = IIf(pRow.Item("工事開始年1").ToString.Length > 0, pRow.Item("工事開始年1"), "") & "." & IIf(pRow.Item("工事開始月1").ToString.Length > 0, pRow.Item("工事開始月1"), "")
        Dim St2 As String = IIf(pRow.Item("工事終了年1").ToString.Length > 0, pRow.Item("工事終了年1"), "") & "." & IIf(pRow.Item("工事終了月1").ToString.Length > 0, pRow.Item("工事終了月1"), "")

        If St1.Length > 1 Then St1 = St1 & "～&#10;"
        If St2.Length > 1 Then St1 = St1 & St2
        '
        pSheet.ValueReplace("{工事着工}", St1)
        pSheet.ValueReplace("{工事計画}", St1)
        '
        If Not IsDBNull(pRow.Item("理由書")) AndAlso pRow.Item("理由書") Then
            pSheet.ValueReplace("{理由書}", "面積超の理由書あり&#10;")
        Else
            pSheet.ValueReplace("{理由書}", "")
        End If

        pSheet.ValueReplace("{目的}", "")
        If Not IsDBNull(pRow.Item("資金計画")) AndAlso pRow.Item("資金計画").ToString.Length > 0 Then
            pSheet.ValueReplace("{資金}", pRow.Item("資金計画").ToString)
        Else
            pSheet.ValueReplace("{資金}", "")
        End If
        区分設定(pSheet, pRow)
    End Sub
    Public Sub 区分設定(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        If Not IsDBNull(pRow.Item("農地の広がり")) Then
            '{農地の広がり}
            Select Case Val(pRow.Item("農地の広がり").ToString)
                Case 0 : pSheet.ValueReplace("{農地の広がり}", "10ha未満")
                Case 1 : pSheet.ValueReplace("{農地の広がり}", "10ha以上")
                Case Else
                    pSheet.ValueReplace("{農地の広がり}", "")
            End Select
        Else
            pSheet.ValueReplace("{農地の広がり}", "")
        End If


        If Not IsDBNull(pRow.Item("土地改良事業の有無")) AndAlso pRow.Item("土地改良事業の有無") Then
            pSheet.ValueReplace("{土地改良事業の有無}", "あり")
        Else
            pSheet.ValueReplace("{土地改良事業の有無}", "なし")
        End If


        If Not IsDBNull(pRow.Item("土地改良区の意見書の有不用")) Then
            Dim pV As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='申請書_土地改良区の意見書について' AND [ID]=" & pRow.Item("土地改良区の意見書の有不用"), "ID", DataViewRowState.CurrentRows)
            If pV.Count > 0 Then
                pSheet.ValueReplace("{土地改良区意見書表示}", pV(0).Item("名称"))
            Else
                pSheet.ValueReplace("{土地改良区意見書表示}", "")
            End If
        Else
            pSheet.ValueReplace("{土地改良区意見書表示}", "")
        End If


        If IsDBNull(pRow.Item("土地改良区の意見書の有無")) Then
            '
            pSheet.ValueReplace("{土地改良区意見書}", "")
        Else
            Select Case pRow.Item("土地改良区の意見書の有無")
                Case True
                    pSheet.ValueReplace("{土地改良区意見書}", "土地改良区の意見書あり")
                Case Else
                    pSheet.ValueReplace("{土地改良区意見書}", "")
            End Select
        End If

        pSheet.ValueReplace("{農地区分}", pRow.Item("農地区分名称").ToString)
        Select Case Val(pRow.Item("農地区分").ToString)
            Case 1 : pSheet.ValueReplace("{諮問農地区分}", 1)
            Case 2 : pSheet.ValueReplace("{諮問農地区分}", 2)
            Case 3 : pSheet.ValueReplace("{諮問農地区分}", 3)
            Case 4, 5 : pSheet.ValueReplace("{諮問農地区分}", "農振農用地")
            Case Else : pSheet.ValueReplace("{諮問農地区分}", "-")
        End Select

        pSheet.ValueReplace("{農振区分名称}", pRow.Item("農振区分名称").ToString)
        pSheet.ValueReplace("{都計区分名称}", pRow.Item("都計区分名称").ToString)


    End Sub

    Public Sub 受付情報(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{受付番号}", pRow.Item("受付番号"))
        pSheet.ValueReplace("{受付年月日和略}", 和暦Format(pRow.Item("受付年月日"), "gyy.MM.dd", ""))
        pSheet.ValueReplace("{受付年月日}", 和暦Format(CDate(pRow.Item("受付年月日"))))

        If DatePart("yyyy", pRow.Item("受付年月日").ToString) >= 2019 AndAlso DatePart("M", pRow.Item("受付年月日").ToString) >= 5 Then
            pSheet.ValueReplace("{受付年}", DatePart("yyyy", pRow.Item("受付年月日").ToString) - 2018)
        Else
            pSheet.ValueReplace("{受付年}", DatePart("yyyy", pRow.Item("受付年月日").ToString) - 1988)
        End If

        pSheet.ValueReplace("{受付月}", DatePart("M", pRow.Item("受付年月日").ToString))
        pSheet.ValueReplace("{受付日}", DatePart("d", pRow.Item("受付年月日").ToString))
    End Sub

    Public Sub 貸借共通(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        Try
            If Val(pRow.Item("小作料").ToString) > 0 Then
                pSheet.ValueReplace("{小作料}", Format(Val(pRow.Item("小作料").ToString), "#,###") & pRow.Item("小作料単位").ToString)
            Else
                pSheet.ValueReplace("{小作料}", "")
            End If

            Select Case Val(pRow.Item("権利種類").ToString)
                Case 1 : pSheet.ValueReplace("{権利種類}", "賃借権") : pSheet.ValueReplace("{契約内容}", "賃借権")
                Case 2, 104 : pSheet.ValueReplace("{権利種類}", "使用貸借権") : pSheet.ValueReplace("{契約内容}", "使用貸借権")
                Case 4, 101 : pSheet.ValueReplace("{契約内容}", "地上権")
                Case 5, 102 : pSheet.ValueReplace("{契約内容}", "永小作権")
                Case 6, 103 : pSheet.ValueReplace("{契約内容}", "質権")
                Case 7 : pSheet.ValueReplace("{契約内容}", "期間借地")
                Case 8 : pSheet.ValueReplace("{契約内容}", "残存小作地")
                Case 9 : pSheet.ValueReplace("{契約内容}", "使用賃借")
                Case 105 : pSheet.ValueReplace("{契約内容}", "賃貸借権")
                Case 106 : pSheet.ValueReplace("{契約内容}", "使用貸借権（期間借地）")
                Case 107 : pSheet.ValueReplace("{契約内容}", "賃貸借権（期間借地）")
                Case 108 : pSheet.ValueReplace("{契約内容}", "使用貸借権（円滑化）")
                Case 109 : pSheet.ValueReplace("{契約内容}", "賃貸借権（円滑化）")
                Case 199 : pSheet.ValueReplace("{契約内容}", "調査中")
                Case Else
                    pSheet.ValueReplace("{権利種類}", "その他") : pSheet.ValueReplace("{契約内容}", "その他")
            End Select

            If IsDBNull(pRow.Item("永久")) AndAlso 1 Then

            End If
            If Not IsDBNull(pRow.Item("始期")) AndAlso Not IsDBNull(pRow.Item("終期")) Then
                Dim pB As TimeSpan = CDate(pRow.Item("終期")).Subtract(pRow.Item("始期"))

                pSheet.ValueReplace("{始期年月日}", 和暦Format(CDate(pRow.Item("始期"))))
                pSheet.ValueReplace("{終期年月日}", 和暦Format(CDate(pRow.Item("終期"))))
                pSheet.ValueReplace("{始期年月日D}", 和暦Format(CDate(pRow.Item("始期")), "gyy.M.d"))
                pSheet.ValueReplace("{終期年月日D}", 和暦Format(CDate(pRow.Item("終期")), "gyy.M.d"))
                pSheet.ValueReplace("{始期年月日S}", 和暦Format(CDate(pRow.Item("始期")), "gyy/M/d"))
                pSheet.ValueReplace("{終期年月日S}", 和暦Format(CDate(pRow.Item("終期")), "gyy/M/d"))

                If Val(pRow.Item("期間").ToString) > 0 Then
                    pSheet.ValueReplace("{契約期間}", Val(pRow.Item("期間").ToString) & "年")
                ElseIf pB.Days > 365 Then
                    pSheet.ValueReplace("{契約期間}", Int(pB.Days / 365) & "年")
                Else
                    pSheet.ValueReplace("{契約期間}", "")
                End If
            ElseIf Val(pRow.Item("期間").ToString) > 0 Then
                pSheet.ValueReplace("{契約期間}", Val(pRow.Item("期間").ToString) & "年")
                pSheet.ValueReplace("{始期年月日}", "")
                pSheet.ValueReplace("{終期年月日}", "")
                pSheet.ValueReplace("{始期年月日D}", "")
                pSheet.ValueReplace("{終期年月日D}", "")
                pSheet.ValueReplace("{始期年月日S}", "")
                pSheet.ValueReplace("{終期年月日S}", "")
            Else
                pSheet.ValueReplace("{契約期間}", "")
                pSheet.ValueReplace("{始期年月日}", "")
                pSheet.ValueReplace("{終期年月日}", "")
                pSheet.ValueReplace("{始期年月日D}", "")
                pSheet.ValueReplace("{終期年月日D}", "")
                pSheet.ValueReplace("{始期年月日S}", "")
                pSheet.ValueReplace("{終期年月日S}", "")
            End If

            pSheet.ValueReplace("{支払方法}", pRow.Item("支払方法").ToString)
            If IsDBNull(pRow.Item("再設定")) Then
                pSheet.ValueReplace("{設定内容}", "新")
            Else
                pSheet.ValueReplace("{設定内容}", IIf(pRow.Item("再設定"), "再", "新"))
            End If
        Catch ex As Exception
            Stop
        End Try

        pSheet.ValueReplace("{契約期間}", "")
        pSheet.ValueReplace("{期間}", "")

    End Sub

    Public Sub 権利内容(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView)
        Select Case CType(pRow.Item("法令"), enum法令)
            Case enum法令.農地法5条所有権
                pSheet.ValueReplace("{権利内容}", "所有権移転")
            Case enum法令.農地法5条貸借
                Select Case Val(pRow.Item("権利種類").ToString)
                    Case 1
                        pSheet.ValueReplace("{権利内容}", "賃借権")
                        pSheet.ValueReplace("{権利内容S}", "賃")
                    Case 2
                        pSheet.ValueReplace("{権利内容}", "使用貸借権")
                        pSheet.ValueReplace("{権利内容S}", "使")
                    Case Else
                        pSheet.ValueReplace("{権利内容}", "その他")
                        pSheet.ValueReplace("{権利内容S}", "他")
                End Select
            Case enum法令.農地法5条一時転用
                pSheet.ValueReplace("{権利内容}", "期間借地")
        End Select
    End Sub


    Public Sub 申請者A(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{氏名A}", pRow.Item("氏名A").ToString)
        pSheet.ValueReplace("{住所A}", pRow.Item("住所A").ToString)

        If pRow.Item("代理人名").ToString.Length > 0 Then
            pSheet.ValueReplace("{代理人}", "[代理人]&#10;" & pRow.Item("代理人住所").ToString & "&#10;" & pRow.Item("代理人名").ToString)
            pSheet.ValueReplace("{代理人名}", "[代理人]" & pRow.Item("代理人名").ToString)
            pSheet.ValueReplace("{代理人名B}", pRow.Item("代理人名").ToString)
            pSheet.ValueReplace("{代理人住所B}", pRow.Item("代理人住所").ToString)
            pSheet.ValueReplace("{代理人名C}", pRow.Item("代理人名").ToString)
            pSheet.ValueReplace("{代理人住所C}", pRow.Item("代理人住所").ToString)
        Else
            pSheet.ValueReplace("{代理人}", "")
            pSheet.ValueReplace("{代理人名}", "")
            pSheet.ValueReplace("{代理人名B}", pRow.Item("氏名A").ToString)
            pSheet.ValueReplace("{代理人住所B}", pRow.Item("住所A").ToString)
            pSheet.ValueReplace("{代理人名C}", "")
            pSheet.ValueReplace("{代理人住所C}", "")
        End If

        pSheet.ValueReplace("{職業A}", pRow.Item("職業A").ToString)
        pSheet.ValueReplace("{年齢A}", pRow.Item("年齢A").ToString)
        pSheet.ValueReplace("{年齢A2}", pRow.Item("年齢A").ToString & "歳")
        pSheet.ValueReplace("{経営面積A}", Val(pRow.Item("経営面積A").ToString).ToString("#,##0"))
        pSheet.ValueReplace("{申請者Ａ経営面積}", Val(pRow.Item("経営面積A").ToString).ToString("#,##0"))

        pSheet.ValueReplace("{申請者Ａ氏名}", pRow.Item("氏名A").ToString)
        pSheet.ValueReplace("{申請者Ａ住所}", pRow.Item("住所A").ToString)
        pSheet.ValueReplace("{申請者Ａ職業}", pRow.Item("職業A").ToString)
        pSheet.ValueReplace("{申請者Ａ年齢}", IIf(Val(pRow.Item("年齢A").ToString) = 0, "-", pRow.Item("年齢A").ToString))
        pSheet.ValueReplace("{申請者Ａ年齢2}", IIf(Val(pRow.Item("年齢A").ToString) = 0, "-", pRow.Item("年齢A").ToString) & "歳")
        pSheet.ValueReplace("{申請者Ａ経営面積}", Val(pRow.Item("経営面積A").ToString).ToString("#,##0"))
        pSheet.ValueReplace("{申請者Ａ申請理由}", pRow.Item("申請理由A").ToString)
        pSheet.ValueReplace("{申請者Ａ集落名}", pRow.Item("集落A").ToString)
    End Sub

    Public Sub 申請者B(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{氏名B}", pRow.Item("氏名B").ToString)
        pSheet.ValueReplace("{住所B}", pRow.Item("住所B").ToString)
        pSheet.ValueReplace("{職業B}", pRow.Item("職業B").ToString)
        pSheet.ValueReplace("{年齢B}", pRow.Item("年齢B").ToString)
        pSheet.ValueReplace("{経営面積B}", Val(pRow.Item("経営面積B").ToString).ToString("#,###"))
        pSheet.ValueReplace("{申請者Ｂ経営面積}", Val(pRow.Item("経営面積B").ToString).ToString("#,###"))
        pSheet.ValueReplace("{申請理由B}", pRow.Item("申請理由B").ToString)

        pSheet.ValueReplace("{申請者Ｂ氏名}", pRow.Item("氏名B").ToString)
        pSheet.ValueReplace("{申請者Ｂ住所}", pRow.Item("住所B").ToString)
        pSheet.ValueReplace("{申請者Ｂ職業}", pRow.Item("職業B").ToString)
        pSheet.ValueReplace("{申請者Ｂ年齢}", pRow.Item("年齢B").ToString)
        pSheet.ValueReplace("{申請者Ｂ経営面積}", GetCVNumber(pRow.Item("経営面積B")))
        pSheet.ValueReplace("{申請者Ｂ申請理由}", pRow.Item("申請理由B").ToString)
        pSheet.ValueReplace("{申請者Ｂ集落名}", pRow.Item("集落B").ToString)
        pSheet.ValueReplace("{申請者Ｂ労力}", pRow.Item("稼動人数").ToString)

    End Sub

    Public Sub 調査委員(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        If Not IsDBNull(pRow.Item("農業委員1")) Then
            Dim pRow1 As DataRow = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員1"), "農業委員"})
            If pRow1 IsNot Nothing Then
                pSheet.ValueReplace("{調査委員}", IIf(IsDBNull(pRow1.Item("名称")), pRow.Item("調査員A").ToString, pRow1.Item("名称").ToString))
                pSheet.ValueReplace("{調査委員1}", IIf(IsDBNull(pRow1.Item("名称")), pRow.Item("調査員A").ToString, pRow1.Item("名称").ToString))
                pSheet.ValueReplace("{農業委員1}", IIf(IsDBNull(pRow1.Item("名称")), pRow.Item("調査員A").ToString, pRow1.Item("名称").ToString))
            End If
        End If
        pSheet.ValueReplace("{調査委員}", pRow.Item("調査員A").ToString)
        pSheet.ValueReplace("{調査委員1}", pRow.Item("調査員A").ToString)
        pSheet.ValueReplace("{農業委員1}", pRow.Item("調査員A").ToString)

        If Not IsDBNull(pRow.Item("農業委員2")) Then
            Dim pRow2 As DataRow = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員2"), "農業委員"})
            If pRow2 IsNot Nothing Then
                pSheet.ValueReplace("{調査委員2}", pRow2.Item("名称").ToString)
                pSheet.ValueReplace("{農業委員2}", pRow2.Item("名称").ToString)
            End If
        End If
        pSheet.ValueReplace("{担当委員}", pRow.Item("調査員B").ToString)
        pSheet.ValueReplace("{調査委員2}", pRow.Item("調査員B").ToString)
        pSheet.ValueReplace("{農業委員2}", pRow.Item("調査員B").ToString)

        If Not IsDBNull(pRow.Item("農業委員3")) Then
            Dim pRow3 As DataRow = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員3"), "農業委員"})
            If pRow3 IsNot Nothing Then
                pSheet.ValueReplace("{調査委員3}", pRow3.Item("名称").ToString)
                pSheet.ValueReplace("{農業委員3}", pRow3.Item("名称").ToString)
            End If
        End If
        pSheet.ValueReplace("{調査委員3}", "")
        pSheet.ValueReplace("{農業委員3}", "")
    End Sub

    Private Function GetCVNumber(pObj As Object) As String
        If pObj Is Nothing Then
            Return ""
        ElseIf IsDBNull(pObj) Then
            Return ""
        ElseIf CDbl(pObj) > 0 Then
            Return Val(pObj).ToString("#,###")
        Else
            Return ""
        End If
    End Function

    Public Overrides Sub SetData(ByRef XMLSS As HimTools2012.Excel.XMLSS2003.CXMLSS2003, ByRef pTab As 申請Page, s処理名称 As String, ByRef pDataCreater As C総会資料Data作成)

        If XMLSS.WorkBook.WorkSheets.Items.ContainsKey("明細") Then
            With XMLSS.WorkBook.WorkSheets.Items("明細")
                Set複数行(._object, pTab, s処理名称, pDataCreater, "")
                mvar総合計.Replace明細総合計(XMLSS.WorkBook.WorkSheets.Items("明細"))

            End With
        Else
            For Each pS As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet In XMLSS.WorkBook.WorkSheets.Items.Values
                Set複数行(pS, pTab, s処理名称, pDataCreater, "")
                mvar総合計.Replace明細総合計(pS)
            Next
        End If

        If MyClass.GetType = GetType(CPrint総会資料作成農用地利用集積計画貸借) Then
            Dim n法令 As enum法令 = Val(HimTools2012.StringF.Mid(pTab.Name, 3))
            If XMLSS.WorkBook.WorkSheets.Items.ContainsKey("総括") Then
                Select Case n法令
                    Case enum法令.利用権設定, enum法令.中間管理機構経由
                        CType(Me, CPrint総会資料作成農用地利用集積計画貸借).Set総括表(XMLSS.WorkBook.WorkSheets.Items("総括"), pTab, s処理名称, pDataCreater)
                    Case enum法令.基盤強化法所有権
                        CType(Me, CPrint総会資料作成農用地利用集積計画所有権).Set総括表(XMLSS.WorkBook.WorkSheets.Items("総括"), pTab, s処理名称, pDataCreater)
                End Select
            Else
                MsgBox("総会資料：利用権設定に総括表のページがありません", MsgBoxStyle.Critical)
            End If
        End If
    End Sub

    Private Function Replace地番(ByVal 土地所在 As String) As String
        Dim 土地所在B As String = ""
        Dim CountB As Integer = 0
        Dim 土地所在２ As String = ""

        If Not IsDBNull(土地所在) AndAlso 土地所在 <> "" Then
            Dim Ar As Object = Nothing
            If InStr(土地所在, "&#10;") > 0 Then
                Ar = Split(土地所在, "&#10;")
                For n As Integer = 0 To UBound(Ar)
                    If InStr(Ar(n), "-") > 0 Then
                        If UBound(Ar) = n Then
                            土地所在 = 土地所在２ & Ar(n)
                        Else
                            土地所在２ = 土地所在２ & Ar(n) & "&#10;"
                        End If
                    Else
                        If UBound(Ar) = n Then
                            土地所在 = 土地所在２ & Ar(n) & "番"
                        Else
                            土地所在２ = 土地所在２ & Ar(n) & "番&#10;"
                        End If
                    End If
                Next
            Else
                If InStr(土地所在, "-") > 0 Then
                Else
                    土地所在 = 土地所在 & "番"
                End If
            End If

            For n As Integer = 1 To Len(土地所在)
                Select Case Mid(土地所在, n, 1)
                    Case "-"
                        If CountB = 0 Then
                            土地所在B = 土地所在B & Replace(Mid(土地所在, n, 1), "-", "番")
                            CountB += 1
                        Else : 土地所在B = 土地所在B & Replace(Mid(土地所在, n, 1), "-", "の")
                        End If
                    Case "&", "#", ";"
                        土地所在B = 土地所在B & Mid(土地所在, n, 1)
                        CountB = 0
                    Case Else : 土地所在B = 土地所在B & Mid(土地所在, n, 1)
                End Select
            Next
        End If

        Return 土地所在B
    End Function
End Class

Public MustInherit Class CPrint単票作成
    Inherits CPrint資料作成共通

    Public MustOverride Sub Set単票Data(ByRef pDataCreater As C総会資料Data作成, pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet)

    Public Sub New()
        MyBase.New()
    End Sub

    Public Overrides Sub SetData(ByRef XMLSS As HimTools2012.Excel.XMLSS2003.CXMLSS2003, ByRef pTab As 申請Page, s処理名称 As String, ByRef pDataCreater As C総会資料Data作成)


        For Each pS As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet In XMLSS.WorkBook.WorkSheets.Items.Values
            Set単票Data(pDataCreater, pS)
        Next
    End Sub

End Class

Public Class C総会資料Data作成
    Inherits HimTools2012.clsAccessor
    Private mvarTab As TabControl
    Public Sub New(pTabC As TabControl)
        MyBase.New()
        mvarTab = pTabC
    End Sub
    Public ReadOnly Property TabCtrl() As TabControl
        Get
            Return mvarTab
        End Get
    End Property

    Public Overrides Sub Execute()
        Dim sFolder As String = SysAD.OutputFolder & String.Format("\総会資料{0}_{1}", Now.Year, Now.Month)
        If IO.Directory.Exists(SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "")) Then
            sFolder = SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "") & String.Format("\総会資料{0}_{1}", Now.Year, Now.Month)
        End If

        If Not IO.Directory.Exists(sFolder) Then
            IO.Directory.CreateDirectory(sFolder)
        End If

        For Each p申請 As 申請Page In mvarTab.TabPages
            If p申請.印刷 Then

                Select Case p申請.Name
                    Case "n.30" : sub総会資料作成("農地法3条", sFolder, New CPrint総会資料作成農地法3条, p申請, "農地法第3条申請について.xml")
                    Case "n.40" : sub総会資料作成("農地法4条", sFolder, New CPrint総会資料作成農地法4条, p申請, "農地法第4条申請について.xml")
                    Case "n.50" : sub総会資料作成("農地法5条", sFolder, New CPrint総会資料作成農地法5条, p申請, "農地法第5条申請について.xml")
                    Case "n.60" : sub総会資料作成("農用地利用集積計画(所有権移転)", sFolder, New CPrint総会資料作成農用地利用集積計画所有権, p申請, "農用地利用集積計画(所有権移転).xml")
                    Case "n.61" : sub総会資料作成("農用地利用集積計画(貸借)", sFolder, New CPrint総会資料作成農用地利用集積計画貸借, p申請, "農用地利用集積計画(貸借).xml")
                    Case "n.62" : sub総会資料作成("農用地利用集積計画(移転)", sFolder, New CPrint総会資料作成農用地利用集積計画移転, p申請, "農用地利用集積計画(移転).xml")
                    Case "n.65" : sub総会資料作成("中間管理機構経由", sFolder, New CPrint総会資料作成農用地利用集積計画貸借, p申請, "農用地利用集積計画(貸借).xml")
                    Case "n.180" : sub総会資料作成("農地法18条", sFolder, New CPrint総会資料作成18条解約, p申請, "農地法第18条申請について.xml")
                    Case "n.210" : sub総会資料作成("合意解約", sFolder, New CPrint総会資料作成合意解約, p申請, "合意解約について.xml")
                    Case "n.302" : sub総会資料作成("農振地整備計画変更", sFolder, New CPrint総会資料作成農振地整備計画変更, p申請, "農振地整備計画の一部変更申出について.xml")
                    Case "n.303" : sub総会資料作成("事業計画変更", sFolder, New CPrint総会資料作成転用事業計画変更, p申請, "農地転用事業計画変更について.xml")
                    Case "n.400" : sub総会資料作成("あっせん申出(売渡・貸付　希望)", sFolder, New CPrint総会資料作成あっせん出し手, p申請, "あっせん申出(売渡・貸付　希望).xml")
                    Case "n.401" : sub総会資料作成("あっせん申出(買受・借受　希望)", sFolder, New CPrint総会資料作成あっせん受け手, p申請, "あっせん申出(買受・借受　希望).xml")

                    Case "n.500" : sub総会資料作成("農地利用目的変更", sFolder, New CPrint総会資料作成農地利用変更, p申請, "農地利用目的変更について.xml")
                    Case "n.600", "n.602" : sub総会資料作成("非農地証明", sFolder, New CPrint総会資料作成非農地証明, p申請, "非農地証明書願いについて.xml")

                    Case "n.801"
                    Case "n.803"
                    Case "n.804"
                    Case Else
                        'Stop
                End Select
            End If
        Next

        SysAD.ShowFolder(sFolder)
    End Sub


    Private Function sub総会資料作成(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal p作成 As CPrint資料作成共通, ByRef p申請 As 申請Page, ByVal sFile As String) As Boolean
        If System.IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile) Then
            If p申請 Is Nothing Then
                If p作成.LoopSub(s処理名称, sDesktopFolder, p作成, sFile) Then
                    Return True
                Else
                    Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)

                    Dim pXMLSS As New HimTools2012.Excel.XMLSS2003.CXMLSS2003(sXML)
                    p作成.SetData(pXMLSS, p申請, s処理名称, Me)

                    If s処理名称 = "中間管理機構経由" Then
                        sFile = "農用地利用集積計画(機構法).xml"
                    End If

                    Dim sNewFile As String = HimTools2012.FileManager.NewFileName(sDesktopFolder & "\" & sFile, 0)
                    HimTools2012.TextAdapter.SaveTextFile(sNewFile, pXMLSS.OutPut(True))
                    Return True
                End If
            ElseIf IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile) Then
                If p作成.LoopSub(s処理名称, sDesktopFolder, p作成, sFile) Then
                    Return True
                Else
                    If p申請.SelectCount > 0 Then
                        Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)
                        sXML = Replace(sXML, "{議案番号}", IIf(p申請.txt議案番号.Text.Length > 0, p申請.txt議案番号.Text, "   "))
                        sXML = Replace(sXML, "{日程番号}", IIf(p申請.txt日程番号.Text.Length > 0, p申請.txt日程番号.Text, "   "))
                        Dim pXMLSS As New HimTools2012.Excel.XMLSS2003.CXMLSS2003(sXML)

                        p作成.SetData(pXMLSS, p申請, s処理名称, Me)

                        If s処理名称 = "中間管理機構経由" Then
                            sFile = "農用地利用集積計画(機構法).xml"
                        End If

                        Dim sNewFile As String = HimTools2012.FileManager.NewFileName(sDesktopFolder & "\" & sFile, 0)
                        HimTools2012.TextAdapter.SaveTextFile(sNewFile, pXMLSS.OutPut(True))

                        Return True
                    Else
                        Return False
                    End If

                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

End Class



