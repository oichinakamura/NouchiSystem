Imports HimTools2012
Imports HimTools2012.NumericFunctions


''' <summary>
''' 多筆型耕作証明
''' </summary>
''' <remarks></remarks>
Public Class CPrint耕作多筆証明願
    Inherits CPrint耕作証明共通

    Private mvarData As PrnData

    Private Structure PrnData
        Dim 住所 As String

        Dim n自筆() As Integer
        Dim n小筆() As Integer
        Dim n自作() As Decimal
        Dim n小作() As Decimal

        Dim n田合計 As Decimal
        Dim n畑合計 As Decimal
        Dim n樹園地 As Decimal
        Dim n採草地 As Decimal

        Dim 都道府県名 As String
        Dim 会長名 As String
        Dim 会長肩書 As String
        Dim 筆数 As Integer
        Dim View As DataView
    End Structure

    Public Overrides Sub Execute()
        Me.DataInit()
        Value = 33

        Me.MakeXMLFile()
        Value = 90
    End Sub

    Public Sub New(ByVal sKey As String)
        MyBase.New(sKey, "\耕作証明願.xml")
        If n発行番号 > 0 Then
            Me.Dialog.StartProc(True, True)

            If Me.Dialog._objException Is Nothing = False Then
                If Me.Dialog._objException.Message = "Cancel" Then
                    MsgBox("処理を中止しました。　", , "処理中止")
                Else
                End If
            Else
                Me.SaveAndOpen(ExcelViewMode.Preview)
                SysAD.DB(sLRDB).DBProperty("耕作証明番号") = n発行番号
            End If
        End If
    End Sub

    Public Sub DataInit()
        Try
            Dim sData As Object = Nothing
            Dim sSQL As String = ""
            Dim cnt As Integer = 0
            Dim n自作(5) As Integer
            Dim n小作(5) As Integer
            Dim nSK As String = ""

            With mvarData
                .都道府県名 = SysAD.DB(sLRDB).DBProperty("都道府県名")
                .会長名 = SysAD.DB(sLRDB).DBProperty("会長名")
                .会長肩書 = IIf(Val(SysAD.DB(sLRDB).DBProperty("会長代理")), "農業委員会会長代理", "農業委員会会長")
            End With


            With mvarData
                .n自筆 = New Integer() {0, 0, 0, 0, 0, 0}
                .n小筆 = New Integer() {0, 0, 0, 0, 0, 0}
                .n自作 = New Decimal() {0, 0, 0, 0, 0, 0}
                .n小作 = New Decimal() {0, 0, 0, 0, 0, 0}
                SysAD.DB(sLRDB).ExecuteSQL("UPDATE V_農地 SET V_農地.農地状況 = 0 WHERE (((V_農地.農地状況) Is Null));")


                Select Case HIMTools.CommonFunc.GetKeyHead(mvarKey)
                    Case "農家"
                        Me.世帯ID = HIMTools.CommonFunc.GetKeyCode(mvarKey)
                        Dim s所有世帯 As String = ""
                        Select Case 出力条件.管理者の影響
                            Case C耕作証明条件.enum管理人.管理人を考慮しない
                                's所有世帯 = String.Format("[V_農地].[耕作世帯ID]={0}", Me.世帯ID)
                                s所有世帯 = String.Format("IIf([自小作別]<>0,[借受世帯ID],[所有世帯ID])={0}", Me.世帯ID)
                            Case C耕作証明条件.enum管理人.管理人を考慮する
                                s所有世帯 = String.Format("[V_農地].[耕作世帯ID]={0}", Me.世帯ID)
                        End Select
                        Select Case 出力条件.市外農地を含む
                            Case C耕作証明条件.enum市外農地.含む

                            Case C耕作証明条件.enum市外農地.含まない
                                s所有世帯 = s所有世帯 & " AND [V_農地].[大字ID] > 0"
                        End Select

                        Dim pTBL世帯 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:世帯Info].世帯主ID,[D:個人Info].* FROM [D:世帯Info] LEFT JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE [D:世帯Info].[ID]=" & Me.世帯ID)

                        If pTBL世帯.Rows.Count > 0 Then
                            Me.個人ID = Val(pTBL世帯.Rows(0).Item("世帯主ID").ToString)

                            sData = Me.申請者名 & ";" & Me.申請者住所 & vbCrLf
                            sSQL = "SELECT False AS [有効農地], V_農地.土地所在, V_農地.田面積, V_農地.畑面積, V_農地.樹園地, V_農地.採草放牧面積, [D:個人Info].氏名 AS 所有者,[D:個人Info_1].氏名 AS 耕作者, V_農地.自小作別, V_農地.小作形態, V_農地.小作開始年月日, V_農地.小作終了年月日, V_農地.小作料 ,V_農地.小作料単位 FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.耕作者ID = [D:個人Info_1].ID WHERE (((V_農地.農地状況) < 20) And " & s所有世帯 & ") ORDER BY V_農地.土地所在;"
                        End If
                    Case "個人"
                        Dim s所有者 As String = ""
                        Me.個人ID = HIMTools.CommonFunc.GetKeyCode(mvarKey)

                        Select Case 出力条件.管理者の影響
                            Case C耕作証明条件.enum管理人.管理人を考慮しない
                                s所有者 = String.Format("IIf([自小作別]<>0,[借受人ID],[所有者ID])={0}", Me.個人ID)
                            Case C耕作証明条件.enum管理人.管理人を考慮する
                                s所有者 = String.Format("[V_農地].[耕作者ID]={0}", Me.個人ID)
                        End Select
                        Select Case 出力条件.市外農地を含む
                            Case C耕作証明条件.enum市外農地.含む

                            Case C耕作証明条件.enum市外農地.含まない
                                s所有者 = s所有者 & " AND [V_農地].[大字ID] > 0"
                        End Select

                        Dim pTBL個人 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [ID]=" & Me.個人ID)

                        If pTBL個人.Rows.Count > 0 Then
                            Me.世帯ID = Val(pTBL個人.Rows(0).Item("世帯ID").ToString)

                            sData = Me.申請者名 & ";" & Me.申請者住所 & vbCrLf
                            sSQL = "SELECT False AS [有効農地], V_農地.自小作別, V_農地.土地所在, V_農地.田面積, V_農地.畑面積, V_農地.樹園地, V_農地.採草放牧面積, [D:個人Info].氏名 AS 所有者,[D:個人Info_1].氏名 AS 耕作者, V_農地.自小作別, V_農地.小作形態, V_農地.小作開始年月日, V_農地.小作終了年月日, V_農地.小作料 ,V_農地.小作料単位 FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) INNER JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.耕作者ID = [D:個人Info_1].ID WHERE [V_農地].[農地状況]<20 AND " & s所有者 & " ORDER BY V_農地.土地所在;"
                        End If
                End Select

                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(sSQL)
                pTBL.Columns.Add("耕作面積", GetType(Decimal))
                pTBL.Columns.Add("現況地目", GetType(String))

                .筆数 = 0

                For Each pRow As DataRow In pTBL.Rows
                    cnt += 1
                    If Val(pRow.Item("田面積").ToString) > 0 Then
                        sData = sData & pRow.Item("土地所在").ToString
                        sData = sData & ";田;" & FormatNumber(pRow.Item("田面積")) & ";" & pRow.Item("所有者") & ";" & pRow.Item("耕作者") & ";" & pRow.Item("自小作別") & ";" & pRow.Item("小作形態") & ";" & pRow.Item("小作開始年月日") & ";" & pRow.Item("小作終了年月日") & ";" & pRow.Item("小作料") & ";" & pRow.Item("小作料単位") & vbCr
                        .n田合計 = .n田合計 + Val(pRow.Item("田面積").ToString)
                        pRow.Item("現況地目") = "田"
                        pRow.Item("耕作面積") = Val(pRow.Item("田面積"))

                        If Val(pRow.Item("自小作別").ToString) = 0 Then
                            .n自筆(1) = .n自筆(1) + 1
                            .n自作(1) = .n自作(1) + pRow.Item("田面積")
                            n自作(0) = n自作(0) + 1
                            n自作(1) = n自作(1) + pRow.Item("田面積")
                        Else
                            .n小筆(1) = .n小筆(1) + 1
                            .n小作(1) = .n小作(1) + pRow.Item("田面積")
                            n小作(0) = n小作(0) + 1
                            n小作(1) = n小作(1) + pRow.Item("田面積")
                        End If
                        pRow.Item("有効農地") = True
                        .筆数 += 1
                    ElseIf Val(pRow.Item("畑面積").ToString) > 0 Then
                        sData = sData & pRow.Item("土地所在").ToString
                        sData = sData & ";畑;" & FormatNumber(pRow.Item("畑面積")) & ";" & pRow.Item("所有者") & ";" & pRow.Item("耕作者") & ";" & pRow.Item("自小作別") & ";" & pRow.Item("小作形態") & ";" & pRow.Item("小作開始年月日") & ";" & pRow.Item("小作終了年月日") & ";" & pRow.Item("小作料") & ";" & pRow.Item("小作料単位") & vbCr
                        .n畑合計 = .n畑合計 + pRow.Item("畑面積")
                        pRow.Item("現況地目") = "畑"
                        pRow.Item("耕作面積") = pRow.Item("畑面積")

                        If Val(pRow.Item("自小作別").ToString) = 0 Then
                            .n自筆(2) = .n自筆(2) + 1
                            .n自作(2) = .n自作(2) + pRow.Item("畑面積")
                            n自作(2) = n自作(2) + 1
                            n自作(3) = n自作(3) + pRow.Item("畑面積")
                        Else
                            .n小筆(2) = .n小筆(2) + 1
                            .n小作(2) = .n小作(2) + pRow.Item("畑面積")
                            n小作(2) = n小作(2) + 1
                            n小作(3) = n小作(3) + pRow.Item("畑面積")
                        End If
                        pRow.Item("有効農地") = True
                        .筆数 += 1
                    ElseIf Val(pRow.Item("樹園地").ToString) > 0 Then
                        sData = sData & pRow.Item("土地所在").ToString
                        sData = sData & ";樹園地;" & FormatNumber(pRow.Item("樹園地")) & ";" & pRow.Item("所有者") & ";" & pRow.Item("耕作者") & ";" & pRow.Item("自小作別") & ";" & pRow.Item("小作形態") & ";" & pRow.Item("小作開始年月日") & ";" & pRow.Item("小作終了年月日") & ";" & pRow.Item("小作料", 0) & ";" & pRow.Item("小作料単位") & vbCr
                        pRow.Item("現況地目") = "樹園地"
                        pRow.Item("耕作面積") = pRow.Item("樹園地")

                        .n樹園地 = .n樹園地 + pRow.Item("樹園地")
                        If pRow.Item("自小作別") = 0 Then
                            .n自筆(3) = .n自筆(3) + 1
                            .n自作(3) = .n自作(3) + pRow.Item("樹園地")
                        Else
                            .n小筆(3) = .n小筆(3) + 1
                            .n小作(3) = .n小作(3) + pRow.Item("樹園地")
                        End If
                        pRow.Item("有効農地") = True
                        .筆数 += 1
                    ElseIf Val(pRow.Item("採草放牧面積").ToString) > 0 Then
                        sData = sData & pRow.Item("土地所在").ToString

                        pRow.Item("現況地目") = "採草放牧地"
                        pRow.Item("耕作面積") = pRow.Item("採草放牧面積")
                        pRow.Item("有効農地") = True
                        .筆数 += 1
                    Else
                        cnt = cnt - 1
                    End If
                Next

                For i = 0 To 5
                    nSK = nSK & ";" & n自作(i) & ";" & n小作(i)
                Next
                For i = 1 To 4
                    .n自筆(5) = .n自筆(5) + .n自筆(i)
                    .n自作(5) = .n自作(5) + .n自作(i)
                    .n小筆(5) = .n小筆(5) + .n小筆(i)
                    .n小作(5) = .n小作(5) + .n小作(i)
                Next

                sData = sData & vbCrLf & .n田合計 & ";" & .n畑合計 & ";" & cnt & nSK

                .View = New DataView(pTBL, "[有効農地]=True", "土地所在", DataViewRowState.CurrentRows)
                Print耕作証明願(pTBL)
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub Print耕作証明願(ByRef pTBL As DataTable)
        Dim nCount As Decimal = pTBL.Rows.Count
        Dim nPage As Integer = Math.Floor(nCount / 10) + 1
        If nCount > (nPage * 10) - 6 Then
            nPage = nPage + 1
        End If

        Select Case nPage
            Case 1 : mvarXML.WorkBook.WorkSheets.Remove("Page2")
            Case 2
            Case Else
                For i = 3 To nPage
                    mvarXML.WorkBook.WorkSheets.CopySheet("Page2", "Page" & i)
                Next
        End Select

        For i = 3 To nPage
            For n = 5 To 14
                Dim pReplaceSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = mvarXML.WorkBook.WorkSheets.Items("Page" & i)
                With pReplaceSheet
                    .ValueReplace("{NowPage}", i)
                    .ValueReplace("{土地の所在" & n & "}", "{土地の所在" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{現況地目" & n & "}", "{現況地目" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{耕作面積" & n & "}", "{耕作面積" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{所有者" & n & "}", "{所有者" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{自小作別" & n & "}", "{自小作別" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{借受人" & n & "}", "{借受人" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{小作形態" & n & "}", "{小作形態" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{開始年月日" & n & "}", "{開始年月日" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{終了年月日" & n & "}", "{終了年月日" & n + 10 * (i - 2) & "}")
                    .ValueReplace("{小作料" & n & "}", "{小作料" & n + 10 * (i - 2) & "}")
                End With
            Next
        Next

        For Each pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet In mvarXML.WorkBook.WorkSheets.Items.Values
            With pSheet
                .ValueReplace("{NowPage}", 2)
                .ValueReplace("{MaxPage}", nPage)
                .ValueReplace("{発行日}", 和暦Format(dt発行日))
                .ValueReplace("{発行番号}", n発行番号)
                .ValueReplace("{住所}", Me.申請者住所)
                .ValueReplace("{氏名}", Me.申請者名)
                .ValueReplace("{市町村名}", SysAD.市町村.市町村名)
                .ValueReplace("{会長名}", mvarData.会長名)

                For i As Integer = 1 To 5
                    Dim s地目別 As String = New String() {"", "田", "畑", "樹", "採", "計"}(i)

                    .ValueReplace("{自" & s地目別 & "筆数}", mvarData.n自筆(i))
                    .ValueReplace("{小" & s地目別 & "筆数}", mvarData.n小筆(i))
                    .ValueReplace("{計" & s地目別 & "筆数}", mvarData.n自筆(i) + mvarData.n小筆(i))
                    .ValueReplace("{自" & s地目別 & "面積}", NumToString(mvarData.n自作(i)))
                    .ValueReplace("{小" & s地目別 & "面積}", NumToString(mvarData.n小作(i)))
                    .ValueReplace("{計" & s地目別 & "面積}", NumToString(mvarData.n自作(i) + mvarData.n小作(i)))
                Next

                If mvarData.筆数 > 0 Then
                    For n As Integer = 1 To mvarData.筆数
                        Dim pRow As New HimTools2012.Data.DataRowPlus(mvarData.View(n - 1).Row)

                        .ValueReplace("{土地の所在" & n & "}", pRow.Item("土地所在", ""))
                        .ValueReplace("{現況地目" & n & "}", pRow.Item("現況地目", ""))
                        .ValueReplace("{耕作面積" & n & "}", NumToString(pRow.Item("耕作面積", 0)))
                        .ValueReplace("{所有者" & n & "}", pRow.Item("所有者", ""))
                        .ValueReplace("{自小作別" & n & "}", pRow.Choose("自小作別", {"自", "小", "農"}))

                        If pRow.Item("自小作別", 0) > 0 Then
                            .ValueReplace("{借受人" & n & "}", pRow.Item("耕作者").ToString)

                            Select Case Val(pRow.Item("小作形態").ToString)
                                Case 0 : .ValueReplace("{小作形態" & n & "}", "-")
                                Case 1 : .ValueReplace("{小作形態" & n & "}", "賃")
                                Case 2 : .ValueReplace("{小作形態" & n & "}", "使")
                            End Select

                            If Not IsDBNull(pRow.Item("小作開始年月日")) AndAlso IsDate(pRow.Item("小作開始年月日")) AndAlso pRow.Item("小作開始年月日") > #1/1/1902# Then
                                .ValueReplace("{開始年月日" & n & "}", 和暦Format(pRow.Item("小作開始年月日"), "gggyy年MM月dd日"))
                            Else
                                .ValueReplace("{開始年月日" & n & "}", "")
                            End If
                            If Not IsDBNull(pRow.Item("小作終了年月日")) AndAlso IsDate(pRow.Item("小作終了年月日")) AndAlso pRow.Item("小作終了年月日") > #1/1/1902# Then
                                .ValueReplace("{終了年月日" & n & "}", 和暦Format(pRow.Item("小作終了年月日"), "gggyy年MM月dd日"))
                            Else
                                .ValueReplace("{終了年月日" & n & "}", "")
                            End If
                            If Val(pRow.Item("小作料").ToString) Then
                                .ValueReplace("{小作料" & n & "}", pRow.Item("小作料").ToString & pRow.Item("小作料単位").ToString)
                            Else
                                .ValueReplace("{小作料" & n & "}", "")
                            End If
                        Else
                            .ValueReplace("{借受人" & n & "}", "")
                            .ValueReplace("{小作形態" & n & "}", "")
                            .ValueReplace("{開始年月日" & n & "}", "")
                            .ValueReplace("{終了年月日" & n & "}", "")
                            .ValueReplace("{小作料" & n & "}", "")
                        End If
                    Next
                End If

                For n = mvarData.筆数 + 1 To nPage * 10 - 6
                    .ValueReplace("{土地の所在" & n & "}", "")
                    .ValueReplace("{現況地目" & n & "}", "")
                    .ValueReplace("{耕作面積" & n & "}", "")
                    .ValueReplace("{所有者" & n & "}", "")
                    .ValueReplace("{自小作別" & n & "}", "")
                    .ValueReplace("{借受人" & n & "}", "")
                    .ValueReplace("{小作形態" & n & "}", "")
                    .ValueReplace("{開始年月日" & n & "}", "")
                    .ValueReplace("{終了年月日" & n & "}", "")
                    .ValueReplace("{小作料" & n & "}", "")
                Next
            End With
        Next

    End Sub

    Public Sub MakeXMLFile()
        Maximum = 100
        Value = 33
        Message = "エクセルファイル作成中.."
        Application.DoEvents()
        If _Cancel Then
            Throw New Exception("Cancel")
        End If
        Value = 90
    End Sub
End Class
