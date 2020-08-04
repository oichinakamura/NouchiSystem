Imports System.ComponentModel
Imports HimTools2012
Imports HimTools2012.CommonFunc

Module mod農地異動関連
    '/****************************/
    '/* 職権による農地の移動処理 */
    '/****************************/
    Public Sub sub異動from職権所有権移転(ByVal sLandList As String, ByVal n新世帯ID As Decimal, ByVal n新所有者ID As Decimal, Optional ByVal n異動事由 As Long = 99997, Optional ByVal s異動内容 As String = "へ職権による所有権移転")
        Dim pInpData As New 異動日
        With New frmInputData("異動日", "異動日時を入力してください", pInpData)
            If .ShowDialog = DialogResult.OK Then
                For Each sKey As String In Split(Replace(sLandList, ";", ","), ",")
                    Dim pRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(GetKeyCode(sKey))
                    Dim pDate As Date = pInpData.異動日

                    If pRow IsNot Nothing Then
                        Dim p農地 As CObj農地 = ObjectMan.GetObjectDB("農地", pRow, GetType(CObj農地))

                        Make農地履歴(pRow.Item("ID"), pDate, pDate, n異動事由, enum法令.職権異動, s異動内容, , 0)
                        p農地.ValueChange("所有世帯ID", n新世帯ID)
                        p農地.ValueChange("所有者ID", n新所有者ID)
                        p農地.ValueChange("管理世帯ID", 0)
                        p農地.ValueChange("管理者ID", 0)
                        p農地.ValueChange("先行異動", True)
                        p農地.ValueChange("先行異動日", pDate)
                        p農地.SaveMyself()
                    End If
                Next
            End If
        End With
    End Sub

    Public Class 異動日
        Inherits CInputDate

        <Category("異動日")>
        Public Property 異動日 As DateTime = Now.Date


        <Browsable(False)>
        Public Overrides ReadOnly Property DataValidate As Boolean
            Get
                Return True
            End Get
        End Property
    End Class

    Public Sub 農地換地処理(ByVal p農地 As CObj農地, ByVal sType As String, ByVal pParams As String())
        If SysAD.page農家世帯.TabPageContainKey("換地処理") Then
            MsgBox("既に換地処理が実行されています。処理が完了するか中断してください", MsgBoxStyle.Critical)
        Else
            If pParams.Length > 0 Then
                SysAD.page農家世帯.中央Tab.AddPage(New CTabPage換地処理(pParams, 0, sType))
            Else
                SysAD.page農家世帯.中央Tab.AddPage(New CTabPage換地処理({p農地.Key.ID.ToString}, 0, sType))
            End If
        End If
    End Sub

    Public Sub 農地履歴手動追加(ByVal p農地 As CObj農地)
        If MsgBox("履歴を追加しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地履歴]([LID],[内容],[更新日],[入力日]) VALUES(" & p農地.ID & ",'新しい履歴',Now,Now)")
            SysAD.page農家世帯.土地履歴リスト.検索開始("[LID]=" & p農地.ID, "[LID]=" & p農地.ID)
        End If
    End Sub

    Public Sub sub異動所有権移転(ByVal p異動日 As Date, ByVal o通知発行日 As Object, ByRef p申請Row As HimTools2012.Data.DataRowPlus, ByVal sFolder As String, ByVal b無異動 As Boolean)
        If Not IsDBNull(o通知発行日) Then
            Dim dt異動日 As DateTime = HimTools2012.DateFunctions.NullCast日付(p異動日.ToString, Now.Date)
            Dim dt通知発行日 As DateTime = HimTools2012.DateFunctions.NullCast日付(o通知発行日.ToString, Now.Date)
            Dim sError As String = ""

            Dim s法令番号 As String = ""
            Dim s内容 As String = ""

            If p申請Row IsNot Nothing Then
                Dim p申請 As New CObj申請(p申請Row.Body, False)
                Dim n異動事由 As Integer = 10000 + p申請.法令
                Dim sOutput As String = ""

                Select Case p申請.法令
                    Case enum法令.農地法3条所有権 : sOutput = "\農地法第3条許可書（所有権の移転）.xml"
                        s内容 = "{0}から{1}へ所有権移転"
                    Case enum法令.農地法3条の3第1項
                        s内容 = "{0}から{1}へ農地法第3条の3第1項にて所有権移転"
                    Case Else
                        CasePrint(p申請.法令)
                End Select

                Dim St As String = p申請Row.Item("農地リスト").ToString
                Dim Cn() As String = Split(St, ";")
                Dim p許可書XML As New C許可書XML(sOutput)

                Dim nArea As Decimal = 0
                Dim nCount As Integer = 0

                Dim okFlag As Boolean = False
                For K As Integer = LBound(Cn) To UBound(Cn)
                    If Len(Cn(K)) = 0 Then
                    ElseIf GetKeyHead(Cn(K)) = "農地" Then
                        nCount += 1
                        Dim TID As Long = GetKeyCode(Cn(K))

                        Dim pRow農地 As DataRow = App農地基本台帳.TBL農地.FindRowByID(TID)
                        If pRow農地 IsNot Nothing Then
                            Dim p農地 As New CObj農地(pRow農地, False)

                            If p農地.自小作別 = 1 Then
                                If okFlag = False Then
                                    If MsgBox("現在「小作」になっている農地が許可対象に含まれていますがよろしいですか？", vbYesNo) = vbYes Then
                                        okFlag = True
                                    Else
                                        Exit Sub
                                    End If
                                End If
                            End If

                            If Not b無異動 AndAlso p農地 IsNot Nothing Then
                                p農地.ValueChange("所有者ID", p申請.GetItem("申請者B", 0))
                                p農地.ValueChange("所有世帯ID", p申請.GetItem("申請世帯B", 0))
                                p農地.ValueChange("管理者ID", 0)
                                p農地.ValueChange("管理世帯ID", 0)
                                p農地.ValueChange("先行異動", True)
                                p農地.ValueChange("先行異動日", dt異動日)
                                p農地.ValueChange("公告年月日", dt異動日)

                                p農地.SaveMyself()
                                '20161007 「p申請.法令」を追加
                                Make農地履歴(p農地.ID, Now.Date, dt異動日, n異動事由, p申請.法令, String.Format(s内容, p申請.氏名A, p申請.氏名B), p申請.申請者A, p申請.申請者B, p申請.ID)
                            Else
                                MsgBox("001:[ID]" & TID & "の農地が見つかりません。")
                            End If
                            If sOutput.Length > 0 Then
                                With p許可書XML
                                    .XMLReplaceWithNo("土地の所在", nCount, 2, p農地.土地所在)
                                    .XMLReplaceWithNo("土地の所在B", nCount, 2, Replace地番(p農地.土地所在))
                                    .XMLReplaceWithNo("登記地目", nCount, 2, p農地.Row.Body.Item("登記簿地目名").ToString)
                                    .XMLReplaceWithNo("現況地目", nCount, 2, p農地.Row.Body.Item("現況地目名").ToString)
                                    .XMLReplaceWithNo("面積", nCount, 2, AreaConv(p農地.GetItem("登記簿面積").ToString))
                                    nArea += Val(p農地.GetItem("登記簿面積").ToString)
                                    .XMLReplaceWithNo("備考", nCount, 2, "")
                                End With

                            End If
                        Else
                            MsgBox("002:[ID]" & TID & "の農地が見つかりません。")
                        End If

                    End If
                Next

                If sOutput.Length > 0 Then
                    With p許可書XML
                        .XMLReplace("発行年月日", 和暦Format(dt通知発行日))
                        If IsDate(dt異動日) Then
                            .XMLReplace("許可月", Month(dt異動日))
                        Else
                            .XMLReplace("許可月", "　")
                        End If
                        .XMLReplace("郵便番号A", "〒" & p申請.GetItem("申請者情報郵便番号A").ToString)
                        .XMLReplace("住所A", p申請.GetItem("住所A").ToString)
                        .XMLReplace("氏名A", p申請.GetItem("氏名A").ToString)
                        .XMLReplace("郵便番号B", "〒" & p申請.GetItem("申請者情報郵便番号B").ToString)
                        .XMLReplace("住所B", p申請.GetItem("住所B").ToString)
                        .XMLReplace("氏名B", p申請.GetItem("氏名B").ToString)

                        .XMLReplace("会長名", SysAD.DB(sLRDB).DBProperty("会長名").ToString)
                        .XMLReplace("市町村名", SysAD.DB(sLRDB).DBProperty("市町村名").ToString)
                        .XMLReplace("許可年月日", 和暦Format(dt異動日))
                        .XMLReplace("年度", Strings.Mid(和暦Format(dt異動日), 3, 2))

                        .XMLReplace("総会番号", p申請.GetItem("受付番号", 0))
                        .XMLReplace("受付番号", p申請.GetItem("受付番号", 0))

                        .XMLReplace("許可番号", p申請.GetItem("許可番号", ""))
                        .XMLReplace("申請年月日", 和暦Format(p申請.GetItem("受付年月日")))
                        .XMLReplace("受付年月日", 和暦Format(p申請.GetItem("受付年月日")))
                        Select Case p申請.GetItem("所有権移転の種類", 0)
                            Case 0 : .XMLReplace("設定種類", "")
                            Case 1 : .XMLReplace("設定種類", "売買")
                            Case 2 : .XMLReplace("設定種類", "贈与")
                            Case 3 : .XMLReplace("設定種類", "交換")
                            Case 4 : .XMLReplace("設定種類", "賃貸借")
                            Case 5 : .XMLReplace("設定種類", "使用貸借")
                        End Select


                        If nCount = 9 Then
                            .XMLReplaceWithNo("土地の所在B", 10, 2, "")
                            .XMLReplaceWithNo("土地の所在", 10, 2, "")
                            .XMLReplaceWithNo("登記地目", 10, 2, "")
                            .XMLReplaceWithNo("現況地目", 10, 2, "合 計")
                            .XMLReplaceWithNo("面積", 10, 2, AreaConv(nArea))
                            .XMLReplaceWithNo("備考", 10, 2, "")

                        ElseIf nCount < 9 Then
                            .XMLReplaceWithNo("土地の所在B", nCount + 1, 2, "以 下 余 白")
                            .XMLReplaceWithNo("土地の所在", nCount + 1, 2, "以 下 余 白")

                            .XMLReplaceWithNo("土地の所在", 10, 2, "")
                            .XMLReplaceWithNo("登記地目", 10, 2, "")
                            .XMLReplaceWithNo("現況地目", 10, 2, "合 計")
                            .XMLReplaceWithNo("面積", 10, 2, AreaConv(nArea))
                            .XMLReplaceWithNo("備考", 10, 2, "")
                        Else
                            .XMLReplaceWithNo("土地の所在B", nCount + 1, 2, "以 下 余 白")
                            .XMLReplaceWithNo("土地の所在", nCount + 1, 2, "以 下 余 白")

                            .XMLReplaceWithNo("土地の所在B", nCount + 2, 2, "")
                            .XMLReplaceWithNo("土地の所在", nCount + 2, 2, "")
                            .XMLReplaceWithNo("登記地目", nCount + 2, 2, "")
                            .XMLReplaceWithNo("現況地目", nCount + 2, 2, "合 計")
                            .XMLReplaceWithNo("面積", nCount + 2, 2, AreaConv(nArea))
                            .XMLReplaceWithNo("備考", nCount + 2, 2, "")
                        End If

                        For i As Integer = nCount + 1 To 60
                            .XMLReplaceWithNo("土地の所在B", i, 2, "")
                            .XMLReplaceWithNo("土地の所在", i, 2, "")
                            .XMLReplaceWithNo("登記地目", i, 2, "")
                            .XMLReplaceWithNo("現況地目", i, 2, "")
                            .XMLReplaceWithNo("面積", i, 2, "")
                            .XMLReplaceWithNo("備考", i, 2, "")
                        Next

                        If Val(p申請.GetItem("許可番号", "").ToString) <> 0 Then
                            .SaveAndOpen(sFolder, p申請.GetItem("許可番号", ""), p申請.GetItem("名称").ToString)
                        End If

                    End With
                End If

                p申請.SetItem("状態", 2)
                p申請.SetItem("許可年月日", p異動日)
                p申請.SaveMyself()
            End If
        End If
    End Sub

    Public Sub sub農地転用事業計画(ByVal p異動日 As Date, ByVal o通知発行日 As Object, ByRef p申請Row As HimTools2012.Data.DataRowPlus, ByVal sFolder As String, ByVal b無異動 As Boolean)
        If Not IsDBNull(o通知発行日) Then
            Dim dt異動日 As DateTime = HimTools2012.DateFunctions.NullCast日付(p異動日.ToString, Now.Date)
            Dim dt通知発行日 As DateTime = HimTools2012.DateFunctions.NullCast日付(o通知発行日.ToString, Now.Date)
            Dim sError As String = ""

            If p申請Row IsNot Nothing Then
                Dim p申請 As New CObj申請(p申請Row.Body, False)
                Dim sOutput As String = ""

                sOutput = "\農地転用事業計画変更通知書.xml"

                Dim St As String = p申請Row.Item("農地リスト").ToString
                Dim Cn() As String = Split(St, ";")
                Dim p許可書XML As New C許可書XML(sOutput)

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

                Dim nArea As Decimal = 0
                Dim nCount As Integer = 0

                Dim n前Area As Decimal = 0
                Dim n前Count As Integer = 0

                Dim Ar異動前農地 As Object = Split(p申請Row.Item("予備1"), vbCrLf)
                For n As Integer = 0 To UBound(Ar異動前農地)
                    n前Count += 1

                    Dim Ar As Object = Split(Ar異動前農地(n), ";")

                    If sOutput.Length > 0 Then
                        With p許可書XML
                            .XMLReplaceWithNo("前土地の所在", n前Count, 2, Ar(0))
                            .XMLReplaceWithNo("前土地の所在B", n前Count, 2, Replace地番(Ar(0)))
                            .XMLReplaceWithNo("前登記地目", n前Count, 2, Ar(1))
                            .XMLReplaceWithNo("前現況地目", n前Count, 2, Ar(2))
                            .XMLReplaceWithNo("前面積", n前Count, 2, AreaConv(Val(Ar(3).ToString)))
                            n前Area += Val(Ar(3).ToString)
                        End With
                    End If
                Next


                For K As Integer = LBound(Cn) To UBound(Cn)
                    If Len(Cn(K)) = 0 Then
                    ElseIf GetKeyHead(Cn(K)) = "転用農地" Then
                        nCount += 1
                        Dim TID As Long = GetKeyCode(Cn(K))

                        Dim pRow転用農地 As DataRow = App農地基本台帳.TBL転用農地.FindRowByID(TID)
                        If pRow転用農地 IsNot Nothing Then
                            Dim p転用農地 As New CObj転用農地(pRow転用農地, False)
                            Dim p復活農地 As New CObj転用農地(pRow転用農地, False)
                            Dim p筆別Row As DataRow = Nothing

                            If Not b無異動 AndAlso p転用農地 IsNot Nothing Then
                                With p転用農地
                                    If pTBL IsNot Nothing AndAlso pTBL.PrimaryKey.Length > 0 Then
                                        p筆別Row = pTBL.Rows.Find("転用農地." & p転用農地.ID)

                                        If Not p筆別Row Is Nothing AndAlso Not IsDBNull(p筆別Row.Item("申請部分面積")) AndAlso p筆別Row.Item("申請部分面積") > 0 Then
                                            If p転用農地.GetIntegerValue("一部現況") = 0 Then
                                                If p転用農地.GetDecimalValue("登記簿面積") > p筆別Row.Item("申請部分面積") Then
                                                    p転用農地.ValueChange("実面積", p転用農地.GetDecimalValue("登記簿面積") - p筆別Row.Item("申請部分面積"))
                                                    If p転用農地.GetDecimalValue("田面積") > 0 Then
                                                        p転用農地.ValueChange("田面積", p転用農地.GetDecimalValue("実面積"))
                                                    ElseIf p転用農地.GetDecimalValue("畑面積") > 0 Then
                                                        p転用農地.ValueChange("畑面積", p転用農地.GetDecimalValue("畑面積"))
                                                    End If
                                                    p転用農地.SetIntegerValue("一部現況", 1)

                                                    p転用農地.SaveMyself()


                                                    p復活農地 = p復活農地.CopyObject()
                                                    p復活農地.SetIntegerValue("一部現況", 2)
                                                    p復活農地.ValueChange("実面積", p筆別Row.Item("申請部分面積"))
                                                    If p復活農地.GetDecimalValue("田面積") > 0 Then
                                                        p復活農地.ValueChange("田面積", p復活農地.GetDecimalValue("実面積"))
                                                    ElseIf p復活農地.GetDecimalValue("畑面積") > 0 Then
                                                        p復活農地.ValueChange("畑面積", p復活農地.GetDecimalValue("畑面積"))
                                                    End If

                                                    If MsgBox("申請外農地を農地へ戻しますか", vbYesNo) = vbYes Then
                                                        p復活農地.Sub転用農地の復活(p復活農地)
                                                    End If
                                                Else
                                                    MsgBox(String.Format("申請部分面積が登記簿面積を越えています。分割はできません。"), MsgBoxStyle.Critical)
                                                End If
                                            Else
                                                MsgBox("対象農地は既に分割されています。再分割はできません。", MsgBoxStyle.Critical)
                                            End If

                                        End If
                                    End If



                                    Make農地履歴(.ID, Now.Date, dt異動日, 0, enum法令.事業計画変更, "事業計画変更", p申請.申請者C, p申請.申請者A, p申請.ID)
                                End With
                            End If
                            If sOutput.Length > 0 Then
                                With p許可書XML
                                    .XMLReplaceWithNo("土地の所在", nCount, 2, p転用農地.土地所在)
                                    .XMLReplaceWithNo("土地の所在B", nCount, 2, Replace地番(p転用農地.土地所在))
                                    .XMLReplaceWithNo("登記地目", nCount, 2, p転用農地.Row.Body.Item("登記簿地目名").ToString)
                                    .XMLReplaceWithNo("現況地目", nCount, 2, p転用農地.Row.Body.Item("現況地目名").ToString)
                                    .XMLReplaceWithNo("面積", nCount, 2, AreaConv(p転用農地.GetItem("登記簿面積").ToString))
                                    nArea += Val(p転用農地.GetItem("登記簿面積").ToString)
                                End With
                            End If
                        End If
                    End If
                Next

                If sOutput.Length > 0 Then
                    With p許可書XML
                        .XMLReplace("会長名", SysAD.DB(sLRDB).DBProperty("会長名").ToString)
                        .XMLReplace("市町村名", SysAD.DB(sLRDB).DBProperty("市町村名").ToString)
                        .XMLReplace("許可年月日", 和暦Format(dt異動日))
                        .XMLReplace("許可番号", p申請.GetItem("許可番号", ""))

                        .XMLReplace("申請者Ａ氏名", p申請.GetItem("氏名A").ToString)
                        .XMLReplace("申請者Ａ住所", p申請.GetItem("住所A").ToString)
                        .XMLReplace("変更前転用目的", p申請.GetItem("予備2").ToString)

                        .XMLReplace("申請者Ｃ氏名", p申請.GetItem("氏名C").ToString)
                        .XMLReplace("申請者Ｃ住所", p申請.GetItem("住所C").ToString)
                        .XMLReplace("変更後転用目的", p申請.GetItem("申請理由A").ToString)

                        If nCount = 4 Then
                            .XMLReplaceWithNo("土地の所在B", 5, 2, "")
                            .XMLReplaceWithNo("土地の所在", 5, 2, "")
                            .XMLReplaceWithNo("登記地目", 5, 2, "")
                            .XMLReplaceWithNo("現況地目", 5, 2, "合 計")
                            .XMLReplaceWithNo("面積", 5, 2, AreaConv(nArea))

                        ElseIf nCount < 4 Then
                            .XMLReplaceWithNo("土地の所在B", nCount + 1, 2, "以 下 余 白")
                            .XMLReplaceWithNo("土地の所在", nCount + 1, 2, "以 下 余 白")

                            .XMLReplaceWithNo("土地の所在", 5, 2, "")
                            .XMLReplaceWithNo("登記地目", 5, 2, "")
                            .XMLReplaceWithNo("現況地目", 5, 2, "合 計")
                            .XMLReplaceWithNo("面積", 5, 2, AreaConv(nArea))
                        Else
                            .XMLReplaceWithNo("土地の所在B", nCount + 1, 2, "以 下 余 白")
                            .XMLReplaceWithNo("土地の所在", nCount + 1, 2, "以 下 余 白")

                            .XMLReplaceWithNo("土地の所在B", nCount + 2, 2, "")
                            .XMLReplaceWithNo("土地の所在", nCount + 2, 2, "")
                            .XMLReplaceWithNo("登記地目", nCount + 2, 2, "")
                            .XMLReplaceWithNo("現況地目", nCount + 2, 2, "合 計")
                            .XMLReplaceWithNo("面積", nCount + 2, 2, AreaConv(nArea))
                        End If

                        For i As Integer = nCount + 1 To 5
                            .XMLReplaceWithNo("土地の所在B", i, 2, "")
                            .XMLReplaceWithNo("土地の所在", i, 2, "")
                            .XMLReplaceWithNo("登記地目", i, 2, "")
                            .XMLReplaceWithNo("現況地目", i, 2, "")
                            .XMLReplaceWithNo("面積", i, 2, "")
                        Next

                        '/***異動前***/
                        If n前Count = 4 Then
                            .XMLReplaceWithNo("前土地の所在B", 5, 2, "")
                            .XMLReplaceWithNo("前土地の所在", 5, 2, "")
                            .XMLReplaceWithNo("前登記地目", 5, 2, "")
                            .XMLReplaceWithNo("前現況地目", 5, 2, "合 計")
                            .XMLReplaceWithNo("前面積", 5, 2, AreaConv(n前Area))

                        ElseIf n前Count < 4 Then
                            .XMLReplaceWithNo("前土地の所在B", n前Count + 1, 2, "以 下 余 白")
                            .XMLReplaceWithNo("前土地の所在", n前Count + 1, 2, "以 下 余 白")

                            .XMLReplaceWithNo("前土地の所在", 5, 2, "")
                            .XMLReplaceWithNo("前登記地目", 5, 2, "")
                            .XMLReplaceWithNo("前現況地目", 5, 2, "合 計")
                            .XMLReplaceWithNo("前面積", 5, 2, AreaConv(n前Area))
                        Else
                            .XMLReplaceWithNo("前土地の所在B", n前Count + 1, 2, "以 下 余 白")
                            .XMLReplaceWithNo("前土地の所在", n前Count + 1, 2, "以 下 余 白")

                            .XMLReplaceWithNo("前土地の所在B", n前Count + 2, 2, "")
                            .XMLReplaceWithNo("前土地の所在", n前Count + 2, 2, "")
                            .XMLReplaceWithNo("前登記地目", n前Count + 2, 2, "")
                            .XMLReplaceWithNo("前現況地目", n前Count + 2, 2, "合 計")
                            .XMLReplaceWithNo("前面積", n前Count + 2, 2, AreaConv(n前Area))
                        End If

                        For i As Integer = n前Count + 1 To 5
                            .XMLReplaceWithNo("前土地の所在B", i, 2, "")
                            .XMLReplaceWithNo("前土地の所在", i, 2, "")
                            .XMLReplaceWithNo("前登記地目", i, 2, "")
                            .XMLReplaceWithNo("前現況地目", i, 2, "")
                            .XMLReplaceWithNo("前面積", i, 2, "")
                        Next
                        .SaveAndOpen(sFolder, p申請.GetItem("許可番号", ""), p申請.GetItem("名称").ToString)
                    End With
                End If

                p申請.SetItem("状態", 2)
                p申請.SetItem("許可年月日", p異動日)
                p申請.SaveMyself()
            End If
        End If
    End Sub


    Public Function fnc設置利用権(ByVal sKey As String, ByVal n異動事由 As Long, ByVal sFolder As String, Optional ByVal b無異動 As Boolean = False, Optional ByVal p異動日 As Object = Nothing, Optional ByVal o通知発行日 As Object = Nothing) As Boolean
        'Try
        Dim p申請Row As DataRow = App農地基本台帳.TBL申請.FindRowByID(GetKeyCode(sKey))

        If p申請Row IsNot Nothing Then
            Dim p申請 As New CObj申請(p申請Row, False)

            Dim p許可通知書 As New 許可通知書
            p許可通知書.dt異動日 = HimTools2012.DateFunctions.NullCast日付(p異動日, Now.Date)

            Select Case p申請.法令
                Case enum法令.農地法3条耕作権 : p許可通知書.OpenFile("\農地法第3条許可書（使用貸借）.xml")
                Case enum法令.利用権設定, enum法令.利用権移転
                    p許可通知書.OpenFile("\利用権設定決定通知書.xml")

                    If Not {"宗像市"}.Contains(SysAD.市町村.市町村名) Then
                        If o通知発行日 Is Nothing OrElse IsDBNull(o通知発行日) OrElse Not IsDate(o通知発行日) Then
                            If p許可通知書.IsExists Then
                                o通知発行日 = InputBox("通知書の発行日を入力してください", "許可処理", Now.Date.ToString.Replace(" 0:00:00", ""))
                            End If
                        End If
                        If o通知発行日 IsNot Nothing AndAlso Not IsDBNull(o通知発行日) Then
                            p許可通知書.dt通知発行日 = HimTools2012.DateFunctions.NullCast日付(o通知発行日, Now.Date)
                        End If
                    End If
                Case Else
                    Stop
            End Select

            Dim St As String = p申請Row.Item("農地リスト").ToString
            Dim Cn() As String = Split(St, ";")

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

            For K As Integer = LBound(Cn) To UBound(Cn)

                If Len(Cn(K)) = 0 Then
                ElseIf GetKeyHead(Cn(K)) = "農地" Then
                    p許可通知書.nCount += 1
                    Dim p農地 As CObj農地 = ObjectMan.GetObject(Cn(K))
                    If p農地 IsNot Nothing Then
                        Dim p筆別Row As DataRow = Nothing
                        If pTBL IsNot Nothing AndAlso pTBL.PrimaryKey.Length > 0 Then
                            p筆別Row = pTBL.Rows.Find(Cn(K))

                            If Not p筆別Row Is Nothing AndAlso Not IsDBNull(p筆別Row.Item("申請部分面積")) AndAlso p筆別Row.Item("申請部分面積") > 0 Then
                                If p農地.GetIntegerValue("一部現況") = 0 Then
                                    If p農地.GetDecimalValue("登記簿面積") > p筆別Row.Item("申請部分面積") Then
                                        p農地.ValueChange("実面積", p農地.GetDecimalValue("登記簿面積") - p筆別Row.Item("申請部分面積"))
                                        If p農地.GetDecimalValue("田面積") > 0 Then
                                            p農地.ValueChange("田面積", p農地.GetDecimalValue("実面積"))
                                        ElseIf p農地.GetDecimalValue("畑面積") > 0 Then
                                            p農地.ValueChange("畑面積", p農地.GetDecimalValue("畑面積"))
                                        End If
                                        p農地.SetIntegerValue("一部現況", 1)

                                        p農地.SaveMyself()

                                        p農地 = p農地.CopyObject()
                                        p農地.SetIntegerValue("一部現況", 2)
                                        p農地.ValueChange("実面積", p筆別Row.Item("申請部分面積"))
                                        If p農地.GetDecimalValue("田面積") > 0 Then
                                            p農地.ValueChange("田面積", p農地.GetDecimalValue("実面積"))
                                        ElseIf p農地.GetDecimalValue("畑面積") > 0 Then
                                            p農地.ValueChange("畑面積", p農地.GetDecimalValue("畑面積"))
                                        End If



                                    Else
                                        MsgBox(String.Format("申請部分面積が登記簿面積を越えています。分割はできません。"), MsgBoxStyle.Critical)
                                    End If
                                Else
                                    MsgBox(String.Format("[{0}]は既に分割されています。再分割はできません。", p農地.所在), MsgBoxStyle.Critical)
                                End If

                            End If
                        End If

                        Select Case Val(p申請.GetItem("権利種類").ToString)
                            Case 1 : p許可通知書.s小作形態 = "賃貸借"
                            Case 2 : p許可通知書.s小作形態 = "使用貸借"
                            Case Else
                        End Select

                        Dim sNameA As String = p申請.GetItem("氏名A").ToString
                        Dim sNameB As String = p申請.GetItem("氏名B").ToString

                        If Not b無異動 AndAlso p農地 IsNot Nothing Then
                            p農地.ValueChange("自小作別", IIF(Val(p申請.GetItem("年金関連").ToString), 2, 1))
                            p農地.ValueChange("借受人ID", Val(p申請.GetItem("申請者B").ToString))
                            p農地.ValueChange("借受世帯ID", Val(p申請.GetItem("申請世帯B").ToString))

                            Select Case p申請.法令
                                Case 30, 31 : p農地.ValueChange("小作地適用法", 1)
                                Case Else : p農地.ValueChange("小作地適用法", 2)
                            End Select

                            p農地.ValueChange("小作形態", Val(p申請.GetItem("権利種類").ToString))
                            p農地.ValueChange("小作開始年月日", IIF(IsDBNull(p申請.GetItem("始期")), DBNull.Value, p申請.GetItem("始期")))
                            p農地.ValueChange("小作終了年月日", IIF(IsDBNull(p申請.GetItem("終期")), DBNull.Value, p申請.GetItem("終期")))

                            If p筆別Row IsNot Nothing AndAlso pTBL.Columns.Contains("賃借料1円10a当たり") AndAlso Not IsDBNull(p筆別Row.Item("賃借料1円10a当たり")) Then
                                '(21)1年間の賃金額
                                Select Case SysAD.市町村.市町村名
                                    Case "姶良市"
                                        p農地.ValueChange("小作料", IIF(IsDBNull(p申請.GetItem("小作料")), DBNull.Value, p申請.GetItem("小作料")))
                                        p農地.ValueChange("小作料単位", IIF(IsDBNull(p申請.GetItem("小作料単位")), DBNull.Value, p申請.GetItem("小作料単位")))
                                    Case Else
                                        p農地.ValueChange("小作料", p筆別Row.Item("賃借料1円10a当たり"))
                                        p農地.ValueChange("小作料単位", "円/10a")
                                End Select

                                p農地.ValueChange("10a賃借料", p筆別Row.Item("賃借料1円10a当たり"))
                            Else
                                p農地.ValueChange("小作料", IIF(IsDBNull(p申請.GetItem("小作料")), DBNull.Value, p申請.GetItem("小作料")))
                                p農地.ValueChange("小作料単位", IIF(IsDBNull(p申請.GetItem("小作料単位")), DBNull.Value, p申請.GetItem("小作料単位")))
                            End If

                            If Not IsDate(p申請.GetItem("公告年月日")) Then
                                p農地.ValueChange("公告年月日", p許可通知書.dt異動日)
                            Else
                                p農地.ValueChange("公告年月日", p申請.GetItem("公告年月日"))
                            End If

                            If Not IsDBNull(p申請.GetItem("経由法人ID")) AndAlso Val(p申請.GetItem("経由法人ID")) <> 0 Then
                                If Val(p申請.GetItem("経由法人ID").ToString) = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) Then
                                    p農地.ValueChange("農業生産法人経由貸借", True)
                                    p農地.ValueChange("経由農業生産法人ID", p申請.GetItem("経由法人ID"))

                                    p農地.ValueChange("中間管理権取得日", p申請.GetItem("機構配分計画中間管理権取得日"))
                                    p農地.ValueChange("意見回答日", p申請.GetItem("機構配分計画意見回答日"))
                                    p農地.ValueChange("知事公告日", p申請.GetItem("機構配分計画知事公告日"))
                                    p農地.ValueChange("認可通知日", p申請.GetItem("機構配分計画認可通知日"))
                                    p農地.ValueChange("権利設定内容", p申請.GetItem("機構配分計画権利設定内容"))

                                    p農地.ValueChange("利用配分計画始期日", p申請.GetItem("機構配分計画利用配分計画始期日"))
                                    p農地.ValueChange("利用配分計画終期日", p申請.GetItem("機構配分計画利用配分計画終期日"))
                                    'p農地.ValueChange("利用配分計画借賃額", p申請.GetItem("機構配分計画利用配分計画借賃額"))
                                    p農地.ValueChange("利用配分計画10a賃借料", p申請.GetItem("機構配分計画利用配分計画10a賃借料"))
                                    p農地.ValueChange("経由農業生産法人ID", p申請.GetItem("経由法人ID"))

                                    Select Case Val(p申請.GetItem("権利種類").ToString)
                                        Case 1 : p許可通知書.s小作形態 = "中間管理機構を介した賃貸借"
                                        Case 2 : p許可通知書.s小作形態 = "中間管理機構を介した使用貸借"
                                        Case Else
                                            p許可通知書.s小作形態 = "中間管理機構を介した" & p許可通知書.s小作形態
                                    End Select
                                End If

                            Else
                                p農地.ValueChange("農業生産法人経由貸借", False)
                                p農地.ValueChange("経由農業生産法人ID", 0)
                            End If

                            If Not IsDBNull(p申請.GetItem("代理人A")) AndAlso Val(p申請.GetItem("代理人A").ToString) <> 0 Then
                                p農地.ValueChange("管理者ID", p申請.GetItem("代理人A"))
                                p農地.ValueChange("農地所有内訳", 2)
                            End If

                            'p農地.SetBoolValue("解除条件付きの農地の貸借", p申請.getItem("解除条件付きの農地の貸借"))

                            Dim p始期 As String = ""
                            Dim p終期 As String = ""
                            If Not IsDBNull(p申請.GetItem("始期")) AndAlso Not IsDBNull(p申請.GetItem("終期")) Then
                                If IsDate(p申請.GetItem("始期")) AndAlso IsDate(p申請.GetItem("終期")) Then
                                    p始期 = 和暦Format(p申請.GetItem("始期"))
                                    p終期 = 和暦Format(p申請.GetItem("終期"))
                                Else
                                    MsgBox("年度が正しく設定されていません。")
                                    Return True
                                End If
                            Else
                                p始期 = "未設定"
                                p終期 = "未設定"
                            End If

                            If Not IsDBNull(p申請.GetItem("再設定")) AndAlso p申請.GetItem("再設定") = True Then
                                Dim s内容 As String = sNameA & "→" & sNameB & "へ" & p許可通知書.s小作形態 & "[再設定] " & p始期 & "-" & p終期 ' & IIf(p申請.getItem("解除条件付きの農地の貸借"), "[条件付き]", "") & ""
                                Make農地履歴(p農地.ID, Now.Date, p許可通知書.dt異動日, n異動事由, p申請.法令, s内容, p申請.申請者A, p申請.申請者B, Val(GetKeyCode(sKey)))
                            Else
                                Dim s内容 As String = sNameA & "→" & sNameB & "へ" & p許可通知書.s小作形態 & " " & p始期 & "-" & p終期 '& IIf(p申請.getItem("解除条件付きの農地の貸借"), "[条件付き]", "") & ""
                                Make農地履歴(p農地.ID, Now.Date, p許可通知書.dt異動日, n異動事由, p申請.法令, s内容, p申請.申請者A, p申請.申請者B, Val(GetKeyCode(sKey)))
                            End If
                            p農地.SaveMyself()
                        End If

                        If p許可通知書.IsExists Then
                            p許可通知書.Set各筆明細(p農地)
                        End If
                        p許可通知書.nArea += Val(p農地.GetItem("登記簿面積").ToString)
                        p許可通知書.cArea += Val(p農地.GetItem("実面積").ToString)
                    End If

                End If
            Next

            If p許可通知書.IsExists Then
                p許可通知書.Set書面設定(p申請)
                Dim s内容 As String = IIF(s代理人 = "", "(" & p申請.GetItem("許可番号", "") & "_" & p申請.GetItem("名称").ToString & ").xml", "(" & p申請.GetItem("許可番号", "") & "_" & p申請.GetItem("名称").ToString & "【" & s代理人 & "宛て】).xml")
                HimTools2012.TextAdapter.SaveTextFile(sFolder & Replace(p許可通知書.sOutput.ToLower, ".xml", s内容), p許可通知書.sXML)
                SysAD.ShowFolder(sFolder)
            End If

            p申請.SetItem("状態", 2)
            p申請.SetItem("許可年月日", p許可通知書.dt異動日)
            If Not IsDate(p申請.GetItem("公告年月日")) Then
                p申請.SetItem("公告年月日", p許可通知書.dt異動日)
            End If

            p申請.SaveMyself()
        End If
        'Catch ex As Exception

        '    Stop
        'End Try
        Return True
    End Function

    Private s代理人 As String
    Public Function fnc通知書発行(ByVal sKey As String, ByVal n異動事由 As Long, ByVal sFolder As String, Optional ByVal b無異動 As Boolean = False, Optional ByVal p異動日 As Object = Nothing, Optional ByVal o通知発行日 As Object = Nothing) As Boolean
        'Try
        Dim p申請Row As DataRow = App農地基本台帳.TBL申請.FindRowByID(GetKeyCode(sKey))

        If p申請Row IsNot Nothing Then
            Dim p申請 As New CObj申請(p申請Row, False)
            Dim bResult As Boolean = True
            Dim p許可通知書 As New 許可通知書
            p許可通知書.dt異動日 = HimTools2012.DateFunctions.NullCast日付(p異動日, Now.Date)

            Select Case p申請.権利種類
                Case 21 : p許可通知書.OpenFile("\農地法第3条解約通知書(賃貸借).xml")
                Case 22 : p許可通知書.OpenFile("\農地法第3条解約通知書(使用貸借).xml")
                Case 23 : p許可通知書.OpenFile("\基盤強化法解約通知書(賃貸借).xml")
                Case 24 : p許可通知書.OpenFile("\基盤強化法解約通知書(使用貸借).xml")
                Case Else
                    MsgBox("申請「" & p申請.名称 & "」において権利種類が設定されていないため通知書の発行ができませんでした。")
                    bResult = False
            End Select

            If bResult = True Then
                Dim St As String = p申請Row.Item("農地リスト").ToString
                Dim Cn() As String = Split(St, ";")

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

                For K As Integer = LBound(Cn) To UBound(Cn)
                    If Len(Cn(K)) = 0 Then
                    ElseIf GetKeyHead(Cn(K)) = "農地" Then
                        p許可通知書.nCount += 1
                        Dim p農地 As CObj農地 = ObjectMan.GetObject(Cn(K))
                        If p農地 IsNot Nothing Then
                            If p許可通知書.IsExists Then
                                p許可通知書.Set各筆明細(p農地)
                            End If
                        End If
                    End If
                Next

                If p許可通知書.IsExists Then
                    p許可通知書.Set書面設定(p申請, False)
                    Dim s内容 As String = "(" & p申請.GetItem("名称").ToString & ").xml"
                    HimTools2012.TextAdapter.SaveTextFile(sFolder & Replace(p許可通知書.sOutput.ToLower, ".xml", s内容), p許可通知書.sXML)

                    s代理人 = ""
                End If
            End If
        End If
        Return True
    End Function


    Public Class 許可通知書
        Public sXML As String = ""
        Public sOutput As String = ""
        Public nCount As Integer = 0
        Public s小作形態 As String = "使用貸借"
        Public nArea As Decimal = 0
        Public cArea As Decimal = 0
        Public dt異動日 As DateTime
        Public dt通知発行日 As DateTime

        Public Sub New()

        End Sub

        Public Sub OpenFile(Optional ByVal pOutFile As String = "")
            If pOutFile.Length > 0 Then
                sOutput = pOutFile
            End If

            Dim sPath As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & sOutput
            If IO.File.Exists(sPath) Then
                sXML = HimTools2012.TextAdapter.LoadTextFile(sPath)
            End If
        End Sub

        Public Function IsExists() As Boolean
            Return sXML.Length > 0
        End Function

        Public Sub Set書面設定(ByVal p申請 As CObj申請, Optional ByVal VisTotalArea As Boolean = True)
            If nCount < 6 And SysAD.DB(sLRDB).DBProperty("市町村名").ToString = "姶良市" Then
                Try
                    Dim sXML2 As String = sXML.Substring(sXML.IndexOf(" <Worksheet ss:Name=""印刷様式A02"">"))
                    Dim sXML3 As String = sXML2.Substring(0, sXML2.LastIndexOf("</Worksheet>") + Len("</Worksheet>") + 2)

                    sXML = Replace(sXML, sXML3, "")
                Catch ex As Exception

                End Try

            End If
            sXML = Replace(sXML, "{発行年月日}", 和暦Format(dt通知発行日))

            p申請.Replace申請者A(sXML)
            p申請.Replace申請者B(sXML)
            Set代理人情報(p申請)
            sXML = Replace(sXML, "{会長名}", SysAD.DB(sLRDB).DBProperty("会長名").ToString)
            sXML = Replace(sXML, "{市町村名}", SysAD.DB(sLRDB).DBProperty("市町村名").ToString)
            sXML = Replace(sXML, "{許可年月日}", 和暦Format(dt異動日))
            sXML = Replace(sXML, "{年度}", Strings.Mid(和暦Format(dt異動日), 3, 2))
            sXML = Replace(sXML, "{許可月}", Strings.Mid(和暦Format(dt異動日), 3, 2))

            sXML = Replace(sXML, "{総会番号}", Val(p申請.Row.Body.Item("総会番号").ToString))
            sXML = Replace(sXML, "{受付番号}", Val(p申請.Row.Body.Item("受付番号").ToString))

            sXML = Replace(sXML, "{許可番号}", p申請.GetItem("許可番号").ToString)

            If IsDate(p申請.GetItem("受付年月日")) Then
                sXML = Replace(sXML, "{申請年月日}", 和暦Format(p申請.GetItem("受付年月日")))
            Else
                If Year(Now) >= 2019 AndAlso Month(Now) >= 5 Then
                    sXML = Replace(sXML, "{申請年月日}", "令和　　年　　月　　日")
                Else
                    sXML = Replace(sXML, "{申請年月日}", "平成　　年　　月　　日")
                End If
            End If

            If Val(p申請.Row.Body.Item("区分地上権").ToString) > 0 Then
                sXML = Replace(sXML, "{権利種類}", "区分地上")
            Else
                Select Case Val(p申請.GetItem("権利種類").ToString)
                    Case 1 : sXML = Replace(sXML, "{権利種類}", "賃借")
                    Case 2 : sXML = Replace(sXML, "{権利種類}", "使用貸借")
                    Case Else : sXML = Replace(sXML, "{権利種類}", "未設定")
                End Select
            End If

            Dim p始期 As String = ""
            Dim p終期 As String = ""
            If Not IsDBNull(p申請.GetItem("始期")) AndAlso Not IsDBNull(p申請.GetItem("終期")) Then
                If IsDate(p申請.GetItem("始期")) AndAlso IsDate(p申請.GetItem("終期")) Then
                    p始期 = 和暦Format(p申請.Row.Body.Item("始期"))
                    p終期 = 和暦Format(p申請.Row.Body.Item("終期"))
                End If
            End If
            Dim sXX As String = p始期

            If p申請.GetBoolValue("永久") Then
                sXX = sXX & "  ～  " & "(永年)"
            Else
                sXX = sXX & "  ～  " & p終期
            End If

            If Val(p申請.GetItem("期間").ToString) > 0 Then
                sXML = Replace(sXML, "{期間}", sXX & " ( " & Val(p申請.GetItem("期間").ToString) & "年 )")
            Else
                If IsDate(p申請.Row.Body.Item("始期")) AndAlso IsDate(p申請.Row.Body.Item("終期")) Then
                    sXML = Replace(sXML, "{期間}", sXX & " ( " & DateDiff(DateInterval.Year, p申請.Row.Body.Item("始期"), p申請.Row.Body.Item("終期")) & "年 )")
                Else
                    sXML = Replace(sXML, "{期間}", "")
                End If
            End If

            If s小作形態 = "賃貸借" Then
                sXML = Replace(sXML, "{小作料表記}", p申請.GetItem("小作料").ToString & p申請.GetItem("小作料単位").ToString)
            Else
                sXML = Replace(sXML, "{小作料表記}", "")
            End If

            If VisTotalArea = True Then
                If nCount = 5 Then
                    sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & 6, 2) & "}", "")
                    sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & 6, 2) & "}", "")
                    sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & 6, 2) & "}", "")
                    sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & 6, 2) & "}", "合 計")
                    sXML = Replace(sXML, "{面積" & Strings.Right("000" & 6, 2) & "}", AreaConv(nArea))
                    sXML = Replace(sXML, "{現況面積" & Strings.Right("000" & 6, 2) & "}", AreaConv(cArea))
                    sXML = Replace(sXML, "{備考" & Strings.Right("000" & 6, 2) & "}", "")
                ElseIf nCount < 5 Then
                    sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount + 1, 2) & "}", "以 下 余 白")
                    sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount + 1, 2) & "}", "以 下 余 白")
                    sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & 6, 2) & "}", "")
                    sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & 6, 2) & "}", "")
                    sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & 6, 2) & "}", "")
                    sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & 6, 2) & "}", "合 計")
                    sXML = Replace(sXML, "{面積" & Strings.Right("000" & 6, 2) & "}", AreaConv(nArea))
                    sXML = Replace(sXML, "{現況面積" & Strings.Right("000" & 6, 2) & "}", AreaConv(cArea))
                    sXML = Replace(sXML, "{備考" & Strings.Right("000" & 6, 2) & "}", "")
                Else
                    sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount + 1, 2) & "}", "以 下 余 白")
                    sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount + 1, 2) & "}", "以 下 余 白")

                    sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount + 2, 2) & "}", "")
                    sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount + 2, 2) & "}", "")
                    sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & nCount + 2, 2) & "}", "")
                    sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & nCount + 2, 2) & "}", "合 計")
                    sXML = Replace(sXML, "{面積" & Strings.Right("000" & nCount + 2, 2) & "}", AreaConv(nArea))
                    sXML = Replace(sXML, "{現況面積" & Strings.Right("000" & nCount + 2, 2) & "}", AreaConv(cArea))
                    sXML = Replace(sXML, "{備考" & Strings.Right("000" & nCount + 2, 2) & "}", "")
                End If
            End If

            For i As Integer = nCount + 1 To 60
                sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{面積" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{現況面積" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{備考" & Strings.Right("000" & i, 2) & "}", "")
            Next
        End Sub

        Public Sub Set各筆明細(ByRef p農地 As CObj農地)
            If Val(p農地.Row.Body.Item("管理者ID").ToString) <> 0 Then
                Dim pRow = App農地基本台帳.TBL個人.FindRowByID(Val(p農地.Row.Body.Item("管理者ID").ToString))
                sXML = Replace(sXML, "{郵便番号A}", "〒" & pRow("郵便番号").ToString)
                sXML = Replace(sXML, "{住所A}", pRow("住所").ToString)
                sXML = Replace(sXML, "{氏名A}", pRow("氏名").ToString)
                s代理人 = pRow("氏名").ToString
            End If

            sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount, 2) & "}", Replace地番(p農地.土地所在))
            sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount, 2) & "}", p農地.土地所在)
            sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & nCount, 2) & "}", p農地.Row.Body.Item("登記簿地目名").ToString)
            sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & nCount, 2) & "}", p農地.Row.Body.Item("現況地目名").ToString)
            sXML = Replace(sXML, "{面積" & Strings.Right("000" & nCount, 2) & "}", AreaConv(Val(p農地.GetItem("登記簿面積").ToString)))
            sXML = Replace(sXML, "{現況面積" & Strings.Right("000" & nCount, 2) & "}", AreaConv(Val(p農地.GetItem("実面積").ToString)))
            sXML = Replace(sXML, "{備考" & Strings.Right("000" & nCount, 2) & "}", "")
        End Sub

        Public Sub Set代理人情報(ByRef p申請 As CObj申請)
            If Val(p申請.GetItem("代理人A").ToString) > 0 Then
                Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(p申請.GetItem("代理人A"))
                If Not pRow Is Nothing Then
                Else
                    pRow = App農地基本台帳.TBL個人.FindRowByID(p申請.GetItem("申請者A"))
                End If
                sXML = Replace(sXML, "{送付先氏名}", pRow.Item("氏名"))
                sXML = Replace(sXML, "{送付先住所}", pRow.Item("住所"))
                sXML = Replace(sXML, "{送付先郵便番号}", IIF(InStr(pRow.Item("郵便番号").ToString, "〒") > 0, pRow.Item("郵便番号").ToString, "〒" & pRow.Item("郵便番号").ToString))
            Else
                Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(p申請.GetItem("申請者A"))
                If Not pRow Is Nothing Then
                    sXML = Replace(sXML, "{送付先氏名}", pRow.Item("氏名"))
                    sXML = Replace(sXML, "{送付先住所}", pRow.Item("住所"))
                    sXML = Replace(sXML, "{送付先郵便番号}", IIF(InStr(pRow.Item("郵便番号").ToString, "〒") > 0, pRow.Item("郵便番号").ToString, "〒" & pRow.Item("郵便番号").ToString))
                Else
                    sXML = Replace(sXML, "{送付先氏名}", "")
                    sXML = Replace(sXML, "{送付先住所}", "")
                    sXML = Replace(sXML, "{送付先郵便番号}", "")
                End If
            End If
        End Sub

    End Class

    Public Sub sub農地転用(ByVal p申請Row As HimTools2012.Data.DataRowPlus, ByVal n異動事由 As Long, ByVal sFolder As String, Optional ByVal b無異動 As Boolean = False, Optional ByVal p異動日 As Object = Nothing)
        Dim dt異動日 As DateTime
        Dim sError As String = ""

        'Try
        If p異動日 Is Nothing OrElse Not IsDate(p異動日) Then
            dt異動日 = Now.Date
        Else
            If Year(p異動日) > 1900 Then
                dt異動日 = p異動日
            Else
                dt異動日 = Now.Date
            End If
        End If

        Dim s法令番号 As String = ""
        Dim p申請 As New CObj申請(p申請Row.Body, False)
        Dim sOutput As String = ""

        Select Case p申請.法令
            Case enum法令.農地法4条, enum法令.農地法4条一時転用 : s法令番号 = "4" : sOutput = "\農地法第4条第１項許可書.xml"
            Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用 : s法令番号 = "5" : sOutput = "\農地法第5条許可書.xml"
            Case Else
                Stop
        End Select

        If {"三股町"}.Contains(SysAD.市町村.市町村名) Then
            sOutput = ""
        End If

        Dim nNo As Integer = Val(p申請.許可番号.ToString)
        If p申請Row IsNot Nothing AndAlso nNo > 0 Then
            Dim St As String = p申請Row.Item("農地リスト").ToString
            Dim Cn() As String = Split(St, ";")
            Dim sXML As String = ""
            If sOutput = "" Then
            Else
                sXML = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & sOutput)
            End If

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

            Dim nCount As Integer = 0
            For K As Integer = LBound(Cn) To UBound(Cn)
                If Len(Cn(K)) = 0 Then
                Else
                    nCount += 1

                    Dim p農地 As Object = Nothing
                    Select Case Split(Cn(K), ".")(0)
                        Case "農地"
                            p農地 = ObjectMan.GetObject(Cn(K))
                        Case "転用農地"
                            p農地 = ObjectMan.GetObject(Cn(K))
                    End Select

                    If p農地 Is Nothing OrElse p農地.Row Is Nothing Then
                    Else
                        Dim p筆別Row As DataRow = Nothing
                        If pTBL IsNot Nothing AndAlso pTBL.PrimaryKey.Length > 0 Then
                            p筆別Row = pTBL.Rows.Find("農地." & p農地.ID)

                            If Not p筆別Row Is Nothing AndAlso Not IsDBNull(p筆別Row.Item("申請部分面積")) AndAlso p筆別Row.Item("申請部分面積") > 0 Then
                                If p農地.GetIntegerValue("一部現況") = 0 Then
                                    If p農地.GetDecimalValue("登記簿面積") > p筆別Row.Item("申請部分面積") Then
                                        p農地.ValueChange("実面積", p農地.GetDecimalValue("登記簿面積") - p筆別Row.Item("申請部分面積"))
                                        If p農地.GetDecimalValue("田面積") > 0 Then
                                            p農地.ValueChange("田面積", p農地.GetDecimalValue("実面積"))
                                        ElseIf p農地.GetDecimalValue("畑面積") > 0 Then
                                            p農地.ValueChange("畑面積", p農地.GetDecimalValue("畑面積"))
                                        End If
                                        p農地.SetIntegerValue("一部現況", 1)

                                        p農地.SaveMyself()

                                        p農地 = p農地.CopyObject()
                                        p農地.SetIntegerValue("一部現況", 2)
                                        p農地.ValueChange("実面積", p筆別Row.Item("申請部分面積"))
                                        If p農地.GetDecimalValue("田面積") > 0 Then
                                            p農地.ValueChange("田面積", p農地.GetDecimalValue("実面積"))
                                        ElseIf p農地.GetDecimalValue("畑面積") > 0 Then
                                            p農地.ValueChange("畑面積", p農地.GetDecimalValue("畑面積"))
                                        End If
                                    Else
                                        MsgBox(String.Format("申請部分面積が登記簿面積を越えています。分割はできません。"), MsgBoxStyle.Critical)
                                    End If
                                Else
                                    MsgBox(String.Format("[{0}]は既に分割されています。再分割はできません。", p農地.土地所在), MsgBoxStyle.Critical)
                                End If

                            End If
                        End If

                        If sXML.Length > 0 Then
                            sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount, 2) & "}", p農地.GetProperty("土地所在"))
                            sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount, 2) & "}", Replace地番(p農地.GetProperty("土地所在")))
                            sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & nCount, 2) & "}", p農地.GetItem("登記簿地目名", ""))
                            sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & nCount, 2) & "}", p農地.GetItem("現況地目名", ""))
                            sXML = Replace(sXML, "{面積" & Strings.Right("000" & nCount, 2) & "}", HimTools2012.NumericFunctions.NumToString(Val(p農地.GetItem("登記簿面積").ToString)))
                            sXML = Replace(sXML, "{備考" & Strings.Right("000" & nCount, 2) & "}", "")
                        End If

                        If Not b無異動 AndAlso p農地 IsNot Nothing Then
                            Select Case Val(p申請Row.Item("権利種類").ToString)
                                Case 7
                                    p農地.ValueChange("小作地適用法", 1)
                                    p農地.ValueChange("小作形態", Val(p申請Row.Item("権利種類").ToString))
                                    p農地.SaveMyself()
                            End Select
                            Select Case p申請.法令
                                Case enum法令.農地法5条所有権
                                    p農地.ValueChange("所有者ID", p申請Row.Item("申請者B"))
                                    p農地.ValueChange("所有世帯ID", p申請Row.Item("申請世帯B"))
                                Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                                    p農地.ValueChange("借受人ID", p申請Row.Item("申請者B"))
                                    p農地.ValueChange("借受世帯ID", p申請Row.Item("申請世帯B"))

                                    p農地.ValueChange("小作開始年月日", p申請Row.Item("始期"))
                                    p農地.ValueChange("小作終了年月日", p申請Row.Item("終期"))
                            End Select

                            p農地.ValueChange("農地状況", IIF(p申請.法令 = enum法令.農地法5条一時転用, 農地状況.一時転用中, 農地状況.転用許可済み))
                            p農地.SaveMyself()

                            If p農地.Key.DataClass = "農地" Then
                                Sub農地転用データ転送(p農地, p農地.ID)
                            End If

                            Make農地履歴(GetKeyCode(Cn(K)), Now.Date, dt異動日, n異動事由, p申請.法令, "転用", , , Val(p申請Row.Item("ID")))
                        End If

                    End If
                End If
            Next
            Dim s許可年月日 = 和暦Format(dt異動日)

            If sXML.Length > 0 Then

                sXML = Replace(sXML, "{土地の所在01}", "")
                sXML = Replace(sXML, "{土地の所在B01}", "")
                sXML = Replace(sXML, "{登記地目01}", "")
                sXML = Replace(sXML, "{現況地目01}", "")
                sXML = Replace(sXML, "{面積01}", "")
                sXML = Replace(sXML, "{備考01}", "")
                p申請.Replace申請者A(sXML)

                sXML = Replace(sXML, "{申請者Ｂ住所}", p申請.GetItem("住所B").ToString)
                sXML = Replace(sXML, "{申請者Ｂ氏名}", p申請.GetItem("氏名B").ToString)

                sXML = Replace(sXML, "{農業委員会会長氏名}", SysAD.DB(sLRDB).DBProperty("会長名").ToString)
                sXML = Replace(sXML, "{市町村}", SysAD.DB(sLRDB).DBProperty("市町村").ToString)
                sXML = Replace(sXML, "{市町村名}", SysAD.DB(sLRDB).DBProperty("市町村名").ToString)
                sXML = Replace(sXML, "{許可申請年月日}", 和暦Format(p申請.GetItem("受付年月日")))
                sXML = Replace(sXML, "{申請年月日}", 和暦Format(p申請.GetItem("受付年月日")))
                sXML = Replace(sXML, "{年度}", Strings.Mid(和暦Format(dt異動日), 3, 2))
                If IsDate(dt異動日) Then
                    sXML = Replace(sXML, "{許可月}", Month(dt異動日))
                Else
                    sXML = Replace(sXML, "{許可月}", "　")
                End If
                sXML = Replace(sXML, "{総会番号}", Val(p申請.GetItem("総会番号").ToString))
                sXML = Replace(sXML, "{受付番号}", Val(p申請.GetItem("受付番号").ToString))
                sXML = Replace(sXML, "{通知年月日}", 和暦Format(Now))
                sXML = Replace(sXML, "{許可年月日}", s許可年月日)
                sXML = Replace(sXML, "{用途}", p申請.GetItem("用途").ToString)
                sXML = Replace(sXML, "{転用目的}", p申請.GetItem("申請理由A").ToString)

                Select Case nCount
                    Case 1 To 10 : sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount + 1, 2) & "}", IIF(nCount > 3, "", "以下余白"))
                    Case Else : sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount + 1, 2) & "}", IIF(nCount > 25, "", "以下余白"))
                End Select

                Select Case nCount
                    Case 1 To 5 : sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount + 1, 2) & "}", IIF(nCount > 3, "", "以下余白"))
                    Case Else : sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount + 1, 2) & "}", IIF(nCount > 25, "", "以下余白"))
                End Select

                For i As Integer = nCount + 1 To 50
                    sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & i, 2) & "}", "")
                    sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & i, 2) & "}", "")
                    sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & i, 2) & "}", "")
                    sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & i, 2) & "}", "")
                    sXML = Replace(sXML, "{面積" & Strings.Right("000" & i, 2) & "}", "")
                    sXML = Replace(sXML, "{備考" & Strings.Right("000" & i, 2) & "}", "")
                Next
                Select Case SysAD.市町村.市町村名
                    Case "日置市"
                        sXML = Replace(sXML, "{農委番号}", String.Format("日農委第{0}号 {1}", s法令番号, nNo))
                    Case "長島町"
                        sXML = Replace(sXML, "{農委番号}", String.Format("農委第{0}-{1}", s法令番号, nNo))
                    Case Else
                        sXML = Replace(sXML, "{農委番号}", String.Format("農委第{0}号 {1}", s法令番号, nNo))
                End Select

                sXML = Replace(sXML, "{許可和年}", Mid(s許可年月日, 3, 2))
                sXML = Replace(sXML, "{許可番号}", nNo.ToString)

                Select Case p申請.法令
                    Case enum法令.農地法4条, enum法令.農地法4条一時転用
                        Dim sFile As String = sFolder & Replace(sOutput.ToLower, ".xml", "(" & nNo & "_" & p申請.GetItem("氏名A").ToString.Trim & ").xml")
                        HimTools2012.TextAdapter.SaveTextFile(sFile, sXML)
                        SysAD.ShowFolder(sFile)
                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                        Select Case p申請.法令
                            Case enum法令.農地法5条所有権
                                Select Case Val(p申請.Row.Body("所有権移転の種類").ToString)
                                    Case 0 : sXML = Replace(sXML, "{権利内容}", "")
                                    Case 1 : sXML = Replace(sXML, "{権利内容}", "所有権移転の売買を許可する土地")
                                    Case 2 : sXML = Replace(sXML, "{権利内容}", "所有権移転の贈与を許可する土地")
                                    Case 3 : sXML = Replace(sXML, "{権利内容}", "所有権移転の交換を許可する土地")
                                    Case 4 : sXML = Replace(sXML, "{権利内容}", "賃貸借を許可する土地")
                                    Case 5 : sXML = Replace(sXML, "{権利内容}", "使用貸借を許可する土地")
                                End Select
                            Case Else
                                Select Case Val(p申請.GetItem("権利種類").ToString)
                                    Case 2 : sXML = Replace(sXML, "{権利内容}", "使用貸借を許可する土地")
                                    Case Else : sXML = Replace(sXML, "{権利内容}", "賃貸借を許可する土地")
                                End Select
                        End Select

                        Select Case p申請.法令
                            Case enum法令.農地法5条所有権
                                sXML = Replace(sXML, "{出申請人}", "譲渡人")
                                sXML = Replace(sXML, "{受申請人}", "譲受人")
                                sXML = Replace(sXML, "{法令}", "所有権の移転")
                            Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                                sXML = Replace(sXML, "{出申請人}", "貸人")
                                sXML = Replace(sXML, "{受申請人}", "借人")
                                Select Case Val(p申請.GetItem("権利種類").ToString)
                                    Case 2 : sXML = Replace(sXML, "{法令}", "使用貸借権の設定")
                                    Case Else : sXML = Replace(sXML, "{法令}", "賃貸借権の設定")
                                End Select
                        End Select

                        Dim sFile As String = sFolder & Replace(sOutput.ToLower, ".xml", "(" & nNo & "_" & p申請.GetItem("氏名B").ToString.Trim & "←" & p申請.GetItem("氏名A").ToString.Trim & ").xml")
                        HimTools2012.TextAdapter.SaveTextFile(sFile, sXML)
                        SysAD.ShowFolder(sFile)
                End Select
            End If

            St = ";" & St
            St = Replace(St, ";農地.", ";転用農地.")
            St = Mid$(St, 2)
            p申請.SetItem("状態", 2)
            p申請.SetItem("許可年月日", dt異動日)

            p申請.SetItem("農地リスト", St)
            p申請.SaveMyself()
        End If
        'Catch ex As Exception
        '    Stop
        'End Try

    End Sub

    Public Sub sub非農地(ByVal p申請Row As DataRow, ByVal n異動事由 As Long, ByVal sFolder As String, Optional ByVal b無異動 As Boolean = False, Optional ByVal p異動日 As Object = Nothing)
        Dim dt異動日 As DateTime
        Dim dt総会日 As DateTime
        Dim sError As String = ""
        Dim Int証明番号 As Integer = 0

        Dim p申請 As New CObj申請(p申請Row, False)
        Dim sOutput As String = ""

        Select Case SysAD.市町村.市町村名
            Case "長島町"
                sOutput = "\非農地証明書.xml"
                Int証明番号 = InputBox("証明番号", "証明番号を入力してください", 1)
            Case Else
                sOutput = "\非農地通知済証明書.xml"
        End Select

        Dim nNo As Integer = Val(p申請.許可番号.ToString)

        If p申請Row IsNot Nothing AndAlso nNo > 0 Then
            Dim St As String = p申請Row.Item("農地リスト").ToString
            Dim Cn() As String = Split(St, ";")
            Dim Int総会番号 As Integer = 1

            Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & sOutput)

            sXML = Replace(sXML, "{市町村名}", SysAD.市町村.市町村名)

            If nNo = 0 Then : nNo = InputBox("許可番号", "許可番号を入力してください", Get許可番号MAX(enum法令.非農地証明願))
            Else : nNo = Val(p申請.許可番号.ToString)
            End If

            If p異動日 Is Nothing OrElse Not IsDate(p異動日) Then
                If IsDBNull(p申請Row.Item("許可年月日")) Then
                    dt異動日 = InputBox("許可日を入力してください", "許可処理", Now.Date.ToString.Replace(" 0:00:00", ""))
                Else
                    If Year(p申請Row.Item("許可年月日")) < 1990 Then : dt異動日 = InputBox("許可日を入力してください", "許可処理", Now.Date.ToString.Replace(" 0:00:00", ""))
                    Else : dt異動日 = p申請Row.Item("許可年月日")
                    End If
                End If
            Else
                dt異動日 = p異動日
            End If

            If IsDBNull(p申請Row.Item("総会番号")) = True Then
                Int総会番号 = InputBox("総会番号", "総会番号を入力してください", 1)
            Else
                Int総会番号 = p申請Row.Item("総会番号")
            End If

            If IsDBNull(p申請Row.Item("総会日")) Then
                dt総会日 = InputBox("総会日を入力してください", "許可処理", Now.Date.ToString.Replace(" 0:00:00", ""))
            Else
                If Year(p申請Row.Item("総会日")) < 1990 Then : dt総会日 = InputBox("総会日を入力してください", "許可処理", Now.Date.ToString.Replace(" 0:00:00", ""))
                Else : dt総会日 = p申請Row.Item("総会日")
                End If
            End If

            Dim nCount As Integer = 0
            For K As Integer = LBound(Cn) To UBound(Cn)
                If Len(Cn(K)) = 0 Then
                Else
                    nCount += 1
                    Dim p農地 As CTargetObjWithView農地台帳
                    p農地 = ObjectMan.GetObject(Cn(K))

                    If p農地 IsNot Nothing Then
                        If p農地 IsNot Nothing Then
                            sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & nCount, 2) & "}", p農地.GetProperty("土地所在").ToString)
                            sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & nCount, 2) & "}", Replace地番(p農地.GetProperty("土地所在").ToString))
                            sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & nCount, 2) & "}", p農地.Row.Body.Item("登記簿地目名").ToString)
                            sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & nCount, 2) & "}", p農地.Row.Body.Item("現況地目名").ToString)
                            sXML = Replace(sXML, "{面積" & Strings.Right("000" & nCount, 2) & "}", AreaConv(p農地.GetItem("登記簿面積").ToString))
                            sXML = Replace(sXML, "{名義人氏名" & Strings.Right("000" & nCount, 2) & "}", IIF(p農地.GetItem("名義人氏名").ToString = "", IIF(p農地.GetItem("名義人氏名").ToString = "", p農地.GetItem("所有者氏名").ToString, p農地.GetItem("管理者氏名").ToString), p農地.GetItem("名義人氏名").ToString))
                            sXML = Replace(sXML, "{申請者氏名" & Strings.Right("000" & nCount, 2) & "}", p申請.GetItem("氏名A").ToString)
                            sXML = Replace(sXML, "{備考" & Strings.Right("000" & nCount, 2) & "}", "")

                            If p農地.Key.DataClass = "農地" Then
                                p農地.ValueChange("農地状況", 農地状況.非農地)
                                p農地.SaveMyself()

                                Sub農地転用データ転送(p農地, p農地.ID)
                                Make農地履歴(GetKeyCode(Cn(K)), Now.Date, dt異動日, n異動事由, p申請.法令, "転用", , , Val(p申請Row.Item("ID")))
                            End If
                        End If
                    End If
                End If
            Next

            sXML = Replace(sXML, "{市町村名}", SysAD.市町村.市町村名)
            sXML = Replace(sXML, "{総会番号}", Int総会番号)
            sXML = Replace(sXML, "{証明番号}", Int証明番号)
            sXML = Replace(sXML, "{発行年月日}", 和暦Format(Now))

            sXML = Replace(sXML, "{農業委員会会長氏名}", SysAD.DB(sLRDB).DBProperty("会長名").ToString)
            sXML = Replace(sXML, "{総会日}", 和暦Format(dt総会日))
            sXML = Replace(sXML, "{許可申請年月日}", 和暦Format(p申請.GetItem("受付年月日")))
            sXML = Replace(sXML, "{申請年月日}", 和暦Format(p申請.GetItem("受付年月日")))
            sXML = Replace(sXML, "{許可年月日}", 和暦Format(dt異動日))

            sXML = Replace(sXML, "{申請者Ａ住所}", p申請.GetItem("住所A").ToString)
            sXML = Replace(sXML, "{申請者Ａ氏名}", p申請.GetItem("氏名A").ToString)
            sXML = Replace(sXML, "{申請者Ｂ住所}", p申請.GetItem("住所B").ToString)
            sXML = Replace(sXML, "{申請者Ｂ氏名}", p申請.GetItem("氏名B").ToString)

            For i As Integer = nCount + 1 To 50
                sXML = Replace(sXML, "{土地の所在" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{土地の所在B" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{登記地目" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{現況地目" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{面積" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{名義人氏名" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{申請者氏名" & Strings.Right("000" & i, 2) & "}", "")
                sXML = Replace(sXML, "{備考" & Strings.Right("000" & i, 2) & "}", "")
            Next
            Select Case SysAD.市町村.市町村名
                Case "日置市"
                    sXML = Replace(sXML, "{農委番号}", String.Format("日農委第{0}号 {1}", 4, nNo))
                Case Else
                    sXML = Replace(sXML, "{農委番号}", String.Format("農委第{0}号 {1}", 4, nNo))
            End Select

            Select Case SysAD.市町村.市町村名
                Case "伊佐市"

                Case Else
                    Dim sFile As String = sFolder & Replace(sOutput.ToLower, ".xml", "(" & nNo & "_" & p申請.GetItem("氏名A").ToString.Trim & ").xml")
                    HimTools2012.TextAdapter.SaveTextFile(sFile, sXML)
                    SysAD.ShowFolder(sFile)
            End Select

            St = ";" & St
            St = Replace(St, ";農地.", ";転用農地.")
            St = Mid$(St, 2)
            p申請.SetItem("状態", 2)
            p申請.SetItem("許可年月日", dt異動日)
            p申請.SetItem("許可番号", nNo)
            p申請.SetItem("農地リスト", St)
            p申請.SetItem("総会番号", Int総会番号)
            p申請.SetItem("総会日", dt総会日)
            p申請.SaveMyself()
        End If
        'Catch ex As Exception
        '    Stop
        'End Try

    End Sub

    Public Function Get許可番号MAX(ByVal n法令 As Integer) As Integer
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Max([D_申請].許可番号) AS 最大番号 FROM [D_申請] WHERE ([D_申請].法令 In ( " & n法令 & " )) AND (DatePart('yyyy',[許可年月日])=DatePart('yyyy',Date()))")
        If pTBL.Rows.Count = 1 Then
            Return Val(pTBL.Rows(0).Item("最大番号").ToString) + 1
        Else
            Return 1
        End If
    End Function

    Public Sub Sub農地転用データ転送(ByRef p農地 As CObj農地, ByVal TID As Long)
        If p農地 IsNot Nothing Then
            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM D_転用農地 WHERE [ID]=" & TID)
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_転用農地([ID]) VALUES(" & TID & ")")

            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [ID]=" & TID)
            Dim p転用 As New CObj転用農地(App農地基本台帳.TBL転用農地.FindRowByID(TID), False)

            For Each pCol As DataColumn In pTBL.Columns
                If pCol.ColumnName <> "ID" AndAlso App農地基本台帳.TBL農地.Columns.Contains(pCol.ColumnName) Then
                    p転用.ValueChange(pCol.ColumnName, p農地.GetItem(pCol.ColumnName))
                End If
            Next
            p転用.SaveMyself()
            p農地.DoCommand("閉じる")
            SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D:農地Info] WHERE [D:農地Info].ID=" & TID)
            p農地.Row.RemoveFromTable()
        End If
    End Sub

    Public Function 農地削除(ByVal bCancelable As Boolean, ByRef pTBL As DataTable, Optional ByVal n履歴 As Integer = 0, Optional ByVal p転送先 As C農地削除.enum転送先 = C農地削除.enum転送先.削除農地, Optional ByVal sMess As String = "", Optional ByVal dt異動日 As DateTime = Nothing) As Integer
        Dim p削除Panel As New CPanel農地削除("農地の削除", bCancelable, sMess, "", dt異動日)
        With p削除Panel
            .対象農地 = pTBL
            .n履歴 = n履歴
            .転送先 = p転送先

            If pTBL.Rows.Count > 0 Then
                SysAD.page農家世帯.BlockPanelCtrl2.BlockPanels.Add(p削除Panel, False)
                For Each pRow As DataRow In pTBL.Rows
                    If SysAD.page農家世帯 IsNot Nothing AndAlso SysAD.page農家世帯.DataViewCollection.ContainsKey("農地." & pRow.Item("ID")) Then
                        With SysAD.page農家世帯.DataViewCollection("農地." & pRow.Item("ID"))
                            Select Case .ClosePage()
                                Case HimTools2012.controls.CloseMode.CloseOK
                                    .DoCommand("閉じる")
                                Case HimTools2012.controls.CloseMode.CancelClose
                                    Return 0
                                Case HimTools2012.controls.CloseMode.NoMessage
                            End Select
                        End With
                    End If
                Next

                .Execute()
            End If

            Return .Value
        End With

    End Function


    Public Function 農地削除(ByRef pTBL As DataTable, Optional ByVal n履歴 As Integer = 0, Optional ByVal p転送先 As C農地削除.enum転送先 = C農地削除.enum転送先.削除農地, Optional ByVal sMess As String = "", Optional ByVal dt異動日 As DateTime = Nothing) As Integer
        With New C農地削除(dt異動日)
            .対象農地 = pTBL
            .n履歴 = n履歴
            .転送先 = p転送先
            .sMess = sMess

            If pTBL.Rows.Count > 0 Then
                For Each pRow As DataRow In pTBL.Rows

                    If SysAD.page農家世帯 IsNot Nothing AndAlso SysAD.page農家世帯.DataViewCollection.ContainsKey("農地." & pRow.Item("ID")) Then
                        With SysAD.page農家世帯.DataViewCollection("農地." & pRow.Item("ID"))
                            Select Case .ClosePage()
                                Case HimTools2012.controls.CloseMode.CloseOK
                                    .DoCommand("閉じる")
                                Case HimTools2012.controls.CloseMode.CancelClose
                                    Return 0
                                Case HimTools2012.controls.CloseMode.NoMessage
                            End Select
                        End With


                    End If
                Next


                .Dialog.StartProc(True, True)
                If .Dialog._objException Is Nothing = False Then
                    If .Dialog._objException.Message = "Cancel" Then
                        MsgBox("処理を中止しました。　", , "処理中止")
                    End If
                End If
            End If

            Return .nCount
        End With
        Return 0
    End Function

    Public Function 農地復元(ByVal nID As Decimal, Optional ByVal p復元元 As C農地削除.enum転送先 = C農地削除.enum転送先.削除農地, Optional ByVal sMess As String = "") As Boolean
        SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D:農地Info] WHERE [ID]=" & nID)
        Dim sRet As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D:農地Info SELECT [D_削除農地].* FROM [D_削除農地] WHERE ((([D_削除農地].ID)=" & nID & "));")


        If sRet = "" Or sRet = "OK" Then
            Dim pRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(nID)
            If pRow IsNot Nothing Then
                Dim sRet2 As String = SysAD.DB(sLRDB).ExecuteSQL("DELETE [D_削除農地].ID FROM [D_削除農地] WHERE ((([D_削除農地].ID)=" & nID & "));")

                If sRet2 = "" Or sRet2 = "OK" Then
                    Dim pDelRow As DataRow = App農地基本台帳.TBL削除農地.FindRowByID(nID)
                    App農地基本台帳.TBL削除農地.Rows.Remove(pDelRow)
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Sub 同一地の結合(ByVal s前List As String, ByRef p後 As CObj農地)
        If MsgBox("この処理を行うと従前の農地は削除され、履歴は結合後の農地に統合されます。よろしいですか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            With New HimTools2012.PropertyGridDialog(New C異動日入力(), "異動日入力", "異動日を入力してください。")
                If .ShowDialog = DialogResult.OK Then
                    Dim sWhere As String = "[ID] In (" & Replace(Replace(s前List, "農地.", ""), ";", ",") & ")"
                    Dim p前ID As Decimal = Replace(Replace(s前List, "農地.", ""), ";", ",")
                    Dim p後ID As Decimal = p後.ID
                    If sWhere.Length > 10 Then

                        Do While InStr(sWhere, "(.") > 0 : sWhere = Replace(sWhere, "(.", "(") : Loop
                        Do While InStr(sWhere, ".)") > 0 : sWhere = Replace(sWhere, ".)", ")") : Loop
                        SysAD.page農家世帯.DataViewCollection.Item("農地." & p前ID).ClosePage()
                        SysAD.page農家世帯.DataViewCollection("農地." & p前ID).DoCommand("閉じる")
                        Dim pDelRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(p前ID)
                        App農地基本台帳.TBL農地.Rows.Remove(pDelRow)
                        SysAD.page農家世帯.DataViewCollection.Item("農地." & p後ID).ClosePage()
                        SysAD.page農家世帯.DataViewCollection("農地." & p後ID).DoCommand("閉じる")

                        SysAD.DB(sLRDB).ExecuteSQL("DELETE [D:農地Info].ID FROM [D:農地Info] WHERE [ID] = " & p後ID)  '換地後の農地削除（IDだけ取得できれば良し）
                        SysAD.DB(sLRDB).ExecuteSQL(String.Format("UPDATE [D:農地Info] SET [D:農地Info].ID = {0} WHERE " & sWhere, p後ID))

                        SysAD.DB(sLRDB).ExecuteSQL(String.Format("UPDATE D_土地履歴 SET D_土地履歴.LID = {1} WHERE D_土地履歴.LID = {0};", p前ID, p後ID)) '履歴IDの更新

                        Dim TBL申請 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_申請")
                        For Each pRow As DataRow In TBL申請.Rows
                            Dim p申請 As New CObj申請(pRow, False)

                            If Not IsDBNull(pRow.Item("パラメータリスト")) Then
                                p申請.SetItem("パラメータリスト", Replace(pRow.Item("パラメータリスト"), "<Key>農地." & p前ID & "</Key>", "<Key>農地." & p後ID & "</Key>"))
                                p申請.SaveMyself()
                            End If

                            If pRow.Item("農地リスト").ToString = "農地." & p前ID AndAlso Not IsDBNull(pRow.Item("農地リスト")) Then
                                p申請.SetItem("農地リスト", Replace(pRow.Item("農地リスト"), "農地." & p前ID, "農地." & p後ID))
                                p申請.SaveMyself()
                            End If

                            If InStr(pRow.Item("農地リスト").ToString, "農地." & p前ID & ";") > 0 Then
                                p申請.SetItem("農地リスト", Replace(pRow.Item("農地リスト"), "農地." & p前ID, "農地." & p後ID))
                                p申請.SaveMyself()
                            End If
                        Next

                        Dim dt異動日 As DateTime = CType(.ResultProperty, C異動日入力).異動日
                        Make農地履歴(p後.ID, dt異動日, dt異動日, 土地異動事由.その他, enum法令.職権異動, "同一地の結合")
                    End If

                End If
            End With
        End If
    End Sub

    Public Sub 換地前後関連付け処理(ByVal s前List As String, ByRef p後 As CObj農地)
        If MsgBox("この処理を行うと従前の農地は削除され、履歴は換地後の農地に統合されます。よろしいですか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            With New HimTools2012.PropertyGridDialog(New C異動日入力(), "異動日入力", "異動日を入力してください。")
                If .ShowDialog = DialogResult.OK Then
                    Dim sWhere As String = "[ID] In (" & Replace(Replace(s前List, "農地.", ""), ";", ",") & ")"
                    If sWhere.Length > 10 Then

                        Do While InStr(sWhere, "(.") > 0 : sWhere = Replace(sWhere, "(.", "(") : Loop
                        Do While InStr(sWhere, ".)") > 0 : sWhere = Replace(sWhere, ".)", ")") : Loop

                        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE " & sWhere)
                        Dim pLst As New List(Of String)
                        Dim p土地 As New List(Of String)
                        App農地基本台帳.TBL農地.MergePlus(pTBL)

                        For Each pRowV As DataRowView In New DataView(App農地基本台帳.TBL農地.Body, sWhere, "", DataViewRowState.CurrentRows)
                            Dim sRet As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地系図]([自ID],[元ID],[元土地所在]) VALUES({0},{1},'{2}')", p後.ID, pRowV.Item("ID"), pRowV.Item("土地所在"))
                            If sRet = "OK" OrElse sRet = "" Then
                                pLst.Add(pRowV.Item("ID"))
                                p土地.Add(pRowV.Item("土地所在"))
                            End If
                        Next

                        Dim dt異動日 As DateTime = CType(.ResultProperty, C異動日入力).異動日
                        Make農地履歴(p後.ID, dt異動日, dt異動日, 土地異動事由.換地処理追加, enum法令.換地処理, "[" & Join(p土地.ToArray, ",") & "]より換地後として作成")
                        農地削除(New DataView(App農地基本台帳.TBL農地.Body, "[ID] IN (" & Join(pLst.ToArray, ",") & ")", "", DataViewRowState.CurrentRows).ToTable, 土地異動事由.換地処理削除, C農地削除.enum転送先.削除農地, "[" & p後.ToString & "]へ換地処理", dt異動日)
                    End If
                End If

            End With


        End If
    End Sub

    Private Function AreaConv(ByVal pArea As Decimal) As String
        If Fix(pArea) = pArea Then
            Return Val(pArea).ToString("#,##0")
        Else
            Return Val(pArea).ToString("#,##0.##") '小数点第2位まで表示
        End If

        Return pArea
    End Function

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

    Public Sub List関連申請(ByVal nID As Long)
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE [農地リスト] LIKE '%" & nID & "%'")
        Dim St As New System.Text.StringBuilder

        For Each pRow As DataRow In pTBL.Rows
            If InStr(pRow.Item("農地リスト").ToString & ";", "." & nID & ";") Then
                St.Append(IIF(St.Length > 0, ",", "") & pRow.Item("ID"))
            End If
        Next

        If St.Length > 0 Then
            If pTBL.Rows.Count > 0 Then
                Dim pList As C申請リスト
                Dim sTitle As String = "関連申請"

                If Not SysAD.page農家世帯.TabPageContainKey(sTitle) Then
                    pList = New C申請リスト(SysAD.page農家世帯, sTitle, sTitle)
                    pList.Name = sTitle
                    SysAD.page農家世帯.中央Tab.AddPage(pList)
                Else
                    pList = SysAD.page農家世帯.GetItem(sTitle)
                End If

                Dim sWhere As String = "[ID] IN (" & St.ToString & ")"
                pList.検索開始(sWhere, sWhere)
            Else
                MsgBox("該当する申請がありません", vbInformation, "農地に関連する申請")
            End If
        Else
            MsgBox("該当する申請がありません", vbInformation, "農地に関連する申請")
        End If


    End Sub

End Module

Public Class C農地削除
    Inherits HimTools2012.clsAccessor
    Public 対象農地 As DataTable
    Public 転送先 As enum転送先 = enum転送先.削除農地
    Public n履歴 As Integer = 0
    Public nCount As Integer = 0
    Public sMess As String = ""
    Public 異動日 As DateTime = Nothing

    Public Sub New(Optional ByVal dt異動日 As DateTime = Nothing)
        If Not IsNothing(dt異動日) Then
            異動日 = dt異動日
        Else
            異動日 = Now.Date
        End If

    End Sub

    Public Enum enum転送先
        削除農地 = 1
        転用農地 = 2
    End Enum

    Public Overrides Sub Execute()
        Select Case 転送先
            Case enum転送先.削除農地

        End Select


        Me.Maximum = 対象農地.Rows.Count + 1
        For Each pRow As DataRow In 対象農地.Rows
            Dim sRet As String = ""
            Select Case 転送先
                Case enum転送先.削除農地
                    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_削除農地] WHERE [ID]=" & pRow.Item("ID"))
                    sRet = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_削除農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & pRow.Item("ID") & "));")
                Case enum転送先.転用農地
                    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_転用農地] WHERE [ID]=" & pRow.Item("ID"))
                    sRet = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_転用農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & pRow.Item("ID") & "));")
            End Select

            If sRet = "" Or sRet = "OK" Then
                Dim St As String = SysAD.DB(sLRDB).ExecuteSQL("DELETE [D:農地Info].ID FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & pRow.Item("ID") & "));")
                If St.Length = 0 Or St = "OK" Then
                    Dim ppRow As DataRow = App農地基本台帳.TBL農地.Rows.Find(pRow.Item("ID"))

                    If ppRow IsNot Nothing Then
                        App農地基本台帳.TBL農地.Rows.Remove(ppRow)
                    End If
                    Dim s異動日 As String = String.Format("#{0}/{1}/{2}#", 異動日.Month, 異動日.Day, 異動日.Year)
                    Select Case n履歴
                        Case 261
                            If sMess = "" Then
                                sMess = "地図システムより非農地確定"
                            End If
                            Dim sRetX As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地履歴]([LID],[異動事由],[内容],[異動日],[更新日],[入力日]) VALUES({0},261,'" & sMess & "'," & s異動日 & "," & s異動日 & "," & s異動日 & ")", pRow.Item("ID"))
                        Case 844
                            If sMess = "" Then
                                sMess = "換地処理"
                            End If
                            Dim sRetX As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地履歴]([LID],[異動事由],[内容],[異動日],[更新日],[入力日]) VALUES({0},261,'" & sMess & "'," & s異動日 & "," & s異動日 & "," & s異動日 & ")", pRow.Item("ID"))
                        Case Else
                            If sMess = "" Then
                                sMess = "システムより削除"
                            End If
                            Dim sRetX As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地履歴]([LID],[異動事由],[内容],[異動日],[更新日],[入力日]) VALUES({0},261,'" & sMess & "'," & s異動日 & "," & s異動日 & "," & s異動日 & ")", pRow.Item("ID"))
                    End Select

                    Me.Message = "処理中(" & nCount & "/" & 対象農地.Rows.Count & ")"
                    Me.Value = nCount
                    nCount += 1
                End If
            End If
            If _Cancel Then
                Throw New Exception("Cancel")
                Exit Sub
            End If
        Next
    End Sub
End Class

Public Class C許可書XML
    Public FileName As String
    Public mvarXML As String
    Public Sub New(ByVal sFilename As String)
        FileName = sFilename

        If sFilename.Length > 0 AndAlso IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & sFilename) Then
            mvarXML = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & sFilename)
        End If
    End Sub
    Public Sub XMLReplace(ByVal sWord As String, ByVal value As Object)
        mvarXML = Replace(mvarXML, "{" & sWord & "}", value)
    End Sub
    Public Sub XMLReplaceWithNo(ByVal sWord As String, ByVal Number As Integer, ByVal FigureLength As Integer, ByVal value As Object)
        mvarXML = Replace(mvarXML, "{" & sWord & Strings.Right("000000" & Number, FigureLength) & "}", value)
    End Sub

    Public Sub SaveAndOpen(ByVal sFolder As String, ByVal nNo As Integer, ByVal sName As String)
        HimTools2012.TextAdapter.SaveTextFile(sFolder & Replace(FileName.ToLower, ".xml", "(" & nNo & "_" & sName & ").xml"), mvarXML)
        SysAD.ShowFolder(sFolder)
    End Sub
End Class

