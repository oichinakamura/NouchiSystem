
Imports HimTools2012
Imports HimTools2012.CommonFunc


'/********************************/
'/* 法令に基づく農地の申請処理作成
'/********************************/
Module mod申請データ作成処理
    Public Sub 農地申請データ作成処理(ByVal sSourcelist As String, pDistObj As HimTools2012.TargetSystem.CTargetObjWithView)
        Dim sSelect As String = ""

        If Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) <> 0 Then
            Dim Ar As Object = Split(sSourcelist, ";")
            Dim pRow As HimTools2012.Data.DataRowEx = App農地基本台帳.TBL農地.FindRowByID(GetKeyCode(Ar(0)))

            If Val(pRow.Item("自小作別").ToString) > 0 AndAlso Val(pRow.Item("借受人ID").ToString) = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) Then
                sSelect = "所有権移転(3条)の申請受付;耕作権設定(3条)の申請受付;-;転用を伴う所有権移転(5条)の申請受付;転用を伴う貸借権設定(5条)の申請受付;5条一時転用の申請受付;-;利用権設定(基盤強化法)の申請受付;中間管理機構から借人へ利用権設定;経営基盤法による所有権移転受付;利用権移転(基盤強化法)の申請受付;農地法第3条の3第1項の届け出;職権による所有権移転;職権による相続移転;職権による貸借設定;職権による時効取得設定"
            Else
                If Val(pRow.Item("自小作別").ToString) > 0 AndAlso Val(pRow.Item("経由農業生産法人ID").ToString) = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) Then
                    sSelect = "所有権移転(3条)の申請受付;耕作権設定(3条)の申請受付;-;転用を伴う所有権移転(5条)の申請受付;転用を伴う貸借権設定(5条)の申請受付;5条一時転用の申請受付;-;利用権設定(基盤強化法)の申請受付;中間管理機構を介した利用権設定;中間管理機構を介した利用権設定の受け手変更;中間管理機構を介した所有権移転;経営基盤法による所有権移転受付;利用権移転(基盤強化法)の申請受付;農地法第3条の3第1項の届け出;職権による所有権移転;職権による相続移転;職権による貸借設定;職権による時効取得設定"
                Else
                    sSelect = "所有権移転(3条)の申請受付;耕作権設定(3条)の申請受付;-;転用を伴う所有権移転(5条)の申請受付;転用を伴う貸借権設定(5条)の申請受付;5条一時転用の申請受付;-;利用権設定(基盤強化法)の申請受付;中間管理機構を介した利用権設定;中間管理機構を介した所有権移転;経営基盤法による所有権移転受付;利用権移転(基盤強化法)の申請受付;農地法第3条の3第1項の届け出;職権による所有権移転;職権による相続移転;職権による貸借設定;職権による時効取得設定"
                End If
            End If
        Else
            sSelect = "所有権移転(3条)の申請受付;耕作権設定(3条)の申請受付;転用を伴う所有権移転(5条)の申請受付;転用を伴う貸借権設定(5条)の申請受付;利用権設定(基盤強化法)の申請受付;利用権移転の申請受付;農地法第3条の3第1項の届け出;職権による所有権移転;職権による相続移転;職権による貸借設定;職権による時効取得設定"
        End If

        Dim p申請作成 As New C申請データ作成(OptionSelect(sSelect, "設定する内容を選択してください。"), sSourcelist, pDistObj)
    End Sub

    Public Sub 転用農地申請データ作成処理(ByVal sSourcelist As String, ByVal pDistObj As HimTools2012.TargetSystem.CTargetObjWithView)

        Dim p転用農地 As CObj転用農地 = ObjectMan.GetObject(sSourcelist)
        Dim sSelect As String = ""
        sSelect = "農地法5条所有権移転逆作成;-;転用を伴う所有権移転(5条)の申請受付;転用を伴う貸借権設定(5条)の申請受付;5条一時転用の申請受付"
        Dim p申請作成 As New C申請データ作成(pDistObj, OptionSelect(sSelect, "設定する内容を選択してください。"), p転用農地)
    End Sub


#Region "特殊処理"
    Public Function 事業計画変更(ByRef p申請 As CObj申請, Optional ByVal bOpenWindow As Boolean = True)

        Dim p出し人 As CObj個人 = p申請.GetProperty("Obj申請者A")
        Dim p受け人 As CObj個人 = p申請.GetProperty("Obj申請者B")

        With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p出し人, p受け人, enum法令.事業計画変更), "事業計画変更申請入力")
            Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.事業計画変更, p申請.農地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)

            Rs.Set受付情報(CType(.ResultProperty, C申請入力支援))


            Select Case p申請.法令
                Case enum法令.農地法4条, enum法令.農地法4条一時転用
                    'Rs.Set申請者A(p出し人, p出し人.氏名 & "(事業計画変更)→")
                    'Rs.Set申請者B(p受け人)
                    'Rs.Set申請者C(p出し人)
                    Rs.SetValue("名称", "(事業計画変更)→")
                Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                    'Rs.Set申請者A(p受け人, p受け人.氏名 & "(事業計画変更)→")
                    'Rs.Set申請者B(p出し人)
                    'Rs.Set申請者C(p受け人)
                    Rs.SetValue("名称", "(事業計画変更)→")
            End Select
            Rs.SetValue("予備2", p申請.GetProperty("申請理由A"))
            Rs.SetValue("予備3", p申請.GetProperty("許可年月日"))
            If p申請.農地リスト.Length > 0 Then
                Dim St As String = Replace(Replace(Replace(p申請.GetProperty("農地リスト"), "転用農地.", ""), ";", ","), "農地.", "")
                Dim sSQL As String = "SELECT [V_転用農地].[土地所在],[V_転用農地].[実面積],[V_転用農地].[田面積],[V_転用農地].[畑面積],[V_地目].名称 AS 登記地目,[地目2].名称 AS [現況],[V_農委地目].名称 AS [農委地目]" &
                 " FROM ((([V_転用農地] LEFT JOIN [V_地目] ON [V_転用農地].[登記簿地目]=[V_地目].ID) LEFT JOIN V_地目 AS 地目2 ON [V_転用農地].[現況地目]=[地目2].ID) LEFT JOIN [V_農委地目] ON [V_転用農地].[農委地目ID]=[V_農委地目].ID) " &
                 " WHERE [V_転用農地].[ID] IN (" & St & ")" &
                 " ORDER BY [V_転用農地].[大字ID],[V_転用農地].[小字ID],val(IIf(IsNull([地番]),0,IIf(InStr([地番],'-')=0,Val([地番]),val(Left([地番],InStr([地番],'-')-1)))) *1000 + Val(IIf(IsNull([地番]),0,IIf(InStr([地番],'-')=0,0,val(Mid([地番],InStr([地番],'-')+1))))))"
                Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(sSQL)
                Dim s土地所在 As New List(Of String) '                    St = "": i = 0
                For Each pRow As DataRow In pTable.Rows
                    Dim nArea As Decimal = Val(pRow.Item("田面積").ToString) + Val(pRow.Item("畑面積"))
                    If nArea = 0 Then nArea = Val(pRow.Item("実面積"))

                    s土地所在.Add(pRow.Item("土地所在").ToString & ";" & pRow.Item("登記地目").ToString & ";" & pRow.Item("現況").ToString & ";" & nArea)
                Next
                Rs.SetValue("予備1", Join(s土地所在.ToArray, vbCrLf))
            End If
            Rs.SetValue("現地調査区分", 2)
            Return Rs.InsertInto(bOpenWindow)
        End With

        Return False
    End Function

    'Public Function 事業計画変更(ByRef p申請 As CObj申請, Optional ByVal bOpenWindow As Boolean = True)

    '    Dim p出し人 As CObj個人 = p申請.GetProperty("Obj申請者A") '5条申請の譲渡人
    '    Dim p受け人 As CObj個人 = p申請.GetProperty("Obj申請者B") '5条申請の譲受人

    '    With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p出し人, p受け人, enum法令.事業計画変更), "事業計画変更申請入力")
    '        Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.事業計画変更, p申請.農地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)

    '        Rs.Set受付情報(CType(.ResultProperty, C申請入力支援))
    '        'Rs.Set申請者A(p出し人, p出し人.氏名 & "(事業計画変更)→")
    '        Rs.Set申請者A(p受け人, p受け人.氏名 & "(事業計画変更)→")
    '        Rs.Set申請者B(p受け人)

    '        Select Case p申請.法令
    '            Case enum法令.農地法4条
    '                Rs.Set申請者C(p出し人) '4条申請の申請人
    '            Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
    '                Rs.Set申請者C(p受け人) '5条申請の譲受人
    '        End Select
    '        Rs.SetValue("予備2", p申請.GetProperty("申請理由A"))
    '        Rs.SetValue("予備3", p申請.GetProperty("許可年月日"))
    '        If p申請.農地リスト.Length > 0 Then
    '            Dim St As String = Replace(Replace(Replace(p申請.GetProperty("農地リスト"), "転用農地.", ""), ";", ","), "農地.", "")
    '            Dim sSQL As String = "SELECT [V_転用農地].[土地所在],[V_転用農地].[実面積],[V_転用農地].[田面積],[V_転用農地].[畑面積],[V_地目].名称 AS 登記地目,[地目2].名称 AS [現況],[V_農委地目].名称 AS [農委地目]" &
    '             " FROM ((([V_転用農地] LEFT JOIN [V_地目] ON [V_転用農地].[登記簿地目]=[V_地目].ID) LEFT JOIN V_地目 AS 地目2 ON [V_転用農地].[現況地目]=[地目2].ID) LEFT JOIN [V_農委地目] ON [V_転用農地].[農委地目ID]=[V_農委地目].ID) " &
    '             " WHERE [V_転用農地].[ID] IN (" & St & ")" &
    '             " ORDER BY [V_転用農地].[大字ID],[V_転用農地].[小字ID],val(IIf(IsNull([地番]),0,IIf(InStr([地番],'-')=0,Val([地番]),val(Left([地番],InStr([地番],'-')-1)))) *1000 + Val(IIf(IsNull([地番]),0,IIf(InStr([地番],'-')=0,0,val(Mid([地番],InStr([地番],'-')+1))))))"
    '            Dim pTable As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(sSQL)
    '            Dim s土地所在 As New List(Of String) '                    St = "": i = 0
    '            For Each pRow As DataRow In pTable.Rows
    '                Dim nArea As Decimal = Val(pRow.Item("田面積").ToString) + Val(pRow.Item("畑面積"))
    '                If nArea = 0 Then nArea = Val(pRow.Item("実面積"))

    '                s土地所在.Add(pRow.Item("土地所在").ToString & ";" & pRow.Item("登記地目").ToString & ";" & pRow.Item("現況").ToString & ";" & nArea)
    '            Next
    '            Rs.SetValue("予備1", Join(s土地所在.ToArray, vbCrLf))
    '        End If
    '        Rs.SetValue("現地調査区分", 2)
    '        Return Rs.InsertInto(bOpenWindow)
    '    End With

    '    Return False
    'End Function
#End Region

#Region "農地法申請取り消し(許可済み)"
    Public Function 農地法取消し(ByRef p申請 As CObj申請) As Boolean
        Dim p申請者A As CObj個人 = p申請.GetProperty("Obj申請者A")
        Dim p申請者B As CObj個人 = p申請.GetProperty("Obj申請者B")

        Select Case p申請.法令
            Case enum法令.農地法4条, enum法令.農地法4条一時転用, enum法令.非農地証明願
                If p申請者A IsNot Nothing AndAlso p申請者A.Key IsNot Nothing Then
                    With New HimTools2012.PropertyGridDialog(New C申請取下げ(p申請者A, p申請者B, enum法令.基盤強化法所有権), p申請.名称 & "の取り消し")
                        If .ShowDialog = DialogResult.OK Then
                            Try
                                Dim sList As String = p申請.農地リスト
                                Dim Ar As String() = Split(sList, ";")
                                Dim pDate As DateTime = CType(.ResultProperty, C申請取下げ).届出年月日
                                For Each sKey As String In Ar
                                    If sKey.Length > 0 Then
                                        If sKey.StartsWith("農地.") Then
                                            Dim p農地 As CObj農地 = ObjectMan.GetObject(sKey)
                                            If p農地 IsNot Nothing Then
                                                Make農地履歴(p農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                p農地.SaveMyself()
                                            End If
                                        ElseIf sKey.StartsWith("転用農地.") Then
                                            Dim p転用農地 As CObj転用農地 = ObjectMan.GetObject(sKey)
                                            If p転用農地 IsNot Nothing Then
                                                Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                p転用農地.SaveMyself()

                                                p転用農地.Sub転用農地の復活(p転用農地)
                                            End If
                                        End If
                                    End If
                                Next

                                Dim St As String = p申請.農地リスト
                                St = ";" & St
                                St = Replace(St, ";転用農地.", ";農地.")
                                St = Mid$(St, 2)
                                p申請.ValueChange("農地リスト", St)
                                p申請.ValueChange("状態", enum申請状況.取消し)
                                p申請.ValueChange("予備1", CType(.ResultProperty, C申請取下げ).理由)
                                p申請.ValueChange("取消年月日", CType(.ResultProperty, C申請取下げ).届出年月日)
                                p申請.SaveMyself()
                            Catch ex As Exception

                            End Try
                        End If
                    End With
                Else
                    MsgBox("申請者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
                End If
            Case Else
                If p申請者A IsNot Nothing AndAlso p申請者A.Key IsNot Nothing AndAlso p申請者B IsNot Nothing AndAlso p申請者B.Key IsNot Nothing Then
                    With New HimTools2012.PropertyGridDialog(New C申請取下げ(p申請者A, p申請者B, enum法令.基盤強化法所有権), p申請.名称 & "の取り消し")
                        If .ShowDialog = DialogResult.OK Then
                            Try
                                Dim sList As String = p申請.農地リスト
                                Dim Ar As String() = Split(sList, ";")
                                Dim pDate As DateTime = CType(.ResultProperty, C申請取下げ).届出年月日
                                For Each sKey As String In Ar
                                    If sKey.Length > 0 Then
                                        If sKey.StartsWith("農地.") Then
                                            Dim p農地 As CObj農地 = ObjectMan.GetObject(sKey)
                                            If p農地 IsNot Nothing Then
                                                Select Case p申請.法令
                                                    Case enum法令.農地法3条所有権
                                                        p農地.ValueChange("所有者ID", p申請者A.ID)
                                                        p農地.ValueChange("所有世帯ID", p申請者A.世帯ID)
                                                        Make農地履歴(p農地.ID, pDate, pDate, 642203, enum法令.農地法3条所有権, "農地法３条(所有権)の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                        p農地.SaveMyself()
                                                    Case enum法令.農地法3条耕作権
                                                        p農地.ValueChange("自小作別", 0)
                                                        Make農地履歴(p農地.ID, pDate, pDate, 642203, enum法令.農地法3条耕作権, "農地法３条(耕作権)の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                        p農地.SaveMyself()
                                                    Case enum法令.農地法18条解約, enum法令.合意解約
                                                        p農地.ValueChange("自小作別", 1)
                                                        Make農地履歴(p農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                        p農地.SaveMyself()
                                                    Case Else
                                                        p農地.ValueChange("自小作別", 0)
                                                        Make農地履歴(p農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                        p農地.SaveMyself()
                                                End Select
                                            End If
                                        ElseIf sKey.StartsWith("転用農地.") Then
                                            Dim p転用農地 As CObj転用農地 = ObjectMan.GetObject(sKey)
                                            If p転用農地 IsNot Nothing Then
                                                Select Case p申請.法令
                                                    Case enum法令.農地法4条, enum法令.農地法4条一時転用
                                                        Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, "農地法４条の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                                                        p転用農地.ValueChange("所有者ID", p申請者A.ID)
                                                        p転用農地.ValueChange("所有世帯ID", p申請者A.世帯ID)
                                                        Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, "農地法５条の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                        p転用農地.SaveMyself()
                                                    Case Else
                                                        p転用農地.ValueChange("自小作別", 0)
                                                        Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り消し:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                        p転用農地.SaveMyself()
                                                End Select

                                                p転用農地.Sub転用農地の復活(p転用農地)
                                            End If
                                        End If
                                    End If
                                Next

                                Dim St As String = p申請.農地リスト
                                St = ";" & St
                                St = Replace(St, ";転用農地.", ";農地.")
                                St = Mid$(St, 2)
                                p申請.ValueChange("農地リスト", St)
                                p申請.ValueChange("状態", enum申請状況.取消し)
                                p申請.ValueChange("予備1", CType(.ResultProperty, C申請取下げ).理由)
                                p申請.ValueChange("取消年月日", CType(.ResultProperty, C申請取下げ).届出年月日)
                                p申請.SaveMyself()
                            Catch ex As Exception

                            End Try
                        End If
                    End With
                Else
                    MsgBox("申請者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
                End If
        End Select

        Return False
    End Function
#End Region
#Region "農地法申請取り下げ（受付中）"
    Public Function 農地法取下げ(ByRef p申請 As CObj申請) As Boolean
        Dim p申請者A As CObj個人 = p申請.GetProperty("Obj申請者A")
        Dim p申請者B As CObj個人 = p申請.GetProperty("Obj申請者B")

        Select Case p申請.法令
            Case enum法令.農地法4条, enum法令.農地法4条一時転用, enum法令.非農地証明願
                If p申請者A IsNot Nothing AndAlso p申請者A.Key IsNot Nothing Then
                    With New HimTools2012.PropertyGridDialog(New C申請取下げ(p申請者A, p申請者B, enum法令.基盤強化法所有権), p申請.名称 & "の取下げ")
                        If .ShowDialog = DialogResult.OK Then
                            Try
                                Dim sList As String = p申請.農地リスト
                                Dim Ar As String() = Split(sList, ";")
                                Dim pDate As DateTime = CType(.ResultProperty, C申請取下げ).届出年月日
                                For Each sKey As String In Ar
                                    If sKey.Length > 0 Then
                                        If sKey.StartsWith("農地.") Then
                                            Dim p農地 As CObj農地 = ObjectMan.GetObject(sKey)
                                            If p農地 IsNot Nothing Then
                                                Make農地履歴(p農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                            End If
                                        ElseIf sKey.StartsWith("転用農地.") Then
                                            Dim p転用農地 As CObj転用農地 = ObjectMan.GetObject(sKey)
                                            If p転用農地 IsNot Nothing Then
                                                Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                            End If
                                        End If
                                    End If
                                Next

                                p申請.ValueChange("状態", enum申請状況.取下げ)
                                p申請.ValueChange("予備1", CType(.ResultProperty, C申請取下げ).理由)
                                p申請.ValueChange("取下年月日", CType(.ResultProperty, C申請取下げ).届出年月日)
                                p申請.SaveMyself()
                            Catch ex As Exception

                            End Try
                        End If
                    End With
                Else
                    MsgBox("申請者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
                End If
            Case Else
                If p申請者A IsNot Nothing AndAlso p申請者A.Key IsNot Nothing AndAlso p申請者B IsNot Nothing AndAlso p申請者B.Key IsNot Nothing Then
                    With New HimTools2012.PropertyGridDialog(New C申請取下げ(p申請者A, p申請者B, enum法令.基盤強化法所有権), p申請.名称 & "の取下げ")
                        If .ShowDialog = DialogResult.OK Then
                            Try
                                Dim sList As String = p申請.農地リスト
                                Dim Ar As String() = Split(sList, ";")
                                Dim pDate As DateTime = CType(.ResultProperty, C申請取下げ).届出年月日
                                For Each sKey As String In Ar
                                    If sKey.Length > 0 Then
                                        If sKey.StartsWith("農地.") Then
                                            Dim p農地 As CObj農地 = ObjectMan.GetObject(sKey)
                                            If p農地 IsNot Nothing Then
                                                Select Case p申請.法令
                                                    Case enum法令.農地法3条所有権
                                                        Make農地履歴(p農地.ID, pDate, pDate, 642203, enum法令.農地法3条所有権, "農地法３条(所有権)の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                    Case enum法令.農地法3条耕作権
                                                        Make農地履歴(p農地.ID, pDate, pDate, 642203, enum法令.農地法3条耕作権, "農地法３条(耕作権)の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                    Case Else
                                                        Make農地履歴(p農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                End Select
                                            End If
                                        ElseIf sKey.StartsWith("転用農地.") Then
                                            Dim p転用農地 As CObj転用農地 = ObjectMan.GetObject(sKey)
                                            If p転用農地 IsNot Nothing Then
                                                Select Case p申請.法令
                                                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                                                        Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, "農地法５条の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                    Case enum法令.農地法4条, enum法令.農地法4条一時転用
                                                        Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, "農地法４条の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                    Case Else
                                                        Make農地履歴(p転用農地.ID, pDate, pDate, 642203, p申請.法令, p申請.名称 & "の取り下げ:" & CType(.ResultProperty, C申請取下げ).理由, , 0)
                                                End Select
                                            End If
                                        End If
                                    End If
                                Next

                                p申請.ValueChange("状態", enum申請状況.取下げ)
                                p申請.ValueChange("予備1", CType(.ResultProperty, C申請取下げ).理由)
                                p申請.ValueChange("取下年月日", CType(.ResultProperty, C申請取下げ).届出年月日)
                                p申請.SaveMyself()
                            Catch ex As Exception

                            End Try
                        End If
                    End With
                Else
                    MsgBox("申請者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
                End If
        End Select


        Return False
    End Function
#End Region
    Public Function Open申請Wnd(ByVal bOpenWindow As Boolean, ByRef pRow As DataRow) As Boolean
        'ID As Integer
        If bOpenWindow Then

            Dim pObj申請 As HimTools2012.TargetSystem.CTargetObjWithView = ObjectMan.GetObjectDB("申請", pRow, GetType(CObj申請), True)
            If pObj申請 IsNot Nothing Then
                pObj申請.OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
            Else
                Return False
            End If
        End If
        Return True
    End Function

    Public Function Get総会番号(ByVal dt作成日 As Date) As Integer
        Return DateDiff("m", DateSerial(SysAD.DB(sLRDB).DBProperty("総会基準年"), SysAD.DB(sLRDB).DBProperty("総会基準月"), 20), dt作成日) + 1
    End Function

End Module

Public Class C申請データ作成
    Inherits HimTools2012.clsAccessor
    Private mvar作成コマンド As String
    Private mvarSourceList As String
    Private mvar所有者 As HimTools2012.TargetSystem.CTargetObjectBase
    Private mvar農地 As HimTools2012.TargetSystem.CTargetObjWithView
    Private mvar受け手 As CObj個人 = Nothing

    Public Sub New(ByVal s作成コマンド As String, ByVal sSourceList As String, ByRef pDistObj As HimTools2012.TargetSystem.CTargetObjectBase)
        mvar作成コマンド = s作成コマンド
        mvarSourceList = sSourceList
        mvar農地 = ObjectMan.GetObject(Strings.Left(sSourceList, InStr(sSourceList & ";", ";") - 1))

        If Val(mvar農地.GetProperty("管理者ID").ToString) <> 0 Then
            If MsgBox("対象農地は農地所有者/管理者が登録されています。こちらを申請人として申請を行いますか？", MsgBoxButton.YesNo) = MsgBoxResult.Yes Then
                mvar所有者 = ObjectMan.GetObject("個人." & mvar農地.GetProperty("管理者ID"))
            Else
                mvar所有者 = ObjectMan.GetObject("個人." & mvar農地.GetProperty("所有者ID"))
            End If
        Else
            mvar所有者 = ObjectMan.GetObject("個人." & mvar農地.GetProperty("所有者ID"))
        End If

        Dim StartFlag As Boolean = False
        If Len(mvar作成コマンド) > 0 Then
            If Val(mvar農地.GetProperty("自小作別").ToString) = 0 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.法令, D_申請.名称 FROM D_申請 WHERE (((D_申請.状態)=0) AND ((D_申請.農地リスト) Like '%{0}%') AND ((D_申請.経由法人ID)=0));", sSourceList))
                If pTBL.Rows.Count > 0 Then
                    If MsgBox("対象農地は既存の受付中申請の対象となっています。このまま申請を作成しますか？", MsgBoxButton.YesNo) = MsgBoxResult.Yes Then
                        StartFlag = True
                    End If
                Else
                    StartFlag = True
                End If
            Else
                If MsgBox("対象農地は貸借中となっています。このまま申請を作成しますか？", MsgBoxButton.YesNo) = MsgBoxResult.Yes Then
                    StartFlag = True
                End If
            End If

            If mvar農地.Row.Body.Table.Columns.Contains("特例農地区分") Then
                If Val(mvar農地.Row.Body.Item("特例農地区分").ToString) <> 0 Then
                    If StartFlag = True Then
                        If MsgBox("対象農地は特例農地に区分されています。このまま申請を作成しますか？", MsgBoxButton.YesNo) = MsgBoxResult.Yes Then
                            StartFlag = True
                        Else
                            StartFlag = False
                        End If
                    End If
                End If
            End If
        End If

        If StartFlag = True Then
            If pDistObj IsNot Nothing Then
                Select Case pDistObj.Key.DataClass
                    Case "農家" : mvar受け手 = CType(pDistObj, CObj農家).GetProperty("世帯主")
                    Case "個人" : mvar受け手 = pDistObj
                    Case Else
                        MsgBox("申請農地の受け手を[" & pDistObj.ToString & "]に設定できません", MsgBoxStyle.Critical)
                        Return
                End Select

                If Not IsDBNull(mvar受け手.Row.Body.Item("経営移譲の有無")) Then
                    If mvar受け手.Row.Body.Item("経営移譲の有無") = True Then
                        If MsgBox("受け手の農家は経営移譲「有」に設定されています。このまま申請を作成しますか？", MsgBoxButton.YesNo) = MsgBoxResult.Yes Then
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            Else
            End If
            Me.Start(True, False)
        End If
    End Sub

    Public Sub New(ByVal p元所有者 As CObj個人, ByVal s作成コマンド As String, ByVal p転用農地 As CObj転用農地)
        mvar所有者 = p元所有者
        mvar受け手 = ObjectMan.GetObject("個人." & p転用農地.GetLongIntValue("所有者ID"))
        Dim sSList As New List(Of String)
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [所有者ID]=" & mvar受け手.ID)
        For Each pRow As DataRow In pTBL.Rows
            sSList.Add("転用農地." & pRow.Item("ID"))
        Next
        mvarSourceList = Join(sSList.ToArray, ",")

        Select Case s作成コマンド
            Case "農地法5条所有権移転逆作成" : 農地法5条所有権移転逆作成(mvar所有者, mvarSourceList, mvar受け手, True)
            Case "転用を伴う所有権移転(5条)の申請受付" : 農地法5条所有権移転(mvar受け手, "転用農地." & p転用農地.GetLongIntValue("ID"), p元所有者, True)
            Case "転用を伴う貸借権設定(5条)の申請受付" : 農地法5条賃借権設定(enum法令.農地法5条貸借, mvar受け手, "転用農地." & p転用農地.GetLongIntValue("ID"), p元所有者, True)
            Case "5条一時転用の申請受付" : 農地法5条賃借権設定(enum法令.農地法5条一時転用, mvar受け手, "転用農地." & p転用農地.GetLongIntValue("ID"), p元所有者, True)
        End Select

    End Sub


    Public Overrides Sub Execute()
        Select Case mvar作成コマンド
            Case "所有権移転(3条)の申請受付" : 農地法3条所有権移転()
            Case "耕作権設定(3条)の申請受付" : 農地法3条耕作権移転(mvar所有者, mvarSourceList, mvar受け手)
            Case "転用農地法4条の受付" : 農地法4条設定(enum法令.農地法4条)
            Case "4条一時転用の申請受付" : 農地法4条設定(enum法令.農地法4条一時転用)
            Case "利用権設定(基盤強化法)の申請受付" : 経営基盤法利用権設定(mvar所有者, mvarSourceList, mvar受け手, "", True)
            Case "中間管理機構を介した利用権設定" : 経営基盤法利用権設定(mvar所有者, mvarSourceList, mvar受け手, "中間管理機構", True)
            Case "中間管理機構を介した所有権移転" : 経営基盤法所有権移転(mvar所有者, mvarSourceList, mvar受け手, "中間管理機構", True)
            Case "中間管理機構から借人へ利用権設定" : 経営基盤法利用権設定(mvar所有者, mvarSourceList, mvar受け手, "中間管理機構から貸人", True)
                'Case "中間管理機構から譲受人へ所有権移転"
            Case "経営基盤法による所有権移転受付" : 経営基盤法所有権移転(mvar所有者, mvarSourceList, mvar受け手, "", True)
            Case "利用権移転(基盤強化法)の申請受付", "利用権移転の申請受付"
                If mvar農地.GetProperty("自小作別") < 1 Or mvar農地.GetProperty("借受人ID") = 0 Then
                    MsgBox("貸借設定がされていません。")
                    Exit Sub
                Else
                    経営基盤法利用権移転(mvar農地, mvarSourceList, mvar受け手, True)
                End If
            Case "転用を伴う所有権移転(5条)の申請受付" : 農地法5条所有権移転(mvar所有者, mvarSourceList, mvar受け手, True)
            Case "転用を伴う貸借権設定(5条)の申請受付" : 農地法5条賃借権設定(enum法令.農地法5条貸借, mvar所有者, mvarSourceList, mvar受け手, True)
            Case "5条一時転用の申請受付" : 農地法5条賃借権設定(enum法令.農地法5条一時転用, mvar所有者, mvarSourceList, mvar受け手, True)
            Case "農地法5条所有権移転逆作成"
            Case "職権による時効取得設定" : sub異動from職権所有権移転(mvarSourceList, mvar受け手.GetProperty("世帯ID"), mvar受け手.ID, 0, "「" & mvar受け手.GetProperty("氏名") & "」へ職権による時効取得設定")
            Case "職権による所有権移転" : sub異動from職権所有権移転(mvarSourceList, mvar受け手.GetProperty("世帯ID"), mvar受け手.ID, 0, "「" & mvar受け手.GetProperty("氏名") & "」へ職権による所有権移転")
            Case "職権による相続移転" : sub異動from職権所有権移転(mvarSourceList, mvar受け手.GetProperty("世帯ID"), mvar受け手.ID, 99996, "「" & mvar受け手.GetProperty("氏名") & "」へ職権による相続移転")
                '        Case "職権による貸借設定" : Sub一括貸借設定(mvarSourceList, CDataviewSK_GetProperty2("世帯ID"), DVProperty.ID)
            Case "農地法第3条の3第1項の届け出" : 農地法3条の3第1項(mvar所有者, mvarSourceList, mvar受け手, True)
            Case "中間管理機構を介した利用権設定の受け手変更"
                sub機構を介した利用権設定の受け手変更(mvarSourceList, mvar受け手)
            Case ""
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Stop
                End If
        End Select
    End Sub

#Region "申請人１人"


    Public Shared Function あっせん申出渡(ByRef p農地 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = True)
        Dim p所有者 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("所有者")
        With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p所有者, enum法令.あっせん出手), "あっせん申出(出し手)申請入力")
            If p所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.あっせん出手, p農地.Key.KeyValue, HimTools2012.Data.UPDateMode.AutoUpdate)
                Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p所有者, Nothing, "{0}(あっせん出)")

                Return True
            End If
        End With

        Return False
    End Function
    Public Shared Function あっせん申出受(ByRef p申請者 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = True) As Boolean

        With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p申請者, enum法令.あっせん出手), "あっせん申出(受け手)申請入力")
            If p申請者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.あっせん受手, "", HimTools2012.Data.UPDateMode.AutoUpdate)
                Rs.SetValue("告示日", Now.Date)
                Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p申請者, Nothing, "{0}(あっせん受)")

                Return True
            End If

        End With

        Return False
    End Function
    Public Function 農地法4条設定(ByVal p法令 As enum法令)
        With New HimTools2012.PropertyGridDialog(New C申請入力一人称(mvar所有者, enum法令.農地法4条), "農地法4条(転用・一時転用)申請入力")
            If mvar所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, p法令, mvarSourceList, HimTools2012.Data.UPDateMode.AutoUpdate)
                Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), mvar所有者, Nothing, IIF(p法令 = enum法令.農地法4条, "{0}(4条申請)", "{0}(4条一時転用)"))

                Select Case p法令
                    Case enum法令.農地法4条
                        Rs.SetValue("権利種類", 0)
                    Case enum法令.農地法4条一時転用
                        Rs.SetValue("権利種類", 7)
                End Select

                Return True
            End If
        End With

        Return False
    End Function
    Public Shared Function 農地改良届(ByVal p農地 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False)
        Dim p所有者 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("所有者")
        With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p所有者, enum法令.農地改良届), "農地改良届申請入力")
            If p所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農地改良届, p農地.Key.KeyValue, HimTools2012.Data.UPDateMode.AutoUpdate)
                Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p所有者, Nothing, "{0}(農地改良届)")

                Return True
            End If
        End With

        Return False
    End Function
    Public Shared Function 農地利用目的変更(ByVal p農地 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False)
        Dim p所有者 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("所有者")
        With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p所有者, enum法令.農地利用目的変更), "農地利用目的変更申請入力")
            If p所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農地利用目的変更, p農地.Key.KeyValue, HimTools2012.Data.UPDateMode.AutoUpdate)
                Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p所有者, Nothing, "{0}(農地利用目的変更)")

                Return True
            End If
        End With
        Return False
    End Function

    Public Shared Function 農用地利用計画変更(ByVal p農地 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False)
        Dim p所有者 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("所有者")
        With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p所有者, enum法令.農用地計画変更), "農用地計画変更申請入力")
            If p所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農用地計画変更, p農地.Key.KeyValue, HimTools2012.Data.UPDateMode.AutoUpdate)
                Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p所有者, Nothing, "{0}(農用地計画変更)")

                Return True
            End If
        End With
        Return False
    End Function

    ''' <summary> 解約申請データを作成します。 </summary>
    ''' <param name="s土地リスト"></param>
    ''' <param name="bOpenWindow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function 解約申請(ByVal s土地リスト As String, Optional ByVal bOpenWindow As Boolean = False) As Boolean
        Dim p農地 As HimTools2012.TargetSystem.CTargetObjWithView = ObjectMan.GetObject(s土地リスト)
        If p農地 IsNot Nothing Then
            Dim n法令 As enum法令 = Nothing
            Dim s名称 As String = ""

            Select Case p農地.GetIntegerValue("小作地適用法")
                Case 1, 3 : n法令 = enum法令.農地法18条解約 : s名称 = "18条解約"
                Case Else : n法令 = enum法令.合意解約 : s名称 = "合意解約"
            End Select

            Dim p所有者 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("所有者")
            If Val(p農地.GetProperty("管理者ID").ToString) <> 0 Then
                If MsgBox("対象農地は農地所有者/管理者が登録されています。こちらを申請人として申請を行いますか？", MsgBoxButton.YesNo) = MsgBoxResult.Yes Then
                    p所有者 = ObjectMan.GetObject("個人." & p農地.GetProperty("管理者ID"))
                End If
            End If

            With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p所有者, n法令), s名称 & " 申請入力")
                If p所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, n法令, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)

                    Rs.Set受付情報(CType(.ResultProperty, C申請入力支援))
                    Dim p貸人 As HimTools2012.TargetSystem.CTargetObjWithView = p所有者
                    Dim p受人 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("Obj受人")

                    If Val(p農地.GetProperty("経由農業生産法人ID").ToString) > 0 Then
                        s名称 = s名称 & "_中間管理機構経由"
                    End If

                    Rs.Set申請者A(p貸人, "{0}(" & s名称 & ")→")
                        Rs.Set申請者B(p受人)
                        Return Rs.InsertInto(bOpenWindow)
                    End If
            End With
        End If
        Return False
    End Function

    Public Shared Function 返還申請(ByVal s土地リスト As String, Optional ByVal bOpenWindow As Boolean = False) As Boolean
        Dim p農地 As HimTools2012.TargetSystem.CTargetObjWithView = ObjectMan.GetObject(s土地リスト)
        If p農地 IsNot Nothing Then
            Dim n法令 As enum法令 = enum法令.中間管理機構へ農地の返還
            Dim s名称 As String = "中間管理機構へ農地の返還"
            Dim p所有者 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("所有者")

            With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p所有者, n法令), s名称 & " 申請入力")
                If p所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, n法令, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)

                    Rs.Set受付情報(CType(.ResultProperty, C申請入力支援))
                    Dim p貸人 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("Obj経由農業生産法人")
                    Dim p受人 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("Obj受人")

                    Rs.Set申請者A(p貸人, "{0}(" & s名称 & ")→")
                    Rs.Set申請者B(p受人)
                    Return Rs.InsertInto(bOpenWindow)
                End If
            End With
        End If
        Return False
    End Function

    Public Shared Function 非農地証明願(ByRef p農地 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = True)
        Dim p所有者 As CObj個人 = p農地.GetProperty("所有者")

        With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p所有者, enum法令.非農地証明願), "非農地証明願申請入力")
            If p所有者 IsNot Nothing AndAlso .ShowDialog = DialogResult.OK Then
                Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.非農地証明願, p農地.Key.KeyValue, HimTools2012.Data.UPDateMode.AutoUpdate)
                Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p所有者, Nothing, "{0}(非農地証明願)")

                Return True
            End If
        End With

        Return False
    End Function

    Public Shared Function 買受適格(ByVal p農地 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal sCmd As String, Optional ByVal bOpenWindow As Boolean = False)
        Dim p農家 As HimTools2012.TargetSystem.CTargetObjWithView = p農地.GetProperty("所有者")
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力一人称(p農家, enum法令.農地法3条所有権), "買受適格申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim n法令 As Integer = 0
                    Dim s名称 As String = ""
                    Select Case sCmd
                        Case "耕作目的－公売" : n法令 = enum法令.買受適格耕公 : s名称 = "{0}(買受適格-耕作目的－公売)"
                        Case "耕作目的－競売" : n法令 = enum法令.買受適格耕競 : s名称 = "{0}(買受適格-耕作目的－競売)"
                        Case "転用目的－公売" : n法令 = enum法令.買受適格転公 : s名称 = "{0}(買受適格-耕作目的－公売)"
                        Case "転用目的－競売" : n法令 = enum法令.買受適格転競 : s名称 = "{0}(買受適格-耕作目的－競売)"
                    End Select

                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, n法令, p農地.Key.KeyValue, HimTools2012.Data.UPDateMode.AutoUpdate)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, Nothing, s名称)

                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function
#End Region


#Region "申請人２人"
    Public Function 農地法3条所有権移転() As Boolean
        If mvar所有者 IsNot Nothing AndAlso mvar所有者.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(mvar所有者, mvar受け手, enum法令.農地法3条所有権), "農地法3条所有権移転申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農地法3条所有権, mvarSourceList, HimTools2012.Data.UPDateMode.AutoUpdate)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), mvar所有者, mvar受け手, "{0}(３条所有権)→")

                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function

    Public Function 農地法3条耕作権移転(ByVal p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView) As Boolean
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, enum法令.農地法3条耕作権), "農地法3条耕作権移転申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農地法3条耕作権, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, p受手農家, "{0}(３条耕作権)→")

                    Return True
                End If
            End With
        End If
        Return False
    End Function

    Public Function 農地法3条の3第1項(ByRef p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False) As Boolean
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, enum法令.農地法3条の3第1項), "農地法3条の3第1項申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農地法3条の3第1項, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, p受手農家, "{0}(農地法3条の3第1項)→")

                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function

    Public Function 農地法5条所有権移転(ByRef p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False)
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, enum法令.農地法5条所有権), "農地法5条所有権移転申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農地法5条所有権, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, p受手農家, "{0}(5条所有権)→")
                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function

    Public Function 農地法5条所有権移転逆作成(ByRef p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False)
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, enum法令.農地法5条所有権), "農地法5条所有権移転_申請逆作成")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.農地法5条所有権, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate, 2)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, p受手農家, "{0}(5条所有権)→")
                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function


    Public Function 農地法5条賃借権設定(ByVal p法令 As enum法令, ByRef p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False)
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, p法令), "農地法5条(貸借・一時転用)申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, p法令, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, p受手農家, IIF(p法令 = enum法令.農地法5条貸借, "{0}(5条貸借)→", "{0}(一時転用)→"))

                    Select Case p法令
                        Case enum法令.農地法5条貸借
                            Rs.SetValue("権利種類", 0)
                        Case enum法令.農地法5条一時転用
                            Rs.SetValue("権利種類", 7)
                    End Select


                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function

    Public Shared Function Sub機構を介した利用権設定の受け手変更(ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView) As Boolean
        Dim sList As String = Replace(Replace(s土地リスト, "農地.", ""), ";", ",")
        If sList.EndsWith(",") Then sList = StringF.Left(sList, sList.Length - 1)

        Dim n所有者ID As Long = 0
        Dim n現借受者 As Long = 0

        For Each sID As String In Split(sList, ",")
            Dim pRow As HimTools2012.Data.DataRowEx = App農地基本台帳.TBL農地.FindRowByID(Val(sID))
            If n所有者ID = 0 Then
                n所有者ID = Val(pRow.Item("所有者ID").ToString())
            ElseIf Not n所有者ID = Val(pRow.Item("所有者ID").ToString()) Then
                MsgBox("異なる所有者の土地が含まれています。同時に処理できません。", MsgBoxStyle.Critical)
                Return False
            End If
            If n現借受者 = 0 Then
                n現借受者 = Val(pRow.Item("借受人ID").ToString())
            ElseIf Not n現借受者 = Val(pRow.Item("借受人ID").ToString()) Then
                MsgBox("異なる借受人の土地が含まれています。同時に処理できません。", MsgBoxStyle.Critical)
                Return False
            End If
        Next

        Dim p所有者 As HimTools2012.TargetSystem.CTargetObjectBase = ObjectMan.GetObject("個人." & n所有者ID)
        Dim p現借受者 As HimTools2012.TargetSystem.CTargetObjectBase = ObjectMan.GetObject("個人." & n現借受者)

        With New HimTools2012.PropertyGridDialog(New C借受者変更(p所有者, p現借受者, p受手農家, enum法令.機構を介した利用権設定の受け手変更), "機構を介した利用権設定の受け手の変更")
            If .ShowDialog = DialogResult.OK Then
                With CType(.ResultProperty, C借受者変更)
                    Dim s内容 As New System.Text.StringBuilder()
                    Dim s新受け手 As String = ""

                    s内容.Append(String.Format("機構を介した利用権設定の[{0}]との契約を{1}に解約し、", p現借受者.GetProperty("氏名"), 和暦Format(.前の貸借の終了日)))


                    '        mvar農地 = ObjectMan.GetObject(Strings.Left(sSourceList, InStr(sSourceList & ";", ";") - 1))

                    For Each sID As String In Split(sList, ",")
                        If sID.Length = 0 Then Exit For
                        Dim p農地 As CObj農地 = ObjectMan.GetObject("農地." & sID)
                        Try
                            If p受手農家 IsNot Nothing Then
                                Select Case p受手農家.Key.DataClass
                                    Case "農家"
                                        p農地.ValueChange("借受者ID", p受手農家.GetProperty("世帯主ID"))
                                        p農地.ValueChange("借受世帯ID", p受手農家.GetProperty("ID"))
                                        s新受け手 = p受手農家.GetProperty("世帯主氏名")
                                    Case "個人"
                                        p農地.ValueChange("借受人ID", p受手農家.GetProperty("ID"))
                                        p農地.ValueChange("借受世帯ID", p受手農家.GetProperty("世帯ID"))
                                        s新受け手 = p受手農家.GetProperty("氏名")
                                End Select
                            End If
                            p農地.ValueChange("小作開始年月日", .新しい貸借の開始日)
                            p農地.ValueChange("小作終了年月日", .新しい貸借の終了日)

                            s内容.Append(String.Format("新しく[{0}]と{1}～{2}で設定した。", s新受け手, 和暦Format(.新しい貸借の開始日), 和暦Format(.新しい貸借の終了日)))

                            p農地.ValueChange("小作形態", .新しい貸借の形態)
                            Select Case .新しい貸借の形態
                                Case enum小作形態.賃貸借
                                    p農地.ValueChange("小作料", .新しい10a当たりの小作料)
                                    p農地.ValueChange("小作料単位", "円/10a")
                                Case enum小作形態.使用貸借
                                    p農地.ValueChange("小作料", 0)
                                    p農地.ValueChange("小作料単位", "")
                            End Select


                            p農地.SaveMyself()
                            Make農地履歴(p農地.ID, Now.Date, .変更を受けた日, 444295, enum法令.機構を介した利用権設定の受け手変更, s内容.ToString, , p所有者.ID, p受手農家.ID, 0)

                        Catch ex As Exception
                            MsgBox(ex.Message)
                            Stop
                        End Try

                    Next

                End With

            End If
        End With
        Return False
    End Function


    Public Function 経営基盤法所有権移転(ByRef p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal sCmd As String, ByVal bOpenWindow As Boolean) As Boolean
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, enum法令.基盤強化法所有権), "基盤強化法所有権移転申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.基盤強化法所有権, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)
                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, p受手農家, "{0}(所有権移転)→")
                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function

    Public Shared Function 経営基盤法利用権設定(ByVal p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal sCmd As String, Optional ByVal bOpenWindow As Boolean = False)
        If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
            With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, enum法令.利用権設定), "基盤強化法利用権設定申請入力")
                If .ShowDialog = DialogResult.OK Then
                    Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.利用権設定, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)

                    Dim s名称追加 As String = ""
                    Select Case sCmd
                        Case "中間管理機構"
                            Dim nID As Decimal = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
                            Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(nID)
                            s名称追加 = "_中間管理機構"
                            If pRow Is Nothing Then
                                MsgBox("中間管理機構の設定がありません。経由法人に設定できませんでした。", MsgBoxStyle.Critical)
                            Else
                                Rs.SetValue("経由法人ID", nID)
                            End If
                        Case "中間管理機構から貸人"
                            Dim nID As Decimal = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
                            'p農家 = ObjectMan.GetObject("個人." & Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")))

                            s名称追加 = "_配分計画に基づく中間管理機構から貸人"
                            Rs.SetValue("経由法人ID", nID)

                            Dim p農地 = ObjectMan.GetObject(s土地リスト)
                            If p農地.GetProperty("自小作別") <> 0 Then
                                Rs.SetValue("始期", p農地.GetProperty("小作開始年月日"))
                                Rs.SetValue("終期", p農地.GetProperty("小作終了年月日"))
                                Rs.SetValue("小作料", p農地.GetProperty("小作料"))
                                Rs.SetValue("小作料単位", p農地.GetProperty("小作料単位"))
                                Rs.SetValue("機構配分計画利用配分計画始期日", p農地.GetProperty("小作開始年月日"))
                                Rs.SetValue("機構配分計画利用配分計画終期日", p農地.GetProperty("小作終了年月日"))
                            End If
                        Case Else

                    End Select
                    Rs.SetValue("始期", DateSerial(Year(Now) - (Month(Now) = 12), Month(Now) + (Month(Now) = 12) * 12 + 1, 1))
                    Rs.SetValue("公告年月日", DateSerial(Year(Now), Month(Now), Choose(Month(Now), 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)))

                    If InStr(sCmd, "再設定") Then
                        Dim p農地 = ObjectMan.GetObject(s土地リスト)
                        Rs.SetValue("小作料", p農地.GetProperty("小作料"))
                        Rs.SetValue("小作料単位", p農地.GetProperty("小作料単位"))
                        Rs.SetValue("権利種類", p農地.GetProperty("小作形態"))

                        Rs.SetValue("再設定", True)
                    End If

                    Dim pA As New 申請人共通処理(Rs, CType(.ResultProperty, C申請入力支援), p農家, p受手農家, "{0}(利用権設定" & s名称追加 & ")→")
                    Return True
                End If
            End With
        Else
            MsgBox("所有者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
        End If
        Return False
    End Function

    Public Function 経営基盤法利用権移転(ByRef p申請農地 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s土地リスト As String, ByVal p受手農家 As HimTools2012.TargetSystem.CTargetObjWithView, Optional ByVal bOpenWindow As Boolean = False) As Boolean
        If p申請農地 Is Nothing OrElse p申請農地.GetProperty("自小作別") < 1 OrElse p申請農地.GetProperty("借受人ID") = 0 Then
            MsgBox("貸借設定がされていません。")
            Return False
        Else
            Dim p農家 As CObj個人 = ObjectMan.GetObject("個人." & p申請農地.GetProperty("借受人ID"))
            If p農家 IsNot Nothing AndAlso p農家.Key IsNot Nothing Then
                With New HimTools2012.PropertyGridDialog(New C申請入力二人称(p農家, p受手農家, enum法令.利用権移転), "基盤強化法利用権移転申請入力")
                    If .ShowDialog = DialogResult.OK Then
                        Dim Rs As New C申請追加(App農地基本台帳.TBL申請.NewRow, enum法令.利用権移転, s土地リスト, HimTools2012.Data.UPDateMode.AutoUpdate)

                        Rs.Set受付情報(CType(.ResultProperty, C申請入力支援))
                        Rs.Set申請者A(p農家, "{0}(利用権移転)→")
                        Rs.Set申請者B(p受手農家)
                        Rs.Set申請者C(p申請農地.GetProperty("所有者"))


                        Return Rs.InsertInto(bOpenWindow)
                    End If
                End With
            Else
                MsgBox("借受者情報が正しく取得できませんでした。データを確認してください。", MsgBoxStyle.OkOnly)
            End If
        End If
        Return False

    End Function

#End Region

    Public Class 申請人共通処理
        Inherits HimTools2012.clsAccessor
        Private Rs As C申請追加
        Private mvar農家 As HimTools2012.TargetSystem.CTargetObjectBase
        Private mvar受手 As HimTools2012.TargetSystem.CTargetObjectBase
        Private pRes As C申請入力支援
        Private mvar名称 As String

        Public Sub New(ByRef pRs As C申請追加, ByRef InputDT As C申請入力支援, ByRef p出手 As HimTools2012.TargetSystem.CTargetObjectBase, ByRef p受手 As HimTools2012.TargetSystem.CTargetObjectBase, ByVal s名称 As String)
            Rs = pRs
            pRes = InputDT
            mvar農家 = p出手
            mvar受手 = p受手
            mvar名称 = s名称
            Me.Start(True, True)
        End Sub

        Public Overrides Sub Execute()
            Me.Message = "基本情報を設定しています。"
            Rs.Set受付情報(pRes)
            Me.Message = "譲渡人（貸人）情報を設定しています。"
            Rs.Set申請者A(mvar農家, mvar名称)
            If mvar受手 IsNot Nothing Then
                Me.Message = "譲受人（借人）情報を設定しています。"
                Rs.Set申請者B(mvar受手)
            End If
            Me.Message = "申請入力画面を準備しています。"
            Rs.InsertInto(True)
        End Sub
    End Class
End Class

