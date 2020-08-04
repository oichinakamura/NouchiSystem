'/20160222霧島

Imports HimTools2012.Excel.XMLSS2003

Public Class CTabPage受付公布簿
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Public mvarStart As ToolStripDateTimePickerWithlabel
    Public mvarEnd As ToolStripDateTimePickerWithlabel
    Public WithEvents mvar検索受付 As ToolStripButton
    Public WithEvents mvar検索許可 As ToolStripButton

    Public mvarTabCtrl As HimTools2012.controls.TabControlBase

    Public Sub New()
        MyBase.New(True, True, "受付公布簿", "受付公布簿")

        mvarStart = New ToolStripDateTimePickerWithlabel("対象期間")
        mvarEnd = New ToolStripDateTimePickerWithlabel("～")
        mvarStart.Value = Now.Date
        mvarEnd.Value = Now.Date

        mvar検索受付 = New ToolStripButton("検索開始(受付中)")
        mvar検索許可 = New ToolStripButton("検索開始(許可済)")

        mvarTabCtrl = New HimTools2012.controls.TabControlBase()

        Me.ControlPanel.Add(mvarTabCtrl)

        Me.ToolStrip.Items.AddRange({mvarStart, mvarEnd, mvar検索受付, New ToolStripSeparator, mvar検索許可})
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

    Private Enum Enum状態
        受付 = 0
        許可 = 2
    End Enum

    Private mvarTab3条申請リスト As CTabPage各受付Page
    Private mvarTab4条申請リスト As CTabPage各受付Page
    Private mvarTab5条申請リスト As CTabPage各受付Page
    Private mvarTab非農地証明願リスト As CTabPage各受付Page
    Private mvarTable As DataTable

    Private Sub mvar検索受付_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索受付.Click
        受付公布簿作成(Enum状態.受付)
    End Sub

    Private Sub mvar検索許可_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索許可.Click
        受付公布簿作成(Enum状態.許可)
    End Sub

    Private Sub 受付公布簿作成(ByVal State As Integer)
        If mvarStart.Value < mvarEnd.Value Then
            Dim 申請状態 As String = ""
            Select Case State
                Case Enum状態.受付 : 申請状態 = "受付"
                Case Enum状態.許可 : 申請状態 = "許可"
            End Select
            Dim sWhere As String = String.Format("[法令] IN (30,31,40,50,51,52,602) AND [状態]={6} AND [{7}年月日]>=#{0}/{1}/{2}# AND [{7}年月日]<=#{3}/{4}/{5}#",
                                                mvarStart.Value.Month, mvarStart.Value.Day, mvarStart.Value.Year,
                                                mvarEnd.Value.Month, mvarEnd.Value.Day, mvarEnd.Value.Year,
                                                State, 申請状態)

            mvarTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE " & sWhere & " ORDER BY [法令],[許可番号]")

            mvarTable.Columns.Add(New DataColumn("選択", GetType(Boolean)))

            For Each pRow As DataRow In mvarTable.Rows
                Select Case Val(pRow.Item("法令").ToString)
                    Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                        If mvarTab3条申請リスト Is Nothing Then
                            mvarTab3条申請リスト = New CTabPage各受付Page(String.Format("３条{0}リスト", 申請状態), "", "[法令] IN (30,31)", mvarTable, "農地法第3条申請受付公布簿.xml")
                            mvarTabCtrl.AddPage(mvarTab3条申請リスト)
                        End If
                    Case enum法令.農地法4条, enum法令.農地法4条一時転用
                        If mvarTab4条申請リスト Is Nothing Then
                            mvarTab4条申請リスト = New CTabPage各受付Page(String.Format("４条{0}リスト", 申請状態), "", "[法令] IN (40,42)", mvarTable, "農地法第4条申請受付公布簿.xml")
                            mvarTabCtrl.AddPage(mvarTab4条申請リスト)
                        End If
                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                        If mvarTab5条申請リスト Is Nothing Then
                            mvarTab5条申請リスト = New CTabPage各受付Page(String.Format("５条{0}リスト", 申請状態), "", "[法令] IN (50,51,52)", mvarTable, "農地法第5条申請受付公布簿.xml")
                            mvarTabCtrl.AddPage(mvarTab5条申請リスト)
                        End If
                    Case enum法令.非農地証明願
                        If mvarTab非農地証明願リスト Is Nothing Then
                            mvarTab非農地証明願リスト = New CTabPage各受付Page(String.Format("非農地証明願({0})", 申請状態), "", "[法令] IN (602)", mvarTable, "非農地証明書願い受付公布簿.xml")
                            mvarTabCtrl.AddPage(mvarTab非農地証明願リスト)
                        End If
                    Case Else

                End Select
            Next
        Else
            MsgBox("日付が正しくありません")
        End If
    End Sub

End Class

Public Class CTabPage各受付Page
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Protected nLoop As Integer = -1
    Private WithEvents mvar全選択 As ToolStripButton
    Private WithEvents mvar全解除 As ToolStripButton
    Public WithEvents mvar出力 As ToolStripButton
    Private mvarGrid As HimTools2012.controls.DataGridViewWithDataView
    Private mvarFileName As String = ""

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CancelClose
        End Get
    End Property

    Public Sub New(ByVal sKey As String, ByVal sTitle As String, ByVal sWhere As String, ByRef pTable As DataTable, ByVal sFileName As String)
        MyBase.New(False, True, sKey, IIf(sTitle = "", sKey, sTitle))

        mvarFileName = sFileName
        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView
        mvarGrid.AllowUserToAddRows = False
        AddCol("選択", "選択", New DataGridViewCheckBoxColumn)
        AddCol("受付番号", "受付番号", New DataGridViewTextBoxColumn)
        AddCol("受付年月日", "受付年月日", New DataGridViewTextBoxColumn)
        AddCol("許可番号", "許可番号", New DataGridViewTextBoxColumn)
        AddCol("許可年月日", "許可年月日", New DataGridViewTextBoxColumn)
        AddCol("名称", "名称", New DataGridViewTextBoxColumn)
        mvarGrid.AutoGenerateColumns = False
        mvarGrid.SetDataView(pTable, sWhere, "受付番号")
        Me.ControlPanel.Add(mvarGrid)

        mvar全選択 = New ToolStripButton("全選択")
        mvar全解除 = New ToolStripButton("全解除")
        mvar出力 = New ToolStripButton("受付公布簿出力")

        Me.ToolStrip.Items.AddRange({mvar全選択, mvar全解除, New ToolStripSeparator, mvar出力})
    End Sub
    Public Sub Reset(ByRef pTable As DataTable, ByVal sWhere As String)
        mvarGrid.SetDataView(pTable, sWhere, "受付番号")
    End Sub

    Private Sub AddCol(sHeader As String, sData As String, ByVal pCol As DataGridViewColumn)
        pCol.HeaderText = sHeader
        pCol.DataPropertyName = sData
        mvarGrid.Columns.Add(pCol)
    End Sub

    Protected mvar総合計 As New C筆明細と集計作成

    Private Sub mvar全選択_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar全選択.Click
        For Each pRow As DataRowView In mvarGrid.DataSource
            pRow.Item("選択") = True
        Next
        mvarGrid.Refresh()
    End Sub
    Private Sub mvar全解除_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar全解除.Click
        For Each pRow As DataRowView In mvarGrid.DataSource
            pRow.Item("選択") = False
        Next
        mvarGrid.Refresh()
    End Sub

    Private Sub mvar出力_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar出力.Click
        Dim sFolder As String = SysAD.OutputFolder & String.Format("\受付公布簿{0}_{1}", Now.Year, Now.Month)
        If IO.Directory.Exists(SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "")) Then
            sFolder = SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "") & String.Format("\受付公布簿{0}_{1}", Now.Year, Now.Month)
        End If

        If Not IO.Directory.Exists(sFolder) Then
            IO.Directory.CreateDirectory(sFolder)
        End If

        sub受付公布簿作成("", sFolder, mvarFileName)
    End Sub

    Private Function sub受付公布簿作成(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal sFile As String) As Boolean
        Dim sNewFile As String = HimTools2012.FileManager.NewFileName(sDesktopFolder & "\" & sFile, 0)
        Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)

        Dim pXMLSS As New CXMLSS2003(sXML)
        Dim LoopRows As XMLLoopRows
        nLoop = -1
        For Each pSheet As XMLSSWorkSheet In pXMLSS.WorkBook.WorkSheets.Items.Values
            LoopRows = New XMLLoopRows(pSheet)
            For Each pRow As DataRowView In mvarGrid.DataSource
                If Not IsDBNull(pRow.Item("選択")) AndAlso pRow.Item("選択") = True Then

                    If nLoop = -1 Then
                    Else
                        For Each pXRow As XMLSSRow In LoopRows
                            Dim pCopyRow = pXRow.CopyRow

                            pSheet.Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                            LoopRows.InsetRow += 1
                        Next
                    End If

                    nLoop += 1

                    Dim p申請 As New CObj申請(App農地基本台帳.TBL申請.FindRowByID(pRow.Item("ID")), False)

                    SetNO(pSheet)
                    pSheet.ValueReplace("{受付番号}", pRow.Item("受付番号").ToString)
                    p申請.Replace申請者A(pSheet)
                    p申請.Replace申請者B(pSheet)

                    pSheet.ValueReplace("{受付日}", 和暦Format(pRow.Item("受付年月日"), "gy.M.d", "-"))
                    複数土地設定(pSheet, pRow, Nothing)

                    pSheet.ValueReplace("{許可番号}", Val(pRow.Item("許可番号").ToString))
                    If IsDate(pRow.Item("許可年月日")) Then
                        pSheet.ValueReplace("{許可年月日}", 和暦Format(pRow.Item("許可年月日"), "gy.M.d", "-"))
                    Else
                        pSheet.ValueReplace("{許可年月日}", "")
                    End If

                    pSheet.ValueReplace("{申請者Ａ申請理由}", pRow.Item("申請理由A").ToString)
                    pSheet.ValueReplace("{申請者Ｂ申請理由}", pRow.Item("申請理由B").ToString)

                    Select Case pRow.Item("法令")
                        Case 30, 32
                            pSheet.ValueReplace("{権利関係}", "所有権移転")

                            Select Case Val(pRow.Item("所有権移転の種類").ToString)
                                Case 1 : pSheet.ValueReplace("{形態}", "売買")
                                Case 2 : pSheet.ValueReplace("{形態}", "贈与")
                                Case 3 : pSheet.ValueReplace("{形態}", "交換")
                                Case Else
                                    pSheet.ValueReplace("{形態}", "")
                            End Select
                            pSheet.ValueReplace("{期間}", "-")
                        Case 31, 33
                            pSheet.ValueReplace("{権利関係}", "耕作権設定")

                            Select Case Val(pRow.Item("権利種類"))
                                Case 1 : pSheet.ValueReplace("{形態}", "賃借権")
                                Case 2 : pSheet.ValueReplace("{形態}", "使用貸借権")
                                Case Else
                                    pSheet.ValueReplace("{形態}", "その他")
                            End Select

                            Dim n期間年 As Integer = 999
                            Dim dt始期 As Object = pRow.Item("始期")
                            Dim dt終期 As Object = pRow.Item("終期")

                            If IsDBNull(pRow.Item("期間")) OrElse pRow.Item("期間") = 0 Then
                                If Not IsDBNull(dt始期) AndAlso Not IsDBNull(dt終期) Then
                                    n期間年 = DateDiff(DateInterval.Year, dt始期, dt終期)
                                End If
                            Else
                                n期間年 = pRow.Item("期間")
                            End If
                            pSheet.ValueReplace("{期間}", IIf(n期間年 = 999, "永久", n期間年))
                        Case 40
                        Case 50
                            pSheet.ValueReplace("{権利関係}", "所有権移転")

                            Select Case Val(pRow.Item("所有権移転の種類").ToString)
                                Case 1 : pSheet.ValueReplace("{形態}", "売買")
                                Case 2 : pSheet.ValueReplace("{形態}", "贈与")
                                Case 3 : pSheet.ValueReplace("{形態}", "交換")
                                Case Else
                                    pSheet.ValueReplace("{形態}", "")
                            End Select
                        Case 51, 52
                            Select Case Val(pRow.Item("権利種類").ToString)
                                Case 1 : pSheet.ValueReplace("{権利関係}", "賃借権設定")
                                Case 2 : pSheet.ValueReplace("{権利関係}", "使用貸借権設定")
                                Case Else : pSheet.ValueReplace("{権利関係}", "")
                            End Select
                            pSheet.ValueReplace("{形態}", "")
                    End Select

                    Select Case Val(pRow.Item("農地区分").ToString)
                        Case 1 : pSheet.ValueReplace("{農地区分}", "第一種農地")
                        Case 2 : pSheet.ValueReplace("{農地区分}", "第二種農地")
                        Case 3 : pSheet.ValueReplace("{農地区分}", "第三種農地")
                        Case 3 : pSheet.ValueReplace("{農地区分}", "甲種農地")
                        Case 3 : pSheet.ValueReplace("{農地区分}", "農用地区域内農地")
                        Case 3 : pSheet.ValueReplace("{農地区分}", "第二種農地その他の農地")
                        Case Else
                            pSheet.ValueReplace("{農地区分}", "")
                    End Select
                    pSheet.ValueReplace("{転用目的}", pRow.Item("申請理由A").ToString)
                End If
            Next

            If pXMLSS.WorkBook.WorkSheets.Items.ContainsKey("明細") Then
                With pXMLSS.WorkBook.WorkSheets.Items("明細")
                    mvar総合計.Replace明細総合計(pXMLSS.WorkBook.WorkSheets.Items("明細"))

                End With
            Else
                For Each pS As XMLSSWorkSheet In pXMLSS.WorkBook.WorkSheets.Items.Values
                    mvar総合計.Replace明細総合計(pS)
                Next
            End If
        Next

        Initialization()

        If nLoop > -1 Then
            HimTools2012.TextAdapter.SaveTextFile(sNewFile, pXMLSS.OutPut(True))
            SysAD.ShowFolder(sDesktopFolder)
        Else
            MsgBox("選択された申請がありません")
        End If

        Return True
    End Function
    Private Sub Initialization()
        mvar総合計.田数計 = 0
        mvar総合計.畑数計 = 0
        mvar総合計.樹数計 = 0
        mvar総合計.他数計 = 0

        mvar総合計.田面計 = 0
        mvar総合計.畑面計 = 0
        mvar総合計.樹面計 = 0
        mvar総合計.他面計 = 0

        mvar総合計.Is田内 = False
        mvar総合計.Is畑内 = False
        mvar総合計.Is樹内 = False
        mvar総合計.Is他内 = False

        mvar総合計.田面計内 = 0
        mvar総合計.畑面計内 = 0
        mvar総合計.樹面計内 = 0
        mvar総合計.他面計内 = 0
    End Sub

    Public Function SetNO(ByRef pSheet As XMLSSWorkSheet) As Integer
        pSheet.ValueReplace("{No}", (nLoop + 1).ToString)
        Return nLoop + 1
    End Function

    Protected Function 複数土地設定(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView, ByRef p総括Item As dt総括表, Optional ByVal b再設定 As Boolean = False) As C筆明細と集計作成
        'Try
        Dim sNList As String = pRow.Item("農地リスト").ToString
        Dim s登記地目 As String = ""
        Dim s現況地目 As String = ""
        Dim s自小作別 As String = ""
        Dim s持分 As String = ""

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
        Dim pView As New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sNID & ")", "", DataViewRowState.CurrentRows)

        For Each pRowV As DataRowView In pView
            If InStr("," & sNID & ",", "," & pRowV.Item("ID") & ",") Then
                sNID = Replace("," & sNID & ",", "," & pRowV.Item("ID") & ",", ",")
            End If
            nCount += 1


            s登記地目 = s登記地目 & p案件内集計.R & pRowV.Item("登記簿地目名").ToString
            s現況地目 = s現況地目 & p案件内集計.R & pRowV.Item("現況地目名").ToString
            s自小作別 = s自小作別 & p案件内集計.R & IIf(Val(pRowV.Item("自小作別").ToString) = 0, "自", "小")
            s持分 = s持分 & p案件内集計.R & IIf(Val(pRowV.Item("共有持分分子").ToString) > 0 And Val(pRowV.Item("共有持分分母").ToString) > 0, Val(pRowV.Item("共有持分分子").ToString) & "/" & Val(pRowV.Item("共有持分分母").ToString), "")
            p案件内集計.Set筆情報(pRowV.Row, p総括Item, b再設定)
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
                s自小作別 = s自小作別 & p案件内集計.R & IIf(Val(pRowV.Item("自小作別").ToString) = 0, "自", "小")
                s持分 = s持分 & p案件内集計.R & IIf(Val(pRowV.Item("共有持分分子").ToString) > 0 And Val(pRowV.Item("共有持分分母").ToString) > 0, Val(pRowV.Item("共有持分分子").ToString) & "/" & Val(pRowV.Item("共有持分分母").ToString), "")
                p案件内集計.Set筆情報(pRowV.Row, p総括Item, b再設定)
            Next
        End If

        With p案件内集計
            pSheet.ValueReplace("{筆数計}", nCount)
            pSheet.ValueReplace("{土地の所在}", .明細作成.To土地所在文字列("&#10;"))
            pSheet.ValueReplace("{地目}", s登記地目)
            pSheet.ValueReplace("{登記地目}", s登記地目)
            pSheet.ValueReplace("{現況地目}", s現況地目)
            pSheet.ValueReplace("{自小作別}", s自小作別)
            pSheet.ValueReplace("{持分}", s持分)

            'p案件内集計.Replace案件毎集計(pSheet)
            mvar総合計.Add合計(p案件内集計)

            pSheet.ValueReplace("{面積}", .明細作成.To面積文字列("&#10;"))
            pSheet.ValueReplace("{面積内}", .明細作成.To面積内文字列("&#10;"))
        End With

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
                Return 田面計 + 畑面計 + 樹面計 + 他面計
            End Get
        End Property

        Public Sub Replace明細総合計(ByRef pSheet As XMLSSWorkSheet)
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

        Public Sub Set筆情報(ByVal pRow As DataRow, ByVal p総括Item As dt総括表, ByVal b再設定 As Boolean)
            Dim n田面積 As Decimal = 0
            Dim n田面積内 As Decimal = 0
            Dim n畑面積 As Decimal = 0
            Dim n畑面積内 As Decimal = 0
            Dim n樹面積 As Decimal = 0
            Dim n樹面積内 As Decimal = 0
            Dim n他面積 As Decimal = 0
            Dim n他面積内 As Decimal = 0

            If pRow.Item("登記簿地目名").ToString.IndexOf("田") > -1 AndAlso pRow.Item("登記簿地目名").ToString <> "塩田" Then
                田数計 += 1
                n田面積 += Val(pRow.Item("登記簿面積").ToString)
                n田面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is田内 = Is田内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                田面計 += n田面積
                田面計内 += n田面積内
            ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("畑") > -1 Then
                畑数計 += 1
                n畑面積 += Val(pRow.Item("登記簿面積").ToString)
                n畑面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is畑内 = Is畑内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                畑面計 += n畑面積
                畑面計内 += n畑面積内
            ElseIf pRow.Item("登記簿地目名").ToString.IndexOf("樹") > -1 Then
                樹数計 += 1
                n樹面積 += Val(pRow.Item("登記簿面積").ToString)
                n樹面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is樹内 = Is樹内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                樹面計 += n樹面積
                樹面計内 += n樹面積内
            Else
                他数計 += 1
                n他面積 += Val(pRow.Item("登記簿面積").ToString)
                n他面積内 = IIf(Val(pRow.Item("部分面積").ToString) > 0, Val(pRow.Item("部分面積").ToString), Val(pRow.Item("登記簿面積").ToString))
                Is他内 = Is他内 AndAlso (Val(pRow.Item("部分面積").ToString) > 0)

                他面計 += n他面積
                他面計内 += n他面積内
            End If

            明細作成.Plus(pRow, Val(pRow.Item("登記簿面積").ToString), Val(pRow.Item("部分面積").ToString), n田面積, n畑面積)
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

            Public Function To面積文字列(Optional ByVal bySp As String = ",") As String
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
            Public Function To面積内文字列(Optional ByVal bySp As String = ",") As String
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

            Public Function To土地所在文字列(Optional ByVal bySp As String = ",") As String
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
End Class



