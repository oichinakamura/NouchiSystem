Public Class CフェーズMainPage
    Inherits HimTools2012.SystemWindows.CMainPageSK

    Public Sub New()
        MyBase.New(True, False, "フェーズ２機能一覧", "フェーズ２機能一覧")

        mvarListView.Groups.Add("操作", "操作>>").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("他システム連携", "他システム連携").HeaderAlignment = HorizontalAlignment.Left
        mvarListView.Groups.Add("設定", "設定").HeaderAlignment = HorizontalAlignment.Left

        With Me
            .ListView.ItemAdd("フェーズ2移行用CSV出力(全件)", "フェーズ2移行用CSV出力(全件)", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("フェーズ2移行用CSV出力(農地)", "フェーズ2移行用CSV出力(農地)", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("フェーズ2移行用CSV出力(個人)", "フェーズ2移行用CSV出力(個人)", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("フェーズ2移行用CSV出力(世帯)", "フェーズ2移行用CSV出力(世帯)", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("フェーズ2移行用CSV出力(データ更新)", "フェーズ2移行用CSV出力(データ更新)", "他システム連携", "他システム連携", AddressOf ClickMenu)
            .ListView.ItemAdd("戻る", "戻る", "作業", "操作", AddressOf ClickMenu)
        End With


    End Sub

    Public Sub ClickMenu(ByVal s As Object, ByVal e As EventArgs)
        Select Case CType(s, ListViewItem).Text
            Case "フェーズ2移行用CSV出力(全件)"
                If MsgBox("フェーズ2移行用CSV出力を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    StartOutPut(EnumOutPutType.全件)
                End If
            Case "フェーズ2移行用CSV出力(農地)"
                If MsgBox("フェーズ2移行用CSV出力を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    StartOutPut(EnumOutPutType.農地)
                End If
            Case "フェーズ2移行用CSV出力(個人)"
                If MsgBox("フェーズ2移行用CSV出力を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    StartOutPut(EnumOutPutType.個人)
                End If
            Case "フェーズ2移行用CSV出力(世帯)"
                If MsgBox("フェーズ2移行用CSV出力を開始しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    StartOutPut(EnumOutPutType.世帯)
                End If
            Case "フェーズ2移行用CSV出力(データ更新)"
                If MsgBox("データベースの更新を行いますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    StartOutPut(EnumOutPutType.データ更新)
                End If
            Case "戻る"
                If SysAD.MainForm.MainTabCtrl.ExistPage("Main") Then
                    SysAD.MainForm.MainTabCtrl.TabPages.Remove(Me)
                    Me.Dispose()
                End If
        End Select
    End Sub

    Private Sub StartOutPut(ByVal pOutPutType As Integer)
        Dim pデータ出力 As New CF2データ出力(pOutPutType)
        My.Application.DoEvents()
        If MessageBox.Show("終了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) = DialogResult.OK Then
        End If
    End Sub
End Class

Public Class C各フェーズCSV出力
    Inherits HimTools2012.clsAccessor

    Public Shared TBL個人 As New DataTable
    Public Shared TBL世帯 As New DataTable
    Public Shared TBL農地 As New DataTable
    Public TBL転用農地 As New DataTable
    Public TBL現況地目 As New DataTable
    Public TBL市町村コード As New DataTable
    Public Shared TBL耕作者 As New DataTable
    Public Shared TBL続柄 As New DataTable
    Public Shared TBL住民区分 As New DataTable

    Public 都道府県ID As String = ""
    Public 市町村CD As String = ""
    Public 市町村名 As String = ""
    Public Shared 中間管理機構ID As Decimal = 0

    Public sCSV As StringBEx
    Public sCSV論理 As StringBEx
    Public Shared sCSVレイアウト As StringBEx

    Public 論理Flg As Boolean = False
    Public 論理連番 As Integer = 1
    Public Shared レイアウトFlg As Boolean = False
    Public Shared レイアウト連番 As Integer = 1

    Public Shared RowCount As Integer = 1

    Public Overrides Sub Execute()

    End Sub

    Public Sub ColumnCheck(ByRef pTBL As DataTable, ByVal pColName As String, ByVal pColType As Type)
        If Not pTBL.Columns.Contains(pColName) Then
            pTBL.Columns.Add(pColName, pColType)
        End If
    End Sub

    Public Sub Conv地番(ByRef pRow As DataRow)
        Dim pAddress As String = Replace(pRow.Item("地番").ToString, "の", "")
        Dim s分岐1 As String = "" : Dim s分岐2 As String = "" : Dim s分岐3 As String = "" : Dim s分岐4 As String = ""
        Dim s本番区分 As String = "" : Dim s本番 As String = ""
        Dim s枝番区分 As String = "" : Dim s枝番 As String = ""
        Dim s孫番区分 As String = "" : Dim s孫番 As String = ""

        pAddress = StrConv(pAddress, vbNarrow)

        If InStr(pAddress, "-") > 0 Then        '地番が"-"を含むかどうか
            s本番 = Val(HimTools2012.StringF.Left(pAddress, InStr(pAddress, "-") - 1))
            s分岐1 = Mid(pAddress, InStr(pAddress, "-") + 1)

            If InStr(s分岐1, "-") > 0 Then        '枝番以降が"-"を含むかどうか
                If Char.IsNumber(s分岐1, 0) Then
                    s枝番 = Val(HimTools2012.StringF.Left(s分岐1, InStr(s分岐1, "-") - 1))
                    s分岐2 = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                    If InStr(s分岐2, "-") > 0 Then
                        If Char.IsNumber(s分岐2, 0) Then
                            s孫番 = Val(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                            s孫番区分 = StrConv(Mid(s分岐2, InStr(s分岐2, "-") + 1), VbStrConv.Wide)
                            '終了
                        Else
                            s枝番区分 = StrConv(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                            s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                s孫番 = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                s孫番区分 = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                '終了
                            Else
                                If Char.IsNumber(s分岐3, 0) Then : s孫番 = Val(s分岐3)
                                Else : s孫番区分 = StrConv(s分岐3, VbStrConv.Wide)
                                End If
                                '終了
                            End If
                        End If
                    Else
                        If Char.IsNumber(s分岐2, 0) Then : s孫番 = Val(s分岐2)
                        Else : s枝番区分 = StrConv(s分岐2, VbStrConv.Wide)
                        End If
                        '終了
                    End If
                Else
                    s本番区分 = StrConv(HimTools2012.StringF.Left(s分岐1, InStr(s分岐1, "-") - 1), VbStrConv.Wide)
                    s分岐2 = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                    If InStr(s分岐2, "-") > 0 Then
                        If Char.IsNumber(s分岐2, 0) Then
                            s枝番 = Val(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                            s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                If Char.IsNumber(s分岐3, 0) Then
                                    s孫番 = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                    s孫番区分 = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                    '終了
                                Else
                                    s枝番区分 = StrConv(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1), VbStrConv.Wide)
                                    s分岐4 = Mid(s分岐3, InStr(s分岐3, "-") + 1)

                                    If InStr(s分岐4, "-") > 0 Then
                                        s孫番 = Val(HimTools2012.StringF.Left(s分岐4, InStr(s分岐4, "-") - 1))
                                        s孫番区分 = StrConv(Mid(s分岐4, InStr(s分岐4, "-") + 1), VbStrConv.Wide)
                                    Else
                                        s孫番区分 = StrConv(Mid(s分岐4, InStr(s分岐4, "-") + 1), VbStrConv.Wide)
                                    End If
                                    '終了
                                End If
                            Else
                                If Char.IsNumber(s分岐3, 0) Then : s孫番 = Val(s分岐3)
                                Else : s枝番区分 = StrConv(s分岐3, VbStrConv.Wide)
                                End If
                                '終了
                            End If
                        Else
                            s枝番区分 = StrConv(HimTools2012.StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                            s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                s孫番 = Val(HimTools2012.StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                s孫番区分 = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                '終了
                            Else
                                If Char.IsNumber(s分岐3, 0) Then : s孫番 = Val(s分岐3)
                                Else : s孫番区分 = StrConv(s分岐3, VbStrConv.Wide)
                                End If
                                '終了
                            End If
                        End If
                    Else
                        If Char.IsNumber(s分岐2, 0) Then : s枝番 = Val(s分岐2)
                        Else : s枝番区分 = StrConv(s分岐2, VbStrConv.Wide)
                        End If
                        '終了
                    End If
                End If
            Else
                If Char.IsNumber(s分岐1, 0) Then : s枝番 = Val(s分岐1)
                Else : s本番区分 = StrConv(s分岐1, VbStrConv.Wide)
                End If
                '終了
            End If
        Else
            s本番 = Val(pAddress)
            '終了
        End If

        pRow.Item("本番区分") = s本番区分
        pRow.Item("本番") = IIf(s本番 = "", DBNull.Value, s本番)
        pRow.Item("枝番区分") = s枝番区分
        pRow.Item("枝番") = IIf(s枝番 = "", DBNull.Value, s枝番)
        pRow.Item("孫番区分") = s孫番区分
        pRow.Item("孫番") = IIf(s孫番 = "", DBNull.Value, s孫番)
    End Sub

    Private cityCode As String = ""
    Private otherJusyo As String = ""
    Private kenCity As String = ""
    Public ReadOnly Property 市町村コード(ByVal sValue As String)
        Get
            Dim CodePath As String = ""
            If System.IO.File.Exists(SysAD.CustomReportFolder("共通様式") & "\code_list.csv") Then
                CodePath = SysAD.CustomReportFolder("共通様式") & "\code_list.csv"
            End If

            Dim cityCodeModel As CitiesCode.Interface.ICityCodeModel = New CitiesCode.Factory.CityCodeFactory().CreateCityCodeModel(CodePath)

            Dim jusyoModel As CitiesCode.Interface.IJusyoModel = cityCodeModel.GetCityCode(sValue)  ' 文字列より市町村コード取得
            If jusyoModel.MatchState = CitiesCode.Types.MatchType.Match Then

                ' Match以外はnull
                cityCode = jusyoModel.CityCode
                otherJusyo = jusyoModel.OtherJusyoText  ' その他の住所
                If Len(jusyoModel.OtherJusyoText) > 0 Then
                    kenCity = jusyoModel.JusyoText.Replace(jusyoModel.OtherJusyoText, "")  ' その他の住所以外
                Else
                    kenCity = ""
                End If

                Return cityCode
            End If

            Return ""
        End Get
    End Property

    Public Function Find市町村コード(ByVal sValue As String)
        Dim Find大分類 As String = ""
        Dim Find小分類 As String = ""

        For Each pRow As DataRow In TBL市町村コード.Rows
            If InStr(sValue, pRow.Item("都道府県名（漢字）")) > 0 Then
                Find大分類 = pRow.Item("都道府県名（漢字）")
                Exit For
            End If
        Next

        For Each pRow As DataRow In TBL市町村コード.Rows
            If InStr(sValue, pRow.Item("市区町村名（漢字）")) > 0 Then
                Find小分類 = pRow.Item("市区町村名（漢字）")
                Exit For
            End If
        Next

        Dim pTBL As DataTable
        If Len(Find大分類) > 0 AndAlso Len(Find小分類) > 0 Then
            pTBL = New DataView(TBL市町村コード, String.Format("[都道府県名（漢字）] = '{0}' And [市区町村名（漢字）] = '{1}'", Find大分類, Find小分類), "", DataViewRowState.CurrentRows).ToTable

            If pTBL.Rows.Count = 1 Then
                Dim FindDataRow() As DataRow = pTBL.Select(String.Format("[都道府県名（漢字）] = '{0}' And [市区町村名（漢字）] = '{1}'", Find大分類, Find小分類), "")
                Return FindDataRow(0).Item("団体コード")
            Else
                Return 都道府県ID & 市町村CD
            End If
        ElseIf Len(Find小分類) > 0 Then
            pTBL = New DataView(TBL市町村コード, String.Format("[市区町村名（漢字）] = '{0}'", Find小分類), "", DataViewRowState.CurrentRows).ToTable

            If pTBL.Rows.Count = 1 Then
                Dim FindDataRow() As DataRow = pTBL.Select(String.Format("[市区町村名（漢字）] = '{0}'", Find小分類), "")
                Return FindDataRow(0).Item("団体コード")
            Else
                Return 都道府県ID & 市町村CD
            End If
        Else
            Return 都道府県ID & 市町村CD
        End If

    End Function

    Public Function CnvID(ByVal pData As Object, Optional ByVal ndigits As Integer = 13) As String
        Try
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                If Val(pData.ToString) = 0 Then
                    Return "0"
                ElseIf pData < 0 Then
                    If ndigits = 8 Then
                        Return System.Math.Abs(Val(pData.ToString)).ToString("99000000")
                    Else
                        Return System.Math.Abs(Val(pData.ToString)).ToString("9900000000000")
                    End If
                Else
                    Return pData.ToString
                End If
            Else
                Return "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return "0"
        End Try
    End Function

    Public Function Cnv小字ID(ByVal pData As Object) As String
        Try
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                If pData = 0 Then
                    Return pData.ToString("0")
                ElseIf pData < 0 Then
                    Return System.Math.Abs(Val(pData.ToString)).ToString("9900000")
                Else
                    If Len(pData.ToString) > 7 Then
                        Return Mid(pData.ToString, 1, 4) & Mid(pData.ToString, 6)
                    Else
                        Return pData.ToString
                    End If
                End If
            Else
                Return "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return "0"
        End Try
    End Function

    Public Function Cnv農地ID(ByVal pData As Object) As String
        Try
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                If pData = 0 Then
                    Return pData.ToString("0")
                ElseIf pData < 0 Then
                    If Len(pData.ToString) = 9 Then
                        Return System.Math.Abs(Val(pData.ToString))
                    ElseIf Len(pData.ToString) = 8 Then
                        Return System.Math.Abs(Val(pData.ToString)).ToString("90000000")
                    Else
                        Return System.Math.Abs(Val(pData.ToString)).ToString("99000000")
                    End If
                Else
                    If Len(pData.ToString) > 8 Then
                        Return Mid(pData.ToString, 1, 1) & Mid(pData.ToString, 3)
                    Else
                        Return pData.ToString
                    End If
                End If
            Else
                Return "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return "0"
        End Try
    End Function

    Public sPath As String = ""
    Public Sub 名前を付けて保存(ByVal sCSV As StringBEx, ByVal SaveFileName As String, Optional ByVal OpenDialog As Boolean = False, Optional ByVal OpenFolder As Boolean = False)
        '/***名前を付けて保存***/
        If OpenDialog = True Then
            With New SaveFileDialog
                .FileName = String.Format("{0}.csv", SaveFileName)
                .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"

                If .ShowDialog = DialogResult.OK Then
                    sPath = .FileName
                End If
            End With
        End If

        Dim ArSavePath As Object = Split(sPath, "\")
        For n As Integer = 0 To UBound(ArSavePath)
            If n = 0 Then : sPath = ArSavePath(0)
            ElseIf n = UBound(ArSavePath) Then : sPath = sPath & "\" & String.Format("{0}.csv", SaveFileName)
            Else : sPath = sPath & "\" & ArSavePath(n)
            End If
        Next

        Dim CSVText As New System.IO.StreamWriter(sPath, False, System.Text.Encoding.UTF8)
        CSVText.Write(sCSV.Body.ToString)
        CSVText.Dispose()

        If OpenFolder = True Then
            MsgBox("CSVの出力が完了しました。")
            SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
        End If
    End Sub

    Public Class StringBEx
        Public mvarBody As System.Text.StringBuilder

        Public Sub New(ByVal s初期値 As String, ByVal bOption As EnumCnv, Optional ByVal pConName As String = "", Optional ByVal pRequired As Boolean = False)
            mvarBody = New System.Text.StringBuilder(s初期値)

            'Select Case bOption
            '    Case EnumCnv.全角
            '        If Not pConName = "" Then
            '            全角半角エラー(s初期値, pConName, EnumCnv.全角, pRequired)
            '        End If
            '    Case EnumCnv.半角
            '        If Not pConName = "" Then
            '            全角半角エラー(s初期値, pConName, EnumCnv.半角, pRequired)
            '        End If
            'End Select
        End Sub
        Public ReadOnly Property Body As System.Text.StringBuilder
            Get
                Return mvarBody
            End Get
        End Property

        Public Sub CnvData(ByVal pData As Object, ByVal bOption As EnumCnv, Optional ByVal sCode As Integer = 0, Optional ByVal pColInfo As String = "", Optional ByVal pRequired As Boolean = False)
            If pData Is Nothing Then : mvarBody.Append(",")
            ElseIf IsDBNull(pData) Then
                Select Case bOption
                    Case EnumCnv.全角, EnumCnv.半角, EnumCnv.日付, EnumCnv.氏名 : mvarBody.Append(",")
                    Case Else : mvarBody.Append("," & 0)
                End Select
            Else
                If Len(pData) > 0 Then
                    If IsDate(pData) Then
                        pData = Format(pData, "yyyy/MM/dd")
                    End If
                    pData = Trim(Replace(Replace(pData, Chr(13), ""), Chr(10), "")) '20161115 Trim追加
                End If

                Select Case bOption
                    Case EnumCnv.氏名
                        mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "、"), vbWide))
                        'mvarBody.Append("," & Replace(pData.ToString, ",", "、"))
                    Case EnumCnv.全角
                        'Dim bytesData As Byte() = System.Text.Encoding.UTF8.GetBytes(pData.ToString)
                        'Dim str As String = System.Text.Encoding.UTF8.GetString(bytesData)
                        'mvarBody.Append("," & StrConv(Replace(str, ",", "、"), vbWide))
                        mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "、"), vbWide)) '全角半角エラー(StrConv(pData.ToString, vbWide), pColInfo, EnumCnv.全角, pRequired)
                    Case EnumCnv.半角
                        If pData.ToString = "0.0000" Then : pData = ""
                        End If
                        mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "."), vbNarrow)) '全角半角エラー(StrConv(pData.ToString, vbNarrow), pColInfo, EnumCnv.半角, pRequired)
                    Case EnumCnv.日付
                        If IsDate(pData) AndAlso Not pData = "0:00:00" Then
                            mvarBody.Append("," & Format(CDate(pData), "yyyy/MM/dd"))
                        Else
                            mvarBody.Append(",") '日付エラー(pData, pColInfo)
                        End If
                    Case EnumCnv.面積  'ここで小数点第２位まで
                        pData = Math.Round(Val(pData), 2)
                        mvarBody.Append("," & pData.ToString) '全角半角エラー(pData, pColInfo, EnumCnv.半角, pRequired)
                    Case EnumCnv.登記簿地目 '読み取り専用のため
                        If pData.ToString = "" Then
                            mvarBody.Append("," & 8) '全角半角エラー(8, pColInfo, EnumCnv.半角, pRequired)
                        Else
                            mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "."), vbNarrow)) '全角半角エラー(StrConv(pData.ToString, vbNarrow), pColInfo, EnumCnv.半角, pRequired)
                        End If
                    Case EnumCnv.現況地目  '読み取り専用のため
                        If pData.ToString = "" Then
                            mvarBody.Append("," & 9) '全角半角エラー(9, pColInfo, EnumCnv.半角, pRequired)
                        Else
                            mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "."), vbNarrow)) '全角半角エラー(StrConv(pData.ToString, vbNarrow), pColInfo, EnumCnv.半角, pRequired)
                        End If
                    Case EnumCnv.字コード
                        If Val(pData.ToString) < 1 Then : pData = ""
                        End If
                        mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "."), vbNarrow))
                    Case EnumCnv.選択
                        If Val(pData.ToString) = sCode AndAlso sCode > 0 Then
                            mvarBody.Append("," & 9) '全角半角エラー(9, pColInfo, EnumCnv.半角, pRequired)
                        Else
                            If Val(pData.ToString) = -1 Then : pData = 1
                            End If
                            mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "."), vbNarrow)) '全角半角エラー(StrConv(pData.ToString, vbNarrow), pColInfo, EnumCnv.半角, pRequired)
                        End If
                    Case EnumCnv.金額
                        If Val(pData.ToString) = "0" Then : pData = ""
                        End If
                        mvarBody.Append("," & StrConv(Replace(pData.ToString, ",", "."), vbNarrow))
                    Case Else : mvarBody.Append("," & Replace(pData.ToString, ",", "."))
                End Select
            End If
        End Sub

        Public Function Cnv続柄(ByRef pTBL As DataTable, ByVal pData As Object)
            Dim pRow As DataRow = pTBL.Rows.Find(IIf(Not IsDBNull(pData), pData, 0))
            Dim pValue As String = ""

            If Not pRow Is Nothing Then
                pValue = StrConv(pRow.Item("名称").ToString, vbWide)

                Select Case pValue
                    Case "世帯主" : Return 1
                    Case "妻", "妻の" : Return 2
                    Case "夫" : Return 3
                    Case "妻（未届）" : Return 4
                    Case "夫（未届）" : Return 5
                    Case "子", "長男", "２男", "３男", "４男", "５男", "６男", "７男", "８男", "９男", "長女", "２女", "３女", "４女", "５女", "６女", "７女", "８女", "９女" : Return 6
                    Case "父" : Return 47
                    Case "母" : Return 48
                    Case "兄" : Return 49
                    Case "弟" : Return 50
                    Case "姉" : Return 51
                    Case "妹" : Return 52
                    Case "祖父" : Return 53
                    Case "祖母" : Return 54
                    Case "曾祖父" : Return 55
                    Case "曾祖母" : Return 56
                    Case "養父" : Return 57
                    Case "養母" : Return 58
                    Case "縁故者" : Return 63
                    Case "擬制世帯主" : Return 64
                    Case "同居人" : Return 65
                    Case "その他" : Return 66
                    Case "調査中" : Return 99
                    Case "－", "続柄が空白", "" : Return 0
                    Case Else
                        Debug.Print("続柄:" & pValue)
                        Return 66
                End Select
            Else
                Return 0
            End If
        End Function

        Public Function Cnv異動区分(ByVal pData As Object)
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                Select Case pData
                    Case Is > 4 : Return 0
                    Case Else : Return pData
                End Select
            Else
                Return 0
            End If
        End Function

        Public Function Cnv住民区分(ByRef pTBL As DataTable, ByVal pData As Object)
            Dim pRow As DataRow = pTBL.Rows.Find(pData)
            Dim pValue As String = ""

            If Not pRow Is Nothing Then
                pValue = StrConv(pRow.Item("名称"), vbWide)
                Select Case pValue
                    Case "住民", "住民（外）", "住民票搭載者", "記録住民", "住登者", "市内住民", "戸籍登録者" : Return 0
                    Case "死亡", "死亡者", "死亡所有者", "死亡所有", "死亡（世帯主）" : Return 2
                    Case "その他", "法人", "特徴法人", "普徴法人", "共有", "共有者", "共有名義" : Return 9
                    Case Else
                        Debug.Print("住民区分:" & pValue)
                        Return 1
                End Select
            Else
                Return 9
            End If
        End Function

        Public Sub Cnv地目(ByRef pTBL As DataTable, ByVal pData As Object, Optional ByVal pConName As String = "")

            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                Dim pRow As DataRow = pTBL.Rows.Find(pData)
                If Not pRow Is Nothing Then
                    mvarBody.Append("," & pRow.Item("ID"))
                Else
                    pRow = pTBL.Rows.Find("その他")
                    mvarBody.Append("," & pRow.Item("ID"))
                End If

                'If Not pConName = "" Then
                '    全角半角エラー(pRow.Item("ID"), pConName, EnumCnv.半角, False)
                'End If
            Else
                Dim pRow As DataRow = pTBL.Rows.Find("その他")
                mvarBody.Append("," & pRow.Item("ID"))

                'If Not pConName = "" Then
                '    全角半角エラー(pRow.Item("ID"), pConName, EnumCnv.半角, False)
                'End If
            End If
        End Sub

        Public Function Fnc耕作者整理番号(ByRef pRow As DataRowView, ByRef pTBL As DataTable, ByVal pOption As Integer) As Decimal
            Try
                Dim 耕作者ID As Decimal = 0
                Dim 耕作者名 As String = ""

                If Val(pRow.Item("自小作別").ToString) > 0 Then
                    耕作者ID = Val(pRow.Item("借受人ID").ToString)
                    耕作者名 = Find農家情報(耕作者ID, Enum農家.氏名)
                Else
                    If Val(pRow.Item("管理者ID").ToString) <> 0 Then
                        耕作者ID = Val(pRow.Item("管理者ID").ToString)
                        耕作者名 = Find農家情報(耕作者ID, Enum農家.氏名)
                    Else
                        耕作者ID = Val(pRow.Item("所有者ID").ToString)
                        耕作者名 = Find農家情報(耕作者ID, Enum農家.氏名)
                    End If
                End If

                Dim p耕作者Row As DataRow = pTBL.Rows.Find(耕作者ID)
                If p耕作者Row Is Nothing AndAlso 耕作者ID <> 0 Then
                    SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_公開用個人]([PID],[氏名]) VALUES({0},'{1}')", 耕作者ID, 耕作者名)

                    Dim pLoadTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_公開用個人 WHERE [PID]=" & 耕作者ID)
                    'Dim pLoadTBL As DataTable = New DataView(TBL耕作者, "[PID] = " & 耕作者ID, "", DataViewRowState.CurrentRows).ToTable
                    pTBL.Merge(pLoadTBL)
                    p耕作者Row = pTBL.Rows.Find(耕作者ID)
                End If

                Select Case pOption
                    Case 1 : Return 耕作者ID
                    Case Else : Return Val(p耕作者Row.Item("AutoID").ToString)
                End Select
            Catch ex As Exception
                MsgBox(ex.Message & "筆ID「" & Val(pRow.Item("ID").ToString) & "」の農地に所有者IDが正しく設定されていません。")
                Return Nothing
            End Try
        End Function

        Public Function Fnc適用法令(ByVal pData As Object)
            Dim pValue As Integer = 0

            If IsNumeric(pData) Then
                If pData >= 100 Then
                    pValue = Val(Right(pData, 2))
                Else
                    Select Case pData
                        Case 0 : pValue = 0
                        Case 1 : pValue = 1
                        Case 2 : pValue = 3
                        Case 3 : pValue = 4
                        Case 4 : pValue = 90
                    End Select
                End If
            Else
                pValue = 0
            End If

            Return pValue
        End Function

        Public Function Fnc小作形態(ByVal pData As Object)
            Dim pValue As Integer = 0

            If IsNumeric(pData) Then
                If pData >= 100 Then
                    pValue = Val(Right(pData, 2))
                Else
                    Select Case pData
                        Case 0 : pValue = 0
                        Case 1 : pValue = 5
                        Case 2 : pValue = 4
                        Case 3 : pValue = 10
                        Case 4 : pValue = 1
                        Case 5 : pValue = 2
                        Case 6 : pValue = 3
                        Case 7 : pValue = 6
                        Case 8 : pValue = 10
                        Case 9 : pValue = 10
                    End Select
                End If
            Else
                pValue = 0
            End If

            Return pValue
        End Function

        Public Function Fnc判定(ByRef pData As Object)
            If IsDBNull(pData) = True Then
                pData = 0
            Else
                If pData = 0 Then : pData = 1
                ElseIf pData = 1 Then : pData = 2
                End If
            End If

            Return pData
        End Function

        '20170616Trim追加
        Public Function Find農家情報(ByVal pID As Decimal, ByVal pOption As Enum農家) As String
            Dim pFindRow As DataRow = TBL個人.Rows.Find(pID)
            If Not pFindRow Is Nothing AndAlso Not pID = 0 Then
                Select Case pOption
                    Case Enum農家.氏名
                        Return Trim(pFindRow.Item("氏名").ToString)
                    Case Enum農家.住所
                        Return Trim(pFindRow.Item("住所").ToString)
                    Case Enum農家.郵便番号
                        Return Trim(Replace(Cnv郵便番号(pFindRow.Item("郵便番号").ToString), "〒", ""))
                    Case Else
                        Return ""
                End Select

            Else
                Return ""
            End If
        End Function

        Public Function Find世帯主情報(ByVal pID As Decimal, ByVal pOption As Enum農家) As Decimal
            Dim pFindRow As DataRow = Nothing
            Select Case pOption
                Case Enum農家.世帯ID : pFindRow = TBL個人.Rows.Find(pID)
                Case Enum農家.世帯主ID : pFindRow = TBL世帯.Rows.Find(pID)
                Case Else : Return 0
            End Select

            If Not pFindRow Is Nothing AndAlso Not pID = 0 Then
                Dim pFindRow2 As DataRow = Nothing
                Select Case pOption
                    Case Enum農家.世帯ID : pFindRow2 = TBL個人.Rows.Find(pFindRow.Item("ID"))
                    Case Enum農家.世帯主ID : pFindRow2 = TBL世帯.Rows.Find(pFindRow.Item("ID"))
                    Case Else : Return 0
                End Select

                If Not pFindRow2 Is Nothing Then
                    Select Case pOption
                        Case Enum農家.世帯ID : Return pFindRow2.Item("世帯ID")
                        Case Enum農家.世帯主ID : Return pFindRow2.Item("世帯主ID")
                        Case Else : Return 0
                    End Select
                Else
                    Return 0
                End If
            Else
                Return 0
            End If

            'Dim pFindRow As DataRow = TBL世帯.Rows.Find(pID)
            'If Not pFindRow Is Nothing AndAlso Not pID = 0 Then
            '    Dim pFindRow2 As DataRow = TBL個人.Rows.Find(pFindRow.Item("世帯主ID"))

            '    If Not pFindRow2 Is Nothing Then
            '        Select Case pOption
            '            Case Enum農家.世帯ID : Return pFindRow2.Item("世帯ID")
            '            Case Enum農家.世帯主ID : Return pFindRow2.Item("ID")
            '            Case Else : Return 0
            '        End Select
            '    Else
            '        Return 0
            '    End If
            'Else
            '    Return 0
            'End If
        End Function

        Public Function Find個人ID(ByVal pID As Object, ByVal pName As String) As Decimal
            If Val(pID.ToString) > 0 Then
                Return pID
            Else
                If Not pName = "" Then
                    Dim pView As DataTable = New DataView(TBL個人, "[氏名] = '" & pName & "'", "", DataViewRowState.CurrentRows).ToTable

                    If pView.Rows.Count > 0 Then
                        Dim pRow As DataRow = pView.Rows(0)
                        Return pRow.Item("ID")
                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            End If
        End Function

        Public Sub Find世帯情報(ByVal pData As Object, ByVal s出力条件 As String)
            Select Case s出力条件
                Case "世帯"
                    If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                        Dim pRow As DataRow = TBL個人.Rows.Find(pData)

                        If Not pRow Is Nothing Then
                            mvarBody.Append("," & StrConv(Replace(pRow.Item("電話番号").ToString, ",", "."), vbNarrow))
                            mvarBody.Append("," & StrConv(Replace(pRow.Item("FAX番号").ToString, ",", "."), vbNarrow))
                            mvarBody.Append("," & StrConv(Replace(pRow.Item("メールアドレス").ToString, ",", "."), vbNarrow))
                        Else
                            mvarBody.Append(",")
                            mvarBody.Append(",")
                            mvarBody.Append(",")
                        End If
                    Else
                        mvarBody.Append(",")
                        mvarBody.Append(",")
                        mvarBody.Append(",")
                    End If
                Case "その他"
                    If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                        Dim pRow As DataRow = TBL個人.Rows.Find(pData)

                        If Not pRow Is Nothing Then
                            mvarBody.Append("," & StrConv(Val(pRow.Item("担い手等の区分").ToString), vbNarrow))
                            If Not IsDBNull(pRow.Item("認定日")) AndAlso IsDate(pRow.Item("認定日")) Then
                                mvarBody.Append("," & Format(CDate(pRow.Item("認定日").ToString), "yyyy/MM/dd"))
                            Else
                                mvarBody.Append(",")
                            End If
                            If Not IsDBNull(pRow.Item("新規就農者認定日")) AndAlso IsDate(pRow.Item("新規就農者認定日")) Then
                                mvarBody.Append("," & Format(CDate(pRow.Item("新規就農者認定日").ToString), "yyyy/MM/dd"))
                            Else
                                mvarBody.Append(",")
                            End If
                        Else
                            mvarBody.Append(",")
                            mvarBody.Append(",")
                            mvarBody.Append(",")
                        End If
                    Else
                        mvarBody.Append(",")
                        mvarBody.Append(",")
                        mvarBody.Append(",")
                    End If
            End Select


        End Sub

        Public Sub Cnv利用状況根拠条項(ByVal pRow As DataRowView, ByVal sKey As String, Optional ByVal pConName As String = "")
            If Val(pRow.Item(String.Format("{0}条件農地法", sKey)).ToString) > 0 Then
                Select Case pRow.Item(String.Format("{0}条件農地法", sKey))
                    Case 1 : mvarBody.Append("," & 1)
                    Case 2 : mvarBody.Append("," & 2)
                    Case 3 : mvarBody.Append("," & 3)
                End Select

                'If Not pConName = "" Then
                '    全角半角エラー(pRow.Item(String.Format("{0}条件農地法", sKey)), pConName, EnumCnv.半角, False)
                'End If
            ElseIf Val(pRow.Item(String.Format("{0}条件基盤強化法", sKey)).ToString) > 0 Then
                Select Case pRow.Item(String.Format("{0}条件基盤強化法", sKey))
                    Case 1 : mvarBody.Append("," & 4)
                    Case 2 : mvarBody.Append("," & 5)
                    Case 3 : mvarBody.Append("," & 6)
                End Select

                'If Not pConName = "" Then
                '    全角半角エラー(pRow.Item(String.Format("{0}条件基盤強化法", sKey)), pConName, EnumCnv.半角, False)
                'End If
            Else
                mvarBody.Append("," & 0)
                '全角半角エラー(0, pConName, EnumCnv.半角, False)
            End If
        End Sub

        Public Function Cnv郵便番号(ByVal pData As Object)
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                Select Case Len(pData)
                    Case 8 : Return pData
                    Case 7 : Return Mid(pData, 1, 3) & "-" & Mid(pData, 4, 4)
                    Case Else : Return ""
                End Select
            Else
                Return ""
            End If
        End Function

        Public Function Cnv性別(ByVal pData As Object)
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                Select Case pData
                    Case 0 : Return 1
                    Case 1 : Return 2
                    Case Else : Return ""
                End Select
            Else
                Return 0
            End If
        End Function

        Public Sub Cnv農地所有区分(ByVal pData As Object)
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                Dim pRow As DataRow = TBL個人.Rows.Find(pData)

                If Not pRow Is Nothing Then
                    If pRow.Item("ID") = 中間管理機構ID Then
                        mvarBody.Append("," & 6)
                    Else
                        Select Case Val(pRow.Item("性別").ToString)
                            Case 3 : mvarBody.Append("," & 2)
                            Case Else : mvarBody.Append("," & 1)
                        End Select
                    End If

                    'If IsDBNull(pRow.Item("性別")) Then
                    '    mvarBody.Append("," & 1)
                    'Else
                    '    If pRow.Item("ID") = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) Then
                    '        mvarBody.Append("," & 6)
                    '    Else
                    '        Select Case pRow.Item("性別")
                    '            Case 3 : mvarBody.Append("," & 2)
                    '            Case Else : mvarBody.Append("," & 1)
                    '        End Select
                    '    End If
                    'End If
                Else
                    mvarBody.Append("," & 1)
                End If
            Else
                mvarBody.Append("," & 1)
            End If

        End Sub

        Public Function Cnv届出事由(ByVal pData As Object)
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                If IsNumeric(pData) Then
                    Return pData
                Else
                    If InStr(pData, "設定無") > 0 Then : Return 0
                    ElseIf InStr(pData, "相続") > 0 Then : Return 1
                    ElseIf InStr(pData, "時効") > 0 Then : Return 2
                    ElseIf InStr(pData, "法人合併") > 0 Then : Return 3
                    ElseIf InStr(pData, "その他") > 0 Then : Return 4
                    ElseIf InStr(pData, "調査中") > 0 Then : Return 9
                    ElseIf Len(pData.ToString) > 1 Then : Return 4
                    Else : Return 0
                    End If
                End If
            Else
                Return 0
            End If
        End Function

        Public Function Cnv部分面積(ByVal pRow As DataRowView)
            If Val(pRow.Item("一部現況").ToString) > 0 Then
                If Not IsDBNull(pRow.Item("部分面積")) Then
                    Return pRow.Item("部分面積")
                Else
                    Return pRow.Item("実面積")
                End If
            Else
                Return pRow.Item("登記簿面積")
            End If
        End Function

        Public Sub Set経営規模(ByVal pData As Object)
            Dim pView As DataView = New DataView(TBL農地, String.Format("[所有世帯ID] = {0}", pData), "[ID]", DataViewRowState.CurrentRows)
            Dim s経営面積 As Decimal = 0
            For Each pViewRow As DataRowView In pView
                If Val(pViewRow.Item("自小作別").ToString) = 0 Then
                    s経営面積 += Val(pViewRow.Item("実面積").ToString)
                End If
            Next

            Select Case s経営面積
                Case 0 : mvarBody.Append("," & 0)
                Case 1 To 2999 : mvarBody.Append("," & 1)
                Case 3000 To 4999 : mvarBody.Append("," & 2)
                Case 5000 To 9999 : mvarBody.Append("," & 3)
                Case 10000 To 14999 : mvarBody.Append("," & 4)
                Case 15000 To 19999 : mvarBody.Append("," & 5)
                Case 20000 To 29999 : mvarBody.Append("," & 6)
                Case 30000 To 49999 : mvarBody.Append("," & 7)
                Case 50000 To 99999 : mvarBody.Append("," & 8)  '8のみ範囲エラー（原因不明：取り込みツール側のエラー？）　
                Case Else : mvarBody.Append("," & 9)
            End Select
        End Sub

        Public Function Cnv農業経営者(ByRef pRow As DataRowView)
            'If pRow.Item("世帯ID") = 167 Then
            '    Stop
            'End If

            If pRow.Item("農業経営者") = True Then
                Return 1
            Else
                'D:世帯Infoの世帯主IDと同じ場合、世帯主判定
                If Me.Find世帯主情報(pRow.Item("世帯ID"), Enum農家.世帯主ID) = pRow.Item("ID") Then
                    Return 1
                Else
                    'D:世帯Infoの世帯主が正しい状態でデータとしてある場合、世帯主でないと判定
                    If Me.Find世帯主情報(Me.Find世帯主情報(pRow.Item("世帯ID"), Enum農家.世帯主ID), Enum農家.世帯ID) = pRow.Item("世帯ID") Then
                        Return 0
                    Else
                        If Me.Cnv続柄(TBL続柄, pRow.Item("続柄1")) = 1 And Not Cnv住民区分(TBL住民区分, pRow.Item("住民区分")) = 2 Then
                            Return 1
                        Else
                            Return 0
                        End If
                    End If
                End If
            End If
        End Function

        '/***レイアウトチェック用***/
        Public Shared Function 全角半角エラー(ByRef pValue As Object, ByVal ColumnInfo As String, ByVal eIME As EnumCnv, ByVal pRequired As Boolean)
            '    If Not ColumnInfo = "" Then
            '        Dim Ar As Object = Split(ColumnInfo, ":")
            '        If pValue Is Nothing Then
            '        Else
            '            pValue = Trim(pValue.ToString)
            '            For n As Integer = 1 To Len(pValue)
            '                Dim OneWord As String = Mid(pValue, n, 1)
            '                Dim Encode_JIS As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
            '                Dim StrLength As Integer = OneWord.Length
            '                Dim ByteLength = Encode_JIS.GetByteCount(OneWord)

            '                If StrLength * 2 = ByteLength Then
            '                    If eIME = EnumCnv.半角 Then
            '                        レイアウトFlg = True
            '                        Exit For
            '                    End If
            '                ElseIf StrLength = ByteLength Then
            '                    If eIME = EnumCnv.全角 Then
            '                        レイアウトFlg = True
            '                        Exit For
            '                    End If
            '                End If
            '            Next

            '            If レイアウトFlg = True Then
            '                Dim pLineRow As New StringBEx(レイアウト連番, EnumCnv.半角) ' 連番
            '                pLineRow.Body.Append("," & RowCount) ' 行番号
            '                pLineRow.Body.Append("," & Ar(0)) ' 列番号
            '                pLineRow.Body.Append("," & Ar(1)) ' 項目名
            '                pLineRow.Body.Append("," & pValue) ' 項目内容
            '                Select Case eIME
            '                    Case EnumCnv.全角
            '                        pLineRow.Body.Append(",全角エラー") ' エラー内容
            '                        pLineRow.Body.Append(",1") ' エラーコード
            '                    Case EnumCnv.半角
            '                        pLineRow.Body.Append(",半角エラー") ' エラー内容
            '                        pLineRow.Body.Append(",2") ' エラーコード
            '                End Select

            '                sCSVレイアウト.Body.AppendLine(pLineRow.Body.ToString)

            '                レイアウトFlg = False
            '                レイアウト連番 += 1
            '            End If

            '            If Len(pValue) > Ar(2) Then
            '                Dim pLineRow As New StringBEx(レイアウト連番, EnumCnv.半角) ' 連番
            '                pLineRow.Body.Append("," & RowCount) ' 行番号
            '                pLineRow.Body.Append("," & Ar(0)) ' 列番号
            '                pLineRow.Body.Append("," & Ar(1)) ' 項目名
            '                pLineRow.Body.Append("," & pValue) ' 項目内容
            '                pLineRow.Body.Append(",桁数エラー") ' エラー内容
            '                pLineRow.Body.Append(",3") ' エラーコード

            '                sCSVレイアウト.Body.AppendLine(pLineRow.Body.ToString)

            '                レイアウトFlg = False
            '                レイアウト連番 += 1
            '            End If

            '            If pRequired = True Then
            '                If pValue Is Nothing Or pValue = "" Then
            '                    Dim pLineRow As New StringBEx(レイアウト連番, EnumCnv.半角) ' 連番
            '                    pLineRow.Body.Append("," & RowCount) ' 行番号
            '                    pLineRow.Body.Append("," & Ar(0)) ' 列番号
            '                    pLineRow.Body.Append("," & Ar(1)) ' 項目名
            '                    pLineRow.Body.Append("," & pValue) ' 項目内容
            '                    pLineRow.Body.Append(",必須エラー") ' エラー内容
            '                    pLineRow.Body.Append(",4") ' エラーコード

            '                    sCSVレイアウト.Body.AppendLine(pLineRow.Body.ToString)

            '                    レイアウトFlg = False
            '                    レイアウト連番 += 1
            '                End If
            '            End If
            '        End If
            '    End If
            Return pValue
        End Function

        Public Shared Function 日付エラー(ByRef pValue As Object, ByVal ColumnInfo As String)
            '    If Not ColumnInfo = "" Then
            '        Dim Ar As Object = Split(ColumnInfo, ":")
            '        If pValue Is Nothing Then
            '        Else
            '            If IsDate(pValue) = True Then
            '                Dim pLineRow As New StringBEx(レイアウト連番, EnumCnv.半角) ' 連番
            '                pLineRow.Body.Append("," & RowCount) ' 行番号
            '                pLineRow.Body.Append("," & Ar(0)) ' 列番号
            '                pLineRow.Body.Append("," & Ar(1)) ' 項目名
            '                pLineRow.Body.Append("," & pValue) ' 項目内容
            '                pLineRow.Body.Append(",日付エラー") ' エラー内容
            '                pLineRow.Body.Append(",8") ' エラーコード

            '                sCSVレイアウト.Body.AppendLine(pLineRow.Body.ToString)

            '                レイアウトFlg = False
            '                レイアウト連番 += 1
            '            End If

            '            If Len(pValue) > Ar(2) Then
            '                Dim pLineRow As New StringBEx(レイアウト連番, EnumCnv.半角) ' 連番
            '                pLineRow.Body.Append("," & RowCount) ' 行番号
            '                pLineRow.Body.Append("," & Ar(0)) ' 列番号
            '                pLineRow.Body.Append("," & Ar(1)) ' 項目名
            '                pLineRow.Body.Append("," & pValue) ' 項目内容
            '                pLineRow.Body.Append(",桁数エラー") ' エラー内容
            '                pLineRow.Body.Append(",3") ' エラーコード

            '                sCSVレイアウト.Body.AppendLine(pLineRow.Body.ToString)

            '                レイアウトFlg = False
            '                レイアウト連番 += 1
            '            End If
            '        End If
            '    End If
            Return pValue
        End Function
    End Class

    Public Class 登記簿地目変換
        Inherits DataTable

        Public Sub New()

        End Sub

        Public Function Init() As 登記簿地目変換
            Dim 登記Row As DataRow
            Dim s登記変換リスト As String() = {"1:田", "1:宅地介在田", "1:市街化田", "1:介在田", "2:畑", "2:宅地介在畑", "2:市街化畑", "2:介在畑", "3:牧場", "4:宅地", "4:宅地(農業用施設用地)", "4:準宅地", "4:雑種地介在宅地", "5:山林", "5:宅地介在山林", "5:市街化山林", "5:農地介在山林", _
                                               "6:原野", "6:宅地介在原野", "6:市街化原野", "6:農地介在原野", "7:雑種地", "7:雑種地(農業用施設用地)", "7:雑種地(田畑)", "7:雑種地(山林)", "7:雑種地(田)", "7:雑種地(畑)", "7:雑種地(宅地)", "7:準雑地", "7:雑地", "7:雑種", "7:その他の雑種地", "8:その他", "8:海成り", "8:川成り", "8:官有地", "8:-", "9:公衆用道路", "9:市道", "9:県道", "9:国道", "9:公衆道", "10:公用地", _
                                               "11:公共用地", "12:公園", "13:鉄道用地", "13:鉄軌道用地", "13:複合利用鉄軌道用地", "13:鉄軌道", "14:学校用地", "14:学校敷地", "15:水道用地", "15:水道", "16:用悪水路", "17:池沼", "18:溜池", "18:ため池", "19:墓場", "19:墓地", "20:境内地", _
                                               "21:堤", "21:堤とう", "22:井溝", "23:運河用地", "24:保安林", "25:塩田", "26:鉱泉地", "27:河川敷地", "27:河川敷"}
            Dim Ar As Object = Nothing

            Me.Columns.Add(New DataColumn("ID", GetType(Integer)))
            Me.Columns.Add(New DataColumn("名称", GetType(String)))
            Me.PrimaryKey = New DataColumn() {Me.Columns("名称")}

            For n As Integer = 0 To UBound(s登記変換リスト)
                登記Row = Me.NewRow

                Ar = Split(s登記変換リスト(n), ":")
                登記Row("ID") = Ar(0)
                登記Row("名称") = Ar(1)

                Me.Rows.Add(登記Row)
            Next

            Return Me
        End Function
    End Class

    Public Class 現況地目変換
        Inherits DataTable

        Public Sub New()

        End Sub

        Public Function Init() As 現況地目変換
            Dim 現況Row As DataRow

            Dim s現況変換リスト As String() = {"1:田", "1:宅地介在田", "1:介在田", "1:宅地介田", "1:市街化田", "2:畑", "2:宅地介在畑", "2:介在畑", "2:宅地介畑", "2:市街化畑", "3:樹園地", "3:樹園地(桑)", "3:樹園地(茶)", "3:樹園地(果樹)", "4:牧草放牧地", "4:採草放牧地", "5:宅地", "5:宅地（農施用地）", "5:宅地（農施）", "5:宅地（農業用施設用地）", "5:雑種地介在宅地", "5:防火水槽", _
                                               "6:山林原野", "6:山林・原野", "6:山林", "6:宅地介在山林", "6:農地介在山林", "6:市街化山林", "6:保安林", "6:砂防指定林", "6:原野", "6:宅地介在原野", "6:市街化原野", "6:農地介在原野", "7:雑種地", "7:雑種地他", "7:その他雑種地", "7:その他の雑種地", "7:太陽光発電", "7:準雑地", "7:準雑", "7:雑種地（田畑）", "7:雑種地（農施用）", "7:雑種地（農施用地）", "7:雑種地（田）", "7:雑種地（畑）", "7:雑種地（宅地）", "7:雑（宅地1）", "7:雑（宅地3）", "7:雑（宅地5）", "7:雑（宅地7）", "7:資材地", _
                                               "7:雑種地（山林）", "7:雑種", "7:雑種（農地）", "7:雑種（山林）", "7:雑種（農施）", "7:遊園地", "7:雑地", "8:農業用施設", "8:農業用施設用地", "8:農用施設用地", "9:その他", "9:-", "9:ため池", "9:池沼", "9:溜池・井溝", "9:墓地", "9:共同利用地等", "9:牧場", "9:農地外", "9:公有地", "9:堤とう", "9:堤", "9:堤塘", "9:貯水池", "9:鉱泉地", "9:公民館", "9:公民館用地", "9:境内地", "9:現地なし", "9:現地確認不能", "9:現確不能", _
                                               "101:農家住宅", "101:農家用宅地", "102:一般個人住宅", "103:集合住宅等", "111:道路", "111:私有道路", "111:公衆用道路", "111:公衆道", "111:私道", "111:道路敷", "112:水路・河川", "112:用悪水路", "112:河川敷", "112:河川敷き", "112:河川敷地", "112:水道用地", "112:防火用水", "112:水道", "112:河川区域", "113:鉄道敷地", "113:鉄軌道用地", "113:複合利用鉄軌道", "113:複合利用鉄軌道用地", "113:鉄軌道", "114:砂利採取", _
                                               "121:個人農林業施設", "122:共同農林業施設", "131:鉱工業用地", "141:運輸通信用地", "151:商業サービス", "152:ゴルフ場", "152:ゴルフ場用地", "153:宿泊施設等", "154:その他サービス", _
                                               "161:公共施設", "162:学校用地", "162:学校敷地", "163:公園・運動場", "163:公園", "163:公園緑地", "164:その他公共施設", "171:植林", "181:基盤強化法転用", "191:露天資材置場", "192:露天駐車場"}
            Dim Ar As Object = Nothing

            Me.Columns.Add(New DataColumn("ID", GetType(Integer)))
            Me.Columns.Add(New DataColumn("名称", GetType(String)))
            Me.PrimaryKey = New DataColumn() {Me.Columns("名称")}

            For n As Integer = 0 To UBound(s現況変換リスト)
                現況Row = Me.NewRow

                Ar = Split(s現況変換リスト(n), ":")
                現況Row("ID") = Ar(0)
                現況Row("名称") = Ar(1)

                Me.Rows.Add(現況Row)
            Next

            Return Me
        End Function
    End Class

    Public Enum EnumCnv
        設定無 = 0
        全角 = 1
        半角 = 2
        日付 = 3
        面積 = 4
        登記簿地目 = 5
        現況地目 = 6
        字コード = 7
        外字 = 8
        選択 = 9
        初期値 = 10
        金額 = 11
        氏名 = 12
    End Enum

    Public Enum Enum農家
        氏名 = 1
        住所 = 2
        郵便番号 = 3
        世帯ID = 10
        世帯主ID = 11

    End Enum


End Class

Public Enum EnumOutPutType
    全件 = 0
    農地 = 1
    個人 = 2
    世帯 = 3
    データ更新 = 4
End Enum
