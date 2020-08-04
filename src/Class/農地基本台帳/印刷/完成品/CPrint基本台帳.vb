Imports HimTools2012.NumericFunctions

Public Enum ExcelViewMode
    AutoPrint = 0
    Preview = 1
    EditMode = 2
End Enum

Public Enum 印刷Mode
    簡易印刷 = 0
    フル印刷 = 1
End Enum

Public Enum enum農地関連Mode
    関連なし = 0
    自作 = 1
    借受 = 2
    貸付 = 3
End Enum

Public Class CPrint基本台帳
    Inherits C印刷Accessor

    Public 世帯ID As Long = 0
    Public 個人ID As Long = 0

    Private mvarXML As HimTools2012.Excel.XMLSS2003.CXMLSS2003

    Public Property mvar人情報Dic As New Dictionary(Of Long, 人情報)

    Public pTBL所有農地 As DataView
    Public pTBL借受農地 As DataView
    Public pTBL貸付農地 As DataView
    Public pTBL関係者 As DataTable
    'Public pTBL耕作者番号 As DataTable
    Private mvar総括表 As C総括表

    Private mvar印刷Mode As 印刷Mode = 印刷Mode.フル印刷 = 0
    Private mvarExecuteMode As ExcelViewMode
    Public Property HasLand As Boolean = False

    Public Sub New(ByVal pXML As HimTools2012.Excel.XMLSS2003.CXMLSS2003, ByVal nID As Long, ByVal PID As Long, ByVal pPrintMode As 印刷Mode, ByVal pExcelViewMode As ExcelViewMode)
        MyBase.New()

        mvarXML = pXML

        mvarExecuteMode = pExcelViewMode
        mvar印刷Mode = pPrintMode
        Me.世帯ID = nID
        Me.個人ID = PID
    End Sub

    Public Property XMLSS As HimTools2012.Excel.XMLSS2003.CXMLSS2003
        Get
            Return mvarXML
        End Get
        Set(ByVal value As HimTools2012.Excel.XMLSS2003.CXMLSS2003)
            mvarXML = value
        End Set
    End Property

    Public Overrides Sub Execute()
        Me.DataInit()
        Value = 33

        Me.MakeXMLFile()
        Value = 90

    End Sub

    Public Sub DataInit()
        Dim mvar個人ID As New IDList
        mvar総括表 = New C総括表

        Message = "データ取り込み中.."
        With SysAD.DB(sLRDB)

            Dim pTbl世帯 = .GetTableBySqlSelect("SELECT * FROM [D:世帯Info] WHERE [ID]<>0 AND [D:世帯Info].[ID]=" & Me.世帯ID & ";")
            If pTbl世帯.Rows.Count = 0 Then
                If Me.個人ID = 0 Then

                    MsgBox("住民番号が不正です。入力されたデータが正しくありません。内容を確認してください。", MsgBoxStyle.Critical)
                    Exit Sub
                Else
                    mvar個人ID.Add(Me.個人ID)
                End If
                Dim p個人 As CObj個人 = ObjectMan.GetObject("個人." & Me.個人ID)

                For Each pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet In mvarXML.WorkBook.WorkSheets.Items.Values
                    pSheet.ValueReplace("{世帯主氏名}", p個人.氏名.ToString)
                    pSheet.ValueReplace("{住所}", p個人.住所.ToString)
                    pSheet.ValueReplace("{郵便番号}", p個人.GetItem("郵便番号").ToString)
                    pSheet.ValueReplace("{電話番号}", p個人.GetItem("電話番号").ToString)
                    pSheet.ValueReplace("{集落名}", "")
                    pSheet.ValueReplace("{番号}", "")
                    pSheet.ValueReplace("{発行日}", IIf(SysAD.DB(sLRDB).DBProperty("印刷日時の表示") = True, 和暦Format(Now), ""))

                    pSheet.ValueReplace("{集積協力金の有無}", IIf(Val(p個人.GetItem("集積協力金の有無").ToString) = 1, "〇", "×"))
                    pSheet.ValueReplace("{集積協力金開始時期}", p個人.GetItem("集積協力金開始時期").ToString)
                    pSheet.ValueReplace("{転換協力金の有無}", IIf(Val(p個人.GetItem("転換協力金の有無").ToString) = 1, "〇", "×"))
                    pSheet.ValueReplace("{転換協力金開始時期}", p個人.GetItem("転換協力金開始時期").ToString)
                Next

                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT * FROM [D:農地Info] WHERE [所有者ID]={0} Or [借受人ID]={0}", Me.個人ID))
                App農地基本台帳.TBL農地.MergePlus(pTBL)

                If SysAD.市町村.市町村名 = "三股町" Then 'テスト
                    pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Me.個人ID & " AND [自小作別]<1) OR ([所有者ID]=" & Me.個人ID & " AND [借受人ID]=" & Me.個人ID & " AND [自小作別]>0)", "", DataViewRowState.CurrentRows)
                    pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受人ID]=" & Me.個人ID & ") AND ([所有者ID]<>" & Me.個人ID & ") AND [自小作別]>0", "", DataViewRowState.CurrentRows)
                    pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Me.個人ID & ") AND ([借受人ID]<>" & Me.個人ID & ") AND ([自小作別]>0)", "", DataViewRowState.CurrentRows)
                ElseIf Val(SysAD.DB(sLRDB).DBProperty("基本台帳管内農地のみ").ToString) = 1 Then
                    pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Me.個人ID & ") AND [自小作別]<1 AND [大字ID]>0", "", DataViewRowState.CurrentRows)
                    pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受人ID]=" & Me.個人ID & ") AND [自小作別]>0 AND [大字ID]>0", "", DataViewRowState.CurrentRows)
                    pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Me.個人ID & ") AND ([借受人ID]<>" & Me.個人ID & ") AND ([自小作別]>0) AND ([大字ID]>0)", "", DataViewRowState.CurrentRows)
                Else
                    pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Me.個人ID & ") AND [自小作別]<1", "", DataViewRowState.CurrentRows)
                    pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受人ID]=" & Me.個人ID & ") AND [自小作別]>0", "", DataViewRowState.CurrentRows)
                    pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有者ID]=" & Me.個人ID & ") AND ([借受人ID]<>" & Me.個人ID & ") AND ([自小作別]>0)", "", DataViewRowState.CurrentRows)
                End If

            Else
                App農地基本台帳.TBL世帯.MergePlus(pTbl世帯)
                Dim pRow世帯 As DataRow = App農地基本台帳.TBL世帯.Rows.Find(Me.世帯ID)

                mvar個人ID.Add(pRow世帯.Item("世帯主ID"))

                Dim p世帯主 As CObj個人 = ObjectMan.GetObject("個人." & pRow世帯.Item("世帯主ID"))
                For Each pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet In mvarXML.WorkBook.WorkSheets.Items.Values
                    pSheet.ValueReplace("{世帯主氏名}", pRow世帯.Item("世帯主氏名").ToString)
                    pSheet.ValueReplace("{住所}", pRow世帯.Item("住所").ToString)
                    pSheet.ValueReplace("{郵便番号}", pRow世帯.Item("世帯主郵便番号").ToString)
                    pSheet.ValueReplace("{電話番号}", pRow世帯.Item("世帯主電話番号").ToString)
                    pSheet.ValueReplace("{集落名}", pRow世帯.Item("世帯主行政区名").ToString)
                    pSheet.ValueReplace("{番号}", Strings.Right("      " & pRow世帯.Item("農家番号").ToString, 14))
                    pSheet.ValueReplace("{発行日}", IIf(SysAD.DB(sLRDB).DBProperty("印刷日時の表示") = True, 和暦Format(Now), ""))

                    pSheet.ValueReplace("{集積協力金の有無}", IIf(Val(p世帯主.GetItem("集積協力金の有無").ToString) = 1, "〇", "×"))
                    pSheet.ValueReplace("{集積協力金開始時期}", p世帯主.GetItem("集積協力金開始時期").ToString)
                    pSheet.ValueReplace("{転換協力金の有無}", IIf(Val(p世帯主.GetItem("転換協力金の有無").ToString) = 1, "〇", "×"))
                    pSheet.ValueReplace("{転換協力金開始時期}", p世帯主.GetItem("転換協力金開始時期").ToString)
                Next
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT * FROM [D:農地Info] WHERE [所有世帯ID]={0} Or [借受世帯ID]={0}", Me.世帯ID))
                App農地基本台帳.TBL農地.MergePlus(pTBL)

                If SysAD.市町村.市町村名 = "三股町" Then 'テスト
                    pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & 世帯ID & " AND [自小作別]<1) OR ([所有世帯ID]=" & 世帯ID & " AND [借受世帯ID]=" & 世帯ID & " AND [経由農業生産法人ID]= " & Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) & "  AND [自小作別]>0 AND [所有者ID] = [借受人ID])", "", DataViewRowState.CurrentRows)
                    pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受世帯ID]=" & 世帯ID & " AND [所有世帯ID]<>" & 世帯ID & " AND [自小作別]>0) OR ([借受世帯ID]=" & 世帯ID & " AND [所有世帯ID]=" & 世帯ID & " AND [所有者ID] <> [借受人ID] AND [自小作別]>0)", "", DataViewRowState.CurrentRows)
                    pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & 世帯ID & " AND [借受世帯ID]<>" & 世帯ID & " AND [自小作別]>0) OR ([所有世帯ID]=" & 世帯ID & " AND [借受世帯ID]=" & 世帯ID & " AND [所有者ID] <> [借受人ID] AND [自小作別]>0)", "", DataViewRowState.CurrentRows)
                ElseIf Val(SysAD.DB(sLRDB).DBProperty("基本台帳管内農地のみ").ToString) = 1 Then
                    pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & 世帯ID & ") AND [自小作別]<1 AND [大字ID]>0", "", DataViewRowState.CurrentRows)
                    pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受世帯ID]=" & 世帯ID & ") AND [自小作別]>0 AND [大字ID]>0", "", DataViewRowState.CurrentRows)
                    pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & 世帯ID & ") AND ([借受世帯ID]<>" & 世帯ID & ") AND ([自小作別]>0) AND ([大字ID]>0)", "", DataViewRowState.CurrentRows)
                Else
                    pTBL所有農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & 世帯ID & ") AND [自小作別]<1", "", DataViewRowState.CurrentRows)
                    pTBL借受農地 = New DataView(App農地基本台帳.TBL農地.Body, "([借受世帯ID]=" & 世帯ID & ") AND [自小作別]>0", "", DataViewRowState.CurrentRows)
                    pTBL貸付農地 = New DataView(App農地基本台帳.TBL農地.Body, "([所有世帯ID]=" & 世帯ID & ") AND ([借受世帯ID]<>" & 世帯ID & ") AND ([自小作別]>0)", "", DataViewRowState.CurrentRows)
                End If

            End If

            For Each pRow As DataRowView In pTBL所有農地
                mvar個人ID.Add(pRow.Item("所有者ID"))
                mvar個人ID.Add(pRow.Item("管理者ID"))
                mvar個人ID.Add(pRow.Item("借受人ID"))
            Next
            For Each pRow As DataRowView In pTBL借受農地
                mvar個人ID.Add(pRow.Item("所有者ID"))
                mvar個人ID.Add(pRow.Item("管理者ID"))
                mvar個人ID.Add(pRow.Item("借受人ID"))
            Next

            HasLand = pTBL所有農地.Count > 0 OrElse pTBL借受農地.Count > 0 OrElse pTBL貸付農地.Count > 0

            If Not Me.世帯ID = 0 Then
                pTBL関係者 = .GetTableBySqlSelect(String.Format("SELECT * FROM [D:個人Info] WHERE [世帯ID]={0} Or [ID] In ({1})", 世帯ID, mvar個人ID.ToString))
            Else
                pTBL関係者 = .GetTableBySqlSelect(String.Format("SELECT * FROM [D:個人Info] WHERE [ID] In ({0})", mvar個人ID.ToString))
            End If
            App農地基本台帳.TBL個人.MergePlus(pTBL関係者)
        End With

    End Sub

    Public Sub MakeXMLFile()
        Maximum = 100
        Value = 20
        Message = "エクセルファイル(世帯員データ)作成中.."
        Application.DoEvents()
        If _Cancel Then
            Throw New Exception("Cancel")
        End If

        sub家族一覧()
        Value = 30
        Message = "エクセルファイル(経営農地)作成中.."
        Application.DoEvents()
        If _Cancel Then
            Throw New Exception("Cancel")
        End If

        sub経営農地一覧()
        Value = 70
        Message = "エクセルファイル(貸付農地)作成中.."
        Application.DoEvents()
        If _Cancel Then
            Throw New Exception("Cancel")
        End If

        sub貸付農地()
        Value = 90
        Message = "エクセルファイル(集計情報)作成中.."
        Application.DoEvents()
        If _Cancel Then
            Throw New Exception("Cancel")
        End If


        sub世帯営農()
    End Sub

    Public Sub SaveAndOpen(ByVal bEditMode As ExcelViewMode)
        Dim sDir As String = SysAD.OutputFolder & "\基本台帳.xml"
        HimTools2012.TextAdapter.SaveTextFile(sDir, Me.XMLSS.OutPut(True))

        Select Case bEditMode
            Case ExcelViewMode.AutoPrint

            Case ExcelViewMode.EditMode
                Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                    pExcel.Show(sDir)
                End Using

            Case ExcelViewMode.Preview
                Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                    pExcel.ShowPreview(sDir)
                End Using
        End Select
    End Sub

    Public Sub sub家族一覧()
        Dim pView家族 As DataView

        pView家族 = New DataView(App農地基本台帳.TBL個人.Body, "[住民区分] IN (0,110) AND [世帯ID]<>0 AND [世帯ID]=" & 世帯ID, "続柄1,続柄2,続柄2", DataViewRowState.CurrentRows)

        Dim pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = mvarXML.WorkBook.WorkSheets.Items("世帯員および就業 ")
        '# 世帯員欄を埋める
        For n As Integer = 1 To 9
            If pView家族.Count >= n Then
                Dim pRow家族 As DataRowView = pView家族(n - 1)
                pSheet.ValueReplace("{家族氏名0" & n & "}", pRow家族.Item("氏名").ToString)

                pSheet.ValueReplace("{続柄0" & n & "}", CObj個人.Get続柄(pRow家族.Row))

                Select Case Val(pRow家族.Item("性別").ToString)
                    Case 0 : pSheet.ValueReplace("{性別0" & n & "}", "男")
                    Case 1 : pSheet.ValueReplace("{性別0" & n & "}", "女")
                    Case Else
                        pSheet.ValueReplace("{性別0" & n & "}", "")
                End Select

                Set和暦TXT(pSheet, pRow家族.Item("生年月日"), "{生年月日0" & n & "}")
                Set正否TXT(pSheet, Not IsDBNull(pRow家族.Item("世帯責任者")) AndAlso pRow家族.Item("世帯責任者") = True, "{世帯責任者0" & n & "}")
                Set正否TXT(pSheet, Not IsDBNull(pRow家族.Item("農業経営者")) AndAlso pRow家族.Item("農業経営者") = True, "{農業経営者0" & n & "}")

                Select Case Val(pRow家族.Item("跡継ぎ区分").ToString)
                    Case 1 : pSheet.ValueReplace(TargetStr2("農業後継ぎ", n, 2), "あ")
                    Case 2 : pSheet.ValueReplace("{農業後継ぎ0" & n & "}", "予")
                    Case 3 : pSheet.ValueReplace("{農業後継ぎ0" & n & "}", "志")
                    Case Else
                        pSheet.ValueReplace("{農業後継ぎ0" & n & "}", "")
                End Select


                Select Case Val(pRow家族.Item("農業改善計画認定").ToString)
                    Case 1
                        pSheet.ValueReplace("{認定農業者0" & n & "}", "認")
                        Set和暦TXT(pSheet, pRow家族.Item("認定日"), "{認定日0" & n & "}")
                        pSheet.ValueReplace("{認定農業者有無}", "○")
                    Case 2
                        pSheet.ValueReplace("{認定農業者0" & n & "}", "担")
                        pSheet.ValueReplace("{認定日0" & n & "}", "")
                        pSheet.ValueReplace("{担い手農家有無}", "○")
                    Case 4
                        pSheet.ValueReplace("{認定農業者0" & n & "}", "認担")
                        Set和暦TXT(pSheet, pRow家族.Item("認定日"), "{認定日0" & n & "}")
                        pSheet.ValueReplace("{認定農業者有無}", "○")
                        pSheet.ValueReplace("{担い手農家有無}", "○")
                    Case Else
                        pSheet.ValueReplace("{認定農業者0" & n & "}", "")
                        pSheet.ValueReplace("{認定日0" & n & "}", "")
                End Select

                '自家農業従事程度
                Select Case Val(pRow家族.Item("自家農業従事程度").ToString)
                    Case 1 : pSheet.ValueReplace("{自家農業従事程度0" & n & "}", "基幹")
                    Case 2 : pSheet.ValueReplace("{自家農業従事程度0" & n & "}", "補助")
                    Case 3 : pSheet.ValueReplace("{自家農業従事程度0" & n & "}", "臨時")
                    Case Else : pSheet.ValueReplace("{自家農業従事程度0" & n & "}", "")
                End Select

                If IsDBNull(pRow家族.Item("農業従事日数")) OrElse pRow家族.Item("農業従事日数") = 0 Then
                    pSheet.ValueReplace("{農業従事日数0" & n & "}", "")
                Else
                    pSheet.ValueReplace("{農業従事日数0" & n & "}", pRow家族.Item("農業従事日数").ToString)
                End If

                Select Case Val(pRow家族.Item("兼業形態").ToString)
                    Case 1 : pSheet.ValueReplace("{兼業形態0" & n & "}", "恒常")
                    Case 2 : pSheet.ValueReplace("{兼業形態0" & n & "}", "出稼")
                    Case 3 : pSheet.ValueReplace("{兼業形態0" & n & "}", "事業形態臨時")
                    Case 4 : pSheet.ValueReplace("{兼業形態0" & n & "}", "自家")
                    Case Else : pSheet.ValueReplace("{兼業形態0" & n & "}", "")
                End Select

                Select Case Val(pRow家族.Item("農年加入受給種別").ToString)
                    Case 1 : pSheet.ValueReplace("{農業者年金0" & n & "}", "旧制度加入者")
                    Case 2 : pSheet.ValueReplace("{農業者年金0" & n & "}", "旧制度受給者")
                    Case 3 : pSheet.ValueReplace("{農業者年金0" & n & "}", "新制度加入者")
                    Case 4 : pSheet.ValueReplace("{農業者年金0" & n & "}", "新制度受給者")
                    Case Else : pSheet.ValueReplace("{農業者年金0" & n & "}", "")
                End Select

                pSheet.ValueReplace("{就労または就学先0" & n & "}", pRow家族.Item("職業").ToString)

                Select Case Val(pRow家族.Item("兼業形態").ToString)
                    Case 1 : pSheet.ValueReplace("{兼業形態0" & n & "}", "恒常")
                    Case 2 : pSheet.ValueReplace("{兼業形態0" & n & "}", "出稼")
                    Case 3 : pSheet.ValueReplace("{兼業形態0" & n & "}", "事業形態臨時")
                    Case 4 : pSheet.ValueReplace("{兼業形態0" & n & "}", "自家")
                    Case Else : pSheet.ValueReplace("{兼業形態0" & n & "}", "")
                End Select

                pSheet.ValueReplace("{1備考0" & n & "}", Replace(pRow家族.Item("備考").ToString, vbLf, "&#10;"))

                If Not mvar人情報Dic.ContainsKey(pRow家族.Item("ID")) Then
                    Dim p人情報 As New 人情報
                    p人情報.ID = pRow家族.Item("ID")
                    p人情報.世帯内 = True
                    p人情報.氏名 = pRow家族.Item("氏名").ToString
                    p人情報.住所 = pRow家族.Item("住所").ToString
                    p人情報.自作地面積 = 0
                    p人情報.借受地面積 = 0
                    p人情報.貸付地面積 = 0

                    mvar人情報Dic.Add(pRow家族.Item("ID"), p人情報)
                End If
            Else
                pSheet.ValueReplace("{家族氏名0" & n & "}", "")
                pSheet.ValueReplace("{続柄0" & n & "}", "")
                pSheet.ValueReplace("{性別0" & n & "}", "")
                pSheet.ValueReplace("{生年月日0" & n & "}", "")
                pSheet.ValueReplace("{世帯責任者0" & n & "}", "")
                pSheet.ValueReplace("{農業経営者0" & n & "}", "")
                pSheet.ValueReplace("{農業後継ぎ0" & n & "}", "")
                pSheet.ValueReplace("{認定農業者0" & n & "}", "")
                pSheet.ValueReplace("{認定日0" & n & "}", "")
                pSheet.ValueReplace("{農業従事日数0" & n & "}", "")
                pSheet.ValueReplace("{自家農業従事程度0" & n & "}", "")
                pSheet.ValueReplace("{兼業形態0" & n & "}", "")
                pSheet.ValueReplace("{加入種別0" & n & "}", "")
                pSheet.ValueReplace("{就労または就学先0" & n & "}", "")
                pSheet.ValueReplace("{農業者年金0" & n & "}", "")
                pSheet.ValueReplace("{1備考0" & n & "}", "")
            End If
        Next
        pSheet.ValueReplace("{認定農業者有無}", "")
        pSheet.ValueReplace("{担い手農家有無}", "")

    End Sub

    '/*********↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑*********/
    Private pSheet01 As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet
    Private pSheet02 As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet
    Private pSheet03 As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet
    Private pSheet04 As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet
    Private pSheet05 As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet

    Public Sub sub経営農地一覧()
        Dim PageCnt As Integer = 10
        If SysAD.市町村.市町村名 = "大崎町" AndAlso mvar印刷Mode = 印刷Mode.簡易印刷 Then
            PageCnt = 7
        End If

        Dim nCount As Decimal = Me.pTBL所有農地.Count + Me.pTBL借受農地.Count
        Dim nPage As Integer = Math.Floor(nCount / PageCnt)
        Me.n置換え桁数 = 3

        If nCount > nPage * PageCnt Then
            nPage = nPage + 1
        End If
        Dim sKK As String = ""
        Message = "エクセルファイル(経営農地_01)作成中.."
        Application.DoEvents()

        If nPage > 1 Then
            Dim pSheet経営(5) As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet
            For i As Integer = 1 To 5
                Select Case mvar印刷Mode
                    Case 印刷Mode.フル印刷
                        pSheet経営(i) = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表" & i)
                    Case 印刷Mode.簡易印刷
                        If i = 1 Then
                            pSheet経営(i) = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表" & i)
                        End If
                End Select
            Next

            For p As Integer = 1 To nPage
                For i As Integer = 1 To 5
                    Select Case mvar印刷Mode
                        Case 印刷Mode.フル印刷
                            Dim pNewSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = mvarXML.WorkBook.WorkSheets.CopySheet(pSheet経営(i).Name, "経営農地等の筆別表" & i & "(" & p & ")")
                        Case 印刷Mode.簡易印刷
                            If i = 1 Then
                                Dim pNewSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = mvarXML.WorkBook.WorkSheets.CopySheet(pSheet経営(i).Name, "経営農地等の筆別表" & i & "(" & p & ")")
                            End If
                    End Select
                Next
            Next

            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表1")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表2")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表3")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表4")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表5")
        ElseIf nCount = 0 Then
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表1")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表2")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表3")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表4")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表5")
            Return
        ElseIf mvar印刷Mode = 印刷Mode.簡易印刷 Then
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表2")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表3")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表4")
            mvarXML.WorkBook.WorkSheets.Remove("経営農地等の筆別表5")
        End If

        Message = "エクセルファイル(経営農地_02)作成中.."
        For nA As Integer = 1 To nPage * PageCnt
            Dim xPage As Integer = Int((nA - 1) / PageCnt) + 1
            If nPage > 1 Then
                sKK = "(" & xPage & ")"
            End If

            Dim n As Integer = ((nA - 1) Mod PageCnt) + 1
            Select Case mvar印刷Mode
                Case 印刷Mode.フル印刷
                    pSheet01 = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表1" & sKK)
                    pSheet02 = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表2" & sKK)
                    pSheet03 = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表3" & sKK)
                    pSheet04 = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表4" & sKK)
                    pSheet05 = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表5" & sKK)

                    pSheet01.ValueReplace(TargetStr2("番号", n), nA)
                    pSheet02.ValueReplace(TargetStr2("番号", n), nA)
                    pSheet03.ValueReplace(TargetStr2("番号", n), nA)
                    pSheet04.ValueReplace(TargetStr2("番号", n), nA)
                    pSheet05.ValueReplace(TargetStr2("番号", n), nA)
                Case 印刷Mode.簡易印刷
                    pSheet01 = mvarXML.WorkBook.WorkSheets.Items("経営農地等の筆別表1" & sKK)
                    pSheet01.ValueReplace(TargetStr2("番号", n), nA)
            End Select

            Message = "エクセル(経営農地_02)作成中.." & n & "/" & (nPage * PageCnt)

            If Me.pTBL所有農地.Count >= nA Then
                Dim p所有Row As DataRowView = Me.pTBL所有農地.Item(nA - 1)
                Dim p所有農地 As CObj農地 = ObjectMan.GetObjectDB("農地", p所有Row.Row, GetType(CObj農地)) ' ObjectMan.GetObject("農地." & p所有Row.Item("ID"))

                print農地共通部分(n, p所有Row, p所有農地)

                If p所有農地.農地状況 >= 20 Then
                    mvar総括表.Add(列.所有_非耕作, 地目.田, Val(p所有Row.Item("田面積").ToString))
                    mvar総括表.Add(列.所有_非耕作, 地目.畑, Val(p所有Row.Item("畑面積").ToString))
                    mvar総括表.Add(列.所有_非耕作, 地目.樹園地, Val(p所有Row.Item("樹園地").ToString))
                    mvar総括表.Add(列.所有_非耕作, 地目.採草放牧地, Val(p所有Row.Item("採草放牧面積").ToString))
                    pSheet01.ValueReplace(TargetStr2("作付状況", n), "×")
                    pSheet01.ValueReplace(TargetStr2("作付状況名", n), p所有Row.Item("農地状況名").ToString)
                Else
                    mvar総括表.Add(列.所有_耕作, 地目.田, Val(p所有Row.Item("田面積").ToString))
                    mvar総括表.Add(列.所有_耕作, 地目.畑, Val(p所有Row.Item("畑面積").ToString))
                    mvar総括表.Add(列.所有_耕作, 地目.樹園地, Val(p所有Row.Item("樹園地").ToString))
                    mvar総括表.Add(列.所有_耕作, 地目.採草放牧地, Val(p所有Row.Item("採草放牧面積").ToString))
                    pSheet01.ValueReplace(TargetStr2("作付状況", n), "")
                    pSheet01.ValueReplace(TargetStr2("作付状況名", n), "")
                End If

                If Not IsDBNull(p所有Row.Item("管理者ID")) AndAlso Not p所有Row.Item("管理者ID") = 0 AndAlso Not p所有Row.Item("所有者ID") = p所有Row.Item("管理者ID") Then
                    Dim p人情報 As 人情報 = Me.Get人情報(p所有Row.Item("管理者ID"), enum農地関連Mode.自作, Val(p所有Row.Item("田面積").ToString), Val(p所有Row.Item("畑面積").ToString), Val(p所有Row.Item("樹園地").ToString), Val(p所有Row.Item("採草放牧面積").ToString))

                    pSheet01.ValueReplace(TargetStr2("所有者名", n), p所有Row.Item("管理者氏名").ToString)
                    pSheet01.ValueReplace(TargetStr2("所有者住所", n), p所有Row.Item("管理者住所").ToString & "&#10;" & "(" & p所有Row.Item("所有者氏名").ToString & ")")
                Else
                    Dim p人情報 As 人情報 = Me.Get人情報(p所有農地.所有者ID, enum農地関連Mode.自作, Val(p所有Row.Item("田面積").ToString), Val(p所有Row.Item("畑面積").ToString), Val(p所有Row.Item("樹園地").ToString), Val(p所有Row.Item("採草放牧面積").ToString))
                    pSheet01.ValueReplace(TargetStr2("所有者名", n), p所有Row.Item("所有者氏名").ToString)
                    pSheet01.ValueReplace(TargetStr2("所有者住所", n), p所有Row.Item("所有者住所").ToString)
                End If

                With pSheet01
                    If Val(p所有Row.Item("共有持分分子").ToString) > 0 AndAlso Val(p所有Row.Item("共有持分分母").ToString) > 0 Then
                        .ValueReplace(TargetStr2("持分割合", n), "(" & p所有Row.Item("共有持分分子") & "/" & p所有Row.Item("共有持分分母") & ")")
                    Else
                        .ValueReplace(TargetStr2("持分割合", n), "")
                    End If

                    Select Case Val(p所有Row.Item("所有者農地意向").ToString)
                        Case 1 : .ValueReplace(TargetStr2("意向内容", n), "所有権移転")
                        Case 2 : .ValueReplace(TargetStr2("意向内容", n), "貸付")
                        Case 3 : .ValueReplace(TargetStr2("意向内容", n), "人・農地プランへの位置づけ")
                        Case 4 : .ValueReplace(TargetStr2("意向内容", n), "農地中間管理機構への貸付")
                        Case 5 : .ValueReplace(TargetStr2("意向内容", n), "その他")
                        Case Else : .ValueReplace(TargetStr2("意向内容", n), "")
                    End Select
                    .ValueReplace(TargetStr2("意向公表", n), IIf(Val(p所有Row.Item("農地法第52公表同意").ToString) = 1, "○", ""))

                    If p所有農地.所有者ID = p所有農地.借受人ID AndAlso Val(p所有Row.Item("自小作別").ToString) > 0 Then
                        Dim p人情報A As 人情報 = Me.Get人情報(p所有農地.借受人ID, enum農地関連Mode.関連なし, 0, 0, 0, 0)
                        If p人情報A Is Nothing Then
                            .ValueReplace(TargetStr2("借受者名", n), "")
                            .ValueReplace(TargetStr2("借受者住所", n), "")
                        Else
                            .ValueReplace(TargetStr2("借受者名", n), p人情報A.氏名.ToString)
                            .ValueReplace(TargetStr2("借受者住所", n), p人情報A.住所.ToString)
                        End If
                        Set小作情報(pSheet01, p所有農地, p所有Row, n, 8)
                    End If

                    .ValueReplace(TargetStr2("耕地番号", n), "")
                    .ValueReplace(TargetStr2("自作借入別", n), "自")
                    .ValueReplace(TargetStr2("借受者名", n), "")
                    .ValueReplace(TargetStr2("借受者住所", n), "")
                    .ValueReplace(TargetStr2("借受者", n), "")
                    .ValueReplace(TargetStr2("経由法人名", n), "")
                    .ValueReplace(TargetStr2("経由法人名B", n), "")
                    .ValueReplace(TargetStr2("適用法", n), "")
                    .ValueReplace(TargetStr2("形態", n), "")
                    .ValueReplace(TargetStr2("契約開始日", n), "")
                    .ValueReplace(TargetStr2("契約終了日", n), "")
                    .ValueReplace(TargetStr2("機構契約開始日", n), "")
                    .ValueReplace(TargetStr2("機構契約終了日", n), "")
                    .ValueReplace(TargetStr2("賃借料", n), "")
                    .ValueReplace(TargetStr2("10ａ賃借料", n), "")
                End With
            ElseIf Me.pTBL所有農地.Count < nA AndAlso Me.pTBL借受農地.Count >= (nA - Me.pTBL所有農地.Count) Then
                Dim p借受Row As DataRowView = Me.pTBL借受農地.Item(nA - 1 - Me.pTBL所有農地.Count)
                Dim p借受農地 As New CObj農地(p借受Row.Row, False)

                print農地共通部分(n, p借受Row, p借受農地)

                If p借受農地.農地状況 >= 20 Then
                    With mvar総括表
                        mvar総括表.Add(列.借受_非耕作, 地目.田, Val(p借受Row.Item("田面積").ToString))
                        mvar総括表.Add(列.借受_非耕作, 地目.畑, Val(p借受Row.Item("畑面積").ToString))
                        mvar総括表.Add(列.借受_非耕作, 地目.樹園地, Val(p借受Row.Item("樹園地").ToString))
                        mvar総括表.Add(列.借受_非耕作, 地目.採草放牧地, Val(p借受Row.Item("採草放牧面積").ToString))
                    End With
                    pSheet01.ValueReplace(TargetStr2("作付状況", n), "×")
                    pSheet01.ValueReplace(TargetStr2("作付状況名", n), p借受Row.Item("農地状況名").ToString)
                Else
                    With mvar総括表
                        mvar総括表.Add(列.借受_耕作, 地目.田, Val(p借受Row.Item("田面積").ToString))
                        mvar総括表.Add(列.借受_耕作, 地目.畑, Val(p借受Row.Item("畑面積").ToString))
                        mvar総括表.Add(列.借受_耕作, 地目.樹園地, Val(p借受Row.Item("樹園地").ToString))
                        mvar総括表.Add(列.借受_耕作, 地目.採草放牧地, Val(p借受Row.Item("採草放牧面積").ToString))
                    End With
                    pSheet01.ValueReplace(TargetStr2("作付状況", n), "")
                    pSheet01.ValueReplace(TargetStr2("作付状況名", n), "")
                End If

                Dim p人情報 As 人情報 = Me.Get人情報(p借受農地.所有者ID, enum農地関連Mode.関連なし, 0, 0, 0, 0)
                If p人情報 Is Nothing Then
                    pSheet01.ValueReplace(TargetStr2("所有者名", n), "")
                    pSheet01.ValueReplace(TargetStr2("所有者住所", n), "")
                Else
                    pSheet01.ValueReplace(TargetStr2("所有者名", n), p人情報.氏名.ToString)
                    pSheet01.ValueReplace(TargetStr2("所有者住所", n), p人情報.住所.ToString)
                End If

                If Val(p借受Row.Item("共有持分分子").ToString) > 0 AndAlso Val(p借受Row.Item("共有持分分母").ToString) > 0 Then
                    pSheet01.ValueReplace(TargetStr2("持分割合", n), "(" & p借受Row.Item("共有持分分子") & "/" & p借受Row.Item("共有持分分母") & ")")
                Else
                    pSheet01.ValueReplace(TargetStr2("持分割合", n), "")
                End If

                Select Case Val(p借受Row.Item("所有者農地意向").ToString)
                    Case 1 : pSheet01.ValueReplace(TargetStr2("意向内容", n), "所")
                    Case 2 : pSheet01.ValueReplace(TargetStr2("意向内容", n), "貸")
                    Case 3 : pSheet01.ValueReplace(TargetStr2("意向内容", n), "人")
                    Case 4 : pSheet01.ValueReplace(TargetStr2("意向内容", n), "機")
                    Case 5 : pSheet01.ValueReplace(TargetStr2("意向内容", n), "他")
                    Case 6 : pSheet01.ValueReplace(TargetStr2("意向内容", n), "自")
                    Case Else : pSheet01.ValueReplace(TargetStr2("意向内容", n), "")
                End Select
                pSheet01.ValueReplace(TargetStr2("意向公表", n), IIf(Val(p借受Row.Item("農地法第52公表同意").ToString) = 1, "○", ""))

                Dim p人情報B As 人情報 = Me.Get人情報(p借受農地.借受人ID, enum農地関連Mode.借受, p借受農地.田面積, p借受農地.畑面積, p借受農地.樹園地, p借受農地.採草放牧地)
                If p人情報B Is Nothing Then
                    pSheet01.ValueReplace(TargetStr2("借受者名", n), "")
                    pSheet01.ValueReplace(TargetStr2("借受者住所", n), "")
                Else
                    pSheet01.ValueReplace(TargetStr2("借受者名", n), p人情報B.氏名.ToString)
                    pSheet01.ValueReplace(TargetStr2("借受者住所", n), p人情報B.住所.ToString)
                End If

                Set小作情報(pSheet01, p借受農地, p借受Row, n, 8)

                pSheet01.ValueReplace(TargetStr2("自作借入別", n), "小")
            Else
                '経営農地の筆別表（1）
                With pSheet01
                    .ValueReplace(TargetStr2("番号", n), "")
                    .ValueReplace(TargetStr2("大字小字", n), "")
                    .ValueReplace(TargetStr2("地番", n), "")
                    .ValueReplace(TargetStr2("耕地番号", n), "")
                    .ValueReplace(TargetStr2("登記地目", n), "")
                    .ValueReplace(TargetStr2("現況地目", n), "")
                    .ValueReplace(TargetStr2("登記面積", n), "")
                    .ValueReplace(TargetStr2("実面積", n), "")
                    .ValueReplace(TargetStr2("本地面積", n), "")
                    .ValueReplace(TargetStr2("農振法", n), "")
                    .ValueReplace(TargetStr2("都市計画法", n), "")
                    .ValueReplace(TargetStr2("土地改良", n), "")
                    .ValueReplace(TargetStr2("生産緑地法に基づく指定", n), "")
                    .ValueReplace(TargetStr2("自作借入別", n), "")
                    .ValueReplace(TargetStr2("所有者名", n), "")
                    .ValueReplace(TargetStr2("所有者住所", n), "")
                    .ValueReplace(TargetStr2("持分割合", n), "")
                    .ValueReplace(TargetStr2("意向内容", n), "")
                    .ValueReplace(TargetStr2("意向公表", n), "")
                    .ValueReplace(TargetStr2("借受者名", n), "")
                    .ValueReplace(TargetStr2("借受者住所", n), "")
                    '.ValueReplace(TargetStr2("整理番号", n), "")
                    .ValueReplace(TargetStr2("経由法人名", n), "")
                    .ValueReplace(TargetStr2("経由法人名B", n), "")
                    .ValueReplace(TargetStr2("適用法", n), "")
                    .ValueReplace(TargetStr2("形態", n), "")
                    .ValueReplace(TargetStr2("契約開始日", n), "")
                    .ValueReplace(TargetStr2("契約終了日", n), "")
                    .ValueReplace(TargetStr2("機構契約開始日", n), "")
                    .ValueReplace(TargetStr2("機構契約終了日", n), "")
                    .ValueReplace(TargetStr2("賃借料", n), "")
                    .ValueReplace(TargetStr2("10ａ賃借料", n), "")
                    .ValueReplace(TargetStr2("作付状況", n), "")
                    .ValueReplace(TargetStr2("作付状況名", n), "")
                    .ValueReplace(TargetStr2("6備考", n), "")
                End With

                Select Case mvar印刷Mode
                    Case 印刷Mode.フル印刷
                        '経営農地の筆別表（2）
                        With pSheet02
                            .ValueReplace(TargetStr2("番号", n), "")
                            .ValueReplace(TargetStr2("利用状況報告の対象", n), "")
                            .ValueReplace(TargetStr2("利用状況報告日", n), "")
                            .ValueReplace(TargetStr2("利用状況勧告日", n), "")
                            .ValueReplace(TargetStr2("是正措置内容", n), "")
                            .ValueReplace(TargetStr2("是正措置期限日", n), "")
                            .ValueReplace(TargetStr2("勧告根拠農地法", n), "")
                            .ValueReplace(TargetStr2("勧告根拠基盤法", n), "")
                            .ValueReplace(TargetStr2("是正確認日", n), "")
                            .ValueReplace(TargetStr2("是正状況", n), "")
                            .ValueReplace(TargetStr2("許可取消日", n), "")
                            .ValueReplace(TargetStr2("取消事由", n), "")
                            .ValueReplace(TargetStr2("許可根拠農地法", n), "")
                            .ValueReplace(TargetStr2("許可根拠基盤法", n), "")
                            .ValueReplace(TargetStr2("相続届出日", n), "")
                            .ValueReplace(TargetStr2("届出事由", n), "")
                            .ValueReplace(TargetStr2("権利取得者名", n), "")
                            .ValueReplace(TargetStr2("相続あっせん希望", n), "")
                        End With

                        '経営農地の筆別表（3）
                        With pSheet03
                            .ValueReplace(TargetStr2("番号", n), "")
                            .ValueReplace(TargetStr2("利用状況調査年月日", n), "")
                            .ValueReplace(TargetStr2("農法32条1項", n), "")
                            .ValueReplace(TargetStr2("荒廃農地分類", n), "")
                            .ValueReplace(TargetStr2("転用分類", n), "")
                            .ValueReplace(TargetStr2("利用意向調査年月日", n), "")
                            .ValueReplace(TargetStr2("利用意向根拠条項", n), "")
                            .ValueReplace(TargetStr2("利用意向意思表明日", n), "")
                            .ValueReplace(TargetStr2("利用意向内容区分", n), "")
                            .ValueReplace(TargetStr2("利用意向調査区分", n), "")
                            .ValueReplace(TargetStr2("利用意向調査結果", n), "")
                            .ValueReplace(TargetStr2("農法32公示日", n), "")
                            .ValueReplace(TargetStr2("農法43通知日", n), "")
                            .ValueReplace(TargetStr2("農法35の1通知日", n), "")
                            .ValueReplace(TargetStr2("農法35の2通知日", n), "")
                            .ValueReplace(TargetStr2("農法35の3通知日", n), "")
                            .ValueReplace(TargetStr2("農法36の1通知日", n), "")
                            .ValueReplace(TargetStr2("中管勧告内容", n), "")
                            .ValueReplace(TargetStr2("中管通知日", n), "")
                            .ValueReplace(TargetStr2("再生利用困難農地", n), "")
                            .ValueReplace(TargetStr2("農法40公告日", n), "")
                            .ValueReplace(TargetStr2("農法43の3公告日", n), "")
                            .ValueReplace(TargetStr2("農法44の1命令日", n), "")
                            .ValueReplace(TargetStr2("農法44の3公告日", n), "")
                        End With

                        '経営農地の筆別表（4）
                        With pSheet04
                            .ValueReplace(TargetStr2("番号", n), "")
                            .ValueReplace(TargetStr2("中管権取得年月日", n), "")
                            .ValueReplace(TargetStr2("意見回答年月日", n), "")
                            .ValueReplace(TargetStr2("知事公告年月日", n), "")
                            .ValueReplace(TargetStr2("認可通知年月日", n), "")
                            .ValueReplace(TargetStr2("権利設定内容", n), "")
                            .ValueReplace(TargetStr2("利用配分設定期間", n), "")
                            .ValueReplace(TargetStr2("利用配分始期日", n), "")
                            .ValueReplace(TargetStr2("利用配分終期日", n), "")
                            .ValueReplace(TargetStr2("借賃額", n), "")
                            .ValueReplace(TargetStr2("10a借賃額", n), "")
                            .ValueReplace(TargetStr2("貸借解除日", n), "")
                            .ValueReplace(TargetStr2("相続税猶予", n), "")
                            .ValueReplace(TargetStr2("贈与税猶予", n), "")
                            .ValueReplace(TargetStr2("納税猶予種別", n), "")
                            .ValueReplace(TargetStr2("納税相続日", n), "")
                            .ValueReplace(TargetStr2("納税適用日", n), "")
                            .ValueReplace(TargetStr2("納税継続日", n), "")
                            .ValueReplace(TargetStr2("租税特別措置法", n), "")
                            .ValueReplace(TargetStr2("営農困難時貸付け", n), "")
                            .ValueReplace(TargetStr2("設定年月日", n), "")
                            .ValueReplace(TargetStr2("仮登記権者氏名", n), "")
                            .ValueReplace(TargetStr2("仮登記権者住所", n), "")
                            .ValueReplace(TargetStr2("農業直接支払交付金", n), "")
                            .ValueReplace(TargetStr2("農地維持支払交付金", n), "")
                            .ValueReplace(TargetStr2("資源向上支払交付金", n), "")
                            .ValueReplace(TargetStr2("中山間地域等直接支払", n), "")
                            .ValueReplace(TargetStr2("特定処分対象農地", n), "")
                        End With

                        '経営農地の筆別表（5）
                        With pSheet05
                            .ValueReplace(TargetStr2("番号", n), "")
                            .ValueReplace(TargetStr2("生産緑地法種別", n), "")
                            .ValueReplace(TargetStr2("生産緑地法指定日", n), "")
                            .ValueReplace(TargetStr2("共有農地区分", n), "")
                            .ValueReplace(TargetStr2("共有者", n), "")
                            .ValueReplace(TargetStr2("共有農地持分割合", n), "")
                            .ValueReplace(TargetStr2("特定作業者", n), "")
                            .ValueReplace(TargetStr2("特定作業作目", n), "")
                            .ValueReplace(TargetStr2("特定作業内容", n), "")
                            .ValueReplace(TargetStr2("6備考", n), "")
                        End With

                    Case 印刷Mode.簡易印刷
                        pSheet01.ValueReplace(TargetStr2("6備考", n), "")
                End Select
            End If
        Next
    End Sub


    Private Sub print農地共通部分(ByRef n As Integer, ByRef p農地Row As DataRowView, ByVal p農地 As CObj農地)
        Me.n置換え桁数 = 3

        With pSheet01
            If p農地.所在.Length > 0 Then
                .ValueReplace(TargetStr2("大字小字", n), p農地.所在)
            Else
                .ValueReplace(TargetStr2("大字小字", n), p農地.大字 & IIf(p農地.小字 = "-", "", "字" & p農地.小字))
            End If
            .ValueReplace(TargetStr2("地番", n), p農地.地番)

            .ValueReplace(TargetStr2("所在", n), p農地.土地所在)
            .ValueReplace(TargetStr2("耕地番号", n), IIf(p農地.耕地番号 = 0, "", p農地.耕地番号))

            If p農地.登記簿地目 <> 0 Then
                Dim p登記地目 As DataRow = App農地基本台帳.TBL地目.Rows.Find(p農地.登記簿地目)

                If p登記地目 Is Nothing Then : .ValueReplace(TargetStr2("登記地目", n), "")
                Else : .ValueReplace(TargetStr2("登記地目", n), p登記地目.Item("名称").ToString)
                End If
            Else
                .ValueReplace(TargetStr2("登記地目", n), "")
            End If

            If p農地.現況地目 <> 0 Then
                Dim p現況地目 As DataRow = App農地基本台帳.TBL現況地目.Rows.Find(p農地.現況地目)

                If p現況地目 Is Nothing Then : .ValueReplace(TargetStr2("現況地目", n), "")
                Else : .ValueReplace(TargetStr2("現況地目", n), p現況地目.Item("名称").ToString)
                End If
            Else
                .ValueReplace(TargetStr2("現況地目", n), "")
            End If

            If p農地.登記簿面積 < 10 Then : .ValueReplace(TargetStr2("登記面積", n), NumToString(p農地.登記簿面積.ToString("F2")))
            Else : .ValueReplace(TargetStr2("登記面積", n), NumToString(p農地.登記簿面積))
            End If

            If p農地.実面積 < 10 Then : .ValueReplace(TargetStr2("実面積", n), NumToString(p農地.実面積.ToString("F2")))
            Else : .ValueReplace(TargetStr2("実面積", n), NumToString(p農地.実面積))
            End If

            If p農地.本地面積 > 0 Then : .ValueReplace(TargetStr2("本地面積", n), NumToString(p農地.本地面積))
            Else : .ValueReplace(TargetStr2("本地面積", n), "")
            End If

            Select Case p農地.農振法区分
                Case enum農振法区分.農用地区域 : .ValueReplace(TargetStr2("農振法", n), "内")
                Case enum農振法区分.農振地域 : .ValueReplace(TargetStr2("農振法", n), "他")
                Case enum農振法区分.農振地域外 : .ValueReplace(TargetStr2("農振法", n), "外")
                Case Else
                    Select Case p農地.旧農振区分
                        Case enum農業振興地域.農用地外 : .ValueReplace(TargetStr2("農振法", n), "他")
                        Case enum農業振興地域.農用地内 : .ValueReplace(TargetStr2("農振法", n), "内")
                        Case enum農業振興地域.振興地域外 : .ValueReplace(TargetStr2("農振法", n), "外")
                        Case Else : .ValueReplace(TargetStr2("農振法", n), "")
                    End Select
            End Select

            Select Case p農地.都市計画法区分
                Case enum都市計画法区分.市街化区域 : .ValueReplace(TargetStr2("都市計画法", n), "市")
                Case enum都市計画法区分.市街化調整区域 : .ValueReplace(TargetStr2("都市計画法", n), "調")
                Case enum都市計画法区分.非線引き都市計画区域の用途地域 : .ValueReplace(TargetStr2("都市計画法", n), "用")
                Case enum都市計画法区分.非線引き都市計画区域内 : .ValueReplace(TargetStr2("都市計画法", n), "内")
                Case enum都市計画法区分.都市計画区域外 : .ValueReplace(TargetStr2("都市計画法", n), "外")
                Case enum都市計画法区分.その他 : .ValueReplace(TargetStr2("都市計画法", n), "他")
                Case Else
                    Select Case p農地.都市計画法
                        Case enum都市計画法.都市計画法外 : .ValueReplace(TargetStr2("都市計画法", n), "外")
                        Case enum都市計画法.都市計画法内 : .ValueReplace(TargetStr2("都市計画法", n), "内")
                        Case enum都市計画法.用途地域内 : .ValueReplace(TargetStr2("都市計画法", n), "用")
                        Case enum都市計画法.調整区域内 : .ValueReplace(TargetStr2("都市計画法", n), "調")
                        Case enum都市計画法.市街化区域内 : .ValueReplace(TargetStr2("都市計画法", n), "市")
                        Case enum都市計画法.都市計画白地 : .ValueReplace(TargetStr2("都市計画法", n), "白地")
                        Case Else : .ValueReplace(TargetStr2("都市計画法", n), "")
                    End Select
            End Select

            Select Case p農地.土地改良法
                Case enum土地改良法.区域外 : .ValueReplace(TargetStr2("土地改良", n), "外")
                Case enum土地改良法.区域内_整備済 : .ValueReplace(TargetStr2("土地改良", n), "内済")
                Case enum土地改良法.区域内_整備中 : .ValueReplace(TargetStr2("土地改良", n), "内中")
                Case Else : .ValueReplace(TargetStr2("土地改良", n), "")
            End Select

            Select Case p農地.生産緑地法
                Case enum有無.有 : .ValueReplace(TargetStr2("生産緑地法に基づく指定", n), "有")
                Case enum有無.無 : .ValueReplace(TargetStr2("生産緑地法に基づく指定", n), "無")
                Case Else : .ValueReplace(TargetStr2("生産緑地法に基づく指定", n), "")
            End Select
        End With

        Select Case mvar印刷Mode
            Case 印刷Mode.簡易印刷
                pSheet01.ValueReplace(TargetStr2("6備考", n), p農地.備考)
            Case 印刷Mode.フル印刷
                '経営農地2
                With pSheet02
                    Select Case p農地.利用状況報告対象
                        Case enum有無.有 : .ValueReplace(TargetStr2("利用状況報告の対象", n), "有")
                        Case enum有無.無 : .ValueReplace(TargetStr2("利用状況報告の対象", n), "無")
                    End Select

                    SubReplace年月日2(p農地, pSheet02, n, "利用状況報告年月日", "利用状況報告日")
                    SubReplace年月日2(p農地, pSheet02, n, "是正勧告日", "利用状況勧告日")
                    SubReplace年月日2(p農地, pSheet02, n, "是正期限", "是正措置期限日")
                    SubReplace年月日2(p農地, pSheet02, n, "是正確認", "是正確認日")
                    SubReplace年月日2(p農地, pSheet02, n, "取消年月日", "許可取消日")
                    SubReplace年月日2(p農地, pSheet02, n, "届出年月日", "相続届出日")

                    .ValueReplace(TargetStr2("是正措置内容", n), p農地.是正内容)

                    Select Case p農地.根拠条件農地法
                        Case enum様式1.不明 : .ValueReplace(TargetStr2("勧告根拠農地法", n), "_")
                        Case enum様式1.農地法1号 : .ValueReplace(TargetStr2("勧告根拠農地法", n), "1号")
                        Case enum様式1.農地法2号 : .ValueReplace(TargetStr2("勧告根拠農地法", n), "2号")
                        Case enum様式1.農地法3号 : .ValueReplace(TargetStr2("勧告根拠農地法", n), "3号")
                        Case Else : .ValueReplace(TargetStr2("勧告根拠農地法", n), "")
                    End Select


                    Select Case p農地.根拠条件基盤強化法
                        Case enum様式1.不明 : .ValueReplace(TargetStr2("勧告根拠基盤法", n), "_")
                        Case enum様式1.農地法1号 : .ValueReplace(TargetStr2("勧告根拠基盤法", n), "1号")
                        Case enum様式1.農地法2号 : .ValueReplace(TargetStr2("勧告根拠基盤法", n), "2号")
                        Case enum様式1.農地法3号 : .ValueReplace(TargetStr2("勧告根拠基盤法", n), "3号")
                        Case Else : .ValueReplace(TargetStr2("勧告根拠基盤法", n), "")
                    End Select

                    .ValueReplace(TargetStr2("是正状況", n), p農地.是正状況)
                    .ValueReplace(TargetStr2("取消事由", n), p農地.取消事由)

                    Select Case p農地.取消条件農地法
                        Case enum様式2.不明 : .ValueReplace(TargetStr2("許可根拠農地法", n), "_")
                        Case enum様式2.農地法1号 : .ValueReplace(TargetStr2("許可根拠農地法", n), "1号")
                        Case enum様式2.農地法2号 : .ValueReplace(TargetStr2("許可根拠農地法", n), "2号")
                        Case Else : .ValueReplace(TargetStr2("許可根拠農地法", n), "")
                    End Select


                    Select Case p農地.取消条件基盤強化法
                        Case enum様式2.不明 : .ValueReplace(TargetStr2("許可根拠基盤法", n), "_")
                        Case enum様式2.農地法1号 : .ValueReplace(TargetStr2("許可根拠基盤法", n), "1号")
                        Case enum様式2.農地法2号 : .ValueReplace(TargetStr2("許可根拠基盤法", n), "2号")
                        Case Else : .ValueReplace(TargetStr2("許可根拠基盤法", n), "")
                    End Select

                    .ValueReplace(TargetStr2("届出事由", n), p農地.GetStringValue("届出事由"))

                    If Not p農地.GetStringValue("届出者氏名") = "" Then
                        .ValueReplace(TargetStr2("権利取得者名", n), p農地.GetStringValue("届出者氏名"))
                    Else
                        Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(p農地.GetLongIntValue("相続届出者ID"))
                        If pRow IsNot Nothing Then
                            .ValueReplace(TargetStr2("権利取得者名", n), pRow.Item("氏名"))
                        Else
                            .ValueReplace(TargetStr2("権利取得者名", n), "")
                        End If
                    End If

                    Select Case p農地.GetIntegerValue("あっせん希望")
                        Case enum有無.有 : .ValueReplace(TargetStr2("相続あっせん希望", n), "有")
                        Case enum有無.無 : .ValueReplace(TargetStr2("相続あっせん希望", n), "無")
                    End Select
                End With

                '経営農地3

                With pSheet03
                    SubReplace年月日2(p農地, pSheet03, n, "利用状況調査日", "利用状況調査年月日")

                    Select Case p農地.GetIntegerValue("利用状況調査農地法")
                        Case enum様式2.不明 : .ValueReplace(TargetStr2("農法32条1項", n), "-")
                        Case enum様式2.農地法1号 : .ValueReplace(TargetStr2("農法32条1項", n), "1号")
                        Case enum様式2.農地法2号 : .ValueReplace(TargetStr2("農法32条1項", n), "2号")
                        Case enum様式2.遊休農地でない : .ValueReplace(TargetStr2("農法32条1項", n), "遊休農地でない")
                        Case Else : .ValueReplace(TargetStr2("農法32条1項", n), "")
                    End Select

                    Select Case p農地.GetIntegerValue("利用状況調査荒廃")
                        Case enum利用状況調査荒廃.不明 : .ValueReplace(TargetStr2("荒廃農地分類", n), "-")
                        Case enum利用状況調査荒廃.A分類 : .ValueReplace(TargetStr2("荒廃農地分類", n), "A分類")
                        Case enum利用状況調査荒廃.B分類 : .ValueReplace(TargetStr2("荒廃農地分類", n), "B分類")
                        Case Else : .ValueReplace(TargetStr2("荒廃農地分類", n), "")
                    End Select

                    Select Case p農地.GetIntegerValue("利用状況調査転用")
                        Case enum利用状況調査転用.不明 : .ValueReplace(TargetStr2("転用分類", n), "-")
                        Case enum利用状況調査転用.一時転用 : .ValueReplace(TargetStr2("転用分類", n), "一時転用")
                        Case enum利用状況調査転用.無断転用 : .ValueReplace(TargetStr2("転用分類", n), "無断転用")
                        Case enum利用状況調査転用.違反転用 : .ValueReplace(TargetStr2("転用分類", n), "違反転用")
                        Case Else : .ValueReplace(TargetStr2("転用分類", n), "")
                    End Select

                    SubReplace年月日2(p農地, pSheet03, n, "利用意向調査日", "利用意向調査年月日")

                    Select Case p農地.GetIntegerValue("利用意向根拠条項")
                        Case enum利用意向根拠条項.不明 : .ValueReplace(TargetStr2("利用意向根拠条項", n), "-")
                        Case enum利用意向根拠条項.農地法第32条第1項 : .ValueReplace(TargetStr2("利用意向根拠条項", n), "農地法第32条第1項")
                        Case enum利用意向根拠条項.農地法第32条第4項 : .ValueReplace(TargetStr2("利用意向根拠条項", n), "農地法第32条第4項")
                        Case enum利用意向根拠条項.農地法第33条第1項 : .ValueReplace(TargetStr2("利用意向根拠条項", n), "農地法第33条第1項")
                        Case Else : .ValueReplace(TargetStr2("利用意向根拠条項", n), "")
                    End Select

                    SubReplace年月日2(p農地, pSheet03, n, "利用意向意思表明日", "利用意向意思表明日")

                    Select Case p農地.GetIntegerValue("利用意向意向内容区分")
                        Case enum利用意向内容区分.不明 : .ValueReplace(TargetStr2("利用意向内容区分", n), "-")
                        Case enum利用意向内容区分.自ら耕作 : .ValueReplace(TargetStr2("利用意向内容区分", n), "自ら耕作")
                        Case enum利用意向内容区分.機構事業 : .ValueReplace(TargetStr2("利用意向内容区分", n), "機構事業")
                        Case enum利用意向内容区分.所有者代理事業 : .ValueReplace(TargetStr2("利用意向内容区分", n), "所有者代理事業")
                        Case enum利用意向内容区分.権利設定または移転 : .ValueReplace(TargetStr2("利用意向内容区分", n), "権利設定または移転")
                        Case enum利用意向内容区分.その他 : .ValueReplace(TargetStr2("利用意向内容区分", n), "その他")
                        Case Else : .ValueReplace(TargetStr2("利用意向内容区分", n), "")
                    End Select

                    Select Case p農地.GetIntegerValue("利用意向権利関係調査区分")
                        Case enum利用意向権利関係調査区分.不明 : .ValueReplace(TargetStr2("利用意向調査区分", n), "-")
                        Case enum利用意向権利関係調査区分.対象外 : .ValueReplace(TargetStr2("利用意向調査区分", n), "対象外")
                        Case enum利用意向権利関係調査区分.調査中 : .ValueReplace(TargetStr2("利用意向調査区分", n), "調査中")
                        Case enum利用意向権利関係調査区分.調査済み : .ValueReplace(TargetStr2("利用意向調査区分", n), "調査済み")
                        Case Else : .ValueReplace(TargetStr2("利用意向調査区分", n), "")
                    End Select

                    .ValueReplace(TargetStr2("利用意向調査結果", n), p農地.利用意向権利関係調査記録)
                    SubReplace年月日2(p農地, pSheet03, n, "利用意向公示年月日", "農法32公示日")
                    SubReplace年月日2(p農地, pSheet03, n, "利用意向通知年月日", "農法43通知日")

                    SubReplace年月日2(p農地, pSheet03, n, "農地法35の1通知日", "農法35の1通知日")
                    SubReplace年月日2(p農地, pSheet03, n, "農地法35の2通知日", "農法35の2通知日")
                    SubReplace年月日2(p農地, pSheet03, n, "農地法35の3通知日", "農法35の3通知日")

                    SubReplace年月日2(p農地, pSheet03, n, "勧告年月日", "農法36の1通知日")

                    Select Case p農地.GetIntegerValue("勧告内容")
                        Case 0 : .ValueReplace(TargetStr2("中管勧告内容", n), "-")
                        Case 1 : .ValueReplace(TargetStr2("中管勧告内容", n), "1号")
                        Case 2 : .ValueReplace(TargetStr2("中管勧告内容", n), "2号")
                        Case 3 : .ValueReplace(TargetStr2("中管勧告内容", n), "3号")
                        Case 4 : .ValueReplace(TargetStr2("中管勧告内容", n), "4号")
                        Case 5 : .ValueReplace(TargetStr2("中管勧告内容", n), "5号")
                        Case Else : .ValueReplace(TargetStr2("中管勧告内容", n), "")
                    End Select

                    SubReplace年月日2(p農地, pSheet03, n, "中間管理勧告日", "中管通知日")

                    Select Case p農地.GetIntegerValue("再生利用困難農地")
                        Case 0 : .ValueReplace(TargetStr2("再生利用困難農地", n), "-")
                        Case 1 : .ValueReplace(TargetStr2("再生利用困難農地", n), "機構法第20条")
                        Case 2 : .ValueReplace(TargetStr2("再生利用困難農地", n), "農地法第35条")
                        Case 3 : .ValueReplace(TargetStr2("再生利用困難農地", n), "農地法第37条")
                        Case 4 : .ValueReplace(TargetStr2("再生利用困難農地", n), "災害")
                        Case 5 : .ValueReplace(TargetStr2("再生利用困難農地", n), "農地法第34条")
                        Case 6 : .ValueReplace(TargetStr2("再生利用困難農地", n), "その他")
                        Case Else : .ValueReplace(TargetStr2("再生利用困難農地", n), "")
                    End Select

                    SubReplace年月日2(p農地, pSheet03, n, "農地法40裁定公告日", "農法40公告日")
                    SubReplace年月日2(p農地, pSheet03, n, "農地法43裁定公告日", "農法43の3公告日")
                    SubReplace年月日2(p農地, pSheet03, n, "農地法44の1裁定公告日", "農法44の1命令日")
                    SubReplace年月日2(p農地, pSheet03, n, "農地法44の3裁定公告日", "農法44の3公告日")
                End With

                '経営農地4
                With pSheet04
                    SubReplace年月日2(p農地, pSheet04, n, "中間管理権取得日", "中管権取得年月日")

                    SubReplace年月日2(p農地, pSheet04, n, "意見回答日", "意見回答年月日")
                    SubReplace年月日2(p農地, pSheet04, n, "知事公告日", "知事公告年月日")
                    SubReplace年月日2(p農地, pSheet04, n, "認可通知日", "認可通知年月日")

                    Select Case p農地.GetIntegerValue("権利設定内容")
                        Case 0 : .ValueReplace(TargetStr2("権利設定内容", n), "-")
                        Case 1 : .ValueReplace(TargetStr2("権利設定内容", n), "使")
                        Case 2 : .ValueReplace(TargetStr2("権利設定内容", n), "賃")
                    End Select

                    SubReplace年月日2(p農地, pSheet04, n, "利用配分設定期間", "利用配分設定期間")
                    SubReplace年月日2(p農地, pSheet04, n, "利用配分計画始期日", "利用配分始期日")
                    SubReplace年月日2(p農地, pSheet04, n, "利用配分計画終期日", "利用配分終期日")

                    .ValueReplace(TargetStr2("借賃額", n), IIf(p農地.利用配分計画借賃額 = 0, "", p農地.利用配分計画借賃額))
                    .ValueReplace(TargetStr2("10a借賃額", n), IIf(p農地.利用配分計画10a賃借料 = 0, "", p農地.利用配分計画10a賃借料))

                    SubReplace年月日2(p農地, pSheet04, n, "貸借契約解除年月日", "貸借解除日")

                    Select Case p農地.GetIntegerValue("納税猶予対象農地")
                        Case enum納税猶予.不明 : .ValueReplace(TargetStr2("相続税猶予", n), "無")
                        Case enum納税猶予.贈与税 : .ValueReplace(TargetStr2("相続税猶予", n), "有")
                        Case enum納税猶予.相続税 : .ValueReplace(TargetStr2("相続税猶予", n), "無")
                    End Select

                    Select Case p農地.GetIntegerValue("納税猶予対象農地")
                        Case enum納税猶予.不明 : .ValueReplace(TargetStr2("贈与税猶予", n), "無")
                        Case enum納税猶予.贈与税 : .ValueReplace(TargetStr2("贈与税猶予", n), "無")
                        Case enum納税猶予.相続税 : .ValueReplace(TargetStr2("贈与税猶予", n), "有")
                    End Select

                    Select Case p農地.GetIntegerValue("納税猶予種別")
                        Case 0 : .ValueReplace(TargetStr2("納税猶予種別", n), "-")
                        Case 1 : .ValueReplace(TargetStr2("納税猶予種別", n), "特例適用外")
                        Case 2 : .ValueReplace(TargetStr2("納税猶予種別", n), "対象外")
                    End Select

                    SubReplace年月日2(p農地, pSheet04, n, "納税猶予相続日", "納税相続日")
                    SubReplace年月日2(p農地, pSheet04, n, "納税猶予適用日", "納税適用日")
                    SubReplace年月日2(p農地, pSheet04, n, "納税猶予継続日", "納税継続日")

                    Select Case p農地.GetIntegerValue("租税処置法")
                        Case enum租税処置法.不明 : .ValueReplace(TargetStr2("租税特別措置法", n), "_")
                        Case enum租税処置法.租税処置法1号 : .ValueReplace(TargetStr2("租税特別措置法", n), "1号")
                        Case enum租税処置法.租税処置法2号 : .ValueReplace(TargetStr2("租税特別措置法", n), "2号")
                        Case enum租税処置法.租税処置法3号 : .ValueReplace(TargetStr2("租税特別措置法", n), "2号")
                        Case Else : pSheet04.ValueReplace(TargetStr2("租税特別措置法", n), "")
                    End Select

                    .ValueReplace(TargetStr2("営農困難時貸付け", n), p農地.GetStringValue("営農困難"))
                    SubReplace年月日2(p農地, pSheet04, n, "仮登記日", "設定年月日")

                    If Not p農地.GetStringValue("仮登記氏名") = "" Then
                        .ValueReplace(TargetStr2("仮登記権者氏名", n), p農地.GetStringValue("仮登記氏名"))
                    Else
                        Dim pRow As DataRow = App農地基本台帳.TBL個人.FindRowByID(p農地.GetLongIntValue("仮登記者ID"))
                        If pRow IsNot Nothing Then
                            .ValueReplace(TargetStr2("仮登記権者氏名", n), pRow.Item("氏名"))
                        Else
                            .ValueReplace(TargetStr2("仮登記権者氏名", n), "")
                        End If
                    End If

                    .ValueReplace(TargetStr2("仮登記権者住所", n), p農地.GetStringValue("仮登記住所"))

                    Select Case p農地.GetIntegerValue("環境保全交付金")
                        Case enum有無.有 : .ValueReplace(TargetStr2("農業直接支払交付金", n), "○")
                        Case enum有無.無 : .ValueReplace(TargetStr2("農業直接支払交付金", n), "")
                    End Select
                    Select Case p農地.GetIntegerValue("農地維持交付金")
                        Case enum有無.有 : .ValueReplace(TargetStr2("農地維持支払交付金", n), "○")
                        Case enum有無.無 : .ValueReplace(TargetStr2("農地維持支払交付金", n), "")
                    End Select
                    Select Case p農地.GetIntegerValue("資源向上交付金")
                        Case enum有無.有 : .ValueReplace(TargetStr2("資源向上支払交付金", n), "○")
                        Case enum有無.無 : .ValueReplace(TargetStr2("資源向上支払交付金", n), "")
                    End Select
                    Select Case p農地.GetIntegerValue("中山間直接支払")
                        Case enum有無.有 : .ValueReplace(TargetStr2("中山間地域等直接支払", n), "○")
                        Case enum有無.無 : .ValueReplace(TargetStr2("中山間地域等直接支払", n), "")
                    End Select
                    Select Case p農地.GetIntegerValue("特定処分対象農地等")
                        Case enum有無.有 : .ValueReplace(TargetStr2("特定処分対象農地", n), "○")
                        Case enum有無.無 : .ValueReplace(TargetStr2("特定処分対象農地", n), "")
                    End Select
                End With


                '経営農地5
                With pSheet05
                    Select Case p農地.GetIntegerValue("生産緑地法種別")
                        Case 0 : .ValueReplace(TargetStr2("生産緑地法種別", n), "設定無")
                        Case 1 : .ValueReplace(TargetStr2("生産緑地法種別", n), "新生産緑地指定")
                        Case 2 : .ValueReplace(TargetStr2("生産緑地法種別", n), "旧長期営農継続農地制度認定")
                        Case 3 : .ValueReplace(TargetStr2("生産緑地法種別", n), "旧第一種生産緑地指定")
                        Case 4 : .ValueReplace(TargetStr2("生産緑地法種別", n), "旧第二種生産緑地指定")
                        Case Else : .ValueReplace(TargetStr2("生産緑地法種別", n), "")
                    End Select

                    SubReplace年月日2(p農地, pSheet05, n, "生産緑地法指定日", "生産緑地法指定日")

                    Select Case p農地.GetIntegerValue("共有地区分")
                        Case 0 : .ValueReplace(TargetStr2("共有農地区分", n), "-")
                        Case 1 : .ValueReplace(TargetStr2("共有農地区分", n), "個")
                        Case 2 : .ValueReplace(TargetStr2("共有農地区分", n), "共")
                        Case 3 : .ValueReplace(TargetStr2("共有農地区分", n), "他")
                        Case Else : .ValueReplace(TargetStr2("共有農地区分", n), "")
                    End Select

                    '/*****機能ができ次第追加*****/
                    .ValueReplace(TargetStr2("共有者", n), "")
                    .ValueReplace(TargetStr2("共有農地持分割合", n), "")

                    .ValueReplace(TargetStr2("特定作業者", n), "")
                    .ValueReplace(TargetStr2("特定作業作目", n), "")
                    .ValueReplace(TargetStr2("特定作業内容", n), "")
                    '/****************************/



                    .ValueReplace(TargetStr2("6備考", n), p農地.備考)
                End With
        End Select
    End Sub

    Private Sub SubReplace年月日2(ByRef p農地 As CObj農地, ByRef TargetSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef n As Integer, ByVal TargetDate As String, ByVal TargetCell As String, Optional ByVal n桁数A As Integer = 0)
        If IsDate(p農地.GetDateValue(TargetDate)) AndAlso p農地.GetDateValue(TargetDate).Year > 1902 Then
            With p農地.GetDateValue(TargetDate)
                Dim sDate As String = ""
                If .Date > #4/30/2019# Then
                    sDate = Strings.Right("00" & (.Year - 2018), 2)
                ElseIf .Date > #1/7/1989# Then
                    sDate = Strings.Right("00" & (.Year - 1988), 2)
                End If

                sDate = sDate & "/" & Strings.Right("00" & .Month, 2)
                sDate = sDate & "/" & Strings.Right("00" & .Day, 2)
                TargetSheet.ValueReplace(TargetStr2(TargetCell, n, n桁数A), sDate)
            End With
        Else
            TargetSheet.ValueReplace(TargetStr2(TargetCell, n, n桁数A), " .  .  ")
        End If
    End Sub


    Public Sub sub貸付農地()
        Dim PageCnt As Integer = 10
        If SysAD.市町村.市町村名 = "大崎町" AndAlso mvar印刷Mode = 印刷Mode.簡易印刷 Then
            PageCnt = 7
        End If

        Dim nCount As Decimal = pTBL貸付農地.Count
        Dim nPage As Integer = Math.Floor(nCount / PageCnt)
        Me.n置換え桁数 = 3

        Try

            If nCount > nPage * PageCnt Then
                nPage += 1
            End If

            Dim PgN As String = ""

            If nPage = 0 Then
                mvarXML.WorkBook.WorkSheets.Remove("貸付地の筆別表")
                Exit Sub
            ElseIf nPage = 1 Then

            Else
                For i = 1 To nPage
                    mvarXML.WorkBook.WorkSheets.CopySheet("貸付地の筆別表", "貸付地の筆別表(" & i & ")")
                Next
                mvarXML.WorkBook.WorkSheets.Remove("貸付地の筆別表")
            End If


            Dim sKK As String = ""
            For nA As Integer = 1 To nPage * PageCnt
                Dim n As Integer = ((nA - 1) Mod PageCnt) + 1
                Dim xPage As Integer = Int((nA - 1) / PageCnt) + 1
                If nPage > 1 Then sKK = "(" & xPage & ")"

                Dim pSheet = mvarXML.WorkBook.WorkSheets.Items("貸付地の筆別表" & sKK)
                If pTBL貸付農地.Count >= nA Then
                    Dim p農地 As New CObj農地(pTBL貸付農地.Item(nA - 1).Row, False)

                    With mvar総括表
                        .Add(列.貸付, 地目.田, p農地.田面積)
                        .Add(列.貸付, 地目.畑, p農地.畑面積)
                        .Add(列.貸付, 地目.樹園地, p農地.樹園地)
                        .Add(列.貸付, 地目.採草放牧地, p農地.採草放牧地)
                    End With

                    pSheet.ValueReplace(TargetStr2("貸番号", n), nA.ToString)
                    pSheet.ValueReplace(TargetStr2("貸付地所在", n), p農地.土地所在)

                    If p農地.登記簿地目 <> 0 Then
                        Dim p登記地目 As DataRow = App農地基本台帳.TBL地目.Rows.Find(p農地.登記簿地目)

                        If p登記地目 Is Nothing Then
                            pSheet.ValueReplace(TargetStr2("貸付登記地目", n), "")
                        Else
                            pSheet.ValueReplace(TargetStr2("貸付登記地目", n), p登記地目.Item("名称").ToString)
                        End If
                    Else
                        pSheet.ValueReplace(TargetStr2("貸付登記地目", n), "")
                    End If

                    If p農地.現況地目 <> 0 Then
                        Dim p現況地目 As DataRow = App農地基本台帳.TBL現況地目.Rows.Find(p農地.現況地目)

                        If p現況地目 Is Nothing Then
                            pSheet.ValueReplace(TargetStr2("貸付現況地目", n), "")
                        Else
                            pSheet.ValueReplace(TargetStr2("貸付現況地目", n), p現況地目.Item("名称").ToString)
                        End If
                    Else
                        pSheet.ValueReplace(TargetStr2("貸付現況地目", n), "")
                    End If

                    If p農地.登記簿面積 < 10 Then
                        pSheet.ValueReplace(TargetStr2("貸付登記面積", n), NumToString(p農地.登記簿面積.ToString("F2")))
                    Else
                        pSheet.ValueReplace(TargetStr2("貸付登記面積", n), NumToString(p農地.登記簿面積))
                    End If
                    If p農地.実面積 < 10 Then
                        pSheet.ValueReplace(TargetStr2("貸付面積", n), NumToString(p農地.実面積.ToString("F2")))
                    Else
                        pSheet.ValueReplace(TargetStr2("貸付面積", n), NumToString(p農地.実面積))
                    End If
                    Dim p人情報 As 人情報 = Me.Get人情報(p農地.所有者ID, enum農地関連Mode.貸付, p農地.田面積, p農地.畑面積, p農地.樹園地, p農地.採草放牧地)
                    If p人情報 Is Nothing Then
                        pSheet.ValueReplace(TargetStr2("貸所有者", n), "")
                        pSheet.ValueReplace(TargetStr2("所有者名", n), "")
                        pSheet.ValueReplace(TargetStr2("所有者住所", n), "")
                        pSheet.ValueReplace(TargetStr2("持分割合", n), "")
                    Else
                        Dim p貸付Row As DataRowView = Me.pTBL貸付農地.Item(nA - 1)
                        If Not IsDBNull(p貸付Row.Item("管理者ID")) AndAlso Not p貸付Row.Item("管理者ID") = 0 AndAlso Not p貸付Row.Item("所有者ID") = p貸付Row.Item("管理者ID") Then
                            'Dim p管理者情報 As 人情報 = Me.Get人情報(p貸付Row.Item("管理者ID"), enum農地関連Mode.自作, Val(p貸付Row.Item("田面積").ToString), Val(p貸付Row.Item("畑面積").ToString), Val(p貸付Row.Item("樹園地").ToString), Val(p貸付Row.Item("採草放牧面積").ToString))
                            pSheet.ValueReplace(TargetStr2("所有者名", n), p貸付Row.Item("管理者氏名").ToString)
                            pSheet.ValueReplace(TargetStr2("所有者住所", n), p貸付Row.Item("管理者住所").ToString & "&#10;" & "(" & p貸付Row.Item("所有者氏名").ToString & ")")
                        Else
                            'Dim p所有者情報 As 人情報 = Me.Get人情報(p農地.所有者ID, enum農地関連Mode.自作, Val(p貸付Row.Item("田面積").ToString), Val(p貸付Row.Item("畑面積").ToString), Val(p貸付Row.Item("樹園地").ToString), Val(p貸付Row.Item("採草放牧面積").ToString))
                            pSheet.ValueReplace(TargetStr2("所有者名", n), p貸付Row.Item("所有者氏名").ToString)
                            pSheet.ValueReplace(TargetStr2("所有者住所", n), p貸付Row.Item("所有者住所").ToString)
                        End If

                        If Val(p貸付Row.Item("共有持分分子").ToString) > 0 AndAlso Val(p貸付Row.Item("共有持分分母").ToString) > 0 Then
                            pSheet.ValueReplace(TargetStr2("持分割合", n), "(" & p貸付Row.Item("共有持分分子") & "/" & p貸付Row.Item("共有持分分母") & ")")
                        Else
                            pSheet.ValueReplace(TargetStr2("持分割合", n), "")
                        End If

                        pSheet.ValueReplace(TargetStr2("貸所有者", n), p人情報.氏名.ToString & "&#10;" & p人情報.住所.ToString)
                    End If

                    Dim p人情報B As 人情報 = Me.Get人情報(p農地.借受人ID, enum農地関連Mode.関連なし, 0, 0, 0, 0)
                    If p人情報B Is Nothing Then
                        pSheet.ValueReplace(TargetStr2("借受者氏名", n), "")
                        pSheet.ValueReplace(TargetStr2("借受者住所", n), "")
                    Else
                        pSheet.ValueReplace(TargetStr2("借受者氏名", n), p人情報B.氏名.ToString)
                        pSheet.ValueReplace(TargetStr2("借受者住所", n), p人情報B.住所.ToString)
                    End If

                    Set小作情報(pSheet, p農地, p農地.Row, n, 6)

                    Me.Set和暦TXT(pSheet, p農地.小作開始年月日, TargetStr2("貸付契約開始日", n))
                    Me.Set和暦TXT(pSheet, p農地.小作終了年月日, TargetStr2("貸付契約終了日", n))

                    Me.Set和暦TXT(pSheet, p農地.機構契約開始年月日, TargetStr2("機構契約開始日", n))
                    Me.Set和暦TXT(pSheet, p農地.機構契約終了年月日, TargetStr2("機構契約終了日", n))

                    Select Case p農地.小作地適用法ID
                        Case enum小作地適用法.不明 : pSheet.ValueReplace(TargetStr2("貸付適用法", n), "-")
                        Case enum小作地適用法.農地法 : pSheet.ValueReplace(TargetStr2("貸付適用法", n), "農")
                        Case enum小作地適用法.基盤法 : pSheet.ValueReplace(TargetStr2("貸付適用法", n), "基")
                        Case enum小作地適用法.特定農地貸付法 : pSheet.ValueReplace(TargetStr2("貸付適用法", n), "特")
                        Case enum小作地適用法.その他 : pSheet.ValueReplace(TargetStr2("貸付適用法", n), "他")
                        Case Else
                            pSheet.ValueReplace(TargetStr2("貸付適用法", n), "")
                    End Select

                    With pSheet
                        Select Case p農地.小作形態
                            Case enum小作形態.不明
                                .ValueReplace(TargetStr2("貸付形態", n), "-")
                                .ValueReplace(TargetStr2("貸付料", n), p農地.小作料表示)
                            Case enum小作形態.賃貸借
                                .ValueReplace(TargetStr2("貸付形態", n), "賃")
                                .ValueReplace(TargetStr2("貸付料", n), p農地.小作料表示)
                            Case enum小作形態.使用貸借
                                .ValueReplace(TargetStr2("貸付形態", n), "使貸")
                                .ValueReplace(TargetStr2("貸付料", n), "-")
                            Case enum小作形態.その他
                                .ValueReplace(TargetStr2("貸付形態", n), "他")
                                .ValueReplace(TargetStr2("貸付料", n), p農地.小作料表示)
                            Case enum小作形態.地上権
                                .ValueReplace(TargetStr2("貸付形態", n), "地")
                                .ValueReplace(TargetStr2("貸付料", n), p農地.小作料表示)
                            Case enum小作形態.永小作権
                                .ValueReplace(TargetStr2("貸付形態", n), "永")
                                .ValueReplace(TargetStr2("貸付料", n), p農地.小作料表示)
                            Case enum小作形態.質権
                                .ValueReplace(TargetStr2("貸付形態", n), "質")
                                .ValueReplace(TargetStr2("貸付料", n), "-")
                            Case enum小作形態.期間借地
                                .ValueReplace(TargetStr2("貸付形態", n), "期")
                                .ValueReplace(TargetStr2("貸付料", n), p農地.小作料表示)
                            Case enum小作形態.残存小作地
                                .ValueReplace(TargetStr2("貸付形態", n), "残")
                                .ValueReplace(TargetStr2("貸付料", n), p農地.小作料表示)
                            Case enum小作形態.使用賃借
                                .ValueReplace(TargetStr2("貸付形態", n), "使賃")
                                .ValueReplace(TargetStr2("貸付料", n), "-")
                            Case Else
                                .ValueReplace("{貸付形態" & n & "}", "")
                                .ValueReplace(TargetStr2("貸付料", n), "-")
                        End Select

                        .ValueReplace(TargetStr2("7備考", n), p農地.備考)
                    End With
                Else
                    pSheet.ValueReplace(TargetStr2("貸番号", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付地所在", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付現況地目", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付登記地目", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付登記面積", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付面積", n), "")
                    pSheet.ValueReplace(TargetStr2("貸所有者", n), "")
                    pSheet.ValueReplace(TargetStr2("所有者名", n), "")
                    pSheet.ValueReplace(TargetStr2("所有者住所", n), "")
                    pSheet.ValueReplace(TargetStr2("持分割合", n), "")
                    pSheet.ValueReplace(TargetStr2("借受者氏名", n), "")
                    pSheet.ValueReplace(TargetStr2("借受者住所", n), "")
                    pSheet.ValueReplace(TargetStr2("経由法人名", n), "")
                    pSheet.ValueReplace(TargetStr2("経由法人名B", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付適用法", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付形態", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付契約開始日", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付契約終了日", n), "")
                    pSheet.ValueReplace(TargetStr2("機構契約開始日", n), "")
                    pSheet.ValueReplace(TargetStr2("機構契約終了日", n), "")
                    pSheet.ValueReplace(TargetStr2("貸付料", n), "")
                    pSheet.ValueReplace(TargetStr2("7備考", n), "")
                End If
            Next
        Catch ex As Exception
            If Not SysAD.IsClickOnceDeployed Then
                Stop
            End If
        End Try

    End Sub

    Public Sub Set正否TXT(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByVal pObj As Object, ByVal sTarget As String)
        If IsDBNull(pObj) OrElse pObj = False Then
            pSheet.ValueReplace(sTarget, "")
        Else
            pSheet.ValueReplace(sTarget, "○")
        End If
    End Sub

    Public Sub Set和暦TXT(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByVal pObj As Object, ByVal sTarget As String)
        If IsDBNull(pObj) OrElse Not IsDate(pObj) OrElse Year(pObj) < 1901 Then
            pSheet.ValueReplace(sTarget, "")
        Else
            pSheet.ValueReplace(sTarget, 和暦Format(pObj))
        End If
    End Sub

    Private Sub Set小作情報(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByVal p農地 As CObj農地, ByVal pRow As Object, ByVal n As Integer, ByVal n桁数 As Integer)
        If Not IsDBNull(pRow.Item("経由農業生産法人名")) AndAlso pRow.Item("経由農業生産法人名").length > 0 Then
            If pRow.Item("経由農業生産法人名").Length > n桁数 Then
                pSheet.ValueReplace(TargetStr2("経由法人名", n), "(" & Left(pRow.Item("経由農業生産法人名"), n桁数) & "... 経由)")
                pSheet.ValueReplace(TargetStr2("経由法人名B", n), pRow.Item("経由農業生産法人名") & " 経由")
            Else
                pSheet.ValueReplace(TargetStr2("経由法人名", n), "(" & pRow.Item("経由農業生産法人名") & " 経由)")
                pSheet.ValueReplace(TargetStr2("経由法人名B", n), pRow.Item("経由農業生産法人名") & " 経由")
            End If
        Else
            pSheet.ValueReplace(TargetStr2("経由法人名", n), "")
            pSheet.ValueReplace(TargetStr2("経由法人名B", n), "")
        End If

        With pSheet
            Select Case p農地.小作地適用法ID
                Case enum小作地適用法.不明 : .ValueReplace(TargetStr2("適用法", n), "-")
                Case enum小作地適用法.農地法 : .ValueReplace(TargetStr2("適用法", n), "農")
                Case enum小作地適用法.基盤法 : .ValueReplace(TargetStr2("適用法", n), "基")
                Case enum小作地適用法.特定農地貸付法 : .ValueReplace(TargetStr2("適用法", n), "特")
                Case enum小作地適用法.その他 : .ValueReplace(TargetStr2("適用法", n), "他")
                Case Else : .ValueReplace(TargetStr2("適用法", n), "")
            End Select


            Select Case p農地.小作形態
                Case enum小作形態.不明
                    .ValueReplace(TargetStr2("形態", n), "-")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", p農地.小作料表示, p農地.物納表示 & "(物納)"))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), IIf(p農地.賃借料10a = 0, "", p農地.賃借料10a))
                Case enum小作形態.賃貸借
                    .ValueReplace(TargetStr2("形態", n), "賃")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", p農地.小作料表示, p農地.物納表示 & "(物納)"))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), IIf(p農地.賃借料10a = 0, "", p農地.賃借料10a))
                Case enum小作形態.使用貸借
                    .ValueReplace(TargetStr2("形態", n), "使貸")
                    .ValueReplace("{貸借区分}", IIf(p農地.物納表示 = "", "賃 借 料", "物 納"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", "-", p農地.物納表示))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), "-")
                Case enum小作形態.その他
                    .ValueReplace(TargetStr2("形態", n), "他")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", p農地.小作料表示, p農地.物納表示 & "(物納)"))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), IIf(p農地.賃借料10a = 0, "", p農地.賃借料10a))
                Case enum小作形態.地上権   '必ずしも小作料発生せず
                    .ValueReplace(TargetStr2("形態", n), "地")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", p農地.小作料表示, p農地.物納表示 & "(物納)"))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), IIf(p農地.賃借料10a = 0, "", p農地.賃借料10a))
                Case enum小作形態.永小作権
                    .ValueReplace(TargetStr2("形態", n), "永")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", p農地.小作料表示, p農地.物納表示 & "(物納)"))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), IIf(p農地.賃借料10a = 0, "", p農地.賃借料10a))
                Case enum小作形態.質権    '不動産など
                    .ValueReplace(TargetStr2("形態", n), "質")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "物 納"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", "-", p農地.物納表示))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), "-")
                Case enum小作形態.期間借地
                    .ValueReplace(TargetStr2("形態", n), "期")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", p農地.小作料表示, p農地.物納表示 & "(物納)"))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), IIf(p農地.賃借料10a = 0, "", p農地.賃借料10a))
                Case enum小作形態.残存小作地
                    .ValueReplace(TargetStr2("形態", n), "残")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", p農地.小作料表示, p農地.物納表示 & "(物納)"))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), IIf(p農地.賃借料10a = 0, "", p農地.賃借料10a))
                Case enum小作形態.使用賃借
                    .ValueReplace(TargetStr2("形態", n), "使賃")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "賃借料(物納を含む)"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", "-", p農地.物納表示))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), "-")
                Case Else
                    .ValueReplace("{形態" & n & "}", "")
                    .ValueReplace("賃 借 料", IIf(p農地.物納表示 = "", "賃 借 料", "物 納"))
                    .ValueReplace(TargetStr2("賃借料", n), IIf(p農地.物納表示 = "", "-", p農地.物納表示))
                    .ValueReplace(TargetStr2("10ａ賃借料", n), "-")
            End Select

            Me.Set和暦TXT(pSheet, p農地.小作開始年月日, TargetStr2("契約開始日", n))
            Me.Set和暦TXT(pSheet, p農地.小作終了年月日, TargetStr2("契約終了日", n))

            Me.Set和暦TXT(pSheet, p農地.機構契約開始年月日, TargetStr2("機構契約開始日", n))
            Me.Set和暦TXT(pSheet, p農地.機構契約終了年月日, TargetStr2("機構契約終了日", n))
        End With
    End Sub

    Public Function Get人情報(ByVal nID As Long, ByVal nMode As enum農地関連Mode, ByVal n田面積 As Decimal, ByVal n畑面積 As Decimal, ByVal n樹面積 As Decimal, ByVal n採面積 As Decimal) As 人情報
        Dim p人情報 As 人情報 = Nothing
        If Not mvar人情報Dic.ContainsKey(nID) Then
            Dim pRow As DataRow = App農地基本台帳.TBL個人.Rows.Find(nID)
            If pRow Is Nothing Then
                If Not nID = 0 Then
                    Dim pTBL所有者 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D:個人INFO WHERE [ID]=" & nID)
                    App農地基本台帳.TBL個人.MergePlus(pTBL所有者)
                    pRow = App農地基本台帳.TBL個人.Rows.Find(nID)
                End If
            End If
            If pRow IsNot Nothing Then
                p人情報 = New 人情報
                p人情報.ID = nID
                p人情報.氏名 = pRow.Item("氏名").ToString
                p人情報.住所 = pRow.Item("住所").ToString
                p人情報.住民区分 = Val(pRow.Item("住民区分").ToString)

                If p人情報.氏名 = "内田　次男" Then
                    Dim a As Integer = 2
                End If

                Select Case nMode
                    Case enum農地関連Mode.関連なし
                        p人情報.自作地面積 = n田面積 + n畑面積 + n樹面積 + n採面積
                    Case enum農地関連Mode.自作
                        p人情報.世帯内 = True
                        p人情報.自作地面積 = n田面積 + n畑面積 + n樹面積 + n採面積
                    Case enum農地関連Mode.借受
                        p人情報.世帯内 = True
                        p人情報.借受地面積 = n田面積 + n畑面積 + n樹面積 + n採面積
                    Case enum農地関連Mode.貸付
                        p人情報.世帯内 = True
                        p人情報.貸付地面積 = n田面積 + n畑面積 + n樹面積 + n採面積
                End Select
                mvar人情報Dic.Add(nID, p人情報)
            End If
        Else
            p人情報 = mvar人情報Dic.Item(nID)

            If p人情報.氏名 = "内田　次男" Then
                Dim a As Integer = 2

            End If

            Select Case nMode
                Case enum農地関連Mode.関連なし
                    p人情報.自作地面積 += n田面積 + n畑面積 + n樹面積 + n採面積
                Case enum農地関連Mode.自作
                    p人情報.自作地面積 += n田面積 + n畑面積 + n樹面積 + n採面積
                Case enum農地関連Mode.借受
                    p人情報.借受地面積 += n田面積 + n畑面積 + n樹面積 + n採面積
                Case enum農地関連Mode.貸付
                    p人情報.世帯内 = True
                    p人情報.貸付地面積 += n田面積 + n畑面積 + n樹面積 + n採面積
            End Select


        End If


        Return p人情報
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Enum 列
        所有_耕作 = 0
        借受_耕作 = 1
        貸付 = 2
        所有_非耕作 = 3
        借受_非耕作 = 4
        面積計_耕作 = 5
        面積計_非耕 = 6
    End Enum

    Public Enum 行
        田面積 = 0
        畑面積 = 1
        樹園地 = 2
        田筆数 = 3
        畑筆数 = 4
        樹筆数 = 5
        採草放牧地 = 6
        採草放牧地筆数 = 7
    End Enum

    Public Enum 地目
        田 = 0
        畑 = 1
        樹園地 = 2
        採草放牧地 = 3
    End Enum


    Public Sub sub世帯営農()
        With SysAD.DB(sLRDB)
            Dim pTBL世帯営農 As DataTable = .GetTableBySqlSelect("SELECT * FROM [D:世帯Info] WHERE [ID]=" & 世帯ID)
            Dim pSheetA As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = mvarXML.WorkBook.WorkSheets.Items("世帯員および就業 ")

            If pTBL世帯営農.Rows.Count > 0 Then
                Dim pRow As DataRow = pTBL世帯営農.Rows(0)
                '2-A 農機具
                Dim sLists As New List(Of String)
                sLists.AddRange({"トラクター", "耕運機", "田植機", "コンバイン", "乾燥機", "噴霧器", "その他機具"})

                Dim nFM As Integer = 1
                For Each sList As String In sLists
                    Select Case sList
                        Case "その他器具"
                            If Len(pRow.Item("その他機具内訳").ToString) > 0 Then
                                pSheetA.ValueReplace("{農機具等0" & nFM & "}", pRow.Item("その他機具内訳").ToString)
                            Else
                                pSheetA.ValueReplace("{農機具等0" & nFM & "}", sList)
                            End If
                        Case Else
                            pSheetA.ValueReplace("{農機具等0" & nFM & "}", sList)
                    End Select

                    If Not IsDBNull(pRow.Item(sList & "台数")) AndAlso pRow.Item(sList & "台数") > 0 Then
                        pSheetA.ValueReplace("{農機具等数量-100" & nFM & "}", pRow.Item(sList & "台数"))
                    Else
                        pSheetA.ValueReplace("{農機具等数量-100" & nFM & "}", "")
                    End If

                    nFM += 1
                Next

                '2-B
                sLists = New List(Of String)
                sLists.AddRange({"-", "米", "畜産", "果樹", "そさい", "養蚕", "その他", "", ""})
                nFM = 1

                For n As Integer = 1 To 8
                    pSheetA.ValueReplace("{販売収入0" & n & "}", sLists(n))
                    pSheetA.ValueReplace("{販売収入-300" & n & "}", "")
                Next
                '

                '2-C 家畜
                sLists = New List(Of String)
                sLists.AddRange({"肉用牛", "乳牛", "豚", "採卵用鶏", "ブロイラー", "その他家畜"})
                nFM = 1
                For Each sList As String In sLists
                    pSheetA.ValueReplace("{家畜0" & nFM & "}", sList)

                    Select Case sList
                        Case "肉用牛", "乳牛", "豚", "その他家畜"
                            If Not IsDBNull(pRow.Item(sList & "頭数")) AndAlso pRow.Item(sList & "頭数") > 0 Then
                                pSheetA.ValueReplace("{家畜数量-200" & nFM & "}", pRow.Item(sList & "頭数"))
                            Else
                                pSheetA.ValueReplace("{家畜数量-200" & nFM & "}", "")
                            End If
                        Case "採卵用鶏", "ブロイラー"
                            If Not IsDBNull(pRow.Item(sList & "羽数")) AndAlso pRow.Item(sList & "羽数") > 0 Then
                                pSheetA.ValueReplace("{家畜数量-200" & nFM & "}", pRow.Item(sList & "羽数"))
                            Else
                                pSheetA.ValueReplace("{家畜数量-200" & nFM & "}", "")
                            End If
                    End Select

                    nFM += 1
                Next



            Else

                '2営農の状況
                pSheetA.ValueReplace("{農機具等01}", "トラクター")
                pSheetA.ValueReplace("{農機具等02}", "耕運機")
                pSheetA.ValueReplace("{農機具等03}", "田植機")
                pSheetA.ValueReplace("{農機具等04}", "コンバイン")
                pSheetA.ValueReplace("{農機具等05}", "乾燥機")
                pSheetA.ValueReplace("{農機具等06}", "噴霧器")
                pSheetA.ValueReplace("{農機具等07}", "その他機具")
                pSheetA.ValueReplace("{農機具等数量-1001}", "")
                pSheetA.ValueReplace("{農機具等数量-1002}", "")
                pSheetA.ValueReplace("{農機具等数量-1003}", "")
                pSheetA.ValueReplace("{農機具等数量-1004}", "")
                pSheetA.ValueReplace("{農機具等数量-1005}", "")
                pSheetA.ValueReplace("{農機具等数量-1006}", "")
                pSheetA.ValueReplace("{農機具等数量-1007}", "")
                pSheetA.ValueReplace("{販売収入01}", "米")
                pSheetA.ValueReplace("{販売収入02}", "畜産")
                pSheetA.ValueReplace("{販売収入03}", "果樹")
                pSheetA.ValueReplace("{販売収入04}", "そさい")
                pSheetA.ValueReplace("{販売収入05}", "養蚕")
                pSheetA.ValueReplace("{販売収入06}", "")
                pSheetA.ValueReplace("{販売収入07}", "")
                pSheetA.ValueReplace("{販売収入08}", "")
                pSheetA.ValueReplace("{販売収入-3001}", "")
                pSheetA.ValueReplace("{販売収入-3002}", "")
                pSheetA.ValueReplace("{販売収入-3003}", "")
                pSheetA.ValueReplace("{販売収入-3004}", "")
                pSheetA.ValueReplace("{販売収入-3005}", "")
                pSheetA.ValueReplace("{販売収入-3006}", "")
                pSheetA.ValueReplace("{販売収入-3007}", "")
                pSheetA.ValueReplace("{販売収入-3008}", "")
                pSheetA.ValueReplace("{家畜01}", "肉用牛")
                pSheetA.ValueReplace("{家畜02}", "乳牛")
                pSheetA.ValueReplace("{家畜03}", "豚")
                pSheetA.ValueReplace("{家畜04}", "採卵用鶏")
                pSheetA.ValueReplace("{家畜05}", "ブロイラー")
                pSheetA.ValueReplace("{家畜06}", "その他家畜")
                pSheetA.ValueReplace("{家畜数量-2001}", "")
                pSheetA.ValueReplace("{家畜数量-2002}", "")
                pSheetA.ValueReplace("{家畜数量-2003}", "")
                pSheetA.ValueReplace("{家畜数量-2004}", "")
                pSheetA.ValueReplace("{家畜数量-2005}", "")
                pSheetA.ValueReplace("{家畜数量-2006}", "")
                pSheetA.ValueReplace("{青色申告}", "青色申告")
                pSheetA.ValueReplace("{申告年}", "　　　")
                pSheetA.ValueReplace("{白色申告}", "白色申告")
                pSheetA.ValueReplace("{その他}", "その他")
                pSheetA.ValueReplace("{認定農業者有無}", "")
                pSheetA.ValueReplace("{担い手農家有無}", "")
                pSheetA.ValueReplace("{家族経営協定}", "")
            End If

            'D_世帯営農
            Dim pSheetB As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet = mvarXML.WorkBook.WorkSheets.Items("土地総括表")

            With mvar総括表
                If .Data(列.所有_非耕作, 行.田筆数) > 0 Then
                    pSheetB.ValueReplace("{田自作地}", NumToString(.Data(列.所有_耕作, 行.田面積)) & "&#10; (" & NumToString(.Data(列.所有_非耕作, 行.田面積)) & ")")
                Else
                    pSheetB.ValueReplace("{田自作地}", NumToString(.Data(列.所有_耕作, 行.田面積)))
                End If

                If .Data(列.所有_非耕作, 行.畑筆数) > 0 Then
                    pSheetB.ValueReplace("{畑自作地}", NumToString(.Data(列.所有_耕作, 行.畑面積)) & "&#10; (" & NumToString(.Data(列.所有_非耕作, 行.畑面積)) & ")")
                Else
                    pSheetB.ValueReplace("{畑自作地}", NumToString(.Data(列.所有_耕作, 行.畑面積)))
                End If

                If .Data(列.所有_非耕作, 行.樹筆数) > 0 Then
                    pSheetB.ValueReplace("{樹園地自作地}", NumToString(.Data(列.所有_耕作, 行.樹園地)) & "&#10; (" & NumToString(.Data(列.所有_非耕作, 行.樹園地)) & ")")
                Else
                    pSheetB.ValueReplace("{樹園地自作地}", NumToString(.Data(列.所有_耕作, 行.樹園地)))
                End If

                If .Data(列.所有_非耕作, 行.採草放牧地筆数) > 0 Then
                    pSheetB.ValueReplace("{採草放牧地自作地}", NumToString(.Data(列.所有_耕作, 行.採草放牧地)) & "&#10; (" & NumToString(.Data(列.所有_非耕作, 行.採草放牧地)) & ")")
                Else
                    pSheetB.ValueReplace("{採草放牧地自作地}", NumToString(.Data(列.所有_耕作, 行.採草放牧地)))
                End If

                If .Data(列.借受_非耕作, 行.田筆数) > 0 Then
                    pSheetB.ValueReplace("{田借入地}", NumToString(.Data(列.借受_耕作, 行.田面積)) & "&#10; (" & NumToString(.Data(列.借受_非耕作, 行.田面積)) & ")")
                Else
                    pSheetB.ValueReplace("{田借入地}", NumToString(.Data(列.借受_耕作, 行.田面積)))
                End If

                If .Data(列.借受_非耕作, 行.畑筆数) > 0 Then
                    pSheetB.ValueReplace("{畑借入地}", NumToString(.Data(列.借受_耕作, 行.畑面積)) & "&#10; (" & NumToString(.Data(列.借受_非耕作, 行.畑面積)) & ")")
                Else
                    pSheetB.ValueReplace("{畑借入地}", NumToString(.Data(列.借受_耕作, 行.畑面積)))
                End If

                If .Data(列.借受_非耕作, 行.樹筆数) > 0 Then
                    pSheetB.ValueReplace("{樹園地借入地}", NumToString(.Data(列.借受_耕作, 行.樹園地)) & "&#10; (" & NumToString(.Data(列.借受_非耕作, 行.樹園地)) & ")")
                Else
                    pSheetB.ValueReplace("{樹園地借入地}", NumToString(.Data(列.借受_耕作, 行.樹園地)))
                End If

                If .Data(列.借受_非耕作, 行.採草放牧地筆数) > 0 Then
                    pSheetB.ValueReplace("{採草放牧地借入地}", NumToString(.Data(列.借受_耕作, 行.採草放牧地)) & "&#10; (" & NumToString(.Data(列.借受_非耕作, 行.採草放牧地)) & ")")
                Else
                    pSheetB.ValueReplace("{採草放牧地借入地}", NumToString(.Data(列.借受_耕作, 行.採草放牧地)))
                End If

                .Calc()
                If .Data(列.面積計_非耕, 行.田筆数) > 0 Then
                    pSheetB.ValueReplace("{田総面積}", NumToString(.Data(列.面積計_耕作, 行.田面積)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.田面積)) & ")")
                    pSheetB.ValueReplace("{田筆数}", NumToString(.Data(列.面積計_耕作, 行.田筆数)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.田筆数)) & ")")
                Else
                    pSheetB.ValueReplace("{田総面積}", NumToString(.Data(列.面積計_耕作, 行.田面積)))
                    pSheetB.ValueReplace("{田筆数}", NumToString(.Data(列.面積計_耕作, 行.田筆数)))
                End If

                If .Data(列.面積計_非耕, 行.畑筆数) > 0 Then
                    pSheetB.ValueReplace("{畑総面積}", NumToString(.Data(列.面積計_耕作, 行.畑面積)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.畑面積)) & ")")
                    pSheetB.ValueReplace("{畑筆数}", NumToString(.Data(列.面積計_耕作, 行.畑筆数)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.畑筆数)) & ")")
                Else
                    pSheetB.ValueReplace("{畑総面積}", NumToString(.Data(列.面積計_耕作, 行.畑面積)))
                    pSheetB.ValueReplace("{畑筆数}", NumToString(.Data(列.面積計_耕作, 行.畑筆数)))
                End If
                If .Data(列.面積計_非耕, 行.樹筆数) > 0 Then
                    pSheetB.ValueReplace("{樹園地総面積}", NumToString(.Data(列.面積計_耕作, 行.樹園地)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.樹園地)) & ")")
                    pSheetB.ValueReplace("{樹園地筆数}", NumToString(.Data(列.面積計_耕作, 行.樹筆数)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.樹筆数)) & ")")
                Else
                    pSheetB.ValueReplace("{樹園地総面積}", NumToString(.Data(列.面積計_耕作, 行.樹園地)))
                    pSheetB.ValueReplace("{樹園地筆数}", NumToString(.Data(列.面積計_耕作, 行.樹筆数)))
                End If
                If .Data(列.面積計_非耕, 行.採草放牧地筆数) > 0 Then
                    pSheetB.ValueReplace("{採草放牧地総面積}", NumToString(.Data(列.面積計_耕作, 行.採草放牧地)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.採草放牧地)) & ")")
                    pSheetB.ValueReplace("{採草放牧地筆数}", NumToString(.Data(列.面積計_耕作, 行.採草放牧地筆数)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.採草放牧地筆数)) & ")")
                Else
                    pSheetB.ValueReplace("{採草放牧地総面積}", NumToString(.Data(列.面積計_耕作, 行.採草放牧地)))
                    pSheetB.ValueReplace("{採草放牧地筆数}", NumToString(.Data(列.面積計_耕作, 行.採草放牧地筆数)))
                End If
                '

                If .Data(列.所有_非耕作, 行.田面積) + .Data(列.所有_非耕作, 行.畑面積) + .Data(列.所有_非耕作, 行.樹園地) + .Data(列.所有_非耕作, 行.採草放牧地) > 0 Then
                    pSheetB.ValueReplace("{自作地耕地計}", NumToString(.Data(列.所有_耕作, 行.田面積) + .Data(列.所有_耕作, 行.畑面積) + .Data(列.所有_耕作, 行.樹園地) + .Data(列.所有_耕作, 行.採草放牧地)) & "&#10;(" & NumToString(.Data(列.所有_非耕作, 行.田面積) + .Data(列.所有_非耕作, 行.畑面積) + .Data(列.所有_非耕作, 行.樹園地) + .Data(列.所有_非耕作, 行.採草放牧地)) & ")")
                Else
                    pSheetB.ValueReplace("{自作地耕地計}", NumToString(.Data(列.所有_耕作, 行.田面積) + .Data(列.所有_耕作, 行.畑面積) + .Data(列.所有_耕作, 行.樹園地) + .Data(列.所有_耕作, 行.採草放牧地)))
                End If

                If .Data(列.借受_非耕作, 行.田面積) + .Data(列.借受_非耕作, 行.畑面積) + .Data(列.借受_非耕作, 行.樹園地) + .Data(列.所有_非耕作, 行.採草放牧地) > 0 Then
                    pSheetB.ValueReplace("{借入地耕地計}", NumToString(.Data(列.借受_耕作, 行.田面積) + .Data(列.借受_耕作, 行.畑面積) + .Data(列.借受_耕作, 行.樹園地) + .Data(列.借受_耕作, 行.採草放牧地)) & "&#10;(" & NumToString(.Data(列.借受_非耕作, 行.田面積) + .Data(列.借受_非耕作, 行.畑面積) + .Data(列.借受_非耕作, 行.樹園地) + .Data(列.借受_非耕作, 行.採草放牧地)) & ")")
                Else
                    pSheetB.ValueReplace("{借入地耕地計}", NumToString(.Data(列.借受_耕作, 行.田面積) + .Data(列.借受_耕作, 行.畑面積) + .Data(列.借受_耕作, 行.樹園地) + .Data(列.借受_耕作, 行.採草放牧地)))
                End If


                If .Data(列.面積計_非耕, 行.田面積) + .Data(列.面積計_非耕, 行.畑面積) + .Data(列.面積計_非耕, 行.樹園地) + .Data(列.所有_非耕作, 行.採草放牧地) > 0 Then
                    pSheetB.ValueReplace("{総面積耕地計}", NumToString(.Data(列.面積計_耕作, 行.田面積) + .Data(列.面積計_耕作, 行.畑面積) + .Data(列.面積計_耕作, 行.樹園地) + .Data(列.面積計_耕作, 行.採草放牧地)) & "&#10;(" & NumToString(.Data(列.面積計_非耕, 行.田面積) + .Data(列.面積計_非耕, 行.畑面積) + .Data(列.面積計_非耕, 行.樹園地) + .Data(列.面積計_非耕, 行.採草放牧地)) & ")")
                    pSheetB.ValueReplace("{筆数計}", NumToString(.Data(列.面積計_耕作, 行.田筆数) + .Data(列.面積計_耕作, 行.畑筆数) + .Data(列.面積計_耕作, 行.樹筆数) + .Data(列.面積計_耕作, 行.採草放牧地筆数)) & "&#10; (" & NumToString(.Data(列.面積計_非耕, 行.田筆数) + .Data(列.面積計_非耕, 行.畑筆数) + .Data(列.面積計_非耕, 行.樹筆数) + .Data(列.面積計_非耕, 行.採草放牧地筆数)) & ")")
                Else
                    pSheetB.ValueReplace("{総面積耕地計}", NumToString(.Data(列.面積計_耕作, 行.田面積) + .Data(列.面積計_耕作, 行.畑面積) + .Data(列.面積計_耕作, 行.樹園地) + .Data(列.面積計_耕作, 行.採草放牧地)))
                    pSheetB.ValueReplace("{筆数計}", NumToString(.Data(列.面積計_耕作, 行.田筆数) + .Data(列.面積計_耕作, 行.畑筆数) + .Data(列.面積計_耕作, 行.樹筆数) + .Data(列.面積計_耕作, 行.採草放牧地筆数)))
                End If


                pSheetB.ValueReplace("{田貸付面積}", NumToString(.Data(列.貸付, 行.田面積)))
                pSheetB.ValueReplace("{畑貸付面積}", NumToString(.Data(列.貸付, 行.畑面積)))
                pSheetB.ValueReplace("{樹園地貸付面積}", NumToString(.Data(列.貸付, 行.樹園地)))
                pSheetB.ValueReplace("{採草放牧地貸付面積}", NumToString(.Data(列.貸付, 行.採草放牧地)))
                pSheetB.ValueReplace("{貸付面積耕地計}", NumToString(.Data(列.貸付, 行.田面積) + .Data(列.貸付, 行.畑面積) + .Data(列.貸付, 行.樹園地) + .Data(列.貸付, 行.採草放牧地)))


                '4～6
                pSheetB.ValueReplace("{意向1}", "1")
                pSheetB.ValueReplace("{意向2}", "2")
                pSheetB.ValueReplace("{意向3}", "3")
                pSheetB.ValueReplace("{意向4}", "4")
                pSheetB.ValueReplace("{計画01}", "1")
                pSheetB.ValueReplace("{計画02}", "2")
                pSheetB.ValueReplace("{計画03}", "3")
                pSheetB.ValueReplace("{拡大01}", "")
                pSheetB.ValueReplace("{拡大02}", "")
                pSheetB.ValueReplace("{拡大03}", "")
                pSheetB.ValueReplace("{拡大04}", "")
                pSheetB.ValueReplace("{拡大05}", "")
                pSheetB.ValueReplace("{拡大06}", "")
                pSheetB.ValueReplace("{部門01}", "")
                pSheetB.ValueReplace("{部門02}", "")
                pSheetB.ValueReplace("{部門03}", "")
                pSheetB.ValueReplace("{部門04}", "")
                pSheetB.ValueReplace("{部門05}", "")
                pSheetB.ValueReplace("{部門06}", "")
                pSheetB.ValueReplace("{縮小01}", "")
                pSheetB.ValueReplace("{縮小02}", "")
                pSheetB.ValueReplace("{縮小03}", "")
                pSheetB.ValueReplace("{縮小04}", "")
                pSheetB.ValueReplace("{縮小05}", "")
                pSheetB.ValueReplace("{縮小06}", "")
                pSheetB.ValueReplace("{拡大方01}", "1")
                pSheetB.ValueReplace("{拡大方02}", "2")
                pSheetB.ValueReplace("{拡大方03}", "3")
                pSheetB.ValueReplace("{拡大方04}", "4")
                pSheetB.ValueReplace("{縮小方01}", "1")
                pSheetB.ValueReplace("{縮小方02}", "2")
                pSheetB.ValueReplace("{縮小方03}", "3")
                pSheetB.ValueReplace("{縮小方04}", "4")
                pSheetB.ValueReplace("{参加01}", "1")
                pSheetB.ValueReplace("{参加02}", "2")
                pSheetB.ValueReplace("{参加03}", "3")
                pSheetB.ValueReplace("{生産組織の参加状況その他}", "　　　　　　　　　　")
                pSheetB.ValueReplace("{改善団体有無}", "有・無")
                pSheetB.ValueReplace("{農業集団有無}", "有・無")


                If pTBL世帯営農.Rows.Count > 0 Then
                    Dim pRow As DataRow = pTBL世帯営農.Rows(0)
                    If pRow IsNot Nothing Then
                        pSheetB.ValueReplace("{土地総括備考}", pRow.Item("備考").ToString)
                    End If
                End If
            End With


            pSheetB.ValueReplace("{総面積耕地計}", "")


            pSheetB.ValueReplace("{採草放牧地総面積}", "0")
            pSheetB.ValueReplace("{採草放牧地自作地}", "0")
            pSheetB.ValueReplace("{採草放牧地借入地}", "0")
            pSheetB.ValueReplace("{採草放牧地筆数}", "0")
            pSheetB.ValueReplace("{採草放牧地貸付面積}", "0")
            pSheetB.ValueReplace("{田備考}", "")
            pSheetB.ValueReplace("{畑備考}", "")
            pSheetB.ValueReplace("{樹園地備考}", "")
            pSheetB.ValueReplace("{耕地備考}", "")
            pSheetB.ValueReplace("{採草放牧地備考}", "")
            pSheetB.ValueReplace("{土地総括備考}", "")

            Dim L As Integer = 1
            For n As Integer = 1 To mvar人情報Dic.Count
                If mvar人情報Dic.Values(n - 1).世帯内 = True Then
                    Dim sPX As String = IIf(mvar人情報Dic.Values(n - 1).住民区分 = 0, "", "(")
                    Dim sBX As String = IIf(mvar人情報Dic.Values(n - 1).住民区分 = 0, "", ")")
                    pSheetB.ValueReplace("{氏名0" & L & "}", sPX & mvar人情報Dic.Values(n - 1).氏名 & sBX)
                    pSheetB.ValueReplace("{自作地0" & L & "}", NumToString(mvar人情報Dic.Values(n - 1).自作地面積))
                    pSheetB.ValueReplace("{借入地0" & L & "}", NumToString(mvar人情報Dic.Values(n - 1).借受地面積))
                    pSheetB.ValueReplace("{特定処分対象農地0" & L & "}", NumToString(mvar人情報Dic.Values(n - 1).特定処分対象地面積))
                    pSheetB.ValueReplace("{耕地計0" & L & "}", NumToString(mvar人情報Dic.Values(n - 1).自作地面積 + mvar人情報Dic.Values(n - 1).借受地面積))
                    pSheetB.ValueReplace("{貸付地0" & L & "}", NumToString(mvar人情報Dic.Values(n - 1).貸付地面積))
                    L += 1
                End If
            Next


            Do Until L > 10
                pSheetB.ValueReplace("{氏名0" & L & "}", "")
                pSheetB.ValueReplace("{自作地0" & L & "}", "")
                pSheetB.ValueReplace("{借入地0" & L & "}", "")
                pSheetB.ValueReplace("{特定処分対象農地0" & L & "}", "")
                pSheetB.ValueReplace("{耕地計0" & L & "}", "")
                pSheetB.ValueReplace("{貸付地0" & L & "}", "")
                L += 1
            Loop


        End With
    End Sub

    'Private Sub 総括面積表示(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByVal 非耕作面積 As Decimal, ByVal 耕作面積 As Decimal, ByVal 置き換え文字 As String)
    '    If 非耕作面積 > 0 Then
    '        pSheet.ValueReplace("{" & 置き換え文字 & "}", String.Format("{0,9}", 耕作面積) & "&#10; (" & String.Format("{0,9}", 非耕作面積) & ")")
    '    Else
    '        pSheet.ValueReplace("{" & 置き換え文字 & "}", String.Format("{0,9}", 耕作面積))
    '    End If
    'End Sub

    Public Class C総括表
        Public Table(7, 7) As Decimal

        Public Sub New()
            For i As Integer = 0 To 7
                For j As Integer = 0 To 7
                    Table(i, j) = 0
                Next
            Next
        End Sub

        Public Sub Add(ByVal X As 列, ByVal Y As 地目, ByVal Value As Decimal)
            Select Case Y
                Case 地目.田 : Table(X, 行.田面積) += Value
                    Table(X, 行.田筆数) += -((Value > 0))
                    If X = 列.所有_非耕作 Then

                    End If
                    'If Value > 0 Then
                    '    Stop
                    'End If
                Case 地目.畑 : Table(X, 行.畑面積) += Value
                    Table(X, 行.畑筆数) += -((Value > 0))
                Case 地目.樹園地 : Table(X, 行.樹園地) += Value
                    Table(X, 行.樹筆数) += -((Value > 0))
                Case 地目.採草放牧地 : Table(X, 行.採草放牧地) += Value
                    Table(X, 行.採草放牧地筆数) += -((Value > 0))
            End Select

        End Sub

        Public Property Data(ByVal X As 列, ByVal Y As 行) As Decimal
            Get
                Return Table(X, Y)
            End Get
            Set(ByVal value As Decimal)
                Table(X, Y) = value
            End Set
        End Property

        Public Sub Calc()
            For n = 0 To 6
                Table(列.面積計_耕作, n) = Data(列.所有_耕作, n) + Data(列.借受_耕作, n)
                Table(列.面積計_非耕, n) = Data(列.所有_非耕作, n) + Data(列.借受_非耕作, n)
            Next


        End Sub



    End Class
End Class

Public Class IDList
    Inherits List(Of Long)

    Public Sub New()

    End Sub

    Public Shadows Sub Add(ByVal pOBJ As Object)
        If Not IsDBNull(pOBJ) AndAlso pOBJ <> 0 AndAlso Not MyBase.Contains(pOBJ) Then
            MyBase.Add(pOBJ)
        End If
    End Sub

    Public Overrides Function ToString() As String
        If Me.Count = 0 Then
            Return ""
        Else
            Try
                Dim sB As New System.Text.StringBuilder
                For Each n As Long In Me
                    sB.Append(IIf(sB.Length > 0, ",", "") & n.ToString)
                Next
                Return sB.ToString
            Catch ex As Exception
                Return ""
            End Try
        End If
    End Function
End Class

Public Class 人情報
    Public ID As Long
    Public 世帯内 As Boolean = False
    Public 住民区分 As Integer = 0
    Public 氏名 As String
    Public 住所 As String
    Public 自作地面積 As Decimal = 0
    Public 特定処分対象地面積 As Decimal = 0
    Public 借受地面積 As Decimal = 0
    Public 貸付地面積 As Decimal = 0
End Class
