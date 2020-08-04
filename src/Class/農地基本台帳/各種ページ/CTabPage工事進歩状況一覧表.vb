Imports HimTools2012.CommonFunc
Imports HimTools2012.Excel.XMLSS2003

Public Class CTabPage工事進歩状況一覧表
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarTable As DataTable

    Public Sub New()
        MyBase.New(True, True, "工事進歩状況一覧表", "工事進歩状況一覧表", HimTools2012.controls.CloseMode.NoMessage)

        Dim pCls As New InputStartAndEndDate
        With New HimTools2012.PropertyGridDialog(pCls, "工事進捗状況一覧", "一覧を作成する転用申請が許可された期間を入力してください。")
            If .ShowDialog() = DialogResult.Yes Then
                'pCls.開始日
                Dim sWhere As String = HimTools2012.StringF.Toリテラル日付(pCls.開始日)
                mvarTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_申請.ID, D_申請.法令, D_申請.状態, D_申請.名称, D_申請.許可年月日, D_申請.許可番号, D_申請.農地リスト, D_申請.申請者A, D_申請.氏名A, D_申請.住所A, D_申請.申請者B, D_申請.氏名B, D_申請.住所B, D_申請.申請理由A, D_申請.所有権移転の種類, D_申請.完了報告年月日 FROM D_申請 WHERE ([D_申請].[法令] IN (40,50,51)) AND (D_申請.状態)=2 ORDER BY D_申請.許可年月日, D_申請.許可番号;")
            End If
        End With



    End Sub


    Public Sub Exec()
        Dim sFile As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\工事進歩状況一覧表.xml"

        If IO.File.Exists(sFile) Then
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_申請.ID, D_申請.法令, D_申請.状態, D_申請.名称, D_申請.許可年月日, D_申請.許可番号, D_申請.農地リスト, D_申請.申請者A, D_申請.氏名A, D_申請.住所A, D_申請.申請者B, D_申請.氏名B, D_申請.住所B, D_申請.申請理由A, D_申請.所有権移転の種類, D_申請.完了報告年月日 FROM D_申請 WHERE (([D_申請].[法令] IN (40,50,51)) AND ((D_申請.状態)=2) AND ((D_申請.完了報告年月日) Is Null Or (D_申請.完了報告年月日)<#1/1/2000#)) ORDER BY D_申請.許可年月日, D_申請.許可番号;")
            Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(sFile)
            Dim XMLSS As New CXMLSS2003(sXML)

            With XMLSS.WorkBook.WorkSheets.Items("一覧表")
                Dim LoopRows As New XMLLoopRows(._object)
                Dim nLoop As Integer = -1

                For Each pRow As DataRow In pTBL.Rows
                    If nLoop = -1 Then

                    Else
                        For Each pXRow As XMLSSRow In LoopRows
                            Dim pCopyRow = pXRow.CopyRow

                            .Table.Rows.InsertRow(LoopRows.InsetRow, pCopyRow)
                            LoopRows.InsetRow += 1
                        Next
                    End If

                    nLoop += 1
                    Dim 申請者ID As Decimal = 0

                    Select Case Val(pRow("法令").ToString)
                        Case enum法令.農地法4条
                            .ValueReplace("{転用事業者氏名}", pRow("氏名A").ToString)
                            .ValueReplace("{転用事業者住所}", pRow("住所A").ToString)
                            申請者ID = Val(pRow.Item("申請者A").ToString)
                        Case enum法令.農地法5条所有権, enum法令.農地法5条貸借
                            .ValueReplace("{転用事業者氏名}", pRow("氏名B").ToString)
                            .ValueReplace("{転用事業者住所}", pRow("住所B").ToString)
                            申請者ID = Val(pRow.Item("申請者B").ToString)
                        Case enum法令.事業計画変更
                            .ValueReplace("{転用事業者氏名}", pRow("氏名C").ToString)
                            .ValueReplace("{転用事業者住所}", pRow("住所C").ToString)
                            申請者ID = Val(pRow.Item("申請者C").ToString)
                    End Select

                    .ValueReplace("{No}", "" & (nLoop + 1).ToString)
                    .ValueReplace("{許可年月日}", Replace(CDate(pRow.Item("許可年月日")).ToString("yyyy.M.d"), "/", "."))
                    .ValueReplace("{許可番号}", pRow.Item("許可番号").ToString)

                    .ValueReplace("{転用目的}", pRow("申請理由A").ToString)

                    Dim pPRow = App農地基本台帳.TBL個人.FindRowByID(GetKeyCode(申請者ID))

                    If pPRow IsNot Nothing Then
                        .ValueReplace("{転用事業者電話番号}", pPRow.Item("電話番号").ToString)
                        .ValueReplace("{転用事業者集落名}", pPRow.Item("行政区名").ToString)
                    End If

                    Dim 転用許可地 As String = ""
                    Dim 転用農地Count As Integer = -1
                    Dim 転用面積_田 As Decimal = 0
                    Dim 転用面積_畑 As Decimal = 0

                    Dim St農地 As String = pRow.Item("農地リスト").ToString
                    For Each s農地Key As String In Split(St農地, ";")
                        If s農地Key.Length > 0 Then
                            Dim pNRow As DataRow = Nothing
                            Select Case GetKeyHead(s農地Key)
                                Case "農地"
                                    pNRow = App農地基本台帳.TBL農地.FindRowByID(GetKeyCode(s農地Key))
                                Case "転用農地"
                                    pNRow = App農地基本台帳.TBL転用農地.FindRowByID(GetKeyCode(s農地Key))
                            End Select

                            If pNRow IsNot Nothing Then
                                If 転用許可地.Length = 0 Then
                                    転用許可地 = pNRow.Item("土地所在")
                                End If
                                If Val(pNRow.Item("田面積").ToString) > 0 OrElse pNRow.Item("登記簿地目名") = "田" Then
                                    転用面積_田 += pNRow.Item("実面積")
                                ElseIf Val(pNRow.Item("畑面積").ToString) > 0 OrElse pNRow.Item("登記簿地目名") = "畑" Then
                                    転用面積_畑 += pNRow.Item("実面積")
                                End If
                            End If

                            転用農地Count += 1
                        End If
                        If 転用農地Count > 0 Then
                            転用許可地 &= String.Format(" 外{0}筆", 転用農地Count)
                        End If
                    Next
                    .ValueReplace("{土地の所在}", 転用許可地)
                    .ValueReplace("{田面積計}", String.Format("{0,8}", 転用面積_田))
                    .ValueReplace("{畑面積計}", String.Format("{0,8}", 転用面積_畑))
                Next

                Dim SavePath As String = SysAD.OutputFolder & String.Format("\工事進捗状況報告書.xml")

                HimTools2012.TextAdapter.SaveTextFile(SavePath, XMLSS.OutPut(True))
                Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                    pExcel.ShowPreview(SavePath)
                End Using
            End With
        End If
    End Sub
End Class

Public Class InputStartAndEndDate
    Inherits HimTools2012.InputSupport.CInputSupport

    Public Sub New()
        MyBase.New(Nothing)
        開始日 = Now.Date
        終了日 = Now.Date
    End Sub

    Public Property 開始日 As Date
    Public Property 終了日 As Date
    Public Overrides Function DataCompleate() As Boolean
        Return IsDate(開始日) AndAlso IsDate(終了日)
    End Function
End Class