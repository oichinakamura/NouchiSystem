Imports HimTools2012.CommonFunc

Public Class CPrint工事進捗状況報告書
    Inherits C印刷処理

    Private 発行年月日 As DateTime
    Private 申請者氏名 As String = ""
    Private 申請者住所 As String = ""
    Private 許可年月日 As DateTime
    Private 許可番号 As Integer = 0
    Private 転用許可地 As String = ""
    Private 転用目的 As String = ""
    Private 転用面積_田 As Decimal = 0
    Private 転用面積_畑 As Decimal = 0
    Private 転用面積_草 As Decimal = 0
    Private mvarFileName As String = ""
    Private mvar申請 As CObj申請

    Public Sub New(ByRef p申請 As CObj申請)
        mvarFileName = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\工事進捗状況報告書.xml"

        If Not IO.File.Exists(mvarFileName) Then
            MsgBox("[工事進捗状況報告書.xml]が見つかりません。作成してください。", vbCritical)
        Else
            mvar申請 = p申請
            Dim s発行日 As String = InputBox("発行年月日を入力してください", "", Now.Date.ToString.Replace(" 0:00:00", ""))
            If s発行日 IsNot Nothing AndAlso IsDate(s発行日) Then
                発行年月日 = s発行日
                Me.ShowDialog()
            End If
        End If
    End Sub


    Public Overrides Sub Execute()

        許可年月日 = mvar申請.許可年月日
        許可番号 = mvar申請.許可番号
        転用目的 = mvar申請.Row.Body("申請理由A").ToString

        Dim 転用農地Count As Integer = -1

        Dim St農地 As String = mvar申請.Row.Body.Item("農地リスト").ToString
        For Each s農地Key As String In Split(St農地, ";")
            If s農地Key.Length > 0 Then
                Dim pRow As DataRow = Nothing
                Select Case GetKeyHead(s農地Key)
                    Case "農地"
                        pRow = App農地基本台帳.TBL農地.FindRowByID(GetKeyCode(s農地Key))
                    Case "転用農地"
                        pRow = App農地基本台帳.TBL転用農地.FindRowByID(GetKeyCode(s農地Key))
                End Select

                If pRow IsNot Nothing Then
                    If 転用許可地.Length = 0 Then
                        転用許可地 = pRow.Item("土地所在")
                    End If
                    If Val(pRow.Item("田面積").ToString) > 0 OrElse pRow.Item("登記簿地目名") = "田" Then
                        転用面積_田 += pRow.Item("実面積")
                    ElseIf Val(pRow.Item("畑面積").ToString) > 0 OrElse pRow.Item("登記簿地目名") = "畑" Then
                        転用面積_畑 += pRow.Item("実面積")
                    End If
                End If

                転用農地Count += 1
            End If
        Next

        If 転用農地Count > 0 Then
            転用許可地 &= String.Format(" 外{0}筆", 転用農地Count)
        End If

        Dim sExcel As String = HimTools2012.TextAdapter.LoadTextFile(mvarFileName)

        Select Case mvar申請.法令
            Case enum法令.農地法4条
                申請者氏名 = mvar申請.Row.Body("氏名A").ToString
                申請者住所 = mvar申請.Row.Body("住所A").ToString
                sExcel = Replace(sExcel, "{農地法}", 4)
                sExcel = Replace(sExcel, "{許可和暦年}", "指令" & 和暦Format(許可年月日, "yy"))
                sExcel = Replace(sExcel, "{本文}", "許可がされている土地の工事進捗状況を下記のとおり報告します")
            Case enum法令.農地法5条所有権, enum法令.農地法5条貸借
                申請者氏名 = mvar申請.Row.Body("氏名B").ToString
                申請者住所 = mvar申請.Row.Body("住所B").ToString
                sExcel = Replace(sExcel, "{農地法}", 5)
                sExcel = Replace(sExcel, "{許可和暦年}", "指令" & 和暦Format(許可年月日, "yy"))
                sExcel = Replace(sExcel, "{本文}", "許可がされている土地の工事進捗状況を下記のとおり報告します")
            Case enum法令.事業計画変更
                申請者氏名 = mvar申請.Row.Body("氏名C").ToString
                申請者住所 = mvar申請.Row.Body("住所C").ToString
                sExcel = Replace(sExcel, "{農地法}", 5)
                sExcel = Replace(sExcel, "{許可和暦年}", "")
                sExcel = Replace(sExcel, "{本文}", "許可（令和　年　月　日付け指令第　号－　号）に係る事業計画の変更承認について、工事進捗状況を下記のとおり報告します")
        End Select

        sExcel = Replace(sExcel, "{発行年月日}", 和暦Format(発行年月日))
        sExcel = Replace(sExcel, "{転用事業者住所}", 申請者住所)
        sExcel = Replace(sExcel, "{転用事業者氏名}", 申請者氏名)
        sExcel = Replace(sExcel, "{許可年月日}", 和暦Format(許可年月日))
        sExcel = Replace(sExcel, "{許可年}", 許可年月日.Year)



        sExcel = Replace(sExcel, "{許可番号}", 許可番号)
        sExcel = Replace(sExcel, "{土地の所在}", 転用許可地)

        sExcel = Replace(sExcel, "{転用目的}", 転用目的)

        sExcel = Replace(sExcel, "{田面積計}", HimTools2012.NumericFunctions.NumToString(転用面積_田))
        sExcel = Replace(sExcel, "{畑面積計}", HimTools2012.NumericFunctions.NumToString(転用面積_畑))
        sExcel = Replace(sExcel, "{採草牧草地面積計}", 0)
        sExcel = Replace(sExcel, "{面積計}", HimTools2012.NumericFunctions.NumToString(転用面積_田 + 転用面積_畑))

        SavePath = SysAD.OutputFolder & String.Format("\工事進捗状況報告書\{0}{1:00}", 発行年月日.Year, 発行年月日.Month)
        HimTools2012.FileManager.CheckAndCleateDirectory(SavePath)
        SaveFileName = "\工事進捗状況報告" & mvar申請.名称 & ".xml"
        SavePath &= SaveFileName

        HimTools2012.TextAdapter.SaveTextFile(SavePath, sExcel)
    End Sub
End Class

Public MustInherit Class C印刷処理
    Inherits HimTools2012.clsAccessor

    Protected SavePath As String = ""
    Protected SaveFileName As String = ""

    Public Sub ShowDialog()
        Me.Dialog.StartProc(True, True)

        If Me.Dialog._objException Is Nothing = False Then
            If Me.Dialog._objException.Message = "Cancel" Then
                MsgBox("処理を中止しました。　", , "処理中止")
            Else
                'Throw objDlg._objException
            End If
        Else
            SysAD.ShowFolder(SavePath)
            Using pExcel As New HimTools2012.Excel.Automation.ExcelAutomation
                pExcel.ShowPreview(SavePath)
            End Using
        End If

    End Sub

End Class