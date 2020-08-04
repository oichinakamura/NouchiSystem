Public Class COutPutCSV農地利用状況調査
    Inherits HimTools2012.clsAccessor

    Public TBL利用状況調査農地 As New DataTable

    Public Sub New()
        Me.Start(True, True)
    End Sub

    Public Overrides Sub Execute()
        Try
            Message = "農地情報読み込み中..."
            TBL利用状況調査農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_農地.土地所在, V_農地.地番, V_地目.名称 AS 登記地目, V_地目.名称 AS 現況地目, V_農地.登記簿面積, V_農地.実面積, V_農地.自小作別, V_小作法令.名称 AS 適用法令, [D:個人Info].[フリガナ] AS 所有者フリガナ, [D:個人Info].氏名 AS 所有者, [D:個人Info].郵便番号 AS 所有者郵便番号, [D:個人Info].住所 AS 所有者住所, [D:個人Info_1].[フリガナ] AS 管理者フリガナ, [D:個人Info_1].郵便番号 AS 管理者郵便番号, [D:個人Info_1].氏名 AS 管理者氏名, [D:個人Info_1].住所 AS 管理者住所, [D:個人Info_2].[フリガナ] AS 借受者フリガナ, [D:個人Info_2].氏名 AS 借受者, [D:個人Info_2].郵便番号 AS 借受者郵便番号, [D:個人Info_2].住所 AS 借受者住所, V_農地.小作開始年月日 AS 貸借開始日, V_農地.小作終了年月日 AS 貸借終了日, V_農地.小作料S AS 小作料S, V_農地.備考 " & _
                                                                      "FROM (((((V_農地 LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON V_農地.現況地目 = V_現況地目.ID) LEFT JOIN V_小作法令 ON V_農地.小作地適用法 = V_小作法令.ID) LEFT JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.管理者ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON V_農地.借受人ID = [D:個人Info_2].ID " & _
                                                                      "ORDER BY V_農地.大字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])), IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),0);")

            Set対象農地()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Set対象農地()
        Dim sCSV As New StringBEx("")
        Dim h取込用農地 As String() = GetHeader利用状況()
        Dim pLineHeader As New StringBEx("土地所在")
        Dim pView As DataView = New DataView(TBL利用状況調査農地, "", "", DataViewRowState.CurrentRows)

        For n As Integer = 0 To UBound(h取込用農地)
            pLineHeader.mvarBody.Append("," & h取込用農地(n))
        Next
        sCSV.Body.AppendLine(pLineHeader.Body.ToString)

        Me.Maximum = pView.Count
        Me.Value = 0

        For Each pRow As DataRowView In pView
            Me.Value += 1
            Message = "データ出力中(" & Me.Value & "/" & pView.Count & ")..."

            Dim pLineRow As New StringBEx(pRow.Item("土地所在").ToString)
            With pLineRow.mvarBody
                If InStr(pRow.Item("地番"), "-") > 0 Then : .Append("," & Left(pRow.Item("地番"), InStr(pRow.Item("地番"), "-") - 1))
                Else : .Append("," & pRow.Item("地番"))
                End If

                If InStr(pRow.Item("地番"), "-") > 0 Then : .Append("," & Mid(pRow.Item("地番"), InStr(pRow.Item("地番"), "-") + 1))
                Else : .Append("," & "")
                End If

                .Append("," & pRow.Item("登記地目").ToString)
                .Append("," & pRow.Item("現況地目").ToString)
                .Append("," & pRow.Item("登記簿面積").ToString)
                .Append("," & pRow.Item("実面積").ToString)
                .Append("," & IIf(pRow.Item("自小作別") <> 0, "小作", "-"))
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("適用法令"), ""))
                .Append("," & pRow.Item("所有者フリガナ").ToString)
                .Append("," & pRow.Item("所有者").ToString)
                .Append("," & pRow.Item("所有者郵便番号").ToString)
                .Append("," & pRow.Item("所有者住所").ToString)
                .Append("," & pRow.Item("管理者フリガナ").ToString)
                .Append("," & pRow.Item("管理者郵便番号").ToString)
                .Append("," & pRow.Item("管理者氏名").ToString)
                .Append("," & pRow.Item("管理者住所").ToString)
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("借受者フリガナ"), ""))
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("借受者"), ""))
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("借受者郵便番号"), ""))
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("借受者住所"), ""))
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("貸借開始日"), ""))
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("貸借終了日"), ""))
                .Append("," & IIf(pRow.Item("自小作別") <> 0, pRow.Item("小作料S"), ""))
                .Append("," & pRow.Item("備考").ToString)

                sCSV.Body.AppendLine(pLineRow.Body.ToString)
            End With
        Next

        名前を付けて保存(sCSV, String.Format("農地利用状況調査（{0}）", Format(Now, "yy")), True, True)
    End Sub

    Private Function GetHeader利用状況()
        Dim sResult As String() = {"本番", "枝番", "登記地目", "現況地目", "登記簿面積", "実面積", "借受状態", "適用法令", _
                                   "所有者フリガナ", "所有者", "所有者郵便番号", "所有者住所", _
                                   "管理者フリガナ", "管理者郵便番号", "管理者氏名", "管理者住所", _
                                   "借受者フリガナ", "借受者", "借受者郵便番号", "借受者住所", _
                                   "貸借開始日", "貸借終了日", "小作料S", "備考"}
        Return sResult
    End Function

    Private sPath As String = ""
    Private Sub 名前を付けて保存(ByVal sCSV As StringBEx, ByVal SaveFileName As String, Optional ByVal OpenDialog As Boolean = False, Optional ByVal OpenFolder As Boolean = False)
        '/***名前を付けて保存***/
        If OpenDialog = True Then
            With New SaveFileDialog
                .FileName = String.Format("{0}.CSV", SaveFileName)
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
            ElseIf n = UBound(ArSavePath) Then : sPath = sPath & "\" & String.Format("{0}.CSV", SaveFileName)
            Else : sPath = sPath & "\" & ArSavePath(n)
            End If
        Next

        Dim CSVText As New System.IO.StreamWriter(sPath, False, System.Text.Encoding.GetEncoding(932))
        CSVText.Write(sCSV.Body.ToString)
        CSVText.Dispose()

        If OpenFolder = True Then
            MsgBox("CSVの出力が完了しました。")
            SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
        End If
    End Sub
End Class
