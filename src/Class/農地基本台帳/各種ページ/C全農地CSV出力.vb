Public Class C全農地CSV出力
    Inherits HimTools2012.clsAccessor

    Public TBL個人 As New DataTable
    Public TBL農地 As New DataTable
    Public TBL転用農地 As New DataTable

    Public Sub New()
        Me.Start(True, True)
    End Sub

    Public Overrides Sub Execute()
        Try
            Message = "個人情報読み込み中..."
            TBL個人 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].郵便番号, [D:個人Info].住所 FROM [D:個人Info];")
            TBL個人.PrimaryKey = {TBL個人.Columns("ID")}

            Message = "農地情報読み込み中..."
            TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_大字.大字, [D:農地Info].地番, V_地目.名称 AS 登地目, V_現況地目.名称 AS 現地目, [D:農地Info].登記簿面積, [D:農地Info].自小作別, IIf([自小作別]=0,'',[D:農地Info].[小作地適用法]) AS 小作地適用法, IIf([自小作別]=0,'',[D:農地Info].[小作形態]) AS 小作形態, [D:農地Info].所有者ID, IIf([自小作別]=0,'',[D:農地Info].[借受人ID]) AS 借受人ID, IIf([自小作別]=0,'',[D:農地Info].[小作開始年月日]) AS 小作開始年月日, IIf([自小作別]=0,'',[D:農地Info].[小作終了年月日]) AS 小作終了年月日, IIf([自小作別]=0,'',[D:農地Info].[小作料]) AS 小作料, IIf([自小作別]=0,'',[D:農地Info].[小作料単位]) AS 小作料単位, [D:農地Info].備考, [D:農地Info].経由農業生産法人ID, [D:農地Info].利用配分計画始期日, [D:農地Info].利用配分計画終期日, [D:農地Info].農業振興地域, [D:農地Info].農振法区分, [D:農地Info].都市計画法, [D:農地Info].都市計画法区分, [大字ID] & '-' & [地番] AS [Key] " &
                                                          "FROM (([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID " &
                                                          "WHERE (((V_大字.大字) Is Not Null)) " &
                                                          "ORDER BY [D:農地Info].大字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])), IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),'');")
            Message = "農地データ出力中..."
            Set出力用農地(TBL農地, "農地")

            Message = "転用農地情報読み込み中..."
            TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT V_大字.大字, [D_転用農地].地番, V_地目.名称 AS 登地目, V_現況地目.名称 AS 現地目, [D_転用農地].登記簿面積, [D_転用農地].自小作別, IIf([自小作別]=0,'',[D_転用農地].[小作地適用法]) AS 小作地適用法, IIf([自小作別]=0,'',[D_転用農地].[小作形態]) AS 小作形態, [D_転用農地].所有者ID, IIf([自小作別]=0,'',[D_転用農地].[借受人ID]) AS 借受人ID, IIf([自小作別]=0,'',[D_転用農地].[小作開始年月日]) AS 小作開始年月日, IIf([自小作別]=0,'',[D_転用農地].[小作終了年月日]) AS 小作終了年月日, IIf([自小作別]=0,'',[D_転用農地].[小作料]) AS 小作料, IIf([自小作別]=0,'',[D_転用農地].[小作料単位]) AS 小作料単位, [D_転用農地].備考, [D_転用農地].経由農業生産法人ID, [D_転用農地].利用配分計画始期日, [D_転用農地].利用配分計画終期日, [D_転用農地].農業振興地域, [D_転用農地].農振法区分, [D_転用農地].都市計画法, [D_転用農地].都市計画法区分, [大字ID] & '-' & [地番] AS [Key] " &
                                                              "FROM (([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_地目 ON [D_転用農地].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D_転用農地].現況地目 = V_現況地目.ID " &
                                                              "WHERE (((V_大字.大字) Is Not Null)) " &
                                                              "ORDER BY [D_転用農地].大字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])), IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),'');")
            Message = "転用農地データ出力中..."
            Set出力用農地(TBL転用農地, "転用農地")

            MsgBox("CSVの出力が完了しました。")
            SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Set出力用農地(ByVal pTBL As DataTable, ByVal sName As String)
        Dim sCSV As New SBuilderEx("")
        Dim pLineHeader As New SBuilderEx("数")
        Dim h取込用農地 As String() = GetHeader農地()

        For n As Integer = 0 To UBound(h取込用農地)
            pLineHeader.mvarBody.Append("," & h取込用農地(n))
        Next
        sCSV.Body.AppendLine(pLineHeader.Body.ToString)

        Me.Maximum = pTBL.Rows.Count
        Me.Value = 0

        For Each pRow As DataRow In pTBL.Rows
            Me.Value += 1
            Message = "農地データ出力中(" & Me.Value & "/" & Me.Maximum & ")..."

            Dim pLineRow As New SBuilderEx("1") '筆
            With pLineRow
                .mvarBody.Append(",") '地域
                .CnvData(pRow.Item("大字"), EnumCnv.全角) '大字
                .CnvData(Conv地番(pRow.Item("地番"), Enum地番.本番), EnumCnv.半角) '本番
                .CnvData(Conv地番(pRow.Item("地番"), Enum地番.枝番), EnumCnv.半角) '枝番
                .CnvData(pRow.Item("登地目"), EnumCnv.全角) '登記地目
                .CnvData(pRow.Item("現地目"), EnumCnv.全角) '現況地目
                .CnvData(pRow.Item("登記簿面積"), EnumCnv.面積) '登記簿面積
                .CnvData(Conv自小作別(pRow.Item("自小作別")), EnumCnv.全角) '借受状況
                .CnvData(Set借受情報(pRow, EnumLoan.適用法令), EnumCnv.全角) '適用法令
                .CnvData(Set借受情報(pRow, EnumLoan.小作形態), EnumCnv.全角) '小作形態
                .CnvData(Set農家情報(pRow.Item("所有者ID"), EnumOwner.フリガナ), EnumCnv.半角) '所有者フリガナ
                .CnvData(Set農家情報(pRow.Item("所有者ID"), EnumOwner.氏名), EnumCnv.全角) '所有者
                .CnvData(Set農家情報(pRow.Item("所有者ID"), EnumOwner.郵便番号), EnumCnv.半角) '所有者郵便番号
                .CnvData(Set農家情報(pRow.Item("所有者ID"), EnumOwner.住所), EnumCnv.全角) '所有者住所
                .CnvData(Set借受者情報(pRow, pRow.Item("借受人ID"), EnumOwner.フリガナ), EnumCnv.半角) '借受者フリガナ
                .CnvData(Set借受者情報(pRow, pRow.Item("借受人ID"), EnumOwner.氏名), EnumCnv.全角) '借受者
                .CnvData(Set借受者情報(pRow, pRow.Item("借受人ID"), EnumOwner.郵便番号), EnumCnv.半角) '借受者郵便番号
                .CnvData(Set借受者情報(pRow, pRow.Item("借受人ID"), EnumOwner.住所), EnumCnv.全角) '借受者住所
                .CnvData(Set借受情報(pRow, EnumLoan.貸借開始日), EnumCnv.日付) '貸借開始日
                .CnvData(Set借受情報(pRow, EnumLoan.貸借終了日), EnumCnv.日付) '貸借終了日
                .CnvData(Set借受情報(pRow, EnumLoan.小作料), EnumCnv.半角) '小作料
                .mvarBody.Append(",") '賃借料複数筆
                .CnvData(pRow.Item("備考"), EnumCnv.全角) '備考
                .CnvData(Set借受情報(pRow, EnumLoan.経由法人名), EnumCnv.全角) '経由法人名
                .CnvData(Set借受情報(pRow, EnumLoan.経由法人住所), EnumCnv.全角) '経由法人住所
                .CnvData(Set借受情報(pRow, EnumLoan.経由貸借開始日), EnumCnv.日付) '経由貸借開始日
                .CnvData(Set借受情報(pRow, EnumLoan.経由貸借終了日), EnumCnv.日付) '経由貸借終了日
                .CnvData(Conv農振法区分(pRow), EnumCnv.全角) '農業振興地域
                .CnvData(Conv都市計画法区分(pRow), EnumCnv.全角) '都市計画区域
                .CnvData(pRow.Item("Key"), EnumCnv.半角) 'Key

                sCSV.Body.AppendLine(.Body.ToString)
            End With
        Next

        名前を付けて保存(sCSV, String.Format("{0}_全農地CSV出力({1})", Format(Now, "yyyyMMdd"), sName), True)
    End Sub

    Private Function GetHeader農地()
        Dim sResult As String() = {"地域", "大字", "本番", "枝番", "登記地目", "現況地目", "登記簿面積", "借受状況", "適用法令", "小作形態", "所有者フリガナ", "所有者", "所有者郵便番号", "所有者住所",
                                   "借受者フリガナ", "借受者", "借受者郵便番号", "借受者住所", "貸借開始日", "貸借終了日", "小作料", "賃借料複数筆", "備考",
                                   "経由法人名", "経由法人住所", "経由貸借開始日", "経由貸借終了日", "農業振興地域", "都市計画区域", "Key"}

        Return sResult
    End Function

    Private Function Conv地番(ByVal pValue As Object, ByVal pOption As Enum地番) As String
        Dim Int本番 As String = ""
        Dim Int枝番 As String = ""
        If Not IsDBNull(pValue) AndAlso Not pValue Is Nothing Then
            If InStr(pValue, "-") > 0 Then
                Int本番 = Left(pValue, InStr(pValue, "-") - 1)
                Int枝番 = Mid(pValue, InStr(pValue, "-") + 1)

                If InStr(Int枝番, "-") > 0 Then
                    Int枝番 = ""
                End If
            Else
                Int本番 = pValue
            End If

            Select Case pOption
                Case Enum地番.本番 : Return Int本番
                Case Enum地番.枝番 : Return Int枝番
                Case Else : Return ""
            End Select
        Else
            Return ""
        End If
    End Function

    Private Function Conv自小作別(ByVal pValue As Object) As String
        Select Case Val(pValue.ToString)
            Case -8, -1 : Return "ヤミ小作"
            Case 1 : Return "小作"
            Case 2 : Return "農年"
            Case Else : Return "自作"
        End Select
    End Function

    Private Function Conv適用法令(ByVal pValue As Object) As String
        Select Case Val(pValue.ToString)
            Case 0, 100 : Return "-"
            Case 1, 101, 102 : Return "農地法"
            Case 2, 103 : Return "基盤強化促進法"
            Case 3, 104 : Return "特定農地貸付法"
            Case Else : Return "その他"
        End Select
    End Function

    Private Function Conv小作形態(ByVal pValue As Object) As String
        Select Case Val(pValue.ToString)
            Case 0 : Return "-"
            Case 1 : Return "賃貸借"
            Case 2 : Return "使用貸借"
            Case 4 : Return "地上権"
            Case 6 : Return "質権"
            Case 7 : Return "期間借地"
            Case 8 : Return "残存小作地"
            Case 9 : Return "使用賃借"
            Case Else : Return "その他"
        End Select
    End Function

    Private Function Set農家情報(ByVal pID As Object, ByVal pOption As EnumOwner) As String
        Dim pRow As DataRow = TBL個人.Rows.Find(pID)

        If Not pRow Is Nothing Then
            Select Case pOption
                Case EnumOwner.氏名 : Return pRow.Item("氏名").ToString
                Case EnumOwner.フリガナ : Return pRow.Item("フリガナ").ToString
                Case EnumOwner.郵便番号 : Return pRow.Item("郵便番号").ToString
                Case EnumOwner.住所 : Return pRow.Item("住所").ToString
                Case Else : Return ""
            End Select
        Else
            Return ""
        End If
    End Function

    Private Function Set借受者情報(ByVal pRow As DataRow, ByVal pID As Object, ByVal pOption As EnumOwner) As String
        If Not Val(pRow.Item("自小作別").ToString) = 0 AndAlso Not IsDBNull(pID) Then
            Return Set農家情報(pID, pOption)
        Else
            Return ""
        End If
    End Function

    Private Function Set借受情報(ByVal pRow As DataRow, ByVal pOption As EnumLoan) As String
        If Not Val(pRow.Item("自小作別").ToString) = 0 Then
            Select Case pOption
                Case EnumLoan.適用法令 : Return Conv適用法令(pRow.Item("小作地適用法"))
                Case EnumLoan.貸借開始日 : Return pRow.Item("小作開始年月日").ToString
                Case EnumLoan.貸借終了日 : Return pRow.Item("小作終了年月日").ToString
                Case EnumLoan.小作料 : Return pRow.Item("小作料").ToString & pRow.Item("小作料単位").ToString
                Case EnumLoan.経由法人名 : Return Set農家情報(pRow.Item("経由農業生産法人ID"), EnumOwner.氏名)
                Case EnumLoan.経由法人住所 : Return Set農家情報(pRow.Item("経由農業生産法人ID"), EnumOwner.住所)
                Case EnumLoan.経由貸借開始日 : Return pRow.Item("利用配分計画始期日").ToString
                Case EnumLoan.経由貸借終了日 : Return pRow.Item("利用配分計画終期日").ToString
                Case EnumLoan.小作形態 : Return Conv小作形態(pRow.Item("小作形態").ToString)
                Case Else : Return ""
            End Select
        Else
            Return ""
        End If
    End Function

    Private Function Conv農振法区分(ByVal pRow As DataRow) As String
        Select Case Val(pRow.Item("農振法区分").ToString)
            Case 1 : Return "農用地区域"
            Case 2 : Return "農振地域"
            Case 3 : Return "農振地域外"
            Case Else
                Select Case Val(pRow.Item("農業振興地域").ToString)
                    Case 0 : Return "農振地域"
                    Case 1 : Return "農用地区域"
                    Case 2 : Return "農振地域外"
                    Case Else : Return "－"
                End Select
        End Select
    End Function

    Private Function Conv都市計画法区分(ByVal pRow As DataRow) As String
        Select Case Val(pRow.Item("都市計画法区分").ToString)
            Case 1 : Return "市街化区域"
            Case 2 : Return "市街化調整区域"
            Case 3 : Return "用途地域"
            Case 4 : Return "都市計画区域外"
            Case 5 : Return "その他"
            Case Else
                Select Case Val(pRow.Item("都市計画法").ToString)
                    Case 0 : Return "都市計画区域外"
                    Case 1 : Return "市街化区域"
                    Case 2 : Return "用途地域"
                    Case 3 : Return "市街化調整区域"
                    Case 4 : Return "市街化区域"
                    Case 5 : Return "用途地域"
                    Case Else : Return "－"
                End Select
        End Select
    End Function

    Private sPath As String = ""
    Private Sub 名前を付けて保存(ByVal sCSV As SBuilderEx, ByVal SaveFileName As String, Optional ByVal OpenDialog As Boolean = False)
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

        Dim CSVText As New System.IO.StreamWriter(sPath, False, System.Text.Encoding.GetEncoding(932))
        CSVText.Write(sCSV.Body.ToString)
        CSVText.Dispose()
    End Sub
End Class

Public Class SBuilderEx
    Public mvarBody As System.Text.StringBuilder

    Public Sub New(ByVal s初期値 As String)
        mvarBody = New System.Text.StringBuilder(s初期値)
    End Sub
    Public ReadOnly Property Body As System.Text.StringBuilder
        Get
            Return mvarBody
        End Get
    End Property

    Public Sub CnvData(ByVal pData As Object, ByVal bOption As EnumCnv, Optional ByVal sCode As Integer = 0)
        If pData Is Nothing Then : mvarBody.Append(",")
        ElseIf IsDBNull(pData) Then
            Select Case bOption
                Case EnumCnv.全角, EnumCnv.半角, EnumCnv.日付 : mvarBody.Append(",")
                Case Else : mvarBody.Append("," & 0)
            End Select
        Else
            If Len(pData) > 0 Then
                pData = Replace(Replace(pData, Chr(13), ""), Chr(10), "")
            End If

            Select Case bOption
                Case EnumCnv.全角 : mvarBody.Append("," & StrConv(pData.ToString, vbWide))
                Case EnumCnv.半角 : mvarBody.Append("," & StrConv(pData.ToString, vbNarrow))
                Case EnumCnv.日付
                    If IsDate(pData) Then : mvarBody.Append("," & Format(CDate(pData), "yyyy/MM/dd"))
                    Else : mvarBody.Append(",")
                    End If
                Case EnumCnv.面積  'ここで小数点第２位まで
                    pData = Math.Round(Val(pData), 2)
                    mvarBody.Append("," & pData.ToString)
                Case Else : mvarBody.Append("," & pData.ToString)
            End Select
        End If
    End Sub
End Class

Public Enum EnumCnv
    全角 = 1
    半角 = 2
    日付 = 3
    面積 = 4
End Enum

Public Enum Enum地番
    本番 = 1
    枝番 = 2
End Enum

Public Enum EnumOwner
    氏名 = 1
    フリガナ = 2
    郵便番号 = 3
    住所 = 4
End Enum

Public Enum EnumLoan
    適用法令 = 1
    貸借開始日 = 2
    貸借終了日 = 3
    小作料 = 4
    経由法人名 = 5
    経由法人住所 = 6
    経由貸借開始日 = 7
    経由貸借終了日 = 8
    小作形態 = 9
End Enum
