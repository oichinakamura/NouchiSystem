Imports HimTools2012

Public Class C標準フォーマット
    Inherits controls.CTabPageWithToolStrip

    Public DSet As DataSet
    Public TBL申請 As DataTable
    Public TBL農家 As DataTable
    Public TBL農地 As DataTable
    Public TBL転用農地 As DataTable
    Public TBL現況地目 As 現況地目変換

    Public Sub New()
        MyBase.New(True, True, "標準フォーマット出力", "標準フォーマット出力")

        TBL現況地目 = New 現況地目変換
        TBL現況地目.Init()
        TBL現況地目.PrimaryKey = New DataColumn() {TBL現況地目.Columns("現況名称")}
    End Sub

    Public Function Get農家情報(ByVal nID As Decimal, ByVal pInfo As PersonInfo) As String
        If Not TBL農家 Is Nothing AndAlso nID <> 0 Then
            Dim pRow As DataRow = TBL農家.Rows.Find(nID)
            If Not pRow Is Nothing Then
                Select Case pInfo
                    Case PersonInfo.Name : Return pRow.Item("氏名").ToString()
                    Case PersonInfo.Address : Return pRow.Item("住所").ToString()
                    Case PersonInfo.Job : Return pRow.Item("職業").ToString()
                End Select
            End If
        End If

        Return ""
    End Function

    Public Function Get農地情報(ByVal nID As Decimal, ByVal pInfo As LandInfo) As String
        If Not TBL農地 Is Nothing And nID <> 0 Then
            Dim pRow As DataRow = TBL農地.Rows.Find(nID)
            If Not pRow Is Nothing Then
                Select Case pInfo
                    Case LandInfo.大字ID : Return pRow.Item("大字ID").ToString()
                    Case LandInfo.大字名 : Return pRow.Item("大字名").ToString()
                    Case LandInfo.小字ID : Return pRow.Item("小字ID").ToString()
                    Case LandInfo.小字名 : Return pRow.Item("小字名").ToString()
                    Case LandInfo.本番区分 : Return pRow.Item("本番区分").ToString()
                    Case LandInfo.本番 : Return pRow.Item("本番").ToString()
                    Case LandInfo.枝番区分 : Return pRow.Item("枝番区分").ToString()
                    Case LandInfo.枝番 : Return pRow.Item("枝番").ToString()
                    Case LandInfo.孫番区分 : Return pRow.Item("孫番区分").ToString()
                    Case LandInfo.孫番 : Return pRow.Item("孫番").ToString()
                    Case LandInfo.一部現況 : Return pRow.Item("一部現況").ToString()
                    Case LandInfo.現況地目 : Return pRow.Item("現況地目").ToString()
                    Case LandInfo.現況地目名 : Return pRow.Item("現況地目名").ToString()
                End Select
            End If
        End If

        Return ""
    End Function

    Public Sub ColumnCheck(ByRef pTBL As DataTable, ByVal pColName As String, ByVal pColType As Type)
        If Not pTBL.Columns.Contains(pColName) Then
            pTBL.Columns.Add(pColName, pColType)
        End If
    End Sub

    Public Function CnvDate(ByVal pData As Object) As String
        If IsDate(pData) AndAlso Not pData = "0:00:00" Then
            Return Format(CDate(pData), "yyyy/MM/dd")
        Else
            Return ""
        End If
    End Function

    Public Function CnvBool(ByVal pData As Object) As Boolean
        If Not IsDBNull(pData) Then
            Return pData
        Else
            Return False
        End If
    End Function

    Public Sub Conv地番(ByRef pRow As DataRow)
        Dim pAddress As String = Replace(pRow.Item("地番").ToString, "の", "")
        Dim s本番区分 As String = "" : Dim s本番 As String = ""
        Dim s枝番区分 As String = "" : Dim s枝番 As String = ""
        Dim s孫番区分 As String = "" : Dim s孫番 As String = ""

        pAddress = StrConv(pAddress, vbNarrow)

        If InStr(pAddress, "-") > 0 Then        '地番が"-"を含むかどうか
            s本番 = Val(StringF.Left(pAddress, InStr(pAddress, "-") - 1))
            Dim s分岐1 As String = Mid(pAddress, InStr(pAddress, "-") + 1)

            If InStr(s分岐1, "-") > 0 Then        '枝番以降が"-"を含むかどうか
                If Char.IsNumber(s分岐1, 0) Then
                    s枝番 = Val(StringF.Left(s分岐1, InStr(s分岐1, "-") - 1))
                    Dim s分岐2 As String = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                    If InStr(s分岐2, "-") > 0 Then
                        If Char.IsNumber(s分岐2, 0) Then
                            s孫番 = Val(StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                            s孫番区分 = StrConv(Mid(s分岐2, InStr(s分岐2, "-") + 1), VbStrConv.Wide)
                            '終了
                        Else
                            s枝番区分 = StrConv(StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                            Dim s分岐3 As String = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                s孫番 = Val(StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
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
                    s本番区分 = StrConv(StringF.Left(s分岐1, InStr(s分岐1, "-") - 1), VbStrConv.Wide)
                    Dim s分岐2 As String = Mid(s分岐1, InStr(s分岐1, "-") + 1)

                    If InStr(s分岐2, "-") > 0 Then
                        If Char.IsNumber(s分岐2, 0) Then
                            s枝番 = Val(StringF.Left(s分岐2, InStr(s分岐2, "-") - 1))
                            Dim s分岐3 As String = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                If Char.IsNumber(s分岐3, 0) Then
                                    s孫番 = Val(StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
                                    s孫番区分 = StrConv(Mid(s分岐3, InStr(s分岐3, "-") + 1), VbStrConv.Wide)
                                    '終了
                                Else
                                    s枝番区分 = StrConv(StringF.Left(s分岐3, InStr(s分岐3, "-") - 1), VbStrConv.Wide)
                                    Dim s分岐4 = Mid(s分岐3, InStr(s分岐3, "-") + 1)

                                    If InStr(s分岐4, "-") > 0 Then
                                        s孫番 = Val(StringF.Left(s分岐4, InStr(s分岐4, "-") - 1))
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
                            s枝番区分 = StrConv(StringF.Left(s分岐2, InStr(s分岐2, "-") - 1), VbStrConv.Wide)
                            Dim s分岐3 = Mid(s分岐2, InStr(s分岐2, "-") + 1)

                            If InStr(s分岐3, "-") > 0 Then
                                s孫番 = Val(StringF.Left(s分岐3, InStr(s分岐3, "-") - 1))
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

    Public Function Conv権利種類(ByVal pRow As DataRow) As Integer
        Dim pValue As Integer = Val(pRow.Item("権利種類").ToString)

        If IsNumeric(pValue) Then
            If pValue >= 100 Then
                Return Val(StringF.Right(pValue, 2))
            Else
                Select Case pValue
                    Case 0 : Return 0
                    Case 1 : Return 5
                    Case 2 : Return 4
                    Case 3 : Return 10
                    Case 4 : Return 1
                    Case 5 : Return 2
                    Case 6 : Return 3
                    Case 7 : Return 6
                    Case 8 : Return 10
                    Case 9 : Return 10
                End Select
            End If
        Else
            pValue = 0
        End If

        Return pValue
    End Function

    Public Function Conv小作料(ByVal pRow As DataRow, ByVal type As FarmRent) As String
        Select Case type
            Case FarmRent.小作料
                If InStr(pRow.Item("小作料単位").ToString(), "円") > 0 Then
                    Return Val(pRow.Item("小作料").ToString())
                End If
            Case FarmRent.物納
                If InStr(pRow.Item("小作料単位").ToString(), "円") = 0 Then
                    Return pRow.Item("小作料").ToString() & pRow.Item("小作料単位").ToString()
                End If
        End Select

        Return ""
    End Function

    Public Function Conv期間(ByVal pRow As DataRow, ByVal type As Time) As String
        Dim s期間 As String = pRow.Item("期間").ToString()

        Select Case type
            Case Time.年
                If InStr(s期間, "年") > 0 Then
                    Return StringF.Left(s期間, InStr(s期間, "年") - 1)
                End If
            Case Time.月
                If InStr(s期間, "年") > 0 And InStr(s期間, "ヵ月") > 0 Then
                    Dim ar As Object = Split(s期間, "年")
                    Return Replace(ar(1), "ヶ月", "")
                End If
        End Select

        Return ""
    End Function

    Public Function Conv解約形態(ByVal pRow As DataRow) As Integer
        Select Case pRow.Item("申請理由A").ToString
            Case "期間満了" : Return 2
            Case ""
                If CnvDate(pRow.Item("許可年月日")) >= CnvDate(pRow.Item("終期")) Then
                    Return 2
                Else
                    Return 1
                End If
            Case Else : Return 1
        End Select

        If pRow.Item("申請理由A").ToString = "期間満了" Then
            Return 2
        Else
            If pRow.Item("申請理由A").ToString = "" Then

            Else
                Return 2
            End If
        End If
    End Function

    Public Function Conv判定地目(ByVal sVal As String) As Integer
        Dim findRow As DataRow = TBL現況地目.Rows.Find(sVal)

        If Not findRow Is Nothing Then
            Return findRow.Item("現況ID")
        Else
            Return 0
        End If
    End Function

    Public Function Cnv農地ID(ByVal pData As Object) As String
        Try
            If Not pData Is Nothing AndAlso Not IsDBNull(pData) Then
                If pData = 0 Then
                    Return pData.ToString("0")
                ElseIf pData < 0 Then
                    If Len(pData.ToString) = 9 Then
                        Return Math.Abs(Val(pData.ToString))
                    ElseIf Len(pData.ToString) = 8 Then
                        Return Math.Abs(Val(pData.ToString)).ToString("90000000")
                    Else
                        Return Math.Abs(Val(pData.ToString)).ToString("99000000")
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

    Public Function 名前を付けて保存(ByVal sFileName As String, ByVal type As OutputType, Optional ByVal OptName As String = "") As String
        Dim SaveFileName As String = IIf(OptName <> "", String.Format("{0}_{1}", sFileName, OptName), sFileName)
        Dim sFilter As String = ""

        Select Case type
            Case OutputType.xlsx
                sFilter = "エクセルファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*"
            Case OutputType.xlsm
                sFilter = "エクセルファイル(*.xlsm)|*.xlsm|すべてのファイル(*.*)|*.*"
            Case OutputType.csv
                sFilter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"
        End Select

        Using sfdlg As New SaveFileDialog()
            sfdlg.FileName = $"{SaveFileName}.{type.ToString()}"
            sfdlg.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            sfdlg.Filter = sFilter

            If sfdlg.ShowDialog() = DialogResult.OK Then
                Return sfdlg.FileName
            End If

        End Using

        Return ""
    End Function

    Public Sub ReleaseObject(ByRef obj As Object)
        Try
            Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
            obj = Nothing
        Catch ex As IO.FileNotFoundException
            obj = Nothing
            Console.WriteLine("オブジェクトを解放できない" + ex.ToString())
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Enum OutputType
        xlsx = 1
        xlsm = 2
        csv = 3
    End Enum

    Public Enum PersonInfo
        Name = 1
        Address = 2
        Job = 3
    End Enum

    Public Enum LandInfo
        大字ID = 1
        大字名 = 2
        小字ID = 3
        小字名 = 4
        本番区分 = 5
        本番 = 6
        枝番区分 = 7
        枝番 = 8
        孫番区分 = 9
        孫番 = 10
        一部現況 = 11
        現況地目 = 12
        現況地目名 = 13
    End Enum

    Public Enum FarmRent
        小作料 = 1
        物納 = 2
    End Enum

    Public Enum Time
        年 = 1
        月 = 2
    End Enum
End Class