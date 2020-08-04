Imports System.CodeDom.Compiler
Imports System.Reflection
Imports System.Text

''' <summary>
''' 宗像市(\\10.1.0.45\HIM\システム配信\宗像市\)
''' </summary>
''' <remarks></remarks>
Public Class C宗像市
    Inherits C市町村別

    Public Sub New()
        MyBase.New("宗像市")
    End Sub

    Public Overrides Function Get選挙世帯一覧() As System.Data.DataTable
        Return Nothing
    End Function
    Public Overrides ReadOnly Property 旧農振都市計画使用 As Boolean
        Get
            Return True
        End Get
    End Property
    Public Overrides Sub InitLocalData()
        With New dlgLoginForm()

            If Not .ShowDialog() = Windows.Forms.DialogResult.OK Then
                Try

                    End
                Catch ex As Exception

                End Try
            Else
                With SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_削除農地] WHERE [ID]=0")
                    If Not .Columns.Contains("調査土地利用区域地目") Then
                        SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [D_削除農地] ADD [調査土地利用区域地目] LONG;")
                    End If
                End With
            End If
        End With
    End Sub

    Public Overrides Sub InitMenu(pMain As HimTools2012.SystemWindows.CMainPageSK)
        With pMain
            .ListView.Clear()

            .ListView.ItemAdd("農家検索", "農地・農家検索", "閲覧・検索", "閲覧・検索", AddressOf 農家一覧)
            .ListView.ItemAdd("住記読込宗像", "住記読込宗像", ImageKey.作業, "操作", AddressOf sub住記読込宗像)
            .ListView.ItemAdd("固定資産比較", "固定資産比較", "他システム連携", "操作", AddressOf 固定資産比較)
            .ListView.ItemAdd("農地台帳一括印刷", "農地台帳一括印刷", "印刷", "印刷", AddressOf ClickMenu)
            .ListView.ItemAdd("農地利用状況調査CSV出力", "農地利用状況調査CSV出力", ImageKey.作業, "操作", AddressOf subCSV農地利用状況調査)
            MyBase.InitMenu(pMain)
        End With

    End Sub

    Private Function StrTrim(sValue As Object) As String
        If IsDBNull(sValue) Then
            Return ",Null"
        ElseIf Len(Trim(Replace(sValue, "　", " "))) > 0 Then
            Return ",'" & Trim(Replace(sValue, "　", " ")) & "'"
        End If
        Return ",Null"
    End Function

    Private Sub sub住記読込宗像()
        With New OpenFileDialog
            If .ShowDialog = DialogResult.OK Then
                Dim St As String = HimTools2012.TextAdapter.LoadTextFile(.FileName)

                SysAD.DB(sLRDB).ExecuteSQL("DELETE * FROM [M_住民情報]")

                Dim Ar As String() = Split(St, vbCrLf)
                For Each SA As String In Ar
                    If Len(SA) > 10 AndAlso Len(SA) <> 278 Then
                        Dim s利用団体コード As String = Left(SA, 5) : SA = Mid(SA, 6) '
                        Dim s住民コード As String = Left(SA, 11) : SA = Mid(SA, 12)
                        Dim s性別 As String = Left(SA, 1) : SA = Mid(SA, 2) '	
                        If Val(s性別) > 0 Then
                            s性別 = "," & Val(s性別) - 1
                        Else
                            s性別 = "," & -1
                        End If
                        Dim s生年月日 As String = Left(SA, 8) : SA = Mid(SA, 9)
                        If IsDate(Mid(s生年月日, 1, 4) & "/" & Mid(s生年月日, 5, 2) & "/" & Mid(s生年月日, 7, 2)) Then

                            s生年月日 = ",#" & Mid(s生年月日, 5, 2) & "/" & Mid(s生年月日, 7, 2) & "/" & Mid(s生年月日, 1, 4) & "#"
                        Else
                            s生年月日 = ",Null"
                        End If

                        Dim sカナ氏名 As String = Left(SA, 64) : SA = Mid(SA, 65) '	64
                        Dim s氏名 As String = Left(SA, 64) : SA = Mid(SA, 65) '	64
                        Dim s住所 As String = Left(SA, 40) : SA = Mid(SA, 41) '	40
                        Dim s方書 As String = Left(SA, 40) : SA = Mid(SA, 41) '	40
                        Dim s郵便番号 As String = Left(SA, 7) : SA = Mid(SA, 8) '	7
                        Dim s市町村コード As String = Left(SA, 5) : SA = Mid(SA, 6) '	5
                        Dim s自治会コード As String = Left(SA, 9) : SA = Mid(SA, 10) '	9
                        Dim s世帯コード As String = Left(SA, 11) : SA = Mid(SA, 12) '	11
                        Dim s住民区分 As String = Left(SA, 2) : SA = Mid(SA, 3) '	2
                        Dim s続柄 As String = Left(SA, 8) : SA = Mid(SA, 9) '	8
                        Select Case s続柄
                            Case "02      ", "12      ", "01      ", "        "
                            Case "20      ", "13      ", "21      "
                            Case "51      ", "52      "
                            Case "1151    ", "0352    ", "96      ", "99      ", "5112    "
                            Case "81      ", "31      ", "54      ", "2012    ", "2020    "
                            Case "5153    ", "74      ", "0351    ", "5152    ", "11      "
                            Case "84      ", "23      ", "7112    ", "0262    "
                            Case "0252    ", "1252    ", "1152    ", "22      ", "1171    "
                            Case "3100    ", "53      ", "92      ", "5252    ", "5284    "
                            Case "5281    ", "3211    ", "1254    ", "34      ", "8111    ", "1284    "

                            Case "3Y      ", "5151    ", "61      ", "1281    ", "33      ", "2011    "
                            Case "1251    ", "63      ", "5174    ", "62      ", "5184    ", "125152  ", "71      ", "3111    "
                            Case "1220    ", "32      ", "9Z      ", "1271    ", "0251    ", "1181    ", "5181    "
                            Case "5452    ", "64      ", "515112  ", "2212    ", "2421    ", "93      ", "5154    ", "40      "
                            Case "93      "
                                Debug.Print(Chr(34) & s続柄 & """")
                                Stop
                        End Select

                        Dim s異動日 As String = Left(SA, 8) : SA = Mid(SA, 9) '	8

                        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [M_住民情報](ID,[住民区分],[フリガナ],[氏名],[住所]," &
                                                "[世帯No],[生年月日],[性別],[郵便番号],[行政区],[続柄TXT]) VALUES(" & s住民コード & "," & s住民区分 &
                                               StrTrim(sカナ氏名) & StrTrim(s氏名) & StrTrim(s住所) & "," &
                                               s世帯コード & s生年月日 & s性別 & StrTrim(s郵便番号) & "," & Val(s自治会コード) & StrTrim(s続柄) & ")")

                        If 1 Then

                        End If
                    Else
                        Stop
                    End If
                Next

                MsgBox("終了")
            End If

        End With
    End Sub

End Class
