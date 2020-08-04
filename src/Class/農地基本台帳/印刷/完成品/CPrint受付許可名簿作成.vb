Imports System.ComponentModel
Imports HimTools2012.Excel.XMLSS2003

Public Class CPrint受付許可名簿作成
    Inherits HimTools2012.clsAccessor

    Public Overrides Sub Execute()
        Dim sFolder As String = SysAD.OutputFolder & String.Format("\申請・許可名簿{0}_{1}", Now.Year, Now.Month)

        If Not IO.Directory.Exists(sFolder) Then
            IO.Directory.CreateDirectory(sFolder)
        End If

        Dim p許可範囲 As New 申請許可名簿

        Dim pFrm As New frmInputData("申請許可名簿", "許可日の範囲を入力してください", p許可範囲)
        pFrm.ShowDialog()

        If pFrm.RetObj IsNot Nothing Then
            p許可範囲 = pFrm.RetObj
            With p許可範囲
                App農地基本台帳.TBL申請.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE [状態]=2 AND " & p許可範囲.ToString))
                Dim pTBL As DataTable =
                    New DataView(App農地基本台帳.TBL申請.Body, "[状態]=2 AND " & p許可範囲.ToString, "", DataViewRowState.CurrentRows).ToTable

                sub受付許可名簿作成("農地法３条申請許可名簿", sFolder, New CPrint申請許可名簿農地法３条, New DataView(pTBL, "[法令] IN (30,31,32,33)", "受付番号,許可番号", DataViewRowState.CurrentRows), "農地法第３条申請・許可名簿.xml")
                sub受付許可名簿作成("農地法４条申請許可名簿", sFolder, New CPrint申請許可名簿農地法４条, New DataView(pTBL, "[法令] IN (40)", "受付番号,許可番号", DataViewRowState.CurrentRows), "農地法第４条申請・許可名簿.xml")
                sub受付許可名簿作成("農地法５条申請許可名簿", sFolder, New CPrint申請許可名簿農地法５条, New DataView(pTBL, "[法令] IN (50,51)", "受付番号,許可番号", DataViewRowState.CurrentRows), "農地法第５条申請・許可名簿.xml")

            End With

            SysAD.ShowFolder(sFolder)
        End If
    End Sub

    Private Sub sub受付許可名簿作成(ByVal s処理名称 As String, ByVal sDesktopFolder As String, ByVal p作成 As CPrint申請許可名簿作成共通, ByRef pView As DataView, ByVal sFile As String)
        'Try
        If pView.Count = 0 Then

        ElseIf Not IO.File.Exists(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile) Then
            MsgBox(sFile & "が見つかりません", MsgBoxStyle.Critical)
            Exit Sub
        Else
            Dim sCount As String = ""

            Do Until Not IO.File.Exists(sDesktopFolder & "\" & Replace(sFile, ".", IIf(sCount.Length > 0, "(", "") & sCount & "."))
                If sCount.Length = 0 Then
                    sCount = "1)"
                Else
                    sCount = Val(sCount) + 1 & ")"
                End If
            Loop

            Dim sXML As String = HimTools2012.TextAdapter.LoadTextFile(SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & sFile)
            Dim XMLSS As New CXMLSS2003(sXML)
            p作成.SetXML(XMLSS, pView, s処理名称, Me)

            HimTools2012.TextAdapter.SaveTextFile(sDesktopFolder & "\" & Replace(sFile, ".", IIf(sCount.Length > 0, "(", "") & sCount & "."), XMLSS.OutPut(True))
        End If
    End Sub

End Class


Public MustInherit Class CPrint申請許可名簿作成共通
    Inherits CPrint資料作成共通


    Public Sub SetXML(XMLSS As CXMLSS2003, pView As DataView, s処理名称 As String, pDataCreater As CPrint受付許可名簿作成)

        For Each pS As XMLSSWorkSheet In XMLSS.WorkBook.WorkSheets.Items.Values
            Set複数行(pS, pView, s処理名称, pDataCreater)
        Next
    End Sub

    Public Sub Set複数行(ByRef pSheet As XMLSSWorkSheet, ByVal pView As DataView, ByVal s処理名称 As String, pDataCreater As CPrint受付許可名簿作成)
        LoopRows = New XMLLoopRows(pSheet)

        pDataCreater.Maximum = pView.Count
        nLoop = -1
        For Each pRow As DataRowView In pView
            nLoop += 1
            pDataCreater.Message = s処理名称 & " 処理中 (" & nLoop + 1 & "/" & pView.Count & ")"
            pDataCreater.Value = nLoop


            Me.SetDataRow(pSheet, pRow)
        Next
    End Sub

    Public Sub 複数土地設定(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As DataRowView, ByRef p総括Item As dt総括表, Optional b再設定 As Boolean = False)
        Try
            Dim sNList As String = pRow.Item("農地リスト").ToString
            Dim s土地所在 As String = ""
            Dim s登記地目 As String = ""
            Dim s現況地目 As String = ""
            Dim n田数 As Decimal = 0
            Dim n畑数 As Decimal = 0
            Dim n他数 As Decimal = 0
            Dim n田面積 As Decimal = 0
            Dim n畑面積 As Decimal = 0
            Dim n他面積 As Decimal = 0


            Dim s面積 As String = ""
            Dim R As String = ""
            Dim Area As Decimal = 0
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
                s土地所在 = s土地所在 & R & IIf(pRowV.Item("所在").ToString.Length > 0, pRowV.Item("所在").ToString, pRowV.Item("大字").ToString & IIf(pRowV.Item("小字").ToString.Length > 0, "字", "") & pRowV.Item("小字").ToString) & pRowV.Item("地番").ToString
                s登記地目 = s登記地目 & R & pRowV.Item("登記簿地目名").ToString
                s現況地目 = s現況地目 & R & pRowV.Item("現況地目名").ToString

                s面積 = s面積 & R & Val(pRowV.Item("登記簿面積").ToString).ToString("#,##0")
                Area += Val(pRowV.Item("登記簿面積").ToString)
                R = "&#10;"
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
                    s土地所在 = s土地所在 & R & IIf(pRowV.Item("所在").ToString.Length > 0, pRowV.Item("所在").ToString, pRowV.Item("大字").ToString & IIf(pRowV.Item("小字").ToString.Length > 0, "字", "") & pRowV.Item("小字").ToString) & pRowV.Item("地番").ToString
                    s登記地目 = s登記地目 & R & pRowV.Item("登記簿地目名").ToString
                    s現況地目 = s現況地目 & R & pRowV.Item("現況地目名").ToString
                    s面積 = s面積 & R & Val(pRowV.Item("登記簿面積").ToString).ToString("#,##0")
                    R = "&#10;"
                    nCount += 1
                Next
            End If

            pSheet.ValueReplace("{筆数計}", nCount)
            pSheet.ValueReplace("{土地の所在}", s土地所在)
            pSheet.ValueReplace("{地目}", s登記地目)
            pSheet.ValueReplace("{登記地目}", s登記地目)
            pSheet.ValueReplace("{現況地目}", s現況地目)
            pSheet.ValueReplace("{田筆数計}", n田数)
            pSheet.ValueReplace("{田面積計}", IIf(n田面積 > 0, n田面積.ToString("#,##0"), ""))
            pSheet.ValueReplace("{畑筆数計}", n畑数)
            pSheet.ValueReplace("{畑面積計}", IIf(n畑面積 > 0, n畑面積.ToString("#,##0"), ""))
            pSheet.ValueReplace("{他筆数計}", n他数)
            pSheet.ValueReplace("{他面積計}", n他面積)

            pSheet.ValueReplace("{面積}", s面積)
            pSheet.ValueReplace("{面積計}", Area.ToString("#,##0"))



        Catch ex As Exception
            Stop
        End Try

    End Sub

    Public Overrides Sub SetData(ByRef XMLSS As CXMLSS2003, ByRef pTab As 申請Page, s処理名称 As String, ByRef pDataCreater As C総会資料Data作成)
    End Sub

End Class


Public Class CPrint申請許可名簿農地法３条
    Inherits CPrint申請許可名簿作成共通



    Public Overrides Sub SetDataRow(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet)

        pSheet.ValueReplace("{受付日}", 和暦Format(pRow.Item("受付年月日"), "gy.M.d", "-"))
        pSheet.ValueReplace("{許可番号}", Val(pRow.Item("許可番号").ToString))
        pSheet.ValueReplace("{許可年月日}", 和暦Format(pRow.Item("許可年月日"), "gy.M.d", "-"))
        pSheet.ValueReplace("{氏名A}", pRow.Item("氏名A").ToString)
        pSheet.ValueReplace("{住所A}", pRow.Item("住所A").ToString)
        pSheet.ValueReplace("{氏名B}", pRow.Item("氏名B").ToString)
        pSheet.ValueReplace("{住所B}", pRow.Item("住所B").ToString)

        複数土地設定(pSheet, pRow, Nothing)

        pSheet.ValueReplace("{農地区分}", pRow.Item("農地区分名称").ToString)
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
        End Select
    End Sub
End Class

Public Class CPrint申請許可名簿農地法４条
    Inherits CPrint申請許可名簿作成共通


    Public Overrides Sub SetDataRow(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        SetNO(pSheet)

        pSheet.ValueReplace("{受付日}", 和暦Format(pRow.Item("受付年月日"), "gy.M.d", "-"))
        pSheet.ValueReplace("{許可番号}", Val(pRow.Item("許可番号").ToString))
        pSheet.ValueReplace("{許可年月日}", 和暦Format(pRow.Item("許可年月日"), "gy.M.d", "-"))
        pSheet.ValueReplace("{氏名A}", pRow.Item("氏名A").ToString)
        pSheet.ValueReplace("{住所A}", pRow.Item("住所A").ToString)
        pSheet.ValueReplace("{氏名B}", pRow.Item("氏名B").ToString)
        pSheet.ValueReplace("{住所B}", pRow.Item("住所B").ToString)

        複数土地設定(pSheet, pRow, Nothing)

        pSheet.ValueReplace("{農地区分}", pRow.Item("農地区分名称").ToString)
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
        End Select
    End Sub
End Class
Public Class CPrint申請許可名簿農地法５条
    Inherits CPrint申請許可名簿作成共通


    Public Overrides Sub SetDataRow(ByRef pSheet As XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        SetNO(pSheet)

        pSheet.ValueReplace("{受付日}", 和暦Format(pRow.Item("受付年月日"), "gy.M.d", "-"))
        pSheet.ValueReplace("{許可番号}", Val(pRow.Item("許可番号").ToString))
        pSheet.ValueReplace("{許可年月日}", 和暦Format(pRow.Item("許可年月日"), "gy.M.d", "-"))
        pSheet.ValueReplace("{氏名A}", pRow.Item("氏名A").ToString)
        pSheet.ValueReplace("{住所A}", pRow.Item("住所A").ToString)
        pSheet.ValueReplace("{氏名B}", pRow.Item("氏名B").ToString)
        pSheet.ValueReplace("{住所B}", pRow.Item("住所B").ToString)

        複数土地設定(pSheet, pRow, Nothing)

        pSheet.ValueReplace("{農地区分}", pRow.Item("農地区分名称").ToString)
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
        End Select
    End Sub
End Class


Public Class 申請許可名簿
    Inherits CInputDate


    Private mvar開始 As DateTime = Nothing
    Private mvar終了 As DateTime = Nothing

    Public Property 許可検索範囲開始() As DateTime
        Get
            Return mvar開始
        End Get
        Set(ByVal value As DateTime)
            mvar開始 = value
        End Set
    End Property
    Public Property 許可検索範囲終了() As DateTime
        Get

            Return mvar終了
        End Get
        Set(ByVal value As DateTime)
            mvar終了 = value
        End Set
    End Property
    <Browsable(False)>
    Public Overrides ReadOnly Property DataValidate As Boolean
        Get
            If IsNothing(mvar開始) AndAlso IsNothing(mvar終了) Then
                Return False
            ElseIf mvar開始.Year > 1 AndAlso mvar終了.Year > 1 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public Overrides Function ToString() As String
        Dim sText As String = String.Format("[許可年月日]>=#{0}/{1}/{2}# and [許可年月日]<=#{3}/{4}/{5}#", mvar開始.Month, mvar開始.Day, mvar開始.Year,
         mvar終了.Month, mvar終了.Day, mvar終了.Year)
        Return sText
    End Function
End Class
