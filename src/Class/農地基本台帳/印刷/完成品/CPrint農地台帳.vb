

Public Class CPrint農地台帳
    Inherits HimTools2012.clsAccessor

    Public 農地ID As Integer = 0
    Private mvarXML As HimTools2012.Excel.XMLSS2003.CXMLSS2003
    Private mvarExecuteMode As ExcelViewMode

    Public 農地台帳DSet As DataSet
    Public pTBL公開用個人 As DataTable

    Public Overrides Sub Execute()
        Me.DataInit()
        Value = 33

        Me.MakeXMLFile()
        Value = 90
    End Sub

    Public Sub New(ByVal pXML As HimTools2012.Excel.XMLSS2003.CXMLSS2003, ByVal nID As Integer, ByVal pExcelViewMode As ExcelViewMode)
        MyBase.New()

        mvarXML = pXML

        mvarExecuteMode = pExcelViewMode
        Me.農地ID = nID
    End Sub

    Public Property XMLSS As HimTools2012.Excel.XMLSS2003.CXMLSS2003
        Get
            Return mvarXML
        End Get
        Set(ByVal value As HimTools2012.Excel.XMLSS2003.CXMLSS2003)
            mvarXML = value
        End Set
    End Property

    Public Sub SaveAndOpen(ByVal bEditMode As ExcelViewMode, ByVal PrintType As String)
        Dim sDir As String = SysAD.OutputFolder & "\農地台帳" & PrintType & ".xml"
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

    Public Sub MakeXMLFile()
        Maximum = 100
        Value = 33
        Message = "エクセルファイル作成中.."
        Application.DoEvents()
        If _Cancel Then
            Throw New Exception("Cancel")
        End If
        Value = 90
    End Sub

    Public Sub DataInit()
        Dim pView農地 As DataView = New DataView(App農地基本台帳.TBL農地.Body, "[ID]=" & Me.農地ID, "", DataViewRowState.CurrentRows)
        Dim pRow農地 As DataRowView = pView農地.Item(0)
        Dim p農地 As New CObj農地(pRow農地.Row, False)

        農地台帳DSet = New DataSet("農地台帳")

        Message = "データ取り込み中.."

        For Each pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet In mvarXML.WorkBook.WorkSheets.Items.Values
            With pSheet
                .ValueReplace("{システム年月日}", 和暦Format(Now))
                .ValueReplace("{市町村名}", SysAD.DB(sLRDB).DBProperty("市町村名").ToString)
                .ValueReplace("{土地所在}", p農地.土地所在)

                If p農地.登記簿地目 <> 0 Then
                    Dim p登記地目 As DataRow = App農地基本台帳.TBL地目.Rows.Find(p農地.登記簿地目)

                    If p登記地目 Is Nothing Then : .ValueReplace("{登記地目}", "-")
                    Else : .ValueReplace("{登記地目}", p登記地目.Item("名称").ToString)
                    End If
                Else
                    .ValueReplace("{登記地目}", "-")
                End If

                .ValueReplace("{登記面積}", HimTools2012.NumericFunctions.NumToString(p農地.登記簿面積))

                If Not IsDBNull(pRow農地.Item("農振法区分")) Then
                    Select Case pRow農地.Item("農振法区分")
                        Case 1 : .ValueReplace("{農振法}", "農用地内")
                        Case 2 : .ValueReplace("{農振法}", "農用地外")
                        Case 3 : .ValueReplace("{農振法}", "振興地域外")
                        Case Else : .ValueReplace("{農振法}", "-")
                    End Select
                Else
                    Select Case p農地.旧農振区分
                        Case enum農業振興地域.農用地外 : .ValueReplace("{農振法}", "農用地外")
                        Case enum農業振興地域.農用地内 : .ValueReplace("{農振法}", "農用地内")
                        Case enum農業振興地域.振興地域外 : .ValueReplace("{農振法}", "振興地域外")
                        Case Else : .ValueReplace("{農振法}", "-")
                    End Select
                End If

                If Not IsDBNull(pRow農地.Item("都市計画法区分")) Then
                    Select Case pRow農地.Item("都市計画法区分")
                        Case 1 : .ValueReplace("{都市計画法}", "市街化区域")
                        Case 2 : .ValueReplace("{都市計画法}", "市街化調整区域")
                        Case 3 : .ValueReplace("{都市計画法}", "その他")
                        Case Else : .ValueReplace("{都市計画法}", "-")
                    End Select
                Else
                    Select Case p農地.都市計画法
                        Case enum都市計画法.都市計画法内 : .ValueReplace("{都市計画法}", "市街化区域")
                        Case enum都市計画法.用途地域内 : .ValueReplace("{都市計画法}", "市街化区域")
                        Case enum都市計画法.調整区域内 : .ValueReplace("{都市計画法}", "市街化調整区域")
                        Case enum都市計画法.市街化区域内 : .ValueReplace("{都市計画法}", "市街化区域")
                        Case enum都市計画法.都市計画白地 : .ValueReplace("{都市計画法}", "その他")
                        Case Else : .ValueReplace("{都市計画法}", "-")
                    End Select
                End If

                Select Case p農地.生産緑地法
                    Case enum有無.有 : .ValueReplace("{生産緑地法}", "生産緑地法")
                    Case enum有無.無 : .ValueReplace("{生産緑地法}", "-")
                    Case Else : .ValueReplace("{生産緑地法}", "-")
                End Select

                If IsDBNull(pRow農地.Item("登記名義人氏名")) Then
                    If p農地.所有者氏名.Length > 0 Then : .ValueReplace("{所有者名}", p農地.所有者氏名)
                    Else : .ValueReplace("{所有者名}", "")
                    End If
                Else
                    .ValueReplace("{所有者名}", pRow農地.Item("登記名義人氏名"))
                End If

                Select Case Val(pRow農地.Item("所有者農地意向").ToString)
                    Case 1 : .ValueReplace("{意向内容}", "所有権移転")
                    Case 2 : .ValueReplace("{意向内容}", "貸付")
                    Case 3 : .ValueReplace("{意向内容}", "人・農地プランへの位置づけ")
                    Case 4 : .ValueReplace("{意向内容}", "農地中間管理機構への貸付")
                    Case 5 : .ValueReplace("{意向内容}", "その他")
                    Case Else : .ValueReplace("{意向内容}", "-")
                End Select

                .ValueReplace("{共有者}", "")

                Dim 耕作者ID As Decimal = 0
                Dim 耕作者名 As String = ""

                If Val(pRow農地.Item("自小作別").ToString) > 0 Then
                    耕作者ID = Val(pRow農地.Item("借受人ID").ToString)
                Else
                    If Val(pRow農地.Item("管理者ID").ToString) <> 0 Then : 耕作者ID = Val(pRow農地.Item("管理者ID").ToString)
                    Else : 耕作者ID = Val(pRow農地.Item("所有者ID").ToString)
                    End If
                End If

                Dim p耕作者名 As DataRow = App農地基本台帳.TBL個人.FindRowByID(耕作者ID)
                If p耕作者名 Is Nothing Then
                    .ValueReplace("{耕作者名}", "")
                    .ValueReplace("{整理番号}", "")
                Else
                    耕作者名 = p耕作者名.Item("氏名").ToString
                    .ValueReplace("{耕作者名}", p耕作者名.Item("氏名").ToString)

                    pTBL公開用個人 = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT * FROM [D_公開用個人] WHERE PID = {0}", 耕作者ID))
                    If pTBL公開用個人.Rows.Count > 0 Then
                        For Each pRow As DataRow In pTBL公開用個人.Rows
                            .ValueReplace("{整理番号}", HimTools2012.StringF.Right("000000000000000000" & pRow.Item("AutoID"), 18))
                        Next
                    Else
                        .ValueReplace("{整理番号}", "")
                    End If
                End If

                If p農地.自小作別 = enum自小作別.自作 Then
                    .ValueReplace("{権利種類}", "")
                    .ValueReplace("{存続期間}", "")
                Else
                    Dim 存続始期期間 As DateTime = p農地.貸借始期
                    Dim 存続終期期間 As DateTime = p農地.貸借終期

                    .ValueReplace("{権利種類}", p農地.小作形態種別)
                    If IsDBNull(存続始期期間) OrElse Not IsDate(存続始期期間) OrElse Year(存続始期期間) < 1901 Then
                        .ValueReplace("{存続期間}", "")
                    Else
                        If IsDBNull(存続終期期間) OrElse Not IsDate(存続終期期間) OrElse Year(存続終期期間) < 1901 Then
                            .ValueReplace("{存続期間}", "")
                        Else
                            .ValueReplace("{存続期間}", 和暦Format(存続始期期間) & " ～ " & 和暦Format(存続終期期間))
                        End If
                    End If
                End If

                If Val(pRow農地.Item("経由農業生産法人ID").ToString) = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID")) Then
                    .ValueReplace("{中間管理権}", "中間管理機構を介した貸借農地")
                Else
                    .ValueReplace("{中間管理権}", "")
                End If


                Select Case p農地.GetIntegerValue("利用状況調査荒廃")
                    Case enum利用状況調査農地法.不明 : .ValueReplace("{利用状況調査結果}", "-")
                    Case enum利用状況調査農地法.農法32条1項1号 : .ValueReplace("{利用状況調査結果}", "現在、耕作されておらず、引き続き、耕作されないと見込まれる農地")
                    Case enum利用状況調査農地法.農法32条1項2号 : .ValueReplace("{利用状況調査結果}", "農業上の利用程度が周辺の農地の利用の程度に比べて著しく劣っていると認められる農地")
                    Case Else : .ValueReplace("{利用状況調査結果}", "遊休農地でない")
                End Select

                Select Case p農地.GetIntegerValue("利用意向意向内容区分")
                    Case enum利用意向内容区分.不明 : .ValueReplace("{利用意向調査結果}", "")
                    Case enum利用意向内容区分.自ら耕作 : .ValueReplace("{利用意向調査結果}", "自ら耕作する")
                    Case enum利用意向内容区分.機構事業 : .ValueReplace("{利用意向調査結果}", "農地中間管理事業を利用する")
                    Case enum利用意向内容区分.所有者代理事業 : .ValueReplace("{利用意向調査結果}", "農地利用集積円滑化団体が行う農地所有者代理事業を利用する")
                    Case enum利用意向内容区分.権利設定または移転 : .ValueReplace("{利用意向調査結果}", "自ら所有者の移転または貸借権その他の使用収益を目的とする権利の設定もしくは移転を行う")
                    Case enum利用意向内容区分.その他 : .ValueReplace("{利用意向調査結果}", "その他の場合")
                    Case Else : .ValueReplace("{利用意向調査結果}", "")
                End Select

                .ValueReplace("{措置の実施状況}", "")
            End With
        Next
    End Sub
End Class
