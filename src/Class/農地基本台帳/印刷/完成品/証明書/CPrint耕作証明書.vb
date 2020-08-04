Imports HimTools2012
Imports HimTools2012.NumericFunctions
Imports System.ComponentModel
Imports System.ComponentModel.TypeConverter
Imports HimTools2012.controls.PropertyGridSupport
Imports HimTools2012.TypeConverterCustom

Public MustInherit Class CPrint耕作証明共通
    Inherits HimTools2012.clsAccessor
    Protected mvarKey As String = ""
    Protected mvarXML As HimTools2012.Excel.XMLSS2003.CXMLSS2003 = Nothing
    Protected mvarFileName As String
    Protected dt発行日 As DateTime = Nothing
    Protected n発行番号 As Integer = 0

    Protected 世帯ID As Integer
    Protected 個人ID As Integer
    Protected 申請者名 As String
    Protected 申請者住所 As String
    Protected 出力条件 As C耕作証明条件

    Public Sub New(ByVal sKey As String, ByVal sFileName As String)
        mvarKey = sKey

        mvarFileName = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & sFileName
        If IO.File.Exists(mvarFileName) Then
            mvarXML = New HimTools2012.Excel.XMLSS2003.CXMLSS2003(HimTools2012.TextAdapter.LoadTextFile(mvarFileName))


            Dim nID As Long = CommonFunc.GetKeyCode(mvarKey)
            Dim 主ID As Long = 0
            Select Case CommonFunc.GetKeyHead(mvarKey)
                Case "農家"
                    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:世帯Info] WHERE [ID]={0}", nID)
                    主ID = Val(pTBL.Rows(0).Item("世帯主ID").ToString)
                    Dim s氏名 As New List(Of String)
                    pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [世帯ID]={0}", nID)
                    For Each pRow As DataRow In pTBL.Rows
                        s氏名.Add(pRow.Item("氏名").ToString)
                        If 主ID = pRow.Item("ID") Then
                            申請者名 = pRow.Item("氏名").ToString()
                            申請者住所 = pRow.Item("住所").ToString()
                        End If
                    Next

                    選択氏名StandardValuesCollection = New StandardValuesCollection(s氏名.ToArray())

                Case "個人"
                    Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [ID]={0}", nID)
                    主ID = nID
                    選択氏名StandardValuesCollection = New StandardValuesCollection(New String() {pTBL.Rows(0).Item("氏名").ToString})
                    申請者名 = pTBL.Rows(0).Item("氏名").ToString
                    申請者住所 = pTBL.Rows(0).Item("住所").ToString()
            End Select

            With New HimTools2012.PropertyGridDialog(New C耕作証明条件(Val(SysAD.DB(sLRDB).DBProperty("耕作証明番号")) + 1, 申請者名, 申請者住所), "耕作証明書")

                If .ShowDialog = DialogResult.OK Then
                    出力条件 = .ResultProperty
                    With CType(.ResultProperty, C耕作証明条件)
                        申請者名 = .申請者氏名
                        申請者住所 = .申請者住所
                        dt発行日 = .発行日
                        n発行番号 = .発行番号
                    End With
                Else
                    n発行番号 = 0
                End If
            End With
        End If
    End Sub

    Public Property XMLSS As HimTools2012.Excel.XMLSS2003.CXMLSS2003
        Get
            Return mvarXML
        End Get
        Set(ByVal value As HimTools2012.Excel.XMLSS2003.CXMLSS2003)
            mvarXML = value
        End Set
    End Property

    Public Sub SaveAndOpen(ByVal bEditMode As ExcelViewMode)
        Dim sDir As String = SysAD.OutputFolder & "\耕作面積証明書.xml"
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
End Class

''' <summary>
''' 耕作面積面積証明
''' </summary>
''' <remarks></remarks>
Public Class CPrint耕作証明書
    Inherits CPrint耕作証明共通

    Private mvarData As PrnData
    Private mvarParamsDic As New Dictionary(Of String, Object)


    Private Structure PrnData
        Dim 住所 As String

        Dim n田合計 As Decimal
        Dim n畑合計 As Decimal
        Dim n樹合計 As Decimal

        Dim 都道府県名 As String
        Dim 会長名 As String
        Dim 会長肩書 As String
        Dim b申請者住所を埋める As Boolean
    End Structure

    Public Overrides Sub Execute()
        Me.DataInit()
        Value = 33

        Me.MakeXMLFile()
        Value = 90
    End Sub

    Public Sub New(ByVal sKey As String)
        MyBase.New(sKey, "\耕作面積証明書.xml")
        If n発行番号 > 0 Then

            Me.Dialog.StartProc(True, True)

            If Me.Dialog._objException Is Nothing = False Then
                If Me.Dialog._objException.Message = "Cancel" Then
                    MsgBox("処理を中止しました。　", , "処理中止")
                Else
                End If
            Else
                Me.SaveAndOpen(ExcelViewMode.Preview)
                SysAD.DB(sLRDB).DBProperty("耕作証明番号") = n発行番号
            End If
        End If

    End Sub

    Public Sub DataInit()

        With mvarData
            .都道府県名 = SysAD.DB(sLRDB).DBProperty("都道府県名")
            .会長名 = SysAD.DB(sLRDB).DBProperty("会長名")
            .会長肩書 = IIf(Val(SysAD.DB(sLRDB).DBProperty("会長代理")), "農業委員会  　会長代理", "農業委員会  　会長")
            .b申請者住所を埋める = Val(SysAD.DB(sLRDB).DBProperty("耕作証明の申請者住所を埋める"))
        End With

        Dim nID As Long = CommonFunc.GetKeyCode(mvarKey)
        mvarParamsDic.Add("発行番号", n発行番号)
        SysAD.DB(sLRDB).ExecuteSQL("UPDATE V_農地 SET V_農地.農地状況 = 0 WHERE (((V_農地.農地状況) Is Null));")

        Select Case CommonFunc.GetKeyHead(mvarKey)
            Case "農家"
                With mvarParamsDic
                    Dim n世帯ID As Long = CommonFunc.GetKeyCode(mvarKey)
                    Dim s所有世帯 As String = ""
                    Dim s市外農地 As String = ""
                    Select Case 出力条件.管理者の影響
                        Case C耕作証明条件.enum管理人.管理人を考慮しない
                            s所有世帯 = String.Format("[V_農地].[所有世帯ID]={0}", n世帯ID)
                        Case C耕作証明条件.enum管理人.管理人を考慮する
                            s所有世帯 = String.Format("IIF([V_農地].[管理世帯ID]<>0,[V_農地].[管理世帯ID]={0},[V_農地].[所有世帯ID]={0})", n世帯ID)
                    End Select

                    Select Case 出力条件.市外農地を含む
                        Case C耕作証明条件.enum市外農地.含む

                        Case C耕作証明条件.enum市外農地.含まない
                            s所有世帯 = s所有世帯 & " AND [大字ID] > 0"
                            s市外農地 = "((V_農地.大字ID)>0) AND "
                    End Select

                    Dim pTBLA As DataTable
                    If s所有世帯 <> "" Then
                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計, Sum([採草放牧面積]) AS 採草合計 FROM V_農地 WHERE (((V_農地.農地状況)<20 Or (V_農地.農地状況) Is Null) AND ((V_農地.自小作別)=0) AND (" & s所有世帯 & ")) OR (((V_農地.農地状況)<20 Or (V_農地.農地状況) Is Null) AND (" & s所有世帯 & ") AND ((V_農地.借受世帯ID)={1}) AND ((V_農地.経由農業生産法人ID) Is Not Null));", s所有世帯, n世帯ID)
                    Else
                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計, Sum([採草放牧面積]) AS 採草合計 FROM V_農地 WHERE ([所有世帯ID]=" & n世帯ID & ") AND [自小作別]<1;")
                    End If

                    With New HimTools2012.Data.DataRowPlus(pTBLA.Rows(0))
                        Add面積("自田面積", .Item("田合計", 0))
                        Add面積("自畑面積", .Item("畑合計", 0))
                        Add面積("自樹面積", .Item("樹園地合計", 0))
                        Add面積("自採面積", .Item("採草合計", 0))
                    End With

                    If s所有世帯 <> "" Then
                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計, Sum([採草放牧面積]) AS 採草合計 FROM V_農地 WHERE (" & s市外農地 & "((V_農地.農地状況)<20 Or (V_農地.農地状況) Is Null) AND ((V_農地.借受世帯ID)={0}) AND ((V_農地.自小作別)>0) AND ((V_農地.所有世帯ID)<>{0}))", n世帯ID)
                    Else
                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計, Sum([採草放牧面積]) AS 採草合計 FROM V_農地 WHERE ([借受世帯ID]=" & n世帯ID & ") AND [自小作別]>0;")
                    End If

                    With New HimTools2012.Data.DataRowPlus(pTBLA.Rows(0))
                        Add面積("小田面積", .Item("田合計", 0))
                        Add面積("小畑面積", .Item("畑合計", 0))
                        Add面積("小樹面積", .Item("樹園地合計", 0))
                        Add面積("小採面積", .Item("採草合計", 0))
                    End With

                    If s所有世帯 <> "" Then
                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計,Sum([採草放牧面積]) AS 採草合計 FROM [V_農地] WHERE ([V_農地].[農地状況]<20 Or [V_農地].[農地状況] Is Null) AND [V_農地].[自小作別]>0 AND [借受世帯ID]<>{0} AND ({1});", n世帯ID, s所有世帯)
                    Else
                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計,Sum([採草放牧面積]) AS 採草合計 FROM [V_農地] WHERE ([所有世帯ID]=" & n世帯ID & ") AND ([借受世帯ID]<>" & n世帯ID & ") AND ([自小作別]>0);")
                    End If

                    With New HimTools2012.Data.DataRowPlus(pTBLA.Rows(0))
                        Add面積("貸田面積", .Item("田合計", 0))
                        Add面積("貸畑面積", .Item("畑合計", 0))
                        Add面積("貸樹面積", .Item("樹園地合計", 0))
                        Add面積("貸採面積", .Item("採草合計", 0))
                    End With

                    Add面積("計田面積", Val(.Item("自田面積") + Val(.Item("小田面積"))))
                    Add面積("計畑面積", Val(.Item("自畑面積") + Val(.Item("小畑面積"))))
                    Add面積("計樹面積", Val(.Item("自樹面積") + Val(.Item("小樹面積"))))
                    Add面積("計採面積", Val(.Item("自採面積") + Val(.Item("小採面積"))))

                    Add面積("計自面積", Val(.Item("自田面積")) + Val(.Item("自畑面積")) + Val(.Item("自樹面積")) + Val(.Item("自採面積")))
                    Add面積("計小面積", Val(.Item("小田面積")) + Val(.Item("小畑面積")) + Val(.Item("小樹面積")) + Val(.Item("小採面積")))
                    Add面積("計貸面積", Val(.Item("貸田面積")) + Val(.Item("貸畑面積")) + Val(.Item("貸樹面積")) + Val(.Item("貸採面積")))
                    Add面積("計面積", Val(.Item("計田面積")) + Val(.Item("計畑面積")) + Val(.Item("計樹面積")) + Val(.Item("計採面積")))
                End With

            Case "個人"
                Dim pTBLA As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [ID]=" & nID)
                Dim pRowA As HimTools2012.Data.DataRowPlus
                pRowA = New HimTools2012.Data.DataRowPlus(pTBLA.Rows(0))

                With mvarParamsDic
                    Dim s所有者 As String = ""
                    Select Case 出力条件.管理者の影響
                        Case C耕作証明条件.enum管理人.管理人を考慮しない
                            s所有者 = "[所有者ID]={0}"
                        Case C耕作証明条件.enum管理人.管理人を考慮する
                            s所有者 = "IIF([V_農地].[管理者ID]<>0,[V_農地].[管理者ID]={0},[V_農地].[所有者ID]={0})"
                    End Select
                    Select Case 出力条件.市外農地を含む
                        Case C耕作証明条件.enum市外農地.含む

                        Case C耕作証明条件.enum市外農地.含まない
                            s所有者 = s所有者 & " AND [大字ID] > 0"
                    End Select

                    Try
                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計,Sum([採草放牧面積]) AS 採草合計 FROM [V_農地] WHERE (((V_農地.農地状況)<20) AND (" & s所有者 & ") AND ((V_農地.自小作別)=0)) OR (((V_農地.農地状況)<20) AND (" & s所有者 & ") AND ((V_農地.借受人ID)={0}) AND ((V_農地.経由農業生産法人ID) Is Not Null));", nID)
                        Add面積("自田面積", pTBLA.Rows(0).Item("田合計"))
                        Add面積("自畑面積", pTBLA.Rows(0).Item("畑合計"))
                        Add面積("自樹面積", pTBLA.Rows(0).Item("樹園地合計"))
                        Add面積("自採面積", pTBLA.Rows(0).Item("採草合計"))

                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計, Sum([採草放牧面積]) AS 採草合計 FROM V_農地 WHERE (((V_農地.農地状況)<20) AND ((V_農地.借受人ID)={0}) AND ((V_農地.自小作別)>0) AND ((V_農地.所有者ID)<>{0}));", nID)
                        Add面積("小田面積", pTBLA.Rows(0).Item("田合計"))
                        Add面積("小畑面積", pTBLA.Rows(0).Item("畑合計"))
                        Add面積("小樹面積", pTBLA.Rows(0).Item("樹園地合計"))
                        Add面積("小採面積", pTBLA.Rows(0).Item("採草合計"))

                        pTBLA = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Sum([V_農地].田面積) AS 田合計, Sum([V_農地].畑面積) AS 畑合計, Sum([V_農地].樹園地) AS 樹園地合計,Sum([採草放牧面積]) AS 採草合計 FROM [V_農地] WHERE [V_農地].[農地状況]<20 AND ([V_農地].借受人ID<>{0} AND (" & s所有者 & ") AND [V_農地].[自小作別]>0);", nID)
                        Add面積("貸田面積", pTBLA.Rows(0).Item("田合計"))
                        Add面積("貸畑面積", pTBLA.Rows(0).Item("畑合計"))
                        Add面積("貸樹面積", pTBLA.Rows(0).Item("樹園地合計"))
                        Add面積("貸採面積", pTBLA.Rows(0).Item("採草合計"))

                        Add面積("計田面積", Val(.Item("自田面積")) + Val(.Item("小田面積")))
                        Add面積("計畑面積", Val(.Item("自畑面積")) + Val(.Item("小畑面積")))
                        Add面積("計樹面積", Val(.Item("自樹面積")) + Val(.Item("小樹面積")))
                        Add面積("計採面積", Val(.Item("自採面積")) + Val(.Item("小採面積")))
                        Add面積("計自面積", Val(.Item("自田面積")) + Val(.Item("自畑面積")) + Val(.Item("自樹面積")) + Val(.Item("自採面積")))
                        Add面積("計小面積", Val(.Item("小田面積")) + Val(.Item("小畑面積")) + Val(.Item("小樹面積")) + Val(.Item("小採面積")))
                        Add面積("計貸面積", Val(.Item("貸田面積")) + Val(.Item("貸畑面積")) + Val(.Item("貸樹面積")) + Val(.Item("貸採面積")))
                        Add面積("計面積", Val(.Item("計田面積")) + Val(.Item("計畑面積")) + Val(.Item("計樹面積")) + Val(.Item("計採面積")))
                    Catch ex As Exception

                    End Try
                End With
        End Select

        For Each pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet In mvarXML.WorkBook.WorkSheets.Items.Values
            With pSheet
                .ValueReplace("{市町村名}", SysAD.市町村.市町村名)
                .ValueReplace("{会長名}", mvarData.会長名)
                .ValueReplace("{発行番号}", mvarParamsDic.Item("発行番号"))
                .ValueReplace("{発行年月日}", 和暦Format(dt発行日))
                .ValueReplace("{住所}", Me.申請者住所)
                .ValueReplace("{氏名}", Me.申請者名)

                Try
                    .ValueReplace("{自田面積}", NumToString(mvarParamsDic.Item("自田面積")))
                    .ValueReplace("{自畑面積}", NumToString(mvarParamsDic.Item("自畑面積")))
                    .ValueReplace("{自樹面積}", NumToString(mvarParamsDic.Item("自樹面積")))
                    .ValueReplace("{自採面積}", NumToString(mvarParamsDic.Item("自採面積")))
                    .ValueReplace("{小田面積}", NumToString(mvarParamsDic.Item("小田面積")))
                    .ValueReplace("{小畑面積}", NumToString(mvarParamsDic.Item("小畑面積")))
                    .ValueReplace("{小樹面積}", NumToString(mvarParamsDic.Item("小樹面積")))
                    .ValueReplace("{小採面積}", NumToString(mvarParamsDic.Item("小採面積")))
                    .ValueReplace("{貸田面積}", NumToString(mvarParamsDic.Item("貸田面積")))
                    .ValueReplace("{貸畑面積}", NumToString(mvarParamsDic.Item("貸畑面積")))
                    .ValueReplace("{貸樹面積}", NumToString(mvarParamsDic.Item("貸樹面積")))
                    .ValueReplace("{貸採面積}", NumToString(mvarParamsDic.Item("貸採面積")))
                    .ValueReplace("{計田面積}", NumToString(mvarParamsDic.Item("計田面積")))
                    .ValueReplace("{計畑面積}", NumToString(mvarParamsDic.Item("計畑面積")))
                    .ValueReplace("{計樹面積}", NumToString(mvarParamsDic.Item("計樹面積")))
                    .ValueReplace("{計採面積}", NumToString(mvarParamsDic.Item("計採面積")))
                    .ValueReplace("{計自面積}", NumToString(mvarParamsDic.Item("計自面積")))
                    .ValueReplace("{計小面積}", NumToString(mvarParamsDic.Item("計小面積")))
                    .ValueReplace("{計貸面積}", NumToString(mvarParamsDic.Item("計貸面積")))
                    .ValueReplace("{計面積}", NumToString(mvarParamsDic.Item("計面積")))
                Catch ex As Exception
                    Stop
                End Try

                Select Case SysAD.市町村.市町村名
                    Case "日置市" : .ValueReplace("{農委区分}", "日農委")
                    Case Else : .ValueReplace("{農委区分}", "")
                End Select
            End With
        Next
    End Sub

    Private Sub Add面積(ByVal sName As String, ByVal oValue As Object)
        If oValue Is Nothing OrElse IsDBNull(oValue) OrElse Not IsNumeric(oValue) Then
            mvarParamsDic.Add(sName, 0)
        Else
            mvarParamsDic.Add(sName, Val(oValue))
        End If
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

End Class

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class C耕作証明条件
    Inherits HimTools2012.InputSupport.CInputSupport

    Private mvar管理者の影響 As enum管理人
    Private mvar対象農地 As enum市外農地

    Public Enum enum管理人
        管理人を考慮する
        管理人を考慮しない
    End Enum

    Public Enum enum市外農地
        含まない
        含む
    End Enum

    Public Sub New(n発行番号 As Integer, s申請者名 As String, s申請者住所 As String)
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)
        Me.発行番号 = n発行番号
        Me.発行日 = Now.Date

        Me.申請者氏名 = s申請者名
        Me.申請者住所 = s申請者住所

        Dim St As String = SysAD.DB(sLRDB).DBProperty("耕作証明集計条件管理人", enum管理人.管理人を考慮しない)
        mvar管理者の影響 = Val(St)
        mvar対象農地 = Val(enum市外農地.含まない)
    End Sub

    <PropertyOrderAttribute(1)> <Category("01 印刷条件")> <DefaultValue("")> <Description("発行日")>
    Public Property 発行日 As DateTime
    <PropertyOrderAttribute(2)> <Category("01 印刷条件")> <DefaultValue("")> <Description("発行番号")>
    Public Property 発行番号 As Integer
    <PropertyOrderAttribute(3)> <TypeConverter(GetType(氏名Converter)), CategoryAttribute("02 申請者情報")>
    Public Property 申請者氏名 As String
    <PropertyOrderAttribute(4)> <Category("02 申請者情報")> <DefaultValue("")> <Description("申請者住所")>
    Public Property 申請者住所 As String
    <PropertyOrderAttribute(5)> <Category("03 申請者情報")> <DefaultValue("")> <Description("耕作者を検索する際、農地の管理者を考慮するか設定する")>
    Public Property 管理者の影響 As enum管理人
        Get
            Return mvar管理者の影響
        End Get
        Set(value As enum管理人)
            SysAD.DB(sLRDB).DBProperty("耕作証明集計条件管理人") = value
            mvar管理者の影響 = value
        End Set
    End Property
    <PropertyOrderAttribute(6)> <Category("04 農地情報")> <DefaultValue("")> <Description("市外(町外)農地を含む")>
    Public Property 市外農地を含む As enum市外農地
        Get
            Return mvar対象農地
        End Get
        Set(value As enum市外農地)
            mvar対象農地 = value
        End Set
    End Property


    Public Overrides Function DataCompleate() As Boolean
        Return MyBase.DataCompleate()
    End Function

End Class


Public Class 氏名Converter
    Inherits StringConverter

    Public Overloads Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overloads Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection

        Return modCommon.選択氏名StandardValuesCollection
    End Function
    Public Overloads Overrides Function GetStandardValuesExclusive(
               ByVal context As ITypeDescriptorContext) As Boolean
        Return False
    End Function

End Class


