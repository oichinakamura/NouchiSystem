Public Class COutPutCSV
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Public Sub New(ByVal bCloseable As Boolean, ByVal toolbarvisible As Boolean, ByVal ObjectName As String, ByVal sTitle As String)
        MyBase.New(bCloseable, toolbarvisible, ObjectName, sTitle)
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property
End Class

Public Enum 変換区分
    全角 = 1
    半角 = 2
    日付 = 3
    面積 = 4
    登記簿地目 = 5
    現況地目 = 6
    字コード = 7
    外字 = 8
End Enum

Public Class StringBEx
    Public mvarBody As System.Text.StringBuilder

    Public Sub New(ByVal s初期値 As String)
        mvarBody = New System.Text.StringBuilder(s初期値)
    End Sub
    Public ReadOnly Property Body As System.Text.StringBuilder
        Get
            Return mvarBody
        End Get
    End Property

    Private List外字 As String() = {";増", ";信", ";勤", ";尭", ";今", ";飫", ";勘", ";枦", ";塚", ";兎", ";溝", ";樋", ";鉋", ";筌", ";刃", ";棚", ";柊", ";葛", ";喰", ";屏", ";籾", ";餅", ";榔", ";己", "棚;高﨑", ";高"}
    Public Sub SetNumber(ByVal pData As Object, ByVal bOption As 変換区分)
        If pData Is Nothing Then
        ElseIf IsDBNull(pData) Then
        Else
            Select Case bOption
                Case 変換区分.全角
                    mvarBody.Append(",""" & StrConv(pData.ToString, VbStrConv.Wide) & """")
                Case 変換区分.半角
                    mvarBody.Append(",""" & StrConv(pData.ToString, VbStrConv.Narrow) & """")
                Case 変換区分.日付
                    If IsDate(pData) Then
                        mvarBody.Append(",""" & Format(CDate(pData), "yyyy/MM/dd") & """")
                    Else
                        mvarBody.Append(",""" & pData & """")
                    End If
                Case 変換区分.面積  'ここで小数点第２位まで
                    pData = Math.Round(pData, 2)
                    mvarBody.Append(",""" & pData.ToString & """")
                Case 変換区分.登記簿地目 '読み取り専用のため
                    If pData.ToString = "" Then
                        mvarBody.Append(",""8""")
                    Else
                        mvarBody.Append(",""" & StrConv(pData.ToString, VbStrConv.Narrow) & """")
                    End If
                Case 変換区分.現況地目  '読み取り専用のため
                    If pData.ToString = "" Then
                        mvarBody.Append(",""9""")
                    Else
                        mvarBody.Append(",""" & StrConv(pData.ToString, VbStrConv.Narrow) & """")
                    End If
                Case 変換区分.字コード
                    If Val(pData.ToString) < 1 Then
                        pData = ""
                    End If
                    mvarBody.Append(",""" & StrConv(pData.ToString, VbStrConv.Narrow) & """")
                Case 変換区分.外字
                    If Len(pData.ToString) > 0 Then
                        For n As Integer = 0 To UBound(List外字)
                            Dim Ar As Object = Split(List外字(n), ";")

                            pData = Replace(pData, Ar(0), Ar(1))
                        Next

                        If pData = "-" Then
                            pData = ""
                        End If
                    End If

                    mvarBody.Append(",""" & StrConv(pData.ToString, VbStrConv.Wide) & """")
                Case Else
                    mvarBody.Append(",""" & pData.ToString & """")
            End Select
        End If
    End Sub
End Class

Public Class 登記簿地目変換
    Inherits DataTable
    '
    Private mvarDSet As New HimTools2012.Data.DataSetEx
    Private tbl農地 As DataTable

    Public Sub New()

    End Sub

    Public Function Init() As 登記簿地目変換
        Dim 登記Row As DataRow
        Dim s登記変換リスト As String() = {"1:田", "1:宅地介在田", "1:市街化田", "1:介在田", "2:畑", "2:宅地介在畑", "2:市街化畑", "3:牧場", "4:宅地", "4:宅地(農業用施設用地)", "4:準宅地", "4:複合利用鉄軌道用地", "5:山林", "5:宅地介在山林", "5:市街化山林", "5:農地介在山林", _
                                           "6:原野", "6:宅地介在原野", "6:市街化原野", "6:農地介在原野", "7:雑種地", "7:雑種地(農業用施設用地)", "7:雑種地(田畑)", "7:雑種地(山林)", "7:準雑地", "7:雑地", "7:雑種", "7:その他の雑種地", "8:その他", "8:海成り", "8:川成り", "8:官有地", "8:-", "9:公衆用道路", "9:市道", "9:県道", "9:国道", "9:公衆道", "10:公用地", _
                                           "11:公共用地", "12:公園", "13:鉄道用地", "13:鉄軌道用地", "14:学校用地", "14:学校敷地", "15:水道用地", "15:水道", "16:用悪水路", "17:池沼", "18:溜池", "18:ため池", "19:墓場", "19:墓地", "20:境内地", _
                                           "21:堤", "21:堤とう", "22:井溝", "23:運河用地", "24:保安林", "25:塩田", "26:鉱泉地", "27:河川敷地", "27:河川敷"}
        Dim Ar As Object = Nothing

        Me.Columns.Add(New DataColumn("登記ID", GetType(Integer)))
        Me.Columns.Add(New DataColumn("登記名称", GetType(String)))
        For n As Integer = 0 To UBound(s登記変換リスト)
            登記Row = Me.NewRow

            Ar = Split(s登記変換リスト(n), ":")
            登記Row("登記ID") = Ar(0)
            登記Row("登記名称") = Ar(1)

            Me.Rows.Add(登記Row)
        Next

        Return Me
    End Function
End Class

Public Class 現況地目変換
    Inherits DataTable
    '
    Private mvarDSet As New DataSet
    Private tbl農地 As DataTable

    Public Sub New()

    End Sub

    Public Function Init() As 現況地目変換
        Dim 現況Row As DataRow

        Dim s現況変換リスト As String() = {"1:田", "1:宅地介在田", "1:介在田", "1:宅地介田", "1:市街化田", "2:畑", "2:宅地介在畑", "2:介在畑", "2:宅地介畑", "2:市街化畑", "3:樹園地", "3:樹園地(桑)", "3:樹園地(茶)", "3:樹園地(果樹)", "4:牧草放牧地", "4:採草放牧地", "5:宅地", "5:宅地（農施用地）", "5:宅地（農施）", "5:宅地（農業用施設用地）", "5:雑種地介在宅地", "5:防火水槽", _
                                           "6:山林原野", "6:山林・原野", "6:山林", "6:宅地介在山林", "6:農地介在山林", "6:市街化山林", "6:保安林", "6:砂防指定林", "6:原野", "6:宅地介在原野", "6:市街化原野", "6:農地介在原野", "7:雑種地", "7:雑種地他", "7:その他雑種地", "7:その他の雑種地", "7:太陽光発電", "7:準雑地", "7:準雑", "7:雑種地（田畑）", "7:雑種地（農施用）", "7:雑種地（農施用地）", "7:雑種地（田）", "7:雑種地（畑）", "7:雑種地（宅地）", "7:雑（宅地1）", "7:雑（宅地3）", "7:雑（宅地5）", "7:雑（宅地7）", "7:資材地", _
                                           "7:雑種地（山林）", "7:雑種", "7:雑種（農地）", "7:雑種（山林）", "7:雑種（農施）", "7:遊園地", "8:農業用施設", "8:農業用施設用地", "8:農用施設用地", "9:その他", "9:-", "9:ため池", "9:池沼", "9:溜池・井溝", "9:墓地", "9:共同利用地等", "9:牧場", "9:農地外", "9:公有地", "9:堤とう", "9:堤", "9:堤塘", "9:貯水池", "9:鉱泉地", "9:公民館", "9:公民館用地", "9:境内地", "9:現地なし", "9:現地確認不能", "9:現確不能", _
                                           "101:農家住宅", "101:農家用宅地", "102:一般個人住宅", "103:集合住宅等", "111:道路", "111:私有道路", "111:公衆用道路", "111:公衆道", "111:私道", "111:道路敷", "112:水路・河川", "112:用悪水路", "112:河川敷", "112:河川敷き", "112:河川敷地", "112:水道用地", "112:防火用水", "112:水道", "112:河川区域", "113:鉄道敷地", "113:鉄軌道用地", "114:砂利採取", _
                                           "121:個人農林業施設", "122:共同農林業施設", "131:鉱工業用地", "141:運輸通信用地", "151:商業サービス", "152:ゴルフ場", "152:ゴルフ場用地", "153:宿泊施設等", "154:その他サービス", _
                                           "161:公共施設", "162:学校用地", "162:学校敷地", "163:公園・運動場", "163:公園", "163:公園緑地", "164:その他公共施設", "171:植林", "181:基盤強化法転用", "191:露天資材置場", "192:露天駐車場"}
        Dim Ar As Object = Nothing

        Me.Columns.Add(New DataColumn("現況ID", GetType(Integer)))
        Me.Columns.Add(New DataColumn("現況名称", GetType(String)))
        'Me.PrimaryKey = New DataColumn() {Me.Columns("現況名称")}

        For n As Integer = 0 To UBound(s現況変換リスト)
            現況Row = Me.NewRow

            Ar = Split(s現況変換リスト(n), ":")
            現況Row("現況ID") = Ar(0)
            現況Row("現況名称") = Ar(1)

            Me.Rows.Add(現況Row)
        Next

        Return Me
    End Function
End Class


