Module mod農地
    Public Enum 土地異動事由
        非農地決定_通知済 = 261
        農地の削除 = 891
        農地法第4条による異動 = 10040
        樹園地へ変換 = 997
        合併世帯に伴う管理世帯変更 = 10008
        合併世帯に伴う借受世帯変更 = 10009
        解約 = 10020
        農地法３条による所有権移転 = 10030
        農地法３条による貸借 = 10031
        農地法県３条による所有権移転 = 10032
        農地法県３条による貸借 = 10033
        農地法４条による転用 = 10040
        農地法５条による転用 = 10050
        基盤強化法にて所有権移転 = 10060
        基盤強化法による使用貸借設定 = 10061
        基盤強化法にて賃貸借設定 = 10062
        第18条解約 = 10200
        利用権の期間満了に伴う貸借終了 = 10201
        解約等の理由による貸借終了 = 10202
        合意解約 = 10210
        期間満了に伴う貸借終了 = 18000
        非非農地証明願の適用 = 18099
        土地の名称変更 = 20115
        台帳修正_貸借 = 99784
        分筆登記 = 99984
        相続 = 99996
        職権による所有権移転 = 99997
        時効取得 = 99998
        その他 = 100000
        非農地証明 = 100001
        農地法３条の３第１項届出 = 100002
        農地の現況報告 = 100003
        その他非農地認定 = 100004
        第３条不許可 = 100005
        許可取消 = 100006
        合筆登記 = 100007
        農地転用事業計画変更 = 100008
        第２号仮登記 = 100009
        許可不要の転用 = 100010
        農地法５１条 = 100011
        市民農園等 = 100012
        換地処理削除 = 844
        換地処理追加 = 845
        一部現況分割 = 97
        一部現況結合 = 98
    End Enum

    '/***中村さんが作成したものです***/
    Public Sub Make農地履歴BK(ByVal nID As Long, ByVal UpDt As Date, ByVal dtDate As Date, ByVal n異動事由 As Long, ByVal s内容 As String, Optional ByVal 申請者A As Long = 0, Optional ByVal 申請者B As Long = 0, Optional ByVal 申請ID As Long = 0, Optional ByVal sDoSQL As String = "")
        Dim sSQL As String = ""
        sSQL = "INSERT INTO D_土地履歴 ( LID, 更新日, 異動日, 入力日, 異動事由, 内容,関係者A,関係者B,申請ID ) "
        sSQL = sSQL & "VALUES(" & nID & ",CDate('" & UpDt & "'),CDate('" & dtDate & "'),CDate('" & UpDt & "')," & n異動事由 & ",'" & s内容 & "'," & 申請者A & "," & 申請者B & "," & 申請ID & ")"
        SysAD.DB(sLRDB).ExecuteSQL(sSQL)
    End Sub

    '/***今回の追加フィールド分も書き込むようにしたものです***/
    Public Sub Make農地履歴(ByVal nID As Long, ByVal UpDt As Date, ByVal dtDate As Date, ByVal n異動事由 As Long, ByVal n法令 As enum法令, ByVal s内容 As String, Optional ByVal 申請者A As Long = 0, Optional ByVal 申請者B As Long = 0, Optional ByVal 申請ID As Long = 0, Optional ByVal sDoSQL As String = "")
        Dim sSQL As String = ""

        Dim pRow As DataRow = App農地基本台帳.TBL農地.FindRowByID(nID)
        If pRow Is Nothing Then
            pRow = App農地基本台帳.TBL転用農地.FindRowByID(nID)
        End If

        If pRow IsNot Nothing Then
            sSQL = "INSERT INTO D_土地履歴 ( LID, 更新日, 異動日, 異動事由, 内容, 関係者A, 関係者B, 申請ID, 申請時大字ID, 申請時小字ID, 申請時所在, 申請時地番, 申請時一部現況, 申請時登記簿地目, 申請時登記簿面積, 申請時共有持分分子, 申請時共有持分分母, 申請時現況地目, 申請時実面積, 申請時本地面積, 申請時部分面積, 申請時田面積, 申請時畑面積, 申請時樹園地, 申請時採草放牧面積, 申請時農地状況, 申請時所有世帯ID, 申請時所有者ID, 申請時所有者氏名, 申請時管理世帯ID, 申請時管理者ID, 申請時自小作別, 申請時借受世帯ID, 申請時借受人ID, 申請時借受者氏名, 申請時農業生産法人経由貸借, 申請時経由農業生産法人ID, 申請時小作地適用法, 申請時小作形態, 申請時小作料, 申請時小作料単位, 申請時10a賃借料, 申請時小作開始年月日, 申請時小作終了年月日, 申請時利用状況調査日, 申請時利用状況調査農地法, 申請時利用状況調査荒廃, 申請時利用意向調査日, 申請時利用意向根拠条項, 申請時利用意向意向内容区分, 入力日) "
            sSQL = sSQL & String.Format("VALUES({0},CDate('{1}'),CDate('{2}'),{3},'{4}',{5},{6},{7},{8}, {9}, '{10}', '{11}', {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, '{28}', {29}, {30}, {31}, {32}, {33}, '{34}', {35}, {36}, {37}, {38}, '{39}', '{40}', {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49},CDate('{1}'))", _
                                        nID, UpDt, dtDate, n異動事由, s内容, 申請者A, 申請者B, 申請ID, Val(pRow.Item("大字ID").ToString), Val(pRow.Item("小字ID").ToString), pRow.Item("所在").ToString, pRow.Item("地番").ToString, Val(pRow.Item("一部現況").ToString), Val(pRow.Item("登記簿地目").ToString), Val(pRow.Item("登記簿面積").ToString), Val(pRow.Item("共有持分分子").ToString), Val(pRow.Item("共有持分分母").ToString), Val(pRow.Item("現況地目").ToString), Val(pRow.Item("実面積").ToString), Val(pRow.Item("本地面積").ToString), Val(pRow.Item("部分面積").ToString), Val(pRow.Item("田面積").ToString), Val(pRow.Item("畑面積").ToString), Val(pRow.Item("樹園地").ToString), Val(pRow.Item("採草放牧面積").ToString), Val(pRow.Item("農地状況").ToString), Val(pRow.Item("所有世帯ID").ToString), Val(pRow.Item("所有者ID").ToString), pRow.Item("所有者氏名").ToString, Val(pRow.Item("管理世帯ID").ToString), Val(pRow.Item("管理者ID").ToString), Val(pRow.Item("自小作別").ToString), Val(pRow.Item("借受世帯ID").ToString), Val(pRow.Item("借受人ID").ToString), pRow.Item("借受人氏名").ToString, pRow.Item("農業生産法人経由貸借"), Val(pRow.Item("経由農業生産法人ID").ToString), Val(pRow.Item("小作地適用法").ToString), Val(pRow.Item("小作形態").ToString), pRow.Item("小作料").ToString, pRow.Item("小作料単位").ToString, Val(pRow.Item("10a賃借料").ToString), CnvDate(pRow.Item("小作開始年月日")), CnvDate(pRow.Item("小作終了年月日")), CnvDate(pRow.Item("利用状況調査日")), Val(pRow.Item("利用状況調査農地法").ToString), Val(pRow.Item("利用状況調査荒廃").ToString), CnvDate(pRow.Item("利用意向調査日")), Val(pRow.Item("利用意向根拠条項").ToString), Val(pRow.Item("利用意向意向内容区分").ToString))
            SysAD.DB(sLRDB).ExecuteSQL(sSQL)
        End If
    End Sub

    Private Function CnvDate(ByVal pValue As Object)
        If pValue IsNot Nothing AndAlso IsDate(pValue) Then
            Return "#" & pValue & "#"
        Else
            Return "Null"
        End If
    End Function
End Module
