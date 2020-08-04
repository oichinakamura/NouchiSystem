
Imports System.ComponentModel
Imports HimTools2012.CommonFunc
Imports HimTools2012.TypeConverterCustom

Public Class C各種申請管理
    Inherits CObj各種
    Public Sub New(ByVal sKey As String)
        MyBase.New(sKey)
    End Sub
    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand ' & "-" & Me.Key.DataClass
            Case "開く", "一覧"
                Select Case Me.Key.DataClass
                    '受付中
                    Case "受付中"
                    Case "３条受付中" : Open申請List("３条受付中.0", "農地法3条受付中", "[法令] In (30,31,32,33) AND [状態]=0")
                    Case "農地法3条1項の届出" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (311) AND [状態]=0")
                    Case "４条受付中" : Open申請List("４条受付中.0", "農地法4条受付中", "[法令] In (40,42) AND [状態]=0")
                    Case "５条受付中" : Open申請List("５条受付中.0", "農地法5条受付中", "[法令] In (50,51,52) AND [状態]=0")
                    Case "18条解約受付中" : Open申請List("申請受付一覧.180", "18条解約受付中", "[法令] In (18,20,180,200,250) AND [状態]=0")
                    Case "所有権移転受付中" : Open申請List("申請受付一覧.62", "基盤強化法所有権受付中", "[法令] In (60) AND [状態]=0")
                    Case "利用権設定受付中" : Open申請List("申請受付一覧.61", "基盤強化法利用権設定受付中", "[法令] In (61) AND [状態]=0")
                    Case "利用権移転受付中" : Open申請List("申請受付一覧.60", "基盤強化法利用権移転受付中", "[法令] In (62) AND [状態]=0")
                    Case "合意解約受付中" : Open申請List("申請受付一覧.210", "合意解約受付中", "[法令] In (210, 250) AND [状態]=0")
                    Case "あっせん申出渡受付中" : Open申請List("申請受付一覧." & enum法令.あっせん出手, "あっせん申出(渡)受付中", "[法令] In (400) AND [状態]=0")
                    Case "あっせん申出受受付中" : Open申請List("申請受付一覧." & enum法令.あっせん受手, "あっせん申出(受)受付中", "[法令] In (401) AND [状態]=0")
                    Case "農地改良届", "農地改良届受付中" : Open申請List("農地改良届.60", "農地改良届受付中", "[法令] In (301) AND [状態]=0")
                    Case "農地利用目的変更受付中" : Open申請List("農地利用目的変更.0", "農地利用目的変更", String.Format("[法令] In ({0}) AND [状態]=0", System.Convert.ToInt32(enum法令.農地利用目的変更)))
                    Case "農用地利用計画変更", "農用地利用計画変更受付中" : Open申請List("農用地利用計画変更.0", "農用地利用計画変更受付中", "[法令] In (302) AND [状態]=0")
                    Case "非農地証明願い受付中" : Open申請List("非農地証明願い.0", "非農地証明願い受付中", "[法令] In (600,602) AND [状態]=0")
                    Case "事業計画変更受付中" : Open申請List("事業計画変更.0", "事業計画変更", String.Format("[法令] In ({0}) AND [状態]=0", System.Convert.ToInt32(enum法令.事業計画変更)))
                    Case "買受適格-耕作-公売", "買受適格-耕作-公売受付中" : Open申請List("買受適格-耕作-公売.0", "買受適格-耕作-公売受付中", "[法令] In (801) AND [状態]=0")
                    Case "買受適格-耕作-競売", "買受適格-耕作-競売受付中" : Open申請List("買受適格-耕作-競売.0", "買受適格-耕作-競売受付中", "[法令] In (802) AND [状態]=0")
                    Case "買受適格-転用-公売", "買受適格-転用-公売受付中" : Open申請List("買受適格-転用-公売.0", "買受適格-転用-公売受付中", "[法令] In (803) AND [状態]=0")
                    Case "買受適格-転用-競売", "買受適格-転用-競売受付中" : Open申請List("買受適格-転用-競売.0", "買受適格-転用-競売受付中", "[法令] In (804) AND [状態]=0")
                    '↓要確認
                    Case "農地改良届受付中" : Open申請List("農地改良届.0", "農地改良届受付中", "[法令] In (301) AND [状態]=0")
                    Case "農地利用目的変更受付中" : Open申請List("農地利用目的変更.0", "農地利用目的変更", String.Format("[法令] In ({0}) AND [状態]=0", System.Convert.ToInt32(enum法令.農地利用目的変更)))
                    Case "農用地利用計画変更受付中" : Open申請List("農用地利用計画変更.0", "農用地利用計画変更", String.Format("[法令] In ({0}) AND [状態]=0", System.Convert.ToInt32(enum法令.農用地計画変更)))
                    Case "事業計画変更受付中" : Open申請List("事業計画変更.0", "事業計画変更", String.Format("[法令] In ({0}) AND [状態]=0", System.Convert.ToInt32(enum法令.事業計画変更)))
                        '審査
                    Case "３条審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (30,31,32,33) AND [状態]=1")
                    Case "４条審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (40,42) AND [状態]=1")
                    Case "５条審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態]=1")

                    Case "所有権移転審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (60) AND [状態]=1")
                    Case "利用権設定審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (61) AND [状態]=1")
                    Case "利用権移転審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (62) AND [状態]=1")

                    Case "あっせん申出渡審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (400) AND [状態]=1")
                    Case "あっせん申出受審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (401) AND [状態]=1")
                        
                    Case "農地利用目的変更審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (500) AND [状態]=1")
                    Case "農用地利用計画変更審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (302) AND [状態]=1")
                    Case "事業計画変更審査中" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (303) AND [状態]=1")
                        '許可
                    Case "３条許可済"
                    Case "３条許可済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (30,31,32,33) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "３条許可済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (30,31,32,33) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "３条１項処理済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (311) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "３条１項処理済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (311) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "４条許可済"
                    Case "４条許可済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (40,42) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "４条許可済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (40,42) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "５条許可済" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態]=2 AND [許可年月日] Is Null")
                    Case "５条許可済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "５条許可済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "18条承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (18,20,180,200,250) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "18条承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (18,20,180,200,250) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))

                    Case "所有権移転承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (60) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "所有権移転承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (60) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "利用権設定承認済"
                    Case "利用権設定承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (61) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "利用権設定承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (61) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "利用権移転承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (62) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "利用権移転承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (62) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "合意解約承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (210,250) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "合意解約承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (210,250) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))

                    Case "あっせん申出渡承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (400) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "あっせん申出渡承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (400) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "あっせん申出受承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (401) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "あっせん申出受承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (401) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))

                    Case "農地改良届承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (301) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "農地改良届承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (301) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "農地利用目的変更承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (500) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "農地利用目的変更承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (500) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "農用地利用計画変更承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (302) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "農用地利用計画変更承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (302) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))

                    Case "非農地証明願済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (602) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "非農地証明願済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (602) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))

                    Case "事業計画変更承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (303) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "事業計画変更承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (303) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))

                    Case "買受適格-耕作-公売承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (801) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "買受適格-耕作-公売承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (801) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "買受適格-耕作-競売承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (802) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "買受適格-耕作-競売承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (802) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "買受適格-転用-公売承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (803) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "買受適格-転用-公売承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (803) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))
                    Case "買受適格-転用-競売承認済年別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (804) AND [状態]=2 AND " & GetBitweenYearStr("許可年月日", Math.Floor(Me.Key.ID / 100)))
                    Case "買受適格-転用-競売承認済月別" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (804) AND [状態]=2 AND " & GetBitweenMonthStr("許可年月日", Year(Now), (Me.Key.ID Mod 100)))

                        '取下げ
                    Case "３条取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (30,31,32,33) AND [状態] IN (4,3)")
                    Case "３条１項取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (311) AND [状態] IN (4,3)")
                    Case "４条取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (40,42) AND [状態] IN (4,3)")
                    Case "５条取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態] IN (4,3)")
                    Case "18条解約取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (18,20,180,200,250) AND [状態] IN (4,3)")
                    Case "所有権移転取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (60) AND [状態] IN (4,3)")
                    Case "利用権設定取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (61) AND [状態] IN (4,3)")
                    Case "利用権移転取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (62) AND [状態] IN (4,3)")
                    Case "合意解約取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (210) AND [状態] IN (4,3)")
                    Case "あっせん申出渡取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (400) AND [状態] IN (4,3)")
                    Case "あっせん申出受取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (401) AND [状態] IN (4,3)")
                    Case "農地改良届取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (301) AND [状態] IN (4,3)")
                    Case "農地利用目的変更取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (500) AND [状態] IN (4,3)")
                    Case "農用地利用計画変更取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (302) AND [状態] IN (4,3)")
                    Case "非農地証明願い取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (600,602) AND [状態] IN (4,3)")
                    Case "事業計画変更取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (303) AND [状態] IN (4,3)")
                    Case "買受適格-耕作-公売取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (801) AND [状態] IN (4,3)")
                    Case "買受適格-耕作-競売取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (802) AND [状態] IN (4,3)")
                    Case "買受適格-転用-公売取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (803) AND [状態] IN (4,3)")
                    Case "買受適格-転用-公売取下げ" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (804) AND [状態] IN (4,3)")
                        '取消し
                    Case "３条取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (30,31,32,33) AND [状態] IN (5)")
                    Case "３条１項取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (311) AND [状態] IN (5)")
                    Case "４条取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (40,42) AND [状態] IN (5)")
                    Case "５条取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態] IN (5)")
                    Case "18条解約取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (18,20,180,200,250) AND [状態] IN (5)")
                    Case "所有権移転取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (60) AND [状態] IN (5)")
                    Case "利用権設定取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (61) AND [状態] IN (5)")
                    Case "利用権移転取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (62) AND [状態] IN (5)")
                    Case "合意解約取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (210) AND [状態] IN (5)")
                    Case "あっせん申出渡取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (400) AND [状態] IN (5)")
                    Case "あっせん申出受取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (401) AND [状態] IN (5)")
                    Case "農地改良届取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (301) AND [状態] IN (5)")
                    Case "農地利用目的変更取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (500) AND [状態] IN (5)")
                    Case "農用地利用計画変更取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (302) AND [状態] IN (5)")
                    Case "非農地証明願い取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (600,602) AND [状態] IN (5)")
                    Case "事業計画変更取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (303) AND [状態] IN (5)")
                    Case "買受適格-耕作-公売取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (801) AND [状態] IN (5)")
                    Case "買受適格-耕作-競売取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (802) AND [状態] IN (5)")
                    Case "買受適格-転用-公売取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (803) AND [状態] IN (5)")
                    Case "買受適格-転用-公売取消し" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (804) AND [状態] IN (5)")
                        '/*不許可*/
                    Case "３条不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (30,31,32,33) AND [状態] IN (42)")
                    Case "農地法3条1項不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (311) AND [状態] IN (42)")
                    Case "４条不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (40,42) AND [状態] IN (42)")
                    Case "５条不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態] IN (42)")
                    Case "18条解約不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (18,20,180,200,250) AND [状態] IN (42)")
                    Case "所有権移転不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (60) AND [状態] IN (42)")
                    Case "利用権設定不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (61) AND [状態] IN (42)")
                    Case "利用権移転不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (62) AND [状態] IN (42)")
                    Case "合意解約不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (210) AND [状態] IN (42)")
                    Case "あっせん申出渡不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (400) AND [状態] IN (42)")
                    Case "あっせん申出受不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (401) AND [状態] IN (42)")
                    Case "農地改良届不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (301) AND [状態] IN (42)")
                    Case "農地利用目的変更不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (500) AND [状態] IN (42)")
                    Case "農用地利用計画変更不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (302) AND [状態] IN (42)")
                    Case "非農地証明願い不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (600,602) AND [状態] IN (42)")
                    Case "事業計画変更不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (303) AND [状態] IN (42)")
                    Case "買受適格-耕作-公売不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (801) AND [状態] IN (42)")
                    Case "買受適格-耕作-競売不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (802) AND [状態] IN (42)")
                    Case "買受適格-転用-公売不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (803) AND [状態] IN (42)")
                    Case "買受適格-転用-公売不許可" : Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (804) AND [状態] IN (42)")

                    Case "現地調査表作成"
                        If Not SysAD.page農家世帯.中央Tab.ExistPage("現地調査表作成.0", True, GetType(CTabPage現地調査表作成)) Then
                        End If
                    Case Else
                        CasePrint(Me.Key.DataClass)
                End Select

            Case "受付簿"
            Case "3条受付中"
            Case "期間(許可日)を指定して抽出"
                Dim pObj As New 期間抽出条件()
                Dim pDlg As New HimTools2012.PropertyGridDialog(pObj, "抽出期間の入力")

                If pDlg.ShowDialog = DialogResult.OK Then
                    Select Case Me.Key.DataClass
                        Case "４条許可済"
                            Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (40,42) AND [状態]=2 AND " & pObj.ToSQLWhere区("許可年月日"))
                        Case "５条許可済"
                            Open申請List(Me.Key.KeyValue, Me.名称, "[法令] In (50,51,52) AND [状態]=2 AND " & pObj.ToSQLWhere区("許可年月日"))
                        Case Else
                            Stop
                    End Select
                End If
            Case Else
                CasePrint(sCommand & "-" & Me.Key.DataClass)
        End Select
        Return Nothing
    End Function

    <TypeConverter(GetType(PropertyOrderConverter))>
    Public Class 期間抽出条件
        Inherits HimTools2012.InputSupport.CInputSupport
        Public Sub New()
            MyBase.New(Nothing)

        End Sub

        <Category("期間～")>
        Public Property 許可日抽出開始日 As Date = Now.Date
        <Category("期間～")>
        Public Property 許可日抽出終了日 As Date = Now.Date

        Public Function ToSQLWhere区(ByVal sField As String) As String
            Dim sB As New List(Of String)
            If 許可日抽出終了日.Date > 許可日抽出開始日.Date Then
                sB.Add("[" & sField & "]>=#" & 許可日抽出開始日.Month & "/" & 許可日抽出開始日.Day & "/" & 許可日抽出開始日.Year & "#")
                sB.Add("[" & sField & "]<=#" & 許可日抽出終了日.Month & "/" & 許可日抽出終了日.Day & "/" & 許可日抽出終了日.Year & "#")

                Return Join(sB.ToArray, " AND ")
            Else
                MsgBox("指定された期間は不正です。", MsgBoxStyle.Critical)
                Return "[" & sField & "]=#" & Now.Month & "/" & Now.Day & "/" & Now.Year & "#"
            End If

        End Function
    End Class

    Public Overrides Function 名称() As String
        Select Case Me.Key.DataClass
            '/*許可・承認*/
            Case "３条許可済月" : Return "３条許可済月別" & Me.Key.ID
            Case "３条許可済年" : Return "３条許可済年別" & Me.Key.ID
            Case "３条許可済月別" : Return "３条許可済月別" & Me.Key.ID
            Case "３条許可済年別" : Return "３条許可済年別" & Me.Key.ID

            Case "３条１項処理済月別" : Return "３条１項処理済月別" & Me.Key.ID
            Case "３条１項処理済年別" : Return "３条１項処理済年別" & Me.Key.ID

            Case "４条許可済月" : Return "４条許可済月" & Me.Key.ID
            Case "４条許可済年" : Return "４条許可済年" & Me.Key.ID
            Case "４条許可済月別" : Return "４条許可済月別" & Me.Key.ID
            Case "４条許可済年別" : Return "４条許可済年別" & Me.Key.ID

            Case "５条許可済" : Return "５条許可済" & Me.Key.ID
            Case "５条許可済年" : Return "５条許可済年" & Me.Key.ID
            Case "５条許可済月" : Return "５条許可済月" & Me.Key.ID
            Case "５条許可済年別" : Return "５条許可済年別" & Me.Key.ID
            Case "５条許可済月別" : Return "５条許可済月別" & Me.Key.ID

            Case "18条承認済年別" : Return "18条承認済年別" & Me.Key.ID
            Case "18条承認済月別" : Return "18条承認済月別" & Me.Key.ID

            Case "所有権移転承認済年別" : Return "所有権移転承認済年別" & Me.Key.ID
            Case "所有権移転承認済月別" : Return "所有権移転承認済月別" & Me.Key.ID

            Case "利用権設定承認済月" : Return "利用権設定承認済月別" & Me.Key.ID
            Case "利用権設定承認済月別" : Return "利用権設定承認済月別" & Me.Key.ID
            Case "利用権設定承認済年別" : Return "利用権設定承認済年別" & Me.Key.ID

            Case "利用権移転承認済月別" : Return "利用権移転承認済月別" & Me.Key.ID
            Case "利用権移転承認済年別" : Return "利用権移転承認済年別" & Me.Key.ID

            Case "合意解約承認済年別" : Return "合意解約承認済年別" & Me.Key.ID
            Case "合意解約承認済月別" : Return "合意解約承認済月別" & Me.Key.ID

            Case "あっせん申出渡承認済年別" : Return "あっせん申出渡承認済年別" & Me.Key.ID
            Case "あっせん申出渡承認済月別" : Return "あっせん申出渡承認済月別" & Me.Key.ID
            Case "あっせん申出受承認済年別" : Return "あっせん申出受承認済年別" & Me.Key.ID
            Case "あっせん申出受承認済月別" : Return "あっせん申出受承認済月別" & Me.Key.ID

            Case "農地改良届承認済月別" : Return "農地改良届承認済月別" & Me.Key.ID
            Case "農地改良届承認済年別" : Return "農地改良届承認済年別" & Me.Key.ID

            Case "農地利用目的変更承認済月別" : Return "農地利用目的変更承認済月別" & Me.Key.ID
            Case "農地利用目的変更承認済年別" : Return "農地利用目的変更承認済年別" & Me.Key.ID

            Case "農用地利用計画変更承認済月別" : Return "農用地利用計画変更承認済月別" & Me.Key.ID
            Case "農用地利用計画変更承認済年別" : Return "農用地利用計画変更承認済年別" & Me.Key.ID

            Case "非農地証明願済年別" : Return "非農地証明願済年別" & Me.Key.ID
            Case "非農地証明願済月別" : Return "非農地証明願済月別" & Me.Key.ID

            Case "事業計画変更承認済年別" : Return "事業計画変更承認済年別" & Me.Key.ID
            Case "事業計画変更承認済月別" : Return "事業計画変更承認済月別" & Me.Key.ID

            Case "買受適格-耕作-公売承認済月別" : Return "買受適格-耕作-公売承認済月別" & Me.Key.ID
            Case "買受適格-耕作-公売承認済年別" : Return "買受適格-耕作-公売承認済年別" & Me.Key.ID
            Case "買受適格-耕作-競売承認済月別" : Return "買受適格-耕作-競売承認済月別" & Me.Key.ID
            Case "買受適格-耕作-競売承認済年別" : Return "買受適格-耕作-競売承認済年別" & Me.Key.ID
            Case "買受適格-転用-公売承認済月別" : Return "買受適格-転用-公売承認済月別" & Me.Key.ID
            Case "買受適格-転用-公売承認済年別" : Return "買受適格-転用-公売承認済年別" & Me.Key.ID
            Case "買受適格-転用-競売承認済月別" : Return "買受適格-転用-競売承認済月別" & Me.Key.ID
            Case "買受適格-転用-競売承認済年別" : Return "買受適格-転用-競売承認済年別" & Me.Key.ID

            Case Else
                CasePrint(Me.Key.DataClass, "Return ")
                Return Me.Key.DataClass
        End Select
    End Function

    Public Overloads Overrides Function CanDropKeyHead(sKey As String, sOption As String) As Boolean
        Select Case GetKeyHead(sKey) & "-" & Me.Key.DataClass
            Case "３条受付中-３条受付中" : Return False
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    CasePrint(GetKeyHead(sKey) & "-" & Me.Key.DataClass & """ : return false")
                    Return MyBase.CanDropKeyHead(sKey, sOption)
                End If

                Return False
        End Select
    End Function
    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuF As New HimTools2012.controls.ContextMenuEx(AddressOf ClickMenu)
        'ID クラスのないオブジェクトで失敗

        Select Case Me.Key.DataClass
            '/*受付中*/
            Case "３条受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "農地法3条1項の届出" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "４条受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "５条受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "18条解約受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "所有権移転受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "利用権設定受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "利用権移転受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "合意解約受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "あっせん申出渡受付中", "あっせん申出受受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "農地改良届受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "農地利用目的変更受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "農用地利用計画変更受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "非農地証明願い受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "事業計画変更受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
            Case "買受適格-耕作-公売受付中", "買受適格-耕作-競売受付中", "買受適格-転用-公売受付中", "買受適格-転用-競売受付中" : pMenuF.AddMenu("受付簿", , AddressOf ClickMenu)
                '/*許可・承認済*/
            Case "３条許可済" : pMenuF.AddMenu("開く", , AddressOf ClickMenu)
            Case "４条許可済" : pMenuF.AddMenu("開く,期間(許可日)を指定して抽出", , AddressOf ClickMenu)
            Case "５条許可済" : pMenuF.AddMenu("開く,期間(許可日)を指定して抽出", , AddressOf ClickMenu)
            Case Else
                CasePrint(GetKeyHead(Me.Key.DataClass), "return pMenuF.AddMenu(""受付簿"",, AddressOf ClickMenu)")
                Return MyBase.GetContextMenu(pMenu, , sParam)
        End Select
        If pMenuF IsNot Nothing Then
            Return pMenuF
        Else
            Return MyBase.GetContextMenu(pMenu, , sParam)
        End If
    End Function

    Public Shared Sub Open申請List(ByVal sKey As String, ByVal sName As String, ByVal sWhere As String)
        Dim pList As C申請リスト
        If Not SysAD.page農家世帯.TabPageContainKey(sKey) Then
            pList = New C申請リスト(SysAD.page農家世帯, sKey, sName)
            pList.Name = sKey
            SysAD.page農家世帯.中央Tab.AddPage(pList)
            pList.ImageKey = pList.IconKey
        Else
            pList = SysAD.page農家世帯.GetItem(sKey)
        End If

        pList.検索開始(sWhere, sWhere, "[受付補助記号],[受付番号]")
    End Sub
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return Nothing
        End Get
    End Property
    Public Overrides Function SaveMyself() As Boolean
        Return False
    End Function

End Class