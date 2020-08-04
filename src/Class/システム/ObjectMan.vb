'20160309霧島

Public Class CObjectMan
    Inherits HimTools2012.TargetSystem.CObjectManSK

    Public Sub New()
    End Sub

    Public Overrides Function GetObject(sKey As String, Optional pDB As HimTools2012.Data.CommonDataBase = Nothing) As HimTools2012.TargetSystem.CTargetObjectBase
        Try
            If ObjectCollection.ContainsKey(sKey) Then
                Return ObjectCollection.Item(sKey)
            End If
        Catch ex As Exception

        End Try

        Dim sHeader As String = Strings.Left(sKey, InStr(sKey, ".") - 1)
        Dim nCode As System.Int64 = Val(Strings.Mid(sKey, InStr(sKey, ".") + 1))

        Select Case sHeader
            Case "農地" : Return GetObjSub(App農地基本台帳.TBL農地.FindRowByID(nCode), GetType(CObj農地), False)
            Case "個人" : Return GetObjSub(App農地基本台帳.TBL個人.FindRowByID(nCode), GetType(CObj個人), False)
            Case "申請" : Return GetObjSub(App農地基本台帳.TBL申請.FindRowByID(nCode), GetType(CObj申請), False)
            Case "農家" : Return GetObjSub(App農地基本台帳.TBL世帯.FindRowByID(nCode), GetType(CObj農家), False)

            Case "営農情報" : Return New CObj営農情報(App農地基本台帳.TBL世帯営農.Rows.Find(nCode), False)

            Case "申請管理" : Return New CObj各種(sKey)
            Case "削除農地" : Return New CObj各種削除(sKey)
            Case "削除個人" : Return New CObj各種削除(sKey)
            Case "削除農家" : Return New CObj各種削除(sKey)

            Case "農地法終期台帳" : Return New CObj各種(sKey)
            Case "利用権終期台帳" : Return New CObj各種(sKey)
            Case "利用権終期年毎" : Return New CObj各種(sKey)
            Case "農地法終期年毎" : Return New CObj各種(sKey)
            Case "受付中" : Return New CObj各種(sKey)
            Case "利用権設定承認済" : Return New CObj各種(sKey)
            Case "利用権移転承認済" : Return New CObj各種(sKey)
            Case "マスタ管理" : Return New CObj各種(sKey)
            Case "農業改善計画認定" : Return New CObj各種(sKey)

            Case "転用農地" : Return GetObjSub(App農地基本台帳.TBL転用農地.FindRowByID(nCode), GetType(CObj転用農地), False)
            Case "固定土地" : Return New CObj各種(App農地基本台帳.TBL固定情報.FindRowByID(nCode), "固定土地", nCode, False)
            Case "行政区" : Return New CObj各種(sKey)
            Case "土地履歴" : Return New CObj土地履歴(App農地基本台帳.TBL土地履歴.Rows.Find(nCode), False)
            Case "不在者管理" : Return New CObj各種(sKey)

            Case "認定・担い手管理" : Return New CObj各種(sKey)

            Case "市内農地管理" : Return New CObj各種(sKey)
            Case "審査中" : Return New CObj各種(sKey)

            Case "市内世帯管理" : Return New CObj各種(sKey)
            Case "市外管理", "市内管理" : Return New CObj各種(sKey)
            Case "市外農地管理" : Return New CObj各種(sKey)
            Case "家族経営協定管理" : Return New CObj各種(sKey)
            Case "パソコン研修管理" : Return New CObj各種(sKey)
            Case "農業法人研修希望" : Return New CObj各種(sKey)
            Case "あっせん希望管理" : Return New CObj各種(sKey)
            Case "農地利用変更" : Return New CObj各種(sKey)
            Case "非農地証明願い" : Return New CObj各種(sKey)
            Case "農地改良届" : Return New CObj各種(sKey)

            Case "農用地利用計画変更" : Return New CObj各種(sKey)
            Case "地域農地" : Return New CObj各種(sKey)
            Case "年金受給者" : Return New CObj各種(sKey)
            Case "年金加入者" : Return New CObj各種(sKey)
            Case "農地台帳" : Return New CObj各種(sKey)
            Case "集計機能" : Return New CObj各種(sKey)
            Case "地目別集計" : Return New CObj各種(sKey)

            Case "農家の取り組み管理" : Return New CObj各種(sKey)
            Case "買受希望者リスト" : Return New CObj各種(sKey)
            Case "買受希望農地リスト" : Return New CObj各種(sKey)
            Case "貸付希望農地リスト" : Return New CObj各種(sKey)
            Case "売渡希望農地リスト" : Return New CObj各種(sKey)
            Case "経営移譲年金" : Return New CObj各種(sKey)
            Case "あっせん申出渡受付中" : Return New CObj各種(sKey)
            Case "農業者老齢年金" : Return New CObj各種(sKey)
            Case "利用権設定始期期間別一覧" : Return New CObj各種(sKey)
            Case "農振地内異動済み農地集計" : Return New CObj各種(sKey)
            Case "農振農用地内一覧" : Return New CObj各種(sKey)
            Case "事業計画変更" : Return New CObj各種(sKey)
            Case "公共機関管理" : Return New CObj各種(sKey)
            Case "公共機関" : Return New CObj各種(sKey)
            Case "住基" : Return New CObj各種(sKey)
            Case "中間管理機構" : Return New CObj各種(sKey)
            Case "中間管理機構を介した貸借農地" : Return New CObj各種(sKey)
            Case "一時転用終期台帳" : Return New CObj各種(sKey)
            Case "農業者年金一覧" : Return New CObj各種(sKey)
            Case "経営移譲年金" : Return New CObj各種(sKey)
            Case "農業者老齢年金" : Return New CObj各種(sKey)
            Case "年金対象者一覧" : Return New CObj各種(sKey)
            Case "農業委員管理" : Return New CObj各種(sKey)
            Case "期別農業委員会" : Return New CObj各種(sKey)
            Case "農委地目" : Return New CObj各種(sKey)
            Case "取消し" : Return New CObj各種(sKey)
            Case "取下げ" : Return New CObj各種(sKey)
            Case "農業委員" : Return New CObj各種(sKey)
            Case "利用権設定承認済過年度" : Return Nothing
            Case "３条許可済過年度" : Return New CObj各種(sKey)
            Case "３条１項処理済過年度" : Return New CObj各種(sKey)
            Case "特例対象農地" : Return New CObj各種(sKey)
            Case Else
                CasePrint(sHeader)
                Return Nothing
        End Select
    End Function

    Public Overrides Function GetObjectDB(sKeyHeader As String, ByRef mvarRow As System.Data.DataRow, pType As System.Type, Optional bAddCollection As Boolean = False) As HimTools2012.TargetSystem.CTargetObjectBase
        Dim sKey As String = sKeyHeader & "." & mvarRow.Item("ID")
        If sKeyHeader.EndsWith(".") Then
            sKey = sKeyHeader & mvarRow.Item("ID")
        End If

        Try
            If ObjectCollection.ContainsKey(sKey) Then
                Return ObjectCollection.Item(sKey)
            End If
        Catch ex As Exception

        End Try

        Try
            Dim pCLASSBASEType As Type = Type.GetType(pType.FullName)
            Dim pObj As HimTools2012.TargetSystem.CTargetObjectBase = Activator.CreateInstance(pCLASSBASEType, {mvarRow, False})

            If bAddCollection Then
                ObjectCollection.Add(sKey, pObj)
            End If

            Return pObj
        Catch ex As Exception
            Return Nothing
        End Try
    End Function



End Class

