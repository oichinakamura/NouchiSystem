Imports HimTools2012
Imports HimTools2012.TargetSystem


Public Class C申請追加
    Inherits HimTools2012.Data.UpdateRow

    Private n中間管ID As Decimal = 0

    Public Sub New(ByVal pRow As DataRow, ByVal n法令 As enum法令, ByVal s農地リスト As String, ByVal pUPDateMode As HimTools2012.Data.UPDateMode, Optional n状態 As Integer = 0)
        MyBase.New(pRow, pUPDateMode)
        SetValue("法令", System.Convert.ToInt32(n法令))
        SetValue("状態", n状態)
        SetValue("農地区分", 0)
        SetValue("農振区分", 0)
        SetValue("現地調査区分", 1)
        SetValue("農地リスト", s農地リスト)
        SetValue("代理人A", 0)
        SetValue("経由法人ID", 0)
        SetValue("更新日", Now)
        n中間管ID = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
    End Sub

    Public Sub Set受付情報(ByRef p入力支援 As C申請入力支援)
        SetValue("受付年月日", p入力支援.受付年月日)
        SetValue("受付番号", p入力支援.受付番号)
        SetValue("通年受付番号", p入力支援.受付通年番号)
    End Sub

    Public Sub Set申請者A(ByVal p農家 As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s名称 As String)
        If p農家 IsNot Nothing Then
            Select Case p農家.Key.DataClass
                Case "農家"
                    SetValue("申請世帯A", p農家.ID)
                    SetValue("申請者A", p農家.GetProperty("世帯主ID"))
                    SetValue("氏名A", p農家.GetProperty("世帯主名").ToString)
                    SetValue("職業A", p農家.GetProperty("世帯主職業").ToString)
                    SetValue("住所A", p農家.GetProperty("住所").ToString)
                    SetValue("集落A", p農家.GetProperty("世帯主集落").ToString)
                    SetValue("年齢A", Val(p農家.GetProperty("年齢")))
                Case "個人"
                    SetValue("氏名A", p農家.GetProperty("氏名"))
                    SetValue("申請者A", p農家.ID)

                    If n中間管ID <> 0 AndAlso p農家.ID = n中間管ID Then
                        SetValue("経営面積A", 0)
                    Else
                        SetValue("申請世帯A", Val(p農家.GetProperty("世帯ID").ToString))

                        SetValue("経営面積A", Val(p農家.GetProperty("経営面積")))
                        SetValue("職業A", p農家.GetStringValue("職業"))
                        SetValue("住所A", p農家.GetProperty("住所"))
                        SetValue("集落A", p農家.GetProperty("集落名"))
                        SetValue("年齢A", Val(p農家.GetProperty("年齢")))
                    End If
                Case Else
                    MsgBox("キー「" & p農家.Key.DataClass & "」に対応していません。システム開発会社に問合せをお願いします。")
                    Stop
            End Select
            SetValue("名称", String.Format(s名称, GetValue("氏名A")))
        End If
    End Sub
                     
    Public Sub Set申請者B(ByVal p受手農家 As CTargetObjWithView)
        If p受手農家 IsNot Nothing Then
            With p受手農家
                Select Case .Key.DataClass
                    Case "農家"
                        If GetValue("名称").ToString.EndsWith("→") Then
                            SetValue("名称", GetValue("名称") & .GetProperty("世帯主名"))
                        End If
                        SetValue("申請世帯B", Val(.GetProperty("ID")))
                        SetValue("申請者B", Val(.GetProperty("世帯主ID")))
                        SetValue("氏名B", .GetProperty("世帯主名").ToString)
                        SetValue("職業B", .GetProperty("世帯主職業").ToString)
                        SetValue("住所B", .GetProperty("住所").ToString)
                        SetValue("集落B", .GetProperty("世帯主集落").ToString)
                        SetValue("世帯員数B", Val(.GetProperty("世帯員数")))
                        SetValue("年齢B", Val(.GetProperty("年齢")))
                        SetValue("経営面積B", Val(.GetProperty("経営面積")))
                        SetValue("借入面積B", Val(.GetProperty("借入面積")))

                        If GetValue("経営面積B") = 0 Then SetValue("現地調査区分", 2)

                    Case "個人"
                        If GetValue("名称").ToString.EndsWith("→") Then
                            SetValue("名称", GetValue("名称") & .GetProperty("氏名"))
                        End If

                        SetValue("申請者B", .ID)
                        SetValue("申請世帯B", .GetProperty("世帯ID"))
                        SetValue("氏名B", .GetProperty("氏名").ToString)
                        SetValue("住所B", .GetProperty("住所").ToString)

                        If n中間管ID <> 0 AndAlso .ID = n中間管ID Then
                        Else
                            SetValue("職業B", .GetProperty("職業").ToString)
                            SetValue("集落B", .GetProperty("集落名").ToString)
                            SetValue("世帯員数B", Val(.GetProperty("世帯員数")))
                            SetValue("年齢B", Val(.GetProperty("年齢")))

                            SetValue("経営面積B", Val(.GetProperty("経営面積").ToString))
                            SetValue("借入面積B", Val(.GetProperty("借入面積").ToString))
                        End If
                    Case Else
                        MsgBox("キー「" & p受手農家.Key.DataClass & "」に対応していません。システム開発会社に問合せをお願いします。")
                        Stop
                End Select
            End With
        End If
    End Sub

    Public Sub Set申請者C(ByVal p農家 As CTargetObjWithView)
        If p農家 IsNot Nothing Then
            Select Case p農家.Key.DataClass
                Case "農家"
                    SetValue("申請世帯C", p農家.ID)
                    SetValue("申請者C", p農家.GetProperty("世帯主ID"))
                    SetValue("氏名A", p農家.GetProperty("世帯主名").ToString)
                    SetValue("職業A", p農家.GetProperty("世帯主職業").ToString)
                    SetValue("住所A", p農家.GetProperty("住所").ToString)
                    SetValue("集落A", p農家.GetProperty("世帯主集落").ToString)
                    SetValue("年齢A", Val(p農家.GetProperty("年齢").ToString))
                Case "個人"
                    SetValue("申請世帯C", p農家.GetProperty("世帯ID"))
                    SetValue("申請者C", p農家.ID)
                    SetValue("氏名C", p農家.GetProperty("氏名"))
                    SetValue("職業C", p農家.GetProperty("職業"))
                    SetValue("経営面積C", p農家.GetProperty("経営面積"))
                    SetValue("住所C", p農家.GetProperty("住所"))
                    SetValue("集落C", p農家.GetProperty("集落名"))
                    SetValue("年齢C", p農家.GetProperty("年齢"))
                Case Else
                    MsgBox("キー「" & p農家.Key.DataClass & "」に対応していません。システム開発会社に問合せをお願いします。")
                    Stop
            End Select

        End If
    End Sub

    Public Function InsertInto(ByVal bOpenWindow As Boolean) As Boolean
        With SysAD.DB(sLRDB).GetInsertRecordByID("D_申請", Me, True)
            If .Success Then
                Body.Item("ID") = .Value
                App農地基本台帳.TBL申請.Rows.Add(Body)
                Return Open申請Wnd(bOpenWindow, Body)
            Else
                MsgBox("データの追加に失敗しました。:" & .Value)
                Return False
            End If
        End With
    End Function

End Class
