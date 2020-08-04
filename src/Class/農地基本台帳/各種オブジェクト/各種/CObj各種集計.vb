'/20160229霧島
Imports HimTools2012.CommonFunc
'

Public Class CObj各種集計 : Inherits CObj各種
    Public Sub New(ByVal sKey As String)
        MyBase.New(sKey)
    End Sub
    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand & "-" & Me.Key.DataClass
            Case "開く-基本集計"
                OpenSQLList("基本集計", "SELECT IIf([所在] Is Null,'市町村管内','当該市町村外') AS 農地の所在, Sum(Sgn([田面積]))+Sum(Sgn([畑面積]))+Sum(Sgn([樹園地]))+Sum(Sgn([採草放牧面積])) AS 全筆数, Sum(Int([田面積]))+Sum(Int([畑面積]))+Sum(Int([樹園地]))+Sum(Int([採草放牧面積])) AS 面積計, Sum(Sgn([田面積])) AS 田筆数, Sum(Int([D:農地Info].[田面積])) AS 田面積の合計, Sum(Sgn([畑面積])) AS 畑筆数, Sum(Int([D:農地Info].[畑面積])) AS 畑面積の合計, Sum(Sgn([樹園地])) AS 樹園地筆数, Sum(Int([D:農地Info].[樹園地])) AS 樹園地面積, Sum(Sgn([採草放牧面積])) AS 採草放牧地筆数, Sum(Int([採草放牧面積])) AS 採草放牧地面積 FROM [D:農地Info] WHERE ((([D:農地Info].田面積)>0) AND (([D:農地Info].大字ID)<>-1) AND (([D:農地Info].農地状況)<=19)) OR ((([D:農地Info].畑面積)>0) AND (([D:農地Info].大字ID)<>-1) AND (([D:農地Info].農地状況)<=19)) OR ((([D:農地Info].樹園地)>0) AND (([D:農地Info].大字ID)<>-1) AND (([D:農地Info].農地状況)<=19)) OR ((([D:農地Info].採草放牧面積)>0) AND (([D:農地Info].大字ID)<>-1) AND (([D:農地Info].農地状況)<=19)) GROUP BY IIf([所在] Is Null,'市町村管内','当該市町村外');")
            Case "開く-農家世帯数"
                OpenSQLList("経営農家件数", "SELECT TBL経営農家.居住区分, Count(TBL経営農家.ID) AS 耕作農家の合計 FROM (SELECT IIf([住民区分]=0,'管内','その他') AS 居住区分, [D:世帯Info].ID FROM (V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID GROUP BY IIf([住民区分]=0,'管内','その他'), [D:世帯Info].ID)  AS TBL経営農家 GROUP BY TBL経営農家.居住区分;", {"耕作農家の合計"})
                '↓過去集計
                'OpenSQLList("経営農家件数", "SELECT DISTINCT IIf([住民区分]=0,'管内','その他') AS 居住区分, Count([D:世帯Info].ID) AS 農家件数 FROM [D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE ((([D:世帯Info].農地との関連)=True)) GROUP BY IIf([住民区分]=0,'管内','その他');", {"農家件数"})
            Case "開く-農業者数"
                OpenSQLList("経営農業者件数", "SELECT TBL経営個人.居住区分, Count(TBL経営個人.ID) AS 耕作者の合計 FROM (SELECT IIf([住民区分]=0,'管内','その他') AS 居住区分, [D:個人Info].ID FROM V_農地 INNER JOIN [D:個人Info] ON V_農地.耕作者ID = [D:個人Info].ID GROUP BY IIf([住民区分]=0,'管内','その他'), [D:個人Info].ID)  AS TBL経営個人 GROUP BY TBL経営個人.居住区分;", {"耕作者の合計"})
                '↓過去集計
                'OpenSQLList("経営農業者件数", "SELECT DISTINCT IIf([住民区分]=0,'管内','その他') AS 居住区分, Count([D:個人Info].ID) AS 農家人口 FROM [D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID WHERE ((([D:世帯Info].農地との関連)=True)) GROUP BY IIf([住民区分]=0,'管内','その他');", {"農家人口"})
            Case "開く-世帯経営面積一覧"
                OpenSQLList("世帯経営面積一覧", "SELECT '農家.' & [D:世帯Info].[ID] AS [KEY], [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].住所, V_行政区.ID, V_行政区.行政区, V_世帯別自作計.自作田, V_世帯別自作計.自作畑, V_世帯別自作計.自作樹, V_世帯別自作計.自作計, V_世帯別小作計.小作田, V_世帯別小作計.小作畑, V_世帯別小作計.小作樹, V_世帯別小作計.小作計, IIf(IsNull([自作田]),0,[自作田])+IIf(IsNull([小作田]),0,[小作田]) AS 経営田, IIf(IsNull([自作畑]),0,[自作畑])+IIf(IsNull([小作畑]),0,[小作畑]) AS 経営畑, IIf(IsNull([自作樹]),0,[自作樹])+IIf(IsNull([小作樹]),0,[小作樹]) AS 経営樹, IIf(IsNull([自作計]),0,[自作計])+IIf(IsNull([小作計]),0,[小作計]) AS 経営計, V_世帯別貸付計.貸付田, V_世帯別貸付計.貸付畑, V_世帯別貸付計.貸付樹, V_世帯別貸付計.貸付計 FROM (((([D:個人Info] INNER JOIN [D:世帯Info] ON [D:個人Info].ID = [D:世帯Info].世帯主ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) LEFT JOIN V_世帯別自作計 ON [D:世帯Info].ID = V_世帯別自作計.農家世帯ID) LEFT JOIN V_世帯別貸付計 ON [D:世帯Info].ID = V_世帯別貸付計.農家世帯ID) LEFT JOIN V_世帯別小作計 ON [D:世帯Info].ID = V_世帯別小作計.借受世帯ID WHERE ((([D:個人Info].行政区ID) Is Not Null) AND (([D:世帯Info].農地との関連)=True)) ORDER BY [D:個人Info].行政区ID;")
            Case "開く-農家別面積集計"
                OpenSQLList("世帯経営面積一覧", "SELECT D_担い手の農地利用集積状況.農家ID, D_担い手の農地利用集積状況.世帯主ID, D_担い手の農地利用集積状況.世帯主, D_担い手の農地利用集積状況.認定区分, Sum(D_担い手の農地利用集積状況.自作所有) AS 自作所有の合計, Sum(D_担い手の農地利用集積状況.自作所有うち田) AS 自作所有うち田の合計, Sum(D_担い手の農地利用集積状況.借入地) AS 借入地の合計, Sum(D_担い手の農地利用集積状況.借入地うち田) AS 借入地うち田の合計 FROM (SELECT [D:世帯Info].ID AS 農家ID, [D:世帯Info].世帯主ID, [D:個人Info].氏名 AS 世帯主, V_農業改善計画認定項目.名称 AS 認定区分, IIf([自小作別]=0,[実面積],IIf([借受世帯ID]=[世帯ID],[実面積],0)) AS 自作所有, IIf([自小作別]=0,[田面積],IIf([借受世帯ID]=[世帯ID],[田面積],0)) AS 自作所有うち田, 0 AS 借入地, 0 AS 借入地うち田 FROM (([D:農地Info] INNER JOIN [D:世帯Info] ON [D:農地Info].所有世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) LEFT JOIN V_農業改善計画認定項目 ON [D:個人Info].農業改善計画認定 = V_農業改善計画認定項目.ID UNION SELECT [D:世帯Info].ID AS 農家ID, [D:世帯Info].世帯主ID, [D:個人Info].氏名 AS 世帯主, V_農業改善計画認定項目.名称 AS 認定区分, 0 AS 自作所有, 0 AS 自作所有うち田, IIf([自小作別]<>0,IIf([所有世帯ID]<>[世帯ID],[実面積],0),0) AS 借入地, IIf([自小作別]<>0,IIf([所有世帯ID]<>[世帯ID],[田面積],0),0) AS 借入地うち田 FROM [D:農地Info] INNER JOIN (([D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) LEFT JOIN V_農業改善計画認定項目 ON [D:個人Info].農業改善計画認定 = V_農業改善計画認定項目.ID) ON [D:農地Info].借受世帯ID = [D:世帯Info].ID)  AS D_担い手の農地利用集積状況 GROUP BY D_担い手の農地利用集積状況.農家ID, D_担い手の農地利用集積状況.世帯主ID, D_担い手の農地利用集積状況.世帯主, D_担い手の農地利用集積状況.認定区分;")
            Case "開く-経営農家一覧"
                OpenSQLList("経営農家一覧", String.Format("SELECT IIf([住民区分]=0,'管内','その他') AS 居住区分, [D:個人Info].ID AS 耕作者ID, [D:個人Info].氏名 AS 耕作者名, [D:個人Info].住所 AS 耕作者住所, V_住民区分.名称 AS 住民区分名, Count(V_農地.ID) AS 経営筆数, Sum(V_農地.実面積) AS 経営面積, IIf(InStr([D:個人Info].[住所],'{0}')>0,'○','×') AS 市町村内判定 FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.耕作者ID = [D:個人Info].ID) LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID GROUP BY IIf([住民区分]=0,'管内','その他'), [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].住所, V_住民区分.名称, IIf(InStr([D:個人Info].[住所],'{0}')>0,'○','×');", SysAD.市町村.市町村名))
            Case "開く-60歳以上の農業従事者"
                OpenSQLList("60歳以上の農業従事者", "SELECT '世帯員.' & [D:個人Info].[ID] AS [Key], [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].住所, IIf([D:個人Info].[性別]=0,'男',IIf([D:個人Info].[性別]=1,'女','-')) AS 性別, [D:個人Info].生年月日, Val(IIf([生年月日]>0,IIf(DateAdd('yyyy',DateDiff('yyyy',[生年月日],Now()),[生年月日])>Now(),DateDiff('yyyy',[生年月日],Now())-1,DateDiff('yyyy',[生年月日],Now())),'-')) AS 年齢, Sum(Int([V_農地].[田面積])) AS 田面積, Sum(Int([V_農地].[畑面積])) AS 畑面積, Sum(Int([V_農地].[樹園地])) AS 樹園地, Sum(Int([V_農地].[採草放牧面積])) AS 採草放牧面積, Sum(Int([V_農地].[田面積])+Int([V_農地].[畑面積])+Int([V_農地].[樹園地])+Int([V_農地].[採草放牧面積])) AS 農地計 FROM [D:個人Info] INNER JOIN V_農地 ON [D:個人Info].ID = V_農地.耕作者ID WHERE (((Int([V_農地].[田面積])+Int([V_農地].[畑面積])+Int([V_農地].[樹園地])+Int([V_農地].[採草放牧面積]))>0)) GROUP BY '世帯員.' & [D:個人Info].[ID], [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].住所, IIf([D:個人Info].[性別]=0,'男',IIf([D:個人Info].[性別]=1,'女','-')), [D:個人Info].生年月日, Val(IIf([生年月日]>0,IIf(DateAdd('yyyy',DateDiff('yyyy',[生年月日],Now()),[生年月日])>Now(),DateDiff('yyyy',[生年月日],Now())-1,DateDiff('yyyy',[生年月日],Now())),'-')) HAVING (((Val(IIf([生年月日]>0,IIf(DateAdd('yyyy',DateDiff('yyyy',[生年月日],Now()),[生年月日])>Now(),DateDiff('yyyy',[生年月日],Now())-1,DateDiff('yyyy',[生年月日],Now())),'-')))>=60)) ORDER BY [D:個人Info].[フリガナ];")
            Case "開く-経営世帯一覧"
                OpenSQLList("経営世帯一覧", String.Format("SELECT IIf([住民区分]=0,'管内','その他') AS 居住区分, [D:世帯Info].ID AS 世帯ID, [D:個人Info].ID AS 世帯主ID, [D:個人Info].氏名 AS 世帯主名, [D:個人Info].住所 AS 世帯主住所, V_住民区分.名称 AS 住民区分名, Count(V_農地.ID) AS 経営筆数, Sum(V_農地.実面積) AS 経営面積, IIf(InStr([D:個人Info].[住所],'{0}')>0,'○','×') AS 市町村内判定 FROM ((V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID GROUP BY IIf([住民区分]=0,'管内','その他'), [D:世帯Info].ID, [D:個人Info].ID, [D:個人Info].氏名, [D:個人Info].住所, V_住民区分.名称, IIf(InStr([D:個人Info].[住所],'{0}')>0,'○','×');", SysAD.市町村.市町村名))
            Case "開く-地区別世帯件数集計"
                OpenSQLList("地区別世帯件数集計", "SELECT [%$##@_Alias].行政区ID, [%$##@_Alias].行政区, Max([%$##@_Alias].世帯数) AS 世帯数, Max([%$##@_Alias].田所有) AS 田所有, Max([%$##@_Alias].畑所有) AS 畑所有, Max([%$##@_Alias].樹所有) AS 樹所有 FROM (SELECT 所有世帯.行政区ID, V_行政区.行政区, Count(所有世帯.農家世帯ID) AS 世帯数, '' AS 田所有, '' AS 畑所有, '' AS 樹所有 FROM (SELECT V_農地.農家世帯ID, [D:個人Info].行政区ID FROM (V_農地 LEFT JOIN [D:世帯Info] ON V_農地.農家世帯ID = [D:世帯Info].ID) LEFT JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID GROUP BY V_農地.農家世帯ID, [D:個人Info].行政区ID)  AS 所有世帯 INNER JOIN V_行政区 ON 所有世帯.行政区ID = V_行政区.ID GROUP BY 所有世帯.行政区ID, V_行政区.行政区, '', '', '', '' UNION SELECT 所有世帯.行政区ID, V_行政区.行政区, '' AS 世帯数, Count(所有世帯.農家世帯ID) AS 田所有, '' AS 畑所有, '' AS 樹所有 FROM (SELECT V_農地.農家世帯ID, [D:個人Info].行政区ID FROM (V_農地 LEFT JOIN [D:世帯Info] ON V_農地.農家世帯ID = [D:世帯Info].ID) LEFT JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE (((V_農地.田面積)>0)) GROUP BY V_農地.農家世帯ID, [D:個人Info].行政区ID)  AS 所有世帯 INNER JOIN V_行政区 ON 所有世帯.行政区ID = V_行政区.ID GROUP BY 所有世帯.行政区ID, V_行政区.行政区, '', '', '', '' UNION SELECT 所有世帯.行政区ID, V_行政区.行政区, '' AS 世帯数, '' AS 田所有, Count(所有世帯.農家世帯ID) AS 畑所有, '' AS 樹所有 FROM (SELECT V_農地.農家世帯ID, [D:個人Info].行政区ID FROM (V_農地 LEFT JOIN [D:世帯Info] ON V_農地.農家世帯ID = [D:世帯Info].ID) LEFT JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE (((V_農地.畑面積)>0)) GROUP BY V_農地.農家世帯ID, [D:個人Info].行政区ID)  AS 所有世帯 INNER JOIN V_行政区 ON 所有世帯.行政区ID = V_行政区.ID GROUP BY 所有世帯.行政区ID, V_行政区.行政区, '', '', '', '' UNION SELECT 所有世帯.行政区ID, V_行政区.行政区, '' AS 世帯数, '' AS 田所有, '' AS 畑所有, Count(所有世帯.農家世帯ID) AS 樹所有 FROM (SELECT V_農地.農家世帯ID, [D:個人Info].行政区ID FROM (V_農地 LEFT JOIN [D:世帯Info] ON V_農地.農家世帯ID = [D:世帯Info].ID) LEFT JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE (((V_農地.樹園地)>0)) GROUP BY V_農地.農家世帯ID, [D:個人Info].行政区ID)  AS 所有世帯 INNER JOIN V_行政区 ON 所有世帯.行政区ID = V_行政区.ID GROUP BY 所有世帯.行政区ID, V_行政区.行政区, '', '', '', '')  AS [%$##@_Alias] GROUP BY [%$##@_Alias].行政区ID, [%$##@_Alias].行政区;")
            Case "開く-地区別貸付世帯件数集計"
                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [V_農地] INNER JOIN [D:世帯Info] ON [V_農地].農家世帯ID = [D:世帯Info].ID SET [D:世帯Info].自作地 = 1 WHERE ((([V_農地].自小作別)=0));")
                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [V_農地] INNER JOIN [D:世帯Info] ON [V_農地].借受世帯ID = [D:世帯Info].ID SET [D:世帯Info].自作地 = 1 WHERE ((([V_農地].自小作別)>0));")
                SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:世帯Info] INNER JOIN [V_農地] ON [D:世帯Info].ID = [V_農地].農家世帯ID SET [D:世帯Info].貸付地 = 1 WHERE ((([V_農地].借受世帯ID)<>[D:世帯Info].[ID]) AND (([V_農地].自小作別)>0));")

                OpenSQLList("地区別貸付世帯件数集計", "SELECT [D:個人Info].行政区ID, V_行政区.行政区, Sum(Int([aaa].[貸付地])) AS 貸付地の合計 FROM ((SELECT [D:世帯Info].ID, [D:世帯Info].世帯主ID, IIf([V_農地].[借受世帯ID]<>[D:世帯Info].[ID],IIf([V_農地].[自小作別]>0,1,0),0) AS 貸付地, IIf([V_農地].[農家世帯ID]=[D:世帯Info].[ID],IIf([V_農地].[自小作別]=0,1,0),0) AS 自作地 FROM V_農地 INNER JOIN [D:世帯Info] ON V_農地.農家世帯ID = [D:世帯Info].ID GROUP BY [D:世帯Info].ID, [D:世帯Info].世帯主ID, IIf([V_農地].[借受世帯ID]<>[D:世帯Info].[ID],IIf([V_農地].[自小作別]>0,1,0),0), IIf([V_農地].[農家世帯ID]=[D:世帯Info].[ID],IIf([V_農地].[自小作別]=0,1,0),0))  AS aaa LEFT JOIN [D:個人Info] ON aaa.世帯主ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE(((aaa.自作地) = 0)) GROUP BY [D:個人Info].行政区ID, V_行政区.行政区;")
            Case "開く-集落別耕作世帯集計"
                OpenSQLList("集落別耕作世帯集計", "SELECT V_行政区.名称 AS [行政区名], Count(耕作世帯.ID) AS 世帯数, Sum(Int(耕作世帯.田面積の合計)) AS 田耕作面積, Sum(Int(耕作世帯.畑面積の合計)) AS 畑耕作面積, Sum(Int(耕作世帯.樹園地の合計)) AS 樹耕作面積, Sum(Int(耕作世帯.採草放牧面積の合計)) AS 採耕作面積 FROM (SELECT [D:世帯Info].ID, [D:個人Info].行政区ID, Sum(Int(V_農地.田面積)) AS 田面積の合計, Sum(Int(V_農地.畑面積)) AS 畑面積の合計, Sum(Int(V_農地.樹園地)) AS 樹園地の合計, Sum(Int(V_農地.採草放牧面積)) AS 採草放牧面積の合計 FROM (V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE ((([V_農地].[田面積] + [V_農地].[畑面積] + [V_農地].[樹園地] + [V_農地].[採草放牧面積]) > 0)) GROUP BY [D:世帯Info].ID, [D:個人Info].行政区ID HAVING ((([D:個人Info].行政区ID)>0))) AS [耕作世帯] INNER JOIN V_行政区 ON 耕作世帯.行政区ID = V_行政区.ID GROUP BY V_行政区.ID, V_行政区.名称 ORDER BY V_行政区.ID;")
            Case "開く-集落別農地関係世帯集計"
                OpenSQLList("集落別農地関係世帯集計", "SELECT '行政区.' & [V_行政区].ID AS [KEY], V_行政区.行政区, Sum(Int([D:農地Info].田面積)) AS 自作地田面積, Sum(Int([D:農地Info].畑面積)) AS 自作地畑面積, Sum(Int([D:農地Info].樹園地)) AS 自作地樹面積, Sum(Int([D:農地Info].採草放牧面積)) AS 自作地採面積, Sum(Int([D:農地Info_1].田面積)) AS 小作地田面積, Sum(Int([D:農地Info_1].畑面積)) AS 小作地畑面積, Sum(Int([D:農地Info_1].樹園地)) AS 小作地樹面積, Sum(Int([D:農地Info_1].採草放牧面積)) AS 小作地採面積, Sum(Int([D:農地Info_2].田面積)) AS 貸付地田面積, Sum(Int([D:農地Info_2].畑面積)) AS 貸付地畑面積, Sum(Int([D:農地Info_2].樹園地)) AS 貸付地樹面積, Sum(Int([D:農地Info_2].採草放牧面積)) AS 貸付地採面積 FROM (([D:農地Info] AS [D:農地Info_2] RIGHT JOIN (([D:世帯Info] LEFT JOIN [D:農地Info] ON [D:世帯Info].ID = [D:農地Info].所有世帯ID) LEFT JOIN [D:農地Info] AS [D:農地Info_1] ON [D:世帯Info].ID = [D:農地Info_1].借受世帯ID) ON [D:農地Info_2].所有世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) INNER JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE ((([D:農地Info].自小作別)=0 Or ([D:農地Info].自小作別) Is Null) AND (([D:農地Info_1].自小作別)>0 Or ([D:農地Info_1].自小作別) Is Null) AND (([D:農地Info_2].自小作別)>0 Or ([D:農地Info_2].自小作別) Is Null)) GROUP BY V_行政区.行政区, V_行政区.ID HAVING(((V_行政区.ID) > 0)) ORDER BY V_行政区.ID;")
            Case "開く-経営規模別世帯集計"
                OpenSQLList("経営規模別世帯集計", "SELECT IIf(Fix([世帯経営面積]/1000)>0,Right$('   ' & Fix([世帯経営面積]/1000)*10,4) & 'アール以上','  10 アール未満') AS 経営規模, Count(世帯経営.ID) AS 世帯数 FROM (SELECT [D:世帯Info].ID, Sum(Int([V_農地].[田面積])+Int([V_農地].[畑面積])+Int([V_農地].[樹園地])+Int([V_農地].[採草放牧面積])) AS 世帯経営面積 FROM V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID GROUP BY [D:世帯Info].ID ) AS 世帯経営 GROUP BY IIf(Fix([世帯経営面積]/1000)>0,Right$('   ' & Fix([世帯経営面積]/1000)*10,4) & 'アール以上','  10 アール未満') ORDER BY IIf(Fix([世帯経営面積]/1000)>0,Right$('   ' & Fix([世帯経営面積]/1000)*10,4) & 'アール以上','  10 アール未満')")
            Case "開く-後継者状況別世帯集計"
                OpenSQLList("後継者状況別世帯集計", "SELECT [D:個人Info].行政区ID, V_行政区.行政区, Count([D:世帯Info].ID) AS 世帯数 FROM ([D:個人Info] INNER JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) INNER JOIN [D:世帯Info] ON [D:個人Info].ID = [D:世帯Info].世帯主ID WHERE ((([D:個人Info].農業跡継ぎ)=True)) GROUP BY [D:個人Info].行政区ID, V_行政区.行政区;")
            Case "開く-都道府県別世帯一覧"
                OpenSQLList("都道府県別世帯一覧", "SELECT '農家.' & [D:世帯Info].[ID] AS [Key], IIf(InStr([D:個人Info].[住所],'東京都')>0,'東京都',IIf(InStr([D:個人Info].[住所],'北海道')>0,'北海道',IIf(InStr([D:個人Info].[住所],'大阪府')>0,'大阪府',Left([D:個人Info].[住所],InStr([D:個人Info].[住所],'県'))))) AS 都道府県, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].住所, [D:世帯Info].確認日時 FROM [D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE (((IIf(InStr([D:個人Info].[住所],'東京都')>0,'東京都',IIf(InStr([D:個人Info].[住所],'北海道')>0,'北海道',IIf(InStr([D:個人Info].[住所],'大阪府')>0,'大阪府',Left([D:個人Info].[住所],InStr([D:個人Info].[住所],'県'))))))<>'' And (IIf(InStr([D:個人Info].[住所],'東京都')>0,'東京都',IIf(InStr([D:個人Info].[住所],'北海道')>0,'北海道',IIf(InStr([D:個人Info].[住所],'大阪府')>0,'大阪府',Left([D:個人Info].[住所],InStr([D:個人Info].[住所],'県')))))) Is Not Null) AND (([D:個人Info].住所) Is Not Null)) ORDER BY IIf(InStr([D:個人Info].[住所],'東京都')>0,'東京都',IIf(InStr([D:個人Info].[住所],'北海道')>0,'北海道',IIf(InStr([D:個人Info].[住所],'大阪府')>0,'大阪府',Left([D:個人Info].[住所],InStr([D:個人Info].[住所],'県')))));")
            Case "開く-管内農地一覧"
                OpenSQLList("管内農地一覧", "SELECT [D:農地Info].ID AS 農地ID, [D:農地Info].大字ID, V_大字.名称 AS 大字, V_小字.名称 AS 小字, [D:農地Info].地番, [D:農地Info].一部現況, V_地目.名称 AS 登記地目, V_現況地目.名称 AS 現況地目, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].所有者ID, [D:個人Info].世帯ID, [D:個人Info].氏名, [D:個人Info].住所, IIf(Format([D:個人Info].[生年月日],'mm/dd')>Format(Date(),'mm/dd'),DateDiff('yyyy',[D:個人Info].[生年月日],Date())-1,DateDiff('yyyy',[D:個人Info].[生年月日],Date())) AS 年齢, IIf([自小作別]>0,'小作','自作') AS 自小作, IIf([自小作別]>0,[借受人ID],Null) AS 借人ID, IIf([自小作別]>0,[D:個人Info_1].[氏名],Null) AS 借人氏名, IIf([自小作別]>0,[D:個人Info_1].[住所],Null) AS 借人住所, IIf([自小作別]>0,IIf(Format([D:個人Info_1].[生年月日],'mm/dd')>Format(Date(),'mm/dd'),DateDiff('yyyy',[D:個人Info_1].[生年月日],Date())-1,DateDiff('yyyy',[D:個人Info_1].[生年月日],Date())),Null) AS 借人年齢 FROM ((((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地Info].借受人ID = [D:個人Info_1].ID WHERE ((([D:農地Info].大字ID)>0)) ORDER BY IIf(Format([D:個人Info].[生年月日],'mm/dd')>Format(Date(),'mm/dd'),DateDiff('yyyy',[D:個人Info].[生年月日],Date())-1,DateDiff('yyyy',[D:個人Info].[生年月日],Date())), [D:農地Info].大字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])), IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),0);")
            Case "開く-大字別面積集計"
                OpenSQLList("大字別面積集計", "SELECT Alias.大字名, Alias.田合計, Alias.田筆数, Alias.畑合計, Alias.畑筆数, Alias.樹園地合計, Alias.樹筆数, Alias.採草放牧地面積, Alias.採筆数, [田合計]+[畑合計]+[樹園地合計]+[採草放牧地面積] AS 農地計, [田筆数]+[畑筆数]+[樹筆数]+[採筆数] AS 筆数計 FROM (SELECT V_大字.名称 AS 大字名, Sum(IIf(IsNull([農委地目ID]),IIf(InStr([V_現況地目].[名称],'田')>0,Int([田面積]),0),IIf(InStr([V_農委地目].[名称],'田')>0,Int([田面積]),0))) AS 田合計, Sum(IIf(IsNull([農委地目ID]),IIf(InStr([V_現況地目].[名称],'田')>0,IIf([田面積]>0,1,0),0),IIf(InStr([V_農委地目].[名称],'田')>0,IIf([田面積]>0,1,0),0))) AS 田筆数, Sum(IIf(IsNull([農委地目ID]),IIf(InStr([V_現況地目].[名称],'畑')>0,Int([畑面積]),0),IIf(InStr([V_農委地目].[名称],'畑')>0,Int([畑面積]),0))) AS 畑合計, Sum(IIf(IsNull([農委地目ID]),IIf(InStr([V_現況地目].[名称],'畑')>0,IIf([畑面積]>0,1,0),0),IIf(InStr([V_農委地目].[名称],'畑')>0,IIf([畑面積]>0,1,0),0))) AS 畑筆数, Sum(Int([樹園地])) AS 樹園地合計, Sum(IIf([樹園地]>0,1,0)) AS 樹筆数, Sum(Int([採草放牧面積])) AS 採草放牧地面積, Sum(IIf([採草放牧面積]>0,1,0)) AS 採筆数 FROM ((V_農地 INNER JOIN V_大字 ON V_農地.大字ID = V_大字.ID) LEFT JOIN V_現況地目 ON V_農地.現況地目 = V_現況地目.ID) LEFT JOIN V_農委地目 ON V_農地.農委地目ID = V_農委地目.ID GROUP BY V_大字.名称, V_大字.ID HAVING(((V_大字.ID) > 0)) ORDER BY V_大字.ID)  AS Alias GROUP BY Alias.大字名, Alias.田合計, Alias.田筆数, Alias.畑合計, Alias.畑筆数, Alias.樹園地合計, Alias.樹筆数, Alias.採草放牧地面積, Alias.採筆数, [田合計]+[畑合計]+[樹園地合計]+[採草放牧地面積], [田筆数]+[畑筆数]+[樹筆数]+[採筆数];")
            Case "開く-登記地目別面積集計"
                OpenSQLList("登記地目別集計", "TRANSFORM Sum(Int([登記簿面積])) AS 登記面積 SELECT [D:農地Info].大字ID, V_大字.大字, Sum(Int([登記簿面積])) AS [合計 登記面積], Sum(Sgn([登記簿面積])) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID GROUP BY [D:農地Info].大字ID, V_大字.大字 PIVOT [V_地目].[ID] & ':' & [V_地目].[名称];")
                '↓過去集計
                'OpenSQLList("登記地目別集計", "SELECT V_地目.ID, V_地目.名称 AS 登記地目, Count(V_地目.ID) AS 筆数, Sum(Int([田面積])+Int([畑面積])+Int([樹園地])+Int([採草放牧面積])) AS 耕作面積 FROM V_農地 INNER JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID GROUP BY V_地目.ID, V_地目.名称 ORDER BY V_地目.ID;")
            Case "開く-現況地目別面積集計"
                OpenSQLList("現況地目別集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT [D:農地Info].大字ID, V_大字.大字, Sum(Int([実面積])) AS [合計 現況面積], Sum(Sgn([実面積])) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID GROUP BY [D:農地Info].大字ID, V_大字.大字 PIVOT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称];", {"合計 現況面積"})
                '↓過去集計
                'OpenSQLList("現況地目別集計", "SELECT V_現況地目.ID, V_現況地目.名称 AS 現況地目, Count(V_現況地目.ID) AS 筆数, Sum(Int([田面積])+Int([畑面積])+Int([樹園地])+Int([採草放牧面積])) AS 耕作面積 FROM V_農地 INNER JOIN V_現況地目 ON V_農地.現況地目 = V_現況地目.ID GROUP BY V_現況地目.ID, V_現況地目.名称 ORDER BY V_現況地目.ID;")
            Case "開く-現況地目別管内面積集計"
                OpenSQLList("現況地目別管内面積集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE ((([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR ((([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR ((([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR ((([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT IIf([所在] Is Null,'市町村管内農地面積','当該市町村外');")
            Case "開く-年齢別耕作農地一覧"
                OpenSQLList("年齢別耕作農地一覧", "SELECT Year(Now())-Year([生年月日]) AS 年齢, Sum(Int(IIf([自小作別]=0,[V_農地].[田面積],0))) AS 自作田, Sum(Int(IIf([自小作別]=0,[V_農地].[畑面積],0))) AS 自作畑, Sum(Int(IIf([自小作別]=0,[V_農地].[樹園地],0))) AS 自作樹, Sum(Int(IIf([自小作別]=0,[V_農地].[採草放牧面積],0))) AS 自作採, Sum(Int(IIf([自小作別]=0,[V_農地].[田面積],0))+Int(IIf([自小作別]=0,[V_農地].[畑面積],0))+Int(IIf([自小作別]=0,[V_農地].[樹園地],0))+Int(IIf([自小作別]=0,[V_農地].[採草放牧面積],0))) AS 自作計, Sum(Int(IIf([自小作別]>0,[V_農地].[田面積],0))) AS 小作田, Sum(Int(IIf([自小作別]>0,[V_農地].[畑面積],0))) AS 小作畑, Sum(Int(IIf([自小作別]>0,[V_農地].[樹園地],0))) AS 小作樹, Sum(Int(IIf([自小作別]>0,[V_農地].[採草放牧面積],0))) AS 小作採, Sum(Int(IIf([自小作別]>0,[V_農地].[田面積],0))+Int(IIf([自小作別]>0,[V_農地].[畑面積],0))+Int(IIf([自小作別]>0,[V_農地].[樹園地],0))+Int(IIf([自小作別]>0,[V_農地].[採草放牧面積],0))) AS 小作計, Sum(Int(IIf([自小作別]=0,[V_農地].[田面積],0))+Int(IIf([自小作別]=0,[V_農地].[畑面積],0))+Int(IIf([自小作別]=0,[V_農地].[樹園地],0))+Int(IIf([自小作別]=0,[V_農地].[採草放牧面積],0))+Int(IIf([自小作別]>0,[V_農地].[田面積],0))+Int(IIf([自小作別]>0,[V_農地].[畑面積],0))+Int(IIf([自小作別]>0,[V_農地].[樹園地],0))+Int(IIf([自小作別]>0,[V_農地].[採草放牧面積],0))) AS 経営地計 FROM (V_農地 INNER JOIN [D:世帯Info] ON V_農地.耕作世帯ID = [D:世帯Info].ID) INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID WHERE ((([D:個人Info].生年月日) Is Not Null)) GROUP BY Year(Now())-Year([生年月日]) ORDER BY Year(Now())-Year([生年月日]);")
            Case "開く-未世帯農地一覧"
                Dim sWhere As String = "([所有世帯ID]=0 Or [所有世帯ID] Is Null)"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-台帳内作成農地一覧"
                Dim sWhere As String = "[ID]<0"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-登記地目が農地以外一覧"
                Dim sWhere As String = "Not [登記簿地目] IN (10,11,20,21)"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-町外・法人所有情報"
                OpenSQLList("町外・法人所有情報", String.Format("SELECT Sum(IIf(IsNull([D:農地Info].[田面積]),0,Int([D:農地Info].[田面積]))) AS 田面積の合計, Sum(IIf(IsNull([D:農地Info].[畑面積]),0,Int([D:農地Info].[畑面積]))) AS 畑面積の合計, Sum(IIf(([D:農地Info].[田面積])<>0,1,0)) AS 田筆数, Sum(IIf([D:農地Info].[畑面積]<>0,1,0)) AS 畑筆数 FROM [D:個人Info] INNER JOIN [D:農地Info] ON [D:個人Info].ID = [D:農地Info].所有者ID WHERE ((([D:個人Info].住民区分)<>0 And ([D:個人Info].住民区分)<>{0}));", SysAD.DB(sLRDB).DBProperty("死亡住民コード")))
            Case "開く-死亡所有農地一覧"
                OpenSQLList("申請無転用農地一覧", String.Format("SELECT [D:農地Info].ID, V_大字.大字, V_小字.小字, [D:農地Info].地番, V_地目.名称 AS 登記地目名, V_現況地目.名称 AS 現況地目名, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].所有者ID, [D:個人Info].氏名, [D:個人Info].住所, IIf([自小作別]=1,'小作',IIf([自小作別]=2,'農年','自作')) AS 自小作, [D:農地Info].借受人ID, [D:個人Info_1].氏名 AS 借受人名, '' AS 代理人 FROM ((((([D:農地Info] LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地Info].借受人ID = [D:個人Info_1].ID WHERE ((([D:個人Info].住民区分)={0}));", SysAD.DB(sLRDB).DBProperty("死亡住民コード")))
            Case "開く-死亡名義人農地一覧"
                OpenSQLList("死亡名義人農地一覧", String.Format("SELECT [D:農地Info].ID, V_大字.大字, V_小字.小字, [D:農地Info].地番, V_地目.名称 AS 登記地目名, V_現況地目.名称 AS 現況地目名, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].所有者ID, [D:個人Info].氏名, [D:個人Info].住所, IIf([自小作別]=1,'小作',IIf([自小作別]=2,'農年','自作')) AS 自小作, [D:農地Info].借受人ID, [D:個人Info_1].氏名 AS 借受人名, [D:農地Info].管理者ID, [D:個人Info_2].氏名 AS 管理者名, IIf([農地所有内訳]=1,'管理人',IIf([農地所有内訳]=2,'代理人',IIf([農地所有内訳]=3,'変更済み',IIf([農地所有内訳]=0,'-','')))) AS 所有内訳 FROM (((((([D:農地Info] LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地Info].借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON [D:農地Info].管理者ID = [D:個人Info_2].ID WHERE ((([D:個人Info].住民区分)={0}));", SysAD.DB(sLRDB).DBProperty("死亡住民コード")))
            Case "開く-転用(非農地)農地一覧"
                If Not SysAD.page農家世帯.TabPageContainKey("転用(非農地)農地一覧", True) Then SysAD.page農家世帯.中央Tab.AddPage(New CTabPage転用農地一覧())
            Case "開く-転用農地申請履歴一覧"
                If Not SysAD.page農家世帯.TabPageContainKey("転用農地申請履歴一覧", True) Then SysAD.page農家世帯.中央Tab.AddPage(New CTabPage転用農地申請履歴一覧())
            Case "開く-申請無転用農地一覧"
                If Not SysAD.page農家世帯.TabPageContainKey("申請無転用農地一覧", True) Then SysAD.page農家世帯.中央Tab.AddPage(New CTabPage申請無転用農地一覧())

            Case "開く-都道府県別世帯一覧"
                Dim sWhere As String = "[都道府県ID]=" & Val(SysAD.DB(sLRDB).DBProperty("都道府県ID").ToString)
                SysAD.page農家世帯.農家リスト.検索開始(sWhere, sWhere)

            Case "開く-規模拡大希望世帯一覧"
                Dim sWhere As String = "([規模拡大希望]=True)"
                SysAD.page農家世帯.農家リスト.検索開始(sWhere, sWhere)
            Case "開く-法人化希望世帯一覧"
                Dim sWhere As String = "([法人化希望]=True)"
                SysAD.page農家世帯.農家リスト.検索開始(sWhere, sWhere)

            Case "開く-利用権設定始期期間別一覧"
                With New dlgInputBWDate(SysAD.GetXMLProperty("集計", "利用権設定始期期間別一覧-開始", Now.Date))
                    If .ShowDialog = DialogResult.OK Then
                        Dim sWhere As String = String.Format("[自小作別]>0 AND [小作地適用法]=2 AND [小作開始年月日]>=#{0}/{1}/{2}# And [小作開始年月日]<=#{3}/{4}/{5}#",
                        .StartDate.Month, .StartDate.Day, .StartDate.Year, .EndDate.Month, .EndDate.Day, .EndDate.Year)

                        SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
                        SysAD.SetXMLProperty("集計", "利用権設定始期期間別一覧-開始", .StartDate)
                    End If
                End With
            Case "開く-利用権設定終期期間別一覧"
                With New dlgInputBWDate(SysAD.GetXMLProperty("集計", "利用権設定終期期間別一覧-開始", Now.Date))
                    If .ShowDialog = DialogResult.OK Then
                        Dim sWhere As String = String.Format("[自小作別]>0 AND [小作地適用法]=2 AND [小作終了年月日]>=#{0}/{1}/{2}# And [小作終了年月日]<=#{3}/{4}/{5}#",
                        .StartDate.Month, .StartDate.Day, .StartDate.Year, .EndDate.Month, .EndDate.Day, .EndDate.Year)

                        SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
                        SysAD.SetXMLProperty("集計", "利用権設定終期期間別一覧-開始", .StartDate)
                    End If
                End With
            Case "開く-申請時始期期間別一覧"
                With New dlgInputBWDate(SysAD.GetXMLProperty("集計", "申請時始期期間別一覧-開始", Now.Date))
                    If .ShowDialog = DialogResult.OK Then
                        Dim Start日付 As String = String.Format("#{0}/{1}/{2}#", .StartDate.Month, .StartDate.Day, .StartDate.Year)
                        Dim End日付 As String = String.Format("#{0}/{1}/{2}#", .EndDate.Month, .EndDate.Day, .EndDate.Year)
                        If Not SysAD.page農家世帯.TabPageContainKey("申請時始期期間別一覧", True) Then SysAD.page農家世帯.中央Tab.AddPage(New CTabPage貸借農地終期期間別一覧("利用権設定始期期間別一覧", Start日付, End日付))
                    End If
                End With
            Case "開く-申請時終期期間別一覧"
                With New dlgInputBWDate(SysAD.GetXMLProperty("集計", "申請時終期期間別一覧-開始", Now.Date))
                    If .ShowDialog = DialogResult.OK Then
                        Dim Start日付 As String = String.Format("#{0}/{1}/{2}#", .StartDate.Month, .StartDate.Day, .StartDate.Year)
                        Dim End日付 As String = String.Format("#{0}/{1}/{2}#", .EndDate.Month, .EndDate.Day, .EndDate.Year)
                        If Not SysAD.page農家世帯.TabPageContainKey("申請時終期期間別一覧", True) Then SysAD.page農家世帯.中央Tab.AddPage(New CTabPage貸借農地終期期間別一覧("利用権設定終期期間別一覧", Start日付, End日付))
                    End If
                End With
            Case "開く-農振農用地内一覧"
                Dim sWhere As String = "([農業振興地域]=1) OR ([農振法区分] In (1,2))"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-農用地集計"
                OpenSQLList("農用地集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE (((IIf([農業振興地域]=1,1,[農振法区分]))=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR (((IIf([農業振興地域]=1,1,[農振法区分]))=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR (((IIf([農業振興地域]=1,1,[農振法区分]))=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR (((IIf([農業振興地域]=1,1,[農振法区分]))=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '農用地面積';")
            Case "開く-農用地外集計"
                OpenSQLList("農用地外集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE (((IIf([農業振興地域]=0,2,[農振法区分]))=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR (((IIf([農業振興地域]=0,2,[農振法区分]))=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR (((IIf([農業振興地域]=0,2,[農振法区分]))=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR (((IIf([農業振興地域]=0,2,[農振法区分]))=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '農用地外面積';")
            Case "開く-農振外集計"
                OpenSQLList("農振外集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE (((IIf([農業振興地域]=2,3,[農振法区分]))=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR (((IIf([農業振興地域]=2,3,[農振法区分]))=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR (((IIf([農業振興地域]=2,3,[農振法区分]))=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR (((IIf([農業振興地域]=2,3,[農振法区分]))=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '農振外面積';")

            Case "開く-都計外集計"
                OpenSQLList("都計外集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE ((([D:農地Info].都市計画法)=0) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR ((([D:農地Info].都市計画法)=0) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR ((([D:農地Info].都市計画法)=0) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR ((([D:農地Info].都市計画法)=0) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '都計外面積';")
            Case "開く-都計内集計"
                OpenSQLList("都計内集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE ((([D:農地Info].都市計画法)=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR ((([D:農地Info].都市計画法)=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR ((([D:農地Info].都市計画法)=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR ((([D:農地Info].都市計画法)=1) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '都計内面積';")
            Case "開く-用途地域内集計"
                OpenSQLList("用途地域内集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE ((([D:農地Info].都市計画法)=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR ((([D:農地Info].都市計画法)=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR ((([D:農地Info].都市計画法)=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR ((([D:農地Info].都市計画法)=2) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '用途地域内面積';")
            Case "開く-調整区域内集計"
                OpenSQLList("調整区域内集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE ((([D:農地Info].都市計画法)=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR ((([D:農地Info].都市計画法)=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR ((([D:農地Info].都市計画法)=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR ((([D:農地Info].都市計画法)=3) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '調整区域内面積';")
            Case "開く-市街化区域内集計"
                OpenSQLList("市街化区域内集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE ((([D:農地Info].都市計画法)=4) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR ((([D:農地Info].都市計画法)=4) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR ((([D:農地Info].都市計画法)=4) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR ((([D:農地Info].都市計画法)=4) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '市街化区域内面積'; ")
            Case "開く-都市計画白地集計"
                OpenSQLList("都市計画白地集計", "TRANSFORM Sum(Int([実面積])) AS 現況面積 SELECT Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] AS 現況地目, Count(1) AS 筆数 FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID WHERE ((([D:農地Info].都市計画法)=5) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].田面積)>0)) OR ((([D:農地Info].都市計画法)=5) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].畑面積)>0)) OR ((([D:農地Info].都市計画法)=5) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].樹園地)>0)) OR ((([D:農地Info].都市計画法)=5) AND (([D:農地Info].所在) Is Null) AND (([D:農地Info].採草放牧面積)>0)) GROUP BY V_現況地目.ID, Format([V_現況地目].[ID],'00') & ':' & [V_現況地目].[名称] ORDER BY V_現況地目.ID PIVOT '都市計画白地面積';")
            Case "開く-経営面積農地地目別集計"
                OpenSQLList("経営面積農地地目別集計", "SELECT V_大字.ID, V_大字.大字, Sum(Int(V_農地.田面積)) AS 田面積の合計, Sum(Int(V_農地.畑面積)) AS 畑面積の合計, Sum(Int(V_農地.樹園地)) AS 樹園地の合計, Sum(Int(V_農地.採草放牧面積)) AS 採草放牧面積の合計 FROM V_農地 INNER JOIN V_大字 ON V_農地.大字ID = V_大字.ID GROUP BY V_大字.ID, V_大字.大字;", {"田面積の合計", "畑面積の合計"})
            Case "開く-利用権設定大字別面積集計"
                'OpenSQLList("利用権設定大字別面積集計", "TRANSFORM Sum(int([D:農地Info].[実面積])) & '(' & Val(' ' & Count([D:農地Info].[ID])) & ' 筆)' AS 実面積の合計 SELECT [D:農地Info].[大字ID], V_大字.大字, Sum(int([D:農地Info].[実面積])) & '(' & Val(' ' & Count([D:農地Info].[ID])) & ' 筆)' AS [合計 実面積] FROM ([D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) INNER JOIN V_小作形態 ON [D:農地Info].小作形態 = V_小作形態.ID WHERE ((([D:農地Info].自小作別)>0) AND (([D:農地Info].小作地適用法)=2)) GROUP BY [D:農地Info].[大字ID], V_大字.大字 PIVOT V_小作形態.名称;", {"実面積の合計"})
                OpenSQLList("利用権設定大字別面積集計", "SELECT Arias.大字ID, Arias.大字, [実面積の合計] & '(' & [筆数] & '筆)' AS 面積計, IIf([田面積の合計]>0,[田面積の合計] & '(' & [田筆数] & '筆)','') AS 田面積計, IIf([畑面積の合計]>0,[畑面積の合計] & '(' & [畑筆数] & '筆)','') AS 畑面積計, IIf([― 面積]>0,[― 面積] & '(' & [― 筆数] & '筆)','') AS ―　計, IIf([― 田面積]>0,[― 田面積] & '(' & [― 田筆数] & '筆)','') AS ―　田計, IIf([― 畑面積]>0,[― 畑面積] & '(' & [― 畑筆数] & '筆)','') AS ―　畑計, IIf([賃貸借 面積]>0,[賃貸借 面積] & '(' & [賃貸借 筆数] & '筆)','') AS 賃貸借　計, IIf([賃貸借 田面積]>0,[賃貸借 田面積] & '(' & [賃貸借 田筆数] & '筆)','') AS 賃貸借　田計, IIf([賃貸借 畑面積]>0,[賃貸借 畑面積] & '(' & [賃貸借 畑筆数] & '筆)','') AS 賃貸借　畑計,  IIf([使用貸借 面積]>0,[使用貸借 面積] & '(' & [使用貸借 筆数] & '筆)','') AS 使用貸借　計, IIf([使用貸借 田面積]>0,[使用貸借 田面積] & '(' & [使用貸借 田筆数] & '筆)','') AS 使用貸借　田計, IIf([使用貸借 畑面積]>0,[使用貸借 畑面積] & '(' & [使用貸借 畑筆数] & '筆)','') AS 使用貸借　畑計 " &
                                                                             "FROM (SELECT [D:農地Info].大字ID, V_大字.大字, Sum([D:農地Info].実面積) AS 実面積の合計, Sum(1) AS 筆数, Sum([D:農地Info].田面積) AS 田面積の合計, Sum(IIF([田面積]>0,1,0)) AS 田筆数, Sum([D:農地Info].畑面積) AS 畑面積の合計, Sum(IIF([畑面積]>0,1,0)) AS 畑筆数,  Sum(IIf([小作形態]=0,[実面積],0)) AS [― 面積], Sum(IIf([小作形態]=0,1,0)) AS [― 筆数], Sum(IIf([小作形態]=0,[田面積],0)) AS [― 田面積], Sum(IIf([小作形態]=0 And [田面積]>0,1,0)) AS [― 田筆数], Sum(IIf([小作形態]=0,[畑面積],0)) AS [― 畑面積], Sum(IIf([小作形態]=0 And [畑面積]>0,1,0)) AS [― 畑筆数],  Sum(IIf([小作形態]=1,[実面積],0)) AS [賃貸借 面積], Sum(IIf([小作形態]=1,1,0)) AS [賃貸借 筆数], Sum(IIf([小作形態]=1,[田面積],0)) AS [賃貸借 田面積], Sum(IIf([小作形態]=1 And [田面積]>0,1,0)) AS [賃貸借 田筆数], Sum(IIf([小作形態]=1,[畑面積],0)) AS [賃貸借 畑面積], Sum(IIf([小作形態]=1 And [畑面積]>0,1,0)) AS [賃貸借 畑筆数],  Sum(IIf([小作形態]=2,[実面積],0)) AS [使用貸借 面積], Sum(IIf([小作形態]=2,1,0)) AS [使用貸借 筆数], Sum(IIf([小作形態]=2,[田面積],0)) AS [使用貸借 田面積], Sum(IIf([小作形態]=2 And [田面積]>0,1,0)) AS [使用貸借 田筆数], Sum(IIf([小作形態]=2,[畑面積],0)) AS [使用貸借 畑面積], Sum(IIf([小作形態]=2 And [畑面積]>0,1,0)) AS [使用貸借 畑筆数] " &
                                                                             "FROM [D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID " &
                                                                             "WHERE ((([D:農地Info].自小作別)>0) AND (([D:農地Info].小作地適用法)=2)) " &
                                                                             "GROUP BY [D:農地Info].大字ID, V_大字.大字 " &
                                                                             "HAVING ((([D:農地Info].大字ID)>0)) " &
                                                                             "ORDER BY [D:農地Info].大字ID)  AS Arias;", {"実面積の合計"})
                '↓過去集計
                'OpenSQLList("利用権設定大字別集計", "SELECT IIf([小作形態]=3,'期間借地',IIf([小作形態]=2,'使用貸借',IIf([小作形態]=1,'賃貸借'))) AS 形態, Sum(Int([田面積])) AS 田計, Sum(Int([畑面積])) AS 畑計, Sum(Int([樹園地])) AS 樹園地計, Sum(Int([採草放牧面積])) AS 採草放牧面積計, Sum(Int([田面積])+Int([畑面積])+Int([樹園地])+Int([採草放牧面積])) AS 農地計, Count([D:農地Info].ID) AS 筆数 FROM [D:農地Info] WHERE ((([D:農地Info].自小作別)>0) AND (([D:農地Info].小作地適用法)=2)) GROUP BY IIf([小作形態]=3,'期間借地',IIf([小作形態]=2,'使用貸借',IIf([小作形態]=1,'賃貸借'))), [D:農地Info].小作形態 ORDER BY [D:農地Info].小作形態;")
            Case "開く-利用権設定をした集落別面積集計"
                OpenSQLList("利用権設定をした集落別面積集計", "SELECT '行政区.' & [行政区ID] AS [Key], V_行政区.行政区, Sum(Int(V_農地.田面積)) AS 利用権を設定した田, Sum(Int(V_農地.畑面積)) AS 利用権を設定した畑, Sum(Int(V_農地.樹園地)) AS 利用権を設定した樹, Sum(Int(V_農地.採草放牧面積)) AS 利用権を設定した採 FROM (V_農地 LEFT JOIN [D:個人Info] ON V_農地.管理人ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE (((V_農地.小作地適用法)=2) AND ((V_農地.自小作別)>0)) GROUP BY '行政区.' & [行政区ID], V_行政区.行政区, V_行政区.ID ORDER BY V_行政区.ID;")
            Case "開く-利用権設定を受けた集落別面積集計"
                OpenSQLList("利用権設定を受けた集落別面積集計", "SELECT '行政区.' & [行政区ID] AS [Key], V_行政区.行政区, Sum(Int(V_農地.田面積)) AS 利用権を受けた田, Sum(Int(V_農地.畑面積)) AS 利用権を受けた畑, Sum(Int(V_農地.樹園地)) AS 利用権を受けた樹, Sum(Int(V_農地.採草放牧面積)) AS 利用権を受けた採 FROM (V_農地 LEFT JOIN [D:個人Info] ON V_農地.借受人ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID WHERE (((V_農地.小作地適用法)=2) AND ((V_農地.自小作別)>0)) GROUP BY '行政区.' & [行政区ID], V_行政区.行政区;")
            Case "開く-大字年度毎10a当賃借料集計"
                OpenSQLList("大字-年度毎10a当賃借料集計", "TRANSFORM Avg([D:農地Info].[10a賃借料]) AS 10a賃借料の平均 SELECT [D:農地Info].[大字ID], V_大字.大字, Avg([D:農地Info].[10a賃借料]) AS [大字毎平均 10a賃借料] FROM [D:農地Info] INNER JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID WHERE ((([D:農地Info].小作地適用法)=2) AND (([D:農地Info].小作形態)=1) AND (([D:農地Info].自小作別)>0) AND (([D:農地Info].[10a賃借料])>0) AND (([D:農地Info].小作開始年月日) Is Not Null)) GROUP BY [D:農地Info].[大字ID], V_大字.大字 PIVOT Format([小作開始年月日],'gggee');")
            Case "開く-年度別許可済み申請集計"
                OpenSQLList("年度別許可済み申請集計", "TRANSFORM Count(D_申請.ID) AS IDのカウント SELECT Format([許可年月日],'gee') AS 許可年度 FROM D_申請 INNER JOIN M_BASICALL ON D_申請.法令 = M_BASICALL.ID WHERE (((D_申請.許可年月日)>1900) AND ((M_BASICALL.Class)='法令')) GROUP BY Format([許可年月日],'gee') ORDER BY Right('000' & [M_BASICALL].[ID],3) & ':' & [M_BASICALL].[名称] PIVOT Right('000' & [M_BASICALL].[ID],3) & ':' & [M_BASICALL].[名称];")
            Case "開く-農振地内異動済み農地集計"
                Act農振地内異動済み農地集計()
            Case "開く-面積別集計"
                OpenSQLList("面積別集計", "SELECT IIf(IIf([田面積] Is Null,0,[田面積])+IIf([畑面積] Is Null,0,[畑面積])+IIf([樹園地] Is Null,0,[樹園地])+IIf([採草放牧面積] Is Null,0,[採草放牧面積])<1000,'  10アール 未満',Right$('   ' & Fix((IIf([田面積] Is Null,0,[田面積])+IIf([畑面積] Is Null,0,[畑面積])+IIf([樹園地] Is Null,0,[樹園地])+IIf([採草放牧面積] Is Null,0,[採草放牧面積]))/1000)*10,4) & 'アール以上') AS 面積, Sum(Int([登記簿面積])) & '(' & Val(' ' & Count([D:農地Info].[ID])) & ' 筆)' AS [合計 登記面積], Sum(Int([実面積])) AS [合計 現況面積], Sum(Int([田面積])) AS [合計 田面積], Sum(Int([畑面積])) AS [合計 畑面積], Sum(Int([樹園地])) AS [合計 樹園地], Sum(Int([採草放牧面積])) AS [合計 採草放牧面積] FROM [D:農地Info] GROUP BY IIf(IIf([田面積] Is Null,0,[田面積])+IIf([畑面積] Is Null,0,[畑面積])+IIf([樹園地] Is Null,0,[樹園地])+IIf([採草放牧面積] Is Null,0,[採草放牧面積])<1000,'  10アール 未満',Right$('   ' & Fix((IIf([田面積] Is Null,0,[田面積])+IIf([畑面積] Is Null,0,[畑面積])+IIf([樹園地] Is Null,0,[樹園地])+IIf([採草放牧面積] Is Null,0,[採草放牧面積]))/1000)*10,4) & 'アール以上') ORDER BY IIf(IIf([田面積] Is Null,0,[田面積])+IIf([畑面積] Is Null,0,[畑面積])+IIf([樹園地] Is Null,0,[樹園地])+IIf([採草放牧面積] Is Null,0,[採草放牧面積])<1000,'  10アール 未満',Right$('   ' & Fix((IIf([田面積] Is Null,0,[田面積])+IIf([畑面積] Is Null,0,[畑面積])+IIf([樹園地] Is Null,0,[樹園地])+IIf([採草放牧面積] Is Null,0,[採草放牧面積]))/1000)*10,4) & 'アール以上');")
            Case "開く-不整合貸借農地一覧"
                Dim sWhere As String = String.Format("(([自小作別]<>1) AND ([小作終了年月日]>#{0}#))", DateTime.Now())
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-面積未設定エラー"
                Dim sWhere As String = "(([農委地目ID]=1) AND ([田面積]=0 Or [田面積] Is Null)) OR (([農委地目ID]=2) AND ([畑面積]=0 Or [畑面積] Is Null)) OR (([現況地目]=1) AND ([農委地目ID] Is Null) AND ([田面積]=0 Or [田面積] Is Null)) OR (([現況地目]=2) AND ([農委地目ID] Is Null) AND ([畑面積]=0 Or [畑面積] Is Null))"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-貸借集計"
                OpenSQLList("貸借集計", "SELECT IIf([小作地適用法]=1,'農地法',IIf([小作地適用法],'基盤強化法','その他')) AS 法令, IIf([小作形態]=1,'賃貸借',IIf([小作形態]=2,'使用貸借','その他')) AS 形態, IIf(DateDiff('yyyy',[小作開始年月日],[小作終了年月日])>0,DateDiff('yyyy',[小作開始年月日],[小作終了年月日]),Null) AS 期間, Format$(Sum(Int([田面積])),'#,##0') AS 田面積の合計, Format$(Sum(Int([畑面積])),'#,##0') AS 畑面積の合計, Format$(Sum(Int([樹園地])),'#,##0') AS 樹園地の合計, Format$(Sum(Int([採草放牧面積])),'#,##0') AS 採草放牧面積の合計 FROM V_農地 WHERE (((V_農地.自小作別)>0)) GROUP BY IIf([小作地適用法]=1,'農地法',IIf([小作地適用法],'基盤強化法','その他')), IIf([小作形態]=1,'賃貸借',IIf([小作形態]=2,'使用貸借','その他')), IIf(DateDiff('yyyy',[小作開始年月日],[小作終了年月日])>0,DateDiff('yyyy',[小作開始年月日],[小作終了年月日]),Null), IIf(IsNull([小作地適用法]),0,[小作地適用法]), IIf(IsNull([小作形態]),0,[小作形態]), V_農地.小作地適用法, V_農地.小作形態, IIf(DateDiff('yyyy',[小作開始年月日],[小作終了年月日])>0,DateDiff('yyyy',[小作開始年月日],[小作終了年月日]),Null) ORDER BY IIf(IsNull([小作地適用法]),0,[小作地適用法]), IIf(IsNull([小作形態]),0,[小作形態]), V_農地.小作地適用法, V_農地.小作形態, IIf(DateDiff('yyyy',[小作開始年月日],[小作終了年月日])>0,DateDiff('yyyy',[小作開始年月日],[小作終了年月日]),Null), V_農地.小作地適用法, V_農地.小作形態, IIf(DateDiff('yyyy',[小作開始年月日],[小作終了年月日])>0,DateDiff('yyyy',[小作開始年月日],[小作終了年月日]),Null);")
            Case "開く-貸借の形態別集計"
                OpenSQLList("貸借の形態別集計", "SELECT IIf([小作地適用法]=0 Or IsNull([小作地適用法]),4 & ':その他',[小作地適用法] & ':' & [M_BASICALL].[名称]) AS 適用法, IIf(IsNull([小作形態]) Or [小作形態]>9 Or [小作形態]=0,3 & ':その他',[小作形態] & ':' & [M_BASICALL_1].[名称]) AS 形態, Count([D:農地Info].ID) AS 筆数, Sum(Int([D:農地Info].登記簿面積)) AS 登記面積の合計, Sum(Int([D:農地Info].実面積)) AS 現況面積の合計, Sum(Int([D:農地Info].田面積)) AS 田面積の合計, Sum(Int([D:農地Info].畑面積)) AS 畑面積の合計, Sum(Int([D:農地Info].樹園地)) AS 樹園地の合計, Sum(Int([D:農地Info].採草放牧面積)) AS 採草放牧面積の合計 FROM ([D:農地Info] LEFT JOIN M_BASICALL ON [D:農地Info].小作地適用法 = M_BASICALL.ID) LEFT JOIN M_BASICALL AS M_BASICALL_1 ON [D:農地Info].小作形態 = M_BASICALL_1.ID WHERE ((([D:農地Info].自小作別)>0) AND ((M_BASICALL.Class)='適用法令') AND ((M_BASICALL_1.Class)='小作形態')) GROUP BY IIf([小作地適用法]=0 Or IsNull([小作地適用法]),4 & ':その他',[小作地適用法] & ':' & [M_BASICALL].[名称]), IIf(IsNull([小作形態]) Or [小作形態]>9 Or [小作形態]=0,3 & ':その他',[小作形態] & ':' & [M_BASICALL_1].[名称]) ORDER BY IIf([小作地適用法]=0 Or IsNull([小作地適用法]),4 & ':その他',[小作地適用法] & ':' & [M_BASICALL].[名称]), IIf(IsNull([小作形態]) Or [小作形態]>9 Or [小作形態]=0,3 & ':その他',[小作形態] & ':' & [M_BASICALL_1].[名称]);")
            Case "開く-貸借農地一覧"
                OpenSQLList("貸借農地一覧", "SELECT V_農地.ID, V_大字.大字, V_小字.小字, V_農地.地番, V_農地.一部現況, V_地目.名称 AS 登記地目名, V_現況地目.名称 AS 現況地目名, V_農地.登記簿面積, V_農地.実面積, IIf([V_農地].[農振法区分]=1,'農用地区域',IIf([V_農地].[農振法区分]=2,'農振地域',IIf([V_農地].[農振法区分]=3,'農振地域外',IIf([V_農地].[農振法区分]=4,'その他',IIf([V_農地].[農振法区分]=5,'調査中','-'))))) AS 農業振興地域, M_BASICALL_1.名称 AS 小作地適用法令, M_BASICALL.名称 AS 権利種類, V_農地.小作開始年月日, V_農地.小作終了年月日, [D:個人Info_1].氏名 AS 借人, [D:個人Info_1].住所 AS 借人住所, IIf([D:個人Info_1].[農業改善計画認定]=1,'認定農業者',IIf([D:個人Info_1].[農業改善計画認定]=2,'担い手農家',IIf([D:個人Info_1].[農業改善計画認定]=3,'農業生産法人',IIf([D:個人Info_1].[農業改善計画認定]=4,'認定農業者＋担い手農家',IIf([D:個人Info_1].[農業改善計画認定]=5,'認定農業者＋農業生産法人','なし'))))) AS 借人認定区分, [D:個人Info].氏名 AS 貸人, [D:個人Info].住所 AS 貸人住所, IIf([D:個人Info].[農業改善計画認定]=1,'認定農業者',IIf([D:個人Info].[農業改善計画認定]=2,'担い手農家',IIf([D:個人Info].[農業改善計画認定]=3,'農業生産法人',IIf([D:個人Info].[農業改善計画認定]=4,'認定農業者＋担い手農家',IIf([D:個人Info].[農業改善計画認定]=5,'認定農業者＋農業生産法人','なし'))))) AS 貸人認定区分, [D:個人Info_2].氏名 AS 経営農業生産法人 FROM ((((((((V_農地 LEFT JOIN V_大字 ON V_農地.大字ID = V_大字.ID) LEFT JOIN V_小字 ON V_農地.小字ID = V_小字.ID) LEFT JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON V_農地.現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON V_農地.経由農業生産法人ID = [D:個人Info_2].ID) LEFT JOIN M_BASICALL ON V_農地.小作形態 = M_BASICALL.ID) LEFT JOIN M_BASICALL AS M_BASICALL_1 ON V_農地.小作地適用法 = M_BASICALL_1.ID WHERE (((V_農地.自小作別)>0) AND ((M_BASICALL.Class)='小作形態') AND ((M_BASICALL_1.Class)='適用法令'));")
            Case "開く-貸付希望地一覧"
                Dim sWhere As String = "([貸付希望]=True)"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-売渡希望地一覧"
                Dim sWhere As String = "([売渡希望]=True)"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-死亡借受農地一覧"
                OpenSQLList("死亡借受農地一覧", String.Format("SELECT '農地.' & [V_農地].[ID] AS [Key], V_農地.土地所在, [D:個人Info].[氏名] AS 所有者名, [D:個人Info].[住所] AS 所有者住所, [D:個人Info_1].[氏名] AS 借受人名, [D:個人Info_1].[住所] AS 借受人住所 FROM (V_農地 LEFT JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.借受人ID = [D:個人Info_1].ID WHERE ((([D:個人Info_1].住民区分)={0})) ORDER BY [D:個人Info_1].ID, V_農地.大字ID, IIf(InStr([地番],'-')>0,Left([地番],InStr([地番],'-')-1),[地番]), IIf(InStr([地番],'-')>0,Mid([地番],InStr([地番],'-')+1),'');", SysAD.DB(sLRDB).DBProperty("死亡住民コード")))
            Case "開く-貸借地所有集計"
                OpenSQLList("貸借地所有集計", "SELECT '個人.' & [所有個人].[ID] AS [Key], 所有個人.氏名, 所有個人.住所, Sum(Int(V_農地.田面積)) AS 小作田面積計, Sum(Int(V_農地.畑面積)) AS 小作畑面積計, Sum(Int(V_農地.樹園地)) AS 小作樹園地計, Sum(Int(V_農地.採草放牧面積)) AS 小作採草放牧面積計 FROM V_農地 INNER JOIN [D:個人Info] AS 所有個人 ON V_農地.管理人ID = 所有個人.ID WHERE (((V_農地.自小作別)>0)) GROUP BY '個人.' & [所有個人].[ID], 所有個人.氏名, 所有個人.住所;")

            Case "開く-認定農業者期間別貸借一覧"
                With New dlgInputBWDate(SysAD.GetXMLProperty("集計", "認定農業者期間別貸借一覧-開始", Now.Date))
                    If .ShowDialog = DialogResult.OK Then
                        OpenSQLList("認定農業者期間別貸借一覧", String.Format("SELECT '農地.' & [V_農地].[ID] AS [Key], V_農地.土地所在, V_地目.名称 AS 登記地目名, V_現況地目.名称 AS 現況地目名, V_農地.登記簿面積, V_農地.実面積, [D:個人Info_1].氏名 AS 貸人, [D:個人Info].認定番号, [D:個人Info].氏名 AS 認定農業者, V_農地.小作開始年月日 AS 始期, V_農地.小作終了年月日 AS 終期 FROM (((V_農地 LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON V_農地.現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID WHERE (((V_農地.小作開始年月日)>=#{0}/{1}/{2}#) AND ((V_農地.小作終了年月日)<=#{3}/{4}/{5}#) AND ((V_農地.小作地適用法)=2) AND ((V_農地.自小作別)=0) AND (([D:個人Info_1].農業改善計画認定)=1)) ORDER BY V_農地.大字ID, V_農地.小字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番]));;", .StartDate.Month, .StartDate.Day, .StartDate.Year, .EndDate.Month, .EndDate.Day, .EndDate.Year))

                        SysAD.SetXMLProperty("集計", "認定農業者期間別貸借一覧-開始", .StartDate)
                    End If
                End With
            Case "開く-認定農業者農地明細"
                OpenSQLList("認定農業者農地明細", "SELECT '農地.' & [V_農地].[ID] AS [Key], V_農地.土地所在, V_地目.名称 AS 登記地目名, V_現況地目.名称 AS 現況地目名, V_農地.登記簿面積, V_農地.実面積, [D:個人Info_1].氏名 AS 貸人, [D:個人Info].氏名 AS 所有者, V_農地.小作開始年月日, V_農地.小作終了年月日, V_農地.小作料, V_農地.小作料単位, V_農地.田面積, V_農地.畑面積, V_農地.樹園地, V_農地.採草放牧面積 FROM (((V_農地 LEFT JOIN V_地目 ON V_農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON V_農地.現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON V_農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID WHERE (((V_農地.自小作別)=0) AND (([D:個人Info].農業改善計画認定)=1)) OR ((([D:個人Info_1].農業改善計画認定)=1)) ORDER BY V_農地.大字ID, V_農地.小字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番]));")
            Case "開く-認定農業者別経営面積"
                OpenSQLList("認定農業者農地明細", "SELECT '個人.' & [D:個人Info].[ID] AS [Key], [D:個人Info].農業改善計画認定, V_農業改善計画認定項目.名称, [D:個人Info].[フリガナ], [D:個人Info].氏名, Sum(IIf([自小作別]=0,IIf([農地状況]<20,[田面積],0),0)) AS 自作田, Sum(IIf([自小作別]=0,IIf([農地状況]<20,[畑面積],0),0)) AS 自作畑, Sum(IIf([自小作別]=0,IIf([農地状況]<20,[田面積]+[畑面積],0),0)) AS 自作計, Sum(IIf([自小作別]>0,IIf([農地状況]<20,[田面積],0),0)) AS 借受田, Sum(IIf([自小作別]>0,IIf([農地状況]<20,[畑面積],0),0)) AS 借受畑, Sum(IIf([自小作別]>0,IIf([農地状況]<20,[田面積]+[畑面積],0),0)) AS 借受計, Sum(IIf([農地状況]<20,[田面積]+[畑面積],0)) AS 耕作計 FROM (V_農地 INNER JOIN [D:個人Info] ON V_農地.耕作者ID = [D:個人Info].ID) INNER JOIN V_農業改善計画認定項目 ON [D:個人Info].農業改善計画認定 = V_農業改善計画認定項目.ID GROUP BY '個人.' & [D:個人Info].[ID], [D:個人Info].農業改善計画認定, V_農業改善計画認定項目.名称, [D:個人Info].[フリガナ], [D:個人Info].氏名 HAVING ((([D:個人Info].農業改善計画認定)>0)) ORDER BY [D:個人Info].農業改善計画認定, [D:個人Info].[フリガナ];")
            Case "開く-担い手区分別経営面積"
                OpenSQLList("担い手区分農地明細", "SELECT '個人.' & [D:個人Info].[ID] AS [Key], [D:個人Info].担い手等の区分, IIf([担い手等の区分]=1,'認定農業者',IIf([担い手等の区分]=2,'新規就農者',IIf([担い手等の区分]=3,'水準到達者',IIf([担い手等の区分]=4,'特定農業団体',IIf([担い手等の区分]=5,'集落営農組織',IIf([担い手等の区分]=6,'育成予定農業者',IIf([担い手等の区分]=7,'農外参入企業',IIf([担い手等の区分]=8,'地域内農業者',IIf([担い手等の区分]=9,'地域外参入者','調査中'))))))))) AS 名称, [D:個人Info].[フリガナ], [D:個人Info].氏名, Sum(IIf([自小作別]=0,IIf([農地状況]<20,[田面積],0),0)) AS 自作田, Sum(IIf([自小作別]=0,IIf([農地状況]<20,[畑面積],0),0)) AS 自作畑, Sum(IIf([自小作別]=0,IIf([農地状況]<20,[田面積]+[畑面積],0),0)) AS 自作計, Sum(IIf([自小作別]>0,IIf([農地状況]<20,[田面積],0),0)) AS 借受田, Sum(IIf([自小作別]>0,IIf([農地状況]<20,[畑面積],0),0)) AS 借受畑, Sum(IIf([自小作別]>0,IIf([農地状況]<20,[田面積]+[畑面積],0),0)) AS 借受計, Sum(IIf([農地状況]<20,[田面積]+[畑面積],0)) AS 耕作計 FROM V_農地 INNER JOIN [D:個人Info] ON V_農地.耕作者ID = [D:個人Info].ID GROUP BY '個人.' & [D:個人Info].[ID], [D:個人Info].担い手等の区分, IIf([担い手等の区分]=1,'認定農業者',IIf([担い手等の区分]=2,'新規就農者',IIf([担い手等の区分]=3,'水準到達者',IIf([担い手等の区分]=4,'特定農業団体',IIf([担い手等の区分]=5,'集落営農組織',IIf([担い手等の区分]=6,'育成予定農業者',IIf([担い手等の区分]=7,'農外参入企業',IIf([担い手等の区分]=8,'地域内農業者',IIf([担い手等の区分]=9,'地域外参入者','調査中'))))))))), [D:個人Info].[フリガナ], [D:個人Info].氏名 HAVING ((([D:個人Info].担い手等の区分)>0)) ORDER BY [D:個人Info].担い手等の区分, [D:個人Info].[フリガナ];")
            Case "開く-集落別選挙有資格世帯数"
                'OpenSQLList("集落別選挙有資格世帯数", "")
            Case "開く-集落別選挙有資格者数"
                OpenSQLList("集落別選挙有資格者数", "SELECT V_行政区.ID AS コード, V_行政区.行政区, Sum(IIf([性別]=0,1,0)) AS 男, Sum(IIf([性別]=1,1,0)) AS 女, Sum(1) AS 計 FROM ([D:個人Info] INNER JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) INNER JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID WHERE ((([D:個人Info].選挙権の有無)=True) AND (([D:個人Info].住民区分)=0)) GROUP BY V_行政区.ID, V_行政区.行政区 ORDER BY V_行政区.ID;")

            Case "開く-遊休農地一覧"
                OpenSQLList("遊休農地一覧", "SELECT IIf([利用状況調査荒廃] Is Null Or [利用状況調査荒廃]=0,'その他',IIf([利用状況調査荒廃]=1,'A分類','B分類')) AS 荒廃農地調査, Sum(Int([登記簿面積])) & '(' & Val(' ' & Count([D:農地Info].[ID])) & ' 筆)' AS 登記面積の合計, Sum(Int([実面積])) AS 現況面積の合計, Sum(Int([田面積])) AS 田面積の合計, Sum(Int([畑面積])) AS 畑面積の合計, Sum(Int([樹園地])) AS 樹園地の合計, Sum(Int([採草放牧面積])) AS 採草放牧面積の合計 FROM(V_農地) GROUP BY IIf([利用状況調査荒廃] Is Null Or [利用状況調査荒廃]=0,'その他',IIf([利用状況調査荒廃]=1,'A分類','B分類'));")
            Case "開く-遊休A分類"
                Dim sWhere As String = "([利用状況調査荒廃]=1)"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-遊休B分類"
                Dim sWhere As String = "([利用状況調査荒廃]=2)"
                SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
            Case "開く-遊休その他"
                If MsgBox("表示に時間がかかる恐れがあります。よろしいですか？", vbOKCancel) = vbOK Then
                    Dim sWhere As String = "([利用状況調査荒廃]=0) OR ([利用状況調査荒廃] Is Null)"
                    SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
                End If
            Case "開く-遊休農地調査年度指定"
                Dim St As String = Trim(InpuText("指定年度", "年度を指定してください", HimTools2012.DateFunctions.西暦年度(Now())))
                If Len(St) = 0 Then
                ElseIf IsDate(St & "/1/1") Then
                    Dim sWhere As String = "([利用状況調査荒廃]=1) AND ([利用状況調査日] > #1/1/" & St & "#) And ([利用状況調査日] < #12/31/" & St & "#)"
                    SysAD.page農家世帯.農地リスト.検索開始(sWhere, sWhere)
                End If
            Case "開く-大字/地区別遊休地集計"
                OpenSQLList("大字/地区別遊休地集計", "SELECT '大字.' & [V_大字].ID AS [KEY],[V_大字].名称, Count([D:農地Info].ID) AS 農地筆数, Sum(IIf([田面積]>0,1,0)) AS 全田筆数, Sum(IIf([畑面積]>0,1,0)) AS 全畑筆数, Sum(Int([田面積])+Int([畑面積])) AS 全農地面積, Sum(Int([田面積])) AS 田合計, Sum(Int([畑面積])) AS 畑合計, Sum(IIf([遊休化],1,0)) AS 遊休化農地筆数, Sum(IIf([遊休化],-1,0)*([田面積]>0)) AS 遊休化田筆数, Sum(IIf([遊休化],-1,0)*([畑面積]>0)) AS 遊休化畑筆数, Sum(Int([田面積]+[畑面積])*IIf([遊休化],1,0)) AS 遊休化面積合計, Sum(Int([田面積])*IIf([遊休化],1,0)) AS 遊休化田面積合計, Sum(Int([畑面積])*IIf([遊休化],1,0)) AS 遊休化畑面積合計 FROM [D:農地Info] INNER JOIN [V_大字] ON [D:農地Info].大字ID = [V_大字].ID WHERE ((([D:農地Info].農委地目ID)=1 Or ([D:農地Info].農委地目ID)=2)) GROUP BY [V_大字].ID, [V_大字].名称 HAVING ((([V_大字].ID)>0)) ORDER BY [V_大字].ID;")
            Case "開く-農地状況別農地一覧"
                OpenSQLList("農地状況別農地一覧", "SELECT Format([M_BASICALL].[ID],'00') & ':' & [M_BASICALL].[名称] AS 農地状況別, V_農地.ID As 農地ID, V_農地.土地所在, V_農地.田面積, V_農地.畑面積, V_農地.樹園地, V_農地.採草放牧面積, [D:個人Info].氏名 AS 所有者, [D:個人Info].住所 AS 所有者住所, V_住民区分.名称 AS 所有者区分 FROM ((V_農地 INNER JOIN [D:個人Info] ON V_農地.所有者ID = [D:個人Info].ID) INNER JOIN M_BASICALL ON V_農地.農地状況 = M_BASICALL.ID) LEFT JOIN V_住民区分 ON [D:個人Info].住民区分 = V_住民区分.ID WHERE (((V_農地.農地状況)<>0 Or (V_農地.農地状況) Is Not Null) AND ((M_BASICALL.Class)='農地状況')) ORDER BY M_BASICALL.ID, V_農地.土地所在;")

            Case "開く-世帯主未設定エラー"
                OpenSQLList("世帯主未設定エラー", "SELECT '農家.' & [D:世帯Info].[ID] AS [Key], [D:世帯Info].世帯主ID, '不明' AS 世帯主名, [D:世帯Info].住所, [D:個人Info].電話番号, V_行政区.行政区, [D:世帯Info].[あっせん希望種別], [D:世帯Info].確認日時, [D:世帯Info].更新日, Count(V_農地.ID) AS 筆数 FROM (([D:世帯Info] LEFT JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) LEFT JOIN V_農地 ON [D:世帯Info].ID = V_農地.耕作世帯ID WHERE ((([D:個人Info].ID) Is Null)) GROUP BY '農家.' & [D:世帯Info].[ID], [D:世帯Info].世帯主ID, '不明', [D:世帯Info].住所, [D:個人Info].電話番号, V_行政区.行政区, [D:世帯Info].[あっせん希望種別], [D:世帯Info].確認日時, [D:世帯Info].更新日;")
            Case "開く-世帯未設定エラー"
                OpenSQLList("世帯未設定エラー", "SELECT '個人.' & [D:個人Info].[ID] AS [Key], [D:個人Info].世帯ID, [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].生年月日, M_BASICALL.名称 AS 性別, V_行政区.行政区, IIf([農年受給の有無]=-1,'有','無') AS 農年受給有無, IIf([老齢受給の有無]=-1,'有','無') AS 老齢受給有無, IIf([経営移譲の有無]=-1,'有','無') AS 経営移譲受給有無, Count(V_農地.ID) AS 筆数 FROM ((([D:個人Info] LEFT JOIN [D:世帯Info] ON [D:個人Info].世帯ID = [D:世帯Info].ID) LEFT JOIN V_行政区 ON [D:個人Info].行政区ID = V_行政区.ID) LEFT JOIN M_BASICALL ON [D:個人Info].性別 = M_BASICALL.ID) LEFT JOIN V_農地 ON [D:個人Info].ID = V_農地.耕作者ID WHERE (((M_BASICALL.Class)='性別') AND (([D:世帯Info].ID) Is Null Or ([D:世帯Info].ID)=0)) GROUP BY '個人.' & [D:個人Info].[ID], [D:個人Info].世帯ID, [D:個人Info].[フリガナ], [D:個人Info].氏名, [D:個人Info].住所, [D:個人Info].生年月日, M_BASICALL.名称, V_行政区.行政区, IIf([農年受給の有無]=-1,'有','無'), IIf([老齢受給の有無]=-1,'有','無'), IIf([経営移譲の有無]=-1,'有','無');")
            Case "開く-農地異動申請別一覧"
                If Not SysAD.page農家世帯.TabPageContainKey("農地異動申請別一覧", True) Then
                    SysAD.page農家世帯.中央Tab.AddPage(New C農地異動申請別一覧)
                End If
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Debug.Print(sCommand & "-" & Me.Key.DataClass)
                    Stop
                End If
        End Select
        Return Nothing
    End Function
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return Nothing
        End Get
    End Property

    Private Sub Act農振地内異動済み農地集計()
        Dim pCls As New InputStartAndEndDate
        Try

            With New HimTools2012.PropertyGridDialog(pCls, "農振地内異動済み農地集計", "一覧を作成する申請が許可された期間を入力してください。")
                If Not .ShowDialog() = DialogResult.Cancel Then

                    If Not SysAD.page農家世帯.TabPageContainKey("農振地内異動済み農地集計") Then
                        Dim pPage As New HimTools2012.controls.CTabPageWithToolStrip(True, True, "農振地内異動済み農地集計", "農振地内異動済み農地集計", HimTools2012.controls.CloseMode.NoMessage)
                        Dim pGrid As New HimTools2012.controls.DataGridViewWithDataView
                        Dim pProg As New ToolStripProgressBar
                        Dim pLabel As New ToolStripLabel
                        pPage.ControlPanel.Add(pGrid)
                        pGrid.Createエクセル出力Ctrl(pPage.ToolStrip)

                        pPage.ToolStrip.Items.Add(pProg)
                        pPage.ToolStrip.Items.Add(pLabel)
                        SysAD.page農家世帯.中央Tab.AddPage(pPage)

                        Application.DoEvents()
                        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_申請.許可年月日, D_申請.法令, D_申請.権利種類, D_申請.名称, D_申請.状態, D_申請.農地リスト FROM D_申請 WHERE (((D_申請.許可年月日)>#{1}/{2}/{0}# And (D_申請.許可年月日)<#{4}/{5}/{3}#) AND ((D_申請.状態)=2));", pCls.開始日.Year, pCls.開始日.Month, pCls.開始日.Day, pCls.終了日.Year, pCls.終了日.Month, pCls.終了日.Day)
                        Dim pViewTBL As New DataTable
                        pViewTBL.Columns.Add("法令", GetType(Integer))
                        pViewTBL.Columns.Add("名称", GetType(String))
                        pViewTBL.Columns.Add("筆数", GetType(Integer))
                        pViewTBL.Columns.Add("面積", GetType(Decimal))
                        pViewTBL.PrimaryKey = New DataColumn() {pViewTBL.Columns("法令")}

                        pProg.Minimum = 0
                        pProg.Maximum = pTBL.Rows.Count

                        For Each pRow As DataRow In pTBL.Rows
                            Dim s法令 As String = ""
                            Dim n法令 As Integer = pRow.Item("法令")
                            Select Case n法令
                                Case "30" : s法令 = "農地法3条所有権移転"
                                Case "31" : s法令 = "農地法3条貸借"
                                Case "311" : s法令 = "農地法3条の3第1項"
                                Case "40" : s法令 = "農地法4条"
                                Case "50" : s法令 = "農地法5条所有権移転"
                                Case "51" : s法令 = "農地法5条貸借"
                                Case "52" : s法令 = "農地法5条一時転用"
                                Case "60" : s法令 = "基盤強化法所有権移転"
                                Case "61"
                                    Select Case Val(pRow.Item("権利種類"))
                                        Case 1
                                            s法令 = "基盤強化法利用権設定(賃貸借)"
                                        Case 2
                                            n法令 = 62
                                            s法令 = "基盤強化法利用権設定(使用貸借)"
                                        Case Else
                                            n法令 = 63
                                            s法令 = "基盤強化法利用権設定(その他)"
                                    End Select
                                Case "180" : s法令 = "農地法18条解約"
                                Case "210" : s法令 = "合意解約"
                                Case "500"
                                Case "602" : s法令 = "非農地証明願"
                                Case "801"
                                Case Else
                            End Select

                            If s法令.Length > 0 Then
                                Dim xRow As DataRow = pViewTBL.Rows.Find(n法令)
                                If xRow Is Nothing Then
                                    xRow = pViewTBL.NewRow()
                                    xRow.Item("法令") = n法令
                                    xRow.Item("名称") = s法令
                                    xRow.Item("筆数") = 0
                                    xRow.Item("面積") = 0
                                    pViewTBL.Rows.Add(xRow)
                                End If
                                Try
                                    If Not IsDBNull(pRow.Item("農地リスト")) AndAlso pRow.Item("農地リスト").ToString.Length > 0 Then
                                        Dim Ar As String() = Split(pRow.Item("農地リスト"), ";")
                                        Dim 農地List As New List(Of String)
                                        Dim 転用農地List As New List(Of String)
                                        For Each St As String In Ar
                                            If St.Length Then
                                                Select Case Left(St, InStr(St, ".") - 1)
                                                    Case "農地"
                                                        農地List.Add(Mid(St, InStr(St, ".") + 1))
                                                    Case "転用農地"
                                                        転用農地List.Add(Mid(St, InStr(St, ".") + 1))
                                                End Select
                                            End If
                                        Next
                                        If 農地List.Count > 0 Then
                                            Dim p農地TBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Count([D:農地Info].ID) AS カウント, Sum([D:農地Info].登記簿面積) AS 登記簿面積の合計 FROM [D:農地Info] WHERE ((([D:農地Info].農業振興地域)=0 Or ([D:農地Info].農業振興地域)=1) AND (([D:農地Info].ID) In ({0})));", Join(農地List.ToArray, ","))
                                            If Not IsDBNull(p農地TBL.Rows(0).Item("登記簿面積の合計")) Then
                                                xRow.Item("筆数") = xRow.Item("筆数") + p農地TBL.Rows(0).Item("カウント")
                                                xRow.Item("面積") = xRow.Item("面積") + p農地TBL.Rows(0).Item("登記簿面積の合計")
                                            End If
                                        End If
                                        If 転用農地List.Count > 0 Then
                                            Dim p転用農地TBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Count(ID) AS カウント, Sum(登記簿面積) AS 登記簿面積の合計 FROM [D_転用農地] WHERE ((([D_転用農地].農業振興地域)=0 Or ([D_転用農地].農業振興地域)=1) AND (([D_転用農地].ID) In ({0})));", Join(転用農地List.ToArray, ","))
                                            If Not IsDBNull(p転用農地TBL.Rows(0).Item("登記簿面積の合計")) Then
                                                xRow.Item("筆数") = xRow.Item("筆数") + p転用農地TBL.Rows(0).Item("カウント")
                                                xRow.Item("面積") = xRow.Item("面積") + p転用農地TBL.Rows(0).Item("登記簿面積の合計")
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception
                                    Stop
                                End Try
                            End If

                            pProg.Value += 1
                            pLabel.Text = String.Format("{0} / {1}", pProg.Value, pProg.Maximum)
                            Application.DoEvents()
                        Next
                        pPage.ToolStrip.Items.Remove(pLabel)
                        pPage.ToolStrip.Items.Remove(pProg)
                        pGrid.SetDataView(pViewTBL, "[筆数]>0", "[法令]")
                    End If
                End If
            End With
        Catch ex As Exception
            Stop
        End Try
    End Sub
End Class

Public Class CTabPage貸借農地終期期間別一覧
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private pTBL農地 As DataTable
    Private pTBL転用農地 As DataTable
    Private pTBL申請 As DataTable
    Private pTBLResult As DataTable
    Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents pBtnExcel As New ToolStripButton("Excel出力")

    Public Sub New(ByVal 表示区分 As String, ByVal pStartDate As String, ByVal pEndDate As String)
        MyBase.New(True, True, 表示区分, 表示区分)

        Try
            Me.ToolStrip.Items.AddRange({pBtnExcel})
            ControlPanel.Add(mvarGrid)

            pTBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID, [D:農地Info].大字ID, V_大字.大字 AS 大字名, V_小字.小字 AS 小字名, [D:農地Info].地番, [D:農地Info].一部現況, V_地目.名称 AS 登記地目名, [D:農地Info].現況地目, V_現況地目.名称 AS 現況地目名, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].田面積, [D:農地Info].畑面積, [D:農地Info].農業振興地域, [D:農地Info].小作開始年月日, [D:農地Info].小作終了年月日, [D:個人Info_1].氏名 AS 借人, [D:個人Info_1].住所 AS 借人住所, [D:個人Info].氏名 AS 貸人, [D:個人Info].住所 AS 貸人住所, [D:個人Info_2].氏名 AS 経由農業生産法人, [D:農地Info].自小作別, [D:農地Info].小作料, [D:農地Info].小作料単位 FROM (((((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地Info].借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON [D:農地Info].経由農業生産法人ID = [D:個人Info_2].ID;")
            pTBL農地.PrimaryKey = {pTBL農地.Columns("ID")}
            'App農地基本台帳.TBL農地.MergePlus(pTBL農地)
            pTBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_転用農地.ID, D_転用農地.大字ID, V_大字.大字 AS 大字名, V_小字.小字 AS 小字名, D_転用農地.地番, D_転用農地.一部現況, V_地目.名称 AS 登記地目名, D_転用農地.現況地目, V_現況地目.名称 AS 現況地目名, D_転用農地.登記簿面積, D_転用農地.実面積, D_転用農地.田面積, D_転用農地.畑面積, D_転用農地.農業振興地域, D_転用農地.小作開始年月日, D_転用農地.小作終了年月日, [D:個人Info_1].氏名 AS 借人, [D:個人Info_1].住所 AS 借人住所, [D:個人Info].氏名 AS 貸人, [D:個人Info].住所 AS 貸人住所, [D:個人Info_2].氏名 AS 経由農業生産法人, D_転用農地.自小作別, D_転用農地.小作料, D_転用農地.小作料単位 FROM ((((((D_転用農地 LEFT JOIN V_大字 ON D_転用農地.大字ID = V_大字.ID) LEFT JOIN V_小字 ON D_転用農地.小字ID = V_小字.ID) LEFT JOIN V_地目 ON D_転用農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON D_転用農地.現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON D_転用農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON D_転用農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON D_転用農地.経由農業生産法人ID = [D:個人Info_2].ID;")
            pTBL転用農地.PrimaryKey = {pTBL転用農地.Columns("ID")}
            'App農地基本台帳.TBL転用農地.MergePlus(pTBL転用農地)

            '↓↓↓日付のWHERE
            Dim sWhere As String = ""
            Dim sTime As String = ""
            Select Case 表示区分
                Case "貸借農地一覧"
                    sWhere = "31,61,62"
                    sTime = "終期"
                Case "利用権設定終期期間別一覧"
                    sWhere = "61,62"
                    sTime = "終期"
                Case "利用権設定始期期間別一覧"
                    sWhere = "61,62"
                    sTime = "始期"
            End Select

            pTBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect(String.Format("SELECT D_申請.ID, D_申請.名称, D_申請.許可年月日, D_申請.法令, D_申請.農地リスト, D_申請.始期, D_申請.終期, D_申請.小作料, D_申請.小作料単位, D_申請.氏名A, D_申請.住所A, [D:個人Info].[農業改善計画認定] AS 認定農家貸人, D_申請.氏名B, D_申請.住所B, [D:個人Info_1].[農業改善計画認定] AS 認定農家借人, D_申請.権利種類, D_申請.申請者A, D_申請.申請者B, D_申請.申請者C, D_申請.経由法人ID, D_申請.代理人A, D_申請.再設定 FROM (D_申請 LEFT JOIN [D:個人Info] ON D_申請.申請者A = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON D_申請.申請者B = [D:個人Info_1].ID WHERE (((D_申請.法令) In ({0})) AND ((D_申請.終期) Is Not Null)) AND ((D_申請.{3}) >={1}) AND ((D_申請.{3}) <={2});", sWhere, pStartDate, pEndDate, sTime))
            pTBL申請.PrimaryKey = {pTBL申請.Columns("ID")}
            'App農地基本台帳.TBL申請.MergePlus(pTBL申請)
            pTBLResult = New DataTable
            CreatepTBLResult(pTBLResult)

            For Each pRow As DataRow In pTBL申請.Rows
                Dim Ar筆リスト As Object = Split(pRow.Item("農地リスト").ToString, ";")
                For n As Integer = 0 To UBound(Ar筆リスト)
                    Dim Ar筆情報 As Object = Split(Ar筆リスト(n), ".")
                    Dim pRowFind As DataRow = Nothing

                    If InStr(Ar筆情報(0), "転用農地") > 0 Then
                        If IsNumeric(Ar筆情報(1)) Then
                            pRowFind = pTBL転用農地.Rows.Find(CDec(Ar筆情報(1)))
                            If pRowFind Is Nothing Then
                                pRowFind = pTBL農地.Rows.Find(CDec(Ar筆情報(1)))
                            End If
                        End If
                    ElseIf InStr(Ar筆情報(0), "農地") > 0 Then
                        If IsNumeric(Ar筆情報(1)) Then
                            pRowFind = pTBL農地.Rows.Find(CDec(Ar筆情報(1)))
                            If pRowFind Is Nothing Then
                                pRowFind = pTBL転用農地.Rows.Find(CDec(Ar筆情報(1)))
                            End If
                        End If
                    End If

                    If Not pRowFind Is Nothing Then
                        Dim pAddRow As DataRow = pTBLResult.NewRow

                        pAddRow.Item("申請ID") = pRow.Item("ID")
                        pAddRow.Item("申請内容") = pRow.Item("名称")
                        pAddRow.Item("許可年月日") = pRow.Item("許可年月日")
                        pAddRow.Item("農地ID") = pRowFind.Item("ID")
                        pAddRow.Item("大字ID") = pRowFind.Item("大字ID")
                        pAddRow.Item("大字") = pRowFind.Item("大字名")
                        pAddRow.Item("小字") = pRowFind.Item("小字名")
                        pAddRow.Item("地番") = pRowFind.Item("地番")
                        pAddRow.Item("一部現況") = pRowFind.Item("一部現況")
                        pAddRow.Item("登記地目名") = pRowFind.Item("登記地目名")
                        pAddRow.Item("現況地目名") = pRowFind.Item("現況地目名")
                        pAddRow.Item("登記簿面積") = pRowFind.Item("登記簿面積")
                        pAddRow.Item("実面積") = pRowFind.Item("実面積")
                        pAddRow.Item("田面積") = pRowFind.Item("田面積")
                        pAddRow.Item("畑面積") = pRowFind.Item("畑面積")
                        pAddRow.Item("農業振興地域") = IIF(Val(pRowFind.Item("農業振興地域").ToString) = 0, "農振白地", IIF(Val(pRowFind.Item("農業振興地域").ToString) = 1, "農用地", "農振外"))
                        'pAddRow.Item("農振法") = IIf(pRowFind.Item("農振法") = 1, "農用地区域", IIf(pRowFind.Item("農振法") = 2, "農振地域", "農振地域外"))
                        pAddRow.Item("権利種類") = IIF(Val(pRow.Item("権利種類").ToString) = 1, "貸借権", IIF(Val(pRow.Item("権利種類").ToString) = 2, "使用貸借", "その他"))
                        pAddRow.Item("小作開始年月日") = pRow.Item("始期")
                        pAddRow.Item("小作終了年月日") = pRow.Item("終期")

                        If Not IsDBNull(pRow.Item("始期")) AndAlso Not IsDBNull(pRow.Item("終期")) Then
                            pAddRow.Item("期間") = IIF(DateDiff("yyyy", pRow.Item("始期"), pRow.Item("終期")) < 1, "1年未満", DateDiff("yyyy", pRow.Item("始期"), pRow.Item("終期")) & "年")
                        Else
                            pAddRow.Item("期間") = ""
                        End If

                        pAddRow.Item("小作料") = pRowFind.Item("小作料").ToString
                        pAddRow.Item("小作料単位") = pRowFind.Item("小作料単位").ToString
                        pAddRow.Item("借人") = pRow.Item("氏名B")
                        pAddRow.Item("借人住所") = pRow.Item("住所B")
                        pAddRow.Item("借人認定区分") = IIF(Val(pRow.Item("認定農家借人").ToString) = 1, "認定農業者", IIF(Val(pRow.Item("認定農家借人").ToString) = 2, "担い手農家", IIF(Val(pRow.Item("認定農家借人").ToString) = 3, "農業生産法人", IIF(Val(pRow.Item("認定農家借人").ToString) = 4, "認定農業者＋担い手農家", IIF(Val(pRow.Item("認定農家借人").ToString) = 5, "認定農業者＋農業生産法人", "なし")))))
                        pAddRow.Item("貸人") = pRow.Item("氏名A")
                        pAddRow.Item("貸人住所") = pRow.Item("住所A")
                        pAddRow.Item("貸人認定区分") = IIF(Val(pRow.Item("認定農家貸人").ToString) = 1, "認定農業者", IIF(Val(pRow.Item("認定農家貸人").ToString) = 2, "担い手農家", IIF(Val(pRow.Item("認定農家貸人").ToString) = 3, "農業生産法人", IIF(Val(pRow.Item("認定農家貸人").ToString) = 4, "認定農業者＋担い手農家", IIF(Val(pRow.Item("認定農家貸人").ToString) = 5, "認定農業者＋農業生産法人", "なし")))))
                        pAddRow.Item("自小作別") = IIF(Val(pRowFind.Item("自小作別").ToString) = 0, "自作", IIF(Val(pRowFind.Item("自小作別").ToString) = 1, "小作", IIF(Val(pRowFind.Item("自小作別").ToString) = 2, "農年", "ヤミ小作")))
                        'pAddRow.Item("耕作者") = IIf(pRowFind.Item("自小作別") = 0, pRowFind.Item("貸人"), pRowFind.Item("借人"))
                        pAddRow.Item("経由農業生産法人") = pRowFind.Item("経由農業生産法人")
                        pAddRow.Item("再設定") = IIF(pRow.Item("法令") = 31, "", IIF(pRow.Item("再設定") = True, "再", "新"))

                        pTBLResult.Rows.Add(pAddRow)
                    End If
                Next
            Next

            mvarGrid.SetDataView(pTBLResult, "", "")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub CreatepTBLResult(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("申請ID", GetType(Integer))
            .Columns.Add("申請内容", GetType(String))
            .Columns.Add("許可年月日", GetType(Date))
            .Columns.Add("農地ID", GetType(Integer))
            .Columns.Add("大字ID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("一部現況", GetType(Integer))
            .Columns.Add("登記地目名", GetType(String))
            .Columns.Add("現況地目名", GetType(String))
            .Columns.Add("登記簿面積", GetType(Decimal))
            .Columns.Add("実面積", GetType(Decimal))
            .Columns.Add("田面積", GetType(Decimal))
            .Columns.Add("畑面積", GetType(Decimal))
            .Columns.Add("農業振興地域", GetType(String))
            .Columns.Add("権利種類", GetType(String))
            .Columns.Add("小作開始年月日", GetType(Date))
            .Columns.Add("小作終了年月日", GetType(Date))
            .Columns.Add("期間", GetType(String))
            .Columns.Add("小作料", GetType(String))
            .Columns.Add("小作料単位", GetType(String))
            .Columns.Add("借人", GetType(String))
            .Columns.Add("借人住所", GetType(String))
            .Columns.Add("借人認定区分", GetType(String))
            .Columns.Add("貸人", GetType(String))
            .Columns.Add("貸人住所", GetType(String))
            .Columns.Add("貸人認定区分", GetType(String))
            .Columns.Add("自小作別", GetType(String))
            '.Columns.Add("耕作者", GetType(String))
            .Columns.Add("経由農業生産法人", GetType(String))
            .Columns.Add("再設定", GetType(String))
        End With
    End Sub

    Private Sub pBtnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pBtnExcel.Click
        mvarGrid.ToExcel()
    End Sub
End Class

Public Class CTabPage転用農地一覧
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private pTBL転用農地 As DataTable
    Private pTBL地図情報 As DataTable
    Private pTBLMaxDay As DataTable
    Private pTBLResult As DataTable
    Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents pBtnExcel As New ToolStripButton("Excel出力")

    Public Sub New()
        MyBase.New(True, True, "転用(非農地)農地一覧", "転用(非農地)農地一覧")

        Try
            Me.ToolStrip.Items.AddRange({pBtnExcel})
            ControlPanel.Add(mvarGrid)

            LoadDataBase()
            pTBLResult = New DataTable
            CreatepTBLResult(pTBLResult)

            For Each pRow As DataRow In pTBL転用農地.Rows
                Dim pFindRow As DataRow = pTBLMaxDay.Rows.Find(pRow.Item("ID"))
                If Not pFindRow Is Nothing Then
                    Dim pView As DataView = New DataView(pTBL転用農地, String.Format("[ID]={0} And [異動日] =#{1}#", pRow.Item("ID"), pFindRow.Item("異動日の最大")), "", DataViewRowState.CurrentRows)
                    pFindRow = pTBLResult.Rows.Find(pRow.Item("ID"))
                    If pView.Count > 0 AndAlso pFindRow Is Nothing Then
                        Dim pAddRow As DataRow = pTBLResult.NewRow

                        pAddRow.Item("ID") = Val(pView.Item(0).Row("ID").ToString)
                        pAddRow.Item("大字") = pView.Item(0).Row("大字").ToString
                        pAddRow.Item("小字") = pView.Item(0).Row("小字").ToString
                        pAddRow.Item("所在") = pView.Item(0).Row("所在").ToString
                        pAddRow.Item("地番") = pView.Item(0).Row("地番").ToString
                        pAddRow.Item("一部現況") = Val(pView.Item(0).Row("一部現況").ToString)
                        pAddRow.Item("登記地目名") = pView.Item(0).Row("登記地目").ToString
                        pAddRow.Item("現況地目名") = pView.Item(0).Row("現況地目").ToString
                        pAddRow.Item("登記簿面積") = Val(pView.Item(0).Row("登記簿面積").ToString)
                        pAddRow.Item("実面積") = Val(pView.Item(0).Row("実面積").ToString)
                        pAddRow.Item("所有者ID") = Val(pView.Item(0).Row("所有者ID").ToString)
                        pAddRow.Item("所有者名") = pView.Item(0).Row("氏名").ToString
                        pAddRow.Item("異動日") = pView.Item(0).Row("異動日").ToString
                        pAddRow.Item("異動事由") = Val(pView.Item(0).Row("異動事由").ToString)
                        pAddRow.Item("異動事由内容") = pView.Item(0).Row("異動事由内容").ToString

                        'Dim pFindView As DataView = New DataView(pTBL地図情報, "[OAza]=", "", DataViewRowState.CurrentRows)

                        Select Case Val(pView.Item(0).Row("異動事由").ToString)
                            Case 10040, 10050 : pAddRow.Item("区分") = "転用"
                            Case 261, 18099, 30001, 100001 : pAddRow.Item("区分") = "非農地"
                            Case 891 : pAddRow.Item("区分") = "削除"
                            Case Else : pAddRow.Item("区分") = "不明"
                        End Select

                        pTBLResult.Rows.Add(pAddRow)
                    End If
                End If
            Next

            mvarGrid.SetDataView(pTBLResult, "", "")
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub LoadDataBase()
        Try
            pTBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_転用農地.ID, V_大字.名称 AS 大字, V_小字.小字, D_転用農地.所在, D_転用農地.地番, D_転用農地.一部現況, V_地目.名称 AS 登記地目, V_現況地目.名称 AS 現況地目, D_転用農地.登記簿面積, D_転用農地.実面積, D_転用農地.所有者ID, [D:個人Info].氏名, D_土地履歴.異動日, D_土地履歴.異動事由, V_土地異動事由.名称 AS 異動事由内容 FROM ((((((D_転用農地 INNER JOIN D_土地履歴 ON D_転用農地.ID = D_土地履歴.LID) LEFT JOIN V_大字 ON D_転用農地.大字ID = V_大字.ID) LEFT JOIN V_小字 ON D_転用農地.小字ID = V_小字.ID) LEFT JOIN [D:個人Info] ON D_転用農地.所有者ID = [D:個人Info].ID) LEFT JOIN V_土地異動事由 ON D_土地履歴.異動事由 = V_土地異動事由.ID) LEFT JOIN V_現況地目 ON D_転用農地.現況地目 = V_現況地目.ID) LEFT JOIN V_地目 ON D_転用農地.登記簿地目 = V_地目.ID WHERE (((D_転用農地.大字ID) Is Not Null)) ORDER BY D_転用農地.大字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])), IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),0);")
            pTBLMaxDay = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_土地履歴.LID, Max(D_土地履歴.異動日) AS 異動日の最大 FROM D_土地履歴 GROUP BY D_土地履歴.LID HAVING (((Max(D_土地履歴.異動日)) Is Not Null));")
            pTBLMaxDay.PrimaryKey = {pTBLMaxDay.Columns("LID")}

            Dim sPath As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\ComparisonDBPath.txt"
            If IO.File.Exists(sPath) Then
                Dim cReader As New System.IO.StreamReader(sPath, System.Text.Encoding.Default)
                While (cReader.Peek() >= 0)
                    Dim stBuffer As String = cReader.ReadLine() ' ファイルを 1 行ずつ読み込む
                    Dim cAr As Object = Split(stBuffer, ":")

                    Select Case cAr(0)
                        Case "地図"
                            pTBL地図情報 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:LotProperty].ID, [D:LotProperty].OAza, [D:LotProperty].Name FROM [D:LotProperty] WHERE ((([D:LotProperty].削除)=False));")
                    End Select

                End While
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub pBtnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pBtnExcel.Click
        mvarGrid.ToExcel()
    End Sub

    Private Sub CreatepTBLResult(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("ID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("所在", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("一部現況", GetType(Integer))
            .Columns.Add("登記地目名", GetType(String))
            .Columns.Add("現況地目名", GetType(String))
            .Columns.Add("登記簿面積", GetType(Decimal))
            .Columns.Add("実面積", GetType(Decimal))
            .Columns.Add("所有者ID", GetType(Decimal))
            .Columns.Add("所有者名", GetType(String))
            .Columns.Add("異動日", GetType(String))
            .Columns.Add("異動事由", GetType(Integer))
            .Columns.Add("異動事由内容", GetType(String))
            .Columns.Add("区分", GetType(String))
            .Columns.Add("地図ID", GetType(Integer))
            .PrimaryKey = { .Columns("ID")}
        End With
    End Sub
End Class

Public Class CTabPage転用農地申請履歴一覧
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private pTBL転用農地 As DataTable
    Private pTBL申請 As DataTable
    Private pTBLResult As DataTable
    Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents pBtnExcel As New ToolStripButton("Excel出力")

    Public Sub New()
        MyBase.New(True, True, "転用農地申請履歴一覧", "転用農地申請履歴一覧")

        Try
            Me.ToolStrip.Items.AddRange({pBtnExcel})
            ControlPanel.Add(mvarGrid)

            pTBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_転用農地.ID, D_転用農地.大字ID, V_大字.大字 AS 大字名, V_小字.小字 AS 小字名, D_転用農地.地番, D_転用農地.一部現況, V_地目.名称 AS 登記地目名, D_転用農地.現況地目, V_現況地目.名称 AS 現況地目名, D_転用農地.登記簿面積, D_転用農地.実面積, D_転用農地.田面積, D_転用農地.畑面積, D_転用農地.農業振興地域, D_転用農地.小作開始年月日, D_転用農地.小作終了年月日, [D:個人Info_1].氏名 AS 借人, [D:個人Info_1].住所 AS 借人住所, [D:個人Info].氏名 AS 貸人, [D:個人Info].住所 AS 貸人住所, [D:個人Info_2].氏名 AS 経由農業生産法人, D_転用農地.自小作別 FROM ((((((D_転用農地 LEFT JOIN V_大字 ON D_転用農地.大字ID = V_大字.ID) LEFT JOIN V_小字 ON D_転用農地.小字ID = V_小字.ID) LEFT JOIN V_地目 ON D_転用農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON D_転用農地.現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON D_転用農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON D_転用農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON D_転用農地.経由農業生産法人ID = [D:個人Info_2].ID;")
            pTBL転用農地.PrimaryKey = {pTBL転用農地.Columns("ID")}

            pTBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_申請;")
            pTBL申請.PrimaryKey = {pTBL申請.Columns("ID")}

            pTBLResult = New DataTable
            CreatepTBLResult(pTBLResult)

            For Each pRow As DataRow In pTBL転用農地.Rows
                Dim pView As DataView = New DataView(pTBL申請, "[農地リスト] Like '*." & pRow.Item("ID") & "*'", "", DataViewRowState.CurrentRows)
                For Each pViewRow As DataRowView In pView
                    Dim Ar As Object = Split(pViewRow.Item("農地リスト"), ";")

                    For n As Integer = 0 To UBound(Ar)
                        If pRow.Item("ID") = Val(Split(Ar(n).ToString, ".")(1)) Then
                            Dim pAddRow As DataRow = pTBLResult.NewRow

                            pAddRow.Item("ID") = Val(pRow.Item("ID").ToString)
                            pAddRow.Item("大字ID") = Val(pRow.Item("大字ID").ToString)
                            pAddRow.Item("大字") = pRow.Item("大字名").ToString
                            pAddRow.Item("小字") = pRow.Item("小字名").ToString
                            pAddRow.Item("地番") = pRow.Item("地番").ToString
                            pAddRow.Item("一部現況") = Val(pRow.Item("一部現況").ToString)
                            pAddRow.Item("登記地目名") = pRow.Item("登記地目名").ToString
                            pAddRow.Item("現況地目名") = pRow.Item("現況地目名").ToString
                            pAddRow.Item("登記簿面積") = Val(pRow.Item("登記簿面積").ToString)
                            pAddRow.Item("実面積") = Val(pRow.Item("実面積").ToString)
                            pAddRow.Item("田面積") = Val(pRow.Item("田面積").ToString)
                            pAddRow.Item("畑面積") = Val(pRow.Item("畑面積").ToString)
                            pAddRow.Item("農業振興地域") = IIF(Val(pRow.Item("農業振興地域").ToString) = 0, "農振白地", IIF(Val(pRow.Item("農業振興地域").ToString) = 1, "農用地", "農振外"))

                            '/*申請内容*/
                            pAddRow.Item("申請名称") = pViewRow.Item("名称").ToString
                            Select Case Val(pViewRow.Item("状態").ToString)
                                Case 0 : pAddRow.Item("申請状態") = "受付中"
                                Case 1 : pAddRow.Item("申請状態") = "審査中"
                                Case 2 : pAddRow.Item("申請状態") = "許可済"
                                Case 4 : pAddRow.Item("申請状態") = "取下げ"
                                Case 5 : pAddRow.Item("申請状態") = "取消し"
                                Case 42 : pAddRow.Item("申請状態") = "不許可"
                            End Select

                            pTBLResult.Rows.Add(pAddRow)
                        End If
                    Next
                Next

            Next

            mvarGrid.SetDataView(pTBLResult, "", "")
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub pBtnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pBtnExcel.Click
        mvarGrid.ToExcel()
    End Sub

    Private Sub CreatepTBLResult(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("ID", GetType(Integer))
            .Columns.Add("大字ID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("一部現況", GetType(Integer))
            .Columns.Add("登記地目名", GetType(String))
            .Columns.Add("現況地目名", GetType(String))
            .Columns.Add("登記簿面積", GetType(Decimal))
            .Columns.Add("実面積", GetType(Decimal))
            .Columns.Add("田面積", GetType(Decimal))
            .Columns.Add("畑面積", GetType(Decimal))
            .Columns.Add("農業振興地域", GetType(String))


            .Columns.Add("申請名称", GetType(String))
            .Columns.Add("申請状態", GetType(String))
        End With
    End Sub
End Class

Public Class CTabPage申請無転用農地一覧
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private pTBL転用農地 As DataTable
    Private pTBL申請 As DataTable
    Private pTBLResult As DataTable
    Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents pBtnExcel As New ToolStripButton("Excel出力")

    Public Sub New()
        MyBase.New(True, True, "申請無転用農地一覧", "申請無転用農地一覧")

        Try
            Me.ToolStrip.Items.AddRange({pBtnExcel})
            ControlPanel.Add(mvarGrid)

            pTBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_転用農地.ID, D_転用農地.大字ID, V_大字.大字 AS 大字名, V_小字.小字 AS 小字名, D_転用農地.地番, D_転用農地.一部現況, V_地目.名称 AS 登記地目名, D_転用農地.現況地目, V_現況地目.名称 AS 現況地目名, D_転用農地.登記簿面積, D_転用農地.実面積, D_転用農地.田面積, D_転用農地.畑面積, D_転用農地.農業振興地域, D_転用農地.小作開始年月日, D_転用農地.小作終了年月日, [D:個人Info_1].氏名 AS 借人, [D:個人Info_1].住所 AS 借人住所, [D:個人Info].氏名 AS 貸人, [D:個人Info].住所 AS 貸人住所, [D:個人Info_2].氏名 AS 経由農業生産法人, D_転用農地.自小作別, D_土地履歴.内容, D_土地履歴.異動日, D_土地履歴.更新日 FROM (((((((D_転用農地 LEFT JOIN V_大字 ON D_転用農地.大字ID = V_大字.ID) LEFT JOIN V_小字 ON D_転用農地.小字ID = V_小字.ID) LEFT JOIN V_地目 ON D_転用農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON D_転用農地.現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON D_転用農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON D_転用農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON D_転用農地.経由農業生産法人ID = [D:個人Info_2].ID) LEFT JOIN D_土地履歴 ON D_転用農地.ID = D_土地履歴.LID ORDER BY D_転用農地.大字ID, D_転用農地.地番;")

            pTBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_申請;")
            pTBL申請.PrimaryKey = {pTBL申請.Columns("ID")}

            pTBLResult = New DataTable
            CreatepTBLResult(pTBLResult)

            For Each pRow As DataRow In pTBL転用農地.Rows
                Dim pView As DataView = New DataView(pTBL申請, "[農地リスト] Like '*." & pRow.Item("ID") & "*'", "", DataViewRowState.CurrentRows)
                If Not pView.Count > 0 Then
                    Dim pAddRow As DataRow = pTBLResult.NewRow

                    pAddRow.Item("ID") = Val(pRow.Item("ID").ToString)
                    pAddRow.Item("大字ID") = Val(pRow.Item("大字ID").ToString)
                    pAddRow.Item("大字") = pRow.Item("大字名").ToString
                    pAddRow.Item("小字") = pRow.Item("小字名").ToString
                    pAddRow.Item("地番") = pRow.Item("地番").ToString
                    pAddRow.Item("一部現況") = Val(pRow.Item("一部現況").ToString)
                    pAddRow.Item("登記地目名") = pRow.Item("登記地目名").ToString
                    pAddRow.Item("現況地目名") = pRow.Item("現況地目名").ToString
                    pAddRow.Item("登記簿面積") = Val(pRow.Item("登記簿面積").ToString)
                    pAddRow.Item("実面積") = Val(pRow.Item("実面積").ToString)
                    pAddRow.Item("田面積") = Val(pRow.Item("田面積").ToString)
                    pAddRow.Item("畑面積") = Val(pRow.Item("畑面積").ToString)
                    pAddRow.Item("農業振興地域") = IIF(Val(pRow.Item("農業振興地域").ToString) = 0, "農振白地", IIF(Val(pRow.Item("農業振興地域").ToString) = 1, "農用地", "農振外"))

                    pAddRow.Item("履歴内容") = pRow.Item("内容").ToString
                    pAddRow.Item("異動日") = pRow.Item("異動日").ToString
                    pAddRow.Item("更新日") = pRow.Item("更新日").ToString

                    pTBLResult.Rows.Add(pAddRow)
                End If
            Next

            mvarGrid.SetDataView(pTBLResult, "", "")
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub pBtnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pBtnExcel.Click
        mvarGrid.ToExcel()
    End Sub

    Private Sub CreatepTBLResult(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("ID", GetType(Integer))
            .Columns.Add("大字ID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("一部現況", GetType(Integer))
            .Columns.Add("登記地目名", GetType(String))
            .Columns.Add("現況地目名", GetType(String))
            .Columns.Add("登記簿面積", GetType(Decimal))
            .Columns.Add("実面積", GetType(Decimal))
            .Columns.Add("田面積", GetType(Decimal))
            .Columns.Add("畑面積", GetType(Decimal))
            .Columns.Add("農業振興地域", GetType(String))


            .Columns.Add("履歴内容", GetType(String))
            .Columns.Add("異動日", GetType(String))
            .Columns.Add("更新日", GetType(String))
        End With
    End Sub
End Class

Public Class CTabPage削除農地一覧
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private pTBL削除農地 As DataTable
    Private pTBL地図情報 As DataTable
    Private pTBLMaxDay As DataTable
    Private pTBLResult As DataTable
    Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents pBtnExcel As New ToolStripButton("Excel出力")

    Public Sub New()
        MyBase.New(True, True, "削除農地一覧", "削除農地一覧")

        Try
            Me.ToolStrip.Items.AddRange({pBtnExcel})
            ControlPanel.Add(mvarGrid)

            LoadDataBase()
            pTBLResult = New DataTable
            CreatepTBLResult(pTBLResult)

            For Each pRow As DataRow In pTBL削除農地.Rows
                Dim pFindRow As DataRow = pTBLMaxDay.Rows.Find(pRow.Item("ID"))
                If Not pFindRow Is Nothing Then
                    Dim pView As DataView = New DataView(pTBL削除農地, String.Format("[ID]={0} And [異動日] =#{1}#", pRow.Item("ID"), pFindRow.Item("異動日の最大")), "", DataViewRowState.CurrentRows)
                    pFindRow = pTBLResult.Rows.Find(pRow.Item("ID"))
                    If pView.Count > 0 AndAlso pFindRow Is Nothing Then
                        Dim pAddRow As DataRow = pTBLResult.NewRow

                        pAddRow.Item("ID") = Val(pView.Item(0).Row("ID").ToString)
                        pAddRow.Item("大字") = pView.Item(0).Row("大字").ToString
                        pAddRow.Item("小字") = pView.Item(0).Row("小字").ToString
                        pAddRow.Item("所在") = pView.Item(0).Row("所在").ToString
                        pAddRow.Item("地番") = pView.Item(0).Row("地番").ToString
                        pAddRow.Item("一部現況") = Val(pView.Item(0).Row("一部現況").ToString)
                        pAddRow.Item("登記地目名") = pView.Item(0).Row("登記地目").ToString
                        pAddRow.Item("現況地目名") = pView.Item(0).Row("現況地目").ToString
                        pAddRow.Item("登記簿面積") = Val(pView.Item(0).Row("登記簿面積").ToString)
                        pAddRow.Item("実面積") = Val(pView.Item(0).Row("実面積").ToString)
                        pAddRow.Item("所有者ID") = Val(pView.Item(0).Row("所有者ID").ToString)
                        pAddRow.Item("所有者名") = pView.Item(0).Row("氏名").ToString
                        pAddRow.Item("異動日") = pView.Item(0).Row("異動日").ToString
                        pAddRow.Item("異動事由") = Val(pView.Item(0).Row("異動事由").ToString)
                        pAddRow.Item("異動事由内容") = pView.Item(0).Row("異動事由内容").ToString

                        'Dim pFindView As DataView = New DataView(pTBL地図情報, "[OAza]=", "", DataViewRowState.CurrentRows)

                        Select Case Val(pView.Item(0).Row("異動事由").ToString)
                            Case 261 : pAddRow.Item("区分") = "削除"
                            Case Else : pAddRow.Item("区分") = "不明"
                        End Select

                        pTBLResult.Rows.Add(pAddRow)
                    End If
                End If
            Next

            mvarGrid.SetDataView(pTBLResult, "", "")
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub LoadDataBase()
        Try
            pTBL削除農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_削除農地.ID, V_大字.名称 AS 大字, V_小字.小字, D_削除農地.所在, D_削除農地.地番, D_削除農地.一部現況, V_地目.名称 AS 登記地目, V_現況地目.名称 AS 現況地目, D_削除農地.登記簿面積, D_削除農地.実面積, D_削除農地.所有者ID, [D:個人Info].氏名, D_土地履歴.異動日, D_土地履歴.異動事由, V_土地異動事由.名称 AS 異動事由内容 FROM ((((((D_削除農地 INNER JOIN D_土地履歴 ON D_削除農地.ID = D_土地履歴.LID) LEFT JOIN V_大字 ON D_削除農地.大字ID = V_大字.ID) LEFT JOIN V_小字 ON D_削除農地.小字ID = V_小字.ID) LEFT JOIN [D:個人Info] ON D_削除農地.所有者ID = [D:個人Info].ID) LEFT JOIN V_土地異動事由 ON D_土地履歴.異動事由 = V_土地異動事由.ID) LEFT JOIN V_現況地目 ON D_削除農地.現況地目 = V_現況地目.ID) LEFT JOIN V_地目 ON D_削除農地.登記簿地目 = V_地目.ID WHERE (((D_削除農地.大字ID) Is Not Null)) ORDER BY D_削除農地.大字ID, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])), IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),0);")
            pTBLMaxDay = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_土地履歴.LID, Max(D_土地履歴.異動日) AS 異動日の最大 FROM D_土地履歴 GROUP BY D_土地履歴.LID HAVING (((Max(D_土地履歴.異動日)) Is Not Null));")
            pTBLMaxDay.PrimaryKey = {pTBLMaxDay.Columns("LID")}

            Dim sPath As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\ComparisonDBPath.txt"
            If IO.File.Exists(sPath) Then
                Dim cReader As New System.IO.StreamReader(sPath, System.Text.Encoding.Default)
                While (cReader.Peek() >= 0)
                    Dim stBuffer As String = cReader.ReadLine() ' ファイルを 1 行ずつ読み込む
                    Dim cAr As Object = Split(stBuffer, ":")

                    Select Case cAr(0)
                        Case "地図"
                            pTBL地図情報 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:LotProperty].ID, [D:LotProperty].OAza, [D:LotProperty].Name FROM [D:LotProperty] WHERE ((([D:LotProperty].削除)=False));")
                    End Select

                End While
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub pBtnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pBtnExcel.Click
        mvarGrid.ToExcel()
    End Sub

    Private Sub CreatepTBLResult(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("ID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("小字", GetType(String))
            .Columns.Add("所在", GetType(String))
            .Columns.Add("地番", GetType(String))
            .Columns.Add("一部現況", GetType(Integer))
            .Columns.Add("登記地目名", GetType(String))
            .Columns.Add("現況地目名", GetType(String))
            .Columns.Add("登記簿面積", GetType(Decimal))
            .Columns.Add("実面積", GetType(Decimal))
            .Columns.Add("所有者ID", GetType(Decimal))
            .Columns.Add("所有者名", GetType(String))
            .Columns.Add("異動日", GetType(String))
            .Columns.Add("異動事由", GetType(Integer))
            .Columns.Add("異動事由内容", GetType(String))
            .Columns.Add("区分", GetType(String))
            .Columns.Add("地図ID", GetType(Integer))
            .PrimaryKey = { .Columns("ID")}
        End With
    End Sub
End Class

Public Class CTabPage利用権設定大字別面積集計
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private pTBL転用農地 As DataTable
    Private pTBL申請 As DataTable
    Private pTBLResult As DataTable
    Private mvarGrid As New HimTools2012.controls.DataGridViewWithDataView
    Private WithEvents pBtnExcel As New ToolStripButton("Excel出力")

    Public Sub New()
        MyBase.New(True, True, "申請無転用農地一覧", "申請無転用農地一覧")

        Try
            Me.ToolStrip.Items.AddRange({pBtnExcel})
            ControlPanel.Add(mvarGrid)

            pTBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_転用農地.ID, D_転用農地.大字ID, V_大字.大字 AS 大字名, V_小字.小字 AS 小字名, D_転用農地.地番, D_転用農地.一部現況, V_地目.名称 AS 登記地目名, D_転用農地.現況地目, V_現況地目.名称 AS 現況地目名, D_転用農地.登記簿面積, D_転用農地.実面積, D_転用農地.田面積, D_転用農地.畑面積, D_転用農地.農業振興地域, D_転用農地.小作開始年月日, D_転用農地.小作終了年月日, [D:個人Info_1].氏名 AS 借人, [D:個人Info_1].住所 AS 借人住所, [D:個人Info].氏名 AS 貸人, [D:個人Info].住所 AS 貸人住所, [D:個人Info_2].氏名 AS 経由農業生産法人, D_転用農地.自小作別, D_土地履歴.内容, D_土地履歴.異動日, D_土地履歴.更新日 FROM (((((((D_転用農地 LEFT JOIN V_大字 ON D_転用農地.大字ID = V_大字.ID) LEFT JOIN V_小字 ON D_転用農地.小字ID = V_小字.ID) LEFT JOIN V_地目 ON D_転用農地.登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON D_転用農地.現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON D_転用農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON D_転用農地.借受人ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON D_転用農地.経由農業生産法人ID = [D:個人Info_2].ID) LEFT JOIN D_土地履歴 ON D_転用農地.ID = D_土地履歴.LID ORDER BY D_転用農地.大字ID, D_転用農地.地番;")

            pTBL申請 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_申請;")
            pTBL申請.PrimaryKey = {pTBL申請.Columns("ID")}

            pTBLResult = New DataTable
            CreatepTBLResult(pTBLResult)

            For Each pRow As DataRow In pTBL転用農地.Rows
                Dim pView As DataView = New DataView(pTBL申請, "[農地リスト] Like '*." & pRow.Item("ID") & "*'", "", DataViewRowState.CurrentRows)
                If Not pView.Count > 0 Then
                    Dim pAddRow As DataRow = pTBLResult.NewRow

                    pAddRow.Item("ID") = Val(pRow.Item("ID").ToString)
                    pAddRow.Item("大字ID") = Val(pRow.Item("大字ID").ToString)
                    pAddRow.Item("大字") = pRow.Item("大字名").ToString
                    pAddRow.Item("小字") = pRow.Item("小字名").ToString
                    pAddRow.Item("地番") = pRow.Item("地番").ToString
                    pAddRow.Item("一部現況") = Val(pRow.Item("一部現況").ToString)
                    pAddRow.Item("登記地目名") = pRow.Item("登記地目名").ToString
                    pAddRow.Item("現況地目名") = pRow.Item("現況地目名").ToString
                    pAddRow.Item("登記簿面積") = Val(pRow.Item("登記簿面積").ToString)
                    pAddRow.Item("実面積") = Val(pRow.Item("実面積").ToString)
                    pAddRow.Item("田面積") = Val(pRow.Item("田面積").ToString)
                    pAddRow.Item("畑面積") = Val(pRow.Item("畑面積").ToString)
                    pAddRow.Item("農業振興地域") = IIF(Val(pRow.Item("農業振興地域").ToString) = 0, "農振白地", IIF(Val(pRow.Item("農業振興地域").ToString) = 1, "農用地", "農振外"))

                    pAddRow.Item("履歴内容") = pRow.Item("内容").ToString
                    pAddRow.Item("異動日") = pRow.Item("異動日").ToString
                    pAddRow.Item("更新日") = pRow.Item("更新日").ToString

                    pTBLResult.Rows.Add(pAddRow)
                End If
            Next

            mvarGrid.SetDataView(pTBLResult, "", "")
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub pBtnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pBtnExcel.Click
        mvarGrid.ToExcel()
    End Sub

    Private Sub CreatepTBLResult(ByRef pTBL As DataTable)
        With pTBL
            .Columns.Add("大字ID", GetType(Integer))
            .Columns.Add("大字", GetType(String))
            .Columns.Add("合計 実面積", GetType(String))

            .Columns.Add("ー 合計", GetType(String))
            .Columns.Add("ー 田面積", GetType(String))
            .Columns.Add("ー 畑面積", GetType(String))

            .Columns.Add("使用貸借 合計", GetType(String))
            .Columns.Add("使用貸借 田面積", GetType(String))
            .Columns.Add("使用貸借 畑面積", GetType(String))

            .Columns.Add("賃貸借 合計", GetType(String))
            .Columns.Add("賃貸借 田面積", GetType(String))
            .Columns.Add("賃貸借 畑面積", GetType(String))
        End With
    End Sub
End Class