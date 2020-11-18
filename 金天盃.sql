USE [Golden]
GO
/****** Object:  StoredProcedure [dbo].[GO_SP_Heroes_Second]    Script Date: 2020/11/18 上午 09:24:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:	靜惠
-- Create date: 2020.06.19
-- Description:	群英盃(參照原惠勤的程式) 將原群英盃 Heroes_2019 更名為 Heroes_Second
-- =============================================
ALTER PROCEDURE [dbo].[GO_SP_Heroes_Second]
AS
BEGIN
--MP群英盃

	declare @p_Year int;
	declare @p_qryString nvarchar(200);
	declare @p_qryString_g nvarchar(200);
	declare @p_BGCode int;
	declare @p_UCode int;
	declare @p_qType_TLevel int;

	set @p_Year=2020;
	set @p_qryString='0';
	set @p_qryString_g='0';
	set @p_BGCode=-1;
	set @p_UCode=-1;
	set @p_qType_TLevel=-1;
	-----------------------------------------------------------

	IF Object_id('tempdb..#titletemp') IS NOT NULL
	BEGIN
		DROP TABLE #titletemp
	END

	IF Object_id('tempdb..#Ins_Code') IS NOT NULL
	BEGIN
		DROP TABLE #Ins_Code
	END

	IF Object_id('tempdb..#allianzCode') IS NOT NULL
	BEGIN
		DROP TABLE #allianzCode
	END

	IF Object_id('tempdb..#MP加碼') IS NOT NULL
	BEGIN
		DROP TABLE #MP加碼
	END

	create table #MP加碼 (SumYear int,Period_s int,Period_e int,Period int
		, BGCode int,Branch nvarchar(50),GCode int,gName nvarchar(50),UCode int,UName nvarchar(50),TLevel tinyint,Title nvarchar(50),Title_mp nvarchar(50)
		, TLevel_mp int,pv1 int,pt1 varchar(50),pv2 int,pt2 varchar(50),pv7 int,pt7 varchar(50),pv5 int,pt5_ActHour int,pv6 int,pt6_ActCnt int
		,PlusValue_t int,FYA money,FYC money,PlusExValue_u varchar(50),PlusValue_u int,Type_u varchar(50));
	insert into #MP加碼 exec ZGolden_PBP_SumRecord @p_Year,@p_qryString,@p_qryString_g,@p_BGCode,@p_UCode,@p_qType_TLevel, 2

	--select UCode as Code,UName,PlusValue_t as 'MP加碼'
	--INTO #MP加碼
	--from #temp


	--先清空資料
	DELETE  dbo.Heroes_Second where iYear=@p_Year
			
	create table #titletemp (
		code int, 
		GCode_ int,
		BGCode_ int
	)
				  
			--在當年度報聘者皆需列入群英盃，組織及職級認定皆以06/30(含)前為準
			insert into #titletemp
				select distinct a.Code,b.O1 as GCode_,ISNULL((select FeatBGCode from V_MAN_FeatBG x where x.MAN_Code=a.Code and x.SDate<'2020/02/01'and Void=0 and x.EDate is null),b.O2)--a.Name,dbo.GetGName(GCode),dbo.GetGName(b.O1) 
				from MAN_Data a left join MAN_Chg b on a.Code=b.Man_Code
				where convert(varchar(6),b.[Date],112)<=202012 and a.Date<='2020/12/31'
				and Lay_Off=0-- and a.Remark not like '%原離職日%'
				and b.[Date]=(select MAX([Date]) from MAN_Chg b where a.Code=b.Man_Code and [Date]<='2020/06/30' )
				and OLine=1  --(異動後事業部)
				and b.Man_Code not in (100015058,100009817,100008683)--特殊

				--職級認定：依2020年度第6工作月組織月結後生效之職級為依據計算。
				--此條件須待6工作月(8/10)人事調整後進行更新，故請先依目前最新職級認定。
			insert into #titletemp
				select distinct a.Code,b.O1 as GCode_,ISNULL((select FeatBGCode from V_MAN_FeatBG x where x.MAN_Code=a.Code and x.SDate<'2020/07/01'and Void=0 and x.EDate is null),b.O2)--a.Name,dbo.GetGName(GCode),dbo.GetGName(b.O1)
				from MAN_Data a left join MAN_Chg b on a.Code=b.Man_Code
				where convert(varchar(6),b.[Date],112)<=202006 --and a.Date<='2020/01/31'
				and Lay_Off=0-- and a.Remark not like '%原離職日%'
				--and Rec_No=(select MAX(Rec_No) from MAN_Chg b where a.Code=b.Man_Code and [Date]<='2017/11/30' )
				and b.[Date]=(select MAX([Date]) from MAN_Chg b where a.Code=b.Man_Code and [Date]<='2020/06/30' )
				and OLine=0  --(異動前事業部)
				and a.Code not in (select code from #titletemp)

			--以下是針對2020/06/30 之後才報聘之業務員 start 
			insert into #titletemp
				select distinct a.Code,b.O1 as GCode_,ISNULL((select FeatBGCode from V_MAN_FeatBG x where x.MAN_Code=a.Code and Void=0 and x.EDate is null),b.O2)--a.Name,dbo.GetGName(GCode),dbo.GetGName(b.O1)
				from MAN_Data a left join MAN_Chg b on a.Code=b.Man_Code
				where convert(varchar(6),b.[Date],112)>202006 and a.Date>'2020/06/30'
				and Lay_Off=0 --and a.Remark not like '%原離職日%'
				--and Rec_No=(select MAX(Rec_No) from MAN_Chg b where a.Code=b.Man_Code and [Date]>'2018/11/30' )-- 190911有生養看異動後的事業部&業務中心
				and OLine=1

			insert into #titletemp
				select distinct a.Code,b.O1 as GCode_,b.O2--,ISNULL((select FeatBGCode from V_MAN_FeatBG x where x.MAN_Code=a.Code and Void=0 and x.EDate is null),b.O2)--a.Name,dbo.GetGName(GCode),dbo.GetGName(b.O1)
				from MAN_Data a left join MAN_Chg b on a.Code=b.Man_Code
				where convert(varchar(6),b.[Date],112)>202006 and a.Date>'2020/06/30'
				and Lay_Off=0 --and a.Remark not like '%原離職日%'
				and a.Code not in (select code from #titletemp)
				and OLine=0
				and Rec_No=(select MIN(Rec_No) from MAN_Chg b where a.Code=b.Man_Code and [Date]>'2020/06/30' )
			--以下是針對2020/06/30 之後才報聘之業務員 end 

			select Ins_Code into #Ins_Code  
			from Product where Pro_Name like '%變額%' and Ins_Code not in ('VXD1A','VXD1B','VXD1C','VXD1D','VXLTAA','VXLTAB','VXLTAC','VXLTAD'
			,'VA2','VA3','VAB6','VAFB6','VXLT7A','VXLT7B','VXLT7C','VXLT7D','VLL','VMRA72C','VMRA43C','VMRB43C','VMRC43C')

			--處理安聯商品的附約,一樣要變成3
			select ins_code allianz_procode
			into #allianzCode
			from Product where SupCode=300000050 and main=0 and [close]=0 

			insert Heroes_Second
			select distinct @p_Year,dbo.GetGName(ISNULL(e.BGCode_,x.FeatBGCode))--dbo.GetUPGName(isnull(e.GCode_,0)) as 業務中心
			,Isnull((SELECT [order] FROM [Group_Order] WHERE GCode=ISNULL(e.BGCode_,x.FeatBGCode)), 0) 業務中心編號
			--,Isnull((SELECT [order] FROM [Group_Order] WHERE GCode=ISNULL(x.FeatBGCode,c.BGCode)), 0) 業務中心編號
			,dbo.GetUPGCode(isnull(e.GCode_,0))as BGCode
			,dbo.GetGNAME(isnull(e.GCode_,c.GCode)) as 事業部
			,isnull(e.GCode_,c.GCode) as GCode
			,dbo.GetName(isnull(a.man_code,0),3) as 姓名
			,c.Code as ManCode
			,f.職級
			,c.TCode as Title
			,dbo.GetSupName(a.supcode,1) as 保險公司
			,a.INo as 保單號碼
			,b.Receive_Date AS 受理日
			,a.iDate AS 生效日
			,isnull(d.Ins_Code,InsCode)  as 險種代碼
			--,isnull([dbo].[GetProName](a.Pro_No,1),'') AS 險種名稱
			,(ISNULL(d.Pro_Name, '') + (Case When b.ECGroup > 2 Then '(行動投保)' Else '' End)) AS 險種名稱   --d.pro_name AS 險種名稱
			,a.Pw as 繳別
			,A.YP as 年期
			,dbo.GetCusName(b.PayerCode,1) as 要保人
			,dbo.GetCusName(b.InsuredCode,1) as 被保人
			--,(isnull(a.FYP,0)isnull(CRCRate,1)) 保費
			,(isnull(a.FYP,0)) 保費
			,(isnull(a.FYA,0)) FYC
			,(case WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 AND b.ECGroup > 2 and a.Period between 202006 and 202012 THEN 3  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績行動投保為3 公勝(公字)202005143 及 202005144號
				   WHEN CHARINDEX('VLFBYA', Ins_Code) > 0 AND a.SupCode=300000050 AND b.ECGroup > 2 and a.Period between 202006 and 202012 THEN 0.5
				   WHEN a.INo='QL14034205' THEN 3	--依據需求單ITSTRF200700020 設定計績 300%
				   --Modify by andy_pao 2020/11/18 start
				   --經檢查,上方的條件,與下方的條件重覆,故移除此項規則
				   -- WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 and a.Period between 202006 and 202012 THEN 2  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績為2 公勝(公字)202005143 及 202005144號
				   --Modify by andy_pao 2020/11/18 end
				   WHEN Ins_Code in (select allianz_procode from #allianzCode ) THEN 3 
				   WHEN a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 and (c.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633) OR CHARINDEX('安聯-吉利亨通', dbo.GetProName(c.Pro_No, 1)) > 0) and c.Period between 202006 and 202012) then 0.5 --公勝(公字)202002008號
				   WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5
				   WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5
				   WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5
				   WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' then 1
				   WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' then 1
				   WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' then 1
				   when Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) 
				   WHEN b.Job_Type IN (100, 150) AND b.ECGroup > 2 AND FS=0.1 THEN 0.5
				   WHEN b.Job_Type IN (100, 150) AND b.ECGroup > 2 AND FS=0.5 THEN 1
				   else FS end) FS		
			--, FS
			,((isnull(a.FYA,0))*(case WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 AND b.ECGroup > 2 and a.Period between 202006 and 202012 THEN 3  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績行動投保為3 公勝(公字)202005143 及 202005144號
									  WHEN CHARINDEX('VLFBYA', Ins_Code) > 0 AND a.SupCode=300000050 AND b.ECGroup > 2 and a.Period between 202006 and 202012 THEN 0.5
									  WHEN a.INo='QL14034205' THEN 3	--依據需求單ITSTRF200700020 設定計績 300%
									  --Modify by andy_pao 2020/11/18 start
									  --經檢查,上方的條件,與下方的條件重覆,故移除此項規則
									  -- WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 and a.Period between 202006 and 202012 THEN 2  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績為2 公勝(公字)202005143 及 202005144號
									  --Modify by andy_pao 2020/11/18 end
									  WHEN Ins_Code in (select allianz_procode from #allianzCode ) THEN 3 
									  WHEN a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 and (c.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633) OR CHARINDEX('安聯-吉利亨通', dbo.GetProName(c.Pro_No, 1)) > 0) and c.Period between 202006 and 202012) then 0.5 --公勝(公字)202002008號
									  WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5
									  WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5
									  WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5
									  WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' then 1
									  WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' then 1
									  WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' then 1
									  when Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) 
									  WHEN b.Job_Type IN (100, 150) AND b.ECGroup > 2 AND FS=0.1 THEN 0.5
									  WHEN b.Job_Type IN (100, 150) AND b.ECGroup > 2 AND FS=0.5 THEN 1
									  else FS end))獎勵計績
			--,(isnull(a.FYA,0)*FS) 獎勵計績
			--,((isnull(a.FYA,0))*(case when a.Period>201810 then ISNULL((select distinct (Case when ISNULL(MP, 0)>MP_Max then MP_Max else ISNULL(MP, 0) END)*1.0/100 from [MP_Heroes] B where B.sumYear=2019 and a.Man_Code = B.UCode),0) else 0 end))MP加碼
			,PlusValue_t	--MP加碼
			--,(((isnull(a.FYA,0))*(case when Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) else FS end))--20200116佳鈺說行動加碼用獎勵計績FYA去乘
			,((isnull(a.FYA,0))
			*(case WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 AND b.ECGroup > 2 and a.Period between 202006 and 202012 THEN 3  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績行動投保為3 公勝(公字)202005143 及 202005144號
				   WHEN CHARINDEX('VLFBYA', Ins_Code) > 0 AND a.SupCode=300000050 AND b.ECGroup > 2 and a.Period between 202006 and 202012 THEN 0.5
				   --Modify by andy_pao 2020/11/18 start
				   --經檢查,上方的條件,與下方的條件重覆,故移除此項規則
				   --WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 and a.Period between 202006 and 202012 THEN 2  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績為2 公勝(公字)202005143 及 202005144號
				   --Modify by andy_pao 2020/11/18 end
				   WHEN Ins_Code in (select allianz_procode from #allianzCode ) THEN 3 
				   WHEN a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 and c.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633) and c.Period between 202006 and 202012) then 0.5 --公勝(公字)202002008號
				   WHEN a.SupCode=300000062 AND b.ECGroup >2 THEN 0 -- 排除安達壽不加行動投保加碼
				   when Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) 
				   When a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 AND d.Job_Type IN (100, 150) and c.Period between 202006 and 202012) then 0--行動投保件產專就不再行動加碼
				   else FS end)
			--*FS
			*(
			--Modify by andy_pao 2020/11/18 start
			--這裡的條件,與上方規則,看似相似但又不同,故換成另一個
			case WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 AND b.ECGroup > 2 and a.Period between 202006 and 202012 THEN 3  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績行動投保為3 公勝(公字)202005143 及 202005144號
			--case when CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 and a.Period between 202006 and 202012 then 0  ---公勝(公字)202002008號
			--Modify by andy_pao 2020/11/18 end
				   --下列為原惠勤程式
				   WHEN a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 and c.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633) and c.Period between 202006 and 202012) then 0.5 --公勝(公字)202002008號
				   when a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 and d.SupCode in (300000171,300000048,300000051,300000052,300000055,300000064,300000116,300000168) and c.Period between 202006 and 202012) then 0.5 
				   WHEN a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 and d.SupCode=300000050) then 0 --安聯不加碼
				   WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5	--安達壽行動投保加碼
			       WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5	--安達壽行動投保加碼
			       WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and b.Effe_Date between '2020/07/01' and '2021/01/01' and b.ECGroup > 2 then 1.5	--安達壽行動投保加碼
				   when a.INo in (select distinct d.Ins_No from V_Feat_SR c left join Insurance d on c.INo=b.Ins_No where d.ECGroup>2 and FType in (1,4,6)) then --這段co保經程式，參照業績統計類別的那個公文第3點
						(CASE WHEN CHARINDEX('變額',dbo.GetProName(a.Pro_No,1)) > 0 THEN 0.1
							  WHEN FType = 1 and MYP >= 5 and charindex('定期',dbo.GetProName(a.Pro_No,1)) > 0 and charindex('終身',dbo.GetProName(a.Pro_No,1)) = 0 THEN 0.5 --公勝表示 商品名稱有含"定期"但不含"終身"為關鍵字 20180629 by Tim
							  WHEN FType = 1 and MYP >= 10 and PW <> 'D' THEN 1 
							  WHEN FType = 1 and MYP between 5 and 9 and PW <> 'D' THEN 0.5 
							  WHEN FType = 6 THEN 1 --產險專案(PA)從50%->100%
							  ELSE (case when FType=4 then 0.5 else 0.1 end) END)--20200211雅瑩回答其他的產險類FType4才是10%->50%
			  else 0 end))行動投保加碼 
			--,(isnull(a.FYA,0) * FS)行動投保加碼 
			,NULL
			,a.Period as HefaPeriod
			,NULL
			,NULL
			,'核發' flag --,f.Pro_No
			,NULL
			,NULL
			from V_Feat_SR a inner join Insurance b on a.INo=b.Ins_No and a.SupCode=b.SupCode
			 LEFT OUTER JOIN V_MAN_FeatBG x 
			ON a.MAN_Code = x.MAN_Code and x.Void = 0 --and a.IDate between isnull(x.SDate,'190011') and isnull(x.EDate,'20401231') --and b.Void = 0
			INNER JOIN MAN_Data C ON A.Man_Code=C.Code 
			left join #titletemp e on a.Man_Code=e.code
			left join 群英盃職級 f on a.Man_Code=f.Code
			left JOIN Product D ON a.Pro_No=d.Pro_No and a.SupCode=d.SupCode
			left join #MP加碼 g on a.Man_Code=g.UCode
			where a.CNO=10000 and b.Void=0 and f.年度=@p_Year AND f.stype=2 --and (a.CT=1 or a.CT=20) --and d.[Close]=0
			--and a.Man_Code=100009594
			--and INo='K002949127'
			AND INo Not IN ('1017982151', '1111546886', 'LVAK017667','LVAK019991') --,FU01190319Z094依據需求單 BUL200200080, BUL200200082, ITSTRF200200016 將此保單不列入「2020年總經理盃」獎勵業績
			and a.Period between 202006 and (case when (select MAX(Period) from SR_Close where CT in (1,5)) > 202012 then 202012
			else (select MAX(Period) from SR_Close where CT in (1,5)) end)  --and ino='1026874935'--and a.Man_Code=100009501 --


			insert Heroes_Second
			select distinct @p_Year,dbo.GetGName(ISNULL(c.BGCode_,x.FeatBGCode)) as 業務中心
				,Isnull((SELECT [order] FROM [Group_Order] WHERE GCode=ISNULL(c.BGCode_,x.FeatBGCode)), 0) 業務中心編號
					,dbo.GetUpGCode(isnull(c.GCode_,b.GCode)) as BGCode
					,dbo.GetGNAME(isnull(c.GCode_,b.GCode)) as 事業部
					,isnull(c.GCode_,b.GCode) as GCode
					,dbo.GetName(isnull(a.FeatCode,0),3) as 姓名
					,a.FeatCode as ManCode
					,d.職級
					,b.TCode as title
					,dbo.GetSupName(supcode,1) as 保險公司
					,ins_No as 保單號碼
					,Receive_Date AS 受理日
					,Effe_Date AS 生效日
					,Ins_Code as 險種代碼
					,(ISNULL(Pro_Name, '') + (Case WHEN (SELECT TOP 1 EcGroup FROM Insurance WHERE Ins_No=a.Ins_No) > 2 THEN '(行動投保)' ELSE '' END)) AS 險種名稱   --d.pro_name AS 險種名稱
					,PayType as 繳別
					,YPeriod as 年期
					,dbo.GetCusName(PayerCode,1) as 要保人
					,dbo.GetCusName(InsuredCode,1) as 被保人
					--,(isnull(a.FYP,0)isnull(CRCRate,1)) 保費
					,round(convert(money,isnull(a.FYP,0)),2) 保費
					,round(convert(money,isnull(a.FYA,0)),2) FYC
					,(case WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 THEN 3  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績行動投保為3 公勝(公字)202005143 及 202005144號
						   WHEN CHARINDEX('VLFBYA', Ins_Code) > 0 AND a.SupCode=300000050 AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2  THEN 0.5
						   WHEN a.Ins_No='QL14034205' THEN 3	--依據需求單ITSTRF200700020 設定計績 300%
						   WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 THEN 2  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績為2 公勝(公字)202005143 及 202005144號
						   WHEN Ins_Code in (select allianz_procode from #allianzCode ) THEN 3 
						   WHEN a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 and (t1.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633) OR CHARINDEX('安聯-吉利亨通', t1.Pro_Name) > 0)) then 0.5 --公勝(公字)202002008號
						   WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' and (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 then 1.5
							WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' and (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 then 1.5
							WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' and (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 then 1.5
							WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' then 1
							WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' then 1
							WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' then 1
						   when Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) 
						   WHEN a.Job_Type IN (100, 150) AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 AND (FS=0.5 OR InsType IN (10, 22)) THEN 1
						   WHEN a.Job_Type IN (100, 150) AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 AND FS=0.1 THEN 0.5
						   
						   else FS end)FS
					,(round(convert(money,isnull(a.FYA,0)),2)*
					(case WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 THEN 3  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績行動投保為3 公勝(公字)202005143 及 202005144號
						  WHEN CHARINDEX('VLFBYA', Ins_Code) > 0 AND a.SupCode=300000050 AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2  THEN 0.5
						  WHEN a.Ins_No='QL14034205' THEN 3	--依據需求單ITSTRF200700020 設定計績 300%
						  WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 THEN 2  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績為2 公勝(公字)202005143 及 202005144號
						  WHEN Ins_Code in (select allianz_procode from #allianzCode ) THEN 3 
						  WHEN a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 and t1.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633)) then 0.5 --公勝(公字)202002008號
						  WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' and (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 then 1.5
						  WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' and (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 then 1.5
						  WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' and (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 then 1.5
						  WHEN CHARINDEX('VMRA43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' then 1
						  WHEN CHARINDEX('VMRB43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' then 1
						  WHEN CHARINDEX('VMRC43C1', Ins_Code) > 0 and a.Effe_Date between '2020/07/01' and '2021/01/01' then 1
						  WHEN Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) 
						  WHEN a.Job_Type IN (100, 150) AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 AND (FS=0.5 OR InsType IN (10, 22)) THEN 1
						  WHEN a.Job_Type IN (100, 150) AND (SELECT TOP 1 ECGroup FROM Insurance WHERE Ins_No=a.Ins_No AND SupCode = a.SupCode) > 2 AND FS=0.1 THEN 0.5					  
					else FS end))獎勵計績
					,PlusValue_t	--MP加碼--(select MP加碼 from #MP加碼 z where a.Man_Code=z.Code)MP加碼
					--,((round(convert(money,isnull(FYA,0)),2)*(case when Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) else FS end))----20200116佳鈺說行動加碼用獎勵計績FYA去乘
					,(round(convert(money,isnull(a.FYA,0)),2)
					*(case WHEN CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 THEN 2  --安聯人壽-時來運轉：自2020/06工作月起(不論生效日)獎勵計績為2 公勝(公字)202005143 及 202005144號
						   WHEN Ins_Code in (select allianz_procode from #allianzCode ) THEN 3 
						   WHEN a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 and t1.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633)) then 0.5 --公勝(公字)202002008號
						   WHEN a.SupCode=300000062 AND (SELECT ECGroup FROM Insurance g WHERE g.Ins_No=a.Ins_No AND g.SupCode=a.SupCode) >2 THEN 0 -- 排除安達壽不加行動投保加碼						   when Ins_Code in (select * from #Ins_Code) then ((FS-FS)+1*0.1) 
						   When a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 AND t1.Job_Type IN (100, 150)) then 0
						   else FS end)
				    *(case when CHARINDEX('VXLT7', Ins_Code) > 0 AND a.SupCode=300000050 then 0
						   --下列為惠勤的程式
						   when a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 and t1.Pro_No in (300012426,300012427,300012551,300012869,300012632,300012633)) then 0  --公勝(公字)202002008號
						   when a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 and t2.SupCode in (300000171,300000048,300000051,300000052,300000055,300000064,300000116,300000168)) then 0.5 	--t2.Ins_No='04000668370' Code=12998651-->需求單表單編號ITSTRF200900001 計入行動投保加碼中 2020/09/22註銷取消加碼
						   WHEN a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 and t2.SupCode=300000050) then 0 --安聯不加碼
						   when a.Ins_No in (select distinct t2.Ins_No from V_Feat t1 left join Insurance t2 on t1.Ins_No=t2.Ins_No where t2.ECGroup>2 and FType in (1,4,6)) then --這段co保經程式，參照業績統計類別的那個公文第3點	
								(CASE --WHEN CHARINDEX('變額',dbo.GetProName(a.Pro_No,1)) > 0 THEN 0.1 -- 安達天生贏家, 贏利高手,鑫龍星 雖是行動投保但不加碼
									  WHEN FType = 1 and MYP >= 5 and charindex('定期',dbo.GetProName(a.Pro_No,1)) > 0 and charindex('終身',dbo.GetProName(a.Pro_No,1)) = 0 THEN 0.5 --公勝表示 商品名稱有含"定期"但不含"終身"為關鍵字 20180629 by Tim
									  WHEN FType = 1 and MYP >= 10 and PayType <> 'D' THEN 1 
									  WHEN FType = 1 and MYP between 5 and 9 and PayType <> 'D' THEN 0.5 
									  WHEN FType = 6 THEN 1 --產險專案(PA)從50%->100%
									  ELSE (case when FType=4 then 0.5 else 0.1 end) END)--20200211雅瑩回答其他的產險類FType4才是10%->50%
							else 0 end))行動投保加碼
					,NULL
					,'' as HefaPeriod
					,NULL
					,NULL
					,'受理未核發' as flag
					,NULL
					,NULL
			from V_Feat a 
			 LEFT OUTER JOIN V_MAN_FeatBG x 
					ON a.FeatCode = x.MAN_Code  --and a.Effe_Date between isnull(SDate,'190011') and isnull(EDate,'20401231') --and b.Void = 0
			join MAN_Data b on a.FeatCode=b.Code join #titletemp c on a.FeatCode=c.Code 
			left join 群英盃職級 d on a.Man_Code=d.Code and d.年度=@p_Year AND d.stype=2
			left join #MP加碼 e on a.Man_Code=e.UCode
			where (Effe_Date between '2020/01/01' and '2020/12/31') --(Select CONVERT(date, CONCAT('2019',(select substring(CAST(max(Period) as varchar),5,2) from SR_Close where CT=1)) +'01', 111)) and GetDate()--and a.Man_Code=100018481 --and PayType'D' --and (main1 or YPeriod=10)--and YPeriod=10
			--and  a.Ins_No='EBR0249R411908039999'
			--and a.Man_Code=100009296
			and a.Ins_No not in(

					select  a.INo as 保單號碼
					from V_Feat_SR a inner join Insurance b on a.INo=b.Ins_No--LEFT OUTER JOIN Feat_Sup b on a.SupCode=b.SupCode
					INNER JOIN MAN_Data C ON A.Man_Code=C.Code
											--INNER JOIN Product D ON a.InsCode=d.Ins_Code and a.SupCode=d.SupCode
					where a.CNO=10000 --and (a.CT=1 or a.CT=20)
					and a.Period between 202006 and (case when (select MAX(Period) from SR_Close where CT in (1,5)) > 202012 then 202012
					else (select MAX(Period) from SR_Close where CT in (1,5)) end) --and ino='TAIN17180233'
					--and b.Effe_Date between '201711' and '20171031'   --生效日
					and a.ino in (select DISTINCT ino from V_Feat_SR sa inner join Insurance sb on sa.ino=sb.Ins_No inner join Ins_Content sc on sb.code=sc.MainCode 
							where sa.Period between 202006 and (case when (select MAX(Period) from SR_Close where CT in (1,5)) > 202012 then 202012
					else (select MAX(Period) from SR_Close where CT in (1,5)) end) --and ino='0053487730'
							
			)) 
			AND (SELECT Count(Rec_No) FROM Feat_SR WHERE INo=a.Ins_No AND SupCode=a.SupCode AND Period <= (select MAX(Period) from SR_Close where CT in (1,5))) < 1 --排除在201912已核發件


			-- 清除所有暫存資料
			drop table #Ins_Code
			drop table #titletemp
			drop table #MP加碼
			--202101前離職者不列入計算
			delete Heroes_Second where iYear=@p_Year and ManCode in (select Code FROM MAN_Data where Lay_Off=1 and Lay_Off_Date < '2021/01/01')

			--修改'中山','安平','北高' 三業務中心業務人員所在業務中心
			UPDATE Heroes_Second SET 業務中心=(SELECT ISNULL(FeatBG, Branch) FROM V_Man_Data WHERE Code=ManCode),
				業務中心編號=(SELECT [Order] FROM Group_Order WHERE GName=(SELECT ISNULL(FeatBG, Branch) FROM V_Man_Data WHERE Code=ManCode))
				WHERE iYear=2020 AND ManCode IN (SELECT ManCode FROM Heroes_Second a LEFT JOIN V_Man_Data b ON ManCode=Code WHERE iYear=@p_Year 
				and ISNULL(b.FeatBG, b.Branch)<>業務中心 AND ISNULL(b.FeatBG, b.Branch) IN ('中山','安平','北高'))
			
			--修改中壢, 和記 事業部 負責人林俐利 ,保單號 ='0101783401', 台灣人壽要在2020/10月工作月返還她2020/05被追扣的佣41022元,
			--公勝會在2020/11/30發放此筆佣 (2020/11/30查V_FEAT_SR) ,需將它由群英盃中排除, 因為該筆資料屬 2020/05總經理盃
			--實務上，企業作融通，所以只要群英盃排除　（年度=2020  AND  保單號碼 ='0101783401' AND HEFAPERIOD =202010）
			--將下列程式加到 GO_SP_HEROES_SECOND ,由 Heroes_Second 排除上述保單號.

				--IF GETDATE() > '2020/11/29' 
					----PRINT '超過2020/11/29後才執行'
					--DELETE  FROM Heroes_Second WHERE 年度=2020  AND  保單號碼 ='0101783401' AND HEFAPERIOD =202010


--彙總為一位業務員一筆記錄顯示,方便快速顯示
--=============================================================================================================
	DELETE Heroes_Data where 年度=@p_Year and stype=2

insert Heroes_Data
SELECT Distinct @p_Year, ManCode,業務中心,事業部,GCode,姓名,(SELECT TOP 1 職級 FROM [Heroes_Second] WHERE ManCode=Y.ManCode AND iYear=@p_Year ORDER BY 職級 DESC) 職級
,[核發原始獎勵計績],[核發行動投保加計],[核發獎勵計績(加碼後)],[受理未核發原始獎勵計績],[受理未核發行動投保加計],[核發+受理未核發獎勵計績(加碼後)],[峇里島1/4差額(20萬)],[峇里島全額差額(80萬)]
, [峇里島名額],[預估峇里島名額], 2
FROM
(SELECT ManCode, 業務中心, 事業部,GCode, 姓名,  [核發+受理未核發獎勵計績(加碼後)], [核發獎勵計績(加碼後)],[受理未核發行動投保加計]
,[核發原始獎勵計績],[核發行動投保加計],[受理未核發原始獎勵計績]
, (Case WHEN [核發獎勵計績(加碼後)] >=400000 THEN 0 ELSE 400000 - [核發獎勵計績(加碼後)] END) [峇里島1/4差額(20萬)]
, (Case WHEN [核發獎勵計績(加碼後)] >=800000 THEN 0 ELSE 800000 - [核發獎勵計績(加碼後)] END) [峇里島全額差額(80萬)]
,峇里島名額 
,預估峇里島名額
FROM (SELECT A.ManCode, 業務中心, 事業部,GCode, 姓名, 
ROUND((SELECT ISNULL(Sum(獎勵計績) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('受理未核發'))*1,0) AS '受理未核發原始獎勵計績',
ROUND((SELECT ISNULL(Sum(行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('受理未核發'))*1,0) AS '受理未核發行動投保加計',
(CASE WHEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發','受理未核發')),0)>=400000
		THEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發','受理未核發')),0)+ISNULL((select distinct MP加碼 from [Heroes_Second] Z where a.iYear=z.iYear and Z.ManCode = A.ManCode and z.MP加碼=a.MP加碼),0)
		ELSE ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發','受理未核發')),0) END) AS '核發+受理未核發獎勵計績(加碼後)',
ROUND((SELECT ISNULL(Sum(獎勵計績) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0) AS '核發原始獎勵計績',  
ROUND((SELECT ISNULL(Sum(行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0) AS '核發行動投保加計',  
(CASE WHEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0)>=400000
		THEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0)+ISNULL((select distinct MP加碼 from [Heroes_Second] Z where a.iYear=z.iYear and Z.ManCode = A.ManCode and z.MP加碼=a.MP加碼),0)
		ELSE ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0) END) AS '核發獎勵計績(加碼後)',                
(floor((CASE WHEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0)>=400000
		THEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0)+ISNULL((select distinct MP加碼 from [Heroes_Second] Z where a.iYear=z.iYear and Z.ManCode = A.ManCode and z.MP加碼=a.MP加碼),0)
		ELSE ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發')),0) END)/400000)*0.5) AS '峇里島名額', 
(floor((CASE WHEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發','受理未核發')),0)>=400000
		THEN ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發','受理未核發')),0)+ISNULL((select distinct MP加碼 from [Heroes_Second] Z where a.iYear=z.iYear and Z.ManCode = A.ManCode and z.MP加碼=a.MP加碼),0)
		ELSE ROUND((SELECT ISNULL(Sum(獎勵計績+行動投保加碼) , 0) FROM [Heroes_Second] Z WHERE a.iYear=z.iYear and Z.ManCode = A.ManCode And flag IN('核發','受理未核發')),0) END)/400000)*0.5) AS '預估峇里島名額'
FROM [dbo].[Heroes_Second] A where iYear=@p_Year --and ManCode=100020789
GROUP BY A.ManCode,業務中心,事業部,GCode,姓名,MP加碼,iYear) X --WHERE [核發獎勵計績(加碼後)] = 100000 OR [核發+受理未核發獎勵計績(加碼後)] = 100000 
) Y 
ORDER BY [核發+受理未核發獎勵計績(加碼後)] DESC


--===============================================================================================================
--彙總成總表顯示各區部及各業務中心統計
--===============================================================================================================
			delete from RPT_GManager_01 WHERE sYear=@p_Year AND sType=2

			insert into RPT_GManager_01 
			SELECT @p_Year, 1, '全部', 200000000, '全部' 
			,ISNULL((select count(distinct mancode) from [Heroes_Second] where iYear=@p_Year and 業務中心 not in ('總公司','網路投保中心','直轄','強制終止','其他','事業部體系','電子商務')),0)人數
			,ISNULL((select sum(CEILING(峇里島名額)) from Heroes_Data x where sType=2 AND 年度=@p_Year and 業務中心 not in ('總公司','網路投保中心','直轄','強制終止','其他','事業部體系','電子商務') and 峇里島名額>0),0)峇里島出團人數
			,ISNULL((select sum(CEILING(預估峇里島名額)) from Heroes_Data x where sType=2 AND 年度=@p_Year and 業務中心 not in ('總公司','網路投保中心','直轄','強制終止','其他','事業部體系','電子商務') and 預估峇里島名額>0),0)預估峇里島出團人數, 2
			union 
			SELECT @p_Year, X.[Order1] 業務中心編號, X.區部, 200000000+X.[Order],X.區部 業務中心, sum(人數), sum(峇里島出團人數), sum(預估峇里島出團人數), 2  FROM (
		   SELECT distinct c.Order1 業務中心編號, d.Area 區部, c.GCode
			,c.GName 業務中心, e.[Order1], e.[Order]
			,ISNULL((select count(distinct mancode) from [Heroes_Second] where iYear=@p_Year and c.GName=業務中心),0)人數
			,ISNULL((select sum(CEILING(峇里島名額)) from Heroes_Data where sType=2 AND 年度=@p_Year and c.GName=業務中心 and 峇里島名額>0),0)峇里島出團人數
			,ISNULL((select sum(CEILING(預估峇里島名額)) from Heroes_Data where sType=2 AND 年度=@p_Year and c.GName=業務中心 and 預估峇里島名額>0),0)預估峇里島出團人數
		   FROM ((Group_Order c LEFT join [Group] b ON b.GCode=c.GCode) LEFT JOIN Group_Area d ON b.ACode=d.ACode)
			 LEFT JOIN [Group_Order] e ON e.GName=d.Area 
		   WHERE c.GCode > 100) X
		   GROUP BY X.[Order1], X.區部, X.[Order]

			insert into RPT_GManager_01
			SELECT distinct @p_Year, Order1 業務中心編號, '',　GCode
			, GName [業務中心]
			,ISNULL((select count(distinct mancode) from [Heroes_Second] x where iYear=@p_Year and GName=x.業務中心),0)人數
			,ISNULL((select sum(CEILING(峇里島名額)) from Heroes_Data x where sType=2 AND 年度=@p_Year and GName=x.業務中心 and 峇里島名額>0),0)峇里島出團人數
			,ISNULL((select sum(CEILING(預估峇里島名額)) from Heroes_Data x where sType=2 AND 年度=@p_Year and GName=x.業務中心 and 預估峇里島名額>0),0)預估峇里島出團人數, 2
			FROM Group_Order WHERE GCode > 100
			order by Order1

END
