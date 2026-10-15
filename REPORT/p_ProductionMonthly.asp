<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim  MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
	Dim strMain, nSpanTotal
	Dim txtyear, txtmonth, txttodaydate, txtDivision, txtBlok, txtPlantingDate,txtComp,v_txtComp
    Dim tempDivision, tempBlok, tempPlantingDate 
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	Dim nSubMandays_Amount_hi, nSubMandays_Amount_shi
	Dim nOutput_Mandays_Amount_hi, nOutput_Mandays_Amount_shi
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,SBIRotasi
	Dim rsBaInfo, strBaInfo, varFcbaName,varCompanyCode, varCompanyName
	Dim rsownership	, txtownership, tempownership, sprevownership, txtstatus
	
	        v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	= Request("txtPlantation")
			v_txtGroupRegion  	= Request("txtGroupRegion")
			v_txtRegion  		= Request("txtRegion")
			v_txtcompany  		= Request("txtCompany")
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtPlantingDate	= Request("txtPlantingDate")
			txtownership = Request("txtownership")
			
			txtstatus = Request("txtstatus")
			
			if "" & txtstatus = "" then txtstatus = "TM"

			
			v_txtyear 		= Request("txtyear")
			v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 9
			
			'response.write "Year = " & v_txtyear & " Month = " & v_txtmonth 
			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
			
			Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
			wherecons = " where company is not null "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
	        if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   fcbaname = rsCat("fcName")
				   CompanyName=rsCat("FCCOMPANYNAME")
				   MgrName = rsCat("FCMANAGER")
				   KTUName = rsCat("FCKTU")
				   CropName=rsCat("FCCROP")
				   FcbaType=rsCat("fcType")
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			
			if txtownership <> "--SUMMARY--"  And txtownership <> "" then
				if txtownership <>  "--ALL--"  then
					whereCons = whereCons & " AND ownership = '" & txtownership & "'"
				end if
				selectCons = selectCons &  " , ownership "
				orderCons  = orderCons & ", ownership "
			else
				nSpanTotal = nSpanTotal - 1 	  
			end if   
			
			   if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & ",DIVISION "
			   else
			    nSpanTotal = nSpanTotal - 1 	  
			   end if   

            
			
				if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderCons  = orderCons & ",PLANTINGYEAR "
				else
			    nSpanTotal = nSpanTotal - 1    
				end if
		
			
				if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , FIELDCODE "
				   orderCons  = orderCons & ",FIELDCODE "
				else
			    nSpanTotal = nSpanTotal - 1    
				end if
            
			
	' strMain = " SELECT   " & strFieldCons & " accountingyear,Division,fieldcode, plantingyear,HA,POKOK,POKOK_PER_HA," & _
					' "	Prdton_01,Amount_01,JJG_01 ,Mandays_01," & _
					' "	Prdton_02,Amount_02,JJG_02 ,Mandays_02," & _
					' "	Prdton_03,Amount_03,JJG_03 ,Mandays_03," & _
					' "	Prdton_04,Amount_04,JJG_04 ,Mandays_04," & _
					' "	Prdton_05,Amount_05,JJG_05 ,Mandays_05," & _
					' "	Prdton_06,Amount_06,JJG_06 ,Mandays_06," & _
					' "	Prdton_07,Amount_07,JJG_07 ,Mandays_07," & _
					' "	Prdton_08,Amount_08,JJG_08 ,Mandays_08," & _
					' "	Prdton_09,Amount_09,JJG_09 ,Mandays_09," & _
					' "	Prdton_10,Amount_10,JJG_10 ,Mandays_10," & _
					' "	Prdton_11,Amount_11,JJG_11 ,Mandays_11," & _
					' "	Prdton_12,Amount_12,JJG_12 ,Mandays_12," & _
					' "	NVL(Prdton_01,0)+NVL(Prdton_02,0)+NVL(Prdton_03,0)+NVL(Prdton_04,0)+NVL(Prdton_05,0)+NVL(Prdton_06,0)+NVL(Prdton_07,0)+NVL(Prdton_08,0)+NVL(Prdton_09,0)+NVL(Prdton_10,0)+NVL(Prdton_11,0)+NVL(Prdton_12,0) prod_ton_SBI," & _
					' "	NVL(Amount_01,0)+NVL(Amount_02,0)+NVL(Amount_03,0)+NVL(Amount_04,0)+NVL(Amount_05,0)+NVL(Amount_06,0)+NVL(Amount_07,0)+NVL(Amount_08,0)+NVL(Amount_09,0)+NVL(Amount_10,0)+NVL(Amount_11,0)+NVL(Amount_12,0) Amount_SBI," & _
					' "	NVL(JJG_01,0)+NVL(JJG_02,0)+NVL(JJG_03,0)+NVL(JJG_04,0)+NVL(JJG_05,0)+NVL(JJG_06,0)+NVL(JJG_07,0)+NVL(JJG_08,0)+NVL(JJG_09,0)+NVL(JJG_10,0)+NVL(JJG_11,0)+NVL(JJG_12,0) JJG_SBI," & _
					' "	NVL(Mandays_01,0)+NVL(Mandays_02,0)+NVL(Mandays_03,0)+NVL(Mandays_04,0)+NVL(Mandays_05,0)+NVL(Mandays_06,0)+NVL(Mandays_07,0)+NVL(Mandays_08,0)+NVL(Mandays_09,0)+NVL(Mandays_10,0)+NVL(Mandays_11,0)+NVL(Mandays_12,0) Mandays_SBI," & _
					' "	nvl(budget_transaction_prodqty_api.GetQuantityUntilMonth(fcba,accountingyear,1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0)/1000 bud_ton_sbi," & _
					' "	nvl(BUDGET_TRANSACTION_ESTATE_API.GetBudgetUntilMonth(fcba,accountingyear,1,1,'','6502101','','','',fieldcode,division,to_number(plantingyear),'',TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0) bud_Amount_sbi," & _      
					' "	nvl(budget_transaction_prodqty_api.GetJanjangUntilMonth(fcba,accountingyear,1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0) bud_jjg_sbi," & _      
					' "	nvl(budget_transaction_prodqty_api.GetMandaysUntilMonth(fcba,accountingyear,1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0) bud_Mandays_sbi" & _
					' "	from ( select fcba,accountingyear,Division,fieldcode, plantingyear," & _
					' "		  field_api.get_HA(fieldcode, fcba) HA,field_api.get_POKOK(fieldcode, fcba) POKOK," & _
					' "		  field_api.get_POKOK_PER_HA(fieldcode, fcba) POKOK_PER_HA," & _
					' "		  sum(Prdton_01) Prdton_01,sum(Prdton_02) Prdton_02,sum(Prdton_03) Prdton_03,sum(Prdton_04) Prdton_04,sum(Prdton_05) Prdton_05,sum(Prdton_06) Prdton_06," & _
					' "		  sum(Prdton_07) Prdton_07,sum(Prdton_08) Prdton_08,sum(Prdton_09) Prdton_09,sum(Prdton_10) Prdton_10,sum(Prdton_11) Prdton_11,sum(Prdton_12) Prdton_12," & _
					' "		  sum(Amount_01) Amount_01,sum(Amount_02) Amount_02,sum(Amount_03) Amount_03,sum(Amount_04) Amount_04,sum(Amount_05) Amount_05,sum(Amount_06) Amount_06," & _
					' "		  sum(Amount_07) Amount_07,sum(Amount_08) Amount_08,sum(Amount_09) Amount_09,sum(Amount_10) Amount_10,sum(Amount_11) Amount_11,sum(Amount_12) Amount_12," & _
					' "		  sum(JJG_01) JJG_01,sum(JJG_02) JJG_02,sum(JJG_03) JJG_03,sum(JJG_04) JJG_04,sum(JJG_05) JJG_05,sum(JJG_06) JJG_06," & _
					' "		  sum(JJG_07) JJG_07,sum(JJG_08) JJG_08,sum(JJG_09) JJG_09,sum(JJG_10) JJG_10,sum(JJG_11) JJG_11,sum(JJG_12) JJG_12," & _
					' "		  sum(Mandays_01) Mandays_01,sum(Mandays_02) Mandays_02,sum(Mandays_03) Mandays_03,sum(Mandays_04) Mandays_04,sum(Mandays_05) Mandays_05,sum(Mandays_06) Mandays_06," & _
					' "		  sum(Mandays_07) Mandays_07,sum(Mandays_08) Mandays_08,sum(Mandays_09) Mandays_09,sum(Mandays_10) Mandays_10,sum(Mandays_11) Mandays_11,sum(Mandays_12) Mandays_12" & _
					' "		  from (  select t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear," & _
					' "						 0 Prdton_01,0 Prdton_02,0 Prdton_03,0 Prdton_04,0 Prdton_05,0 Prdton_06," & _
					' "						 0 Prdton_07,0 Prdton_08,0 Prdton_09,0 Prdton_10,0 Prdton_11,0 Prdton_12," & _
					' "						 sum(decode(t.accountingperiod, 1, t.Amount, 0)) Amount_01," & _
					' "						 sum(decode(t.accountingperiod, 2, t.Amount, 0)) Amount_02," & _
					' "						 sum(decode(t.accountingperiod, 3, t.Amount, 0)) Amount_03," & _
					' "						 sum(decode(t.accountingperiod, 4, t.Amount, 0)) Amount_04," & _
					' "						 sum(decode(t.accountingperiod, 5, t.Amount, 0)) Amount_05," & _
					' "						 sum(decode(t.accountingperiod, 6, t.Amount, 0)) Amount_06," & _
					' "						 sum(decode(t.accountingperiod, 7, t.Amount, 0)) Amount_07," & _
					' "						 sum(decode(t.accountingperiod, 8, t.Amount, 0)) Amount_08," & _
					' "						 sum(decode(t.accountingperiod, 9, t.Amount, 0)) Amount_09," & _
					' "						 sum(decode(t.accountingperiod, 10, t.Amount, 0)) Amount_10," & _
					' "						 sum(decode(t.accountingperiod, 11, t.Amount, 0)) Amount_11," & _
					' "						 sum(decode(t.accountingperiod, 12, t.Amount, 0)) Amount_12," & _
					' "						 0 JJG_01,0 JJG_02,0 JJG_03,0 JJG_04,0 JJG_05,0 JJG_06," & _
					' "						 0 JJG_07,0 JJG_08,0 JJG_09,0 JJG_10,0 JJG_11,0 JJG_12," & _
					' "						 0 Mandays_01,0 Mandays_02,0 Mandays_03,0 Mandays_04,0 Mandays_05,0 Mandays_06," & _
					' "						 0 Mandays_07,0 Mandays_08,0 Mandays_09,0 Mandays_10,0 Mandays_11,0 Mandays_12" & _
					' "					from view_dailyharvesting_prod t" & _
					' "				   where to_char(t.fddate, 'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
					' "					 and to_char(t.fddate, 'YYYY') = '" & v_txtyear &"'" & _
					' "				   group by t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear" & _
					' "				  union " & _
					' "				  select t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear," & _
					' "						 sum(decode(t.accountingperiod, 1, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_01," & _
					' "						 sum(decode(t.accountingperiod, 2, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_02," & _
					' "						 sum(decode(t.accountingperiod, 3, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_03," & _
					' "						 sum(decode(t.accountingperiod, 4, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_04," & _
					' "						 sum(decode(t.accountingperiod, 5, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_05," & _
					' "						 sum(decode(t.accountingperiod, 6, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_06," & _
					' "						 sum(decode(t.accountingperiod, 7, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_07," & _
					' "						 sum(decode(t.accountingperiod, 8, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_08," & _
					' "						 sum(decode(t.accountingperiod, 9, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_09," & _
					' "						 sum(decode(t.accountingperiod, 10, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_10," & _
					' "						 sum(decode(t.accountingperiod, 11, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_11," & _
					' "						 sum(decode(t.accountingperiod, 12, nvl(t.totaldailyreceivetonase1,0), 0)) Prdton_12," & _
					' "						 0 Amount_01,0 Amount_02,0 Amount_03,0 Amount_04,0 Amount_05,0 Amount_06," & _
					' "						 0 Amount_07,0 Amount_08,0 Amount_09,0 Amount_10,0 Amount_11,0 Amount_12," & _
					' "						 sum(decode(t.accountingperiod, 1, t.TOTALDAILYRECEIVE, 0)) JJG_01," & _
					' "						 sum(decode(t.accountingperiod, 2, t.TOTALDAILYRECEIVE, 0)) JJG_02," & _
					' "						 sum(decode(t.accountingperiod, 3, t.TOTALDAILYRECEIVE, 0)) JJG_03," & _
					' "						 sum(decode(t.accountingperiod, 4, t.TOTALDAILYRECEIVE, 0)) JJG_04," & _
					' "						 sum(decode(t.accountingperiod, 5, t.TOTALDAILYRECEIVE, 0)) JJG_05," & _
					' "						 sum(decode(t.accountingperiod, 6, t.TOTALDAILYRECEIVE, 0)) JJG_06," & _
					' "						 sum(decode(t.accountingperiod, 7, t.TOTALDAILYRECEIVE, 0)) JJG_07," & _
					' "						 sum(decode(t.accountingperiod, 8, t.TOTALDAILYRECEIVE, 0)) JJG_08," & _
					' "						 sum(decode(t.accountingperiod, 9, t.TOTALDAILYRECEIVE, 0)) JJG_09," & _
					' "						 sum(decode(t.accountingperiod, 10, t.TOTALDAILYRECEIVE, 0)) JJG_10," & _
					' "						 sum(decode(t.accountingperiod, 11, t.TOTALDAILYRECEIVE, 0)) JJG_11," & _
					' "						 sum(decode(t.accountingperiod, 12, t.TOTALDAILYRECEIVE, 0)) JJG_12," & _
					' "						 0 Mandays_01,0 Mandays_02,0 Mandays_03,0 Mandays_04,0 Mandays_05,0 Mandays_06," & _
					' "						 0 Mandays_07,0 Mandays_08,0 Mandays_09,0 Mandays_10,0 Mandays_11,0 Mandays_12" & _
					' "					from view_dailyharvesting_mill t" & _
					' "				   where to_char(t.receptiondate, 'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
					' "					 and to_char(t.receptiondate, 'YYYY') = '" & v_txtyear &"'" & _
					' "				   group by t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear" & _
					' "				  union" & _
					' "				  select t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear," & _
					' "						  0 Prdton_01,0 Prdton_02,0 Prdton_03,0 Prdton_04,0 Prdton_05,0 Prdton_06," & _
					' "						 0 Prdton_07,0 Prdton_08,0 Prdton_09,0 Prdton_10,0 Prdton_11,0 Prdton_12," & _
					' "						 0 Amount_01,0 Amount_02,0 Amount_03,0 Amount_04,0 Amount_05,0 Amount_06," & _
					' "						 0 Amount_07,0 Amount_08,0 Amount_09,0 Amount_10,0 Amount_11,0 Amount_12," & _
					' "						 0 JJG_01,0 JJG_02,0 JJG_03,0 JJG_04,0 JJG_05,0 JJG_06," & _
					' "						 0 JJG_07,0 JJG_08,0 JJG_09,0 JJG_10,0 JJG_11,0 JJG_12," & _
					' "						 sum(decode(t.accountingperiod, 1, t.totalDailyMandays, 0)) Mandays_01," & _
					' "						 sum(decode(t.accountingperiod, 2, t.totalDailyMandays, 0)) Mandays_02," & _
					' "						 sum(decode(t.accountingperiod, 3, t.totalDailyMandays, 0)) Mandays_03," & _
					' "						 sum(decode(t.accountingperiod, 4, t.totalDailyMandays, 0)) Mandays_04," & _
					' "						 sum(decode(t.accountingperiod, 5, t.totalDailyMandays, 0)) Mandays_05," & _
					' "						 sum(decode(t.accountingperiod, 6, t.totalDailyMandays, 0)) Mandays_06," & _
					' "						 sum(decode(t.accountingperiod, 7, t.totalDailyMandays, 0)) Mandays_07," & _
					' "						 sum(decode(t.accountingperiod, 8, t.totalDailyMandays, 0)) Mandays_08," & _
					' "						 sum(decode(t.accountingperiod, 9, t.totalDailyMandays, 0)) Mandays_09," & _
					' "						 sum(decode(t.accountingperiod, 10, t.totalDailyMandays, 0)) Mandays_10," & _
					' "						 sum(decode(t.accountingperiod, 11, t.totalDailyMandays, 0)) Mandays_11," & _
					' "						 sum(decode(t.accountingperiod, 12, t.totalDailyMandays, 0)) Mandays_12" & _
					' "					from view_dailyharvesting_mandays t" & _
					' "				   where to_char(t.fddate, 'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
					' "					 and to_char(t.fddate, 'YYYY') = '" & v_txtyear &"'" & _
					' "				   group by t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear" & _
					' "			  )" & _
					' "		   group by fcba, accountingyear, Division, fieldcode, plantingyear" & _
					' "		) "
					
	    ' Remarked By Ukie.. 11 November 2015
		' strMain = 	"  select * from PRODUCTION_MONTHLY_REPORT_TAB t " & _
					' "  where t.year = '" & v_txtyear &"' and  t.month='" & v_txtmonth &"' and USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "',t.fcba) ='TRUE' "
	Dim wherestatus				
	if "" & txtstatus <> "" and "" & txtstatus <> "--ALL--" then
		wherestatus = "       and field_api.GetValid_status(t.fieldcode,t.fcba,TO_DATE(t.month||'/1/'||t.year, 'mm/dd/yyyy') ) = '" & txtstatus & "'"
	end if
					
	strMain = " select year" & _
              "        , month" & _
              "        , groupplantation" & _
              "        , plantation" & _
              "        , groupregion" & _
              "        , region" & _
              "        , company" & _
              "        , fcba" & _
              "        , accountingyear" & _
              "        , ownership" & _
              "        , division" & _
              "        , fieldcode" & _
              "        , sum(prdton_01) prdton_01" & _
              "        , sum(amount_01) amount_01" & _
              "        , sum(jjg_01) jjg_01" & _
              "        , sum(mandays_01) mandays_01" & _
              "        , sum(prdton_02) prdton_02" & _
              "        , sum(amount_02) amount_02" & _
              "        , sum(jjg_02) jjg_02" & _
              "        , sum(mandays_02) mandays_02" & _
              "        , sum(prdton_03) prdton_03" & _
              "        , sum(amount_03) amount_03" & _
              "        , sum(jjg_03) jjg_03" & _
              "        , sum(mandays_03) mandays_03" & _
              "        , sum(prdton_04) prdton_04" & _
              "        , sum(amount_04) amount_04" & _
              "        , sum(jjg_04) jjg_04" & _
              "        , sum(mandays_04) mandays_04" & _
              "        , sum(prdton_05) prdton_05" & _
              "        , sum(amount_05) amount_05" & _
              "        , sum(jjg_05) jjg_05" & _
              "        , sum(mandays_05) mandays_05" & _
              "        , sum(prdton_06) prdton_06" & _
              "        , sum(amount_06) amount_06" & _
              "        , sum(jjg_06) jjg_06" & _
              "        , sum(mandays_06) mandays_06" & _
              "        , sum(prdton_07) prdton_07" & _
              "        , sum(amount_07) amount_07" & _
              "        , sum(jjg_07) jjg_07" & _
              "        , sum(bjr_07) bjr_07" & _
              "        , sum(mandays_07) mandays_07" & _
              "        , sum(prdton_08) prdton_08" & _
              "        , sum(amount_08) amount_08" & _
              "        , sum(jjg_08) jjg_08" & _
              "        , sum(mandays_08) mandays_08" & _
              "        , sum(prdton_09) prdton_09" & _
              "        , sum(amount_09) amount_09" & _
              "        , sum(jjg_09) jjg_09" & _
              "        , sum(mandays_09) mandays_09" & _
              "        , sum(prdton_10) prdton_10" & _
              "        , sum(amount_10) amount_10" & _
              "        , sum(jjg_10) jjg_10" & _
              "        , sum(mandays_10) mandays_10" & _
              "        , sum(prdton_11) prdton_11" & _
              "        , sum(amount_11) amount_11" & _
              "        , sum(jjg_11) jjg_11" & _
              "        , sum(mandays_11) mandays_11" & _
              "        , sum(prdton_12) prdton_12" & _
              "        , sum(amount_12) amount_12" & _
              "        , sum(jjg_12) jjg_12" & _
              "        , sum(bjr_12) bjr_12" & _
              "        , sum(mandays_12) mandays_12" & _
              "        , sum(prod_ton_sbi) prod_ton_sbi" & _
              "        , sum(amount_sbi) amount_sbi" & _
              "        , sum(jjg_sbi) jjg_sbi" & _
              "        , sum(mandays_sbi) mandays_sbi" & _
              "        , sum(bud_ton_sbi) bud_ton_sbi" & _
              "        , sum(bud_amount_sbi) bud_amount_sbi" & _
              "        , sum(bud_jjg_sbi) bud_jjg_sbi" & _
              "        , sum(bud_mandays_sbi) bud_mandays_sbi" & _
              "        , sum(bud_ton_bi) bud_ton_bi" & _
              "        , sum(bud_ton_annual) bud_ton_annual" & _
              "        , sum(rotasi_01) rotasi_01" & _
              "        , sum(rotasi_02) rotasi_02" & _
              "        , sum(rotasi_03) rotasi_03" & _
              "        , sum(rotasi_04) rotasi_04" & _
              "        , sum(rotasi_05) rotasi_05" & _
              "        , sum(rotasi_06) rotasi_06" & _
              "        , sum(rotasi_07) rotasi_07" & _
              "        , sum(rotasi_08) rotasi_08" & _
              "        , sum(rotasi_09) rotasi_09" & _
              "        , sum(rotasi_10) rotasi_10" & _
              "        , sum(rotasi_11) rotasi_11" & _
              "        , sum(rotasi_12) rotasi_12" & _
              "        , sum(rotasi_sbi) rotasi_sbi" & _
              " from production_monthly_report_tab t" & _
              " where t.year = '" & v_txtyear & "'" & _
			  "       and t.month = '" & v_txtmonth & "'" & _
			  "" & wherestatus & "" & _
			  "       and USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', t.fcba) = 'TRUE'"
			  
	strMain = strMain & " group by year" & _
						"          , month" & _
						"          , groupplantation" & _
						"          , plantation" & _
						"          , groupregion" & _
						"          , region" & _
						"          , company" & _
						"          , fcba" & _
						"          , accountingyear" & _
						"          , ownership" & _
						"          , division" & _
						"          , fieldcode"

	strMain = " select aa.*" & _
	          "        , production_monthly_report_api.get_tm_plantingyear(fcba, fieldcode, '" & v_txtyear & "'||LPAD('" & v_txtmonth & "', 2, 0)) plantingyear" & _
			  "        , production_monthly_report_api.get_tm_ha(fcba, fieldcode, '" & v_txtyear & "'||LPAD('" & v_txtmonth & "', 2, 0)) ha" & _
			  "        , production_monthly_report_api.get_tm_pokok(fcba, fieldcode, '" & v_txtyear & "'||LPAD('" & v_txtmonth & "', 2, 0)) pokok" & _
			  "        , decode(production_monthly_report_api.get_tm_ha(fcba, fieldcode, '" & v_txtyear & "'||LPAD('" & v_txtmonth & "', 2, 0))" & _
			  "                 , 0, 0" & _
			  "                 , production_monthly_report_api.get_tm_pokok(fcba, fieldcode, '" & v_txtyear & "'||LPAD('" & v_txtmonth & "', 2, 0))/production_monthly_report_api.get_tm_ha(fcba, fieldcode, '" & v_txtyear & "'||LPAD('" & v_txtmonth & "', 2, 0))" & _
			  "          ) pokok_per_ha" & _
              " from (" & strMain & ") aa"
					
	
								 
	 strEPS2  =  " ,SUM(HA) HA,SUM(POKOK) POKOK,decode(SUM(HA),0,0,SUM(POKOK)/SUM(HA) ) POKOK_PER_HA" & _
	             " ,SUM(Prdton_01) Prdton_01, decode(SUM(HA),0,0,SUM(Prdton_01)/SUM(HA)) Prdton_HA_01,SUM(Amount_01) Amount_01,decode(SUM(prdton_01),0,0,SUM(Amount_01)/(SUM(prdton_01)*1000)) Amount_KG_01,SUM(JJG_01) JJG_01,decode(sum(JJG_01),0,0,(sum(Prdton_01)*1000)/sum(JJG_01)) BJR_01,decode(SUM(POKOK),0,0,SUM(JJG_01)/SUM(POKOK)) JJG_PKK_01,SUM(Mandays_01) Mandays_01,decode(sum(Mandays_01),0,0,(sum(Prdton_01)*1000)/sum(Mandays_01)) KG_HK_01,nvl(decode(SUM(HA),0,0,SUM((ROTASI_01*HA))/SUM(HA)) ,0) ROTASI_01" & _
				 " ,SUM(Prdton_02) Prdton_02, decode(SUM(HA),0,0,SUM(Prdton_02)/SUM(HA)) Prdton_HA_02,SUM(Amount_02) Amount_02,decode(SUM(prdton_02),0,0,SUM(Amount_02)/(SUM(prdton_02)*1000)) Amount_KG_02,SUM(JJG_02) JJG_02,decode(sum(JJG_02),0,0,(sum(Prdton_02)*1000)/sum(JJG_02)) BJR_02,decode(SUM(POKOK),0,0,SUM(JJG_02)/SUM(POKOK)) JJG_PKK_02,SUM(Mandays_02) Mandays_02,decode(sum(Mandays_02),0,0,(sum(Prdton_02)*1000)/sum(Mandays_02)) KG_HK_02,nvl(decode(SUM(HA),0,0,SUM((ROTASI_02*HA))/SUM(HA)) ,0)  ROTASI_02" & _
			     " ,SUM(Prdton_03) Prdton_03, decode(SUM(HA),0,0,SUM(Prdton_03)/SUM(HA)) Prdton_HA_03,SUM(Amount_03) Amount_03,decode(SUM(prdton_03),0,0,SUM(Amount_03)/(SUM(prdton_03)*1000)) Amount_KG_03,SUM(JJG_03) JJG_03,decode(sum(JJG_03),0,0,(sum(Prdton_03)*1000)/sum(JJG_03)) BJR_03,decode(SUM(POKOK),0,0,SUM(JJG_03)/SUM(POKOK)) JJG_PKK_03,SUM(Mandays_03) Mandays_03,decode(sum(Mandays_03),0,0,(sum(Prdton_03)*1000)/sum(Mandays_03)) KG_HK_03,nvl(decode(SUM(HA),0,0,SUM((ROTASI_03*HA))/SUM(HA)) ,0)  ROTASI_03" & _
                 " ,SUM(Prdton_04) Prdton_04, decode(SUM(HA),0,0,SUM(Prdton_04)/SUM(HA)) Prdton_HA_04,SUM(Amount_04) Amount_04,decode(SUM(prdton_04),0,0,SUM(Amount_04)/(SUM(prdton_04)*1000)) Amount_KG_04,SUM(JJG_04) JJG_04,decode(sum(JJG_04),0,0,(sum(Prdton_04)*1000)/sum(JJG_04)) BJR_04,decode(SUM(POKOK),0,0,SUM(JJG_04)/SUM(POKOK)) JJG_PKK_04,SUM(Mandays_04) Mandays_04,decode(sum(Mandays_04),0,0,(sum(Prdton_04)*1000)/sum(Mandays_04)) KG_HK_04,nvl(decode(SUM(HA),0,0,SUM((ROTASI_04*HA))/SUM(HA)) ,0)  ROTASI_04" & _
                 " ,SUM(Prdton_05) Prdton_05, decode(SUM(HA),0,0,SUM(Prdton_05)/SUM(HA)) Prdton_HA_05,SUM(Amount_05) Amount_05,decode(SUM(prdton_05),0,0,SUM(Amount_05)/(SUM(prdton_05)*1000)) Amount_KG_05,SUM(JJG_05) JJG_05,decode(sum(JJG_05),0,0,(sum(Prdton_05)*1000)/sum(JJG_05)) BJR_05,decode(SUM(POKOK),0,0,SUM(JJG_05)/SUM(POKOK)) JJG_PKK_05,SUM(Mandays_05) Mandays_05,decode(sum(Mandays_05),0,0,(sum(Prdton_05)*1000)/sum(Mandays_05)) KG_HK_05,nvl(decode(SUM(HA),0,0,SUM((ROTASI_05*HA))/SUM(HA)) ,0)  ROTASI_05" & _
                 " ,SUM(Prdton_06) Prdton_06, decode(SUM(HA),0,0,SUM(Prdton_06)/SUM(HA)) Prdton_HA_06,SUM(Amount_06) Amount_06,decode(SUM(prdton_06),0,0,SUM(Amount_06)/(SUM(prdton_06)*1000)) Amount_KG_06,SUM(JJG_06) JJG_06,decode(sum(JJG_06),0,0,(sum(Prdton_06)*1000)/sum(JJG_06)) BJR_06,decode(SUM(POKOK),0,0,SUM(JJG_06)/SUM(POKOK))  JJG_PKK_06,SUM(Mandays_06) Mandays_06,decode(sum(Mandays_06),0,0,(sum(Prdton_06)*1000)/sum(Mandays_06)) KG_HK_06,nvl(decode(SUM(HA),0,0,SUM((ROTASI_06*HA))/SUM(HA)) ,0)  ROTASI_06" & _
                 " ,SUM(Prdton_07) Prdton_07, decode(SUM(HA),0,0,SUM(Prdton_07)/SUM(HA)) Prdton_HA_07,SUM(Amount_07) Amount_07,decode(SUM(prdton_07),0,0,SUM(Amount_07)/(SUM(prdton_07)*1000)) Amount_KG_07,SUM(JJG_07) JJG_07,decode(sum(JJG_07),0,0,(sum(Prdton_07)*1000)/sum(JJG_07)) BJR_07,decode(SUM(POKOK),0,0,SUM(JJG_07)/SUM(POKOK))  JJG_PKK_07,SUM(Mandays_07) Mandays_07,decode(sum(Mandays_07),0,0,(sum(Prdton_07)*1000)/sum(Mandays_07)) KG_HK_07,nvl(decode(SUM(HA),0,0,SUM((ROTASI_07*HA))/SUM(HA)) ,0)  ROTASI_07" & _
                 " ,SUM(Prdton_08) Prdton_08, decode(SUM(HA),0,0,SUM(Prdton_08)/SUM(HA)) Prdton_HA_08,SUM(Amount_08) Amount_08,decode(SUM(prdton_08),0,0,SUM(Amount_08)/(SUM(prdton_08)*1000)) Amount_KG_08,SUM(JJG_08) JJG_08,decode(sum(JJG_08),0,0,(sum(Prdton_08)*1000)/sum(JJG_08)) BJR_08,decode(SUM(POKOK),0,0,SUM(JJG_08)/SUM(POKOK))  JJG_PKK_08,SUM(Mandays_08) Mandays_08,decode(sum(Mandays_08),0,0,(sum(Prdton_08)*1000)/sum(Mandays_08)) KG_HK_08,nvl(decode(SUM(HA),0,0,SUM((ROTASI_08*HA))/SUM(HA)) ,0)  ROTASI_08" & _
				 " ,SUM(Prdton_09) Prdton_09, decode(SUM(HA),0,0,SUM(Prdton_09)/SUM(HA)) Prdton_HA_09,SUM(Amount_09) Amount_09,decode(SUM(prdton_09),0,0,SUM(Amount_09)/(SUM(prdton_09)*1000)) Amount_KG_09,SUM(JJG_09) JJG_09,decode(sum(JJG_09),0,0,(sum(Prdton_09)*1000)/sum(JJG_09)) BJR_09,decode(SUM(POKOK),0,0,SUM(JJG_09)/SUM(POKOK))  JJG_PKK_09,SUM(Mandays_09) Mandays_09,decode(sum(Mandays_09),0,0,(sum(Prdton_09)*1000)/sum(Mandays_09)) KG_HK_09,nvl(decode(SUM(HA),0,0,SUM((ROTASI_09*HA))/SUM(HA)) ,0)  ROTASI_09" & _
				 " ,SUM(Prdton_10) Prdton_10, decode(SUM(HA),0,0,SUM(Prdton_10)/SUM(HA)) Prdton_HA_10,SUM(Amount_10) Amount_10,decode(SUM(prdton_10),0,0,SUM(Amount_10)/(SUM(prdton_10)*1000)) Amount_KG_10,SUM(JJG_10) JJG_10,decode(sum(JJG_10),0,0,(sum(Prdton_10)*1000)/sum(JJG_10)) BJR_10,decode(SUM(POKOK),0,0,SUM(JJG_10)/SUM(POKOK))  JJG_PKK_10,SUM(Mandays_10) Mandays_10,decode(sum(Mandays_10),0,0,(sum(Prdton_10)*1000)/sum(Mandays_10)) KG_HK_10,nvl(decode(SUM(HA),0,0,SUM((ROTASI_10*HA))/SUM(HA)) ,0) ROTASI_10" & _
				 " ,SUM(Prdton_11) Prdton_11, decode(SUM(HA),0,0,SUM(Prdton_11)/SUM(HA)) Prdton_HA_11,SUM(Amount_11) Amount_11,decode(SUM(prdton_11),0,0,SUM(Amount_11)/(SUM(prdton_11)*1000)) Amount_KG_11,SUM(JJG_11) JJG_11,decode(sum(JJG_11),0,0,(sum(Prdton_11)*1000)/sum(JJG_11)) BJR_11,decode(SUM(POKOK),0,0,SUM(JJG_11)/SUM(POKOK))  JJG_PKK_11,SUM(Mandays_11) Mandays_11,decode(sum(Mandays_11),0,0,(sum(Prdton_11)*1000)/sum(Mandays_11)) KG_HK_11,nvl(decode(SUM(HA),0,0,SUM((ROTASI_11*HA))/SUM(HA)) ,0)  ROTASI_11" & _
				 " ,SUM(Prdton_12) Prdton_12, decode(SUM(HA),0,0,SUM(Prdton_12)/SUM(HA)) Prdton_HA_12,SUM(Amount_12) Amount_12,decode(SUM(prdton_12),0,0,SUM(Amount_12)/(SUM(prdton_12)*1000)) Amount_KG_12,SUM(JJG_12) JJG_12,decode(sum(JJG_12),0,0,(sum(Prdton_12)*1000)/sum(JJG_12)) BJR_12,decode(SUM(POKOK),0,0,SUM(JJG_12)/SUM(POKOK))  JJG_PKK_12,SUM(Mandays_12) Mandays_12,decode(sum(Mandays_12),0,0,(sum(Prdton_12)*1000)/sum(Mandays_12)) KG_HK_12,nvl(decode(SUM(HA),0,0,SUM((ROTASI_12*HA))/SUM(HA)) ,0) ROTASI_12" & _
                 " ,SUM(prod_ton_SBI) prod_ton_SBI,decode(SUM(HA),0,0,SUM(prod_ton_SBI)/SUM(HA)) prod_ton_HA_SBI,SUM(Amount_SBI) Amount_SBI,decode(SUM(prod_ton_SBI),0,0,SUM(Amount_SBI)/(SUM(prod_ton_SBI)*1000)) Amount_KG_SBI " & _
				 " ,SUM(JJG_SBI) JJG_SBI,decode(sum(JJG_SBI),0,0,(sum(Prod_ton_SBI)*1000)/sum(JJG_SBI))  BJR_SBI,DECODE(SUM(POKOK),0,0,SUM(JJG_SBI)/ SUM(POKOK)) JJG_PKK_SBI" & _
                 " ,SUM(Mandays_SBI) Mandays_SBI,decode(sum(Mandays_SBI),0,0,(sum(Prod_ton_SBI)*1000)/sum(Mandays_SBI)) KG_HK_SBI " & _ 
				 " ,SUM(bud_ton_sbi) bud_ton_sbi,decode(SUM(HA),0,0,SUM(bud_ton_sbi)/SUM(HA)) bud_prod_ton_HA_SBI,SUM(bud_Amount_sbi) bud_Amount_sbi,decode(SUM(bud_ton_sbi),0,0,SUM(bud_Amount_sbi)/(SUM(bud_ton_sbi) *1000)) bud_Amount_KG_SBI " & _ 
				 " ,SUM(bud_JJG_sbi) bud_jjg_sbi,decode(SUM(bud_JJG_sbi),0,0,(SUM(bud_ton_sbi)*1000)/SUM(bud_JJG_sbi)) BUD_BJR_SHI,decode(SUM(POKOK),0,0,SUM(bud_JJG_sbi)/SUM(POKOK)) BUD_JJG_PKK_SBI " & _ 
				 " ,SUM(bud_mandays_sbi) bud_mandays_sbi,decode(SUM(bud_mandays_sbi),0,0,(SUM(bud_ton_sbi)*1000)/ SUM(bud_mandays_sbi)) bud_KG_HK_SBI ,nvl(SUM(ROTASI_SBI),0) ROTASI_SBI"
								 

	strEPS =  "select GroupPlantation " & selectCons & strEPS2 &"  From (" & strMain & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  GroupPlantation" & selectCons & "  Order by GroupPlantation" & orderCons  	



'response.write strEPS
				
Set rsEPS = Server.CreateObject("ADODB.Recordset")

rsEPS.Open strEPS , connstr, adOpenStatic 



strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Analisa Pencapaian Produksi</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link href="/css/font/font-awesome.css" rel="stylesheet">
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<style id="KG_my_style"></style>
	<style id="TON_HA_my_style"></style>
	<style id="RP_my_style"></style>
	<style id="RP_KG_my_style"></style>
	<style id="JJG_my_style"></style>
	<style id="BJR_my_style"></style>
	<style id="JJG_PKK_my_style"></style>
	<style id="HK_my_style"></style>
	<style id="KG_HK_my_style"></style>
	<style id="ROTASI_my_style"></style>
<script type="text/javascript">

    function toggle(id){
		var id_button = id + "_button";
		var button = document.getElementById(id_button);
		var jan = document.getElementById("jan");
		var feb = document.getElementById("feb");
        var mar = document.getElementById("mar");
        var apr = document.getElementById("apr");
        var may = document.getElementById("may");
		var jun = document.getElementById("jun");
		var jul = document.getElementById("jul");
		var aug = document.getElementById("aug");
		var sep = document.getElementById("sep");
		var oct = document.getElementById("oct");
		var nov = document.getElementById("nov");
		var des = document.getElementById("des");
		var sbi = document.getElementById("sbi");
		var bgt = document.getElementById("bgt");
		//for (c = 1; c < 15; c++) {
			//var id_ = id + c;
			//alert(id_);
			//var e = document.getElementById(id_);
			//if(e.style.display == 'table-cell'){
			//	e.style.display = 'none';
			//}else if(e.style.display == 'none'){
			//	e.style.display = 'table-cell';
			//}
		//}
		
			if (button.className == "uncheck"){
				button.className = "check";
				button.style.background = "#5bb75b";
				jan.colSpan = jan.colSpan + 1;
				feb.colSpan = feb.colSpan + 1;
				mar.colSpan = mar.colSpan + 1;
				apr.colSpan = apr.colSpan + 1;
				may.colSpan = may.colSpan + 1;
				jun.colSpan = jun.colSpan + 1;
				jul.colSpan = jul.colSpan + 1;
				aug.colSpan = aug.colSpan + 1;
				sep.colSpan = sep.colSpan + 1;
				oct.colSpan = oct.colSpan + 1;
				nov.colSpan = nov.colSpan + 1;
				des.colSpan = des.colSpan + 1;
				sbi.colSpan = sbi.colSpan + 1;
				if (id === "ROTASI"){
						bgt.colSpan = bgt.colSpan;
					}else{
						bgt.colSpan = bgt.colSpan + 1;
					}
				document.getElementById(id + "_my_style").innerHTML = "." + id + "{display:table-cell;}";
			}else if (button.className == "check"){
				if (bgt.colSpan === 1 ){
					alert("Tidak Bisa di-hide semua Kolom !");
				}else{
					button.className = "uncheck";
					button.style.background = "#ccc";
					jan.colSpan = jan.colSpan - 1;
					feb.colSpan = feb.colSpan - 1;
					mar.colSpan = mar.colSpan - 1;
					apr.colSpan = apr.colSpan - 1;
					may.colSpan = may.colSpan - 1;
					jun.colSpan = jun.colSpan - 1;
					jul.colSpan = jul.colSpan - 1;
					aug.colSpan = aug.colSpan - 1;
					sep.colSpan = sep.colSpan - 1;
					oct.colSpan = oct.colSpan - 1;
					nov.colSpan = nov.colSpan - 1;
					des.colSpan = des.colSpan - 1;
					sbi.colSpan = sbi.colSpan - 1;
					if (id === "ROTASI"){
						bgt.colSpan = bgt.colSpan;
					}else{
						bgt.colSpan = bgt.colSpan - 1;
					}
					document.getElementById(id + "_my_style").innerHTML = "." + id + "{display:none;}";
				}	
			}
			xRemoveEventListener(window, 'load','', false);
	}				

</script>
<script language="JavaScript">
xAddEventListener(window, 'load',function() {new xTableHeaderFixed('xthf-site', window);}, false);

//xRemoveEventListener(ele, sEventType, fnEventListener[, bCapture])
function xRemoveEventListener(e,eT,eL,cap)
{
  if(!(e=xGetElementById(e)))return;
  eT=eT.toLowerCase();
  if(e.removeEventListener)e.removeEventListener(eT,eL,cap||false);
  else if(e.detachEvent)e.detachEvent('on'+eT,eL);
  else e['on'+eT]=null;
}


</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>

<!--<script type="text/javascript" src="http://jqueryjs.googlecode.com/files/jquery-1.3.1.min.js" > </script> -->
<script type="text/javascript">

    function PrintElem(elem)
    {
        Popup($(elem).html());
    }

    function Popup(data) 
    {
        var mywindow = window.open('', 'my div', 'height=400,width=600');
        mywindow.document.write('<html><head><title>my div</title>');
        mywindow.document.write('<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />');
        mywindow.document.write('</head><body >');
        mywindow.document.write(data);
        mywindow.document.write('</body></html>');

        mywindow.print();
        mywindow.close();

        return true;
    }

</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

	
	
</head>

<body>
<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	</div>
<div class="top_select" style=" position:absolute; top:0; left:0px; width:100%;">
<form name="frmEntry" method="post" action="p_ProductionMonthly.asp">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	

<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>BULAN</th>
		<th>TAHUN</th>
		<th>STATUS</th>
		<th>OWNERSHIP</th>
		<th>AFDELING</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
		<th>&nbsp;</th>
	</tr>
	<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %-->
		<% 
					MM=Request("txtMonth")
					yyyy=Request("txtYear")

					mon = Month(Session("v_Costbook"))
					If len(MM)=0 then MM=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = 2015
					MaxYr = Yr  + 2
					If len(yyyy)= 0 then yyyy = Yr
				%>
		
		   <td>
				<font size="2" color="#0000FF" ></font>
					<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
						<%
							Next
						%>
					</select>
		  </td>	
		   <td>
				<font size="2"  color="#0000FF" ></font>
				
				<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
				</select>
			</td>
			
		<td align="center">
			<select name="txtstatus">	
				<option value="--ALL--" <%if txtstatus = "--ALL--" then%> Selected <%end if%>>--ALL--</option>
				<option value="TM" <%if txtstatus = "TM" then%> Selected <%end if%>>TM</option>
				<option value="TBM" <%if txtstatus = "TBM" then%> Selected <%end if%>>TBM</option>
			</select>		
		</td>
			
				
		<td align="center">
			<select name="txtownership">	
				<option value="--SUMMARY--" <%if txtownership = "--SUMMARY--" then%> Selected <%end if%>>--SUMMARY--</option>
				<option value="--ALL--" <%if txtownership = "--ALL--" then%> Selected <%end if%>>--ALL--</option>
				<%
					strsql =  " SELECT DISTINCT fccode FROM parameter WHERE parcode = 'F0004'"
					Set rsownership = Server.CreateObject("ADODB.Recordset")
					rsownership.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsownership.eof then
						DO while not rsownership.eof
				%>			<option value="<%=rsownership("fccode")%>" <%if txtownership = rsownership("fccode") then%>Selected<%end if%>><%=rsownership("fccode")%></option>
				<%			rsownership.movenext
						loop
					End if
					rsownership.Close
					set rsownership = nothing
				%>
			</select>		
		</td>
		
		<td align="center">
			<select name="txtDivision">	
				<option value="--SUMMARY--">--SUMMARY--</option>
				<%
					Dim rsDivision
					Set rsDivision = Server.CreateObject("ADODB.Recordset")
					rsDivision.CursorLocation = adUseClient
					

			
					if len(trim(v_txtDivision)) > 0 then 
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
					end if
					
					
					rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsDivision.eof then
						DO while not rsDivision.eof
						if v_txtDivision = rsDivision("Division") then
				%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
						<%else%>
							<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
						<%end if%>
				<%			rsDivision.movenext
						loop
					End if
					rsDivision.Close
					set rsDivision = nothing
				%>
			</select>		
		</td>


		<td align="center">
			<select name="txtPlantingDate">	
					<option value="--SUMMARY--">--SUMMARY--</option>
					<%
						Dim rsPlantingDate
						Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
						rsPlantingDate.CursorLocation = adUseClient
				
						if len(trim(v_txtPlantingDate)) > 0 then 
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
						else
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
						end if
						
						
						rsPlantingDate.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
						if not rsPlantingDate.eof then
							DO while not rsPlantingDate.eof
							if v_txtPlantingDate = rsPlantingDate("PlantingDate") then
					%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
							<%else%>
								<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
							<%end if%>
					<%			rsPlantingDate.movenext
							loop
						End if
						rsPlantingDate.Close
						set rsPlantingDate = nothing
					%>
			</select>		
		</td>
		<td align="center">
			<select name="txtBlok">	
				<option value="--SUMMARY--">--SUMMARY--</option>
				<%
					Dim rsBlok
					Set rsBlok = Server.CreateObject("ADODB.Recordset")
					rsBlok.CursorLocation = adUseClient

			
					if len(trim(v_txtBlok)) > 0 then 
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
					end if
					
					rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsBlok.eof then
						DO while not rsBlok.eof
						if v_txtBlok = rsBlok("Blok") then
				%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%></option>
						<%else%>
							<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%></option>
						<%end if%>
				<%			rsBlok.movenext
						loop
					End if
					rsBlok.Close
					set rsBlok = nothing
				%>
			</select>		
		</td>
		
		
		<td align="center"><input type="submit" value="View" /></td>

	</tr>
    <tr>
	<td colspan="140" style=" ">
		<div id="header-wrapper">
			<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
				<% End if%>
			</div>
			<div id="subheader"class="centered">
				<h2>LAPORAN ANALISA PENCAPAIAN PRODUKSI</h2>
				<h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%></h5>

				<h3>
					<input id="ROTASI_button" class="check" type="button" onclick="javascript:toggle('ROTASI')" value="ROTASI">
					<input id="KG_button" class="check" type="button" onclick="javascript:toggle('KG')" value="KG">
					<input id="TON_HA_button" class="check" type="button" onclick="javascript:toggle('TON_HA')" value="TON/HA">
					<input id="RP_button" class="check" type="button" onclick="javascript:toggle('RP')" value="RP">
					<input id="RP_KG_button" class="check" type="button" onclick="javascript:toggle('RP_KG')" value="RP/KG">
					<input id="JJG_button" class="check" type="button" onclick="javascript:toggle('JJG')" value="JJG">
					<input id="BJR_button" class="check" type="button" onclick="javascript:toggle('BJR')" value="BJR">
					<input id="JJG_PKK_button" class="check" type="button" onclick="javascript:toggle('JJG_PKK')" value="JJG/PKK">
					<input id="HK_button" class="check" type="button" onclick="javascript:toggle('HK')" value="HK" >
					<input id="KG_HK_button" class="check" type="button" onclick="javascript:toggle('KG_HK')" value="KG/HK">
					
				</h3>
				</br>
			</div>
			
		</div>

	<%
		if len(trim(request("txtYear"))) = 0 then
			v_txtyear = year(Now)	 
		end if

		if len(trim(request("txtmonth"))) = 0 then
			v_txtmonth = Month(Now)	 
		end if
	%>

	</td>
	<td>
	<style>
		.uncheck{background-color:#ccc;}
		.check{background-color:#000;}
	</style>

	</td>
</tr>
</table>
</div> 

<div id="mydiv" style="margin-top:300px; z-index:10;">

<table border="0" width="100%" id="product-table" class="xthf-site" style="background-color:#FFFFFF;">
<thead>
	<tr>
		 <th rowspan="2" style="  background-color:#00A300;color:#fff; "></th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% 			
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% 			
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% 		
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% 		
	    end if %>

		<% if txtownership <> "--SUMMARY--"   and txtownership <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">OWNERSHIP</th>
		<% end if %>
		
		<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% 		
	    end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" And v_txtPlantingDate <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">TT</th>
		<% 		
	    end if %>


		<% if v_txtBlok <> "--SUMMARY--" And v_txtBlok <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<% 		
	    end if %>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">HA</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">PKK</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">PKK/HA</th>
		
		<th id="jan" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=1&LayerType=3">
		<font color="white">JAN</font>
		</a>
		</th>
		<th id="feb" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=2&LayerType=3">
		<font color="white">FEB</font>
		</a>
		</th>
        <th id="mar" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=3&LayerType=3">
		<font color="white">MAR</font>
		</a>
		</th>
        <th id="apr" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=4&LayerType=3">
		<font color="white">APR</font>
		</a>
		</th>
        <th id="may" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=5&LayerType=3">
		<font color="white">MAY</font>
		</a>
		</th>		
		<th id="jun" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=6&LayerType=3">
		<font color="white">JUN</font>
		</a>
		</th>
		<th id="jul" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=7&LayerType=3">
		<font color="white">JUL</font>
		</a>
		</th>
		<th id="aug" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=8&LayerType=3">
		<font color="white">AUG</font>
		</a>
		</th>
		<th id="sep" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=9&LayerType=3">
		<font color="white">SEP</font>
		</a>
		</th>
		<th id="oct" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=10&LayerType=3">
		<font color="white">OCT</font>
		</a>
		</th>
		<th id="nov" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=11&LayerType=3">
		<font color="white">NOV</font>
		</a>
		</th>
		<th id="des" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=12&LayerType=3">
		<font color="white">DES</font>
		</a>
		</th>
		<th id="sbi" colspan="10"  style="background-color:#00A300;color:#fff;">
		<a target="_blank" href="http://172.30.1.122/ffb/Content/FFB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&year=<%=v_txtYear%>&month=<%=Lpad(v_txtMonth,"0",2)%>&LayerType=2">
		<font color="white">SBI</font>
		</a>
		
		</th>
		
		<th id="bgt" colspan="9"  style="background-color:#00A300;color:#fff;">BGT SBI</th>
	</tr>

	<tr>
		<% DIM K
		   FOR K = 1 TO 14 
		 IF K <> 14 then %> 
		<th class="ROTASI" style="background-color:#00A300;color:#fff;">ROTASI</th>
		<% end if %>
		<th class="KG" style="background-color:#00A300;color:#fff;">KG</th>
		<th class="TON_HA" style="background-color:#00A300;color:#fff;">TON/HA</th>
		<th class="RP" style="background-color:#00A300;color:#fff;">RP</th>
		<th class="RP_KG" style="background-color:#00A300;color:#fff;">RP/KG</th>
		<th class="JJG" style="background-color:#00A300;color:#fff;">JJG</th>
		<th class="BJR" style="background-color:#00A300;color:#fff;">BJR</th>
		<th class="JJG_PKK" style="background-color:#00A300;color:#fff;">JJG/PKK</th>
		<th class="HK" style="background-color:#00A300;color:#fff;">HK</th>
		<th class="KG_HK" style="background-color:#00A300;color:#fff;">KG/HK</th>
		<% NEXT %>
	</tr>
	
        </thead>
 <tbody>
<%
Dim nsubhi, nsubshi, nsubbi, nsubsbi, nsubtonase_hi
Dim nsubtonase_shi, nsubtonase_bi, nsubtonase_sbi, nsubpom_hi, nsubpom_shi
Dim nsubpom_bi, nsubpom_sbi, nsubpom_tonase_hi, nsubpom_tonase_shi, nsubpom_tonase_bi, nsubpom_tonase_sbi, nsubrestan
Dim nsubbudjjg_shi,nsubbudjjg_sbi,nsubbudqty_shi,nsubbudqty_sbi,nsubbudmandays_shi,nsubbudmandays_sbi,nsubbudOutput_shi,nsubbudOutput_sbi
Dim nsubmandays_hi, nsubmandays_shi, nsubmandays_bi, nsubmandays_sbi
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim nsubOutput_hi, nsubOutput_shi, nsubOutput_bi, nsubOutput_sbi
Dim nOutput_hi, nOutput_shi, nOutput_bi, nOutput_sbi ,nbudOutput_shi,nbudOutput_sbi
Dim nha,npokok,nPrdton_01,nAmount_01,nJJG_01,nMandays_01 ,nPrdton_02,nAmount_02,nJJG_02,nMandays_02 ,nPrdton_03,nAmount_03,nJJG_03,nMandays_03 ,nPrdton_04,nAmount_04,nJJG_04,nMandays_04 
Dim nPrdton_05,nAmount_05,nJJG_05,nMandays_05,nPrdton_06,nAmount_06,nJJG_06,nMandays_06 ,nPrdton_07,nAmount_07,nJJG_07,nMandays_07 ,nPrdton_08,nAmount_08,nJJG_08,nMandays_08  
Dim nPrdton_09,nAmount_09,nJJG_09,nMandays_09,nPrdton_10,nAmount_10,nJJG_10,nMandays_10 ,nPrdton_11,nAmount_11,nJJG_11,nMandays_11 ,nPrdton_12,nAmount_12,nJJG_12,nMandays_12  
Dim nPrdton_SBI,nAmount_SBI,nJJG_SBI,nMandays_SBI,nRotasi_SBI,nPrdton_BUD,nAmount_BUD,nJJG_BUD,nMandays_BUD 
	
dim nPrdton_(12)		  
dim nAmount_(12)		   
dim nJJG_(12)	  
dim nMandays_(12)
DIM nRotasi_(12)
DIM nRotasiMultiHa_(12)

	if not rsEPS.EOF then
		NO = 1
		nsubhi = 0
		nsubshi = 0
		nsubbi = 0
		nsubsbi = 0
		nsubtonase_hi = 0
		nsubtonase_shi = 0
		nsubtonase_bi = 0
		nsubtonase_sbi = 0
		nsubpom_hi = 0
		nsubpom_shi = 0
		nsubpom_bi = 0
		nsubpom_sbi = 0
		nsubpom_tonase_hi = 0
		nsubpom_tonase_shi = 0
		nsubpom_tonase_bi = 0
		nsubpom_tonase_sbi = 0
		nsubrestan = 0
		nbudOutput_shi = 0
		nbudOutput_sbi = 0
		
		nsubmandays_hi = 0
		nsubmandays_shi = 0
		nsubmandays_bi = 0
		nsubmandays_sbi = 0
		
		nsubOutput_hi = 0
		nsubOutput_shi = 0
		nsubOutput_bi = 0
		nsubOutput_sbi = 0
		
		nSubMandays_Amount_hi = 0
		nSubMandays_Amount_shi = 0
		
		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		while not rsEPS.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"></td>
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% tempPlantation = rsEPS("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsEPS("Plantation") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsEPS("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsEPS("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsEPS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsEPS("GroupRegion") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsEPS("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsEPS("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsEPS("Region") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("region")%>
							</a>
						</td>
						<% SprevRegion = rsEPS("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsEPS("Company")
					   tempfcba = "--ALL--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsEPS("Company") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("company")%>
							</a>
						</td>
						<% SprevCompany = rsEPS("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsEPS("fcba")
					   tempownership = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsEPS("fcba") then %>
						<td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if txtownership <> "--SUMMARY--" and txtownership <> "" then %>
				   <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsEPS("fcba")
					   else
					       tempfcba = ""
					   end if
					   if txtownership <> "--SUMMARY--"  and txtownership <> "" then
					      tempownership = rsEPS("ownership")
					   else
					      tempownership = ""
                       end if					   
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevownership <> rsEPS("ownership") then %>
					    <td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("ownership")%>
							</a>
						</td>
						<% sprevownership = rsEPS("ownership") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
				<% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> "" then %>
				   <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsEPS("fcba")
					   else
					       tempfcba = ""
					   end if
					   if txtownership <> "--SUMMARY--"  and txtownership <> "" then
					      tempownership = rsEPS("ownership")
					   else
					      tempownership = ""
                       end if
					   if v_txtDivision <> "--SUMMARY--"  and v_txtDivision <> "" then
					      tempDivision = rsEPS("division")
					   else
					      tempDivision = ""
                       end if					   
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevDivision <> rsEPS("division") then %>
					    <td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("division")%>
							</a>
						</td>
						<% sprevDivision = rsEPS("division") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>

				<% if v_txtPlantingDate <> "--SUMMARY--"  and  v_txtPlantingDate <> "" then %>
				   <%  if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsEPS("fcba")
					   else
					       tempfcba = ""
					   end if
					   if txtownership <> "--SUMMARY--"  and txtownership <> "" then
					      tempownership = rsEPS("ownership")
					   else
					      tempownership = ""
                       end if
					   if v_txtDivision <> "--SUMMARY--"  and v_txtDivision <> "" then
					      tempDivision = rsEPS("division")
					   else
					      tempDivision = ""
                       end if	
					   tempPlantingDate = rsEPS("plantingyear")
					   tempBlok = "--ALL--"					   
					 %>
					
					<% tempPlantingDate = rsEPS("plantingyear") %>
					
					
				    <% if sprevPlantingYear <> rsEPS("plantingyear") then %>
				        <td align="center">
							<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("plantingyear")%>
							</a>
						</td>
						<% sprevPlantingYear = rsEPS("plantingyear") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>


				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then %>
				    <% if sprevFieldCode <> rsEPS("fieldcode") then %>				
					    <td align="center"><%=rsEPS("fieldcode")%></td>				
						<% sprevFieldCode = rsEPS("fieldcode") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>				
				
			
				<td align="right"><%=FormatNumber(rsEPS("HA"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("POKOK"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("POKOK_PER_HA"),2)%></td>
				
	
			<% 
			   SBIRotasi = 0
			   FOR K = 1 TO 12 %>	

				<td class="ROTASI" align="right"><%=FormatNumber(rsEPS("ROTASI_" & lpad(K,"0",2)),1)%></td>
				<td class="KG" align="right"><%=FormatNumber( cdbl(rsEPS("Prdton_" & lpad(K,"0",2))) * 1000,2)%></td>
				<td class="TON_HA" align="right"><%=FormatNumber(rsEPS("Prdton_HA_" & lpad(K,"0",2)),2)%></td>
				<td class="RP" align="right"><%=FormatNumber(rsEPS("Amount_" & lpad(K,"0",2)),2)%></td>
				<td class="RP_KG" align="right"><%=FormatNumber(rsEPS("Amount_KG_" & lpad(K,"0",2)),2)%></td>
				<td class="JJG" align="right"><%=FormatNumber(rsEPS("JJG_" & lpad(K,"0",2)),2)%></td>
				<td class="BJR" align="right"><%=FormatNumber(rsEPS("BJR_" & lpad(K,"0",2)),2)%></td>
				<td class="JJG_PKK" align="right"><%=FormatNumber(rsEPS("JJG_PKK_" & lpad(K,"0",2)),1)%></td>
				<td class="HK" align="right"><%=FormatNumber(rsEPS("Mandays_" & lpad(K,"0",2)),1)%></td>
                <td class="KG_HK" align="right"><%=FormatNumber(rsEPS("KG_HK_" & lpad(K,"0",2)),2)%></td>
			<% 
			     SBIRotasi = SBIRotasi +   cdbl(rsEPS("ROTASI_" & lpad(K,"0",2) ))
			   NEXT 
			   
			   %>
			
				<td class="ROTASI" align="right"><%=FormatNumber(SBIRotasi,2)%></td>
				<td class="KG" align="right"><%=FormatNumber(cdbl(rsEPS("prod_ton_SBI"))* 1000,2)%></td>
				<td class="TON_HA" align="right"><%=FormatNumber(rsEPS("prod_ton_HA_SBI"),2)%></td>
				<td class="RP" align="right"><%=FormatNumber(rsEPS("Amount_SBI"),2)%></td>
				<td class="RP_KG" align="right"><%=FormatNumber(rsEPS("Amount_KG_SBI"),2)%></td>
				<td class="JJG" align="right"><%=FormatNumber(rsEPS("JJG_SBI"),2)%></td>
				<td class="BJR" align="right"><%=FormatNumber(rsEPS("BJR_SBI"),2)%></td>
				<td class="JJG_PKK" align="right"><%=FormatNumber(rsEPS("JJG_PKK_SBI"),1)%></td>
				<td class="HK" align="right"><%=FormatNumber(rsEPS("Mandays_SBI"),1)%></td>
                <td class="KG_HK" align="right"><%=FormatNumber(cdbl(rsEPS("KG_HK_SBI"))/1000,2)%></td>
				
				
				<td class="KG" align="right"><%=FormatNumber(cdbl(rsEPS("bud_ton_sbi"))* 1000,2)%></td>
				<td class="TON_HA" align="right"><%=FormatNumber(rsEPS("bud_prod_ton_HA_SBI"),2)%></td>
				<td class="RP" align="right"><%=FormatNumber(rsEPS("bud_Amount_sbi"),2)%></td>
				<td class="RP_KG" align="right"><%=FormatNumber(rsEPS("bud_Amount_KG_SBI"),2)%></td>
				<td class="JJG" align="right"><%=FormatNumber(rsEPS("bud_jjg_sbi"),2)%></td>
				<td class="BJR" align="right"><%=FormatNumber(rsEPS("BUD_BJR_SHI"),2)%></td>
				<td class="JJG_PKK" align="right"><%=FormatNumber(rsEPS("BUD_JJG_PKK_SBI"),1)%></td>
				<td class="HK" align="right"><%=FormatNumber(rsEPS("bud_Mandays_sbi"),1)%></td>
                <td class="KG_HK" align="right"><%=FormatNumber(cdbl(rsEPS("bud_KG_HK_SBI"))/1000,2)%></td>
			
				
								
				
			</tr>
<%
			  
			  
			   nha = nha + cdbl(rsEPS("HA"))
			   npokok = npokok + cdbl(rsEPS("POKOK"))
			    
				FOR i = 1 TO 12 
					'response.write("<script>alert('" & lpad(i,"0",2) & "');</script>")
					
					nPrdton_(i) = nPrdton_(i) + cdbl(rsEPS("Prdton_" & lpad(i,"0",2) ))	* 1000
					nAmount_(i)	= nAmount_(i) + cdbl(rsEPS("Amount_" & lpad(i,"0",2) ))		   
					nJJG_(i) =  nJJG_(i) + cdbl(rsEPS("JJG_" & lpad(i,"0",2) ))
					nMandays_(i) = nMandays_(i) + cdbl(rsEPS("Mandays_" & lpad(i,"0",2) ))		
					nRotasi_(i) = nRotasi_(i) + cdbl(rsEPS("ROTASI_" & lpad(i,"0",2) ))		
					nRotasi_SBI = nRotasi_SBI + cdbl(rsEPS("ROTASI_" & lpad(i,"0",2) ))	
                    nRotasiMultiHa_(i) = nRotasiMultiHa_(i) + cdbl(rsEPS("ROTASI_" & lpad(i,"0",2) ) )* cdbl(rsEPS("HA"))
					
				NEXT 
			
			
		
		'	   nPrdton_01   = nPrdton_01 + cdbl(rsEPS("Prdton_01"))			  
		'	   nAmount_01 = nAmount_01 + cdbl(rsEPS("Amount_01"))		   
		'	   nJJG_01 = nJJG_01 + cdbl(rsEPS("JJG_01"))		  
		'	   nMandays_01 = nMandays_01 + cdbl(rsEPS("Mandays_01"))
		'	   
		'	   nPrdton_02   = nPrdton_02 + cdbl(rsEPS("Prdton_02"))			  
		'	   nAmount_02 = nAmount_02 + cdbl(rsEPS("Amount_02"))		   
		'	   nJJG_02 = nJJG_02 + cdbl(rsEPS("JJG_02"))		  
		'	   nMandays_02 = nMandays_02 + cdbl(rsEPS("Mandays_02"))
        '        
	'		   nPrdton_03   = nPrdton_03 + cdbl(rsEPS("Prdton_03"))			  
'			   nAmount_03 = nAmount_03 + cdbl(rsEPS("Amount_03"))		   
'			   nJJG_03 = nJJG_03 + cdbl(rsEPS("JJG_03"))		  
'			   nMandays_03 = nMandays_03 + cdbl(rsEPS("Mandays_03"))
'			    
'			   nPrdton_04   = nPrdton_04 + cdbl(rsEPS("Prdton_04"))			  
'			   nAmount_04 = nAmount_04 + cdbl(rsEPS("Amount_04"))		   
'			   nJJG_04 = nJJG_04 + cdbl(rsEPS("JJG_04"))		  
'			   nMandays_04 = nMandays_04 + cdbl(rsEPS("Mandays_04"))
'			    
'			   nPrdton_05   = nPrdton_05 + cdbl(rsEPS("Prdton_05"))			  
'			   nAmount_05 = nAmount_05 + cdbl(rsEPS("Amount_05"))		   
''			   nJJG_05 = nJJG_05 + cdbl(rsEPS("JJG_05"))		  
	'		   nMandays_05 = nMandays_05 + cdbl(rsEPS("Mandays_05"))
	'		    
	'		   nPrdton_06   = nPrdton_06 + cdbl(rsEPS("Prdton_06"))			  
	'		   nAmount_06 = nAmount_06 + cdbl(rsEPS("Amount_06"))		   
	'		   nJJG_06 = nJJG_06 + cdbl(rsEPS("JJG_06"))		  
	'		   nMandays_06 = nMandays_06 + cdbl(rsEPS("Mandays_06"))
	'		    
	'		   nPrdton_07   = nPrdton_07 + cdbl(rsEPS("Prdton_07"))			  
	'		   nAmount_07 = nAmount_07 + cdbl(rsEPS("Amount_07"))		   
	'		   nJJG_07 = nJJG_07 + cdbl(rsEPS("JJG_07"))		  
	'		   nMandays_07 = nMandays_07 + cdbl(rsEPS("Mandays_07"))
	'		    
	'		   nPrdton_08   = nPrdton_08 + cdbl(rsEPS("Prdton_08"))			  
	'		   nAmount_08 = nAmount_08 + cdbl(rsEPS("Amount_08"))		   
	'		   nJJG_08 = nJJG_08 + cdbl(rsEPS("JJG_08"))		  
	'		   nMandays_08 = nMandays_08 + cdbl(rsEPS("Mandays_08"))
	'		    
	'		   nPrdton_09   = nPrdton_09 + cdbl(rsEPS("Prdton_09"))			  
	'		   nAmount_09 = nAmount_09 + cdbl(rsEPS("Amount_09"))		   
	'		   nJJG_09 = nJJG_09 + cdbl(rsEPS("JJG_09"))		  
	'		   nMandays_09 = nMandays_09 + cdbl(rsEPS("Mandays_09"))
	'		    
	'		   nPrdton_10   = nPrdton_10 + cdbl(rsEPS("Prdton_10"))			  
	'		   nAmount_10 = nAmount_10 + cdbl(rsEPS("Amount_10"))		   
	'		   nJJG_10 = nJJG_10 + cdbl(rsEPS("JJG_10"))		  
	'		   nMandays_10 = nMandays_10 + cdbl(rsEPS("Mandays_10"))
	'		   
	'		   nPrdton_11   = nPrdton_11 + cdbl(rsEPS("Prdton_11"))			  
	'		   nAmount_11 = nAmount_11 + cdbl(rsEPS("Amount_11"))		   
	'		   nJJG_11 = nJJG_11 + cdbl(rsEPS("JJG_11"))		  
	'		   nMandays_11 = nMandays_11 + cdbl(rsEPS("Mandays_11")) 
	'		   	    
	'		   nPrdton_12   = nPrdton_12 + cdbl(rsEPS("Prdton_12"))			  
	'		   nAmount_12 = nAmount_12 + cdbl(rsEPS("Amount_12"))		   
	'		   nJJG_12 = nJJG_12 + cdbl(rsEPS("JJG_12"))		  
	'		   nMandays_12 = nMandays_12 + cdbl(rsEPS("Mandays_12"))
			   
			   nPrdton_SBI   = nPrdton_SBI + cdbl(rsEPS("prod_ton_SBI")) * 1000			  
			   nAmount_SBI = nAmount_SBI + cdbl(rsEPS("Amount_SBI"))		   
			   nJJG_SBI = nJJG_SBI + cdbl(rsEPS("JJG_SBI"))		  
			   nMandays_SBI = nMandays_SBI + cdbl(rsEPS("Mandays_SBI"))
			   
			   
			    
			   nPrdton_bud   = nPrdton_bud + cdbl(rsEPS("bud_ton_sbi"))	* 1000		  
			   nAmount_bud = nAmount_bud + cdbl(rsEPS("bud_Amount_sbi"))		   
			   nJJG_bud = nJJG_bud + cdbl(rsEPS("bud_jjg_sbi"))		  
			   nMandays_bud = nMandays_bud + cdbl(rsEPS("bud_Mandays_sbi"))
			   NO = NO + 1
			rsEPS.MoveNext			
		Wend
	end if
	rsEPS.Close
	set rsEPS = Nothing
	
	' tambahan akbar
	
	
	
	
	if nsubpom_tonase_hi > 0 then
	   nSubMandays_Amount_hi = (nsubmandays_amount_hi/(nsubpom_tonase_hi*1000)) 
	else
	   nSubMandays_Amount_hi = 0
	end if
	if nsubpom_tonase_shi > 0 then
	   nSubMandays_Amount_shi = (nsubmandays_amount_shi/(nsubpom_tonase_shi*1000)) 
	else
	   nSubMandays_Amount_shi = 0
	end if
	'nRotasi_SBI = nRotasi_SBI/NO
	' tambahan akbar
%>

			<tr bgcolor="#F4F4A0">
		
				<td align="center" colspan="<%= 1+ nSpanTotal %>">TOTAL</td>				
				<td align="right"><%=FormatNumber(nha,2)%></td>
				<td align="right"><%=FormatNumber(npokok,2)%></td>
				
				<% IF cdbl(nha) > 0 then %>
				  <td align="right"><%=FormatNumber(cdbl(npokok)/cdbl(nha),2)%></td>
				<% ELSE %>
				   <td align="right">0</td>
				<% END IF%>
				

				<% FOR i = 1 TO 12 %>
				   <%
     				  if cdbl(NO) > 0 then
					      if cdbl(nha) = 0 then
						      nRotasi_(i) = 0
						  else
						      nRotasi_(i) =  cdbl(nRotasiMultiHa_(i))/cdbl(nha)  ''nRotasi_(i)/NO
					
					      end if
					  end if
				  	 %> 
					<td class="ROTASI" align="right"><%=FormatNumber(nRotasi_(i),1)%></td>
					<td class="KG" align="right"><%=FormatNumber(nPrdton_(i),2)%></td>
					
					<% IF cdbl(nha) > 0 then %>
					  <td class="TON_HA" align="right"><%=FormatNumber(cdbl(nPrdton_(i))/cdbl(nha)/ 1000,2)%></td>
					<% ELSE %>
					   <td class="TON_HA" align="right">0</td>
					<% END IF%>
					
					<td class="RP" align="right"><%=FormatNumber(nAmount_(i),2)%></td>
					
					<% IF cdbl(nPrdton_(i)) > 0 then %>
					  <td class="RP_KG" align="right"><%=FormatNumber(cdbl(nAmount_(i))/(cdbl(nPrdton_(i))),2)%></td>
					<% ELSE %>
					   <td class="RP_KG" align="right">0</td>
					<% END IF%>		
					
					<td class="JJG" align="right"><%=FormatNumber(nJJG_(i),2)%></td>
					
					<% IF cdbl(nJJG_(i)) > 0 then %>
					  <td class="BJR" align="right"><%=FormatNumber( (cdbl(nPrdton_(i)))/cdbl(nJJG_(i)),2)%></td>
					<% ELSE %>
					   <td class="BJR" align="right">0</td>
					<% END IF%>
					
					<% IF cdbl(npokok) > 0 then %>
					  <td class="JJG_PKK" align="right"><%=FormatNumber(cdbl(nJJG_(i))/cdbl(npokok),1)%></td>
					<% ELSE %>
					   <td class="JJG_PKK" align="right">0</td>
					<% END IF%>
					
					<td class="HK" align="right"><%=FormatNumber(nMandays_(i),1)%></td>
					
					<% IF cdbl(nMandays_(i)) > 0 then %>
					  <td class="KG_HK" align="right"><%=FormatNumber((cdbl(nPrdton_(i)))/cdbl(nMandays_(i)),2)%></td>
					<% ELSE %>
					   <td class="KG_HK" align="right">0</td>
					<% END IF%>
					
				
				<% next %>	
				
				<td class="ROTASI" align="right"><%=FormatNumber(nRotasi_SBI,2)%></td>
				<td class="KG" align="right"><%=FormatNumber(nPrdton_SBI,2)%></td>
				
				<% IF cdbl(nha) > 0 then %>
				  <td class="TON_HA" align="right"><%=FormatNumber(cdbl(nPrdton_SBI)/cdbl(nha)/1000,2)%></td>
				<% ELSE %>
				   <td class="TON_HA" align="right">0</td>
				<% END IF%>
				
				<td class="RP" align="right"><%=FormatNumber(nAmount_SBI,2)%></td>
				
				<% IF cdbl(nPrdton_SBI) > 0 then %>
				  <td class="RP_KG" align="right"><%=FormatNumber(cdbl(nAmount_SBI)/(cdbl(nPrdton_SBI)),2)%></td>
				<% ELSE %>
				   <td class="RP_KG" align="right">0</td>
				<% END IF%>
				
				<td class="JJG" align="right"><%=FormatNumber(nJJG_SBI,2)%></td>
				
				<% IF cdbl(nJJG_SBI) > 0 then %>
				  <td class="BJR" align="right"><%=FormatNumber((cdbl(nPrdton_SBI))/cdbl(nJJG_SBI),2)%></td>
				<% ELSE %>
				   <td class="BJR" align="right">0</td>
				<% END IF%>
				
				<% IF cdbl(npokok) > 0 then %>
				  <td class="JJG_PKK" align="right"><%=FormatNumber(cdbl(nJJG_SBI)/cdbl(npokok),1)%></td>
				<% ELSE %>
				   <td class="JJG_PKK" align="right">0</td>
				<% END IF%>
				
				<td class="HK" align="right"><%=FormatNumber(nMandays_SBI,2)%></td>
				
				<% IF cdbl(nMandays_SBI) > 0 then %>
				  <td class="KG_HK" align="right"><%=FormatNumber((cdbl(nPrdton_SBI)*1000)/cdbl(nMandays_SBI),2)%></td>
				<% ELSE %>
				   <td class="KG_HK" align="right">0</td>
				<% END IF%>
				
     
				
				<td class="KG" align="right"><%=FormatNumber(nPrdton_BUD,2)%></td>
				<% IF cdbl(nha) > 0 then %>
				  <td class="TON_HA" align="right"><%=FormatNumber(cdbl(nPrdton_BUD)/cdbl(nha)/1000,2)%></td>
				<% ELSE %>
				   <td class="TON_HA" align="right">0</td>
				<% END IF%>
				
				<td class="RP" align="right"><%=FormatNumber(nAmount_BUD,2)%></td>
				
				<% IF cdbl(nPrdton_BUD) > 0 then %>
				  <td class="RP_KG" align="right"><%=FormatNumber(cdbl(nAmount_BUD)/(cdbl(nPrdton_BUD)),2)%></td>
				<% ELSE %>
				   <td class="RP_KG" align="right">0</td>
				<% END IF%>		
				
				<td class="JJG" align="right"><%=FormatNumber(nJJG_BUD,2)%></td>
				
				<% IF cdbl(nJJG_BUD) > 0 then %>
				  <td class="BJR" align="right"><%=FormatNumber((cdbl(nPrdton_BUD))/cdbl(nJJG_BUD),2)%></td>
				<% ELSE %>
				   <td class="BJR" align="right">0</td>
				<% END IF%>
				
				<% IF cdbl(npokok) > 0 then %>
				  <td class="JJG_PKK" align="right"><%=FormatNumber(cdbl(nJJG_BUD)/cdbl(npokok),1)%></td>
				<% ELSE %>
				   <td class="JJG_PKK" align="right">0</td>
				<% END IF%>
				
				<td class="HK" align="right"><%=FormatNumber(nMandays_BUD,2)%></td>
				
				<% IF cdbl(nMandays_BUD) > 0 then %>
				  <td class="KG_HK" align="right"><%=FormatNumber((cdbl(nPrdton_BUD)*1000)/cdbl(nMandays_BUD),2)%></td>
				<% ELSE %>
				   <td class="KG_HK" align="right">0</td>
				<% END IF%>
				
				
				
			</tr>
</tbody>
</table>
</form>

</div>
</body>
</html>
