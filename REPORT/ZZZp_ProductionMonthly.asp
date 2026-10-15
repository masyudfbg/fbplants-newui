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
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	
	        v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	= Request("txtPlantation")
			v_txtGroupRegion  	= Request("txtGroupRegion")
			v_txtRegion  		= Request("txtRegion")
			v_txtcompany  		= Request("txtCompany")
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtPlantingDate	= Request("txtPlantingDate")

			
			v_txtyear 		= Request("txtyear")
			v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 8
			
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
			    selectCons = selectCons & ",norut,fcba"
			    orderCons  = orderCons & ",norut "
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
            
			
	strMain = " SELECT   " & strFieldCons & " accountingyear,Division,fieldcode, plantingyear,HA,POKOK,POKOK_PER_HA," & _
					"	Prdton_01,decode(HA,0,0,Prdton_01/HA) Prdton_HA_01,Amount_01,decode(Prdton_01,0,0,Amount_01/(Prdton_01*1000)) Amount_KG_01,JJG_01 , decode(JJG_01,0,0,(Prdton_01*1000)/JJG_01) BJR_01,decode(POKOK,0,0,JJG_01/POKOK) JJG_PKK_01,Mandays_01,decode(Mandays_01,0,0,(Prdton_01*1000)/Mandays_01) KG_HK_01," & _
					"	Prdton_02,decode(HA,0,0,Prdton_02/HA) Prdton_HA_02,Amount_02,decode(Prdton_02,0,0,Amount_02/(Prdton_02*1000)) Amount_KG_02,JJG_02 , decode(JJG_02,0,0,(Prdton_02*1000)/JJG_02) BJR_02,decode(POKOK,0,0,JJG_02/POKOK) JJG_PKK_02,Mandays_02,decode(Mandays_02,0,0,(Prdton_02*1000)/Mandays_02) KG_HK_02," & _
					"	Prdton_03,decode(HA,0,0,Prdton_03/HA) Prdton_HA_03,Amount_03,decode(Prdton_03,0,0,Amount_03/(Prdton_03*1000)) Amount_KG_03,JJG_03 , decode(JJG_03,0,0,(Prdton_03*1000)/JJG_03) BJR_03,decode(POKOK,0,0,JJG_03/POKOK) JJG_PKK_03,Mandays_03,decode(Mandays_03,0,0,(Prdton_03*1000)/Mandays_03) KG_HK_03," & _
					"	Prdton_04,decode(HA,0,0,Prdton_04/HA) Prdton_HA_04,Amount_04,decode(Prdton_04,0,0,Amount_04/(Prdton_04*1000)) Amount_KG_04,JJG_04 , decode(JJG_04,0,0,(Prdton_04*1000)/JJG_04) BJR_04,decode(POKOK,0,0,JJG_04/POKOK) JJG_PKK_04,Mandays_04,decode(Mandays_04,0,0,(Prdton_04*1000)/Mandays_04) KG_HK_04," & _
					"	Prdton_05,decode(HA,0,0,Prdton_05/HA) Prdton_HA_05,Amount_05,decode(Prdton_05,0,0,Amount_05/(Prdton_05*1000)) Amount_KG_05,JJG_05 , decode(JJG_05,0,0,(Prdton_05*1000)/JJG_05) BJR_05,decode(POKOK,0,0,JJG_05/POKOK) JJG_PKK_05,Mandays_05,decode(Mandays_05,0,0,(Prdton_05*1000)/Mandays_05) KG_HK_05," & _
					"	Prdton_06,decode(HA,0,0,Prdton_06/HA) Prdton_HA_06,Amount_06,decode(Prdton_06,0,0,Amount_06/(Prdton_06*1000)) Amount_KG_06,JJG_06 , decode(JJG_06,0,0,(Prdton_06*1000)/JJG_06) BJR_06,decode(POKOK,0,0,JJG_06/POKOK) JJG_PKK_06,Mandays_06,decode(Mandays_06,0,0,(Prdton_06*1000)/Mandays_06) KG_HK_06," & _
					"	Prdton_07,decode(HA,0,0,Prdton_07/HA) Prdton_HA_07,Amount_07,decode(Prdton_07,0,0,Amount_07/(Prdton_07*1000)) Amount_KG_07,JJG_07 , decode(JJG_07,0,0,(Prdton_07*1000)/JJG_07) BJR_07,decode(POKOK,0,0,JJG_07/POKOK) JJG_PKK_07,Mandays_07,decode(Mandays_07,0,0,(Prdton_07*1000)/Mandays_07) KG_HK_07," & _
					"	Prdton_08,decode(HA,0,0,Prdton_08/HA) Prdton_HA_08,Amount_08,decode(Prdton_08,0,0,Amount_08/(Prdton_08*1000)) Amount_KG_08,JJG_08 , decode(JJG_08,0,0,(Prdton_08*1000)/JJG_08) BJR_08,decode(POKOK,0,0,JJG_08/POKOK) JJG_PKK_08,Mandays_08,decode(Mandays_08,0,0,(Prdton_08*1000)/Mandays_08) KG_HK_08," & _
					"	Prdton_09,decode(HA,0,0,Prdton_09/HA) Prdton_HA_09,Amount_09,decode(Prdton_09,0,0,Amount_09/(Prdton_09*1000)) Amount_KG_09,JJG_09 , decode(JJG_09,0,0,(Prdton_09*1000)/JJG_09) BJR_09,decode(POKOK,0,0,JJG_09/POKOK) JJG_PKK_09,Mandays_09,decode(Mandays_09,0,0,(Prdton_09*1000)/Mandays_09) KG_HK_09," & _
					"	Prdton_10,decode(HA,0,0,Prdton_10/HA) Prdton_HA_10,Amount_10,decode(Prdton_10,0,0,Amount_10/(Prdton_10*1000)) Amount_KG_10,JJG_10 , decode(JJG_10,0,0,(Prdton_10*1000)/JJG_10) BJR_10,decode(POKOK,0,0,JJG_10/POKOK) JJG_PKK_10,Mandays_10,decode(Mandays_10,0,0,(Prdton_10*1000)/Mandays_10) KG_HK_10," & _
					"	Prdton_11,decode(HA,0,0,Prdton_11/HA) Prdton_HA_11,Amount_11,decode(Prdton_11,0,0,Amount_11/(Prdton_11*1000)) Amount_KG_11,JJG_11 , decode(JJG_11,0,0,(Prdton_11*1000)/JJG_11) BJR_11,decode(POKOK,0,0,JJG_11/POKOK) JJG_PKK_11,Mandays_11,decode(Mandays_11,0,0,(Prdton_11*1000)/Mandays_11) KG_HK_11," & _
					"	Prdton_12,decode(HA,0,0,Prdton_12/HA) Prdton_HA_12,Amount_12,decode(Prdton_12,0,0,Amount_12/(Prdton_12*1000)) Amount_KG_12,JJG_12 , decode(JJG_12,0,0,(Prdton_12*1000)/JJG_12) BJR_12,decode(POKOK,0,0,JJG_12/POKOK) JJG_PKK_12,Mandays_12,decode(Mandays_12,0,0,(Prdton_12*1000)/Mandays_12) KG_HK_12," & _
					"	NVL(Prdton_01,0)+NVL(Prdton_02,0)+NVL(Prdton_03,0)+NVL(Prdton_04,0)+NVL(Prdton_05,0)+NVL(Prdton_06,0)+NVL(Prdton_07,0)+NVL(Prdton_08,0)+NVL(Prdton_09,0)+NVL(Prdton_10,0)+NVL(Prdton_11,0)+NVL(Prdton_12,0) prod_ton_SBI," & _
					"	NVL(Amount_01,0)+NVL(Amount_02,0)+NVL(Amount_03,0)+NVL(Amount_04,0)+NVL(Amount_05,0)+NVL(Amount_06,0)+NVL(Amount_07,0)+NVL(Amount_08,0)+NVL(Amount_09,0)+NVL(Amount_10,0)+NVL(Amount_11,0)+NVL(Amount_12,0) Amount_SBI," & _
					"	NVL(JJG_01,0)+NVL(JJG_02,0)+NVL(JJG_03,0)+NVL(JJG_04,0)+NVL(JJG_05,0)+NVL(JJG_06,0)+NVL(JJG_07,0)+NVL(JJG_08,0)+NVL(JJG_09,0)+NVL(JJG_10,0)+NVL(JJG_11,0)+NVL(JJG_12,0) JJG_SBI," & _
					"	NVL(Mandays_01,0)+NVL(Mandays_02,0)+NVL(Mandays_03,0)+NVL(Mandays_04,0)+NVL(Mandays_05,0)+NVL(Mandays_06,0)+NVL(Mandays_07,0)+NVL(Mandays_08,0)+NVL(Mandays_09,0)+NVL(Mandays_10,0)+NVL(Mandays_11,0)+NVL(Mandays_12,0) Mandays_SBI," & _
					"	nvl(budget_transaction_prodqty_api.GetQuantityUntilMonth(fcba,accountingyear,1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0)/1000 bud_ton_sbi," & _
					"	nvl(BUDGET_TRANSACTION_ESTATE_API.GetBudgetUntilMonth(fcba,accountingyear,1,1,'','6502101','','','',fieldcode,division,to_number(plantingyear),'',TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0) bud_Amount_sbi," & _      
					"	nvl(budget_transaction_prodqty_api.GetJanjangUntilMonth(fcba,accountingyear,1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0) bud_jjg_sbi," & _      
					"	nvl(budget_transaction_prodqty_api.GetMandaysUntilMonth(fcba,accountingyear,1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('05/30/2013', 'MM/DD/YYYY')),0) bud_Mandays_sbi" & _
					"	from ( select fcba,accountingyear,Division,fieldcode, plantingyear," & _
					"		  field_api.get_HA(fieldcode, fcba) HA,field_api.get_POKOK(fieldcode, fcba) POKOK," & _
					"		  field_api.get_POKOK_PER_HA(fieldcode, fcba) POKOK_PER_HA," & _
					"		  sum(Prdton_01) Prdton_01,sum(Prdton_02) Prdton_02,sum(Prdton_03) Prdton_03,sum(Prdton_04) Prdton_04,sum(Prdton_05) Prdton_05,sum(Prdton_06) Prdton_06," & _
					"		  sum(Prdton_07) Prdton_07,sum(Prdton_08) Prdton_08,sum(Prdton_09) Prdton_09,sum(Prdton_10) Prdton_10,sum(Prdton_11) Prdton_11,sum(Prdton_12) Prdton_12," & _
					"		  sum(Amount_01) Amount_01,sum(Amount_02) Amount_02,sum(Amount_03) Amount_03,sum(Amount_04) Amount_04,sum(Amount_05) Amount_05,sum(Amount_06) Amount_06," & _
					"		  sum(Amount_07) Amount_07,sum(Amount_08) Amount_08,sum(Amount_09) Amount_09,sum(Amount_10) Amount_10,sum(Amount_11) Amount_11,sum(Amount_12) Amount_12," & _
					"		  sum(JJG_01) JJG_01,sum(JJG_02) JJG_02,sum(JJG_03) JJG_03,sum(JJG_04) JJG_04,sum(JJG_05) JJG_05,sum(JJG_06) JJG_06," & _
					"		  sum(JJG_07) JJG_07,sum(JJG_08) JJG_08,sum(JJG_09) JJG_09,sum(JJG_10) JJG_10,sum(JJG_11) JJG_11,sum(JJG_12) JJG_12," & _
					"		  sum(Mandays_01) Mandays_01,sum(Mandays_02) Mandays_02,sum(Mandays_03) Mandays_03,sum(Mandays_04) Mandays_04,sum(Mandays_05) Mandays_05,sum(Mandays_06) Mandays_06," & _
					"		  sum(Mandays_07) Mandays_07,sum(Mandays_08) Mandays_08,sum(Mandays_09) Mandays_09,sum(Mandays_10) Mandays_10,sum(Mandays_11) Mandays_11,sum(Mandays_12) Mandays_12" & _
					"		  from (  select t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear," & _
					"						 sum(decode(t.accountingperiod, 1, nvl(t.totalDailyProdTonase,0), 0)) Prdton_01," & _
					"						 sum(decode(t.accountingperiod, 2, nvl(t.totalDailyProdTonase,0), 0)) Prdton_02," & _
					"						 sum(decode(t.accountingperiod, 3, nvl(t.totalDailyProdTonase,0), 0)) Prdton_03," & _
					"						 sum(decode(t.accountingperiod, 4, nvl(t.totalDailyProdTonase,0), 0)) Prdton_04," & _
					"						 sum(decode(t.accountingperiod, 5, nvl(t.totalDailyProdTonase,0), 0)) Prdton_05," & _
					"						 sum(decode(t.accountingperiod, 6, nvl(t.totalDailyProdTonase,0), 0)) Prdton_06," & _
					"						 sum(decode(t.accountingperiod, 7, nvl(t.totalDailyProdTonase,0), 0)) Prdton_07," & _
					"						 sum(decode(t.accountingperiod, 8, nvl(t.totalDailyProdTonase,0), 0)) Prdton_08," & _
					"						 sum(decode(t.accountingperiod, 9, nvl(t.totalDailyProdTonase,0), 0)) Prdton_09," & _
					"						 sum(decode(t.accountingperiod, 10, nvl(t.totalDailyProdTonase,0), 0)) Prdton_10," & _
					"						 sum(decode(t.accountingperiod, 11, nvl(t.totalDailyProdTonase,0), 0)) Prdton_11," & _
					"						 sum(decode(t.accountingperiod, 12, nvl(t.totalDailyProdTonase,0), 0)) Prdton_12," & _
					"						 sum(decode(t.accountingperiod, 1, t.Amount, 0)) Amount_01," & _
					"						 sum(decode(t.accountingperiod, 2, t.Amount, 0)) Amount_02," & _
					"						 sum(decode(t.accountingperiod, 3, t.Amount, 0)) Amount_03," & _
					"						 sum(decode(t.accountingperiod, 4, t.Amount, 0)) Amount_04," & _
					"						 sum(decode(t.accountingperiod, 5, t.Amount, 0)) Amount_05," & _
					"						 sum(decode(t.accountingperiod, 6, t.Amount, 0)) Amount_06," & _
					"						 sum(decode(t.accountingperiod, 7, t.Amount, 0)) Amount_07," & _
					"						 sum(decode(t.accountingperiod, 8, t.Amount, 0)) Amount_08," & _
					"						 sum(decode(t.accountingperiod, 9, t.Amount, 0)) Amount_09," & _
					"						 sum(decode(t.accountingperiod, 10, t.Amount, 0)) Amount_10," & _
					"						 sum(decode(t.accountingperiod, 11, t.Amount, 0)) Amount_11," & _
					"						 sum(decode(t.accountingperiod, 12, t.Amount, 0)) Amount_12," & _
					"						 0 JJG_01,0 JJG_02,0 JJG_03,0 JJG_04,0 JJG_05,0 JJG_06," & _
					"						 0 JJG_07,0 JJG_08,0 JJG_09,0 JJG_10,0 JJG_11,0 JJG_12," & _
					"						 0 Mandays_01,0 Mandays_02,0 Mandays_03,0 Mandays_04,0 Mandays_05,0 Mandays_06," & _
					"						 0 Mandays_07,0 Mandays_08,0 Mandays_09,0 Mandays_10,0 Mandays_11,0 Mandays_12" & _
					"					from view_dailyharvesting_prod t" & _
					"				   where to_char(t.fddate, 'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
					"					 and to_char(t.fddate, 'YYYY') = '" & v_txtyear &"'" & _
					"				   group by t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear" & _
					"				  union " & _
					"				  select t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear," & _
					"						 0 Prdton_01,0 Prdton_02,0 Prdton_03,0 Prdton_04,0 Prdton_05,0 Prdton_06," & _
					"						 0 Prdton_07,0 Prdton_08,0 Prdton_09,0 Prdton_10,0 Prdton_11,0 Prdton_12," & _
					"						 0 Amount_01,0 Amount_02,0 Amount_03,0 Amount_04,0 Amount_05,0 Amount_06," & _
					"						 0 Amount_07,0 Amount_08,0 Amount_09,0 Amount_10,0 Amount_11,0 Amount_12," & _
					"						 sum(decode(t.accountingperiod, 1, t.TOTALDAILYRECEIVE, 0)) JJG_01," & _
					"						 sum(decode(t.accountingperiod, 2, t.TOTALDAILYRECEIVE, 0)) JJG_02," & _
					"						 sum(decode(t.accountingperiod, 3, t.TOTALDAILYRECEIVE, 0)) JJG_03," & _
					"						 sum(decode(t.accountingperiod, 4, t.TOTALDAILYRECEIVE, 0)) JJG_04," & _
					"						 sum(decode(t.accountingperiod, 5, t.TOTALDAILYRECEIVE, 0)) JJG_05," & _
					"						 sum(decode(t.accountingperiod, 6, t.TOTALDAILYRECEIVE, 0)) JJG_06," & _
					"						 sum(decode(t.accountingperiod, 7, t.TOTALDAILYRECEIVE, 0)) JJG_07," & _
					"						 sum(decode(t.accountingperiod, 8, t.TOTALDAILYRECEIVE, 0)) JJG_08," & _
					"						 sum(decode(t.accountingperiod, 9, t.TOTALDAILYRECEIVE, 0)) JJG_09," & _
					"						 sum(decode(t.accountingperiod, 10, t.TOTALDAILYRECEIVE, 0)) JJG_10," & _
					"						 sum(decode(t.accountingperiod, 11, t.TOTALDAILYRECEIVE, 0)) JJG_11," & _
					"						 sum(decode(t.accountingperiod, 12, t.TOTALDAILYRECEIVE, 0)) JJG_12," & _
					"						 0 Mandays_01,0 Mandays_02,0 Mandays_03,0 Mandays_04,0 Mandays_05,0 Mandays_06," & _
					"						 0 Mandays_07,0 Mandays_08,0 Mandays_09,0 Mandays_10,0 Mandays_11,0 Mandays_12" & _
					"					from view_dailyharvesting_mill t" & _
					"				   where to_char(t.receptiondate, 'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
					"					 and to_char(t.receptiondate, 'YYYY') = '" & v_txtyear &"'" & _
					"				   group by t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear" & _
					"				  union" & _
					"				  select t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear," & _
					"						  0 Prdton_01,0 Prdton_02,0 Prdton_03,0 Prdton_04,0 Prdton_05,0 Prdton_06," & _
					"						 0 Prdton_07,0 Prdton_08,0 Prdton_09,0 Prdton_10,0 Prdton_11,0 Prdton_12," & _
					"						 0 Amount_01,0 Amount_02,0 Amount_03,0 Amount_04,0 Amount_05,0 Amount_06," & _
					"						 0 Amount_07,0 Amount_08,0 Amount_09,0 Amount_10,0 Amount_11,0 Amount_12," & _
					"						 0 JJG_01,0 JJG_02,0 JJG_03,0 JJG_04,0 JJG_05,0 JJG_06," & _
					"						 0 JJG_07,0 JJG_08,0 JJG_09,0 JJG_10,0 JJG_11,0 JJG_12," & _
					"						 sum(decode(t.accountingperiod, 1, t.totalDailyMandays, 0)) Mandays_01," & _
					"						 sum(decode(t.accountingperiod, 2, t.totalDailyMandays, 0)) Mandays_02," & _
					"						 sum(decode(t.accountingperiod, 3, t.totalDailyMandays, 0)) Mandays_03," & _
					"						 sum(decode(t.accountingperiod, 4, t.totalDailyMandays, 0)) Mandays_04," & _
					"						 sum(decode(t.accountingperiod, 5, t.totalDailyMandays, 0)) Mandays_05," & _
					"						 sum(decode(t.accountingperiod, 6, t.totalDailyMandays, 0)) Mandays_06," & _
					"						 sum(decode(t.accountingperiod, 7, t.totalDailyMandays, 0)) Mandays_07," & _
					"						 sum(decode(t.accountingperiod, 8, t.totalDailyMandays, 0)) Mandays_08," & _
					"						 sum(decode(t.accountingperiod, 9, t.totalDailyMandays, 0)) Mandays_09," & _
					"						 sum(decode(t.accountingperiod, 10, t.totalDailyMandays, 0)) Mandays_10," & _
					"						 sum(decode(t.accountingperiod, 11, t.totalDailyMandays, 0)) Mandays_11," & _
					"						 sum(decode(t.accountingperiod, 12, t.totalDailyMandays, 0)) Mandays_12" & _
					"					from view_dailyharvesting_mandays t" & _
					"				   where to_char(t.fddate, 'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
					"					 and to_char(t.fddate, 'YYYY') = '" & v_txtyear &"'" & _
					"				   group by t.fcba,t.accountingyear,Division,t.fieldcode,t.plantingyear" & _
					"			  )" & _
					"		   group by fcba, accountingyear, Division, fieldcode, plantingyear" & _
					"		) "
	
	
								 
	 strEPS2  =  " ,SUM(HA) HA,SUM(POKOK) POKOK,SUM(POKOK_PER_HA) POKOK_PER_HA" & _
	             " ,SUM(Prdton_01) Prdton_01, SUM(Prdton_HA_01) Prdton_HA_01,SUM(Amount_01) Amount_01,SUM(Amount_KG_01) Amount_KG_01,SUM(JJG_01) JJG_01,SUM(BJR_01) BJR_01,decode(SUM(POKOK),0,0,SUM(JJG_01)/SUM(POKOK)) JJG_PKK_01,SUM(Mandays_01) Mandays_01,SUM(KG_HK_01) KG_HK_01" & _
				 " ,SUM(Prdton_02) Prdton_02, SUM(Prdton_HA_02) Prdton_HA_02,SUM(Amount_02) Amount_02,SUM(Amount_KG_02) Amount_KG_02,SUM(JJG_02) JJG_02,SUM(BJR_02) BJR_02,decode(SUM(POKOK),0,0,SUM(JJG_02)/SUM(POKOK)) JJG_PKK_02,SUM(Mandays_02) Mandays_02,SUM(KG_HK_02) KG_HK_02" & _
			     " ,SUM(Prdton_03) Prdton_03, SUM(Prdton_HA_03) Prdton_HA_03,SUM(Amount_03) Amount_03,SUM(Amount_KG_03) Amount_KG_03,SUM(JJG_03) JJG_03,SUM(BJR_03) BJR_03,decode(SUM(POKOK),0,0,SUM(JJG_03)/SUM(POKOK)) JJG_PKK_03,SUM(Mandays_03) Mandays_03,SUM(KG_HK_03) KG_HK_03" & _
                 " ,SUM(Prdton_04) Prdton_04, SUM(Prdton_HA_04) Prdton_HA_04,SUM(Amount_04) Amount_04,SUM(Amount_KG_04) Amount_KG_04,SUM(JJG_04) JJG_04,SUM(BJR_04) BJR_04,decode(SUM(POKOK),0,0,SUM(JJG_04)/SUM(POKOK)) JJG_PKK_04,SUM(Mandays_04) Mandays_04,SUM(KG_HK_04) KG_HK_04" & _
                 " ,SUM(Prdton_05) Prdton_05, SUM(Prdton_HA_05) Prdton_HA_05,SUM(Amount_05) Amount_05,SUM(Amount_KG_05) Amount_KG_05,SUM(JJG_05) JJG_05,SUM(BJR_05) BJR_05,decode(SUM(POKOK),0,0,SUM(JJG_05)/SUM(POKOK)) JJG_PKK_05,SUM(Mandays_05) Mandays_05,SUM(KG_HK_05) KG_HK_05" & _
                 " ,SUM(Prdton_06) Prdton_06, SUM(Prdton_HA_06) Prdton_HA_06,SUM(Amount_06) Amount_06,SUM(Amount_KG_06) Amount_KG_06,SUM(JJG_06) JJG_06,SUM(BJR_06) BJR_06,decode(SUM(POKOK),0,0,SUM(JJG_06)/SUM(POKOK))  JJG_PKK_06,SUM(Mandays_06) Mandays_06,SUM(KG_HK_06) KG_HK_06" & _
                 " ,SUM(Prdton_07) Prdton_07, SUM(Prdton_HA_07) Prdton_HA_07,SUM(Amount_07) Amount_07,SUM(Amount_KG_07) Amount_KG_07,SUM(JJG_07) JJG_07,SUM(BJR_07) BJR_07,decode(SUM(POKOK),0,0,SUM(JJG_07)/SUM(POKOK))  JJG_PKK_07,SUM(Mandays_07) Mandays_07,SUM(KG_HK_07) KG_HK_07" & _
                 " ,SUM(Prdton_08) Prdton_08, SUM(Prdton_HA_08) Prdton_HA_08,SUM(Amount_08) Amount_08,SUM(Amount_KG_08) Amount_KG_08,SUM(JJG_08) JJG_08,SUM(BJR_08) BJR_08,decode(SUM(POKOK),0,0,SUM(JJG_08)/SUM(POKOK))  JJG_PKK_08,SUM(Mandays_08) Mandays_08,SUM(KG_HK_08) KG_HK_08" & _
				 " ,SUM(Prdton_09) Prdton_09, SUM(Prdton_HA_09) Prdton_HA_09,SUM(Amount_09) Amount_09,SUM(Amount_KG_09) Amount_KG_09,SUM(JJG_09) JJG_09,SUM(BJR_09) BJR_09,decode(SUM(POKOK),0,0,SUM(JJG_09)/SUM(POKOK))  JJG_PKK_09,SUM(Mandays_09) Mandays_09,SUM(KG_HK_09) KG_HK_09" & _
				 " ,SUM(Prdton_10) Prdton_10, SUM(Prdton_HA_10) Prdton_HA_10,SUM(Amount_10) Amount_10,SUM(Amount_KG_10) Amount_KG_10,SUM(JJG_10) JJG_10,SUM(BJR_10) BJR_10,decode(SUM(POKOK),0,0,SUM(JJG_10)/SUM(POKOK))  JJG_PKK_10,SUM(Mandays_10) Mandays_10,SUM(KG_HK_10) KG_HK_10" & _
				 " ,SUM(Prdton_11) Prdton_11, SUM(Prdton_HA_11) Prdton_HA_11,SUM(Amount_11) Amount_11,SUM(Amount_KG_11) Amount_KG_11,SUM(JJG_11) JJG_11,SUM(BJR_11) BJR_11,decode(SUM(POKOK),0,0,SUM(JJG_11)/SUM(POKOK))  JJG_PKK_11,SUM(Mandays_11) Mandays_11,SUM(KG_HK_11) KG_HK_11" & _
				 " ,SUM(Prdton_12) Prdton_12, SUM(Prdton_HA_12) Prdton_HA_12,SUM(Amount_12) Amount_12,SUM(Amount_KG_12) Amount_KG_12,SUM(JJG_12) JJG_12,SUM(BJR_12) BJR_12,decode(SUM(POKOK),0,0,SUM(JJG_12)/SUM(POKOK))  JJG_PKK_12,SUM(Mandays_12) Mandays_12,SUM(KG_HK_12) KG_HK_12" & _
                 " ,SUM(prod_ton_SBI) prod_ton_SBI,decode(SUM(HA),0,0,SUM(prod_ton_SBI)/SUM(HA)) prod_ton_HA_SBI,SUM(Amount_SBI) Amount_SBI,decode(SUM(prod_ton_SBI),0,0,SUM(Amount_SBI)/(SUM(prod_ton_SBI)*1000)) Amount_KG_SBI " & _
				 " ,SUM(JJG_SBI) JJG_SBI,SUM(NVL(BJR_01,0)+NVL(BJR_02,0)+NVL(BJR_03,0)+NVL(BJR_04,0)+NVL(BJR_05,0)+NVL(BJR_06,0)+NVL(BJR_07,0)+NVL(BJR_08,0)+NVL(BJR_09,0)+NVL(BJR_10,0)+NVL(BJR_11,0)+NVL(BJR_12,0)) BJR_SBI,DECODE(SUM(POKOK),0,0,SUM(JJG_SBI)/ SUM(POKOK)) JJG_PKK_SBI" & _
                 " ,SUM(Mandays_SBI) Mandays_SBI,SUM(NVL(KG_HK_01,0)+NVL(KG_HK_02,0)+NVL(KG_HK_03,0)+NVL(KG_HK_04,0)+NVL(KG_HK_05,0)+NVL(KG_HK_06,0)+NVL(KG_HK_07,0)+NVL(KG_HK_08,0)+NVL(KG_HK_09,0)+NVL(KG_HK_10,0)+NVL(KG_HK_11,0)+NVL(KG_HK_12,0)) KG_HK_SBI " & _ 
				 " ,SUM(bud_ton_sbi) bud_ton_sbi,decode(SUM(HA),0,0,SUM(bud_ton_sbi)/SUM(HA)) bud_prod_ton_HA_SBI,SUM(bud_Amount_sbi) bud_Amount_sbi,decode(SUM(bud_ton_sbi),0,0,SUM(bud_Amount_sbi)/(SUM(bud_ton_sbi) *1000)) bud_Amount_KG_SBI " & _ 
				 " ,SUM(bud_JJG_sbi) bud_jjg_sbi,decode(SUM(bud_JJG_sbi),0,0,(SUM(bud_ton_sbi)*1000)/SUM(bud_JJG_sbi)) BUD_BJR_SHI,decode(SUM(POKOK),0,0,SUM(bud_JJG_sbi)/SUM(POKOK)) BUD_JJG_PKK_SBI " & _ 
				 " ,SUM(bud_mandays_sbi) bud_mandays_sbi,decode(SUM(bud_mandays_sbi),0,0,(SUM(bud_ton_sbi)*1000)/ SUM(bud_mandays_sbi)) bud_KG_HK_SBI"
								 

	strEPS =  "select GroupPlantation " & selectCons & strEPS2 &"  From (" & strMain & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  GroupPlantation" & selectCons & "  Order by GroupPlantation" & orderCons  	

	' strEPS = strEPS &  " FROM ( " & strMain & " )  " 
	
	' if len(trim(v_txtFcba)) > 0 then 
	   ' if v_txtFcba <> "--SUMMARY--" then
	    ' strEPS = strEPS & " WHERE fcba = '" & v_txtFcba & "' " 
	   ' end if
	' end if

	
	
	' ' Where Condion
	' if len(trim(v_txtDivision)) > 0 then 
		' if v_txtDivision <> "--SUMMARY--" then
			' strEPS = strEPS & " AND DIVISION = '" & v_txtDivision & "'"
		' end if
	' end if
	' if len(trim(v_txtBlok)) > 0 then 
		' if  v_txtBlok <> "--SUMMARY--" then
			' strEPS = strEPS & " AND FIELDCODE = '" & v_txtBlok & "'"
		' end if
	' end if
	' if len(trim(v_txtPlantingDate)) > 0  then
		' if v_txtPlantingDate <> "--SUMMARY--" then
			' strEPS = strEPS & " AND PLANTINGYEAR = " & cint(v_txtplantingdate)
		' end if 
	' end if
    
	
	' ' Group By
	' strEPS = strEPS & " GROUP BY FCBA "
	' if len(trim(v_txtDivision)) > 0 then 
		' if v_txtDivision <> "--SUMMARY--" then
			' strEPS = strEPS & " , DIVISION "
		' end if
	' else
		' strEPS = strEPS & " , DIVISION "
	' end if

	' if len(trim(v_txtBlok)) > 0 then 
		' if v_txtBlok <> "--SUMMARY--" then
			' strEPS = strEPS & " , FIELDCODE "
		' end if
	' else
		' strEPS = strEPS & " , FIELDCODE "	
	' end if

	' if len(trim(v_txtPlantingDate)) > 0 then 
		' if v_txtPlantingDate <> "--SUMMARY--" then
			' strEPS = strEPS & " , PLANTINGYEAR "
		' end if
	' else
	    ' strEPS = strEPS & " , PLANTINGYEAR "
	' end if

	' ' Order By
	' strEPS = strEPS & " ORDER BY FCBA "
	' if len(trim(v_txtDivision)) > 0 then 
		' if v_txtDivision <> "--SUMMARY--" then
			' strEPS = strEPS & " , DIVISION "
		' end if
	' else
		' strEPS = strEPS & " , DIVISION "
	' end if

	' if len(trim(v_txtPlantingDate)) > 0 then 
		' if v_txtPlantingDate <> "--SUMMARY--" then
			' strEPS = strEPS & " , PLANTINGYEAR "
		' end if
	' else
	    ' strEPS = strEPS & " , PLANTINGYEAR "
	' end if
	
	' if len(trim(v_txtBlok)) > 0 then 
		' if v_txtBlok <> "--SUMMARY--" then
			' strEPS = strEPS & " , FIELDCODE "
		' end if
	' else
		' strEPS = strEPS & " , FIELDCODE "	
	' end if


				
Set rsEPS = Server.CreateObject("ADODB.Recordset")
'response.write strEPS

rsEPS.Open strEPS , connstr, adOpenStatic 


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Harian Produksi</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
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
<div class="top_select" style=" position:fixed; z-index:-999;  top:0; left:0px; width:100%;"> 
<form name="frmEntry" method="post" action="ZZZp_ProductionMonthly.asp">


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
					MinYr = Yr - 2
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
<td colspan="140" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
<!--<input type="button" value="Print Div" onClick="PrintElem('#mydiv')" />-->
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
		</div>
		<div id="subheader"class="centered" style="">
			<h2>LAPORAN ANALISA PENCAPAIAN PRODUKSI</h2>
			
				<h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> , KEBUN : <%=v_txtFcba%></h5>
		
			<h3></br></h3>
			<h3></br></h3>
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
</tr>
</table>

</div> 

<div id="mydiv" style="margin-top:280px;">
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
		<th colspan="9"   style="background-color:#00A300;color:#fff;">JAN</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">FEB</th>
        <th colspan="9"   style="background-color:#00A300;color:#fff;">MAR</th>
        <th colspan="9"   style="background-color:#00A300;color:#fff;">APR</th>
        <th colspan="9"   style="background-color:#00A300;color:#fff;">MAY</th>		
		<th colspan="9"   style="background-color:#00A300;color:#fff;">JUN</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">JUL</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">AUG</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">SEP</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">OCT</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">NOV</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">DES</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">SBI</th>
		<th colspan="9"   style="background-color:#00A300;color:#fff;">BGT SBI</th>
	</tr>
	
	<tr>
		<% DIM K
		   FOR K = 1 TO 14 %>
		<th  style="background-color:#00A300;color:#fff;">TON</th>
		<th  style="background-color:#00A300;color:#fff;">TON/HA</th>
		<th  style="background-color:#00A300;color:#fff;">RP</th>
		<th  style="background-color:#00A300;color:#fff;">RP/KG</th>
		<th  style="background-color:#00A300;color:#fff;">JJG</th>
		<th  style="background-color:#00A300;color:#fff;">BJR</th>
		<th  style="background-color:#00A300;color:#fff;">JJG/PKK</th>
		<th  style="background-color:#00A300;color:#fff;">HK</th>
		<th  style="background-color:#00A300;color:#fff;">KG/HK</th>
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
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsEPS("Plantation") then %>
						<td align="center">
							<a href="ZZZp_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsEPS("GroupRegion") then %>
						<td align="center">
							<a href="ZZZp_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsEPS("Region") then %>
						<td align="center">
							<a href="ZZZp_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsEPS("Company") then %>
						<td align="center">
							<a href="ZZZp_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsEPS("fcba") then %>
						<td align="center">
							<a href="ZZZp_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") %>	
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
							<a href="ZZZp_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
							<a href="ZZZp_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_01"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_01"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_02"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_02"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_03"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_03"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_04"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_04"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_05"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_05"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_06"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_06"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_07"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_07"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_08"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_08"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_09"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_09"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_10"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_10"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_11"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_11"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Prdton_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Prdton_HA_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_12"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_12"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("prod_ton_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("prod_ton_HA_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Amount_KG_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BJR_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("JJG_PKK_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Mandays_SBI"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("KG_HK_SBI"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("bud_ton_sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("bud_prod_ton_HA_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("bud_Amount_sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("bud_Amount_KG_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("bud_jjg_sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BUD_BJR_SHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("BUD_JJG_PKK_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("bud_Mandays_sbi"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("bud_KG_HK_SBI"),2)%></td>
				
								
				
			</tr>
<%
			NO = NO + 1
			' nsubhi = nsubhi + Cdbl(rsEPS("hi"))
			' nsubshi = nsubshi + Cdbl(rsEPS("shi"))			
			' nsubsbi = nsubsbi + Cdbl(rsEPS("sbi"))
			
			' nsubtonase_hi = nsubtonase_hi + Cdbl(rsEPS("tonase_hi"))
			' nsubtonase_shi = nsubtonase_shi + Cdbl(rsEPS("tonase_shi"))			
			' nsubtonase_sbi = nsubtonase_sbi + Cdbl(rsEPS("tonase_sbi"))
			
			' nsubpom_hi = nsubpom_hi + Cdbl(rsEPS("pom_hi"))
			' nsubpom_shi = nsubpom_shi + Cdbl(rsEPS("pom_shi"))
			' nsubpom_sbi = nsubpom_sbi + Cdbl(rsEPS("pom_sbi"))
			
			' nsubbudjjg_shi = nsubbudjjg_shi + Cdbl(rsEPS("bud_jjg_shi"))
			' nsubbudjjg_sbi = nsubbudjjg_sbi + Cdbl(rsEPS("bud_jjg_sbi"))
			
			' nsubpom_tonase_hi = nsubpom_tonase_hi + Cdbl(rsEPS("pom_tonase_hi"))
			' nsubpom_tonase_shi = nsubpom_tonase_shi + Cdbl(rsEPS("pom_tonase_shi"))
			' nsubpom_tonase_sbi = nsubpom_tonase_sbi + Cdbl(rsEPS("pom_tonase_sbi"))
			
			' nsubbudqty_shi = nsubbudqty_shi + Cdbl(rsEPS("bud_qty_shi"))
			' nsubbudqty_sbi = nsubbudqty_sbi + Cdbl(rsEPS("bud_qty_sbi"))
			
			' ' tambahan akbar
			' nsubmandays_amount_hi = nsubmandays_amount_hi + Cdbl(rsEPS("mandays_amount_hi"))
			' nsubmandays_amount_shi = nsubmandays_amount_shi + Cdbl(rsEPS("mandays_amount_shi"))
			' ' tambahan akbar
			
			
			' nsubbudmandays_shi = nsubbudmandays_shi + Cdbl(rsEPS("bud_mandays_shi"))
			' nsubbudmandays_sbi = nsubbudmandays_sbi + Cdbl(rsEPS("bud_mandays_sbi"))
			
			' nsubrestan = nsubrestan + Cdbl(rsEPS("restan"))
			
			' nsubmandays_hi = nsubmandays_hi + Cdbl(rsEPS("mandays_hi"))
			' nsubmandays_shi = nsubmandays_shi + Cdbl(rsEPS("mandays_shi"))
			
			' nsubmandays_sbi = nsubmandays_sbi + Cdbl(rsEPS("mandays_sbi"))		

			' if nsubmandays_hi > 0 then
			   ' nsubOutput_hi = (nsubpom_tonase_hi/nsubmandays_hi) * 1000
			' else
			   ' nsubOutput_hi = 0
			' end if   
			' if nsubmandays_shi > 0 then
			   ' nsubOutput_shi =(nsubpom_tonase_shi /nsubmandays_shi) * 1000
			' else
                ' nsubOutput_shi = 0
            ' end if			
			' if nsubmandays_bi > 0 then
			   ' nsubOutput_bi = (nsubpom_tonase_bi /nsubmandays_bi) * 1000
			' else
			   ' nsubOutput_bi = 0
			' end if
			' if nsubmandays_sbi > 0 then
			   ' nsubOutput_sbi = (nsubpom_tonase_sbi /nsubmandays_sbi) * 1000
			' else
			   ' nsubOutput_sbi = 0
			' end if
			
		    ' if nsubbudmandays_shi > 0 then
				' nsubbudOutput_shi = (nsubbudqty_shi / nsubbudmandays_shi) * 1000
		    ' else
			    ' nsubbudOutput_shi = 0
            ' end if
            ' if nsubbudmandays_sbi > 0 then		
				' nsubbudOutput_sbi = (nsubbudqty_sbi / nsubbudmandays_sbi) * 1000
			' else
			    ' nsubbudOutput_sbi = 0
            ' end if

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
	
	' tambahan akbar
%>
			<tr bgcolor="#F4F4A0">
			   
				<td align="center" colspan="<%= 1+ nSpanTotal %>">TOTAL</td>				
				<td align="right"><%=FormatNumber(nsubhi,2)%></td>
				<td align="right"><%=FormatNumber(nsubshi,2)%></td>
				<td align="right"><%=FormatNumber(nsubsbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubtonase_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubtonase_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubtonase_sbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubpom_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_sbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubbudjjg_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubbudjjg_sbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubpom_tonase_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_sbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubbudqty_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubbudqty_sbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubrestan,2)%></td>

				<td align="right"><%=FormatNumber(nsubmandays_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubmandays_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubmandays_sbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubbudmandays_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubbudmandays_sbi,2)%></td>

				<td align="right"><%=FormatNumber(nsubOutput_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubOutput_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubOutput_sbi,2)%></td>
				
				<td align="right"><%=FormatNumber(nsubbudOutput_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubbudOutput_sbi,2)%></td>

				<td align="right"><%=FormatNumber(nSubMandays_Amount_hi,2)%></td>
				<td align="right"><%=FormatNumber(nSubMandays_Amount_shi,2)%></td>
				
			</tr>
</tbody>
</table>

</form>

</div>

