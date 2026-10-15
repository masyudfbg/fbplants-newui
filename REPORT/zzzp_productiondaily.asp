<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
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
			v_txttodaydate		= Request("txttodaydate")
			
			txtyear 		= right(Request("txttodaydate"),4)
			txtmonth 		= left(Request("txttodaydate"),2)
			txttodaydate 	= Request("txttodaydate")
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 8
			v_txtComp           =   txtComp
			v_txtyear			=	txtyear
			v_txtmonth			=	txtmonth
			
			
			
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
            
			
	strMain = " SELECT   " & strFieldCons & " division, fieldcode, fieldname, plantingyear, " & _
								 " round(SUM(nvl(hi,0)),3) hi, round(SUM(nvl(shi,0)),3) shi, round(SUM(nvl(bi,0)),3) bi, round(SUM(nvl(sbi,0)),3) sbi, " & _
								 " round(SUM(nvl(tonase_hi,0)),3) tonase_hi, round(SUM(nvl(tonase_shi,0)),3) tonase_shi, " & _
								 " round(SUM(nvl(tonase_bi,0)),3) tonase_bi, round(SUM(nvl(tonase_sbi,0)),3) tonase_sbi, " & _
								 " round(SUM(nvl(pom_hi,0)),3) pom_hi, round(SUM(nvl(pom_shi,0)),3) pom_shi, " & _
								 " round(SUM(nvl(pom_bi,0)),3) pom_bi, round(SUM(nvl(pom_sbi,0)),3) pom_sbi, " & _
								 " round(SUM(nvl(pom_tonase_hi,0)),3) pom_tonase_hi, " & _
								 " round(SUM(nvl(pom_tonase_shi,0)),3) pom_tonase_shi, " & _
								 " round(SUM(nvl(pom_tonase_bi,0)),3) pom_tonase_bi, " & _
								 " nvl(budget_transaction_prodqty_api.GetJanjang_SHI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_jjg_shi, " & _
                                 " nvl(budget_transaction_prodqty_api.GetJanJang_SBI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_jjg_sbi, " & _       
								 " nvl(budget_transaction_prodqty_api.GetQuantity_SHI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_qty_shi, " & _
								 " nvl(budget_transaction_prodqty_api.GetQuantity_SBI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_qty_sbi, " & _
								 " nvl(budget_transaction_prodqty_api.GetMandays_SHI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_Mandays_shi, " & _
								 " nvl(budget_transaction_prodqty_api.GetMandays_SBI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_Mandays_sbi, " & _
								 " nvl(budget_transaction_prodqty_api.GetOutput_SHI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_output_shi, " & _
								 " nvl(budget_transaction_prodqty_api.GetOutput_SBI(fcba,'" & v_txtyear & "',1,1,fieldcode,division,to_number(plantingyear),TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY')),0) bud_output_sbi, " & _
								 " round(SUM(nvl(pom_tonase_sbi,0)),3) pom_tonase_sbi, round(SUM(nvl(restan,0)),3) restan, " & _
								 " round(SUM(nvl(mandays_hi,0)),3) mandays_hi, round(SUM(nvl(mandays_shi,0)),3) mandays_shi, round(SUM(nvl(mandays_bi,0)),3) mandays_bi, round(SUM(nvl(mandays_sbi,0)),3) mandays_sbi, " & _
								 " round(SUM(nvl(mandays_amount_hi,0)),3) mandays_amount_hi, round(SUM(nvl(mandays_amount_shi,0)),3) mandays_amount_shi " & _
								 " FROM (SELECT fcba, division, fieldcode, fieldname, plantingyear, " & _
										 " totaldailyproduction hi, 0 shi, 0 bi, 0 sbi, " & _
										 " totaldailyprodtonase tonase_hi, 0 tonase_shi, " & _
										 " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										 " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										 " 0 pom_tonase_shi, 0 pom_tonase_bi, 0 pom_tonase_sbi, " & _
										 " 0 restan, " & _
										 " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _										 
                                         " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									" FROM view_dailyharvesting_prod " & _
								   " WHERE " & _
									 " accountingyear = '" & v_txtyear & "' " & _
									 " AND accountingperiod = '" & v_txtmonth & "' " & _
									 " AND fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, SUM (totaldailyproduction) shi, 0 bi, 0 sbi, " & _
										   " 0 tonase_hi, SUM (totaldailyprodtonase) tonase_shi, " & _
										   " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										   " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, 0 restan , " & _
										   " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _										   
                                           " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									  " FROM view_dailyharvesting_prod " & _
									 " WHERE  " & _
									   " accountingyear = '" & v_txtyear & "' " & _
									   " AND accountingperiod = '" & v_txtmonth & "' " & _
									   " AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, " & _
										   " accountingyear, " & _
										   " accountingperiod, " & _
										   " division, " & _
										   " fieldcode, " & _
										   " fieldname, " & _
										   " plantingyear " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, SUM (totaldailyproduction) sbi, " & _
										   " 0 tonase_hi, 0 tonase_shi, 0 tonase_bi, " & _
										   " SUM (totaldailyprodtonase) tonase_sbi, 0 pom_hi, " & _
										   " 0 pom_shi, 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, 0 restan, " & _
										   " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _										   
                                           " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									  " FROM view_dailyharvesting_prod " & _
									 " WHERE  " & _
									   " accountingyear = '" & v_txtyear & "' " & _
									   " AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, division, fieldcode, fieldname, plantingyear " & _
								  " UNION " & _
								  " SELECT fcba, division, fieldcode, fieldname, plantingyear, " & _
										 " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, 0 tonase_shi, " & _
										 " 0 tonase_bi, 0 tonase_sbi, totaldailyreceive pom_hi, " & _
										 " 0 pom_shi, 0 pom_bi, 0 pom_sbi, " & _
										 " totaldailyreceivetonase1 pom_tonase_hi, " & _
										 " 0 pom_tonase_shi, 0 pom_tonase_bi, 0 pom_tonase_sbi, " & _
										 " 0 restan, " & _
										 " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _										 
                                         " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									" FROM view_dailyharvesting_mill " & _
								   " WHERE  " & _
									 " accountingyear = '" & v_txtyear & "' " & _
									 " AND accountingperiod = '" & v_txtmonth & "' " & _
									 " AND receptiondate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, " & _
										   " 0 tonase_shi, 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, " & _
										   " SUM (totaldailyreceive) pom_shi, 0 pom_bi, " & _
										   " 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " SUM (totaldailyreceivetonase1) pom_tonase_shi, " & _
										   " 0 pom_tonase_bi, 0 pom_tonase_sbi, 0 restan, " & _
										   " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _										   
                                           " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									  " FROM view_dailyharvesting_mill " & _
									 " WHERE " & _
									   " accountingyear = '" & v_txtyear & "' " & _
									   " AND accountingperiod = '" & v_txtmonth & "' " & _
									   " AND receptiondate <=  TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, " & _
										   " accountingyear, " & _
										   " accountingperiod, " & _
										   " division, " & _
										   " fieldcode, " & _
										   " fieldname, " & _
										   " plantingyear " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, " & _
										   " 0 tonase_shi, 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, " & _
										   " 0 pom_shi, 0 pom_bi, " & _
										   " SUM (totaldailyreceive) pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " SUM (totaldailyreceivetonase1) pom_tonase_sbi, " & _
										   " 0 restan, " & _
										   " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _										   
                                           " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									  " FROM view_dailyharvesting_mill " & _
									 " WHERE " & _
									   " accountingyear = '" & v_txtyear & "' " & _
									   " AND receptiondate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, division, fieldcode, fieldname, plantingyear " & _
								  " UNION " & _
								  " select FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR,  " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, " & _
										   " 0 tonase_shi, 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, " & _
										   " 0 pom_shi, 0 pom_bi, " & _
										   " 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, " & _
										   " sum(variantemp) + sum(restantemp) restan , " & _
										   " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _
                                           " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
										" from (	 " & _
											" select FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR,  varian variantemp,  " & _
											" 0 restantemp from view_dailyharvesting_varian " & _
											" where " & _
											" DATEPROD = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
											" UNION " & _
											" select FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR,  0 variantemp,  " & _
											" sum(varian) restantemp from view_dailyharvesting_varian " & _
											" where " & _
											" DATEPROD < TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
											" group by FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR " & _
										" )  " & _
										" group by FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR " & _
									" UNION " & _
									" SELECT fcba, division, fieldcode, fieldname, plantingyear, " & _
										 " 0 hi, 0 shi, 0 bi, 0 sbi, " & _
										 " 0 tonase_hi, 0 tonase_shi, " & _
										 " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										 " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										 " 0 pom_tonase_shi, 0 pom_tonase_bi, 0 pom_tonase_sbi, " & _
										 " 0 restan, " & _
										 " totaldailymandays mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _
                                         " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									" FROM view_dailyharvesting_mandays " & _
								   " WHERE " & _
									 " accountingyear = '" & v_txtyear & "' " & _
									 " AND accountingperiod = '" & v_txtmonth & "' " & _
									 " AND fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, " & _
										   " 0 tonase_hi, 0 tonase_shi, " & _
										   " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										   " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, 0 restan, " & _
										   " 0 mandays_hi, SUM(totaldailymandays) mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _
                                           " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									  " FROM view_dailyharvesting_mandays " & _
									 " WHERE  " & _
									   " accountingyear = '" & v_txtyear & "' " & _
									   " AND accountingperiod = '" & v_txtmonth & "' " & _
									   " AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, " & _
										   " accountingyear, " & _
										   " accountingperiod, " & _
										   " division, " & _
										   " fieldcode, " & _
										   " fieldname, " & _
										   " plantingyear " & _
								  " UNION " & _
									" SELECT fcba, division, fieldcode, fieldname, plantingyear, " & _
										 " 0 hi, 0 shi, 0 bi, 0 sbi, " & _
										 " 0 tonase_hi, 0 tonase_shi, " & _
										 " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										 " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										 " 0 pom_tonase_shi, 0 pom_tonase_bi, 0 pom_tonase_sbi, " & _
										 " 0 restan, " & _
										 " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _	
                                         " mandays_amount mandays_amount_hi, 0 mandays_amount_shi " & _					 
									" FROM view_dailyharvesting_amount " & _
								   " WHERE " & _
									 " accountingyear = '" & v_txtyear & "' " & _
									 " AND accountingperiod = '" & v_txtmonth & "' " & _
									 " AND fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, " & _
										   " 0 tonase_hi, 0 tonase_shi, " & _
										   " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										   " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, 0 restan , " & _
										   " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, 0 mandays_sbi, " & _										   
                                           " 0 mandays_amount_hi, sum(mandays_amount) mandays_amount_shi " & _										 
									  " FROM view_dailyharvesting_amount " & _
									 " WHERE  " & _
									   " accountingyear = '" & v_txtyear & "' " & _
									   " AND accountingperiod = '" & v_txtmonth & "' " & _
									   " AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, " & _
										   " accountingyear, " & _
										   " accountingperiod, " & _
										   " division, " & _
										   " fieldcode, " & _
										   " fieldname, " & _
										   " plantingyear " & _										   
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, " & _
										   " 0 tonase_hi, 0 tonase_shi, 0 tonase_bi, " & _
										   " 0 tonase_sbi, 0 pom_hi, " & _
										   " 0 pom_shi, 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, 0 restan, " & _
										   " 0 mandays_hi, 0 mandays_shi, 0 mandays_bi, SUM(totaldailymandays) mandays_sbi, " & _
                                           " 0 mandays_amount_hi, 0 mandays_amount_shi " & _										 
									  " FROM view_dailyharvesting_mandays " & _
									 " WHERE  " & _
									   " accountingyear = '" & v_txtyear & "' " & _
									   " AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, division, fieldcode, fieldname, plantingyear " & _									  
								  " ) " & _
								 
						" GROUP BY BUSINESSUNIT_API.get_OrderData(fcba),fcba, division, fieldcode, fieldname, plantingyear "
	
	' strEPS = " SELECT fcba"
	' ' if len(trim(v_txtcompany )) > 0 then 
		' ' if v_txtcompany  <> "--SUMMARY--" then
			' ' strEPS = strEPS & "  company"
		' ' end if
	' ' else
		' ' strEPS = strEPS & "  company "
	' ' end if
	' ' if len(trim(v_txtfcba)) > 0 then 
		' ' if v_txtfcba <> "--SUMMARY--" then
			' ' strEPS = strEPS & "  fcba"
		' ' end if
	' ' else
		' ' strEPS = strEPS & "  fcba "
	' ' end if
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
	
	
	' strEPS = strEPS &            ",SUM (hi) hi, SUM (shi) shi, SUM (bi) bi, SUM (sbi) sbi, " & _
								 ' " SUM (tonase_hi) tonase_hi, SUM (tonase_shi) tonase_shi, " & _
								 ' " SUM (tonase_bi) tonase_bi, SUM (tonase_sbi) tonase_sbi, " & _
								 ' " SUM (pom_hi) pom_hi, SUM (pom_shi) pom_shi, " & _
								 ' " SUM (pom_bi) pom_bi, SUM (pom_sbi) pom_sbi, " & _
								 ' " SUM (pom_tonase_hi) pom_tonase_hi, " & _
								 ' " SUM (pom_tonase_shi) pom_tonase_shi, " & _
								 ' " SUM (pom_tonase_bi) pom_tonase_bi, " & _
								 ' " SUM (pom_tonase_sbi) pom_tonase_sbi, SUM (restan) restan ,  " & _
								 ' " SUM (mandays_hi) mandays_hi, SUM(mandays_shi) mandays_shi, SUM(mandays_bi) mandays_bi, SUM(mandays_sbi) mandays_sbi, " & _
								 ' " SUM (mandays_amount_hi) mandays_amount_hi, SUM(mandays_amount_shi) mandays_amount_shi "
  
   	' strEPS2 =                     ",SUM (hi) hi, SUM (shi) shi, SUM (bi) bi, SUM (sbi) sbi,sum(bud_jjg_shi) bud_jjg_shi,sum(bud_jjg_sbi) bud_jjg_sbi, " & _
								 ' " SUM (tonase_hi) tonase_hi, SUM (tonase_shi) tonase_shi, " & _
								 ' " SUM (tonase_bi) tonase_bi, SUM (tonase_sbi) tonase_sbi, " & _
								 ' " SUM (pom_hi) pom_hi, SUM (pom_shi) pom_shi, " & _
								 ' " SUM (pom_bi) pom_bi, SUM (pom_sbi) pom_sbi, " & _
								 ' " SUM (pom_tonase_hi) pom_tonase_hi, " & _
								 ' " SUM (pom_tonase_shi) pom_tonase_shi, " & _
								 ' " SUM (pom_tonase_bi) pom_tonase_bi, " & _
								 ' " SUM (pom_tonase_sbi) pom_tonase_sbi, SUM (restan) restan ,  " & _
								 ' " SUM (mandays_hi) mandays_hi, SUM(mandays_shi) mandays_shi, SUM(mandays_bi) mandays_bi, SUM(mandays_sbi) mandays_sbi, " & _
								 ' " SUM (mandays_amount_hi) mandays_amount_hi, SUM(mandays_amount_shi) mandays_amount_shi 
								 
	 strEPS2  =  " ,SUM(hi) hi, " & _
				 "  SUM(shi) shi, " & _
				 "  SUM(sbi) sbi, " & _
				 "  SUM(tonase_hi) tonase_hi, " & _
				 "  SUM(tonase_shi) tonase_shi, " & _
				 "  SUM(tonase_sbi) tonase_sbi,  " & _
				 "  SUM(pom_hi) pom_hi, " & _
				 "  SUM(pom_shi) pom_shi, " & _
				 "  SUM(pom_sbi) pom_sbi, " & _
				 "  sum(bud_jjg_shi) bud_jjg_shi,sum(bud_jjg_sbi) bud_jjg_sbi, " & _
				 "  SUM(pom_tonase_hi) pom_tonase_hi, " & _
				 "  SUM(pom_tonase_shi) pom_tonase_shi, " & _
				 "  SUM(pom_tonase_sbi) pom_tonase_sbi, " & _
				 "  sum(bud_qty_shi)/1000 bud_qty_shi,sum(bud_qty_sbi)/1000 bud_qty_sbi, " & _
				 "  SUM(restan) restan, " & _
				 "  SUM(mandays_hi) mandays_hi, " & _
				 "  SUM(mandays_shi) mandays_shi, " & _
				 "  SUM(mandays_sbi) mandays_sbi, " & _
				 "  sum(bud_mandays_shi) bud_mandays_shi,sum(bud_mandays_sbi) bud_mandays_sbi, " & _
				 "  SUM(mandays_amount_hi) mandays_amount_hi, " & _
				 "  SUM(mandays_amount_shi) mandays_amount_shi,	" & _	
                 "  sum(bud_output_shi) bud_output_shi,sum(bud_output_sbi) bud_output_sbi " 				 
								 
								 

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
<link rel="stylesheet" href="/css/stylereport(ams).css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<!--<link rel="stylesheet" href="/css/test.css" type="text/css" />
TABLE STYLE -->
 		<link href="css/960.css" rel="stylesheet" media="screen" />
        <link href="css/defaultTheme.css" rel="stylesheet" media="screen" />
        <link href="css/myTheme.css" rel="stylesheet" media="screen" />
        <style>
		.headrowspan {	height:127px;}
		.leftwidth{}
		table {
			border-collapse: collapse;
			border-spacing: 0;
		}
		</style>
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.5.2/jquery.min.js"></script>

        <script src="js/jquery.fixedheadertable.js"></script>
       <!-- <script src="js/fix.js"></script>
        
<!-- END TABLE STYLE -->
<!--<script src="/js/jquery-1.9.1.js"></script>
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
</script>-->
<script src="/js/jquery-ui.js"></script>
 <script>
$(function() {
$( "#datepicker" ).datepicker();
});
</script>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
 <script>
            function toggle(id){
				var e = document.getElementById(id);
				if(e.style.display == 'block'){
					e.style.display = 'none';
				}else if(e.style.display == 'none'){
					e.style.display = 'block';
					
				}
			}
            </script>
            
     <script>
	$(document).ready(function() {

		$('#runFix').bind('click', function(e) {
		e.preventDefault();
		$('#product-table').fixedHeaderTable({ altClass: 'odd', footer: true });
		});

		$('#runPlugin').bind('click', function(e) {
			e.preventDefault();
			$('#product-table').fixedHeaderTable({ altClass: 'odd', footer: true ,fixedColumns: <%= 1+ nSpanTotal %>  });
			
		});
	});
				
	</script>
</head>

<body>

<form name="frmEntry" method="post" action="ZZZp_ProductionDaily.asp">


	
<div class="setting"> 
<div id="left_form" style=" display:none;">	

<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>TAHUN</th>
		<th>BULAN</th>
		<th>TANGGAL</th>
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
		<td align="center"><input type="text" name="txtyear" value="<%=v_txtyear%>"></td>
		<td align="center"><input type="text" name="txtmonth" value="<%=v_txtmonth%>"></td>
		<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
         <input type="text" id="datepicker" name="txttodaydate" value="<%=v_txttodaydate%>" />
			<!--<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >-->
		<% else%>
          <input type="text" id="datepicker" name="txttodaydate" value="<%=textDate(Date)%>" />
			<!--<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >-->
		<% end if%>
			<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<!--<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>-->
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
</table>

</div>
<a id="tog"  onclick="javascript:toggle('left_form','tog')" href="#">SETTING</a>
</div>
<br/><br/><br/>
<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
        <tr>
			<td colspan="2" align="center"><input type="submit" id="runPlugin"  value="Run Fix Head table" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" id="runFix" value="freze left colom" disabled></td>
			
		</tr>
	</table>
<!--<div id="wrapper">
<tr>
<th colspan="38" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">-->
<div class="container_12" style="width:97%; padding-bottom:100px;"> 
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
			<h2>LAPORAN HARIAN PRODUKSI</h2>
			<% if v_txttodaydate <> "" and v_txtFcba <> "" then %>
				<h5>PERIODE : <%=v_txttodaydate%>, KEBUN : <%=v_txtFcba%></h5>
			<% end if %>

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
<!--
</th>

</tr>-->



<div class="grid_4 height400" style="width:100%;">
<table border="0" width="100%" id="product-table" class="fancyTable" >
<thead>
	<tr>
		 <th rowspan="4" style=" background-color:#00A300;color:#fff; "></th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% 			
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% 			
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;">REGION</th>
		<%
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="4" style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% 		
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="4" style="background-color:#00A300;color:#fff;">SITE</th>
		<% 		
	    end if %>

		
		<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% 		
	    end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" And v_txtPlantingDate <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">TT</th>
		<% 		
	    end if %>


		<% if v_txtBlok <> "--SUMMARY--" And v_txtBlok <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">BLOK</th>
		<% 		
	    end if %>
		
		<th colspan="6" style="background-color:#00A300;color:#fff;">PRODUKSI TBS KEBUN</th>
		<th colspan="10" style="background-color:#00A300;color:#fff;">TBS DITERIMA PMKS</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">RESTAN (JJG)</th>
		<th colspan="10" style="background-color:#00A300;color:#fff;">TENAGA KERJA POTONG BUAH</th>
		<th colspan="2" rowspan="3" style="background-color:#00A300;color:#fff;">AMOUNT (RP/KG)</th>		
	</tr>
	
		<th colspan="3" style="background-color:#00A300;color:#fff;">JANJANG PANEN</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">TONASE</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;">JANJANG KIRIM</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;">TONASE</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;">HARI KERJA (HK)</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;">OUTPUT (KG/HK)</th>
	</tr>
	
	
	<tr>
  
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
	</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
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
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
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
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtDivision <> "--SUMMARY--" then %>
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
					   tempDivision = rsEPS("division")
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevDivision <> rsEPS("division") then %>
					    <td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("division")%>
							</a>
						</td>
						<% sprevDivision = rsEPS("division") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>

				<% if v_txtPlantingDate <> "--SUMMARY--" then %>
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
					   tempDivision = rsEPS("division")
					   tempPlantingDate = rsEPS("plantingyear")
					   tempBlok = "--ALL--"					   
					 %>
					
					<% tempPlantingDate = rsEPS("plantingyear") %>
					
					
				    <% if sprevPlantingYear <> rsEPS("plantingyear") then %>
				        <td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("plantingyear")%>
							</a>
						</td>
						<% sprevPlantingYear = rsEPS("plantingyear") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>


				<% if v_txtBlok <> "--SUMMARY--" then %>
				    <% if sprevFieldCode <> rsEPS("fieldcode") then %>				
					    <td align="center"><%=rsEPS("fieldcode")%></td>				
						<% sprevFieldCode = rsEPS("fieldcode") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>				
				
				
				<td align="right"><%=FormatNumber(rsEPS("hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("sbi"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("tonase_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("tonase_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("tonase_sbi"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("pom_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_sbi"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("bud_jjg_shi"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("bud_jjg_sbi"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("pom_tonase_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_tonase_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_tonase_sbi"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("bud_qty_shi"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("bud_qty_sbi"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("restan"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("mandays_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("mandays_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("mandays_sbi"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("bud_mandays_shi"),2)%></td>
                <td align="right"><%=FormatNumber(rsEPS("bud_mandays_sbi"),2)%></td>
				
				<% if cdbl(rsEPS("mandays_hi")) <> 0 then %>
				    <% nOutput_hi = cdbl(rsEPS("pom_tonase_hi")) * 1000 / cdbl(rsEPS("mandays_hi")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("mandays_shi")) <> 0 then %>
					<% nOutput_shi = cdbl(rsEPS("pom_tonase_shi")) * 1000 / cdbl(rsEPS("mandays_shi")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if cdbl(rsEPS("mandays_sbi")) <> 0 then %>
					<% nOutput_sbi = cdbl(rsEPS("pom_tonase_sbi")) * 1000 / cdbl(rsEPS("mandays_sbi")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,2)%></td>
				
				<% if cdbl(rsEPS("bud_mandays_shi")) <> 0 then %>
					<% nbudOutput_shi = (cdbl(rsEPS("bud_qty_shi")) * 1000 ) /  cdbl(rsEPS("bud_mandays_shi")) %>
				<% else %>
				    <% nbudOutput_shi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("bud_mandays_sbi")) <> 0 then %>
					<% nbudOutput_sbi = (cdbl(rsEPS("bud_qty_sbi")) * 1000 )  /  cdbl(rsEPS("bud_mandays_sbi")) %>
				<% else %>
				    <% nbudOutput_sbi = 0 %>
				<% end if %>
				
				<td align="right"><%=FormatNumber(nbudOutput_shi,2)%></td>
                <td align="right"><%=FormatNumber(nbudOutput_sbi,2)%></td>
				
				<% if cdbl(rsEPS("pom_tonase_hi")) <> 0 then %>
					<% nOutput_Mandays_Amount_hi = cdbl(rsEPS("mandays_amount_hi")) / (cdbl(rsEPS("pom_tonase_hi")) * 1000)  %>
				<% else %>
				    <% nOutput_Mandays_Amount_hi = 0 %>
				<% end if %>

				<% if cdbl(rsEPS("pom_tonase_shi")) <> 0 then %>
					<% nOutput_Mandays_Amount_shi = cdbl(rsEPS("mandays_amount_shi")) / (cdbl(rsEPS("pom_tonase_shi")) * 1000)  %>
				<% else %>
				    <% nOutput_Mandays_Amount_shi = 0 %>
				<% end if %>
				
				
				
				<td align="right"><%=FormatNumber(nOutput_Mandays_Amount_hi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_Mandays_Amount_shi,2)%></td>
				
				
				
			</tr>
<%
			NO = NO + 1
			nsubhi = nsubhi + Cdbl(rsEPS("hi"))
			nsubshi = nsubshi + Cdbl(rsEPS("shi"))			
			nsubsbi = nsubsbi + Cdbl(rsEPS("sbi"))
			
			nsubtonase_hi = nsubtonase_hi + Cdbl(rsEPS("tonase_hi"))
			nsubtonase_shi = nsubtonase_shi + Cdbl(rsEPS("tonase_shi"))			
			nsubtonase_sbi = nsubtonase_sbi + Cdbl(rsEPS("tonase_sbi"))
			
			nsubpom_hi = nsubpom_hi + Cdbl(rsEPS("pom_hi"))
			nsubpom_shi = nsubpom_shi + Cdbl(rsEPS("pom_shi"))
			nsubpom_sbi = nsubpom_sbi + Cdbl(rsEPS("pom_sbi"))
			
			nsubbudjjg_shi = nsubbudjjg_shi + Cdbl(rsEPS("bud_jjg_shi"))
			nsubbudjjg_sbi = nsubbudjjg_sbi + Cdbl(rsEPS("bud_jjg_sbi"))
			
			nsubpom_tonase_hi = nsubpom_tonase_hi + Cdbl(rsEPS("pom_tonase_hi"))
			nsubpom_tonase_shi = nsubpom_tonase_shi + Cdbl(rsEPS("pom_tonase_shi"))
			nsubpom_tonase_sbi = nsubpom_tonase_sbi + Cdbl(rsEPS("pom_tonase_sbi"))
			
			nsubbudqty_shi = nsubbudqty_shi + Cdbl(rsEPS("bud_qty_shi"))
			nsubbudqty_sbi = nsubbudqty_sbi + Cdbl(rsEPS("bud_qty_sbi"))
			
			' tambahan akbar
			nsubmandays_amount_hi = nsubmandays_amount_hi + Cdbl(rsEPS("mandays_amount_hi"))
			nsubmandays_amount_shi = nsubmandays_amount_shi + Cdbl(rsEPS("mandays_amount_shi"))
			' tambahan akbar
			
			
			nsubbudmandays_shi = nsubbudmandays_shi + Cdbl(rsEPS("bud_mandays_shi"))
			nsubbudmandays_sbi = nsubbudmandays_sbi + Cdbl(rsEPS("bud_mandays_sbi"))
			
			nsubrestan = nsubrestan + Cdbl(rsEPS("restan"))
			
			nsubmandays_hi = nsubmandays_hi + Cdbl(rsEPS("mandays_hi"))
			nsubmandays_shi = nsubmandays_shi + Cdbl(rsEPS("mandays_shi"))
			
			nsubmandays_sbi = nsubmandays_sbi + Cdbl(rsEPS("mandays_sbi"))		

			if nsubmandays_hi > 0 then
			   nsubOutput_hi = (nsubpom_tonase_hi/nsubmandays_hi) * 1000
			else
			   nsubOutput_hi = 0
			end if   
			if nsubmandays_shi > 0 then
			   nsubOutput_shi =(nsubpom_tonase_shi /nsubmandays_shi) * 1000
			else
                nsubOutput_shi = 0
            end if			
			if nsubmandays_bi > 0 then
			   nsubOutput_bi = (nsubpom_tonase_bi /nsubmandays_bi) * 1000
			else
			   nsubOutput_bi = 0
			end if
			if nsubmandays_sbi > 0 then
			   nsubOutput_sbi = (nsubpom_tonase_sbi /nsubmandays_sbi) * 1000
			else
			   nsubOutput_sbi = 0
			end if
			
		    if nsubbudmandays_shi > 0 then
				nsubbudOutput_shi = (nsubbudqty_shi / nsubbudmandays_shi) * 1000
		    else
			    nsubbudOutput_shi = 0
            end if
            if nsubbudmandays_sbi > 0 then		
				nsubbudOutput_sbi = (nsubbudqty_sbi / nsubbudmandays_sbi) * 1000
			else
			    nsubbudOutput_sbi = 0
            end if

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
</tbody>
<tfoot>
			<tr bgcolor="#F4F4A0">
			   
				<td align="center" colspan="<%= 1+ nSpanTotal %>">TOTAL</td>
                <!--<td align="right"></td>	
                <td align="right"></td>
                
                <td align="right"></td>	
                <td align="right"></td>	
                <td align="right">TOTAL</td>-->	
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
</tfoot>
</table>
</div>
</form>
<div class="clear">
</body>
</html>

