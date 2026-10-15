 <% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDateFrom, v_txtDateTo
Dim ScriptName
Dim rsAnalisaAngkutTBS, strAnalisaAngkutTBS, strAnalisaAngkutTBS2, strMain
Dim cClass, nday

Dim var_colspan : var_colspan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtOwn				= Request("txtOwn")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where company is not null "

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

v_txtDateFrom	=	Request("txtDateFrom")
v_txtDateTo		=	Request("txtDateTo")


If Len(Trim(v_txtDateFrom)) = 0 Then
	v_txtDateFrom = textDate(Date)
End If

If Len(Trim(v_txtDateTo)) = 0 Then
	v_txtDateTo = textDate(Date)
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT " & strFieldCons & "" & _
			"					,transporttype, " & _
			"					 own, " & _
			"                    SUM (kg_bunch_own_direct) kg_bunch_own_direct," & _
			"                    SUM (kg_bunch_own_indirect) kg_bunch_own_indirect," & _
			"                    SUM (Kg_Bunch_contract_direct) Kg_Bunch_contract_direct," & _
			"                    SUM (Kg_Bunch_contract_indirect) Kg_Bunch_contract_indirect," & _
			"                    SUM (kg_bunch_own_direct + kg_bunch_own_indirect)  Totalown, " & _
			"                    SUM (Kg_Bunch_contract_direct + Kg_Bunch_contract_indirect) Totalcontract" & _ 
			"               FROM (SELECT bunchbucket.fcba," & _
			"                            bunchbucket.spbno," & _
			"                            bunchbucket.own," & _
			"                            bunchbucket.transporttype," & _
			"                            bunchbucket.receptiondate," & _
			"                            bunchbucket.vehicle," & _
			"                            bunchbucket.driver," & _
			"                            bunchbucket.division," & _
			"                            bunchbucket.total_bunch," & _
			"                            iplasprod.wb_harvesting.bjr_per_chitno," & _
			"                            (  bunchbucket.total_bunch" & _
			"                             * iplasprod.wb_harvesting.bjr_per_chitno)" & _
			"                               AS kg_bunch_own_direct," & _
			"                            0 kg_bunch_own_indirect," & _
			"                            0 Kg_Bunch_contract_direct," & _
			"                            0 Kg_Bunch_contract_indirect," & _
			"                            bunchbucket.total_bucket," & _
			"                            iplasprod.wb_harvesting.kg_per_bucket," & _
			"                            (  bunchbucket.total_bucket" & _
			"                             * iplasprod.wb_harvesting.kg_per_bucket)" & _
			"                               AS Kg_Bucket" & _
			"                       FROM (  SELECT fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division," & _
			"                                      SUM (total_bunch) AS total_bunch," & _
			"                                      SUM (total_bucket) AS total_bucket" & _
			"                                 FROM (  SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                NVL (" & _
			"                                                   SUM (iplasprod.harvestingspb.bunch)," & _
			"                                                   0)" & _
			"                                                   AS total_bunch," & _
			"                                                0 AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bunch > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division" & _
			"                                       UNION ALL" & _
			"                                         SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                0 AS total_bunch," & _
			"                                                NVL (" & _
			"                                                   SUM (" & _
			"                                                      iplasprod.harvestingspb.bucket)," & _
			"                                                   0)" & _
			"                                                   AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bucket > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division)" & _
			"                                WHERE     receptiondate >= TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                                      AND receptiondate <= TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                                      AND own = 'OWN'" & _
			"                                      AND transporttype = 'DIRECTTRANSPORT'" & _
			"                             GROUP BY fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division) bunchbucket" & _
			"                            INNER JOIN iplasprod.wb_harvesting" & _
			"                               ON (    bunchbucket.fcba =" & _
			"                                          iplasprod.wb_harvesting.fcba_estate" & _
			"                                   AND bunchbucket.spbno =" & _
			"                                          iplasprod.wb_harvesting.spbno" & _
			"                                   AND bunchbucket.vehicle =" & _
			"                                          iplasprod.wb_harvesting.vehiclecode)" & _
			"                     UNION ALL" & _
			"                     SELECT bunchbucket.fcba," & _
			"                            bunchbucket.spbno," & _
			"                            bunchbucket.own," & _
			"                            bunchbucket.transporttype," & _
			"                            bunchbucket.receptiondate," & _
			"                            bunchbucket.vehicle," & _
			"                            bunchbucket.driver," & _
			"                            bunchbucket.division," & _
			"                            bunchbucket.total_bunch," & _
			"                            iplasprod.wb_harvesting.bjr_per_chitno," & _
			"                            0 Kg_Bunch_own_direct," & _
			"                            (  bunchbucket.total_bunch" & _
			"                             * iplasprod.wb_harvesting.bjr_per_chitno)" & _
			"                               AS Kg_Bunch_own_indirect," & _
			"                            0 Kg_Bunch_contract_direct," & _
			"                            0 Kg_Bunch_contract_indirect," & _
			"                            bunchbucket.total_bucket," & _
			"                            iplasprod.wb_harvesting.kg_per_bucket," & _
			"                            (  bunchbucket.total_bucket" & _
			"                             * iplasprod.wb_harvesting.kg_per_bucket)" & _
			"                               AS Kg_Bucket" & _
			"                       FROM (  SELECT fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division," & _
			"                                      SUM (total_bunch) AS total_bunch," & _
			"                                      SUM (total_bucket) AS total_bucket" & _
			"                                 FROM (  SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                NVL (" & _
			"                                                   SUM (iplasprod.harvestingspb.bunch)," & _
			"                                                   0)" & _
			"                                                   AS total_bunch," & _
			"                                                0 AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bunch > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division" & _
			"                                       UNION ALL" & _
			"                                         SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                0 AS total_bunch," & _
			"                                                NVL (" & _
			"                                                   SUM (" & _
			"                                                      iplasprod.harvestingspb.bucket)," & _
			"                                                   0)" & _
			"                                                   AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bucket > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division)" & _
			"                                WHERE     receptiondate >= TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                                      AND receptiondate <= TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                                      AND own = 'OWN'" & _
			"                                      AND transporttype = 'INDIRECTTRANSPORT'" & _
			"                             GROUP BY fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division) bunchbucket" & _
			"                            INNER JOIN iplasprod.wb_harvesting" & _
			"                               ON (    bunchbucket.fcba =" & _
			"                                          iplasprod.wb_harvesting.fcba_estate" & _
			"                                   AND bunchbucket.spbno =" & _
			"                                          iplasprod.wb_harvesting.spbno" & _
			"                                   AND bunchbucket.vehicle =" & _
			"                                          iplasprod.wb_harvesting.vehiclecode)" & _
			"                     UNION ALL" & _
			"                     SELECT bunchbucket.fcba," & _
			"                            bunchbucket.spbno," & _
			"                            bunchbucket.own," & _
			"                            bunchbucket.transporttype," & _
			"                            bunchbucket.receptiondate," & _
			"                            bunchbucket.vehicle," & _
			"                            bunchbucket.driver," & _
			"                            bunchbucket.division," & _
			"                            bunchbucket.total_bunch," & _
			"                            iplasprod.wb_harvesting.bjr_per_chitno," & _
			"                            0 Kg_Bunch_own_direct," & _
			"                            0 Kg_Bunch_own_indirect," & _
			"                            (  bunchbucket.total_bunch" & _
			"                             * iplasprod.wb_harvesting.bjr_per_chitno)" & _
			"                               AS kg_bunch_contract_direct," & _
			"                            0 kg_bunch_contract_indirect," & _
			"                            bunchbucket.total_bucket," & _
			"                            iplasprod.wb_harvesting.kg_per_bucket," & _
			"                            (  bunchbucket.total_bucket" & _
			"                             * iplasprod.wb_harvesting.kg_per_bucket)" & _
			"                               AS Kg_Bucket" & _
			"                       FROM (  SELECT fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division," & _
			"                                      SUM (total_bunch) AS total_bunch," & _
			"                                      SUM (total_bucket) AS total_bucket" & _
			"                                 FROM (  SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                NVL (" & _
			"                                                   SUM (iplasprod.harvestingspb.bunch)," & _
			"                                                   0)" & _
			"                                                   AS total_bunch," & _
			"                                                0 AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bunch > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division" & _
			"                                       UNION ALL" & _
			"                                         SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                0 AS total_bunch," & _
			"                                                NVL (" & _
			"                                                   SUM (" & _
			"                                                      iplasprod.harvestingspb.bucket)," & _
			"                                                   0)" & _
			"                                                   AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bucket > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division)" & _
			"                                WHERE     receptiondate >= TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                                      AND receptiondate <= TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                                      AND own = 'CONTRACT'" & _
			"                                      AND transporttype = 'DIRECTTRANSPORT'" & _
			"                             GROUP BY fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division) bunchbucket" & _
			"                            LEFT JOIN iplasprod.wb_harvesting" & _
			"                               ON (    bunchbucket.fcba =" & _
			"                                          iplasprod.wb_harvesting.fcba_estate" & _
			"                                   AND bunchbucket.spbno =" & _
			"                                          iplasprod.wb_harvesting.spbno )" & _
			"                     UNION ALL" & _
			"                     SELECT bunchbucket.fcba," & _
			"                            bunchbucket.spbno," & _
			"                            bunchbucket.own," & _
			"                            bunchbucket.transporttype," & _
			"                            bunchbucket.receptiondate," & _
			"                            bunchbucket.vehicle," & _
			"                            bunchbucket.driver," & _
			"                            bunchbucket.division," & _
			"                            bunchbucket.total_bunch," & _
			"                            iplasprod.wb_harvesting.bjr_per_chitno," & _
			"                            0 Kg_Bunch_own_direct," & _
			"                            0 Kg_Bunch_own_indirect," & _
			"                            0 Kg_Bunch_contract_direct," & _
			"                            (  bunchbucket.total_bunch" & _
			"                             * iplasprod.wb_harvesting.bjr_per_chitno)" & _
			"                               AS Kg_Bunch_contract_indirect," & _
			"                            bunchbucket.total_bucket," & _
			"                            iplasprod.wb_harvesting.kg_per_bucket," & _
			"                            (  bunchbucket.total_bucket" & _
			"                             * iplasprod.wb_harvesting.kg_per_bucket)" & _
			"                               AS Kg_Bucket" & _
			"                       FROM (  SELECT fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division," & _
			"                                      SUM (total_bunch) AS total_bunch," & _
			"                                      SUM (total_bucket) AS total_bucket" & _
			"                                 FROM (  SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                NVL (" & _
			"                                                   SUM (iplasprod.harvestingspb.bunch)," & _
			"                                                   0)" & _
			"                                                   AS total_bunch," & _
			"                                                0 AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bunch > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division" & _
			"                                       UNION ALL" & _
			"                                         SELECT iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division," & _
			"                                                0 AS total_bunch," & _
			"                                                NVL (" & _
			"                                                   SUM (" & _
			"                                                      iplasprod.harvestingspb.bucket)," & _
			"                                                   0)" & _
			"                                                   AS total_bucket" & _
			"                                           FROM iplasprod.harvestingspb" & _
			"                                                INNER JOIN iplasprod.field" & _
			"                                                   ON (    iplasprod.harvestingspb.fcba =" & _
			"                                                              iplasprod.field.fcba" & _
			"                                                       AND iplasprod.harvestingspb.fieldcode =" & _
			"                                                              iplasprod.field.fccode)" & _
			"                                          WHERE iplasprod.harvestingspb.bucket > 0" & _
			"                                       GROUP BY iplasprod.harvestingspb.fcba," & _
			"                                                iplasprod.harvestingspb.spbno," & _
			"                                                iplasprod.harvestingspb.own," & _
			"                                                iplasprod.harvestingspb.transporttype," & _
			"                                                iplasprod.harvestingspb.receptiondate," & _
			"                                                iplasprod.harvestingspb.vehicle," & _
			"                                                iplasprod.harvestingspb.driver," & _
			"                                                iplasprod.field.division)" & _
			"                                WHERE     receptiondate >= TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                                      AND receptiondate <= TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                                      AND own = 'CONTRACT'" & _
			"                                      AND transporttype = 'INDIRECTTRANSPORT'" & _
			"                             GROUP BY fcba," & _
			"                                      spbno," & _
			"                                      own," & _
			"                                      transporttype," & _
			"                                      receptiondate," & _
			"                                      vehicle," & _
			"                                      driver," & _
			"                                      division) bunchbucket" & _
			"                            LEFT JOIN iplasprod.wb_harvesting" & _
			"                               ON (    bunchbucket.fcba =" & _
			"                                          iplasprod.wb_harvesting.fcba_estate" & _
			"                                   AND bunchbucket.spbno =" & _
			"                                          iplasprod.wb_harvesting.spbno ))" & _ 
			" where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E'" & _
			" GROUP BY FCBA, TRANSPORTTYPE, OWN ORDER BY FCBA "
			
strAnalisaAngkutTBS2	=	" 		 , transporttype," & _
							" 		   own," & _
							"          SUM (kg_bunch_own_direct) kg_bunch_own_direct," & _
							"          SUM (kg_bunch_own_indirect) kg_bunch_own_indirect," & _
							"          SUM (Kg_Bunch_contract_direct) Kg_Bunch_contract_direct," & _
							"          SUM (Kg_Bunch_contract_indirect) Kg_Bunch_contract_indirect," & _
							"          SUM (kg_bunch_own_direct + kg_bunch_own_indirect) Totalown," & _
							"          SUM (Kg_Bunch_contract_direct + Kg_Bunch_contract_indirect) Totalcontract "
			
			
strAnalisaAngkutTBS	=	" select GroupPlantation " & selectCons & strAnalisaAngkutTBS2 & "  From (" & strMain & " )" & whereCons & _
						" group by GroupPlantation" & selectCons & ", transporttype, own Order by GroupPlantation" & selectCons
						
' response.write strAnalisaAngkutTBS
%>
<html>

<head>
<title>ePlaS Reporting - Analisa Angkut TBS</title>
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

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSumE.html"-->
				<table id="form-table" class="centered-margin">	
					<tr>
						<td><font face="verdana" size="1">DARI TANGGAL</font></td>
						<td align="center">
						<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
						<% else %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a href="javascript:show_calendar('frmEntry.txtDateFrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						<td><font face="verdana" size="1">SAMPAI</font></td>
						<td align="center">
						<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
						<% else %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a href="javascript:show_calendar('frmEntry.txtDateTo');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						
						<td><input type="submit" value="View" /></td>
					</tr>				
				</table>
			</div>
			</form>
		</div>

<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
        <tr>
			<th colspan="<%=var_colspan+8%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">    
				<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>ANALISA PENGANGKUTAN TBS</h3>
					<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
						<h5>TANGGAL: <%=v_txtdateto%></h5>
					<% else %>
						<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
					<% end if %>
				</div>
			</th>
		</tr>
	<tr>
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
	
		<th rowspan="2" style="background-color:#00A300;color:#fff;">PENGANGKUTAN</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">LANGSUNG</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">TIDAK LANGSUNG</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">% DOUBLE HANDLING</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">KG</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">KG</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">KG</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
</thead>
<tbody>
<%
Set rsAnalisaAngkutTBS = Server.CreateObject("ADODB.Recordset")
rsAnalisaAngkutTBS.Open strAnalisaAngkutTBS, connstr, adOpenStatic

Dim rsGetTotal, strGetTotal
strGetTotal = strAnalisaAngkutTBS
Set rsGetTotal = Server.CreateObject("ADODB.Recordset")
rsGetTotal.Open strGetTotal, connstr, adOpenStatic

Dim Temp_TotalownLangsung : Temp_TotalownLangsung = 0
Dim Temp_TotalownTidakLangsung : Temp_TotalownTidakLangsung = 0
Dim Temp_TotalcontractLangsung : Temp_TotalcontractLangsung = 0
Dim Temp_TotalcontractTidakLangsung : Temp_TotalcontractTidakLangsung = 0
Dim Temp_TotalownTotal : Temp_TotalownTotal = 0
Dim Temp_TotalcontractTotal : Temp_TotalcontractTotal = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

If Not rsGetTotal.EOF Then
	Do While Not rsGetTotal.EOF
		Temp_TotalownLangsung = Temp_TotalownLangsung + Cdbl(to_nol(rsGetTotal("kg_bunch_own_direct")))
		Temp_TotalownTidakLangsung = Temp_TotalownTidakLangsung + Cdbl(to_nol(rsGetTotal("kg_bunch_own_indirect")))
		Temp_TotalcontractLangsung = Temp_TotalcontractLangsung + Cdbl(to_nol(rsGetTotal("kg_bunch_contract_direct")))
		Temp_TotalcontractTidakLangsung = Temp_TotalcontractTidakLangsung + Cdbl(to_nol(rsGetTotal("kg_bunch_contract_indirect")))
		Temp_TotalownTotal = Temp_TotalownTotal + Cdbl(to_nol(rsGetTotal("totalown")))
		Temp_TotalcontractTotal = Temp_TotalcontractTotal + Cdbl(to_nol(rsGetTotal("totalcontract")))
		rsGetTotal.MoveNext
	Loop
End If
rsGetTotal.Close
set rsGetTotal = Nothing

if (rsAnalisaAngkutTBS.bof and rsAnalisaAngkutTBS.eof) then
%>
	<tr>
		<td colspan="<%=var_colspan+8%>">NO DATA AVAILABLE</td>
	</tr>
<%else

Dim TotalLangsung : TotalLangsung = 0
Dim TotalTidakLangsung : TotalTidakLangsung = 0
Dim TotalTotal : TotalTotal = 0
		while not rsAnalisaAngkutTBS.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
			
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsAnalisaAngkutTBS("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsAnalisaAngkutTBS("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsAnalisaAngkutTBS("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsAnalisaAngkutTBS("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsAnalisaAngkutTBS("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsAnalisaAngkutTBS("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
						    <%=rsAnalisaAngkutTBS("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if rsAnalisaAngkutTBS("OWN") = "OWN" then %>
					<td align="center"><%=rsAnalisaAngkutTBS("OWN")%></td>
					
					<td align="right"><%=FormatNumber(rsAnalisaAngkutTBS("kg_bunch_own_direct"), 2)%></td>
					<%If Temp_TotalownLangsung = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("kg_bunch_own_direct"))/Temp_TotalownLangsung)*100, 0), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsAnalisaAngkutTBS("kg_bunch_own_indirect"), 2)%></td>
					<%If Temp_TotalownTidakLangsung = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("kg_bunch_own_indirect"))/Temp_TotalownTidakLangsung)*100, 0), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsAnalisaAngkutTBS("TOTALOWN"), 2)%></td>
					<%If Temp_TotalownTotal = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("TOTALOWN"))/Temp_TotalownTotal)*100, 0), 2)%></td>
					<%End If%>
					
					<%If Cdbl(rsAnalisaAngkutTBS("kg_bunch_own_direct")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsAnalisaAngkutTBS("kg_bunch_own_indirect"))/Cdbl(rsAnalisaAngkutTBS("kg_bunch_own_direct"))*100, 2)%></td>
					<%End If%>
				<%else%>
					<td align="center"><%=rsAnalisaAngkutTBS("OWN")%></td>
				
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsAnalisaAngkutTBS("kg_bunch_contract_direct"))), 2)%></td>
					<%If Temp_TotalcontractLangsung = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("kg_bunch_contract_direct"))/Temp_TotalcontractLangsung)*100, 0), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(rsAnalisaAngkutTBS("kg_bunch_contract_indirect"), 2)%></td>
					<%If Temp_TotalcontractTidakLangsung = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("kg_bunch_contract_indirect"))/Temp_TotalcontractTidakLangsung)*100, 0), 2)%></td>
					<%End If%>
					
					<td align="right"><%=FormatNumber(cdbl(to_nol(rsAnalisaAngkutTBS("TOTALCONTRACT"))), 2)%></td>
					<%If Temp_TotalcontractTotal = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("TOTALCONTRACT"))/Temp_TotalcontractTotal)*100, 0), 2)%></td>
					<%End If%>
					
					<%If Cdbl(to_nol(rsAnalisaAngkutTBS("kg_bunch_contract_direct")) )= 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsAnalisaAngkutTBS("kg_bunch_contract_indirect"))/Cdbl(rsAnalisaAngkutTBS("kg_bunch_contract_direct"))*100, 2)%></td>
					<%End If%>
				<%end if%>
				
			</tr>
<%
			TotalownLangsung = TotalownLangsung + Cdbl(to_nol(rsAnalisaAngkutTBS("kg_bunch_own_direct")) )
			TotalownTidakLangsung = TotalownTidakLangsung + Cdbl(to_nol(rsAnalisaAngkutTBS("kg_bunch_own_indirect")))
			TotalcontractLangsung = TotalcontractLangsung + Cdbl(to_nol(rsAnalisaAngkutTBS("kg_bunch_contract_direct")) )
			TotalcontractTidakLangsung = TotalcontractTidakLangsung + Cdbl(to_nol(rsAnalisaAngkutTBS("kg_bunch_contract_indirect")))
			TotalOWNTotal = TotalOWNTotal + Cdbl(to_nol(rsAnalisaAngkutTBS("TOTALOWN")))
			
			v_TotalLangsung = TotalownLangsung + TotalcontractLangsung
			v_TotaltidakLangsung = TotalownTidakLangsung + TotalcontractTidakLangsung
			v_total = v_TotalLangsung + v_TotaltidakLangsung
			' TotalContractTotal = TotalContractTotal + Cdbl(to_nol(rsAnalisaAngkutTBS("TOTALCONTRACT")))
			OWN = rsAnalisaAngkutTBS("OWN")  
			rsAnalisaAngkutTBS.MoveNext
		Wend
	rsAnalisaAngkutTBS.Close
	set rsAnalisaAngkutTBS = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%=1+var_colspan%>"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FormatNumber(v_TotalLangsung, 2)%></strong></td>
				<%If v_TotalLangsung = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(v_TotalLangsung/v_total*100, 2)%></strong></td>
				<%End if%>
				<td align="right"><strong><%=FormatNumber(v_TotaltidakLangsung, 2)%></strong></td>
				<%If v_TotaltidakLangsung = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(v_TotaltidakLangsung/v_total*100, 2)%></strong></td>
				<%End if%>				
				<td align="right"><strong><%=FormatNumber(v_total, 2)%></strong></td>
				<%If v_total = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(v_total/v_total*100, 2)%></strong></td>
				<%End if%>
				<%If v_TotalLangsung = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(v_TotaltidakLangsung/v_TotalLangsung*100, 2)%></strong></td>
				<%End if%>				
			</tr>
<%End if%>
</tbody>
</table>
		</div>
	</body>
</html>