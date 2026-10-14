<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth
	Dim ScriptName
	Dim rsGrnInv, strGrnInv
	Dim v_txtPOCode
	Dim NO, cClass
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtFcba = Request("txtFcba")
	v_txtMonth = Request("txtMonth")
	v_txtYear = Request("txtYear")
	v_txtsectionname = Request("txtsectionname")
	v_txtempcode = Request("txtempcode")

	txttdate = txtMonth & "" & txtYear
	If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
	If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")

	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
		
       if  v_txtfcba <> "" then
	       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
		end if
	    ' selectCons = selectCons & ",fcba"
	    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		   fcbaname = rsCat("fcName")
		   CompanyName=rsCat("FCCOMPANYNAME")
		   MgrName = rsCat("FCMANAGER")
		   KTUName = rsCat("FCKTU")
		   CropName=rsCat("FCCROP")
		   FcbaType=rsCat("fcType")
	else
	    nSpanTotal = nSpanTotal - 1 
	end if
	
	q = "select type_2, type_3 from allowded where fccode = 'A12' and fcba ='" & v_txtfcba & "' "
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   type_2 = rsCat("type_2")
	   type_3 = cdbl(rsCat("type_3"))
	end if
	
	If type_2 = "M" Then 
	
	strGrnInv		=	" select" & _
						" fcba,employeecode, employee_api.get_employeename(fcba, employeecode) employeename, maritalstatus," & _
						" employee_api.get_sectionname(fcba, employeecode) division," & _
						" beras_1,beras_2,beras_3,beras_4,beras_5,beras_6,beras_7,beras_8,beras_9,beras_10," & _
						" beras_11,beras_12,beras_13,beras_14,beras_15,beras_16,beras_17,beras_18,beras_19,beras_20," & _
						" beras_21,beras_22,beras_23,beras_24,beras_25,beras_26,beras_27,beras_28,beras_29,beras_30,beras_31," & _
						" total_HKE,total_HK_Lebih_Basis," & _
						" case " & _
						"     when kg_beras = 0 then 0 " & _
						"     else kg_beras" & _
						"     end as kg_beras" & _
						" from (" & _
						"    select hk_beras_khl.fcba,hk_beras_khl.employeecode,hk_beras_khl.maritalstatus," & _
						"    hk_beras_khl.beras_1,hk_beras_khl.beras_2,hk_beras_khl.beras_3,hk_beras_khl.beras_4,hk_beras_khl.beras_5,hk_beras_khl.beras_6,hk_beras_khl.beras_7,hk_beras_khl.beras_8,hk_beras_khl.beras_9,hk_beras_khl.beras_10," & _
						"    hk_beras_khl.beras_11,hk_beras_khl.beras_12,hk_beras_khl.beras_13,hk_beras_khl.beras_14,hk_beras_khl.beras_15,hk_beras_khl.beras_16,hk_beras_khl.beras_17,hk_beras_khl.beras_18,hk_beras_khl.beras_19,hk_beras_khl.beras_20," & _
						"    hk_beras_khl.beras_21,hk_beras_khl.beras_22,hk_beras_khl.beras_23,hk_beras_khl.beras_24,hk_beras_khl.beras_25,hk_beras_khl.beras_26,hk_beras_khl.beras_27,hk_beras_khl.beras_28,hk_beras_khl.beras_29,hk_beras_khl.beras_30,hk_beras_khl.beras_31," & _
						"    total_HKE,total_HK_Lebih_Basis," & _
						"    case" & _
						"        when total_HK_Lebih_Basis > 21 then ((total_HKE*par.rate_6)+par.rate_7)" & _
						"        else (total_HKE*par.rate_6)" & _
						"        end as kg_beras" & _
						"    from (" & _
						"        select hk_beras.fcba,hk_beras.employeecode,emp.maritalstatus," & _
						"        hk_beras.beras_1,hk_beras.beras_2,hk_beras.beras_3,hk_beras.beras_4,hk_beras.beras_5,hk_beras.beras_6,hk_beras.beras_7,hk_beras.beras_8,hk_beras.beras_9,hk_beras.beras_10," & _
						"        hk_beras.beras_11,hk_beras.beras_12,hk_beras.beras_13,hk_beras.beras_14,hk_beras.beras_15,hk_beras.beras_16,hk_beras.beras_17,hk_beras.beras_18,hk_beras.beras_19,hk_beras.beras_20," & _
						"        hk_beras.beras_21,hk_beras.beras_22,hk_beras.beras_23,hk_beras.beras_24,hk_beras.beras_25,hk_beras.beras_26,hk_beras.beras_27,hk_beras.beras_28,hk_beras.beras_29,hk_beras.beras_30,hk_beras.beras_31," & _
						"        (" & _
						"        hk_beras.HKE_1+hk_beras.HKE_2+hk_beras.HKE_3+hk_beras.HKE_4+hk_beras.HKE_5+hk_beras.HKE_6+hk_beras.HKE_7+hk_beras.HKE_8+hk_beras.HKE_9+hk_beras.HKE_10+" & _
						"        hk_beras.HKE_11+hk_beras.HKE_12+hk_beras.HKE_13+hk_beras.HKE_14+hk_beras.HKE_15+hk_beras.HKE_16+hk_beras.HKE_17+hk_beras.HKE_18+hk_beras.HKE_19+hk_beras.HKE_20+" & _
						"        hk_beras.HKE_21+hk_beras.HKE_22+hk_beras.HKE_23+hk_beras.HKE_24+hk_beras.HKE_25+hk_beras.HKE_26+hk_beras.HKE_27+hk_beras.HKE_28+hk_beras.HKE_29+hk_beras.HKE_30+hk_beras.HKE_31" & _
						"        ) as total_HKE,       " & _
						"        (" & _
						"        hk_beras.beras_1+hk_beras.beras_2+hk_beras.beras_3+hk_beras.beras_4+hk_beras.beras_5+hk_beras.beras_6+hk_beras.beras_7+hk_beras.beras_8+hk_beras.beras_9+hk_beras.beras_10+" & _
						"        hk_beras.beras_11+hk_beras.beras_12+hk_beras.beras_13+hk_beras.beras_14+hk_beras.beras_15+hk_beras.beras_16+hk_beras.beras_17+hk_beras.beras_18+hk_beras.beras_19+hk_beras.beras_20+" & _
						"        hk_beras.beras_21+hk_beras.beras_22+hk_beras.beras_23+hk_beras.beras_24+hk_beras.beras_25+hk_beras.beras_26+hk_beras.beras_27+hk_beras.beras_28+hk_beras.beras_29+hk_beras.beras_30+hk_beras.beras_31" & _
						"        ) as total_HK_Lebih_Basis" & _
						"        from (" & _
						"            select" & _
						"            fcba,employeecode," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),1,HKE,0)) as HKE_1," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),2,HKE,0)) as HKE_2," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),3,HKE,0)) as HKE_3," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),4,HKE,0)) as HKE_4," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),5,HKE,0)) as HKE_5," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),6,HKE,0)) as HKE_6," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),7,HKE,0)) as HKE_7," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),8,HKE,0)) as HKE_8," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),9,HKE,0)) as HKE_9," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),10,HKE,0)) as HKE_10," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),11,HKE,0)) as HKE_11," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),12,HKE,0)) as HKE_12," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),13,HKE,0)) as HKE_13," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),14,HKE,0)) as HKE_14," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),15,HKE,0)) as HKE_15," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),16,HKE,0)) as HKE_16," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),17,HKE,0)) as HKE_17," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),18,HKE,0)) as HKE_18," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),19,HKE,0)) as HKE_19," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),20,HKE,0)) as HKE_20," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),21,HKE,0)) as HKE_21," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),22,HKE,0)) as HKE_22," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),23,HKE,0)) as HKE_23," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),24,HKE,0)) as HKE_24," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),25,HKE,0)) as HKE_25," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),26,HKE,0)) as HKE_26," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),27,HKE,0)) as HKE_27," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),28,HKE,0)) as HKE_28," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),29,HKE,0)) as HKE_29," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),30,HKE,0)) as HKE_30," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),31,HKE,0)) as HKE_31,           " & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),1,hitung_lebih_basis,0)) as beras_1," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),2,hitung_lebih_basis,0)) as beras_2," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),3,hitung_lebih_basis,0)) as beras_3," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),4,hitung_lebih_basis,0)) as beras_4," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),5,hitung_lebih_basis,0)) as beras_5," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),6,hitung_lebih_basis,0)) as beras_6," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),7,hitung_lebih_basis,0)) as beras_7," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),8,hitung_lebih_basis,0)) as beras_8," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),9,hitung_lebih_basis,0)) as beras_9," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),10,hitung_lebih_basis,0)) as beras_10," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),11,hitung_lebih_basis,0)) as beras_11," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),12,hitung_lebih_basis,0)) as beras_12," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),13,hitung_lebih_basis,0)) as beras_13," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),14,hitung_lebih_basis,0)) as beras_14," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),15,hitung_lebih_basis,0)) as beras_15," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),16,hitung_lebih_basis,0)) as beras_16," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),17,hitung_lebih_basis,0)) as beras_17," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),18,hitung_lebih_basis,0)) as beras_18," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),19,hitung_lebih_basis,0)) as beras_19," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),20,hitung_lebih_basis,0)) as beras_20," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),21,hitung_lebih_basis,0)) as beras_21," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),22,hitung_lebih_basis,0)) as beras_22," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),23,hitung_lebih_basis,0)) as beras_23," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),24,hitung_lebih_basis,0)) as beras_24," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),25,hitung_lebih_basis,0)) as beras_25," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),26,hitung_lebih_basis,0)) as beras_26," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),27,hitung_lebih_basis,0)) as beras_27," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),28,hitung_lebih_basis,0)) as beras_28," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),29,hitung_lebih_basis,0)) as beras_29," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),30,hitung_lebih_basis,0)) as beras_30," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),31,hitung_lebih_basis,0)) as beras_31" & _
						"            from (" & _
						"                 select " & _
						"                 fcba,fddate,employeecode,hitung_lebih_basis,count(hitung_lebih_basis) as HKE" & _
						"                 from (           " & _
						"                     select" & _
						"                     fcba,fddate,employeecode," & _
						"                     case" & _
						"                     when extraquantity > (select type_3 from iplasprod.allowded where fcba = '"& v_txtFcba &"' and fccode = 'A12') then 1" & _
						"                     else 0" & _
						"                     end as hitung_lebih_basis" & _
						"                     from (" & _
						"                         select iplasprod.harvesting_daily_bjr.fcba," & _
						"                         iplasprod.harvesting_daily_bjr.fddate," & _
						"                         iplasprod.harvesting_daily_bjr.employeecode," & _
						"                         case" & _
						"                             when ald.type_2 = 'D'" & _
						"                             then sum (iplasprod.harvesting_daily_bjr.extraquantity_bydaily)" & _
						"                             else sum (iplasprod.harvesting_daily_bjr.extraquantity_bymonthly)" & _
						"                             end as extraquantity" & _
						"                         from iplasprod.harvesting_daily_bjr" & _
						"                         inner join (" & _
						"                             select iplasprod.allowded.fcba," & _
						"                             iplasprod.allowded.fccode," & _
						"                             iplasprod.allowded.type_2" & _
						"                             from iplasprod.allowded" & _
						"                             where iplasprod.allowded.fccode = 'A12') ald" & _
						"                         on (iplasprod.harvesting_daily_bjr.fcba = ald.fcba)" & _
						"                         where iplasprod.harvesting_daily_bjr.presentgrade = 'KHL'" & _
						"                         and iplasprod.harvesting_daily_bjr.fcba = '"& v_txtFcba &"'" & _
						"                         and to_number ( to_char ( to_date (iplasprod.harvesting_daily_bjr.fddate, 'mm/dd/yyyy'), 'mm')) = '"& v_txtMonth &"'" & _
						"                         and to_number ( to_char ( to_date (iplasprod.harvesting_daily_bjr.fddate, 'mm/dd/yyyy'), 'yyyy')) = '"& v_txtYear &"'" & _
						"                         group by " & _
						"                         iplasprod.harvesting_daily_bjr.fcba," & _
						"                         iplasprod.harvesting_daily_bjr.fddate," & _
						"                         iplasprod.harvesting_daily_bjr.employeecode," & _
						"                         ald.type_2" & _
						"                     )" & _
						"                )" & _
						"                group by" & _
						"                fcba,fddate,employeecode,hitung_lebih_basis" & _
						"            )" & _
						"            group by" & _
						"            fcba,employeecode" & _
						"        ) hk_beras" & _
						"        inner join iplasprod.employee emp" & _
						"        on (" & _
						"        hk_beras.fcba = emp.fcba and" & _
						"        hk_beras.employeecode = emp.fccode )" & _
						"    ) hk_beras_khl" & _
						"    inner join (" & _
						"        select fcba,fccode,rate_6,rate_7 from iplasprod.parameter" & _
						"        where parcode = 'E0006'" & _
						"        and parname = 'MARITAL STATUS'" & _
						"    ) par" & _
						"    on (" & _
						"    hk_beras_khl.fcba = par.fcba and" & _
						"    hk_beras_khl.maritalstatus = par.fccode )" & _
						" )" & _
						" where kg_beras > 0" 
	ELSE
		strGrnInv=		" select" & _
						" fcba,employeecode, employee_api.get_employeename(fcba, employeecode) employeename, maritalstatus," & _
						" employee_api.get_sectionname(fcba, employeecode) division," & _
						" beras_1,beras_2,beras_3,beras_4,beras_5,beras_6,beras_7,beras_8,beras_9,beras_10," & _
						" beras_11,beras_12,beras_13,beras_14,beras_15,beras_16,beras_17,beras_18,beras_19,beras_20," & _
						" beras_21,beras_22,beras_23,beras_24,beras_25,beras_26,beras_27,beras_28,beras_29,beras_30,beras_31," & _
						" total_HKE,total_HK_Lebih_Basis," & _
						" case " & _
						"     when kg_beras = 0 then 0 " & _
						"     else kg_beras" & _
						"     end as kg_beras" & _
						" from (" & _
						"    select hk_beras_khl.fcba,hk_beras_khl.employeecode,hk_beras_khl.maritalstatus," & _
						"    hk_beras_khl.beras_1,hk_beras_khl.beras_2,hk_beras_khl.beras_3,hk_beras_khl.beras_4,hk_beras_khl.beras_5,hk_beras_khl.beras_6,hk_beras_khl.beras_7,hk_beras_khl.beras_8,hk_beras_khl.beras_9,hk_beras_khl.beras_10," & _
						"    hk_beras_khl.beras_11,hk_beras_khl.beras_12,hk_beras_khl.beras_13,hk_beras_khl.beras_14,hk_beras_khl.beras_15,hk_beras_khl.beras_16,hk_beras_khl.beras_17,hk_beras_khl.beras_18,hk_beras_khl.beras_19,hk_beras_khl.beras_20," & _
						"    hk_beras_khl.beras_21,hk_beras_khl.beras_22,hk_beras_khl.beras_23,hk_beras_khl.beras_24,hk_beras_khl.beras_25,hk_beras_khl.beras_26,hk_beras_khl.beras_27,hk_beras_khl.beras_28,hk_beras_khl.beras_29,hk_beras_khl.beras_30,hk_beras_khl.beras_31," & _
						"    total_HKE,total_HK_Lebih_Basis," & _
						"    case" & _
						"        when total_HK_Lebih_Basis > 21 then ((total_HKE*par.rate_6)+par.rate_7)" & _
						"        else (total_HKE*par.rate_6)" & _
						"        end as kg_beras" & _
						"    from (" & _
						"        select hk_beras.fcba,hk_beras.employeecode,emp.maritalstatus," & _
						"        hk_beras.beras_1,hk_beras.beras_2,hk_beras.beras_3,hk_beras.beras_4,hk_beras.beras_5,hk_beras.beras_6,hk_beras.beras_7,hk_beras.beras_8,hk_beras.beras_9,hk_beras.beras_10," & _
						"        hk_beras.beras_11,hk_beras.beras_12,hk_beras.beras_13,hk_beras.beras_14,hk_beras.beras_15,hk_beras.beras_16,hk_beras.beras_17,hk_beras.beras_18,hk_beras.beras_19,hk_beras.beras_20," & _
						"        hk_beras.beras_21,hk_beras.beras_22,hk_beras.beras_23,hk_beras.beras_24,hk_beras.beras_25,hk_beras.beras_26,hk_beras.beras_27,hk_beras.beras_28,hk_beras.beras_29,hk_beras.beras_30,hk_beras.beras_31," & _
						"        (" & _
						"        hk_beras.HKE_1+hk_beras.HKE_2+hk_beras.HKE_3+hk_beras.HKE_4+hk_beras.HKE_5+hk_beras.HKE_6+hk_beras.HKE_7+hk_beras.HKE_8+hk_beras.HKE_9+hk_beras.HKE_10+" & _
						"        hk_beras.HKE_11+hk_beras.HKE_12+hk_beras.HKE_13+hk_beras.HKE_14+hk_beras.HKE_15+hk_beras.HKE_16+hk_beras.HKE_17+hk_beras.HKE_18+hk_beras.HKE_19+hk_beras.HKE_20+" & _
						"        hk_beras.HKE_21+hk_beras.HKE_22+hk_beras.HKE_23+hk_beras.HKE_24+hk_beras.HKE_25+hk_beras.HKE_26+hk_beras.HKE_27+hk_beras.HKE_28+hk_beras.HKE_29+hk_beras.HKE_30+hk_beras.HKE_31" & _
						"        ) as total_HKE," & _
						"        (" & _
						"        hk_beras.beras_1+hk_beras.beras_2+hk_beras.beras_3+hk_beras.beras_4+hk_beras.beras_5+hk_beras.beras_6+hk_beras.beras_7+hk_beras.beras_8+hk_beras.beras_9+hk_beras.beras_10+" & _
						"        hk_beras.beras_11+hk_beras.beras_12+hk_beras.beras_13+hk_beras.beras_14+hk_beras.beras_15+hk_beras.beras_16+hk_beras.beras_17+hk_beras.beras_18+hk_beras.beras_19+hk_beras.beras_20+" & _
						"        hk_beras.beras_21+hk_beras.beras_22+hk_beras.beras_23+hk_beras.beras_24+hk_beras.beras_25+hk_beras.beras_26+hk_beras.beras_27+hk_beras.beras_28+hk_beras.beras_29+hk_beras.beras_30+hk_beras.beras_31" & _
						"        ) as total_HK_Lebih_Basis" & _
						"        from (" & _
						"            select" & _
						"            fcba,employeecode," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),1,HKE,0)) as HKE_1," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),2,HKE,0)) as HKE_2," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),3,HKE,0)) as HKE_3," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),4,HKE,0)) as HKE_4," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),5,HKE,0)) as HKE_5," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),6,HKE,0)) as HKE_6," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),7,HKE,0)) as HKE_7," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),8,HKE,0)) as HKE_8," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),9,HKE,0)) as HKE_9," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),10,HKE,0)) as HKE_10," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),11,HKE,0)) as HKE_11," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),12,HKE,0)) as HKE_12," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),13,HKE,0)) as HKE_13," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),14,HKE,0)) as HKE_14," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),15,HKE,0)) as HKE_15," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),16,HKE,0)) as HKE_16," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),17,HKE,0)) as HKE_17," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),18,HKE,0)) as HKE_18," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),19,HKE,0)) as HKE_19," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),20,HKE,0)) as HKE_20," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),21,HKE,0)) as HKE_21," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),22,HKE,0)) as HKE_22," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),23,HKE,0)) as HKE_23," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),24,HKE,0)) as HKE_24," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),25,HKE,0)) as HKE_25," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),26,HKE,0)) as HKE_26," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),27,HKE,0)) as HKE_27," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),28,HKE,0)) as HKE_28," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),29,HKE,0)) as HKE_29," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),30,HKE,0)) as HKE_30," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),31,HKE,0)) as HKE_31," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),1,hitung_lebih_basis,0)) as beras_1," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),2,hitung_lebih_basis,0)) as beras_2," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),3,hitung_lebih_basis,0)) as beras_3," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),4,hitung_lebih_basis,0)) as beras_4," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),5,hitung_lebih_basis,0)) as beras_5," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),6,hitung_lebih_basis,0)) as beras_6," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),7,hitung_lebih_basis,0)) as beras_7," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),8,hitung_lebih_basis,0)) as beras_8," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),9,hitung_lebih_basis,0)) as beras_9," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),10,hitung_lebih_basis,0)) as beras_10," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),11,hitung_lebih_basis,0)) as beras_11," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),12,hitung_lebih_basis,0)) as beras_12," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),13,hitung_lebih_basis,0)) as beras_13," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),14,hitung_lebih_basis,0)) as beras_14," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),15,hitung_lebih_basis,0)) as beras_15," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),16,hitung_lebih_basis,0)) as beras_16," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),17,hitung_lebih_basis,0)) as beras_17," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),18,hitung_lebih_basis,0)) as beras_18," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),19,hitung_lebih_basis,0)) as beras_19," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),20,hitung_lebih_basis,0)) as beras_20," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),21,hitung_lebih_basis,0)) as beras_21," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),22,hitung_lebih_basis,0)) as beras_22," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),23,hitung_lebih_basis,0)) as beras_23," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),24,hitung_lebih_basis,0)) as beras_24," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),25,hitung_lebih_basis,0)) as beras_25," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),26,hitung_lebih_basis,0)) as beras_26," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),27,hitung_lebih_basis,0)) as beras_27," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),28,hitung_lebih_basis,0)) as beras_28," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),29,hitung_lebih_basis,0)) as beras_29," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),30,hitung_lebih_basis,0)) as beras_30," & _
						"            sum(decode(to_number(to_char(to_date(fddate,'mm/dd/yyyy'),'dd')),31,hitung_lebih_basis,0)) as beras_31         " & _
						"            from (" & _
						"                 select " & _
						"                 fcba,fddate,employeecode,hitung_lebih_basis,count(hitung_lebih_basis) as HKE" & _
						"                 from (" & _
						"                     select" & _
						"                     fcba,fddate,employeecode," & _
						"                     case" & _
						"                     when extraquantity > (select type_3 from iplasprod.allowded where fcba = '"& v_txtFcba &"' and fccode = 'A12') then 1" & _
						"                     else 0" & _
						"                     end as hitung_lebih_basis" & _
						"                     from (" & _
						"                         select iplasprod.harvesting_daily_bjr.fcba," & _
						"                         iplasprod.harvesting_daily_bjr.fddate," & _
						"                         iplasprod.harvesting_daily_bjr.employeecode," & _
						"                         case" & _
						"                             when ald.type_2 = 'D'" & _
						"                             then sum (iplasprod.harvesting_daily_bjr.extraquantity_bydaily)" & _
						"                             else sum (iplasprod.harvesting_daily_bjr.extraquantity_bymonthly)" & _
						"                             end as extraquantity" & _
						"                         from iplasprod.harvesting_daily_bjr" & _
						"                         inner join (" & _
						"                             select iplasprod.allowded.fcba," & _
						"                             iplasprod.allowded.fccode," & _
						"                             iplasprod.allowded.type_2" & _
						"                             from iplasprod.allowded" & _
						"                             where iplasprod.allowded.fccode = 'A12') ald" & _
						"                         on (iplasprod.harvesting_daily_bjr.fcba = ald.fcba)" & _
						"                         where iplasprod.harvesting_daily_bjr.presentgrade = 'KHL'" & _
						"                         and iplasprod.harvesting_daily_bjr.fcba = '"& v_txtFcba &"'" & _
						"                         and to_number ( to_char ( to_date (iplasprod.harvesting_daily_bjr.fddate, 'mm/dd/yyyy'), 'mm')) = '"& v_txtMonth &"'" & _
						"                         and to_number ( to_char ( to_date (iplasprod.harvesting_daily_bjr.fddate, 'mm/dd/yyyy'), 'yyyy')) = '"& v_txtYear &"'" & _
						"                         group by " & _
						"                         iplasprod.harvesting_daily_bjr.fcba," & _
						"                         iplasprod.harvesting_daily_bjr.fddate," & _
						"                         iplasprod.harvesting_daily_bjr.employeecode," & _
						"                         ald.type_2" & _
						"                     )" & _
						"                 )" & _
						"                 group by " & _
						"                 fcba,fddate,employeecode,hitung_lebih_basis" & _
						"            )" & _
						"            group by" & _
						"            fcba,employeecode" & _
						"        ) hk_beras" & _
						"        inner join iplasprod.employee emp" & _
						"        on (" & _
						"        hk_beras.fcba = emp.fcba and" & _
						"        hk_beras.employeecode = emp.fccode )" & _
						"    ) hk_beras_khl" & _
						"    inner join (" & _
						"        select fcba,fccode,rate_6,rate_7 from iplasprod.parameter" & _
						"        where parcode = 'E0006'" & _
						"        and parname = 'MARITAL STATUS'" & _
						"    ) par" & _
						"    on (" & _
						"    hk_beras_khl.fcba = par.fcba and" & _
						"    hk_beras_khl.maritalstatus = par.fccode )" & _
						" )" & _
						" where kg_beras > 0" 
	End If

   
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
	end if 
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND employee_api.get_sectionname(fcba, employeecode) ='"& v_txtsectionname &"'"
	end if 


	' if  cdbl(to_nol(compare1)) = 1 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != vehicleactivity.vh_employeecode"
	' end if
	' if  cdbl(to_nol(compare1)) = 2 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != harvestingspb.harvesting_employeecode" 
	' end if
	' if  cdbl(to_nol(compare1)) = 3 then
	'     strGrnInv	=	strGrnInv & " and vehicleactivity.vh_employeecode != harvestingspb.harvesting_employeecode" 
	' end if

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY employeecode"

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Natura Pemanen KHL</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
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
				<table id="form-table" class="centered-margin">
					<tr> 
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">KEBUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">TAHUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">BULAN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DIVISION</font>
						</td> 
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">EMPLOYEE CODE</font>
						</td>
					</tr>
					<tr>	 
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<option></option>
							<%
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
							  
							
							rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsFcba.eof then
								DO while not rsFcba.eof
									if v_txtFcba  = rsFcba("fcba") then
									%>			
										<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
									<% Else %>	
										<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
									<% End If
									
									rsFcba.movenext
								loop
							End if
							rsFcba.Close
							set rsFcba = Nothing
							%>
							</select>	
						</td> 
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
						%>
						<td width="4">
							<select name="txtYear">
							<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%Next%>
							</select>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td> 
						<td>
							<select name="txtsectionname" >
								<option value="">-- ALL --</option>
									<%
									Dim rssectionname, strsectionname
									strsectionname	=	"select fccode, fcname from sectionname where fccode in ('AFD-01','AFD-02','AFD-03','AFD-04','AFD-05','AFD-06','AFD-07','AFD-08','AFD-09','AFD-10','AFD-11','AFD-12','AFD-13') order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtsectionname then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<select name="txtempcode" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode, fcname from employee where fcba = '"& v_txtfcba &"' order by fcCode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%> </option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td>
						<!--
						<td>
							<font face="verdana" size="1">VEHICLE CODE</font>
						</td>
						<td width="2">
							<input type="text" name="txtvehiclecode" value="<%=txtvehiclecode%>">
						</td> -->
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
			
			</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="38" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>LAPORAN NATURA PEMANEN KHL</h2>
				<h4>PERIODE : Year <%=v_txtyear%>, Month <%=v_txtmonth%></h4>
                <br/>
			</div>
		</div>
</th>
		</tr>
	<tr>
			<th style="background-color:#00A300;color:#fff;">Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">Marital Status</th>
			<th style="background-color:#00A300;color:#fff;">Division</th>
			<th style="background-color:#00A300;color:#fff;">1</th> 
			<th style="background-color:#00A300;color:#fff;">2</th> 
			<th style="background-color:#00A300;color:#fff;">3</th> 
			<th style="background-color:#00A300;color:#fff;">4</th> 
			<th style="background-color:#00A300;color:#fff;">5</th> 
			<th style="background-color:#00A300;color:#fff;">6</th> 
			<th style="background-color:#00A300;color:#fff;">7</th> 
			<th style="background-color:#00A300;color:#fff;">8</th> 
			<th style="background-color:#00A300;color:#fff;">9</th> 
			<th style="background-color:#00A300;color:#fff;">10</th> 
			<th style="background-color:#00A300;color:#fff;">11</th> 
			<th style="background-color:#00A300;color:#fff;">12</th> 
			<th style="background-color:#00A300;color:#fff;">13</th> 
			<th style="background-color:#00A300;color:#fff;">14</th> 
			<th style="background-color:#00A300;color:#fff;">15</th> 
			<th style="background-color:#00A300;color:#fff;">16</th> 
			<th style="background-color:#00A300;color:#fff;">17</th> 
			<th style="background-color:#00A300;color:#fff;">18</th> 
			<th style="background-color:#00A300;color:#fff;">19</th> 
			<th style="background-color:#00A300;color:#fff;">20</th> 
			<th style="background-color:#00A300;color:#fff;">21</th> 
			<th style="background-color:#00A300;color:#fff;">22</th> 
			<th style="background-color:#00A300;color:#fff;">23</th> 
			<th style="background-color:#00A300;color:#fff;">24</th> 
			<th style="background-color:#00A300;color:#fff;">25</th> 
			<th style="background-color:#00A300;color:#fff;">26</th> 
			<th style="background-color:#00A300;color:#fff;">27</th> 
			<th style="background-color:#00A300;color:#fff;">28</th> 
			<th style="background-color:#00A300;color:#fff;">29</th> 
			<th style="background-color:#00A300;color:#fff;">30</th> 
			<th style="background-color:#00A300;color:#fff;">31</th> 
			<th style="background-color:#00A300;color:#fff;">Total HKE</th>
			<th style="background-color:#00A300;color:#fff;">Total HK Lebih Basis</th> 
			<th style="background-color:#00A300;color:#fff;">KG Beras</th> 
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="28" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_beras_1 = 0
		v_beras_2 = 0
		v_beras_3 = 0
		v_beras_4 = 0
		v_beras_5 = 0
		v_beras_6 = 0
		v_beras_7 = 0
		v_beras_8 = 0
		v_beras_9 = 0
		v_beras_10 = 0
		v_beras_11 = 0
		v_beras_12 = 0
		v_beras_13 = 0
		v_beras_14 = 0
		v_beras_15 = 0
		v_beras_16 = 0
		v_beras_17 = 0
		v_beras_18 = 0
		v_beras_19 = 0
		v_beras_20 = 0
		v_beras_21 = 0
		v_beras_22 = 0
		v_beras_23 = 0
		v_beras_24 = 0
		v_beras_25 = 0
		v_beras_26 = 0
		v_beras_27 = 0
		v_beras_28 = 0
		v_beras_29 = 0
		v_beras_30 = 0
		v_beras_31 = 0
		v_total = 0
		v_total_hk_lebih_basis = 0
		v_kg_beras = 0
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="center"><%=rsGrnInv("employeecode")%></td>
			<td align="left"><%=rsGrnInv("employeename")%></td>
			<td align="center"><%=rsGrnInv("maritalstatus")%></td>
			<td align="center"><%=rsGrnInv("division")%></td>
			<td align="left"><%=rsGrnInv("beras_1")%></td>
			<td align="left"><%=rsGrnInv("beras_2")%></td>
			<td align="left"><%=rsGrnInv("beras_3")%></td>
			<td align="left"><%=rsGrnInv("beras_4")%></td>
			<td align="left"><%=rsGrnInv("beras_5")%></td>
			<td align="left"><%=rsGrnInv("beras_6")%></td>
			<td align="left"><%=rsGrnInv("beras_7")%></td>
			<td align="left"><%=rsGrnInv("beras_8")%></td>
			<td align="left"><%=rsGrnInv("beras_9")%></td>
			<td align="left"><%=rsGrnInv("beras_10")%></td>
			<td align="left"><%=rsGrnInv("beras_11")%></td>
			<td align="left"><%=rsGrnInv("beras_12")%></td>
			<td align="left"><%=rsGrnInv("beras_13")%></td>
			<td align="left"><%=rsGrnInv("beras_14")%></td>
			<td align="left"><%=rsGrnInv("beras_15")%></td>
			<td align="left"><%=rsGrnInv("beras_16")%></td>
			<td align="left"><%=rsGrnInv("beras_17")%></td>
			<td align="left"><%=rsGrnInv("beras_18")%></td>
			<td align="left"><%=rsGrnInv("beras_19")%></td>
			<td align="left"><%=rsGrnInv("beras_20")%></td>
			<td align="left"><%=rsGrnInv("beras_21")%></td>
			<td align="left"><%=rsGrnInv("beras_22")%></td>
			<td align="left"><%=rsGrnInv("beras_23")%></td>
			<td align="left"><%=rsGrnInv("beras_24")%></td>
			<td align="left"><%=rsGrnInv("beras_25")%></td>
			<td align="left"><%=rsGrnInv("beras_26")%></td>
			<td align="left"><%=rsGrnInv("beras_27")%></td>
			<td align="left"><%=rsGrnInv("beras_28")%></td>
			<td align="left"><%=rsGrnInv("beras_29")%></td>
			<td align="left"><%=rsGrnInv("beras_30")%></td>
			<td align="left"><%=rsGrnInv("beras_31")%></td>
			<td align="left"><%=rsGrnInv("total_hke")%></td>
			<td align="left"><%=rsGrnInv("total_HK_Lebih_Basis")%></td>
			<td align="left"><%=rsGrnInv("kg_beras")%></td>
		</tr>
<%			
			v_beras_1 = v_beras_1 + cdbl(to_nol(rsGrnInv("beras_1"))) 
			v_beras_2 = v_beras_2 + cdbl(to_nol(rsGrnInv("beras_2"))) 
			v_beras_3 = v_beras_3 + cdbl(to_nol(rsGrnInv("beras_3"))) 
			v_beras_4 = v_beras_4 + cdbl(to_nol(rsGrnInv("beras_4"))) 
			v_beras_5 = v_beras_5 + cdbl(to_nol(rsGrnInv("beras_5"))) 
			v_beras_6 = v_beras_6 + cdbl(to_nol(rsGrnInv("beras_6"))) 
			v_beras_7 = v_beras_7 + cdbl(to_nol(rsGrnInv("beras_7"))) 
			v_beras_8 = v_beras_8 + cdbl(to_nol(rsGrnInv("beras_8"))) 
			v_beras_9 = v_beras_9 + cdbl(to_nol(rsGrnInv("beras_9"))) 
			v_beras_10 = v_beras_10 + cdbl(to_nol(rsGrnInv("beras_10"))) 
			v_beras_11 = v_beras_11 + cdbl(to_nol(rsGrnInv("beras_11"))) 
			v_beras_12 = v_beras_12 + cdbl(to_nol(rsGrnInv("beras_12"))) 
			v_beras_13 = v_beras_13 + cdbl(to_nol(rsGrnInv("beras_13"))) 
			v_beras_14 = v_beras_14 + cdbl(to_nol(rsGrnInv("beras_14"))) 
			v_beras_15 = v_beras_15 + cdbl(to_nol(rsGrnInv("beras_15"))) 
			v_beras_16 = v_beras_16 + cdbl(to_nol(rsGrnInv("beras_16"))) 
			v_beras_17 = v_beras_17 + cdbl(to_nol(rsGrnInv("beras_17"))) 
			v_beras_18 = v_beras_18 + cdbl(to_nol(rsGrnInv("beras_18"))) 
			v_beras_19 = v_beras_19 + cdbl(to_nol(rsGrnInv("beras_19"))) 
			v_beras_20 = v_beras_20 + cdbl(to_nol(rsGrnInv("beras_20"))) 
			v_beras_21 = v_beras_21 + cdbl(to_nol(rsGrnInv("beras_21"))) 
			v_beras_22 = v_beras_22 + cdbl(to_nol(rsGrnInv("beras_22"))) 
			v_beras_23 = v_beras_23 + cdbl(to_nol(rsGrnInv("beras_23"))) 
			v_beras_24 = v_beras_24 + cdbl(to_nol(rsGrnInv("beras_24"))) 
			v_beras_25 = v_beras_25 + cdbl(to_nol(rsGrnInv("beras_25"))) 
			v_beras_26 = v_beras_26 + cdbl(to_nol(rsGrnInv("beras_26"))) 
			v_beras_27 = v_beras_27 + cdbl(to_nol(rsGrnInv("beras_27"))) 
			v_beras_28 = v_beras_28 + cdbl(to_nol(rsGrnInv("beras_28"))) 
			v_beras_29 = v_beras_29 + cdbl(to_nol(rsGrnInv("beras_29"))) 
			v_beras_30 = v_beras_30 + cdbl(to_nol(rsGrnInv("beras_30"))) 
			v_beras_31 = v_beras_31 + cdbl(to_nol(rsGrnInv("beras_31"))) 
			v_total = v_total + cdbl(to_nol(rsGrnInv("total_hke"))) 
			v_total_hk_lebih_basis = v_total_hk_lebih_basis + cdbl(to_nol(rsGrnInv("total_HK_Lebih_Basis"))) 
			v_kg_beras = v_kg_beras + cdbl(to_nol(rsGrnInv("kg_beras")))

			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="4"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_1)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_2)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_3)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_4)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_5)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_6)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_7)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_8)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_9)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_10)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_11)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_12)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_13)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_14)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_15)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_16)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_17)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_18)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_19)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_20)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_21)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_22)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_23)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_24)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_25)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_26)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_27)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_28)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_29)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_30)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_beras_31)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_hk_lebih_basis)),0)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_beras)),0)%></font></td>
		</tr>
	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	var config = {
	  '.chosen-select'           : {},
	  '.chosen-select-deselect'  : {allow_single_deselect:true},
	  '.chosen-select-no-single' : {disable_search_threshold:10},
	  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	  '.chosen-select-width'     : {width:"45%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");
	
</script>