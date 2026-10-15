<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsPerawatanTM, strPerawatanTM, strGroup, strOrder, strQuery
	Dim v_txtdatefrom, v_txtdateto, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass, v_txtStatusTmTbm, v_txtjobcode, v_txtGroup
	Dim nSpanTotal
	Dim txtdatefrom, txtdateto, txtDivision, txtBlok, txtPlantingDate, txtStatusTmTbm, txtjobcode, txtGroup
	Dim sprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode, sprevStatus, sprevJobCode, sprevGroup
    Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempDivision, tempPlantingDate,  tempBlok,  tempStatus,  tempJobcode,  tempGroup 
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	Dim ExtendedQuery, supportExtendedQuery, strExtendedQuery,strEPS2 
    'Dim rscomp , strcomp,txtcompany ,txcompanyname,txtComp,v_txtComp
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
		
	 nSpanTotal = 10
	
	v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
	v_txtfcba 			   = Request("txtfcba")
	v_txtDivision		   = Request("txtDivision")
	v_txtBlok			   = Request("txtBlok")
	v_txtPlantingDate	   = Request("txtPlantingDate")
	v_txtdatefrom		   = Request("txtdatefrom")
	v_txtdateto			   = Request("txtdateto")
	v_txtStatusTmTbm       = Request("txtStatusTmTbm")
	v_txtjobcode     	   = Request("txtjobcode")
	v_txtGroup			   = Request("txtGroup")

	        'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " TT" & v_txtPlantingDate
			'response.write  "TM/TBM " & v_txtStatusTmTbm & "  jobcode " & v_txtjobcode & "  Group " &  v_txtGroup

					
			Dim selectFields,strFieldCons,selectCons,whereCons
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',a.fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',a.fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',a.fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',a.fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',a.fcba) company, BUSINESSUNIT_API.get_OrderData(a.fcba) norut, a.fcba, " 

						   
			wherecons = " where company is not null "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
			if  v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "" then
				if v_txtStatusTmTbm <> "--ALL--" then
					whereCons = whereCons &  " AND STATUS = '" &  v_txtStatusTmTbm & "'"
				end if 
                selectCons = selectCons &  " , STATUS "				   
			else
			   nSpanTotal = nSpanTotal - 1 
			end if	
			if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",norut,fcba"
			    
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
			if  v_txtDivision <> "--SUMMARY--" AND v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
			else
			   nSpanTotal = nSpanTotal - 1 
			end if
            
			if v_txtPlantingDate <> "--SUMMARY--"  AND  v_txtPlantingDate <> "" then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
			else
			   nSpanTotal = nSpanTotal - 1 
			end if
			
			if  v_txtBlok <> "--SUMMARY--" AND v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , FIELDCODE "	
			else
			   nSpanTotal = nSpanTotal - 1 
            end if
			
	        
			if v_txtjobcode <> "--SUMMARY--" AND v_txtjobcode <> "--SUMMARY--" then
				   if v_txtjobcode <> "--ALL--" then
					  whereCons = whereCons &  " AND  jobcode = '" & v_txtjobcode & "'"
				   end if
			      selectCons = selectCons &  " , jobcode"				   
			else
			   nSpanTotal = nSpanTotal - 1 
			end if	
	
    'strPerawatanTM = " SELECT c.param_2, a.status "
    ' strPerawatanTM = " SELECT c.param_2, c.param_2_name "
	
	' supportExtendedQuery = 0

	

	strPerawatanTM 		=      " select c.param_2,c.param_2_name,a.status,a.jobcode," & strFieldCons & "a.DIVISION,a.FIELDCODE,a.fcname,a.plantingyear,round(SUM(nvl(a.sku,0)),2) sku, round(SUM(nvl(a.khl,0)),2) khl, " & _ 
							   " round(SUM(nvl(a.total,0)),2) total, round(SUM(nvl(a.ha_harian,0)),2) ha_harian, " & _ 
							   " round(SUM(nvl(a.bo_harian,0)),2) bo_harian, round(SUM(nvl(a.total_bo_harian,0)),2) total_bo_harian, " & _ 
							   " round(SUM(nvl(a.gli_ltr, 0)), 2) gli_ltr, " & _
							   " round(SUM(nvl(a.par_ltr, 0)), 2) par_ltr, " & _
							   " round(SUM(nvl(a.met_kg, 0)), 2) met_kg, " & _
							   " round(SUM(nvl(a.flo_kg, 0)), 2) flo_kg, " & _
							   " round(SUM(nvl(a.amount_hi, 0)), 2) amount_hi " & _									  
						  " FROM view_attendance_tm_smry a, job b, " & _
								"( select grouphead.fcba, grouphead.param_1, grouphead.param_2, groupname.fcname param_2_name from parameterdetail grouphead, " & _
                                 " (select fcba, fccode, fcname from parameterdetail " & _
									" where " & _
									" parcode = 'SR002' and " & _
									" parheadcode = 'GROUPJOB') groupname " & _
							 " where   " & _
							 " grouphead.parcode = 'SR002' and " & _
							 " grouphead.parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') and " & _
							 " groupname.fcba = grouphead.fcba and " & _
							 " groupname.fccode = grouphead.param_2 )  c " & _
						 " WHERE (a.fcba = b.fcba AND a.jobcode = b.fccode) and " & _ 
                                 " (a.fcba = c.fcba AND a.jobcode = c.param_1)  " & _ 
								 " AND fddate BETWEEN TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY') " & _ 
								 " AND TO_DATE ('" & v_txtdateto & "', 'MM/DD/YYYY') " & _ 
								" GROUP BY c.param_2, c.param_2_name,a.status,a.jobcode ,a.fcba,a.DIVISION,a.PLANTINGYEAR,a.FIELDCODE,a.fcname"
	
                      strEPS2    =   " ,round(SUM(nvl(sku,0)),2) sku, " & _
									 " round(SUM(nvl(khl,0)),2) khl,  " & _
									 " round(SUM(nvl(total,0)),2) total,  " & _
									 " round(SUM(nvl(ha_harian,0)),2) ha_harian, " & _
									 " round(SUM(nvl(bo_harian,0)),2) bo_harian,  " & _
									 " round(SUM(nvl(total_bo_harian,0)),2) total_bo_harian, " & _
									 " ROUND(DECODE(round(SUM(nvl(ha_harian, 0)), 2), " & _
									 " 0, " & _
									 " 0, " & _
									 " round(SUM(nvl(total, 0)), 2) / " & _
									 " round(SUM(nvl(ha_harian, 0)), 2)), " & _
									 " 2) ha_per_hk, " & _
									 "	   round(SUM(nvl(gli_ltr, 0)), 2) gli_ltr, " & _
									 "	   ROUND(DECODE(round(SUM(nvl(total_bo_harian, 0)), 2), " & _
									 "					0, " & _
									 "					0, " & _
									 "					round(SUM(nvl(gli_ltr, 0)), 2) / " & _
									 "					round(SUM(nvl(total_bo_harian, 0)), 2)), " & _
									 "			 2) gli_ltr_ha, " & _
									 "	   round(SUM(nvl(par_ltr, 0)), 2) par_ltr, " & _
									 "	   ROUND(DECODE(round(SUM(nvl(total_bo_harian, 0)), 2), " & _
									 "					0, " & _
									 "					0, " & _
									 "					round(SUM(nvl(par_ltr, 0)), 2) / " & _
									 "					round(SUM(nvl(total_bo_harian, 0)), 2)), " & _
									 "			 2) par_ltr_ha, " & _
									 "	   round(SUM(nvl(met_kg, 0)), 2) met_kg, " & _
									 "	   ROUND(DECODE(round(SUM(nvl(total_bo_harian, 0)), 2), " & _
								     "					0, " & _
									 "					0, " & _
									 "					round(SUM(nvl(met_kg, 0)), 2) / " & _
									 "					round(SUM(nvl(total_bo_harian, 0)), 2)), " & _
									 "			 2) met_kg_ha, " & _
									 "	   round(SUM(nvl(flo_kg, 0)), 2) flo_kg, " & _
									 "	   ROUND(DECODE(round(SUM(nvl(total_bo_harian, 0)), 2), " & _
									 "					0, " & _
									 "					0, " & _
									 "					round(SUM(nvl(flo_kg, 0)), 2) / " & _
									 "					round(SUM(nvl(total_bo_harian, 0)), 2)), " & _
									 "			 2) flo_kg_ha, " & _
									 "	   ROUND(DECODE(round(SUM(nvl(ha_harian, 0)), 2), " & _
									 "					0, " & _
									 "					0, " & _
									 "					round(SUM(nvl(amount_hi, 0)), 2) / " & _
									 "					round(SUM(nvl(ha_harian, 0)), 2)), " & _
									 "			 2) amount_hi_per_ha "
									 
	strPerawatanTM =  "select param_2,param_2_name,GroupPlantation " & selectCons & strEPS2 & "  From (" & strPerawatanTM & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  param_2,param_2_name,GroupPlantation" & selectCons & "  Order by param_2,param_2_name,GroupPlantation" & selectCons  
	
	
	
	
    ' ' Where Condion
	
	' if len(trim(v_txtGroup)) > 0 then 
		' if v_txtGroup <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " AND c.param_2 = '" & v_txtGroup & "'"
		' end if
	' end if
	
	' if len(trim(v_txtStatusTmTbm)) > 0 then 
		' if v_txtStatusTmTbm <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " AND a.STATUS = '" & v_txtStatusTmTbm & "'"
		' end if
	' end if
	
    ' if len(trim(v_txtjobcode)) > 0  then
		' if v_txtjobcode <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " AND a.JOBCODE = '" & v_txtjobcode & "'"
		' end if 
	' end if	

	
	' if len(trim(v_txtDivision)) > 0 then 
		' if  v_txtDivision <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " AND a.DIVISION = '" & v_txtDivision & "'"
		' end if
	' end if
	
	' if len(trim(v_txtPlantingDate)) > 0  then
		' if v_txtPlantingDate <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " AND a.PLANTINGYEAR = '" & v_txtplantingdate & "'"
		' end if 
	' end if
	
	
    ' if len(trim(v_txtBlok)) > 0  then
		' if v_txtBlok <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " AND a.FIELDCODE = '" & v_txtBlok & "'"
		' end if 
	' end if
	


	' ' Group By
	' strPerawatanTM = strPerawatanTM & " GROUP BY c.param_2, c.param_2_name "
	
	' if len(trim(v_txtStatusTmTbm)) > 0 then 
		' if v_txtStatusTmTbm <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , a.status "
		' end if
	' else
	    ' strPerawatanTM = strPerawatanTM & " , a.status "
	' end if	
	
	' if len(trim(v_txtjobcode)) > 0 then 
		' if v_txtjobcode <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , a.jobcode, b.fcname "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , a.jobcode, b.fcname  "	
	' end if	
	
	' if len(trim(v_txtfcba)) > 0 then 
		' if v_txtfcba <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , a.fcba "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , a.fcba "	
	' end if	
	
	' if len(trim(v_txtDivision)) > 0 then 
		' if v_txtDivision <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , a.DIVISION "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , a.DIVISION "
	' end if
	
	' if len(trim(v_txtPlantingDate)) > 0 then 
		' if v_txtPlantingDate <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , a.PLANTINGYEAR "
		' end if
	' else
	    ' strPerawatanTM = strPerawatanTM & " , a.PLANTINGYEAR "
	' end if	

	' if len(trim(v_txtBlok)) > 0 then 
		' if v_txtBlok <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , a.FIELDCODE, a.fcname "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , a.FIELDCODE, a.fcname "	
	' end if


	
	' ' Order By
	' strPerawatanTM = strPerawatanTM & " ORDER BY param_2 "
	
	' if len(trim(v_txtStatusTmTbm)) > 0 then 
		' if v_txtStatusTmTbm <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , status "
		' end if
	' else
	    ' strPerawatanTM = strPerawatanTM & " , status "
	' end if	
	
	' if len(trim(v_txtjobcode)) > 0 then 
		' if v_txtjobcode <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , jobcode "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , jobcode  "	
	' end if	

	' if len(trim(v_txtfcba)) > 0 then 
		' if v_txtfcba <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , fcba "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , fcba "	
	' end if

	
	' if len(trim(v_txtDivision)) > 0 then 
		' if v_txtDivision <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , DIVISION "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , DIVISION "
	' end if
	
	' if len(trim(v_txtPlantingDate)) > 0 then 
		' if v_txtPlantingDate <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , PLANTINGYEAR "
		' end if
	' else
	    ' strPerawatanTM = strPerawatanTM & " , PLANTINGYEAR "
	' end if	

	' if len(trim(v_txtBlok)) > 0 then 
		' if v_txtBlok <> "--SUMMARY--" then
			' strPerawatanTM = strPerawatanTM & " , FIELDCODE "
		' end if
	' else
		' strPerawatanTM = strPerawatanTM & " , FIELDCODE "	
	' end if



	' Make String combination if the Value of Jobcode is ZERO
	' By NAF
	' the fcba, division, plantingdate and block must be in --SUMMARY-- condition
	
	
	' ExtendedQuery = 0
	' if v_txtfcba = "--SUMMARY--" then 
	    ' if v_txtDivision = "--SUMMARY--" then 
			' if v_txtPlantingDate = "--SUMMARY--" then 
			    ' if v_txtJobCode = "--SUMMARY--" then 
					' if v_txtblok = "--SUMMARY--" then 
						' if v_txtStatusTmTbm = "--SUMMARY--" then 
							' ExtendedQuery = 1				
						' end if
					' end if
				' end if
			' end if
		' end if
	' end if
			    
    ' if ExtendedQuery = 1 then 
	    ' strExtendedQuery = " select param_2, param_2_name, " & _
									 ' " round(SUM(nvl(sku,0)),2) sku, " & _
									 ' " round(SUM(nvl(khl,0)),2) khl,  " & _
									 ' " round(SUM(nvl(total,0)),2) total,  " & _
									 ' " round(SUM(nvl(ha_harian,0)),2) ha_harian, " & _
									 ' " round(SUM(nvl(bo_harian,0)),2) bo_harian,  " & _
									 ' " round(SUM(nvl(total_bo_harian,0)),2) total_bo_harian, " & _
									 ' " round(SUM(nvl(ha_per_hk,0)),2) ha_per_hk, " & _
									 ' " round(SUM(nvl(gli_ltr,0)),2) gli_ltr, " & _
									 ' " round(SUM(nvl(gli_ltr_ha,0)),2) gli_ltr_ha, " & _
									 ' " round(SUM(nvl(par_ltr,0)),2) par_ltr, " & _
									 ' " round(SUM(nvl(par_ltr_ha,0)),2) par_ltr_ha, " & _
									 ' " round(SUM(nvl(met_kg,0)),2) met_kg, " & _
									 ' " round(SUM(nvl(met_kg_ha,0)),2) met_kg_ha, " & _
									 ' " round(SUM(nvl(flo_kg,0)),2) flo_kg, " & _
									 ' " round(SUM(nvl(flo_kg_ha,0)),2) flo_kg_ha, " & _
									 ' " round(SUM(nvl(amount_hi,0)),2) amount_hi_per_ha " & _
							' " from ( " & _
							' " select b_temp.param_2, b_temp.param_2_name, " & _
							  ' " 0 sku, 0 khl, 0 total,  " & _
							  ' " 0 ha_harian, 0 bo_harian, 0 total_bo_harian, " & _
							  ' " 0 ha_per_hk, " & _
							  ' " 0 gli_ltr, 0 gli_ltr_ha, " & _
							  ' " 0 par_ltr, 0 par_ltr_ha, " & _
							  ' " 0 met_kg, 0 met_kg_ha, " & _
							  ' " 0 flo_kg, 0 flo_kg_ha, 0 amount_hi from job a_temp, " & _
							' "( select grouphead.fcba, grouphead.param_1, grouphead.param_2, groupname.fcname param_2_name from parameterdetail grouphead, " & _
                                 ' " (select fcba, fccode, fcname from parameterdetail " & _
									' " where " & _
									' " parcode = 'SR002' and " & _
									' " parheadcode = 'GROUPJOB') groupname " & _
							 ' " where   " & _
							 ' " grouphead.parcode = 'SR002' and " & _
							 ' " grouphead.parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') and " & _
							 ' " groupname.fcba = grouphead.fcba and " & _
							 ' " groupname.fccode = grouphead.param_2 ) b_temp " & _
							' " where " & _
							' " a_temp.fccode = b_temp.param_1 " & _
							' " and a_temp.fcba = b_temp.fcba "  
							' '" and  ACCESSRIGHTSCOMPANY_API.IsAccessRights('" & Session("CodeName") & "',a_temp.fcba,'" & txtcompany & "') = 'TRUE' "
		' if len(trim(v_txtStatusTmTbm)) > 0  then
			' if v_txtStatusTmTbm <> "--SUMMARY--" then
				' strExtendedQuery = strExtendedQuery & " AND a_temp.job_forfieldstatus = '" & v_txtStatusTmTbm & "' "
			' end if 
		' end if							
		
		' strExtendedQuery = strExtendedQuery & " UNION " & strPerawatanTM
        ' strExtendedQuery = strExtendedQuery & ") " & _ 
												' " GROUP BY param_2, param_2_name " & _
												' " ORDER BY param_2 " 
												
	
	
		' Set rsPerawatanTM = Server.CreateObject("ADODB.Recordset")
		' 'response.write "SATU " & strExtendedQuery 
		
		' rsPerawatanTM.Open strExtendedQuery, connstr, adOpenStatic 
	' else
		' Set rsPerawatanTM = Server.CreateObject("ADODB.Recordset")
		' 'response.write "DUA " & strPerawatanTM
		
		' rsPerawatanTM.Open strPerawatanTM, connstr, adOpenStatic 	
    ' end if	

	
	
	' End of Modification
	
	Set rsPerawatanTM = Server.CreateObject("ADODB.Recordset")
	'response.write  strPerawatanTM
		
	rsPerawatanTM.Open strPerawatanTM, connstr, adOpenStatic 	
	


	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Maintenance Report</title>
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
</head>
<form name="frmEntry" method="post" action="p_PerawatanTM_BA_NAF.asp">

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	
	</div>
<!--#include file="FormSelectNewSum.html"-->	
<table id="form-table" class="centered-margin">
	<tr>
		<th>DARI</th>
		<th>SAMPAI</th>
		<th>GROUP</th>
		<th>STATUS</th>
		<th>PEKERJAAN</th>
		<th>AFDELING</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
		<th>&nbsp;</th>
	</tr>
	
	<tr>
		<td align="center">
		<% if v_txtdatefrom <> "" and v_txtdatefrom = v_txtdatefrom then %>
			<input type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtdatefrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
		<% if v_txtdateto <> "" and v_txtdateto = v_txtdateto then %>
			<input type="text" maxlength="10" name="txtdateto" size="10" value="<%=v_txtdateto%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtdateto" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtdateto');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		
		
		
		<td align="center">
		<select name="txtGroup">	
			<option value="">--ALL--</option>
		<%
			Dim rsGroup
			Set rsGroup = Server.CreateObject("ADODB.Recordset")
			rsGroup.CursorLocation = adUseClient
			
			
			if len(trim(v_txtStatusTmTbm)) > 0 then 
			    strsql =  "  select distinct fccode, fcname from parameterdetail " & _
							" where " & _
							" parcode = 'SR002' and " & _
							" parheadcode = 'GROUPJOB' " 
							
			else
			    strsql =  "  select distinct fccode, fcname from parameterdetail " & _
							" where " & _
							" parcode = 'SR002' and " & _
							" parheadcode = 'GROUPJOB' " 
							
			end if
			
			
			
			rsGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsGroup.eof then
				DO while not rsGroup.eof
					if v_txtGroup = rsGroup("fccode") then
			%>			<option value="<%=rsGroup("fccode")%>" Selected><%=rsGroup("fccode")%> - <%= rsGroup("fcname") %></option>
					<%else%>
						<option value="<%=rsGroup("fccode")%>"><%=rsGroup("fccode")%> - <%= rsGroup("fcname") %></option>
					<%end if%>
		<%			rsGroup.movenext
				loop
			End if
			rsGroup.Close
			set rsGroup = nothing
		%>
		</select>		
		</td>
		
		
		<td align="center">
		<select name="txtStatusTmTbm">	
			<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsStatusTmTbm
			Set rsStatusTmTbm = Server.CreateObject("ADODB.Recordset")
			rsStatusTmTbm.CursorLocation = adUseClient
			
			
			' if v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "" then 
				' strsql =  "  SELECT '--ALL--' as STATUSTMTBM from DUAL UNION SELECT FCCODE as STATUSTMTBM from PARAMETER where parcode = 'F0006' and FCBA = '" & v_txtFcba & "'"
			' else
				strsql =  "  SELECT '--ALL--' as STATUSTMTBM from DUAL UNION SELECT DISTINCT FCCODE as STATUSTMTBM from PARAMETER where parcode = 'F0006' " 

			' end if
			
			
			rsStatusTmTbm.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStatusTmTbm.eof then
				DO while not rsStatusTmTbm.eof
				if v_txtStatusTmTbm = rsStatusTmTbm("STATUSTMTBM") then
		%>			<option value="<%=rsStatusTmTbm("STATUSTMTBM")%>" Selected><%=rsStatusTmTbm("STATUSTMTBM")%></option>
				<%else%>
					<option value="<%=rsStatusTmTbm("STATUSTMTBM")%>"><%=rsStatusTmTbm("STATUSTMTBM")%></option>
				<%end if%>
		<%			rsStatusTmTbm.movenext
				loop
			End if
			rsStatusTmTbm.Close
			set rsStatusTmTbm = nothing
		%>
		</select>		
		</td>		
		
		
		<td>
				<select name="txtjobcode">	
				<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsJob
			Set rsJob = Server.CreateObject("ADODB.Recordset")
			rsJob.CursorLocation = adUseClient
			
			' Get the Activity from ParameterDetail Table
			
			if len(trim(v_txtBlok)) > 0 then 
			    strsql = " SELECT '--ALL--' as fccode, '' as fcname  from DUAL UNION SELECT FCCODE, FCNAME FROM Job WHERE FCCODE IN ( " & _
						 " select param_1 from parameterdetail " & _
						 " where " & _
						 " fcba = '" & v_txtFcba & "' and " & _
						 " parcode = 'SR002' and " & _
						 " parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 " ) AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC " 
			
			else
			    strsql = " SELECT '--ALL--' as fccode, '' as fcname  from DUAL UNION SELECT FCCODE, FCNAME FROM Job WHERE FCCODE IN ( " & _
						 " select param_1 from parameterdetail " & _
						 " where " & _
						 " fcba = '" & Session("fcBA") & "' and " & _
						 " parcode = 'SR002' and " & _
						 " parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 " ) AND FCBA = '" & Session("fcBA") & "' ORDER BY FCCODE ASC " 

			end if		
			
			rsJob.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsJob.eof then
				DO while not rsJob.eof
				if v_txtjobcode = rsJob("FCCODE") then
		%>			<option value="<%=rsJob("FCCODE")%>" Selected><%=rsJob("FCCODE")%> - <%=rsJob("FCNAME")%></option>
				<%else%>
					<option value="<%=rsJob("FCCODE")%>"><%=rsJob("FCCODE")%> - <%=rsJob("FCNAME")%></option>
				<%end if%>
		<%			rsJob.movenext
				loop
			End if
			rsJob.Close
			set rsJob = nothing
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
			

			if len(trim(v_txtStatusTmTbm)) > 0 then 
				strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & _
				v_txtFcba & "'"
			else
				strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & _
				Session("fcBA") & "'"
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

<div id="header-wrapper">
		<div id="header">
		 <% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
		 <% End if%>
		 <% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
		 <% End if%>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN HARIAN MERUMPUT</h3>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>
	</div>
</div>

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP</th>
		<% if v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "--SUMMARY--" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" >STATUS</th>
		<% end if %>

		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" >PLANTATION</th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" >GROUP REGION</th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" >REGION</th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;" >COMPANY</th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;"  >SITE</th>
		<% end if %>

		
		<% if v_txtDivision <> "--SUMMARY--" AND v_txtDivision <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;"  >AFDELING</th>
		<% end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" AND v_txtPlantingDate <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;"  >TT</th>
		<% end if %>


		<% if v_txtBlok <> "--SUMMARY--" AND v_txtBlok <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;"  >BLOK</th>
		<% end if %>
		

		
		
		
		<th colspan="3" style="background-color:#00A300;color:#fff;" >TENAGA KERJA</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >HASIL KERJA</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;" >RASIO (HK/HA)</th>
		<th colspan="8" style="background-color:#00A300;color:#fff;" >PEMAKAIAN BAHAN</th>
		<th style="background-color:#00A300;color:#fff;">AMOUNT</th>
	</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >SKU</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >KHL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >TOTAL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >HARIAN</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >BORONGAN</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >TOTAL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >AKTUAL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >BUDGET</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;" >GLIFOSAT(LTR/HA)</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;" >PARAQUAT(LTR/HA)</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;" >METIL(LTR/HA)</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;" >FLOROXIPIR(LTR/HA)</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >RP/HA</th>		
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" >AKTUAL</th>
		<th style="background-color:#00A300;color:#fff;" >BUDGET</th>
		<th style="background-color:#00A300;color:#fff;" >AKTUAL</th>
		<th style="background-color:#00A300;color:#fff;" >BUDGET</th>
		<th style="background-color:#00A300;color:#fff;" >AKTUAL</th>
		<th style="background-color:#00A300;color:#fff;" >BUDGET</th>
		<th style="background-color:#00A300;color:#fff;" >AKTUAL</th>
		<th style="background-color:#00A300;color:#fff;" >BUDGET</th>
	</tr>
	
<%
Dim nsubsku, nsubkhl, nsubtotal, nsubha_harian, nsubbo_harian
Dim nsubtotal_bo_harian, nsubha_per_hk, nsubgli_ltr, nsubgli_ltr_ha, nsubpar_ltr
Dim nsubpar_ltr_ha, nsubmet_kg, nsubmet_kg_ha, nsubflo_kg, nsubflo_kg_ha, nsubamount_hi_per_ha
	if not rsPerawatanTM.EOF then
		NO = 1
		nsubsku = 0
		nsubkhl = 0
		nsubtotal = 0
		nsubha_harian = 0
		nsubbo_harian = 0
		nsubtotal_bo_harian = 0
		nsubha_per_hk = 0
		nsubgli_ltr = 0
		nsubgli_ltr_ha = 0
		nsubpar_ltr = 0
		nsubpar_ltr_ha = 0
		nsubmet_kg = 0
		nsubmet_kg_ha = 0
		nsubflo_kg = 0
		nsubflo_kg_ha = 0
		
		nsubamount_hi_per_ha = 0

		while not rsPerawatanTM.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				
				
				<% if v_txtGroup <> "--SUMMARY--"   then %>	
				        <% tempGroup = rsPerawatanTM("param_2") 
						   tempStatus = "--ALL--" 
						   tempJobcode = "--SUMMARY--" 
						   tempPlantation = "--SUMMARY--" 
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempDivision = "--SUMMARY--"
						   tempPlantingDate = "--SUMMARY--"
						   tempBlok = "--SUMMARY--"
						  %> 
				
						<% if sprevGroup <> rsPerawatanTM("param_2") then %>
							<td align="left">
								<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
								<%=rsPerawatanTM("param_2_name")%>
								</a>
							</td>
							<% sprevGroup = rsPerawatanTM("param_2") %>	
						<% else	%>
							<td align="left"></td>					
						<% end if  
				     end if 
				  %>	
				
				
				<% if v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "" then %>	
                     <%    tempGroup = rsPerawatanTM("param_2") 
						   tempStatus = rsPerawatanTM("status")
						   tempJobcode = "--SUMMARY--" 
						   tempPlantation = "--ALL--" 
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempDivision = "--SUMMARY--"
						   tempPlantingDate = "--SUMMARY--"
						   tempBlok = "--SUMMARY--"
						  %> 				
					<% if sprevStatus <> rsPerawatanTM("status") then %>

						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsPerawatanTM("status")%> 
							</a>
						</td>
						<% sprevStatus = rsPerawatanTM("status") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>	
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     tempGroup = rsPerawatanTM("param_2") 
					       if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then 
						      tempStatus = rsPerawatanTM("status")
						   else
						      tempStatus =""
						   end if
						   tempJobcode = "--SUMMARY--" 
						   tempPlantation = rsPerawatanTM("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempDivision = "--SUMMARY--"
						   tempPlantingDate = "--SUMMARY--"
						   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsPerawatanTM("Plantation") then %>
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsPerawatanTM("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsPerawatanTM("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  tempGroup = rsPerawatanTM("param_2") 
						tempStatus = rsPerawatanTM("status")
						tempJobcode = "--SUMMARY--" 
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsPerawatanTM("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsPerawatanTM("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsPerawatanTM("GroupRegion") then %>
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsPerawatanTM("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsPerawatanTM("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  tempGroup = rsPerawatanTM("param_2") 
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then 
						      tempStatus = rsPerawatanTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPerawatanTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPerawatanTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsPerawatanTM("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsPerawatanTM("Region") then %>
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsPerawatanTM("region")%>
							</a>
						</td>
						<% SprevRegion = rsPerawatanTM("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					    tempGroup = rsPerawatanTM("param_2") 
					    
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then 
						      tempStatus = rsPerawatanTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPerawatanTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPerawatanTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPerawatanTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsPerawatanTM("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsPerawatanTM("Company") then %>
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsPerawatanTM("company")%>
							</a>
						</td>
						<% SprevCompany = rsPerawatanTM("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <%  tempGroup = rsPerawatanTM("param_2") 
					    
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then 
						      tempStatus = rsPerawatanTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPerawatanTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPerawatanTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPerawatanTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsPerawatanTM("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsPerawatanTM("fcba")
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsPerawatanTM("fcba") then %>
						<td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
						    <%=rsPerawatanTM("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsPerawatanTM("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtDivision <> "--SUMMARY--" then %>
				   <%   tempGroup = rsPerawatanTM("param_2") 
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then 
						      tempStatus = rsPerawatanTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--"  
				       if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPerawatanTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPerawatanTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPerawatanTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsPerawatanTM("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsPerawatanTM("fcba")
					   else
					       tempfcba = ""
					   end if
					   tempDivision = rsPerawatanTM("division")
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevDivision <> rsPerawatanTM("division") then %>
					    <td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsPerawatanTM("division")%>
							</a>
						</td>
						<% sprevDivision = rsPerawatanTM("division") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>

				<% if v_txtPlantingDate <> "--SUMMARY--" then %>
				   <%   tempGroup = rsPerawatanTM("param_2") 
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then 
						      tempStatus = rsPerawatanTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
				       if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPerawatanTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPerawatanTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPerawatanTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsPerawatanTM("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsPerawatanTM("fcba")
					   else
					       tempfcba = ""
					   end if
					   tempDivision = rsPerawatanTM("division")
					   tempPlantingDate = rsPerawatanTM("plantingyear")
					   tempBlok = "--ALL--"					   
					 %>
					
					
					
				    <% if sprevPlantingYear <> rsPerawatanTM("plantingyear") then %>
				        <td align="center">
							<a href="p_PerawatanTM_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsPerawatanTM("plantingyear")%>
							</a>
						</td>
						<% sprevPlantingYear = rsPerawatanTM("plantingyear") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
                <% if v_txtBlok <> "--SUMMARY--" then %>
				    <% if sprevFieldCode <> rsPerawatanTM("fieldcode") then %>				
					    <td align="center"><%=rsPerawatanTM("fieldcode")%></td>				
						<% sprevFieldCode = rsPerawatanTM("fieldcode") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>	

				
				
				<td align="right"><%=FormatNumber(rsPerawatanTM("sku"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("khl"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("total"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("ha_harian"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("bo_harian"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("total_bo_harian"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("ha_per_hk"),2)%></td>
				<td align="right">0</td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("gli_ltr"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("gli_ltr_ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("par_ltr"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("par_ltr_ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("met_kg"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("met_kg_ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("flo_kg"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("flo_kg_ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsPerawatanTM("amount_hi_per_ha"),2)%></td>
			</tr>
<%
			NO = NO + 1			
			nsubsku = nsubsku + Cdbl(rsPerawatanTM("sku"))
			nsubkhl = nsubkhl + Cdbl(rsPerawatanTM("khl"))
			nsubtotal = nsubtotal + Cdbl(rsPerawatanTM("total"))
			nsubha_harian = nsubha_harian + Cdbl(rsPerawatanTM("ha_harian"))
			nsubbo_harian = nsubbo_harian + Cdbl(rsPerawatanTM("bo_harian"))
			nsubtotal_bo_harian = nsubtotal_bo_harian + Cdbl(rsPerawatanTM("total_bo_harian"))
			nsubha_per_hk = nsubha_per_hk + Cdbl(rsPerawatanTM("ha_per_hk"))
			nsubgli_ltr = nsubgli_ltr + Cdbl(rsPerawatanTM("gli_ltr"))
			nsubgli_ltr_ha = nsubgli_ltr_ha + Cdbl(rsPerawatanTM("gli_ltr_ha"))
			nsubpar_ltr = nsubpar_ltr + Cdbl(rsPerawatanTM("par_ltr"))
			nsubpar_ltr_ha = nsubpar_ltr_ha + Cdbl(rsPerawatanTM("par_ltr_ha"))
			nsubmet_kg = nsubmet_kg + Cdbl(rsPerawatanTM("met_kg"))
			nsubmet_kg_ha = nsubmet_kg_ha + Cdbl(rsPerawatanTM("met_kg_ha"))
			nsubflo_kg = nsubflo_kg + Cdbl(rsPerawatanTM("flo_kg"))
			nsubflo_kg_ha = nsubflo_kg_ha + Cdbl(rsPerawatanTM("flo_kg_ha"))
			
			nsubamount_hi_per_ha = nsubamount_hi_per_ha + cdbl(rsPerawatanTM("amount_hi_per_ha"))

			rsPerawatanTM.MoveNext
		Wend
	end if
	rsPerawatanTM.Close
	set rsPerawatanTM = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 1 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nsubsku,2)%></td>
				<td align="right"><%=FormatNumber(nsubkhl,2)%></td>
				<td align="right"><%=FormatNumber(nsubtotal,2)%></td>
				<td align="right"><%=FormatNumber(nsubha_harian,2)%></td>
				<td align="right"><%=FormatNumber(nsubbo_harian,2)%></td>
				<td align="right"><%=FormatNumber(nsubtotal_bo_harian,2)%></td>
				<% if cdbl(nsubha_harian) > 0 then %>
				   <td align="right"><%=FormatNumber(nsubtotal/nsubha_harian,2)%></td>
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right">0</td>
				<td align="right"><%=FormatNumber(nsubgli_ltr,2)%></td>
				<td align="right"><%=FormatNumber(nsubgli_ltr_ha,2)%></td>
				<td align="right"><%=FormatNumber(nsubpar_ltr,2)%></td>
				<td align="right"><%=FormatNumber(nsubpar_ltr_ha,2)%></td>
				<td align="right"><%=FormatNumber(nsubmet_kg,2)%></td>
				<td align="right"><%=FormatNumber(nsubmet_kg_ha,2)%></td>
				<td align="right"><%=FormatNumber(nsubflo_kg,2)%></td>
				<td align="right"><%=FormatNumber(nsubflo_kg_ha,2)%></td>
				<td align="right"><%=FormatNumber(nsubamount_hi_per_ha,2)%></td>
			</tr>
</table>

</form>


