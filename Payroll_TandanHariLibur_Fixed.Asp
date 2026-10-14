<%
	'Amandemen Premi Tandan Hari Libur
	'1. Code Allowance Panen = A25
	'   Qty Basis = Rate1 
	'   Premi Pertandan = OpHarvRate
	'   Unit Quantity didapat dari attendance_gad untuk KK (Kerja Kontanan)
	'   Type_1 = Alokasi daily(D) / Monthly (M)
	'   Type_2 = Kerja Kontanan (KK) / Non Kontanan (NK)
	'2. Code D20 premi potongan tandan hari libur
	'   tempIncludePayrollProcessingBasisKontanan 
	
	
  
	Dim vTerrainType_kk : vTerrainType_kk = ""
	Dim nRatePremiHariLibur
	
    ' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_TandanHariLibur_Fixed.asp"
	tempRefCodeReceive = "ATTENDANCE_GAD"

	qReference =  " select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
				  " from process_tableref_tab " & _
				  " where " & _
				  " processcode = '" & tempProcessCode & "' and " & _
				  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
				  " fccode =  '" & tempRefCodeReceive & "'"

	Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
	rsReferenceReceive.CursorType = 3
	rsReferenceReceive.CursorLocation = 3
	rsReferenceReceive.Open qReference, connstr

	doneReceive = 1
	if not rsReferenceReceive.eof then
		txtRefNumberKeys_Receive = rsReferenceReceive("REF_NUMBERKEYS")
		txtRefStringSeparator_Receive = rsReferenceReceive("REF_STRINGSEPARATOR")
		txtRefStringKeys_Receive = rsReferenceReceive("REF_STRINGKEYS")
		doneReceive = 1
	else
		doneReceive = 0
	end if
	rsReferenceReceive.close
	set rsReferenceReceive = Nothing
	' End of Picking Referential Integrity

	
	
	
    ' Cari RateBuah Mentah
	' Type_1 ==> nPaymentType ==> Daily / Monthly

	qBA="select nvl(Rate_1,0) as Rate_1, type_2,INCLUDEDPAYROLLPROCESSING from ALLOWDED where FCCODE='A25' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nRatePremiHariLibur=cdbl(rsBA("Rate_1"))
	   'nPaymentType ==> Daily / Monthly
	   nPaymentType = rsBA("type_2")
	   tempIncludePayrollProcessingBasis = rsBA("INCLUDEDPAYROLLPROCESSING")
	else
	   nRatePremiHariLibur=0
	   nPaymentType = "NK"
	   tempIncludePayrollProcessingBasis = "FALSE"
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Tandan Hari Libur is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	
	qBA="select INCLUDEDPAYROLLPROCESSING from ALLOWDED where FCCODE='D20' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	  
	   tempIncludePayrollProcessingBasisKontanan = rsBA("INCLUDEDPAYROLLPROCESSING")
	else
	  
	   tempIncludePayrollProcessingBasisKontanan = "FALSE"
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Potongan Tandan Hari Libur  is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

    if tempIncludePayrollProcessingBasis = "TRUE" then 
	
		' qReceive=" Select  " & _
				 ' " hv.FCBA,  " & _
				 ' " hv.FDDATE, " & _
				 ' " hv.EMPLOYEECODE, " & _
				 ' " hv.UNIT, " & _
				 ' " hv.PRESENTRATE, " & _
				 ' " hv.TOTALAMOUNT, " & _
				 ' " hv.A01, " & _
				 ' " hv.A25 " & _
				 ' " from " & _
				 ' " harvesting_premi hv,Employee Emp " & _
				 ' " where (hv.EMPLOYEECODE = Emp.FCCODE and hv.FCBA = Emp.FCBA)  " & _
				 ' " AND Emp.DATETERMINATE is null and hv.ATTENDANCE='KK' " & _
				 ' " AND hv.FCBA = '" & Session("fcBA") & "' " & _
				 ' " and hv.UNIT > 0 " & _
				 ' " AND TO_CHAR(TO_DATE(hv.FDDATE,'dd/mm/yyyy'),'fmMM') = " & Month(Session("v_Employee")) & _
				 ' " AND TO_CHAR(TO_DATE(hv.FDDATE,'dd/mm/yyyy'),'yyyy') = " & Year(Session("v_Employee"))
		
		qReceive = "   SELECT fcba," & _
						"          fddate," & _
						"          employeecode," & _
						"          employee_api.get_employeename (fcba, employeecode) employeename," & _
						"          employee_api.get_sectionname (fcba, employeecode) division," & _
						"          total_tbs," & _
						"          kg_tbs_extra_harian," & _
						"          premi_lebih_borong_tbs_harian," & _
						"          total_bucket," & _
						"          kg_brondolan_harian," & _
						"          premi_brondolan_harian," & _
						"          total_premi_kontanan" & _
						"     FROM (  SELECT premi_kontanan.fcba," & _
						"                    premi_kontanan.fddate," & _
						"                    premi_kontanan.employeecode," & _
						"                    ROUND (SUM (premi_kontanan.total_tbs), 0) AS total_tbs," & _
						"                    ROUND (SUM (premi_kontanan.kg_tbs_extra_harian), 0)" & _
						"                       AS kg_tbs_extra_harian," & _
						"                    ROUND (SUM (premi_kontanan.premi_lebih_borong_tbs_harian), 2)" & _
						"                       AS premi_lebih_borong_tbs_harian," & _
						"                    ROUND (SUM (premi_kontanan.total_bucket), 0) AS total_bucket," & _
						"                    ROUND (SUM (premi_kontanan.kg_brondolan_harian), 0)" & _
						"                       AS kg_brondolan_harian," & _
						"                    ROUND (SUM (premi_kontanan.premi_brondolan_harian), 2)" & _
						"                       AS premi_brondolan_harian," & _
						"                    ROUND (" & _
						"                       NVL (" & _
						"                          (  SUM (premi_kontanan.premi_lebih_borong_tbs_harian)" & _
						"                           + SUM (premi_kontanan.premi_brondolan_harian))," & _
						"                          0)," & _
						"                       2)" & _
						"                       AS total_premi_kontanan" & _
						"               FROM (SELECT bkm.fcba," & _
						"                            bkm.fddate," & _
						"                            bkm.employeecode," & _
						"                            hdb.division AS division_blok," & _
						"                            bkm.locationcode," & _
						"                            NVL (bkm.total_tbs, 0) AS total_tbs," & _
						"                            NVL (hdb.kg_tbs_extra_harian, 0)" & _
						"                               AS kg_tbs_extra_harian," & _
						"                            NVL (hdb.premi_lebih_borong_tbs_harian, 0)" & _
						"                               AS premi_lebih_borong_tbs_harian," & _
						"                            NVL (bkm.total_bucket, 0) AS total_bucket," & _
						"                            NVL (hmwn_daily.bbr, 0) AS bbr," & _
						"                            NVL ( (bkm.total_bucket * hmwn_daily.bbr), 0)" & _
						"                               AS kg_brondolan_harian," & _
						"                            NVL (mhr.rate_brondolan, 0) AS rate_brondolan," & _
						"                            NVL (" & _
						"                               (  NVL ( (bkm.total_bucket * hmwn_daily.bbr), 0)" & _
						"                                * NVL (mhr.rate_brondolan, 0))," & _
						"                               0)" & _
						"                               AS premi_brondolan_harian" & _
						"                       FROM (  SELECT iplasprod.attendance_gad.fcba," & _
						"                                      iplasprod.attendance_gad.fddate," & _
						"                                      iplasprod.attendance_gad.employeecode," & _
						"                                      iplasprod.attendance_gad.locationcode," & _
						"                                      SUM (iplasprod.attendance_gad.unit)" & _
						"                                         AS total_tbs," & _
						"                                      SUM (iplasprod.attendance_gad.bucket)" & _
						"                                         AS total_bucket" & _
						"                                 FROM iplasprod.attendance_gad" & _
						"                                WHERE " & _
						"										iplasprod.attendance_gad.fcba  = '"& session("fcba") &"' " & _
						"           							and to_number(to_char(iplasprod.attendance_gad.fddate,'mm')) = " & Month(Session("v_Employee")) & "" & _
						"           							and to_number(to_char(iplasprod.attendance_gad.fddate,'yyyy'))= " & Year(Session("v_Employee")) & ""	& _			
						"										and iplasprod.attendance_gad.jobcode IN ('505030101', '112030101')" & _
						"                                       and iplasprod.attendance_gad.attendance = 'KK'" & _
						"                             GROUP BY iplasprod.attendance_gad.fcba," & _
						"                                      iplasprod.attendance_gad.fddate," & _
						"                                      iplasprod.attendance_gad.employeecode," & _
						"                                      iplasprod.attendance_gad.locationcode) bkm" & _
						"                            INNER JOIN" & _
						"                            (  SELECT iplasprod.harvesting_daily_bjr.fcba," & _
						"                                      iplasprod.harvesting_daily_bjr.fddate," & _
						"                                      iplasprod.harvesting_daily_bjr.employeecode," & _
						"                                      iplasprod.harvesting_daily_bjr.division," & _
						"                                      iplasprod.harvesting_daily_bjr.plantingyear," & _
						"                                      iplasprod.harvesting_daily_bjr.presentgrade," & _
						"                                      iplasprod.harvesting_daily_bjr.attendance," & _
						"                                      iplasprod.harvesting_daily_bjr.fieldcode," & _
						"                                      SUM (" & _
						"                                         iplasprod.harvesting_daily_bjr.extraquantity_bydaily)" & _
						"                                         AS kg_tbs_extra_harian," & _
						"                                      SUM (" & _
						"                                         iplasprod.harvesting_daily_bjr.extraamount_bydaily)" & _
						"                                         AS premi_lebih_borong_tbs_harian" & _
						"                                 FROM iplasprod.harvesting_daily_bjr" & _
						"                                WHERE " & _
						"									   iplasprod.harvesting_daily_bjr.fcba = '"& session("fcba") &"'" & _
						"                                      AND iplasprod.harvesting_daily_bjr.attendance ='KK'" & _
						"                             GROUP BY iplasprod.harvesting_daily_bjr.fcba," & _
						"                                      iplasprod.harvesting_daily_bjr.fddate," & _
						"                                      iplasprod.harvesting_daily_bjr.employeecode," & _
						"                                      iplasprod.harvesting_daily_bjr.division," & _
						"                                      iplasprod.harvesting_daily_bjr.plantingyear," & _
						"                                      iplasprod.harvesting_daily_bjr.presentgrade," & _
						"                                      iplasprod.harvesting_daily_bjr.attendance," & _
						"                                      iplasprod.harvesting_daily_bjr.fieldcode)" & _
						"                            hdb" & _
						"                               ON (    bkm.fcba = hdb.fcba" & _
						"                                   AND TO_NUMBER (TO_CHAR (bkm.fddate, 'mmyyyy')) =" & _
						"                                          TO_NUMBER (" & _
						"                                             TO_CHAR (" & _
						"                                                TO_DATE (hdb.fddate, 'mm/dd/yyyy')," & _
						"                                                'mmyyyy'))" & _
						"                                   AND bkm.employeecode = hdb.employeecode" & _
						"                                   AND bkm.locationcode = hdb.fieldcode)" & _
						"                            INNER JOIN" & _
						"                            (SELECT iplasprod.harvesting_mill_weight_netto.fcba_estate," & _
						"                                    iplasprod.harvesting_mill_weight_netto.harvestdate," & _
						"                                    iplasprod.harvesting_mill_weight_netto.fieldcode," & _
						"                                    iplasprod.harvesting_mill_weight_netto.bbr" & _
						"                               FROM iplasprod.harvesting_mill_weight_netto" & _
						"								WHERE " & _
						"									 iplasprod.harvesting_mill_weight_netto.FCBA_ESTATE = '"& session("fcba") &"' " & _
						"                            )hmwn_daily" & _
						"                               ON (    bkm.fcba = hmwn_daily.fcba_estate" & _
						"                                   AND bkm.fddate = hmwn_daily.harvestdate" & _
						"                                   AND bkm.locationcode = hmwn_daily.fieldcode)" & _
						"                            INNER JOIN" & _
						"                            (SELECT distinct iplasprod.multilevel_harvesting_rate.fcba," & _
						"                                    iplasprod.multilevel_harvesting_rate.division," & _
						"                                    iplasprod.multilevel_harvesting_rate.planting_year," & _
						"                                    iplasprod.multilevel_harvesting_rate.employeetype," & _
						"                                    iplasprod.multilevel_harvesting_rate.attendance," & _
						"                                    iplasprod.multilevel_harvesting_rate.rate_brondolan" & _
						"                               FROM iplasprod.multilevel_harvesting_rate" & _
						"                              WHERE " & _
						"										iplasprod.multilevel_harvesting_rate.fcba  = '"& session("fcba") &"' " & _						
						"									AND iplasprod.multilevel_harvesting_rate.attendance =" & _
						"                                       'KK') mhr" & _
						"                               ON (    hdb.fcba = mhr.fcba" & _
						"                                   AND hdb.division = mhr.division" & _
						"                                   AND hdb.plantingyear = mhr.planting_year" & _
						"                                   AND hdb.presentgrade = mhr.employeetype" & _
						"                                   AND hdb.attendance = mhr.attendance))" & _
						"                    premi_kontanan" & _
						"           GROUP BY premi_kontanan.fcba," & _
						"                    premi_kontanan.fddate," & _
						"                    premi_kontanan.employeecode)" & _
						"    WHERE     fcba = '"& session("fcba") &"'" & _
						"           and to_number(to_char(fddate,'mm')) = " & Month(Session("v_Employee")) & "" & _
						"           and to_number(to_char(fddate,'yyyy'))= " & Year(Session("v_Employee")) & "" 					
		
		' response.Write qReceive
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Premi Tandan Hari Libur ..." & "<br>" 
		response.Flush()
		rsEmpAllDed.Open qReceive, connstr
		if not rsEmpAllDed.eof then
		   Do While Not rsEmpAllDed.eof

					maxLineNo = maxLineNo + 1
					qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A25" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & CLng(rsEmpAllDed("total_premi_kontanan")) & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
						 
					connstr.execute(qAllowance)
			  			  
						'detailLineNo = 0

							'	Insert the History of EmpallDed Data 
							' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
							
						'	txtRefStringValues_Receive = rsEmpAllDed("FDDATE") & txtRefStringSeparator_Receive & _
						'								 rsEmpAllDed("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
						'								 Session("fcBA")					 
						'	
						'	detailLineNo = detailLineNo + 1
						'	qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
						'										" LINEHEADERNO," & _
						'										" LINENO," & _
						'										" FCBA," & _
						'										" REF_PROCESS," & _
						'										" REF_PROCESSNAME," & _
						'										" REF_TABLESOURCE," & _
						'										" REF_NUMBERKEYS," & _
						'										" REF_STRINGSEPARATOR," & _
						'										" REF_STRINGKEYS," & _
						'										" REF_STRINGVALUES," & _
						'										" UNIT, " & _
						'										" RATE, " & _
						'										" AMOUNT, " & _
						'										" FCENTRY," & _
						'										" FCEDIT," & _
						'										" FCIP," & _
						'										" LASTUPDATE," & _
						'										" LASTTIME)" & _
						'								" VALUES ( " & _
						'										" '" & maxLineNo & "'," & _
						'										" '" & detailLineNo & "'," & _
						'										" '" & Session("fcBA") & "'," & _
						'										" '" & tempProcessCode & "'," & _
						'										" '" & tempProcessNameCode & "'," & _
						'										" '" & tempRefCodeReceive & "'," & _
						'										" '" & txtRefNumberKeys_Receive & "'," & _
						'										" '" & txtRefStringSeparator_Receive & "'," & _
						'										" '" & txtRefStringKeys_Receive & "'," & _
						'										" '" & txtRefStringValues_Receive & "',0,0," & _
						'										rsEmpAllDed("total_premi_kontanan") & "," & _ 
						'										" '" & Session("CodeName") & "'," & _
						'										" '" & Session("CodeName") & "'," & _
						'										" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
						'										" sysdate, " & _
						'										" to_char(sysdate, 'HH24:MI') " & _
						'										" ) "					
						'	connstr.execute(qAllowanceHistory)
										
					' BELOM SELESAI..... SILAHKAN Dilanjutkan
					' if nPaymentType = "KK" then 
					    ' if tempIncludePayrollProcessingBasisKontanan = "TRUE" Then
							' maxLineNo = maxLineNo + 1							
										
							' qAllowance= "Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
										' "(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D20" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','dd/mm/yyyy') " & "," & -1 * CLng(rsEmpAllDed("A25"))  & _
										' ",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
										' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"	 
							' connstr.execute(qAllowance)
							
								' detailLineNo = 0

										' Insert the History of EmpallDed Data 
									' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
									
									' txtRefStringValues_Receive = rsEmpAllDed("FDDATE") & txtRefStringSeparator_Receive & _
																 ' rsEmpAllDed("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
																 ' Session("fcBA")
																 
									' detailLineNo = detailLineNo + 1
									' qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
																		' " LINEHEADERNO," & _
																		' " LINENO," & _
																		' " FCBA," & _
																		' " REF_PROCESS," & _
																		' " REF_PROCESSNAME," & _
																		' " REF_TABLESOURCE," & _
																		' " REF_NUMBERKEYS," & _
																		' " REF_STRINGSEPARATOR," & _
																		' " REF_STRINGKEYS," & _
																		' " REF_STRINGVALUES," & _
																		' " UNIT, " & _
																		' " RATE, " & _
																		' " AMOUNT, " & _
																		' " FCENTRY," & _
																		' " FCEDIT," & _
																		' " FCIP," & _
																		' " LASTUPDATE," & _
																		' " LASTTIME)" & _
																' " VALUES ( " & _
																		' " '" & maxLineNo & "'," & _
																		' " '" & detailLineNo & "'," & _
																		' " '" & Session("fcBA") & "'," & _
																		' " '" & tempProcessCode & "'," & _
																		' " '" & tempProcessNameCode & "'," & _
																		' " '" & tempRefCodeReceive & "'," & _
																		' " '" & txtRefNumberKeys_Receive & "'," & _
																		' " '" & txtRefStringSeparator_Receive & "'," & _
																		' " '" & txtRefStringKeys_Receive & "'," & _
																		' " '" & txtRefStringValues_Receive & "'," & _
																		' rsEmpAllDed("UNIT") & "," & _
																		' rsEmpAllDed("A25") & "," & _
																		' rsEmpAllDed("TOTALAMOUNT") & "," & _  
																		' " '" & Session("CodeName") & "'," & _
																		' " '" & Session("CodeName") & "'," & _
																		' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
																		' " sysdate, " & _
																		' " to_char(sysdate, 'HH24:MI') " & _
																		' " ) "					
									' connstr.execute(qAllowanceHistory)
					    ' end if
					' end if
			    rsEmpAllDed.MoveNext
		    loop
		End if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
	end if
%>