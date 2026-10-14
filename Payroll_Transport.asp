<%
	'Amandemen
	'1. Allowded Code Untuk Premi Basis Transport = A32
	'2. Allowded Code Untuk Premi Operator = A33
	'3. Allowded Code Untuk Premi Basis Lebih Transport = A32A	/ A64
	'4. Allowded Code Untuk Premi Basis Lebih Transport = A32B  / A62
	'5. Allowded Code Untuk Premi Basis Lebih Transport = A32C  / A63
	
	'By Iwa

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Transport.asp"
	tempRefCodeReceive = "ATTENDANCE_GAD"

	qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
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
	
	' get allowded A32
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select nvl(RATE_1,0) as Rate_1, includedpayrollprocessing " & _
				" From ALLOWDED where FCCODE='A32'  " & _
				" and man_auto = 'A' " & _
				" and fcBA='" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		' nRatePremi will be inputted by 100 (percent), if 100%
		nRatePremiBasis= cdbl(rsParameter("Rate_1")) 
		tempIncludePayrollProcessingBasis = rsParameter("includedpayrollprocessing")
	else
		nRatePremiBasis = 0 
		tempIncludePayrollProcessingBasis = "FALSE"		 
	End if
	rsParameter.Close
	set rsParameter = Nothing
	' end get allowded A32
	
	' get allowded A33
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select includedpayrollprocessing " & _
				" From ALLOWDED where FCCODE='A33'  " & _
				" and man_auto = 'A' " & _
				" and fcBA='" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		tempIncludePayrollProcessingOpr = rsParameter("includedpayrollprocessing")
	else
		nRatePremiBasis = 0 
		tempIncludePayrollProcessingOpr = "FALSE"		 
	End if
	rsParameter.Close
	set rsParameter = Nothing
	' end get allowded A33
	
	' get allowded A64
    Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select nvl(RATE_1,0) as Rate_1, includedpayrollprocessing " & _
				" From ALLOWDED where FCCODE='A64'  " & _
				" and man_auto = 'A' " & _
				" and fcBA='" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		' nRatePremi will be inputted by 100 (percent), if 100%
		nRatePremiLebihBasis= cdbl(rsParameter("Rate_1")) 
		tempIncludePayrollProcessingLebihBasis = rsParameter("includedpayrollprocessing")
	else
		nRatePremiLebihBasis = 0 
		tempIncludePayrollProcessingLebihBasis = "FALSE"		 
	End if
	rsParameter.Close
	set rsParameter = Nothing
	' endget allowded A64
	
	' get allowded A62
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select TYPE_2 as type2, includedpayrollprocessing " & _
				" From ALLOWDED where FCCODE='A62'  " & _
				" and man_auto = 'A' " & _
				" and fcBA='" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		nBasisPremiType= rsParameter("type2") 
		tempIncludePayrollProcessingByDiv = rsParameter("includedpayrollprocessing")
	else
		tempIncludePayrollProcessingByDiv = "FALSE"
	End if
	rsParameter.Close
	set rsParameter = Nothing
	'end get allowded A62
	
	if tempIncludePayrollProcessingBasis = "TRUE" then 
	 
		qReceive = " 	SELECT " & _
				   "			   fcba," & _
				   " 			   fddate," & _
				   "			   employeecode," & _
				   "			   firstreference reference," & _
				   " 			   SUM (totalUnitPerDay) totalUnitPerDay" & _
				   "			FROM" & _
				   " 			   (	SELECT " & _
				   " 					   fcba, " & _
				   "					   fddate, " & _
				   "					   employeecode, " & _
				   " 					   reference, " & _
				   " 					   oppremirate_api.getFirstReference (fcba, fddate, employeecode,'A32') firstReference, " & _
				   "					   oppremirate_api.convertQty (fcba, " & _
			 	   "												   fddate, " & _
				   " 												   employeecode, " & _
				   "												   reference, " & _
				   "												   'A32','TRANSPORT', " & _
				   "												   TotalUnitPerDay) " & _
				   "						   AS TotalUnitPerDay " & _
				   "					FROM " & _
				   "					   (	SELECT " & _
				   "							   gad.fcba, " & _
				   "							   gad.FDDATE, " & _
				   "							   gad.EMPLOYEECODE, " & _
				   "							   gad.reference reference, " & _
				   "							   NVL (SUM (gad.unit), 0) TotalUnitPerDay " & _
				   "							FROM " & _
				   "							   Employee emp, ATTENDANCE_GAD gad " & _
				   "							WHERE " & _
				   "							   (emp.fccode = gad.EMPLOYEECODE AND " & _
				   "								emp.fcba = gad.fcba) AND " & _
				   "							   gad.attendance IN ('KJ', 'WH', 'WS') AND " & _
				   "							   DATETERMINATE IS NULL AND " & _
				   "							   process_posting_api.GetPostingValue ('PAYROLL', " & _
				   "																	'A32', " & _
				   "																	'JOB', " & _
				   "																	gad.JOBCODE, " & _
				   "																	'" & Session("fcBA") & "') = 'TRUE' AND " & _
				   "							   gad.unit > 0 AND " & _
				   "							   TO_NUMBER (TO_CHAR (gad.FDDATE, 'mm')) = " & Month(Session("v_Employee")) & " AND " & _
				   "							   TO_NUMBER (TO_CHAR (gad.FDDATE, 'yyyy')) = " & Year(Session("v_Employee")) & " AND " & _
				   "							   gad.fcBA = '" & Session("fcBA") & "' " & _
				   "							GROUP BY " & _
				   "							   gad.fcba, " & _
				   "							   gad.FDDATE, " & _
				   "							   gad.EMPLOYEECODE, " & _
				   "							   gad.reference) " & _
				   "					GROUP BY " & _
				   "					   fcba, " & _
				   "					   fddate, " & _
				   "					   employeecode, " & _
				   "					   reference, " & _
				   "					   oppremirate_api.getFirstReference (fcba, fddate, employeecode,'A32'), " & _
				   "									  oppremirate_api.convertQty (fcba, " & _
				   "												   fddate, " & _
				   "												   employeecode, " & _
				   "												   reference, " & _
				   "												   'A32','TRANSPORT'," & _
				   "												   TotalUnitPerDay)) " & _
				   "			GROUP BY " & _
				   "			   fcba, " & _
				   "			   fddate, " & _
				   "			   employeecode, " & _
				   "			   firstreference		 "  
				 

		    
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Premi Transport ..."& "<br>" 
		response.Flush()
		rsEmpAllDed.Open qReceive, connstr
		if not rsEmpAllDed.eof then
		    Do While Not rsEmpAllDed.eof
		   
				Set rsOpPremiRate = Server.CreateObject("ADODB.Recordset")
				qParameter = " select basis, basisrate, extrabasisrate from oppremirate " & _
							  " where " & _
							  " premitypecode ='TRANSPORT' and " & _
							  " fccode = '" & rsEmpAllDed("REFERENCE") & "' and " & _
							  " fcba = '" & Session("fcBA") & "' " 
							
				'response.Write " qParameter = " & qParameter				  
				rsOpPremiRate.CursorType = 3
				rsOpPremiRate.CursorLocation = 3
				rsOpPremiRate.Open qParameter, connstr
				if not rsOpPremiRate.eof then
				    ntempBasis = cdbl(rsOpPremiRate("basis"))
					ntempBasisRate = cdbl(rsOpPremiRate("basisrate"))
					ntempExtraBasisRate = cdbl(rsOpPremiRate("extrabasisrate"))
				else
				    ntempBasis = 0
					ntempBasisRate = 0
					ntempExtraBasisRate = 0
				End if
				rsOpPremiRate.Close
				set rsOpPremiRate = Nothing		   
				
				' Get Total TREE/ ambil jumlah Pokok-nya/Pohon yang di Transport oleh Employee
				tempTotalUnitPerDay = cdbl(rsEmpAllDed ("TotalUnitPerDay"))
				
				if tempTotalUnitPerDay >= ntempBasis then 
				    ' the Employee will have the BASIS.
		   
					maxLineNo = maxLineNo + 1
					nPremiBasis = ntempBasisRate * nRatePremiBasis / 100
					qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
							"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A32" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nPremiBasis & _
							",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
							
					connstr.execute(qAllowance)
					
					
					if tempIncludePayrollProcessingLebihBasis = "TRUE" then 
					    'the Employee will have extra basis 
						tempTotalLebihBasis = tempTotalUnitPerDay - ntempBasis
						nPremiLebihBasis = tempTotalLebihBasis * ntempExtraBasisRate * nRatePremiBasis  / 100
						
						if nPremiLebihBasis <> 0 then 
							
							maxLineNo = maxLineNo + 1
							qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
									"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A64" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nPremiLebihBasis & _
									",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
									
							connstr.execute(qAllowance)
							maxLineNoLebih = maxLineNo 
						end if
					end if							

				end if
				rsEmpAllDed.MoveNext
			loop
		End if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
	end if
	
	' A33
	if tempIncludePayrollProcessingOpr = "TRUE" then
		Response.Write "<br> Start Calculating Payroll Operator.."& "<br>" 
			qOperator = " select                                            																																																														" & _
						" fcba,                                                                                                                                                                         " & _
						" vehicledate,                                                                                                                                                                  " & _
						" employeecode,                                                                                                                                                                 " & _
						" employeename,                                                                                                                                                                 " & _
						" sum(nvl(premi_operator1,0))+sum(nvl(premi_operator2,0)) as total_premi_operator                                                                                                                                   " & _
						" from (                                                                                                                                                                        " & _		
						" SELECT DISTINCT " & _
						"        vhact_fix.fcba," & _
						"        vhact_fix.vehicledate," & _
						"        vhact_fix.employeecode," & _
						"        emp.fcname AS employeename," & _
						"        vhact_fix.vehiclecode," & _
						"        vhact_fix.categoryname," & _
						"        vhact_fix.hasilkerja," & _
						"        vhact_fix.basis," & _
						"        vhact_fix.lebihbasis," & _
						"        vhact_fix.rate," & _
						"        CASE WHEN vhact_fix.vehiclegroupcode <> 'EX' THEN " & _
						"           CASE WHEN (    cat_atas.category_true_teratas = vhact_fix.categorycode" & _
						"                 AND vhact_fix.statusbasis = 'TRUE'" & _
						"                 AND vhact_fix.daystatus = 'Hari Kerja')" & _
						"           THEN" & _
						"              /*lebihbasis * rate */" & _
						"              round((hasilkerja/sumhasilkerja) * ((sumhasilkerja-basis)*rate),0) " & _
						"           ELSE" & _
						"              vhact_fix.hasilkerja * rate" & _
						"           END " & _
						"        ELSE " & _
						"           CASE WHEN vhact_fix.lebihbasis >=  vhact_fix.basis AND vhact_fix.basis <> 0 THEN  " & _
						"                  CASE WHEN vhact_fix.daystatus = 'Hari Kerja' THEN " & _
						"                       vhact_fix.basis * Rate  " & _
						"                       ELSE " & _
						"                       CASE WHEN sumhasilkerja > 8 THEN " & _
						"                          8 * Rate " & _
						"                       ELSE " & _
						"                          sumhasilkerja * Rate " & _
						"                       END " & _
						"                   END " & _
						"           Else " & _
						"                  CASE WHEN vhact_fix.daystatus = 'Hari Kerja' THEN " & _
						"                       vhact_fix.lebihbasis * Rate  " & _
						"                       ELSE " & _
						"                       CASE WHEN sumhasilkerja > 8 THEN " & _
						"                          8 * Rate " & _
						"                       ELSE " & _
						"                          sumhasilkerja * Rate " & _
						"                       END " & _
						"                   END " & _
						"           END " & _
						"        END " & _
						"        AS premi_operator1," & _
						"        CASE WHEN vhact_fix.vehiclegroupcode <> 'EX' THEN " & _
						"           0 " & _
						"        ELSE " & _
						"        CASE WHEN vhact_fix.daystatus = 'Hari Kerja' AND vhact_fix.basis <> 0 THEN " & _
						"           CASE WHEN (vhact_fix.lebihbasis >  vhact_fix.basis) AND ( (vhact_fix.lebihbasis - vhact_fix.basis) > 0 ) THEN  " & _
						"                  (vhact_fix.lebihbasis - vhact_fix.basis) * Rate2 " & _
						"           Else " & _
						"                  0 " & _
						"           END " & _
						"        ELSE " & _
						"            CASE WHEN sumhasilkerja > 8 THEN (sumhasilkerja-8) * Rate2 ELSE 0 END " & _
						"        END END" & _
						"        AS premi_operator2" & _
						"   FROM (SELECT fcba," & _
						"                employeecode," & _
						"                CASE" & _
						"                   WHEN harilibur IS NULL THEN 'Hari Kerja'" & _
						"                   ELSE 'Hari Libur'" & _
						"                END" & _
						"                   AS daystatus," & _
						"                vehiclegroupcode," & _
						"                vehicledate," & _
						"                vehiclecode," & _
						"                categorycode," & _
						"                categoryname," & _
						"                method_by," & _
						"                hasilkerja," & _
						"                basis," & _
						"                CASE" & _
						"                   WHEN hasilkerja - basis < 0 THEN 0" & _
						"                   ELSE hasilkerja - basis" & _
						"                END" & _
						"                   AS lebihbasis," & _
						"                CASE WHEN sUMhasilkerja - basis < 0 THEN 'FALSE' ELSE 'TRUE' END" & _
						"                   AS statusbasis," & _
						"                rate," & _
						"                rate2," & _
						"                sUMhasilkerja" & _
						"           FROM (  SELECT vehicleactivity.fcba," & _
						"                          vehicleactivity.employeecode," & _
						"                          harilibur.harilibur," & _
						"                          vehiclegroup.vehiclegroupcode," & _
						"                          vehicleactivity.vehicledate," & _
						"                          vehicleactivity.vehiclecode," & _
						"                          to_number(vehicleactivitycategory.fccode) AS categorycode," & _
						"                          vehicleactivitycategory.fcname AS categoryname," & _
						"                          NVL (SUM (vehicleactivity.trip), 0) AS trip," & _
						"                          NVL (SUM (vehicleactivity.unit), 0) AS KMHM," & _
						"                          NVL (SUM (vehicleactivity.output), 0) AS output," & _
						"                          operator_category_rate.method_by," & _
						"                          CASE" & _
						"                             WHEN operator_category_rate.method_by = 'OUTPUT'" & _
						"                             THEN" & _
						"                                NVL (SUM (vehicleactivity.output), 0)" & _
						"                             WHEN operator_category_rate.method_by = 'HM'" & _
						"                             THEN" & _
						"                                NVL (SUM (vehicleactivity.unit), 0)" & _
						"                             ELSE" & _
						"                                NVL (SUM (vehicleactivity.trip), 0)" & _
						"                          END" & _
						"                             AS hasilkerja," & _
						"                          CASE " & _
						"                             WHEN operator_category_rate.method_by = 'OUTPUT' " & _
						"                             THEN " & _
						"                                NVL (smax.sumoutput, 0) " & _
						"                             WHEN operator_category_rate.method_by = 'HM' " & _
						"                             THEN " & _
						"                                NVL (SUM (vehicleactivity.unit), 0) " & _
						"                             ELSE " & _
						"                                NVL (smax.sumtrip, 0) " & _
						"                          END " & _
						"                             AS sUMhasilkerja, " & _
						"                          operator_category_rate.basis," & _
						"                          operator_category_rate.rate," & _
						"                          operator_category_rate.rate2" & _
						"                     FROM vehicleactivity" & _
						"                          INNER JOIN vehicleactivitycategory" & _
						"                             ON (vehicleactivity.categorytransport =" & _
						"                                    vehicleactivitycategory.fccode)" & _
						"                          INNER JOIN" & _
						"                          (SELECT fcba," & _
						"                                  fccode," & _
						"                                  fcname," & _
						"                                  vehiclegroupcode" & _
						"                             FROM vehicle WHERE fcba='"& Session("fcBA") & "') vehicle" & _
						"                             ON (    vehicleactivity.fcba = vehicle.fcba" & _
						"                                 AND vehicleactivity.vehiclecode =" & _
						"                                        vehicle.fccode)" & _
						"                          INNER JOIN" & _
						"                          (SELECT fcba, vehiclegroupcode, description" & _
						"                             FROM vehiclegroup WHERE fcba='"& Session("fcBA") & "') vehiclegroup" & _
						"                             ON (    vehicleactivity.fcba = vehiclegroup.fcba" & _
						"                                 AND vehicle.vehiclegroupcode =" & _
						"                                        vehiclegroup.vehiclegroupcode)" & _
						"                          INNER JOIN" & _
						"                          (SELECT DISTINCT fcba," & _
						"                                           fddate," & _
						"                                           employeecode," & _
						"                                           locationcode," & _
						"                                           case WHEN attendance IN ('KJ','WS','WH','KK','DL') AND  To_Char(fddate,'D')='1' then 'WS' " & _
						"                                           ELSE " & _
						"												case when attendance IN ('KJ','WS','WH','KK','DL' ) AND TO_CHAR(fddate,'yyyymmdd')  IN (SELECT DISTINCT TO_CHAR(holidaydate,'yyyymmdd') FROM budget_transaction_holiday where To_Char(holidaydate,'yyyy')='" & Year(Session("v_Employee")) & "') THEN" & _
						"                                           	'WH' " & _
						" 												ELSE " & _
						" 												attendance " & _
						" 												END " & _
						" 											END attendance " & _						
						"                             FROM attendance_gad" & _
						"                            WHERE     fcba = '" & Session("fcBA") & "'" & _
						"                                  AND locationtype = 'VH'" & _
						"                     			   AND TO_NUMBER (TO_CHAR (fddate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
						"                     			   AND TO_NUMBER (TO_CHAR (fddate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
						"                                  AND attendance IN ('KJ'," & _
						"                                                     'WH'," & _
						"                                                     'WS'," & _
						"                                                     'KK'," & _
						"                                                     'DL')) attendance_gad" & _
						"                             ON (    attendance_gad.fcba = vehicleactivity.fcba" & _
						"                                 AND attendance_gad.fddate =" & _
						"                                        vehicleactivity.vehicledate" & _
						"                                 AND attendance_gad.employeecode =" & _
						"                                        vehicleactivity.employeecode" & _
						"                                 AND attendance_gad.locationcode =" & _
						"                                        vehicleactivity.vehiclecode)" & _
						"                          INNER JOIN operator_category_rate" & _
						"                             ON (    vehicleactivity.fcba =" & _
						"                                        operator_category_rate.fcba" & _
						"                                 AND vehicleactivity.categorytransport =" & _
						"                                        operator_category_rate.categorytransport" & _
						"                                 AND vehicle.vehiclegroupcode =" & _
						"                                        operator_category_rate.vehicle_group" & _
						"                                 AND attendance_gad.attendance =" & _
						"                                        operator_category_rate.attendance)" & _
						"                          LEFT JOIN" & _
						"                          (SELECT DISTINCT HariLibur" & _
						"                             FROM (SELECT dt AS HariLibur" & _
						"                                     FROM (WITH dts" & _
						"                                                AS (    SELECT   DATE '"& Year(Session("v_Employee")) &"-01-01'" & _
						"                                                               + ROWNUM -1" & _
						"                                                                  dt" & _
						"                                                          FROM DUAL" & _
						"                                                    CONNECT BY LEVEL <= 366)" & _
						"                                           SELECT *" & _
						"                                             FROM dts" & _
						"                                            WHERE TO_CHAR (dt, 'fmday') =" & _
						"                                                     'sunday')" & _
						"                                   UNION ALL" & _
						"                                   SELECT DISTINCT holidaydate" & _
						"                                     FROM budget_transaction_holiday" & _
						"                                    WHERE fcba IN (SELECT fccode FROM businessunit where fccompanycode='" & Session("COMPANYCODE") & "')" & _
						"										)) harilibur" & _
						"                             ON (vehicleactivity.vehicledate =" & _
						"                                    harilibur.HariLibur)" & _
						"          INNER JOIN  " & _
                        "        ( Select EmployeeCode,fcba,VehicleDate,categoryTransport,sum(Trip) sumTrip,sum(output) sumoutput " & _
                        "          from vehicleactivity " & _
                        "                WHERE     vehicleactivity.fcba = '" & Session("fcBA") & "' " & _
						"                     	   AND TO_NUMBER (TO_CHAR (vehicleactivity.vehicledate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
						"                          AND TO_NUMBER (TO_CHAR (vehicleactivity.vehicledate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
						"                          group by EmployeeCode,fcba,VehicleDate,categoryTransport " & _
                        "        ) smax on (vehicleactivity.EmployeeCode = smax.EmployeeCode and vehicleactivity.fcba = smax.fcba  " & _
                        "                   and vehicleactivity.VehicleDate = smax.VehicleDate " & _
                        "                   and vehicleactivity.categoryTransport = smax.categoryTransport " & _
                        "                   ) " & _
						"                    WHERE     vehicleactivity.fcba = '" & Session("fcBA") & "'" & _
						"                     	   AND TO_NUMBER (TO_CHAR (vehicleactivity.vehicledate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
						"                          AND TO_NUMBER (TO_CHAR (vehicleactivity.vehicledate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
						" 	                GROUP BY vehicleactivity.fcba," & _
						"                          vehicleactivity.employeecode," & _
						"                          attendance_gad.attendance," & _
						"                          harilibur.harilibur," & _
						"                          vehiclegroup.vehiclegroupcode," & _
						"                          vehicleactivity.vehicledate," & _
						"                          vehicleactivity.vehiclecode," & _
						"                          vehicleactivitycategory.fccode," & _
						"                          vehicleactivitycategory.fcname," & _
						"                          operator_category_rate.method_by," & _
						"                          operator_category_rate.basis," & _
						"                          operator_category_rate.rate," & _
						"                          operator_category_rate.rate2," & _
						"                          smax.sumtrip," & _
                        "                          smax.sumoutput" & _
						"                          )" & _
						"		  ) vhact_fix" & _
						"        INNER JOIN" & _
						"        (SELECT fcba," & _
						"                employeecode," & _
						"                vehiclegroupcode," & _
						"                vehicledate," & _
						"                vehiclecode," & _
						"                categorycode," & _
						"                statusbasis," & _
						"                category_true_teratas" & _
						"           FROM (SELECT fcba," & _
						"                        employeecode," & _
						"                        vehiclegroupcode," & _
						"                        vehicledate," & _
						"                        vehiclecode," & _
						"                        categorycode," & _
						"                        statusbasis," & _
						"                        MIN (" & _
						"                           categorycode)" & _
						"                        OVER (" & _
						"                           PARTITION BY fcba || employeecode || vehicledate)" & _
						"                           AS category_true_teratas" & _
						"                   FROM (SELECT fcba," & _
						"                                employeecode," & _
						"                                vehiclegroupcode," & _
						"                                vehicledate," & _
						"                                vehiclecode," & _
						"                                categorycode," & _
						"                                CASE" & _
						"                                 WHEN (    daystatus = 'Hari Kerja'" & _
						"										AND hasilkerja - basis < 0)" & _
						"								  THEN" & _
						"									 'FALSE'" & _
						"								  WHEN (    daystatus = 'Hari Libur'" & _
						"										AND hasilkerja - basis < 0)" & _
						"								  THEN" & _
						"									 'TRUE'" & _
						"								  ELSE" & _
						"									 'TRUE'" & _
						"                                END" & _
						"                                   AS statusbasis" & _
						"                           FROM (  SELECT vehicleactivity.fcba," & _
						"                                          vehicleactivity.employeecode," & _
						"                                          vehiclegroup.vehiclegroupcode," & _
						"                                          vehicleactivity.vehicledate," & _
						"										   CASE " & _
						"											WHEN harilibur.harilibur IS NULL" & _
						"											THEN" & _
						"											   'Hari Kerja'" & _
						"											ELSE" & _
						"											   'Hari Libur'" & _
						"										 END" & _
						"											AS daystatus," & _
						"                                          vehicleactivity.vehiclecode," & _
						"                                          to_number(vehicleactivitycategory.fccode) AS categorycode," & _
						"                                          NVL (SUM (vehicleactivity.trip), 0)" & _
						"                                             AS trip," & _
						"                                          NVL (SUM (vehicleactivity.unit), 0)" & _
						"                                             AS KMHM," & _
						"                                          NVL (SUM (vehicleactivity.output), 0)" & _
						"                                             AS output," & _
						"                                          operator_category_rate.method_by," & _
						"                                          CASE" & _
						"                                             WHEN operator_category_rate.method_by =" & _
						"                                                     'OUTPUT'" & _
						"                                             THEN" & _
						"                                                NVL (" & _
						"                                                   SUM (vehicleactivity.output)," & _
						"                                                   0)" & _
						"                                             WHEN operator_category_rate.method_by =" & _
						"                                                     'HM'" & _
						"                                             THEN" & _
						"                                                NVL (SUM (vehicleactivity.unit)," & _
						"                                                     0)" & _
						"                                             ELSE" & _
						"                                                NVL (SUM (vehicleactivity.trip)," & _
						"                                                     0)" & _
						"                                          END" & _
						"                                             AS hasilkerja," & _
						"                                          operator_category_rate.basis" & _
						"                                     FROM vehicleactivity" & _
						"                                          INNER JOIN vehicleactivitycategory" & _
						"                                             ON (vehicleactivity.categorytransport =" & _
						"                                                    vehicleactivitycategory.fccode)" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT fcba," & _
						"                                                  fccode," & _
						"                                                  fcname," & _
						"                                                  vehiclegroupcode" & _
						"                                             FROM vehicle WHERE fcba='"& Session("fcBA") & "') vehicle" & _
						"                                             ON (    vehicleactivity.fcba =" & _
						"                                                        vehicle.fcba" & _
						"                                                 AND vehicleactivity.vehiclecode =" & _
						"                                                        vehicle.fccode)" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT fcba," & _
						"                                                  vehiclegroupcode," & _
						"                                                  description" & _
						"                                             FROM vehiclegroup WHERE fcba='"& Session("fcBA") & "') vehiclegroup" & _
						"                                             ON (    vehicleactivity.fcba =" & _
						"                                                        vehiclegroup.fcba" & _
						"                                                 AND vehicle.vehiclegroupcode =" & _
						"                                                        vehiclegroup.vehiclegroupcode)" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT DISTINCT fcba," & _
						"                                                           fddate," & _
						"                                                           employeecode," & _
						"                                                           locationcode," & _
						"                                                           attendance" & _
						"                                             FROM attendance_gad" & _
						"                                            WHERE     fcba = '" & Session("fcBA") & "'" & _
						"                                                  AND locationtype = 'VH'" & _
						"                     							   AND TO_NUMBER (TO_CHAR (fddate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
						"                     							   AND TO_NUMBER (TO_CHAR (fddate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
						"                                                  AND attendance IN ('KJ'," & _
						"                                                                     'WH'," & _
						"                                                                     'WS'," & _
						"                                                                     'KK'," & _
						"                                                                     'DL'))" & _
						"                                          attendance_gad" & _
						"                                             ON (    attendance_gad.fcba =" & _
						"                                                        vehicleactivity.fcba" & _
						"                                                 AND attendance_gad.fddate =" & _
						"                                                        vehicleactivity.vehicledate" & _
						"                                                 AND attendance_gad.employeecode =" & _
						"                                                        vehicleactivity.employeecode" & _
						"                                                 AND attendance_gad.locationcode =" & _
						"                                                        vehicleactivity.vehiclecode)" & _
						"                                          INNER JOIN operator_category_rate" & _
						"                                             ON (    vehicleactivity.fcba =" & _
						"                                                        operator_category_rate.fcba" & _
						"                                                 AND vehicleactivity.categorytransport =" & _
						"                                                        operator_category_rate.categorytransport" & _
						"                                                 AND vehicle.vehiclegroupcode =" & _
						"                                                        operator_category_rate.vehicle_group" & _
						"                                                 AND attendance_gad.attendance =" & _
						"                                                        operator_category_rate.attendance)" & _
						"                                         LEFT JOIN" & _
                        "                 (SELECT DISTINCT HariLibur" & _
                        "                    FROM (SELECT dt AS HariLibur" & _
                        "                            FROM (WITH dts" & _
                        "                                       AS (    SELECT   DATE '"& Year(Session("v_Employee")) &"-01-01'" & _
                        "                                                      + ROWNUM" & _
                        "                                                      - 1" & _
                        "                                                         dt" & _
                        "                                                 FROM DUAL" & _
                        "                                           CONNECT BY LEVEL <=" & _
                        "                                                         366)" & _
                        "                                  SELECT *" & _
                        "                                    FROM dts" & _
                        "                                   WHERE TO_CHAR (" & _
                        "                                            dt," & _
                        "                                            'fmday') =" & _
                        "                                            'sunday')" & _
                        "                          UNION ALL" & _
                        "                          SELECT DISTINCT holidaydate" & _
                        "                            FROM budget_transaction_holiday" & _
                        "                           WHERE fcba IN (SELECT fccode FROM businessunit where fccompanycode='" & Session("COMPANYCODE") & "')" & _
                        "                            ))" & _
                        "                 harilibur" & _
                        "                    ON (vehicleactivity.vehicledate = harilibur.HariLibur)" & _
						"                     WHERE vehicleactivity.fcba = '" & Session("fcBA") & "' AND                                                                                                      " & _
						"                     TO_NUMBER (TO_CHAR (vehicleactivity.vehicledate, 'mm')) = " & Month(Session("v_Employee")) & " AND                                                        " & _
						"                     TO_NUMBER (TO_CHAR (vehicleactivity.vehicledate, 'yyyy')) = " & Year(Session("v_Employee")) & "                                                           " & _
						"                                 GROUP BY vehicleactivity.fcba," & _
						"                                          vehicleactivity.employeecode," & _
						"                                          vehiclegroup.vehiclegroupcode," & _
						"                                          vehicleactivity.vehicledate," & _
						"                                          vehicleactivity.vehiclecode," & _
						"                                          harilibur.harilibur," & _
						"                                          vehicleactivitycategory.fccode," & _
						"                                          operator_category_rate.method_by," & _
						"                                          operator_category_rate.basis," & _
						"                                          operator_category_rate.rate))" & _
						"                  WHERE statusbasis = 'TRUE')" & _
						"          WHERE category_true_teratas = categorycode) cat_atas" & _
						"           ON (    vhact_fix.fcba = cat_atas.fcba" & _
						"               AND vhact_fix.employeecode = cat_atas.employeecode" & _
						"               AND vhact_fix.vehicledate = cat_atas.vehicledate" & _
						"              /* AND vhact_fix.vehiclecode = cat_atas.vehiclecode*/" & _
						"              )" & _
						"        INNER JOIN (SELECT fcba, fccode, fcname" & _
						"                      FROM employee" & _
						"                     WHERE fcba = '" & Session("fcBA") & "') emp" & _
						"           ON (    vhact_fix.fcba = emp.fcba" & _
						"               AND vhact_fix.employeecode = emp.fccode)"  & _
						" )" & _
						" group by fcba, vehicledate, employeecode, employeename"
		
		'response.write qOperator &"</BR>"
		Set rsOperator = Server.CreateObject("ADODB.Recordset")
		rsOperator.CursorType = 3
		rsOperator.CursorLocation = 3
		rsOperator.Open qOperator, connstr		
		
		If Not rsOperator.EOF then
			Do While Not rsOperator.EOF
				If cdbl(rsOperator("total_premi_operator")) <> 0 Then				
					maxLineNo	=	maxLineNo + 1
					qAllowance	=	" Insert Into EmpAllded (" & _
									"     EMPCODE" & _
									"     , ALLOWDEDCODE" & _
									"     , ADHOCDATE" & _
									"     , AMOUNT" & _
									"     , FIXEDALLDED" & _
									"     , FCENTRY" & _
									"     , FCEDIT" & _
									"     , FCIP" & _
									"     , FCBA" & _
									"     , REMARKS" & _
									"     , LASTUPDATE" & _
									"     , LASTTIME" & _
									"     , LINENO" & _
									" ) Values (" & _
									"     '" & rsOperator("EMPLOYEECODE") & "'" & _
									"     , 'A33'" & _
									"     , to_date('" & rsOperator("VEHICLEDATE") &"','mm/dd/yyyy')" & _
									"     , " & cdbl(rsOperator("total_premi_operator")) & "" & _
									"     , ''" & _
									"     , '" & Session("CodeName") & "'" & _
									"     , '" & Session("CodeName") & "'" & _
									"     , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"     , '" & Session("fcBA") & "'" & _
									"     , ''" & _
									"     , sysdate" & _
									"     , to_char(sysdate,'HH24:MI')" & _
									"     , " & maxLineNo & "" & _
									" )"
					connstr.execute(qAllowance)
					
					'Response.Write "."
					'End Remarks By Iwa : 08/01/2019
				End If
				
				rsOperator.MoveNext
			Loop
		End If
		Response.Write "<br> Calculating Payroll Operator Success.."& "<br>" 
		rsOperator.close
		Set rsOperator = Nothing
	end if
	' End A33 
	
	'A62
	if tempIncludePayrollProcessingByDiv = "TRUE" then 
	'the Employee will have extra basis 
		'response.write nBasisPremiType
		if nBasisPremiType = "MIN" then
			'normalisasi tabel SUM_PREMI_OPR_DT_MINBASIS
			
			qNormalisasiTempTable = " delete SUM_PREMI_OPR_DT_MINBASIS" & _
									" where fcba = '" & Session("fcBA") & "'" 
			connstr.execute(qNormalisasiTempTable)
			
			
			qUpdateTempTable = "insert into SUM_PREMI_OPR_DT_MINBASIS select * from SUMMARY_PREMI_OPR_DT_MINBASIS " & _
							   " where fcba='" & Session("fcBA") & "' and period_year = " & Year(Session("v_Employee")) & " and " & _
							   " period_month = " & Month(Session("v_Employee")) & ""
			connstr.execute(qUpdateTempTable)
				
			qA62 =	" select " & _
					" FCBA," & _
					" PERIOD_YEAR," & _
					" PERIOD_MONTH," & _
					" LAST_DAY(TO_DATE(PERIOD_YEAR||PERIOD_MONTH,'yyyymm')) FDDATE," & _
					" OPERATORCODE EMPLOYEECODE," & _
					" TOTAL_PREMI_OPERATOR_MIN_BASIS TOTAL_AMOUNT" & _
					" FROM " & _
					" SUM_PREMI_OPR_DT_MINBASIS" & _
					" WHERE" & _
					" fcba = '" & Session("fcBA") & "'" & _
					" and PERIOD_YEAR = " & Year(Session("v_Employee")) & "" & _
					" and PERIOD_MONTH = " & Month(Session("v_Employee")) & ""

		else
			'normalisasi tabel SUM_PREMI_OPR_DT_MINBASIS
			
			qNormalisasiTempTable = " delete SUM_PREMI_OPR_DT_AVGBASIS" & _
									" where fcba = '" & Session("fcBA") & "'"  
			connstr.execute(qNormalisasiTempTable)
			
			qUpdateTempTable = "insert into SUM_PREMI_OPR_DT_AVGBASIS select * from SUMMARY_PREMI_OPR_DT_AVGBASIS " & _
							   " where fcba='" & Session("fcBA") & "' and period_year = " & Year(Session("v_Employee")) & " and " & _
							   " period_month = " & Month(Session("v_Employee")) & ""
			connstr.execute(qUpdateTempTable)
			
			qA62 =	" select " & _
					" FCBA," & _
					" PERIOD_YEAR," & _
					" PERIOD_MONTH," & _
					" LAST_DAY(TO_DATE(PERIOD_YEAR||PERIOD_MONTH,'yyyymm')) FDDATE," & _
					" OPERATORCODE EMPLOYEECODE," & _
					" TOTAL_PREMI_OPERATOR_MIN_BASIS TOTAL_AMOUNT" & _
					" FROM " & _
					" SUM_PREMI_OPR_DT_AVGBASIS" & _
					" WHERE" & _
					" fcba = '" & Session("fcBA") & "'" & _
					" and PERIOD_YEAR = " & Year(Session("v_Employee")) & "" & _
					" and PERIOD_MONTH = " & Month(Session("v_Employee")) & ""
			
		end if
	
		Set rsA62 = Server.CreateObject("ADODB.Recordset")
					rsA62.CursorType = 3
					rsA62.CursorLocation = 3
					response.Write "Calculate Premi Transport Operator By Division..."& "<br>" 
					response.Flush()
					rsA62.Open qA62, connstr
					
					'response.write qA62
					if not rsA62.eof then
						Do While Not rsA62.eof						
							
							if Cdbl(rsA62("TOTAL_AMOUNT")) <> 0 then 								
								maxLineNo = maxLineNo + 1
								qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
											"(" & maxLineNo & ",'" & rsA62("EMPLOYEECODE") & "','" & "A62" & "'," &  " to_date('" & rsA62("FDDATE") & "','mm/dd/yyyy') " & "," & Cdbl(rsA62("TOTAL_AMOUNT")) & _
											",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
											Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
								
								'response.write qAllowance&"</BR>"

								connstr.execute(qAllowance)
								maxLineNoLebih = maxLineNo 
								
							end if
							rsA62.MoveNext
						loop
					end if
					rsA62.Close
					set rsA62 = Nothing
	end if
	
	
	' get allowded A63
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select includedpayrollprocessing " & _
				" From ALLOWDED where FCCODE='A63'  " & _
				" and man_auto = 'A' " & _
				" and fcBA='" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		tempIncludePayrollProcessingByTrip = rsParameter("includedpayrollprocessing")
	else
		tempIncludePayrollProcessingByTrip = "FALSE"
	End if
	rsParameter.Close
	set rsParameter = Nothing
	'end get allowded A63
	
	'A63
	if tempIncludePayrollProcessingByTrip = "TRUE" then 
			
			qA63 =	" select" & _
					" fcba, year, month, " & _
					" LAST_DAY(TO_DATE(year||month,'yyyymm')) FDDATE," & _
					" EMPLOYEECODE,premi TOTAL_AMOUNT" & _
					" from (" & _
					"     select" & _
					"     fcba,year,month,employeecode,sum(premi) as premi" & _
					"     from premi_operator_dt_bytrip" & _
					" 	  where " & _
					"     fcba = '" & Session("fcBA") & "'" & _
					"     and year = " & Year(Session("v_Employee")) & "" & _
					"     and month = " & Month(Session("v_Employee")) & "" & _
					"     and premi > 0" & _
					"     group by fcba,year,month,employeecode" & _
					" )"
		
			'the Employee will have extra basis 
			Set rsA63 = Server.CreateObject("ADODB.Recordset")
						rsA63.CursorType = 3
						rsA63.CursorLocation = 3
						response.Write "Calculate Premi Transport Operator By Division..."& "<br>" 
						response.Flush()
						rsA63.Open qA63, connstr
												
						if not rsA63.eof then
							Do While Not rsA63.eof						
								
								if Cdbl(rsA63("TOTAL_AMOUNT")) <> 0 then 								
									maxLineNo = maxLineNo + 1
									qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"(" & maxLineNo & ",'" & rsA63("EMPLOYEECODE") & "','" & "A63" & "'," &  " to_date('" & rsA63("FDDATE") & "','mm/dd/yyyy') " & "," & Cdbl(rsA63("TOTAL_AMOUNT")) & _
												",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
							
									connstr.execute(qAllowance)
									maxLineNoLebih = maxLineNo 
									
								end if
								rsA63.MoveNext
							loop
						end if
						rsA63.Close
			set rsA63 = Nothing
			
	end if
%>