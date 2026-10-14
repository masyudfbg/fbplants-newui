<%
		' Premi Insentif

		allowance_kerja_panen = "A28"
		
		q = " SELECT parameterdetail_api.listfccodeforwherestmt ('" & Session("fcBA") & "', 'HARDCODE', 'JOBCODE_PANEN') wherestmt FROM DUAL"
		' response.write q
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.Open q, connstr
		IF NOT rsEmpAllDed.EOF THEN
			wherestmt = rsEmpAllDed("wherestmt")
		END IF
		rsEmpAllDed.Close : set rsEmpAllDed = Nothing
		
		private function TotalRecordPerDay(arr_, fddate_, empcode_)
			result_ = 0
			arr1 = arr_
			FOR xyz = LBound(arr1, 2) TO UBound(arr1, 2)
				if cstr(""&arr1(2,xyz)) = cstr(""&empcode_) and cdate(arr1(1,xyz)) = cdate(fddate_) then
					result_ = result_ +1
				end if
			NEXT
			TotalRecordPerDay = result_
		end function
		
		private function TotalBunchPerDay(arr_, fddate_, empcode_)
			result_ = 0
			arr1 = arr_
			FOR xyz = LBound(arr1, 2) TO UBound(arr1, 2)
				if cstr(""&arr1(2,xyz)) = cstr(""&empcode_) and cdate(arr1(1,xyz)) = cdate(fddate_) then
					result_ = result_ +cdbl(to_nol(arr1(6,xyz)))
				end if
			NEXT
			TotalBunchPerDay = result_
		end function
		
		private function RoundOracle(number_)
			result_ = 0
			q = " SELECT ROUND('"&number_&"', 0) amount FROM DUAL"
			' response.write q
			Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
			rsEmpAllDed.Open q, connstr
			IF NOT rsEmpAllDed.EOF THEN
				result_ = cdbl(to_nol(rsEmpAllDed("amount")))
			END IF
			rsEmpAllDed.Close : set rsEmpAllDed = Nothing
			
			RoundOracle = result_
		end function
		
		arr_data_kp_eof = true
		
		'strMain = " Select "& _
		'		  " FCBA, "& _
		'		  "	FDDATE, "& _
		'		  "	EMPLOYEECODE, "& _
		'		  "	PRESENTRATE, "& _
		'		  "	TOTALAMOUNT, "& _
		'		  "	A01, "& _
		'		  "	A24 "& _
		'		  "	from "& _
		'		  "	harvesting_premi "& _
		'		  " Where FCBA = '" & Session("fcBA") & "'"& _
		'		  " AND TO_CHAR(TO_DATE(FDDATE,'dd/mm/yyyy'),'fmMM') = " & Month(Session("v_Employee")) & _
		'		 " AND TO_CHAR(TO_DATE(FDDATE,'dd/mm/yyyy'),'yyyy') = " & Year(Session("v_Employee"))
		
		strMain = " select"& _
				  " fcba,"& _
				  " fddate,"& _
				  " employeecode,"& _
				  " case when amount >= 0 then amount else 0 end amount"& _
				  " from ("& _
				  " select"& _
				  " hdb.fcba,"& _
				  " hdb.fddate,"& _
				  " hdb.employeecode,"& _
				  " pm.param_1,"& _
				  " hdb.extraamount_bydaily,"& _
				  " hdb.deductionamount_bydaily,"& _
				  " hdb.extraamount_bymonthly,"& _
				  " hdb.deductionamount_bymonthly,"& _
				  " case "& _
				  " when pm.param_1 = 'D' then (hdb.extraamount_bydaily+hdb.deductionamount_bydaily)"& _
 				  " else (hdb.extraamount_bymonthly+hdb.deductionamount_bymonthly)"& _
				  " end amount"& _
				  " from iplasprod.harvesting_daily_bjr hdb"& _
				  " inner join iplasprod.parameter pm on (hdb.fcba = pm.fcba)"& _
				  " where hdb.fcba ='" & Session("fcBA") & "'"& _
				  " and TO_CHAR(TO_DATE(FDDATE,'dd/mm/yyyy'),'fmMM') = " & Month(Session("v_Employee")) & _
				  " and TO_CHAR(TO_DATE(FDDATE,'dd/mm/yyyy'),'yyyy') = " & Year(Session("v_Employee")) & _
				  " and pm.parcode = 'HARDCODE'"& _
				  " and pm.fccode = 'JOBCODE_PANEN'"& _
				  " )"

		'response.Write " sql " & strMain

		response.Write " Calculate Insentif Panen ..." & "<br>" 
		response.Flush()
		
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.Open strMain, connstr
		IF NOT rsEmpAllDed.EOF THEN
			arr_data_kp = rsEmpAllDed.getrows()
			arr_data_kp_eof = false
		END IF
		rsEmpAllDed.Close : set rsEmpAllDed = Nothing
		
		IF NOT arr_data_kp_eof THEN
			FOR xxxx=0 TO UBound(arr_data_kp, 2)
														
				PostAmountInsentifPanen = RoundOracle(arr_data_kp(3,xxxx))
				if PostAmountInsentifPanen > 0 then
							maxLineNo = maxLineNo + 1
							qAllowance = " INSERT INTO empallded ( lineno" & _
										  "                               , empcode" & _
										  "                               , allowdedcode" & _
										  "                               , adhocdate" & _
										  "                               , amount" & _
										  "                               , fixedallded" & _
										  "                               , remarks" & _
										  "                               , fcentry" & _
										  "                               , fcedit" & _
										  "                               , fcip" & _
										  "                               , fcba" & _
										  "                               , lastupdate" & _
										  "                               , lasttime)" & _
										  " VALUES ( " & maxLineNo & "" & _
										  "          , '" & arr_data_kp(2, xxxx) & "'" & _
										  "          , '" & allowance_kerja_panen & "'" & _
										  "          , to_date('" & arr_data_kp(1, xxxx) & "','dd/mm/yyyy')" & _
										  "          , " & PostAmountInsentifPanen & "" & _
										  "          , 0" & _
										  "          , ''" & _
										  "          , '" & Session("CodeName") & "'" & _
										  "          , '" & Session("CodeName") & "'" & _
										  "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										  "          , '" & Session("fcBA") & "'" & _
										  "          , SYSDATE" & _
										  "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
										  
							'response.Write "<br/>qAllowance : " & qAllowance
							connstr.execute(qAllowance)
				end if				
			NEXT
		END IF
%>