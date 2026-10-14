<%

		' **** Konsep Panen Borongan **** '
		' Dikarenakan konsep alokasi employee iPlaS memakai upah rata-rata, sehingga proses perhitungan panen borongan di iPlaS tetap memakai absensi KJ
		' maka akan terbentuk gaji 1 hk, tetapi gaji 1 hk tersebut langsung di potong memakai potongan tidak dapat basis (D19) sehingga gaji menjadi 0
		' Dan upah borongan panen dalam 1 hari memakai allowded Premi Lebih Borong (A24) dengan konsep perhitungan JJG x BJR x Rp/Ton
		' Kondisi memotong upah, yaitu: sama dengan A01 yang didapatkan karyawan


		allowance_kerja_panen = "A24"
		deduction_kerja_panen = "D19"
		
		allowance_gaji = "A01"
		
		arr_data_kp_eof = true
		
		qReceive = " SELECT fcba" & _
		           "        , fddate" & _
		           "        , employeecode" & _
		           "        , jobcode" & _
		           "        , locationtype" & _
		           "        , locationcode" & _
		           "        , NVL(unit, 0) unit" & _
		           "        , NVL(field_api.getvalid_harvestingbased_abw(locationcode, fcba, fddate), 0) harvestingbased_abw" & _
		           "        , NVL(harvesting_rate_api.get_rate ( field_api.getvalid_croptype(locationcode, fcba, fddate)," & _
		           "                                         field_api.getvalid_afdeling(locationcode, fcba, fddate)," & _
		           "                                         to_char(field_api.getvalid_plantingdate(locationcode, fcba, fddate), 'yyyy')," & _
		           "                                         field_api.getvalid_terraintype(locationcode, fcba, fddate)," & _
		           "                                         fcba), 0) rate" & _
				   "        , presentgrade" & _
				   "        , presentrate" & _
				   "        , PAYROLL_PANEN_API.get_actual_panen_hk(fcba, employeecode, fddate) panen" & _
		           " FROM (" & _
		           "     SELECT ag.fcba" & _
		           "            , ag.fddate" & _
		           "            , ag.employeecode" & _
		           "            , em.presentgrade" & _
		           "            , em.presentrate" & _
		           "            , ag.jobcode" & _
		           "            , ag.locationtype" & _
		           "            , ag.locationcode" & _
		           "            , nvl(sum(ag.unit), 0) unit" & _
		           "     FROM attendance_gad ag" & _
		           "     INNER JOIN employee em ON (ag.employeecode = em.fccode AND ag.fcba = em.fcba)" & _
		           "     WHERE em.dateterminate IS NULL" & _
		           "          AND ag.attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
		           "          AND ag.jobcode IN ('1323001', '6502101')" & _
		           "          AND ag.locationtype = 'FF'" & _
		           "          AND ag.mandays > 0" & _
		           "          AND ag.unit > 0" & _
		           "          AND to_number(to_char(ag.fddate,'mm')) = " & Month(Session("v_Employee")) & "" & _
		           "          AND to_number(to_char(ag.fddate,'yyyy')) = " & Year(Session("v_Employee")) & "" & _
		           "          AND ag.fcba = '" & Session("fcBA") & "'" & _
		           "     GROUP BY ag.fcba" & _
		           "              , ag.fddate" & _
		           "              , ag.employeecode" & _
				   "              , em.presentgrade" & _
		           "              , em.presentrate" & _
		           "              , ag.Jobcode" & _
		           "              , ag.locationtype" & _
		           "              , ag.locationcode" & _
		           " )"
				
		qReceive = " SELECT * FROM ( " & qReceive & " ) WHERE panen <> 0"
				
		'response.Write " sql " & qReceive
		response.Write "Calculate Kerja Panen ..." & "<br>" 
		response.Flush()
		
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.Open qReceive, connstr
		IF NOT rsEmpAllDed.EOF THEN
			arr_data_kp = rsEmpAllDed.getrows()
			arr_data_kp_eof = false
		END IF
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
		
		IF NOT arr_data_kp_eof THEN
			FOR xxxx = LBound(arr_data_kp, 2) TO UBound(arr_data_kp, 2)
				amount_kp = 0
				amount_kp = Round( ( Cdbl(arr_data_kp(6, xxxx)) * Cdbl(arr_data_kp(7, xxxx)) )* Cdbl(arr_data_kp(8, xxxx)), 0)
				
				IF amount_kp > 0 THEN
					maxLineNo = maxLineNo + 1
					qAllowance = " INSERT INTO empallded ( lineno" & _
								  "                                , empcode" & _
								  "                                , allowdedcode" & _
								  "                                , adhocdate" & _
								  "                                , amount" & _
								  "                                , fixedallded" & _
								  "                                , fcentry" & _
								  "                                , fcedit" & _
								  "                                , fcip" & _
								  "                                , fcba" & _
								  "                                , lastupdate" & _
								  "                                , lasttime)" & _
								  " values ( " & maxLineNo & "" & _
								  "          , '" & arr_data_kp(2, xxxx) & "'" & _
								  "          , '" & allowance_kerja_panen & "'" & _
								  "          , to_date('" & arr_data_kp(1, xxxx) & "','mm/dd/yyyy')" & _
								  "          , " & amount_kp & "" & _
								  "          , 0" & _
								  "          , '" & Session("CodeName") & "'" & _
								  "          , '" & Session("CodeName") & "'" & _
								  "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								  "          , '" & Session("fcBA") & "'" & _
								  "          , sysdate" & _
								  "          , to_char(SYSDATE, 'HH24:MI'))"
								  
					'response.Write "<br/>qAllowance : " & qAllowance
					connstr.execute(qAllowance)
						

				END IF
					
			NEXT
		END IF
		
		arr_data_kp_eof = true		
		
		qReceive = " SELECT KP.FCBA" & _
                   "        , KP.ADHOCDATE" & _
                   "        , KP.EMPCODE" & _
                   "        , KP.AMOUNT AMOUNT_KP" & _
                   "        , GJ.AMOUNT AMOUNT_GJ" & _
                   "        , PAYROLL_PANEN_API.GET_TOTAL_BASIS_HK(KP.FCBA, KP.EMPCODE, TO_CHAR(KP.ADHOCDATE, 'YYYYMM')) BASIS" & _
                   "        , PAYROLL_PANEN_API.GET_TOTAL_PANEN_HK(KP.FCBA, KP.EMPCODE, TO_CHAR(KP.ADHOCDATE, 'YYYYMM')) PANEN" & _
                   " FROM ( SELECT FCBA" & _
                   "               , TRUNC(ADHOCDATE, 'MM') ADHOCDATE" & _
                   "               , EMPCODE" & _
                   "               , NVL(SUM(AMOUNT), 0) AMOUNT" & _
                   "        FROM EMPALLDED" & _
                   "        WHERE FCBA = '" & Session("fcBA") & "'" & _
                   "              AND TO_CHAR(ADHOCDATE, 'YYYYMM') = " & Year(Session("v_Employee")) & "||LPAD(" & Month(Session("v_Employee")) & ", 2, 0)" & _
                   "              AND ALLOWDEDCODE = '" & allowance_kerja_panen & "'" & _
                   "        GROUP BY FCBA" & _
                   "                 , TRUNC(ADHOCDATE, 'MM')" & _
                   "                 , EMPCODE" & _
                   " ) KP" & _
                   " LEFT JOIN ( SELECT FCBA" & _
                   "                    , TRUNC(ADHOCDATE, 'MM') ADHOCDATE" & _
                   "                    , EMPCODE" & _
                   "                    , NVL(SUM(AMOUNT), 0) AMOUNT" & _
                   "             FROM EMPALLDED" & _
                   "             WHERE FCBA = '" & Session("fcBA") & "'" & _
                   "                   AND TO_CHAR(ADHOCDATE, 'YYYYMM') = " & Year(Session("v_Employee")) & "||LPAD(" & Month(Session("v_Employee")) & ", 2, 0)" & _
                   "                   AND ALLOWDEDCODE = '" & allowance_gaji & "'" & _
                   "             GROUP BY FCBA" & _
                   "                      , TRUNC(ADHOCDATE, 'MM')" & _
                   "                      , EMPCODE" & _
                   " ) GJ" & _
                   " ON ( KP.FCBA = GJ.FCBA" & _
                   "      AND KP.ADHOCDATE = GJ.ADHOCDATE" & _
                   "      AND KP.EMPCODE = GJ.EMPCODE)"
				   
		' response.Write " sql " & qReceive
				   
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.Open qReceive, connstr
		IF NOT rsEmpAllDed.EOF THEN
			arr_data_kp = rsEmpAllDed.getrows()
			arr_data_kp_eof = false
		END IF
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
		
		IF NOT arr_data_kp_eof THEN
			FOR xxxx = LBound(arr_data_kp, 2) TO UBound(arr_data_kp, 2)
			
				amount_deduction_hk = 0
				
				IF "" & arr_data_kp(4, xxxx) = "" THEN
					amount_deduction_hk = 0
				ELSE
					amount_deduction_hk = ( Cdbl( arr_data_kp(4, xxxx) ) /Cdbl( arr_data_kp(5, xxxx) ) )* Cdbl( arr_data_kp(6, xxxx) )
				END IF
		
					maxLineNo = maxLineNo + 1
					qAllowance = " INSERT INTO empallded ( lineno" & _
								  "                               , empcode" & _
								  "                               , allowdedcode" & _
								  "                               , adhocdate" & _
								  "                               , amount" & _
								  "                               , fixedallded" & _
								  "                               , fcentry" & _
								  "                               , fcedit" & _
								  "                               , fcip" & _
								  "                               , fcba" & _
								  "                               , lastupdate" & _
								  "                               , lasttime)" & _
								  " values ( " & maxLineNo & "" & _
								  "          , '" & arr_data_kp(2, xxxx) & "'" & _
								  "          , '" & deduction_kerja_panen & "'" & _
								  "          , to_date('" & arr_data_kp(1, xxxx) & "','mm/dd/yyyy')" & _
								  "          , " & amount_deduction_hk & "" & _
								  "          , 0" & _
								  "          , '" & Session("CodeName") & "'" & _
								  "          , '" & Session("CodeName") & "'" & _
								  "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								  "          , '" & Session("fcBA") & "'" & _
								  "          , sysdate" & _
								  "          , to_char(SYSDATE, 'HH24:MI'))"
								  
					'response.Write "<br/>qAllowance : " & qAllowance
					connstr.execute(qAllowance)
		
			NEXT
		END IF
%>