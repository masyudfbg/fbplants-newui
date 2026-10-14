<%

		' **** Konsep Panen Borongan **** '
		' Dikarenakan konsep alokasi employee iPlaS memakai upah rata-rata, sehingga proses perhitungan panen borongan di iPlaS tetap memakai absensi KJ
		' maka akan terbentuk gaji 1 hk, tetapi gaji 1 hk tersebut langsung di potong memakai potongan tidak dapat basis (D19) sehingga gaji menjadi 0
		' Dan upah borongan panen dalam 1 hari memakai allowded Premi Lebih Borong (A24) dengan konsep perhitungan JJG x BJR x Rp/Ton
		' Kondisi memotong upah, yaitu: sama dengan A01 yang didapatkan karyawan


		allowance_gaji = "A01"
		allowance_kerja_panen = "A24"
		deduction_kerja_panen = "D19"
		
		allowance_gaji = "A01"
		
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
		strMain = " SELECT ag.fcba" & _
                  "        , ag.fddate" & _
                  "        , ag.employeecode" & _
                  "        , ag.jobcode" & _
                  "        , ag.locationtype" & _
                  "        , ag.locationcode" & _
                  "        , NVL(ag.unit, 0) unit" & _
				  "        , em.presentgrade" & _
                  "        , NVL(em.presentrate, 0) presentrate" & _
				  "        , field_api.getValid_Afdeling(ag.locationcode, ag.fcba, ag.fddate) division" & _
				  "        , field_api.getValid_Plantingyear(ag.locationcode, ag.fcba, ag.fddate) plantingdate" & _
				  "        , field_api.GetValid_terraintype(ag.locationcode, ag.fcba, ag.fddate) terraintype" & _
                  " FROM attendance_gad ag" & _
                  " INNER JOIN employee em ON (ag.employeecode = em.fccode AND ag.fcba = em.fcba)" & _
                  " WHERE em.dateterminate IS NULL" & _
                  "       AND ag.attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
                  "       AND ag.jobcode "&wherestmt&"" & _
                  "       AND ag.locationtype = 'FF'" & _
                  "       AND ag.mandays > 0" & _
                  "       AND ag.unit > 0" & _
                  "       AND TO_NUMBER(TO_CHAR(ag.fddate,'mm')) = " & Month(Session("v_Employee")) & "" & _
                  "       AND TO_NUMBER(TO_CHAR(ag.fddate,'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                  "       AND ag.fcba = '" & Session("fcBA") & "'" & _
                  " ORDER BY ag.fcba" & _
                  "          , ag.fddate" & _
                  "          , ag.employeecode" & _
                  "          , ag.linenokey"
				
		'response.Write " sql " & strMain
		response.Write "Calculate Panen ..." & "<br>" 
		response.Flush()
		
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.Open strMain, connstr
		IF NOT rsEmpAllDed.EOF THEN
			arr_data_kp = rsEmpAllDed.getrows()
			arr_data_kp_eof = false
		END IF
		rsEmpAllDed.Close : set rsEmpAllDed = Nothing
		
		IF NOT arr_data_kp_eof THEN
			FirstRecord = 0 : BunchTodate = 0 : TotalBunch = 0 : Max_level = 0 : CurrentLevel = 0 : TotalAmountPanen = 0 : PostAmountPanen = 0
			FOR xxxx=0 TO UBound(arr_data_kp, 2)
				FirstRecord = FirstRecord +1
				CurrentBunch = 0
				
				CurrentBunch = cdbl(to_nol(arr_data_kp(6,xxxx)))
				BunchTodate = BunchTodate +CurrentBunch 
				TotalBunch = TotalBunchPerDay(arr_data_kp, arr_data_kp(1,xxxx), arr_data_kp(2,xxxx))
				
				if FirstRecord = 1 then
					FirstDivision = arr_data_kp(9,xxxx)
					FirstPlantingYear = arr_data_kp(10,xxxx)
					FirstTerrainType = arr_data_kp(11,xxxx)
					
					q = " SELECT NVL(levelno, 0) levelno" & _
						" FROM multilevel_harvesting_rate" & _
						" WHERE fcba = '" & Session("fcBA") & "'" & _
						"       AND (cropcode = 'ALL' OR cropcode = 'OP')" & _
						"       AND (division = 'ALL' OR division = '"&FirstDivision&"')" & _
						"       AND (planting_year = 'ALL' OR planting_year = '"&FirstPlantingYear&"')" & _
						"       AND (terrain_type = 'ALL' OR terrain_type = '"&FirstTerrainType&"')" & _
						"       AND '"&TotalBunch&"' BETWEEN min_bunch AND max_bunch"
					' response.write q
					Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
					rsEmpAllDed.Open q, connstr
					IF NOT rsEmpAllDed.EOF THEN
						Max_level = cdbl(to_nol(rsEmpAllDed("levelno")))
					END IF
					rsEmpAllDed.Close : set rsEmpAllDed = Nothing
					
					CurrentLevel = CurrentLevel +1
				end if
				
				' response.write "<br/>"
				' response.write "empcode: "&arr_data_kp(2,xxxx)
				' response.write "| fddate: "&arr_data_kp(1,xxxx)
				' response.write "| blok: "&arr_data_kp(5,xxxx)
				' response.write "| tt: "&arr_data_kp(11,xxxx)
				' response.write "| acuan tt: "&FirstPlantingYear
				' response.write "| jjg: "&CurrentBunch
				' response.write "| jjg todate: "&BunchTodate
				' response.Flush()
				
				rate_level = 0 : HitungBunch = 0
				q = " SELECT NVL(rate, 0) rate" & _
					" FROM multilevel_harvesting_rate" & _
					" WHERE fcba = '" & Session("fcBA") & "'" & _
					"       AND (cropcode = 'ALL' OR cropcode = 'OP')" & _
					"       AND (division = 'ALL' OR division = '"&arr_data_kp(9,xxxx)&"')" & _
					"       AND (planting_year = 'ALL' OR planting_year = '"&arr_data_kp(10,xxxx)&"')" & _
					"       AND (terrain_type = 'ALL' OR terrain_type = '"&arr_data_kp(11,xxxx)&"')" & _
					"       AND levelno = '"&CurrentLevel&"'" 
				' response.write q
				Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
				rsEmpAllDed.Open q, connstr
				IF NOT rsEmpAllDed.EOF THEN
					rate_level = cdbl(to_nol(rsEmpAllDed("rate")))
				END IF
				rsEmpAllDed.Close : set rsEmpAllDed = Nothing
				
				q = " SELECT NVL(min_bunch, 0) min_bunch, NVL(max_bunch, 0) max_bunch" & _
					" FROM multilevel_harvesting_rate" & _
					" WHERE fcba = '" & Session("fcBA") & "'" & _
					"       AND (cropcode = 'ALL' OR cropcode = 'OP')" & _
					"       AND (division = 'ALL' OR division = '"&FirstDivision&"')" & _
					"       AND (planting_year = 'ALL' OR planting_year = '"&FirstPlantingYear&"')" & _
					"       AND (terrain_type = 'ALL' OR terrain_type = '"&FirstTerrainType&"')" & _
					"       AND levelno = '"&CurrentLevel&"'"
				' response.write q
				Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
				rsEmpAllDed.Open q, connstr
				IF NOT rsEmpAllDed.EOF THEN
					MinBunch = cdbl(to_nol(rsEmpAllDed("min_bunch")))
					MaxBunch = cdbl(to_nol(rsEmpAllDed("max_bunch")))
				END IF
				rsEmpAllDed.Close : set rsEmpAllDed = Nothing
				
				if BunchTodate > MaxBunch then
					HitungBunch = MaxBunch
					' response.write "| Hitung: "&HitungBunch&" x "&rate_level
					' response.Flush()
					TotalAmountPanen = TotalAmountPanen +( cdbl(to_nol(HitungBunch)) *rate_level )
					CurrentLevel = CurrentLevel +1
				else
					HitungBunch = (MinBunch -1) +BunchTodate
					' response.write "| Hitung: "&HitungBunch&" x "&rate_level
					' response.Flush()
					TotalAmountPanen = TotalAmountPanen +( cdbl(to_nol(HitungBunch)) *rate_level )
				end if
				
				
				if FirstRecord = TotalRecordPerDay(arr_data_kp, arr_data_kp(1,xxxx), arr_data_kp(2,xxxx)) then
					
					' response.write "| TotalAmountPanen: "&TotalAmountPanen
					' response.Flush()
					
					Max_Gaji = 0 : PostJenis = ""
					if arr_data_kp(7,xxxx)="KBT" then
						Max_Gaji = RoundOracle( cdbl(to_nol(arr_data_kp(8,xxxx))) /30 )
					elseif arr_data_kp(7,xxxx)="KHT" then
						Max_Gaji = RoundOracle( cdbl(to_nol(arr_data_kp(8,xxxx))) /25 )
					else
						Max_Gaji = cdbl(to_nol(arr_data_kp(8,xxxx)))
					end if
					
					if TotalAmountPanen > Max_Gaji then
						PostAmountPanen = TotalAmountPanen -Max_Gaji
						PostJenis = "insert"
					else
						PostAmountPanen = TotalAmountPanen
						PostJenis = "update"
					end if
					' response.write "| PostAmountPanen: "&PostAmountPanen
					' response.Flush()
					
					if PostAmountPanen > 0 then
						if PostJenis = "insert" then
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
										  "          , to_date('" & arr_data_kp(1, xxxx) & "','mm/dd/yyyy')" & _
										  "          , " & PostAmountPanen & "" & _
										  "          , 0" & _
										  "          , '" & arr_data_kp(3, xxxx) & "'" & _
										  "          , '" & Session("CodeName") & "'" & _
										  "          , '" & Session("CodeName") & "'" & _
										  "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										  "          , '" & Session("fcBA") & "'" & _
										  "          , SYSDATE" & _
										  "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
										  
							'response.Write "<br/>qAllowance : " & qAllowance
							connstr.execute(qAllowance)
						elseif PostJenis = "update" then
							qAllowance = " UPDATE empallded" & _
										 " SET amount = '" & PostAmountPanen & "'" & _
										 "     , remarks = '" & arr_data_kp(3, xxxx) & "'" & _
										 "     , lastupdate = SYSDATE" & _
										 "     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
										 " WHERE adhocdate = TO_DATE('" & arr_data_kp(1, xxxx) & "','mm/dd/yyyy')" & _
										 "       AND allowdedcode = '" & allowance_gaji & "'" & _
										 "       AND empcode = '" & arr_data_kp(2, xxxx) & "'" & _
										 "       AND fcba = '" & Session("fcBA") & "'"
										  
							'response.Write "<br/>qAllowance : " & qAllowance
							connstr.execute(qAllowance)
						end if
					end if
					
					FirstRecord = 0 : BunchTodate = 0 : TotalBunch = 0 : CurrentLevel = 0 : TotalAmountPanen = 0 : PostAmountPanen = 0
				end if
			NEXT
		END IF
%>