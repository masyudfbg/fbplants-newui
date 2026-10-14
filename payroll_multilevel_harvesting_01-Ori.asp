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
		
		'private function RoundOracle(number_)
		'	result_ = 0
		'	q = " SELECT ROUND('"&number_&"', 0) amount FROM DUAL"
		'	' response.write q
		'	Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		'	rsEmpAllDed.Open q, connstr
		'	IF NOT rsEmpAllDed.EOF THEN
		'		result_ = cdbl(to_nol(rsEmpAllDed("amount")))
		'	END IF
		'	rsEmpAllDed.Close : set rsEmpAllDed = Nothing
			
		'	RoundOracle = result_
		'end function
		
		'arr_data_kp_eof = true
		
		'strMain = " select "& _
		'		  " fcba, "& _
		'		  " fddate, "& _
		'		  " employeecode,"& _
		'		  " max(amount) "& _
		'		  " from "& _	
		'		  " ( "& _
		'	      " select "& _
		'		  " fcba, "& _
		'		  " fddate, "& _
		'		  " employeecode,"& _
		'		  " case when amount >= 0 then amount else 0 end amount "& _
		'		  " from ( "& _
		'		  " select "& _
		'		  " hdb.fcba, "& _
		'		  " hdb.fddate, "& _
		'		  " hdb.employeecode, "& _
		'		  " pm.param_1, "& _
		'		  " hdb.extraamount_bydaily, "& _
		'		  " hdb.deductionamount_bydaily, "& _
		'		  " hdb.extraamount_bymonthly, "& _
		'		  " hdb.deductionamount_bymonthly, "& _
		'		  " case "& _
		'		  " when pm.param_1 = 'D' then (hdb.extraamount_bydaily+hdb.deductionamount_bydaily) "& _
 		'		  " else (hdb.extraamount_bymonthly+hdb.deductionamount_bymonthly) "& _
		'		  " end amount "& _
		'		  " from iplasprod.harvesting_daily_bjr hdb "& _
		'		  " inner join iplasprod.parameter pm on (hdb.fcba = pm.fcba)"& _
		'		  " where hdb.fcba ='" & Session("fcBA") & "'"& _
		'		  " and TO_CHAR(TO_DATE(FDDATE,'mm/dd/yyyy'),'fmMM') = " & Month(Session("v_Employee")) & _
		'		  " and TO_CHAR(TO_DATE(FDDATE,'mm/dd/yyyy'),'yyyy') = " & Year(Session("v_Employee")) & _
		'		  " and pm.parcode = 'HARDCODE'"& _
		'		  " and pm.fccode = 'JOBCODE_PANEN'"& _
		'		  " )"& _
		'		  " )"& _
		'		  " group by fcba,"& _
		'		  " fddate,"& _
		'		  " employeecode"
		'
		'response.Write " sql " & strMain

		'response.Write " Calculate Premi Panen ..." & "<br>" 
		'response.Flush()
		
		'Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		'rsEmpAllDed.Open strMain, connstr
		'IF NOT rsEmpAllDed.EOF THEN
		'	arr_data_kp = rsEmpAllDed.getrows()
		'	arr_data_kp_eof = false
		'END IF
		'rsEmpAllDed.Close : set rsEmpAllDed = Nothing
		
		'IF NOT arr_data_kp_eof THEN
		'	FOR xxxx=0 TO UBound(arr_data_kp, 2)
														
		'		PostAmountPanen = RoundOracle(arr_data_kp(3,xxxx))
		'		if PostAmountPanen > 0 then						
		'					maxLineNo = maxLineNo + 1
		'					qAllowance = " INSERT INTO empallded ( lineno" & _
		'								  "                               , empcode" & _
		'								  "                               , allowdedcode" & _
		'								  "                               , adhocdate" & _
		'								  "                               , amount" & _
		'								  "                               , fixedallded" & _
		'								  "                               , remarks" & _
		'								  "                               , fcentry" & _
		'								  "                               , fcedit" & _
		'								  "                               , fcip" & _
		'								  "                               , fcba" & _
		'								  "                               , lastupdate" & _
		'								  "                               , lasttime)" & _
		'								  " VALUES ( " & maxLineNo & "" & _
		'								  "          , '" & arr_data_kp(2, xxxx) & "'" & _
		'								  "          , '" & allowance_kerja_panen & "'" & _
		'								  "          , to_date('" & arr_data_kp(1, xxxx) & "','mm/dd/yyyy')" & _
		'								  "          , " & PostAmountPanen & "" & _
		'								  "          , 0" & _
		'								  "          , ''" & _
		'								  "          , '" & Session("CodeName") & "'" & _
		'								  "          , '" & Session("CodeName") & "'" & _
		'								  "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
		'								  "          , '" & Session("fcBA") & "'" & _
		'								  "          , SYSDATE" & _
		'								  "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
		'								  
		'					'response.Write "<br/>qAllowance : " & qAllowance
		'					connstr.execute(qAllowance)
		'		end if				
		'	NEXT
		'END IF
		
		response.Write " Calculate Premi Panen ..." & "<br>" 
		response.Flush()
		
		'**Premi Dapat Basis Insentif A27 **'			
		'**A27 -> Premi_capai_basis**'
		'**A24 -> premi_lebih_basis**'
		'**D19 -> potongan_tidak_dapat_basis**'
		'**A26 -> premi_brondolan**'
			
			qPremiInsentifBasis =	" SELECT fcba," & _
									"        fddate," & _
									"        employeecode," & _
									"        attendance," & _
									"        jobcode," & _
									"        division_final," & _
									"        planting_year_final," & _
									"        terraintype_final," & _
									"        presentgrade," & _
									"        presentrate," & _
									"        potongan_tidak_dapat_basis," & _
									"        total_janjang," & _
									"        total_ha," & _
									"        premi_capai_basis," & _
									"        jjg_lebih_basis," & _
									"        premi_level_1," & _
									"        premi_level_2," & _
									"        premi_level_3," & _
									"        premi_level_4," & _
									"        premi_lebih_basis," & _
									"        kg_brondolan_fix," & _
									"        (kg_brondolan_fix * rate_brondolan) AS premi_brondolan," & _
									"          presentrate" & _
									"        + premi_capai_basis" & _
									"        + premi_lebih_basis" & _
									"        - potongan_tidak_dapat_basis" & _
									"        + ( (kg_brondolan_fix * rate_brondolan))" & _
									"           AS total_pendapatan_panen" & _
									"   FROM (  SELECT fcba," & _
									"                  fddate," & _
									"                  employeecode," & _
									"                  attendance," & _
									"                  jobcode," & _
									"                  division_final," & _
									"                  planting_year_final," & _
									"                  terraintype_final," & _
									"                  presentgrade," & _
									"                  presentrate," & _
									"                  presentrate_fix," & _
									"                  CASE" & _
									"                     WHEN presentrate = presentrate_fix THEN 0" & _
									"                     ELSE presentrate - presentrate_fix" & _
									"                  END" & _
									"                     AS potongan_tidak_dapat_basis," & _
									"                  total_janjang," & _
									"                  total_ha," & _
									"                  premi_capai_basis," & _
									"                  jjg_lebih_basis," & _
									"                  premi_level_1," & _
									"                  premi_level_2," & _
									"                  premi_level_3," & _
									"                  premi_level_4," & _
									"                  (premi_level_1 + premi_level_2 + premi_level_3 + premi_level_4)" & _
									"                     AS premi_lebih_basis," & _
									"                  NVL (kg_brondolan, 0) AS kg_brondolan," & _
									"                  CASE" & _
									"                     WHEN     (presentrate = presentrate_fix)" & _
									"                          AND (premi_capai_basis > 0)" & _
									"                     THEN" & _
									"                        NVL (kg_brondolan, 0)" & _
									"                     ELSE" & _
									" 0" & _
									"                  END" & _
									"                     AS kg_brondolan_fix," & _
									"                  NVL (rate_brondolan_level_1, 0) AS rate_brondolan" & _
									"             FROM (SELECT gad_extra_basis.fcba," & _
									"                          gad_extra_basis.fddate," & _
									"                          gad_extra_basis.employeecode," & _
									"                          gad_extra_basis.attendance," & _
									"                          gad_extra_basis.jobcode," & _
									"                          gad_extra_basis.division_final," & _
									"                          gad_extra_basis.planting_year_final," & _
									"                          gad_extra_basis.terraintype_final," & _
									"                          gad_extra_basis.jml_baris," & _
									"                          gad_extra_basis.presentgrade," & _
									"                          gad_extra_basis.presentrate_original AS presentrate," & _
									"                          mhr_level.rate_nobasis_level_1," & _
									"                          gad_extra_basis.total_janjang," & _
									"                          gad_extra_basis.total_ha," & _
									"                          gad_extra_basis.max_ha," & _
									"                          gad_extra_basis.basis_ha_max_jjg," & _
									"                          gad_extra_basis.basis_ha_max_ha," & _
									"                          CASE" & _
									"                             WHEN     (gad_extra_basis.total_ha >=" & _
									"                                          gad_extra_basis.max_ha)" & _
									"                                  AND (DECODE (" & _
									"                                          gad_extra_basis.max_ha," & _
									"                                          0, 0," & _
									"                                            gad_extra_basis.total_ha" & _
									"                                          / gad_extra_basis.max_ha) <" & _
									"                                          gad_extra_basis.jml_baris)" & _
									"                                  AND (gad_extra_basis.jml_baris > 1)" & _
									"                             THEN" & _
									"                                (CASE" & _
									"                                    WHEN gad_extra_basis.total_ha >=" & _
									"                                            gad_extra_basis.basis_ha_max_ha" & _
									"                                    THEN" & _
									"                                       gad_extra_basis.presentrate_original" & _
									"                                    WHEN     (gad_extra_basis.total_ha <" & _
									"                                                 gad_extra_basis.basis_ha_max_ha)" & _
									"                                         AND (gad_extra_basis.total_janjang >=" & _
									"                                                 gad_extra_basis.basis_jjg_max_ha)" & _
									"                                    THEN" & _
									"                                         (  gad_extra_basis.total_ha" & _
									"                                          / gad_extra_basis.basis_ha_max_ha)" & _
									"                                       * gad_extra_basis.presentrate_original" & _
									"                                    ELSE" & _
									"                                       (  mhr_level.rate_nobasis_level_1" & _
									"                                        * gad_extra_basis.total_janjang)" & _
									"                                 END)" & _
									"                             WHEN     (gad_extra_basis.total_ha >=" & _
									"                                          gad_extra_basis.max_ha)" & _
									"                                  AND (DECODE (" & _
									"                                          gad_extra_basis.max_ha," & _
									"                                          0, 0," & _
									"                                            gad_extra_basis.total_ha" & _
									"                                          / gad_extra_basis.max_ha) =" & _
									"                                          gad_extra_basis.jml_baris)" & _
									"                                  AND (gad_extra_basis.jml_baris > 1)" & _
									"                             THEN" & _
									"                                (CASE" & _
									"                                    WHEN gad_extra_basis.total_ha >=" & _
									"                                            gad_extra_basis.basis_ha_max_jjg" & _
									"                                    THEN" & _
									"                                       gad_extra_basis.presentrate_original" & _
									"                                    WHEN     (gad_extra_basis.total_ha <" & _
									"                                                 gad_extra_basis.basis_ha_max_jjg)" & _
									"                                         AND (gad_extra_basis.total_janjang >=" & _
									"                                                 gad_extra_basis.basis_jjg_max_jjg)" & _
									"                                    THEN" & _
									"                                         (  gad_extra_basis.total_ha" & _
									"                                          / gad_extra_basis.basis_ha_max_jjg)" & _
									"                                       * gad_extra_basis.presentrate_original" & _
									"                                    ELSE" & _
									"                                       (  mhr_level.rate_nobasis_level_1" & _
									"                                        * gad_extra_basis.total_janjang)" & _
									"                                 END)" & _
									"                             WHEN     (gad_extra_basis.total_ha =" & _
									"                                          gad_extra_basis.max_ha)" & _
									"                                  AND (DECODE (" & _
									"                                          gad_extra_basis.max_ha," & _
									"                                          0, 0," & _
									"                                            gad_extra_basis.total_ha" & _
									"                                          / gad_extra_basis.max_ha) =" & _
									"                                          gad_extra_basis.jml_baris)" & _
									"                                  AND (gad_extra_basis.jml_baris = 1)" & _
									"                             THEN" & _
									"                                (CASE" & _
									"                                    WHEN gad_extra_basis.total_ha >=" & _
									"                                            gad_extra_basis.basis_ha_max_ha" & _
									"                                    THEN" & _
									"                                       gad_extra_basis.presentrate_original" & _
									"                                    WHEN     (gad_extra_basis.total_ha <" & _
									"                                                 gad_extra_basis.basis_ha_max_ha)" & _
									"                                         AND (gad_extra_basis.total_janjang >=" & _
									"                                                 gad_extra_basis.basis_jjg_max_ha)" & _
									"                                    THEN" & _
									"                                         (  gad_extra_basis.total_ha" & _
									"                                          / gad_extra_basis.basis_ha_max_ha)" & _
									"                                       * gad_extra_basis.presentrate_original" & _
									"                                    ELSE" & _
									"                                       (  mhr_level.rate_nobasis_level_1" & _
									"                                        * gad_extra_basis.total_janjang)" & _
									"                                 END)" & _
									"                             ELSE" & _
									" 0" & _
									"                          END" & _
									"                             AS presentrate_fix," & _
									"                          CASE" & _
									"                             WHEN gad_extra_basis.nfactor_basis = 1" & _
									"                             THEN" & _
									"                                mhr_level.standart_level_1" & _
									"                             ELSE" & _
									" 0" & _
									"                          END" & _
									"                             AS premi_capai_basis," & _
									"                          gad_extra_basis.jjg_lebih_basis," & _
									"                          CASE" & _
									"                             WHEN (gad_extra_basis.jjg_lebih_basis >" & _
									"                                      mhr_level.max_bunch_level_1)" & _
									"                             THEN" & _
									"                                  mhr_level.max_bunch_level_1" & _
									"                                * mhr_level.rate_level_1" & _
									"                             WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
									"                                          mhr_level.min_bunch_level_1)" & _
									"                                  AND (gad_extra_basis.jjg_lebih_basis <=" & _
									"                                          mhr_level.max_bunch_level_1)" & _
									"                             THEN" & _
									"                                  gad_extra_basis.jjg_lebih_basis" & _
									"                                * mhr_level.rate_level_1" & _
									"                             ELSE" & _
									" 0" & _
									"                          END" & _
									"                             AS premi_level_1," & _
									"                          CASE" & _
									"                             WHEN (gad_extra_basis.jjg_lebih_basis >" & _
									"                                      mhr_level.max_bunch_level_2)" & _
									"                             THEN" & _
									"                                  mhr_level.max_bunch_level_1" & _
									"                                * mhr_level.rate_level_2" & _
									"                             WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
									"                                          mhr_level.min_bunch_level_2)" & _
									"                                  AND (gad_extra_basis.jjg_lebih_basis <=" & _
									"                                          mhr_level.max_bunch_level_2)" & _
									"                             THEN" & _
									"                                  (  gad_extra_basis.jjg_lebih_basis" & _
									"                                   - mhr_level.max_bunch_level_1)" & _
									"                                * mhr_level.rate_level_2" & _
									"                             ELSE" & _
									" 0" & _
									"                          END" & _
									"                             AS premi_level_2," & _
									"                          CASE" & _
									"                             WHEN (gad_extra_basis.jjg_lebih_basis >" & _
									"                                      mhr_level.max_bunch_level_3)" & _
									"                             THEN" & _
									"                                  mhr_level.max_bunch_level_1" & _
									"                                * mhr_level.rate_level_3" & _
									"                             WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
									"                                          mhr_level.min_bunch_level_3)" & _
									"                                  AND (gad_extra_basis.jjg_lebih_basis <=" & _
									"                                          mhr_level.max_bunch_level_3)" & _
									"                             THEN" & _
									"                                  (  gad_extra_basis.jjg_lebih_basis" & _
									"                                   - mhr_level.max_bunch_level_2)" & _
									"                                * mhr_level.rate_level_3" & _
									"                             ELSE" & _
									" 0" & _
									"                          END" & _
									"                             AS premi_level_3," & _
									"                          CASE" & _
									"                             WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
									"                                          mhr_level.min_bunch_level_4)" & _
									"                                  AND (gad_extra_basis.jjg_lebih_basis <=" & _
									"                                          mhr_level.max_bunch_level_4)" & _
									"                             THEN" & _
									"                                  (  gad_extra_basis.jjg_lebih_basis" & _
									"                                   - mhr_level.max_bunch_level_3)" & _
									"                                * mhr_level.rate_level_4" & _
									"                             ELSE" & _
									" 0" & _
									"                          END" & _
									"                             AS premi_level_4," & _
									"                          brd.kg_brondolan," & _
									"                          mhr_level.rate_brondolan_level_1" & _
									"                     FROM (SELECT DISTINCT fcba," & _
									"                                           fddate," & _
									"                                           employeecode," & _
									"                                           jobcode," & _
									"                                           jml_baris," & _
									"                                           presentgrade," & _
									"                                           presentrate_original," & _
									"                                           attendance," & _
									"                                           total_janjang," & _
									"                                           total_ha," & _
									"                                           max_ha," & _
									"                                           basis_ha_max_jjg," & _
									"                                           basis_ha_max_ha," & _
									"                                           max_janjang," & _
									"                                           basis_jjg_max_jjg," & _
									"                                           basis_jjg_max_ha," & _
									"                                           division_final," & _
									"                                           planting_year_final," & _
									"                                           terraintype_final," & _
									"                                           jjg_lebih_basis," & _
									"                                           nfactor_basis" & _
									"                             FROM (SELECT gad.fcba," & _
									"                                          gad.fddate," & _
									"                                          gad.employeecode," & _
									"                                          gad.jobcode," & _
									"                                          gad.jml_baris," & _
									"                                          emp.presentgrade," & _
									"                                          emp.presentrate" & _
									"                                             AS presentrate_original," & _
									"                                          gad.attendance," & _
									"                                          gad.total_janjang," & _
									"                                          gad.total_ha," & _
									"                                          gad_max_jjg.max_janjang," & _
									"                                          gad_max_ha.max_ha," & _
									"                                          mhr_max_jjg.basis_jjg_max_jjg," & _
									"                                          mhr_max_jjg.basis_ha_max_jjg," & _
									"                                          mhr_max_ha.basis_jjg_max_ha," & _
									"                                          mhr_max_ha.basis_ha_max_ha," & _
									"                                          CASE" & _
									"                                             WHEN gad.total_ha >" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                gad_max_ha.division_max_ha" & _
									"                                             WHEN gad.total_ha =" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                (CASE" & _
									"                                                    WHEN gad.total_janjang >=" & _
									"                                                            gad_max_jjg.max_janjang" & _
									"                                                    THEN" & _
									"                                                       gad_max_jjg.division_max_jjg" & _
									"                                                    ELSE" & _
									"                                                       ''" & _
									"                                                 END)" & _
									"                                             ELSE" & _
									"                                                ''" & _
									"                                          END" & _
									"                                             AS division_final," & _
									"                                          CASE" & _
									"                                             WHEN gad.total_ha >" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                gad_max_ha.plantingyear_max_ha" & _
									"                                             WHEN gad.total_ha =" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                (CASE" & _
									"                                                    WHEN gad.total_janjang >=" & _
									"                                                            gad_max_jjg.max_janjang" & _
									"                                                    THEN" & _
									"                                                       gad_max_jjg.plantingyear_max_jjg" & _
									"                                                    ELSE" & _
									"                                                       ''" & _
									"                                                 END)" & _
									"                                             ELSE" & _
									"                                                ''" & _
									"                                          END" & _
									"                                             AS planting_year_final," & _
									"                                          CASE" & _
									"                                             WHEN gad.total_ha >" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                gad_max_ha.terraintype_max_ha" & _
									"                                             WHEN gad.total_ha =" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                (CASE" & _
									"                                                    WHEN gad.total_janjang >=" & _
									"                                                            gad_max_jjg.max_janjang" & _
									"                                                    THEN" & _
									"                                                       gad_max_jjg.terraintype_max_jjg" & _
									"                                                    ELSE" & _
									"                                                       ''" & _
									"                                                 END)" & _
									"                                             ELSE" & _
									"                                                ''" & _
									"                                          END" & _
									"                                             AS terraintype_final," & _
									"                                          CASE" & _
									"                                             WHEN gad.total_ha >" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                (CASE" & _
									"                                                    WHEN gad.total_janjang >" & _
									"                                                            mhr_max_ha.basis_jjg_max_ha" & _
									"                                                    THEN" & _
									"                                                         gad.total_janjang" & _
									"                                                       - mhr_max_ha.basis_jjg_max_ha" & _
									"                                                    ELSE" & _
									" 0" & _
									"                                                 END)" & _
									"                                             WHEN gad.total_ha =" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                (CASE" & _
									"                                                    WHEN gad.total_janjang >=" & _
									"                                                            gad_max_jjg.max_janjang" & _
									"                                                    THEN" & _
									"                                                       (CASE" & _
									"                                                           WHEN gad.total_janjang >" & _
									"                                                                   mhr_max_jjg.basis_jjg_max_jjg" & _
									"                                                           THEN" & _
									"                                                                gad.total_janjang" & _
									"                                                              - mhr_max_jjg.basis_jjg_max_jjg" & _
									"                                                           ELSE" & _
									" 0" & _
									"                                                        END)" & _
									"                                                    ELSE" & _
									" 0" & _
									"                                                 END)" & _
									"                                             ELSE" & _
									" 0" & _
									"                                          END" & _
									"                                             AS jjg_lebih_basis," & _
									"                                          CASE" & _
									"                                             WHEN gad.total_ha >" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                (CASE" & _
									"                                                    WHEN gad.total_janjang >=" & _
									"                                                            mhr_max_ha.basis_jjg_max_ha" & _
									"                                                    THEN" & _
									" 1" & _
									"                                                    ELSE" & _
									" 0" & _
									"                                                 END)" & _
									"                                             WHEN gad.total_ha =" & _
									"                                                     gad_max_ha.max_ha" & _
									"                                             THEN" & _
									"                                                (CASE" & _
									"                                                    WHEN gad.total_janjang >=" & _
									"                                                            gad_max_jjg.max_janjang" & _
									"                                                    THEN" & _
									"                                                       (CASE" & _
									"                                                           WHEN gad.total_janjang >=" & _
									"                                                                   mhr_max_jjg.basis_jjg_max_jjg" & _
									"                                                           THEN" & _
									" 1" & _
									"                                                           ELSE" & _
									" 0" & _
									"                                                        END)" & _
									"                                                    ELSE" & _
									" 0" & _
									"                                                 END)" & _
									"                                             ELSE" & _
									" 0" & _
									"                                          END" & _
									"                                             AS nfactor_basis" & _
									"                                     FROM (  SELECT attendance_gad.fcba," & _
									"                                                    attendance_gad.fddate," & _
									"                                                    attendance_gad.employeecode," & _
									"                                                    attendance_gad.jobcode," & _
									"                                                    attendance_gad.attendance," & _
									"                                                    ff.division," & _
									"                                                    ff.plantingyear," & _
									"                                                    ff.terraintype," & _
									"                                                       attendance_gad.fcba" & _
									"                                                    || attendance_gad.fddate" & _
									"                                                    || attendance_gad.employeecode" & _
									"                                                    || attendance_gad.jobcode" & _
									"                                                    || attendance_gad.attendance" & _
									"                                                       AS n," & _
									"                                                    SUM (" & _
									"                                                       attendance_gad.unit)" & _
									"                                                    OVER (" & _
									"                                                       PARTITION BY    attendance_gad.fcba" & _
									"                                                                    || attendance_gad.fddate" & _
									"                                                                    || attendance_gad.employeecode" & _
									"                                                                    || attendance_gad.jobcode" & _
									"                                                                    || attendance_gad.attendance)" & _
									"                                                       AS total_janjang," & _
									"                                                    SUM (" & _
									"                                                       attendance_gad.output)" & _
									"                                                    OVER (" & _
									"                                                       PARTITION BY    attendance_gad.fcba" & _
									"                                                                    || attendance_gad.fddate" & _
									"                                                                    || attendance_gad.employeecode" & _
									"                                                                    || attendance_gad.jobcode" & _
									"                                                                    || attendance_gad.attendance)" & _
									"                                                       AS total_ha," & _
									"                                                    COUNT (" & _
									"                                                       attendance_gad.fcba)" & _
									"                                                    OVER (" & _
									"                                                       PARTITION BY    attendance_gad.fcba" & _
									"                                                                    || attendance_gad.fddate" & _
									"                                                                    || attendance_gad.employeecode" & _
									"                                                                    || attendance_gad.jobcode" & _
									"                                                                    || attendance_gad.attendance)" & _
									"                                                       AS jml_baris" & _
									"                                               FROM attendance_gad" & _
									"                                                    INNER JOIN" & _
									"                                                    (SELECT field.fcba," & _
									"                                                            field.fccode," & _
									"                                                            field.division," & _
									"                                                            TO_CHAR (" & _
									"                                                               field.plantingdate," & _
									"                                                               'yyyy')" & _
									"                                                               AS plantingyear," & _
									"                                                            field.terraintype" & _
									"                                                       FROM field" & _
									"                                                      WHERE     field.fcba = '"& session("fcba") &"'" & _
									"                                                            AND field.activation =" & _
									"                                                                   'Y') ff" & _
									"                                                       ON (    attendance_gad.fcba =" & _
									"                                                                  ff.fcba" & _
									"                                                           AND attendance_gad.locationcode =" & _
									"                                                                  ff.fccode)" & _
									"                                              WHERE     attendance_gad.jobcode =" & _
									"                                                           '6502101'" & _
									"                                                    AND attendance_gad.attendance IN" & _
									"                                                           ('KJ'," & _
									"                                                            'KK'," & _
									"                                                            'WH'," & _
									"                                                            'WS')" & _
									"                                                    AND attendance_gad.fcba =" & _
									"                                                           '"& session("fcba") &"'" & _
									"                                           ORDER BY attendance_gad.fcba," & _
									"                                                    attendance_gad.fddate," & _
									"                                                    attendance_gad.employeecode," & _
									"                                                    attendance_gad.jobcode," & _
									"                                                    attendance_gad.attendance)" & _
									"                                          gad" & _
									"                                          INNER JOIN" & _
									"                                          (SELECT employee.fcba," & _
									"                                                  employee.fccode," & _
									"                                                  employee.presentgrade," & _
									"                                                  CASE" & _
									"                                                     WHEN employee.presentgrade IN" & _
									"                                                             ('KHT', 'KBT')" & _
									"                                                     THEN" & _
									"                                                          employee.presentrate" & _
									"                                                        / 25" & _
									"                                                     ELSE" & _
									"                                                        employee.presentrate" & _
									"                                                  END" & _
									"                                                     AS presentrate" & _
									"                                             FROM employee" & _
									"                                            WHERE     employee.fcba = '"& session("fcba") &"'" & _
									"                                                  AND employee.dateterminate" & _
									"                                                         IS NULL) emp" & _
									"                                             ON (    gad.fcba = emp.fcba" & _
									"                                                 AND gad.employeecode =" & _
									"                                                        emp.fccode)" & _
									"                                          INNER JOIN" & _
									"                                          (SELECT gad2.fcba," & _
									"                                                  gad2.fddate," & _
									"                                                  gad2.employeecode," & _
									"                                                  gad2.jobcode," & _
									"                                                  gad2.attendance," & _
									"                                                  gad2.locationcode," & _
									"                                                  ff.division" & _
									"                                                     AS division_max_jjg," & _
									"                                                  ff.plantingyear" & _
									"                                                     AS plantingyear_max_jjg," & _
									"                                                  ff.terraintype" & _
									"                                                     AS terraintype_max_jjg," & _
									"                                                  gad2.max_janjang," & _
									"                                                  gad2.linenokey" & _
									"                                             FROM (SELECT gad.fcba," & _
									"                                                          gad.fddate," & _
									"                                                          gad.employeecode," & _
									"                                                          gad.jobcode," & _
									"                                                          gad.attendance," & _
									"                                                          gad.locationcode," & _
									"                                                          gad.max_jjg" & _
									"                                                             AS max_janjang," & _
									"                                                          gad.linenokey" & _
									"                                                             linenokey," & _
									"                                                          ROW_NUMBER ()" & _
									"                                                          OVER (" & _
									"                                                             PARTITION BY    gad.fcba" & _
									"                                                                          || gad.fddate" & _
									"                                                                          || gad.employeecode" & _
									"                                                                          || gad.max_jjg" & _
									"                                                             ORDER BY" & _
									"                                                                gad.linenokey ASC)" & _
									"                                                             AS minimal" & _
									"                                                     FROM (SELECT gad.fcba," & _
									"                                                                  gad.fddate," & _
									"                                                                  gad.employeecode," & _
									"                                                                  gad.jobcode," & _
									"                                                                  gad.attendance," & _
									"                                                                  gad.locationcode," & _
									"                                                                  gad_max_jjg.max_jjg" & _
									"                                                                     AS max_jjg," & _
									"                                                                  gad.linenokey" & _
									"                                                             FROM attendance_gad" & _
									"                                                                  gad" & _
									"                                                                  INNER JOIN" & _
									"                                                                  (  SELECT fcba," & _
									"                                                                            fddate," & _
									"                                                                            employeecode," & _
									"                                                                            jobcode," & _
									"                                                                            MAX (" & _
									"                                                                               unit)" & _
									"                                                                               AS max_jjg" & _
									"                                                                       FROM attendance_gad" & _
									"                                                                      WHERE     jobcode =" & _
									"                                                                                   '6502101'" & _
									"                                                                            AND attendance IN" & _
									"                                                                                   ('KJ'," & _
									"                                                                                    'KK'," & _
									"                                                                                    'WH'," & _
									"                                                                                    'WS')" & _
									"                                                                            AND fcba =" & _
									"                                                                                   '"& session("fcba") &"'" & _
									"                                                                   GROUP BY fcba," & _
									"                                                                            fddate," & _
									"                                                                            employeecode," & _
									"                                                                            jobcode)" & _
									"                                                                  gad_max_jjg" & _
									"                                                                     ON (    gad.fcba =" & _
									"                                                                                gad_max_jjg.fcba" & _
									"                                                                         AND gad.fddate =" & _
									"                                                                                gad_max_jjg.fddate" & _
									"                                                                         AND gad.employeecode =" & _
									"                                                                                gad_max_jjg.employeecode" & _
									"                                                                         AND gad.jobcode =" & _
									"                                                                                gad_max_jjg.jobcode" & _
									"                                                                         AND gad.unit =" & _
									"                                                                                gad_max_jjg.max_jjg))" & _
									"                                                          gad) gad2" & _
									"                                                  INNER JOIN" & _
									"                                                  (SELECT field.fcba," & _
									"                                                          field.fccode," & _
									"                                                          field.division," & _
									"                                                          TO_CHAR (" & _
									"                                                             field.plantingdate," & _
									"                                                             'yyyy')" & _
									"                                                             AS plantingyear," & _
									"                                                          field.terraintype" & _
									"                                                     FROM field" & _
									"                                                    WHERE     field.fcba = '"& session("fcba") &"'" & _
									"                                                          AND field.activation =" & _
									"                                                                 'Y') ff" & _
									"                                                     ON (    gad2.fcba = ff.fcba" & _
									"                                                         AND gad2.locationcode =" & _
									"                                                                ff.fccode)" & _
									"                                            WHERE gad2.minimal = 1) gad_max_jjg" & _
									"                                             ON (    gad.fcba = gad_max_jjg.fcba" & _
									"                                                 AND gad.fddate =" & _
									"                                                        gad_max_jjg.fddate" & _
									"                                                 AND gad.employeecode =" & _
									"                                                        gad_max_jjg.employeecode" & _
									"                                                 AND gad.jobcode =" & _
									"                                                        gad_max_jjg.jobcode)" & _
									"                                          INNER JOIN" & _
									"                                          (SELECT gad2.fcba," & _
									"                                                  gad2.fddate," & _
									"                                                  gad2.employeecode," & _
									"                                                  gad2.jobcode," & _
									"                                                  gad2.attendance," & _
									"                                                  gad2.locationcode," & _
									"                                                  ff.division AS division_max_ha," & _
									"                                                  ff.plantingyear" & _
									"                                                     AS plantingyear_max_ha," & _
									"                                                  ff.terraintype" & _
									"                                                     AS terraintype_max_ha," & _
									"                                                  gad2.max_ha," & _
									"                                                  gad2.linenokey" & _
									"                                             FROM (SELECT gad.fcba," & _
									"                                                          gad.fddate," & _
									"                                                          gad.employeecode," & _
									"                                                          gad.jobcode," & _
									"                                                          gad.attendance," & _
									"                                                          gad.locationcode," & _
									"                                                          gad.max_ha AS max_ha," & _
									"                                                          gad.linenokey" & _
									"                                                             linenokey," & _
									"                                                          ROW_NUMBER ()" & _
									"                                                          OVER (" & _
									"                                                             PARTITION BY    gad.fcba" & _
									"                                                                          || gad.fddate" & _
									"                                                                          || gad.employeecode" & _
									"                                                                          || gad.max_ha" & _
									"                                                             ORDER BY" & _
									"                                                                gad.linenokey ASC)" & _
									"                                                             AS minimal" & _
									"                                                     FROM (SELECT gad.fcba," & _
									"                                                                  gad.fddate," & _
									"                                                                  gad.employeecode," & _
									"                                                                  gad.jobcode," & _
									"                                                                  gad.attendance," & _
									"                                                                  gad.locationcode," & _
									"                                                                  gad_max_ha.max_ha" & _
									"                                                                     AS max_ha," & _
									"                                                                  gad.linenokey" & _
									"                                                             FROM attendance_gad" & _
									"                                                                  gad" & _
									"                                                                  INNER JOIN" & _
									"                                                                  (  SELECT fcba," & _
									"                                                                            fddate," & _
									"                                                                            employeecode," & _
									"                                                                            jobcode," & _
									"                                                                            MAX (" & _
									"                                                                               output)" & _
									"                                                                               AS max_ha" & _
									"                                                                       FROM attendance_gad" & _
									"                                                                      WHERE     jobcode =" & _
									"                                                                                   '6502101'" & _
									"                                                                            AND attendance IN" & _
									"                                                                                   ('KJ'," & _
									"                                                                                    'KK'," & _
									"                                                                                    'WH'," & _
									"                                                                                    'WS')" & _
									"                                                                            AND fcba =" & _
									"                                                                                   '"& session("fcba") &"'" & _
									"                                                                   GROUP BY fcba," & _
									"                                                                            fddate," & _
									"                                                                            employeecode," & _
									"                                                                            jobcode)" & _
									"                                                                  gad_max_ha" & _
									"                                                                     ON (    gad.fcba =" & _
									"                                                                                gad_max_ha.fcba" & _
									"                                                                         AND gad.fddate =" & _
									"                                                                                gad_max_ha.fddate" & _
									"                                                                         AND gad.employeecode =" & _
									"                                                                                gad_max_ha.employeecode" & _
									"                                                                         AND gad.jobcode =" & _
									"                                                                                gad_max_ha.jobcode" & _
									"                                                                         AND gad.output =" & _
									"                                                                                gad_max_ha.max_ha))" & _
									"                                                          gad) gad2" & _
									"                                                  INNER JOIN" & _
									"                                                  (SELECT field.fcba," & _
									"                                                          field.fccode," & _
									"                                                          field.division," & _
									"                                                          TO_CHAR (" & _
									"                                                             field.plantingdate," & _
									"                                                             'yyyy')" & _
									"                                                             AS plantingyear," & _
									"                                                          field.terraintype" & _
									"                                                     FROM field" & _
									"                                                    WHERE     field.fcba = '"& session("fcba") &"'" & _
									"                                                          AND field.activation =" & _
									"                                                                 'Y') ff" & _
									"                                                     ON (    gad2.fcba = ff.fcba" & _
									"                                                         AND gad2.locationcode =" & _
									"                                                                ff.fccode)" & _
									"                                            WHERE gad2.minimal = 1) gad_max_ha" & _
									"                                             ON (    gad.fcba = gad_max_ha.fcba" & _
									"                                                 AND gad.fddate =" & _
									"                                                        gad_max_ha.fddate" & _
									"                                                 AND gad.employeecode =" & _
									"                                                        gad_max_ha.employeecode" & _
									"                                                 AND gad.jobcode =" & _
									"                                                        gad_max_ha.jobcode)" & _
									"                                          INNER JOIN" & _
									"                                          (SELECT DISTINCT" & _
									"                                                  multilevel_harvesting_rate.fcba," & _
									"                                                  multilevel_harvesting_rate.employeetype," & _
									"                                                  multilevel_harvesting_rate.attendance," & _
									"                                                  multilevel_harvesting_rate.planting_year," & _
									"                                                  multilevel_harvesting_rate.terrain_type," & _
									"                                                  multilevel_harvesting_rate.basis" & _
									"                                                     AS basis_jjg_max_jjg," & _
									"                                                  multilevel_harvesting_rate.basis_ha" & _
									"                                                     AS basis_ha_max_jjg" & _
									"                                             FROM multilevel_harvesting_rate" & _
									"                                            WHERE multilevel_harvesting_rate.fcba =" & _
									"                                                     '"& session("fcba") &"') mhr_max_jjg" & _
									"                                             ON (    gad.fcba = mhr_max_jjg.fcba" & _
									"                                                 AND emp.presentgrade =" & _
									"                                                        mhr_max_jjg.employeetype" & _
									"                                                 AND gad.attendance =" & _
									"                                                        mhr_max_jjg.attendance" & _
									"                                                 AND gad_max_jjg.plantingyear_max_jjg =" & _
									"                                                        mhr_max_jjg.planting_year" & _
									"                                                 AND gad_max_jjg.terraintype_max_jjg =" & _
									"                                                        mhr_max_jjg.terrain_type)" & _
									"                                          INNER JOIN" & _
									"                                          (SELECT DISTINCT" & _
									"                                                  multilevel_harvesting_rate.fcba," & _
									"                                                  multilevel_harvesting_rate.employeetype," & _
									"                                                  multilevel_harvesting_rate.attendance," & _
									"                                                  multilevel_harvesting_rate.planting_year," & _
									"                                                  multilevel_harvesting_rate.terrain_type," & _
									"                                                  multilevel_harvesting_rate.basis" & _
									"                                                     AS basis_jjg_max_ha," & _
									"                                                  multilevel_harvesting_rate.basis_ha" & _
									"                                                     AS basis_ha_max_ha" & _
									"                                             FROM multilevel_harvesting_rate" & _
									"                                            WHERE multilevel_harvesting_rate.fcba =" & _
									"                                                     '"& session("fcba") &"') mhr_max_ha" & _
									"                                             ON (    gad_max_ha.fcba =" & _
									"                                                        mhr_max_ha.fcba" & _
									"                                                 AND emp.presentgrade =" & _
									"                                                        mhr_max_ha.employeetype" & _
									"                                                 AND gad.attendance =" & _
									"                                                        mhr_max_ha.attendance" & _
									"                                                 AND gad_max_ha.plantingyear_max_ha =" & _
									"                                                        mhr_max_ha.planting_year" & _
									"                                                 AND gad_max_ha.terraintype_max_ha =" & _
									"                                                        mhr_max_ha.terrain_type)))" & _
									"                          gad_extra_basis" & _
									"                          LEFT JOIN" & _
									"                          (  SELECT fcba," & _
									"                                    employeetype," & _
									"                                    attendance," & _
									"                                    division," & _
									"                                    planting_year," & _
									"                                    terrain_type," & _
									"                                    SUM (DECODE (levelno, 1, min_bunch, 0))" & _
									"                                       min_bunch_level_1," & _
									"                                    SUM (DECODE (levelno, 2, min_bunch, 0))" & _
									"                                       min_bunch_level_2," & _
									"                                    SUM (DECODE (levelno, 3, min_bunch, 0))" & _
									"                                       min_bunch_level_3," & _
									"                                    SUM (DECODE (levelno, 4, min_bunch, 0))" & _
									"                                       min_bunch_level_4," & _
									"                                    SUM (DECODE (levelno, 1, max_bunch, 0))" & _
									"                                       max_bunch_level_1," & _
									"                                    SUM (DECODE (levelno, 2, max_bunch, 0))" & _
									"                                       max_bunch_level_2," & _
									"                                    SUM (DECODE (levelno, 3, max_bunch, 0))" & _
									"                                       max_bunch_level_3," & _
									"                                    SUM (DECODE (levelno, 4, max_bunch, 0))" & _
									"                                       max_bunch_level_4," & _
									"                                    SUM (DECODE (levelno, 1, standart, 0))" & _
									"                                       standart_level_1," & _
									"                                    SUM (DECODE (levelno, 2, standart, 0))" & _
									"                                       standart_level_2," & _
									"                                    SUM (DECODE (levelno, 3, standart, 0))" & _
									"                                       standart_level_3," & _
									"                                    SUM (DECODE (levelno, 4, standart, 0))" & _
									"                                       standart_level_4," & _
									"                                    SUM (DECODE (levelno, 1, rate, 0))" & _
									"                                       rate_level_1," & _
									"                                    SUM (DECODE (levelno, 2, rate, 0))" & _
									"                                       rate_level_2," & _
									"                                    SUM (DECODE (levelno, 3, rate, 0))" & _
									"                                       rate_level_3," & _
									"                                    SUM (DECODE (levelno, 4, rate, 0))" & _
									"                                       rate_level_4," & _
									"                                    SUM (DECODE (levelno, 1, rate_brondolan, 0))" & _
									"                                       rate_brondolan_level_1," & _
									"                                    SUM (DECODE (levelno, 2, rate_brondolan, 0))" & _
									"                                       rate_brondolan_level_2," & _
									"                                    SUM (DECODE (levelno, 3, rate_brondolan, 0))" & _
									"                                       rate_brondolan_level_3," & _
									"                                    SUM (DECODE (levelno, 4, rate_brondolan, 0))" & _
									"                                       rate_brondolan_level_4," & _
									"                                    SUM (DECODE (levelno, 1, rate_nobasis, 0))" & _
									"                                       rate_nobasis_level_1," & _
									"                                    SUM (DECODE (levelno, 2, rate_nobasis, 0))" & _
									"                                       rate_nobasis_level_2," & _
									"                                    SUM (DECODE (levelno, 3, rate_nobasis, 0))" & _
									"                                       rate_nobasis_level_3," & _
									"                                    SUM (DECODE (levelno, 4, rate_nobasis, 0))" & _
									"                                       rate_nobasis_level_4" & _
									"                               FROM (SELECT iplasprod.multilevel_harvesting_rate.fcba," & _
									"                                            iplasprod.multilevel_harvesting_rate.levelno," & _
									"                                            iplasprod.multilevel_harvesting_rate.employeetype," & _
									"                                            iplasprod.multilevel_harvesting_rate.attendance," & _
									"                                            iplasprod.multilevel_harvesting_rate.division," & _
									"                                            iplasprod.multilevel_harvesting_rate.planting_year," & _
									"                                            iplasprod.multilevel_harvesting_rate.terrain_type," & _
									"                                            iplasprod.multilevel_harvesting_rate.min_bunch," & _
									"                                            iplasprod.multilevel_harvesting_rate.max_bunch," & _
									"                                            iplasprod.multilevel_harvesting_rate.standart," & _
									"                                            iplasprod.multilevel_harvesting_rate.rate," & _
									"                                            iplasprod.multilevel_harvesting_rate.rate_brondolan," & _
									"                                            iplasprod.multilevel_harvesting_rate.rate_nobasis" & _
									"                                       FROM iplasprod.multilevel_harvesting_rate)" & _
									"                              WHERE fcba = '"& session("fcba") &"'" & _
									"                           GROUP BY fcba," & _
									"                                    employeetype," & _
									"                                    attendance," & _
									"                                    division," & _
									"                                    planting_year," & _
									"                                    terrain_type) mhr_level" & _
									"                             ON (    gad_extra_basis.fcba = mhr_level.fcba" & _
									"                                 AND gad_extra_basis.presentgrade =" & _
									"                                        mhr_level.employeetype" & _
									"                                 AND gad_extra_basis.attendance =" & _
									"                                        mhr_level.attendance" & _
									"                                 AND gad_extra_basis.division_final =" & _
									"                                        mhr_level.division" & _
									"                                 AND gad_extra_basis.planting_year_final =" & _
									"                                        mhr_level.planting_year" & _
									"                                 AND gad_extra_basis.terraintype_final =" & _
									"                                        mhr_level.terrain_type)" & _
									"                          LEFT JOIN" & _
									"                          (SELECT fcba," & _
									"                                  fddate," & _
									"                                  employeecode," & _
									"                                  attendance," & _
									"                                  jobcode," & _
									"                                  division," & _
									"                                  terraintype," & _
									"                                  plantingyear," & _
									"                                  SUM (" & _
									"                                     kg_brondolan)" & _
									"                                  OVER (" & _
									"                                     PARTITION BY    fcba" & _
									"                                                  || fddate" & _
									"                                                  || employeecode" & _
									"                                                  || attendance)" & _
									"                                     AS kg_brondolan" & _
									"                             FROM (  SELECT attendance_gad.fcba," & _
									"                                            attendance_gad.fddate," & _
									"                                            attendance_gad.employeecode," & _
									"                                            attendance_gad.jobcode," & _
									"                                            attendance_gad.attendance," & _
									"                                            ff.division," & _
									"                                            ff.plantingyear," & _
									"                                            ff.terraintype," & _
									"                                            SUM (attendance_gad.kg_brondolan)" & _
									"                                               AS kg_brondolan" & _
									"                                       FROM attendance_gad" & _
									"                                            INNER JOIN" & _
									"                                            (SELECT field.fcba," & _
									"                                                    field.fccode," & _
									"                                                    field.division," & _
									"                                                    TO_CHAR (field.plantingdate," & _
									"                                                             'yyyy')" & _
									"                                                       AS plantingyear," & _
									"                                                    field.terraintype" & _
									"                                               FROM field" & _
									"                                              WHERE     field.fcba = '"& session("fcba") &"'" & _
									"                                                    AND field.activation = 'Y') ff" & _
									"                                               ON (    attendance_gad.fcba =" & _
									"                                                          ff.fcba" & _
									"                                                   AND attendance_gad.locationcode =" & _
									"                                                          ff.fccode)" & _
									"                                      WHERE     attendance_gad.jobcode = '6502101'" & _
									"                                            AND attendance_gad.attendance IN" & _
									"                                                   ('KJ'," & _
									"                                                    'KK'," & _
									"                                                    'WH'," & _
									"                                                    'WS')" & _
									"                                            AND attendance_gad.fcba = '"& session("fcba") &"'" & _
									"                                   GROUP BY attendance_gad.fcba," & _
									"                                            attendance_gad.fddate," & _
									"                                            attendance_gad.employeecode," & _
									"                                            attendance_gad.jobcode," & _
									"                                            attendance_gad.attendance," & _
									"                                            ff.division," & _
									"                                            ff.plantingyear," & _
									"                                            ff.terraintype)" & _
									"                            WHERE kg_brondolan > 0) brd" & _
									"                             ON (    gad_extra_basis.fcba = brd.fcba" & _
									"                                 AND gad_extra_basis.fddate = brd.fddate" & _
									"                                 AND gad_extra_basis.employeecode =" & _
									"                                        brd.employeecode" & _
									"                                 AND gad_extra_basis.attendance = brd.attendance" & _
									"                                 AND gad_extra_basis.jobcode = brd.jobcode" & _
									"                                 AND gad_extra_basis.division_final =" & _
									"                                        brd.division" & _
									"                                 AND gad_extra_basis.planting_year_final =" & _
									"                                        brd.plantingyear" & _
									"                                 AND gad_extra_basis.terraintype_final =" & _
									"                                        brd.terraintype))" & _
									"            WHERE     TO_NUMBER (TO_CHAR (fddate, 'yyyy')) = '"& year(session("v_Employee")) &"'" & _
									"                  AND TO_NUMBER (TO_CHAR (fddate, 'mm')) = '"& Month(session("v_Employee")) &"'" & _
									"         ORDER BY fcba, fddate, employeecode)" 
			
			'response.Write " sql " & qPremiInsentifBasis &"<br>"
			response.Write " Start Premi Insentif ..." & "<br>" 
			response.Flush()
			Set rsPremiInsentifBasis = Server.CreateObject("ADODB.Recordset")
			rsPremiInsentifBasis.CursorType = 3
			rsPremiInsentifBasis.CursorLocation = 3
			rsPremiInsentifBasis.Open qPremiInsentifBasis, connstr
			if not rsPremiInsentifBasis.eof then
			  Do While Not rsPremiInsentifBasis.eof
			  'response.write "Masuk"
				 '**A27 -> Premi_capai_basis**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("PREMI_CAPAI_BASIS"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "A27" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 'response.Write "A27 : " & qAllowance & "<br>"
					 connstr.execute(qAllowance)
				  end if
				'** End A27 -> Premi_capai_basis**'
				
				'**A24 -> premi_lebih_basis**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("PREMI_LEBIH_BASIS"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "A24" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 'response.Write "A24 : " & qAllowance & "<br>"
					 connstr.execute(qAllowance)
				  end if
				'** End A24 -> premi_lebih_basis**'
				
				'**D19 -> potongan_tidak_dapat_basis**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("POTONGAN_TIDAK_DAPAT_BASIS"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "D19" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 'response.Write "D19 : " & qAllowance & "<br>"
					 connstr.execute(qAllowance)
				  end if
				'**D19 -> potongan_tidak_dapat_basis**'
				
				'**A26 -> premi_brondolan**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("PREMI_BRONDOLAN"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "A26" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 'response.Write "A26 : " & qAllowance & "<br>"
					 connstr.execute(qAllowance)
				  end if
				'**A26 -> premi_brondolan**'
				
				rsPremiInsentifBasis.MoveNext
			  loop
			End if
			rsPremiInsentifBasis.Close
			set rsPremiInsentifBasis = Nothing
			response.Write " Finish Premi Insentif ..." & "<br>" 
			response.Flush()
		'**Premi Tidak Dapat Basis A27 Insentif **'
%>