<%

		' **** Konsep Panen Borongan **** '
		' Dikarenakan konsep alokasi employee iPlaS memakai upah rata-rata, sehingga proses perhitungan panen borongan di iPlaS tetap memakai absensi KJ
		' maka akan terbentuk gaji 1 hk, tetapi gaji 1 hk tersebut langsung di potong memakai potongan tidak dapat basis (D19) sehingga gaji menjadi 0
		' Dan upah borongan panen dalam 1 hari memakai allowded Premi Lebih Borong (A24) dengan konsep perhitungan JJG x BJR x Rp/Ton
		' Kondisi memotong upah, yaitu: sama dengan A01 yang didapatkan karyawan

		allowance_gaji = "A01"
		allowance_kerja_panen = "A24"
		deduction_kerja_panen = "D19"
		allowance_kerajinan = "A23"
		
		'Dim tglnow
		tglnow = Session("tglnow")
		
		
		response.Write " Calculate Premi Panen ..." & "<br>" 
		response.Flush()
		
		'**Premi Dapat Basis Insentif A27 **'			
		'**A27 -> Premi_capai_basis**'
		'**A24 -> premi_lebih_basis**'
		'**D19 -> potongan_tidak_dapat_basis**'
		'**A26 -> premi_brondolan**'
		
		qReference = " SELECT NVL( PARAMETER_API.GetParam1('BASIS', 'PAYROLL_PANEN', '" & Session("fcBA") & "' ), 'JJG') OPTIONS FROM DUAL"

		Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
		rsReferenceReceive.CursorType = 3
		rsReferenceReceive.CursorLocation = 3
		rsReferenceReceive.Open qReference, connstr

		if not rsReferenceReceive.eof then
			Default_Basis = rsReferenceReceive("OPTIONS")
		else
			Default_Basis = "JJG"
		end if
		rsReferenceReceive.close
		set rsReferenceReceive = Nothing
		
		connstr.BeginTrans
		qNormalisasiDetail = 	" DELETE FROM employee_premi_panen_tab  " & _
								" WHERE fcba = '" & Session("fcBA") & "' " & _
								" AND fddate=to_date('" & tglnow & "','mm/dd/yyyy')  "
								
				
		connstr.execute(qNormalisasiDetail)
				
		
			
		qNormalisasiDetail =	"INSERT INTO employee_premi_panen_tab(  " & _
								"GANGCODE, 		    " & _        
								"FDDATE,            " & _
								"SUPERVISION_1,     " & _
								"SUPERVISION_2,     " & _
								"SUPERVISION_3,     " & _
								"SUPERVISION_4,     " & _
								"EMPLOYEECODE,      " & _
								"EMPLOYEENAME,      " & _
								"ATTENDANCE,        " & _
								"JOBCODE,           " & _
								"LOCATIONTYPE,      " & _
								"LOCATIONCODE,      " & _
								"TOT_UNIT,          " & _
								"TOT_MANDAYS,       " & _
								"TOT_HA,		    " & _
								"UNIT,              " & _
								"MANDAYS,           " & _
								"HA,                " & _
								"BASIS,             " & _
								"PRESENTRATE,       " & _
								"RATE_HK,           " & _ 
								"PREMI_CAPAI_BASIS, " & _
								"LEBIH_BASIS,       " & _
								"PREMI_1_BASIS,     " & _
								"PREMI_1_RATE,      " & _
								"PREMI_1,           " & _
								"PREMI_2_BASIS,     " & _
								"PREMI_2_RATE,      " & _
								"PREMI_2,           " & _
								"PREMI_3_BASIS,     " & _
								"PREMI_3_RATE,      " & _
								"PREMI_3,           " & _
								"PREMI_4_BASIS,     " & _
								"PREMI_4_RATE,      " & _
								"PREMI_4,           " & _
								"PREMI_LEBIH_BASIS, " & _
								"PREMI_BONUS,       " & _
								"PREMI_KERAJINAN,   " & _
								"PREMI_TOTAL,       " & _
								"POT_BASIS,         " & _
								"POT_DENDA,         " & _
								"POT_TOTAL,         " & _
								"BRONDOLAN_KG,      " & _
								"BRONDOLAN_RP,      " & _
								"BRONDOLAN_TOTAL,   " & _
								"TOTAL,             " & _
								"FCBA,              " & _
								"LASTUPDATE,        " & _
								"LASTTIME,          " & _
								"LINENOKEY          " & _
								")                  " & _
								"SELECT  " & _
								"gangcode, " & _
								"fddate, " & _
								"supervision_1, " & _
								"supervision_2, " & _
								"supervision_3, " & _
								"supervision_4, " & _
								"employeecode, " & _
								"employee_api.get_employeename(fcba,employeecode)employeename, " & _
								"attendance, " & _
								"jobcode, " & _
								"'FF' locationtype, " & _
								"blok, " & _
								"total_janjang TOT_UNIT, " & _
								"presentrate_fix_hk TOT_MANDAYS, " & _
								"total_ha TOT_HA, " & _
								"qtyunit UNIT, " & _
								"hk_final MANDAYS, " & _
								"0 HA, " & _
								"basis_jjg_max_jjg BASIS, " & _
								"presentrate, " & _
								"Round(hk_final*presentrate,0)As RateHK, " & _
								"Round(hk_final*premi_capai_basis,0)premi_capai_basis, " & _
								"jjg_lebih_basis, " & _
								"0 premi_basis_1, " & _
								"0 premi_rate_1, " & _
								"Round(hk_final*premi_level_1,0)premi_level_1, " & _
								"0 premi_basis_2, " & _
								"0 premi_rate_2, " & _
								"Round(hk_final*premi_level_2,0)premi_level_2, " & _
								"0 premi_basis_3, " & _
								"0 premi_rate_3, " & _
								"Round(hk_final*premi_level_3,0)premi_level_3, " & _
								"0 premi_basis_4, " & _
                                "0 premi_rate_4, " & _
								"Round(hk_final*premi_level_4,0)premi_level_4, " & _
								"Round(hk_final*premi_lebih_basis,0)premi_lebih_basis, " & _
								"Round(hk_final*bonus_premi,0)bonus_premi, " & _
								"0 premi_kerajinan, " & _
								"Round((hk_final*premi_capai_basis),0) " & _
								"+Round((hk_final*premi_lebih_basis),0) " & _
								"+Round((hk_final*bonus_premi),0)AS Premi_total, " & _
								"0 Pot_basis, " & _
								"0 Pot_Denda, " & _
								"0 Pot_total, " & _
								"kg_brondolan_fix, " & _
								"rate_brondolan, " & _
								"(kg_brondolan_fix*rate_brondolan)AS brondolan_total, " & _
								"Round((hk_final*presentrate),0) " & _
								"+Round((hk_final*premi_capai_basis),0) " & _
								"+Round((hk_final*premi_lebih_basis),0) " & _
								"+Round((hk_final*bonus_premi),0) " & _
								"+(kg_brondolan_fix*rate_brondolan) " & _
								"AS total, " & _
								"fcba, " & _
								"SYSDATE lastupdate, " & _
								"'00:00' lasttime, " & _
								"Linenokey  " & _
								"FROM( SELECT  " & _
								"supervision_1, " & _
								"supervision_2, " & _
								"supervision_3, " & _
								"supervision_4, " & _
								"gangcode, " & _
								" " & _
								"fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"attendance, " & _
								"linenokey, " & _
								" " & _
								"jobcode, " & _
								"blok, " & _
								"qtyunit, " & _
								"Round((qtyunit/total_janjang)*presentrate_fix_hk,2)Hk_final, " & _
								"division_final, " & _
								"planting_year_final, " & _
								"terraintype_final, " & _
								"presentgrade, " & _
								"presentrate, " & _
								"presentrate_fix, " & _
								" " & _
								"presentrate_fix_jjg, " & _
								"presentrate_fix_hk, " & _
								" " & _
								"CASE " & _
								"WHEN presentrate=presentrate_fix_jjg THEN 0 " & _
								"ELSE presentrate-presentrate_fix_jjg " & _
								"END " & _
								"AS potongan_tidak_dapat_basis, " & _
								" " & _
								"total_janjang, " & _
								"total_ha, " & _
								"premi_capai_basis, " & _
								"jjg_lebih_basis, " & _
								"premi_level_1, " & _
								"premi_level_2, " & _
								"premi_level_3, " & _
								"premi_level_4, " & _
								"(premi_level_1+premi_level_2+premi_level_3+premi_level_4)AS premi_lebih_basis, " & _
								"NVL(kg_brondolan,0)AS kg_brondolan, " & _
								"/*AMH-UPDJAN22BrondolanNonbasispanen*/ " & _
								"/*CASE*/ " & _
								"/*WHEN(presentrate=presentrate_fix)*/ " & _
								"/*WHEN(presentrate=presentrate_fix_jjg)*/ " & _
								"/*AND(premi_capai_basis>0)*/ " & _
								"/*THEN*/ " & _
								"/*NVL(kg_brondolan,0)*/ " & _
								"/*ELSE*/ " & _
								"/*0*/ " & _
								"/*END*/ " & _
								"NVL(kg_brondolan,0)AS kg_brondolan_fix, " & _
								"NVL(rate_brondolan_level_1,0)AS rate_brondolan, " & _
								"/*AMH-ADDAPR21BonusPremiperApr21*/ " & _
								"/*CASEWHENTO_CHAR(fddate,'yyyymm')>='202104'THEN*/ " & _
								"/*CASEWHENplanting_year_final<=2013THEN*/ " & _
								"/*CASEWHENtotal_janjang>=166THEN15500ELSE0END*/ " & _
								"/*ELSE*/ " & _
								"/*CASEWHENtotal_janjang>=233THEN15500ELSE0END*/ " & _
								"/*END*/ " & _
								"/*ELSE0END*/ " & _
								"CASE WHEN total_janjang>=bonus_basis then bonus_premi ELSE 0 END AS Bonus_premi, " & _
								"basis_jjg_max_jjg  " & _
								"FROM  " & _
								"(SELECT  " & _
								"gad_extra_basis.supervision_1, " & _
								"gad_extra_basis.supervision_2, " & _
								"gad_extra_basis.supervision_3, " & _
								"gad_extra_basis.supervision_4, " & _
								"gad_extra_basis.gangcode, " & _
								" " & _
								"gad_extra_basis.fcba, " & _
								"gad_extra_basis.fddate, " & _
								"gad_extra_basis.employeecode, " & _
								"gad_extra_basis.attendance, " & _
								" " & _
								"gad_extra_basis.linenokey, " & _
								"gad_extra_basis.jobcode, " & _
								" " & _
								"gad_extra_basis.blok, " & _
								"gad_extra_basis.qtyunit, " & _
								" " & _
								"gad_extra_basis.division_final, " & _
								"gad_extra_basis.planting_year_final, " & _
								"gad_extra_basis.terraintype_final, " & _
								"gad_extra_basis.jml_baris, " & _
								"gad_extra_basis.presentgrade, " & _
								"gad_extra_basis.presentrate_original AS presentrate, " & _
								"mhr_level.rate_nobasis_level_1, " & _
								"gad_extra_basis.total_janjang, " & _
								"gad_extra_basis.total_ha, " & _
								"gad_extra_basis.max_ha, " & _
								"gad_extra_basis.basis_ha_max_jjg, " & _
								"gad_extra_basis.basis_ha_max_ha, " & _
								"gad_extra_basis.basis_jjg_max_jjg, " & _
								" " & _
								"/*AMHBasisJjg*/ " & _
								"CASE  " & _
								"WHEN(gad_extra_basis.total_janjang>=gad_extra_basis.basis_jjg_max_jjg) THEN " & _
								"gad_extra_basis.presentrate_original  " & _
								"ELSE  " & _
								"round((gad_extra_basis.total_janjang/gad_extra_basis.basis_jjg_max_jjg)*gad_extra_basis.presentrate_original,0)  " & _
								"END  " & _
								"AS presentrate_fix_jjg, " & _
								" " & _
								"/*AMHHKPanen*/ " & _
								"CASE  " & _
								"WHEN(gad_extra_basis.total_janjang>=gad_extra_basis.basis_jjg_max_jjg) THEN  " & _
								"1 " & _
								"ELSE  " & _
								"round((gad_extra_basis.total_janjang/gad_extra_basis.basis_jjg_max_jjg),2)  " & _
								"END " & _
								"AS presentrate_fix_hk, " & _
								" " & _
								"CASE  " & _
								"WHEN(gad_extra_basis.total_ha>=gad_extra_basis.max_ha) " & _
								"AND(gad_extra_basis.total_ha/gad_extra_basis.max_ha < gad_extra_basis.jml_baris) " & _
								"AND(gad_extra_basis.jml_baris>1) THEN " & _
								"(CASE  " & _
								"WHEN gad_extra_basis.total_ha>=gad_extra_basis.basis_ha_max_ha " & _
								"THEN  " & _
								"gad_extra_basis.presentrate_original " & _
								"WHEN (gad_extra_basis.total_ha<gad_extra_basis.basis_ha_max_ha) " & _
								"AND (gad_extra_basis.total_janjang>=gad_extra_basis.basis_jjg_max_ha) " & _
								"THEN  " & _
								"(gad_extra_basis.total_ha/gad_extra_basis.basis_ha_max_ha)*gad_extra_basis.presentrate_original " & _
								"ELSE  " & _
								"(mhr_level.rate_nobasis_level_1*gad_extra_basis.total_janjang) " & _
								"END  " & _
								") " & _
								"WHEN (gad_extra_basis.total_ha>=gad_extra_basis.max_ha) " & _
								"AND (gad_extra_basis.total_ha/gad_extra_basis.max_ha=gad_extra_basis.jml_baris) " & _
								"AND (gad_extra_basis.jml_baris>1)THEN " & _
								"(CASE  " & _
								"WHEN gad_extra_basis.total_ha>=gad_extra_basis.basis_ha_max_jjg " & _
								"THEN  " & _
								"gad_extra_basis.presentrate_original " & _
								"WHEN (gad_extra_basis.total_ha<gad_extra_basis.basis_ha_max_jjg) " & _
								"AND (gad_extra_basis.total_janjang>=gad_extra_basis.basis_jjg_max_jjg) " & _
								"THEN  " & _
								"(gad_extra_basis.total_ha/gad_extra_basis.basis_ha_max_jjg)*gad_extra_basis.presentrate_original " & _
								"ELSE  " & _
								"(mhr_level.rate_nobasis_level_1*gad_extra_basis.total_janjang) " & _
								"END  " & _
								") " & _
								"WHEN (gad_extra_basis.total_ha=gad_extra_basis.max_ha) " & _
								"AND (gad_extra_basis.total_ha/gad_extra_basis.max_ha=gad_extra_basis.jml_baris) " & _
								"AND (gad_extra_basis.jml_baris=1)THEN " & _
								"(CASE  " & _
								"WHEN gad_extra_basis.total_ha>=gad_extra_basis.basis_ha_max_ha " & _
								"THEN  " & _
								"gad_extra_basis.presentrate_original " & _
								"WHEN (gad_extra_basis.total_ha<gad_extra_basis.basis_ha_max_ha) " & _
								"AND (gad_extra_basis.total_janjang>=gad_extra_basis.basis_jjg_max_ha) " & _
								"THEN  " & _
								"(gad_extra_basis.total_ha/gad_extra_basis.basis_ha_max_ha)*gad_extra_basis.presentrate_original " & _
								"ELSE  " & _
								"(mhr_level.rate_nobasis_level_1*gad_extra_basis.total_janjang) " & _
								"END  " & _
								") " & _
								"ELSE  " & _
								"0  " & _
								"END  " & _
								"AS presentrate_fix, " & _
								" " & _
								"CASE  " & _
								"WHEN gad_extra_basis.nfactor_basis=1 THEN " & _
								"mhr_level.standart_level_1 " & _
								"ELSE  " & _
								"0 " & _
								"END  " & _
								"AS premi_capai_basis, " & _
								" " & _
								"gad_extra_basis.jjg_lebih_basis, " & _
								" " & _
								"CASE  " & _
								"WHEN (gad_extra_basis.jjg_lebih_basis>mhr_level.max_bunch_level_1)  " & _
								"THEN  " & _
								"mhr_level.max_bunch_level_1*mhr_level.rate_level_1 " & _
								"WHEN (gad_extra_basis.jjg_lebih_basis>=mhr_level.min_bunch_level_1) " & _
								"AND (gad_extra_basis.jjg_lebih_basis<=mhr_level.max_bunch_level_1) " & _
								"THEN " & _
								"gad_extra_basis.jjg_lebih_basis*mhr_level.rate_level_1 " & _
								"ELSE " & _
								"0 " & _
								"END " & _
								"AS premi_level_1, " & _
								" " & _
								"CASE " & _
								"WHEN (gad_extra_basis.jjg_lebih_basis>mhr_level.max_bunch_level_2) " & _
								"THEN " & _
								"mhr_level.max_bunch_level_1*mhr_level.rate_level_2 " & _
								"WHEN (gad_extra_basis.jjg_lebih_basis>=mhr_level.min_bunch_level_2) " & _
								"AND (gad_extra_basis.jjg_lebih_basis<=mhr_level.max_bunch_level_2) " & _
								"THEN " & _
								"(gad_extra_basis.jjg_lebih_basis-mhr_level.max_bunch_level_1)*mhr_level.rate_level_2 " & _
								"ELSE " & _
								"0 " & _
								"END " & _
								"AS premi_level_2, " & _
								" " & _
								"CASE " & _
								"WHEN (gad_extra_basis.jjg_lebih_basis>mhr_level.max_bunch_level_3) " & _
								"THEN " & _
								"mhr_level.max_bunch_level_1*mhr_level.rate_level_3 " & _
								"WHEN (gad_extra_basis.jjg_lebih_basis>=mhr_level.min_bunch_level_3) " & _
								"AND(gad_extra_basis.jjg_lebih_basis<=mhr_level.max_bunch_level_3) " & _
								"THEN  " & _
								"(gad_extra_basis.jjg_lebih_basis-mhr_level.max_bunch_level_2)*mhr_level.rate_level_3 " & _
								"ELSE  " & _
								"0  " & _
								"END " & _
								"AS premi_level_3, " & _
								" " & _
								"CASE  " & _
								"WHEN(gad_extra_basis.jjg_lebih_basis>=mhr_level.min_bunch_level_4) " & _
								"AND(gad_extra_basis.jjg_lebih_basis<=mhr_level.max_bunch_level_4) " & _
								"THEN  " & _
								"(gad_extra_basis.jjg_lebih_basis-mhr_level.max_bunch_level_3)*mhr_level.rate_level_4 " & _
								"ELSE  " & _
								"0  " & _
								"END  " & _
								"AS premi_level_4, " & _
								" " & _
								"brd.kg_brondolan, " & _
								"mhr_level.rate_brondolan_level_1, " & _
								"mhr_level.bonus_basis,mhr_level.bonus_premi " & _
								"FROM (SELECT DISTINCT  " & _
								"supervision_1, " & _
								"supervision_2, " & _
								"supervision_3, " & _
								"supervision_4, " & _
								"gangcode, " & _
								" " & _
								"fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"jobcode, " & _
								"blok, " & _
								"qtyunit, " & _
								"jml_baris, " & _
								"presentgrade, " & _
								"presentrate_original, " & _
								"attendance, " & _
								" " & _
								"linenokey, " & _
								" " & _
								"total_janjang, " & _
								"total_ha, " & _
								"max_ha, " & _
								"basis_ha_max_jjg, " & _
								"basis_ha_max_ha, " & _
								"max_janjang, " & _
								"basis_jjg_max_jjg, " & _
								"basis_jjg_max_ha, " & _
								"division_final, " & _
								"planting_year_final, " & _
								"terraintype_final, " & _
								"jjg_lebih_basis, " & _
								"nfactor_basis " & _
								"FROM(SELECT  " & _
								"gad.supervision_1, " & _
								"gad.supervision_2, " & _
								"gad.supervision_3, " & _
								"gad.supervision_4, " & _
								"gad.gangcode, " & _
								" " & _
								"gad.fcba, " & _
								"gad.fddate, " & _
								"gad.employeecode, " & _
								"gad.jobcode, " & _
								"gad.blok, " & _
								"gad.qtyunit, " & _
								"gad.jml_baris, " & _
								"emp.presentgrade, " & _
								"emp.presentrate AS presentrate_original, " & _
								"gad.attendance, " & _
								"gad.linenokey, " & _
								"gad.total_janjang, " & _
								"gad.total_ha, " & _
								"gad_max_jjg.max_janjang, " & _
								"gad_max_ha.max_ha, " & _
								"mhr_max_jjg.basis_jjg_max_jjg, " & _
								"mhr_max_jjg.basis_ha_max_jjg, " & _
								"mhr_max_ha.basis_jjg_max_ha, " & _
								"mhr_max_ha.basis_ha_max_ha, " & _
								"CASE " & _
								"WHEN gad.total_ha>gad_max_ha.max_ha " & _
								"THEN " & _
								"gad_max_ha.division_max_ha " & _
								"WHEN gad.total_ha=gad_max_ha.max_ha " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>=gad_max_jjg.max_janjang " & _
								"THEN " & _
								"gad_max_jjg.division_max_jjg " & _
								"ELSE " & _
								"'' " & _
								"END) " & _
								"ELSE " & _
								"'' " & _
								"END " & _
								"AS division_final, " & _
								" " & _
								"CASE " & _
								"WHEN gad.total_ha>gad_max_ha.max_ha " & _
								"THEN " & _
								"gad_max_ha.plantingyear_max_ha " & _
								"WHEN gad.total_ha=gad_max_ha.max_ha " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>=gad_max_jjg.max_janjang " & _
								"THEN  " & _
								"gad_max_jjg.plantingyear_max_jjg " & _
								"ELSE " & _
								"'' " & _
								"END) " & _
								"ELSE " & _
								"'' " & _
								"END " & _
								"AS planting_year_final, " & _
								" " & _
								"CASE " & _
								"WHEN gad.total_ha>gad_max_ha.max_ha " & _
								"THEN " & _
								"gad_max_ha.terraintype_max_ha " & _
								"WHEN gad.total_ha=gad_max_ha.max_ha " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>=gad_max_jjg.max_janjang " & _
								"THEN " & _
								"gad_max_jjg.terraintype_max_jjg " & _
								"ELSE " & _
								"'' " & _
								"END) " & _
								"ELSE " & _
								"'' " & _
								"END " & _
								"AS terraintype_final, " & _
								" " & _
								"CASE " & _
								"WHEN gad.total_ha>gad_max_ha.max_ha " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>mhr_max_ha.basis_jjg_max_ha " & _
								"THEN " & _
								"gad.total_janjang-mhr_max_ha.basis_jjg_max_ha " & _
								"ELSE " & _
								"0 " & _
								"END) " & _
								"WHEN gad.total_ha=gad_max_ha.max_ha " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>=gad_max_jjg.max_janjang " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>mhr_max_jjg.basis_jjg_max_jjg " & _
								"THEN " & _
								"gad.total_janjang-mhr_max_jjg.basis_jjg_max_jjg " & _
								"ELSE " & _
								"0 " & _
								"END) " & _
								"ELSE " & _
								"0 " & _
								"END) " & _
								"ELSE " & _
								"0 " & _
								"END " & _
								"AS jjg_lebih_basis, " & _
								" " & _
								"CASE " & _
								"WHEN gad.total_ha>gad_max_ha.max_ha " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>=mhr_max_ha.basis_jjg_max_ha " & _
								"THEN " & _
								"1 " & _
								"ELSE " & _
								"0 " & _
								"END) " & _
								"WHEN gad.total_ha=gad_max_ha.max_ha " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>=gad_max_jjg.max_janjang " & _
								"THEN " & _
								"(CASE " & _
								"WHEN gad.total_janjang>=mhr_max_jjg.basis_jjg_max_jjg " & _
								"THEN " & _
								"1 " & _
								"ELSE " & _
								"0 " & _
								"END) " & _
								"ELSE " & _
								"0 " & _
								"END) " & _
								"ELSE " & _
								"0 " & _
								"END " & _
								"AS nfactor_basis " & _
								"FROM(SELECT  " & _
								"attendance_gad.supervision_1, " & _
								"attendance_gad.supervision_2, " & _
								"attendance_gad.supervision_3, " & _
								"attendance_gad.supervision_4, " & _
								"attendance_gad.gangcode, " & _
								" " & _
								"attendance_gad.fcba, " & _
								"attendance_gad.fddate, " & _
								"attendance_gad.employeecode, " & _
								"attendance_gad.jobcode, " & _
								"attendance_gad.attendance, " & _
								"attendance_gad.linenokey, " & _
								"ff.fccode blok, " & _
								"CASE WHEN Upper('" & Default_Basis & "')='JJG' THEN attendance_gad.unit ELSE attendance_gad.bunch_weight END Qtyunit, " & _
								"ff.division, " & _
								"ff.plantingyear, " & _
								"ff.terraintype, " & _
								"attendance_gad.fcba " & _
								"||attendance_gad.fddate " & _
								"||attendance_gad.employeecode " & _
								"||attendance_gad.jobcode " & _
								"||attendance_gad.attendance " & _
								"AS n, " & _
								"SUM(CASE WHEN Upper('" &Default_Basis & "')='JJG' THEN attendance_gad.unit ELSE attendance_gad.bunch_weight END) " & _
								"OVER( " & _
								"PARTITION BY attendance_gad.fcba " & _
								"||attendance_gad.fddate " & _
								"||attendance_gad.employeecode " & _
								"||attendance_gad.jobcode " & _
								"||attendance_gad.attendance) " & _
								"AS total_janjang, " & _
								"SUM(attendance_gad.output) " & _
								"OVER( " & _
								"PARTITION BY attendance_gad.fcba " & _
								"||attendance_gad.fddate " & _
								"||attendance_gad.employeecode " & _
								"||attendance_gad.jobcode " & _
								"||attendance_gad.attendance) " & _
								"AS total_ha, " & _
								"COUNT(attendance_gad.fcba) " & _
								"OVER( " & _
								"PARTITION BY attendance_gad.fcba " & _
								"||attendance_gad.fddate " & _
								"||attendance_gad.employeecode " & _
								"||attendance_gad.jobcode " & _
								"||attendance_gad.attendance) " & _
								"AS jml_baris " & _
								"FROM attendance_gad " & _
								"INNER JOIN " & _
								"(SELECT field.fcba, " & _
								"field.fccode, " & _
								"field.division, " & _
								"TO_CHAR( " & _
								"field.plantingdate, " & _
								"'yyyy') " & _
								"AS plantingyear, " & _
								"field.terraintype " & _
								"FROM field " & _
								"WHERE field.fcba='"& session("fcba") &"' " & _
								"AND field.activation='Y' " & _
								")ff " & _
								"ON(attendance_gad.fcba=ff.fcba " & _
								"AND attendance_gad.locationcode=ff.fccode) " & _
								"WHERE attendance_gad.jobcode in('6502101','1321410') " & _
								"AND attendance_gad.attendance IN('KJ','KK','WH','WS') " & _
								"AND attendance_gad.fcba='"& session("fcba") &"' " & _
								"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy')  " & _
								"ORDER BY attendance_gad.fcba, " & _
								"/*GROUPBYattendance_gad.fcba,*/ " & _
								"attendance_gad.fddate, " & _
								"attendance_gad.employeecode, " & _
								"attendance_gad.jobcode, " & _
								"attendance_gad.attendance/*, " & _
								"ff.fccode " & _
								"/*ff.division,*/ " & _
								"/*ff.plantingyear,*/ " & _
								"/*ff.terraintype,*/ " & _
								"/*attendance_gad.fcba*/ " & _
								"/*||attendance_gad.fddate*/ " & _
								"/*||attendance_gad.employeecode*/ " & _
								"/*||attendance_gad.jobcode*/ " & _
								"/*||attendance_gad.attendance*/ " & _
								") " & _
								"gad " & _
								"INNER JOIN " & _
								"(SELECT employee.fcba, " & _
								"employee.fccode, " & _
								"employee.presentgrade, " & _
								"CASE " & _
								"WHEN employee.presentgrade IN ('KHT','KBT') " & _
								"THEN " & _
								"employee.presentrate/25 " & _
								"ELSE " & _
								"employee.presentrate " & _
								"END " & _
								"AS presentrate " & _
								"FROM employee " & _
								"WHERE employee.fcba='"& session("fcba") &"' " & _
								"AND (employee.dateterminate IS NULL  " & _
								"  OR employee.dateterminate = to_date('" & tglnow & "','mm/dd/yyyy') ) " & _
								")emp " & _
								"ON(gad.fcba=emp.fcba " & _
								"AND gad.employeecode=emp.fccode) " & _
								"INNER JOIN " & _
								"(SELECT gad2.fcba, " & _
								"gad2.fddate, " & _
								"gad2.employeecode, " & _
								"gad2.jobcode, " & _
								"gad2.attendance, " & _
								"gad2.locationcode, " & _
								"ff.division AS division_max_jjg, " & _
								"ff.plantingyear AS plantingyear_max_jjg, " & _
								"ff.terraintype AS terraintype_max_jjg, " & _
								"gad2.max_janjang, " & _
								"gad2.linenokey " & _
								"FROM(SELECT gad.fcba, " & _
								"gad.fddate, " & _
								"gad.employeecode, " & _
								"gad.jobcode, " & _
								"gad.attendance, " & _
								"gad.locationcode, " & _
								"gad.max_jjg AS max_janjang, " & _
								"gad.linenokey linenokey, " & _
								"ROW_NUMBER() " & _
								"OVER( " & _
								"PARTITION BY gad.fcba " & _
								"||gad.fddate " & _
								"||gad.employeecode " & _
								"||gad.max_jjg " & _
								"ORDER BY " & _
								"gad.linenokey ASC) " & _
								"AS minimal " & _
								"FROM(SELECT gad.fcba, " & _
								"gad.fddate, " & _
								"gad.employeecode, " & _
								"gad.jobcode, " & _
								"gad.attendance, " & _
								"gad.locationcode, " & _
								"gad_max_jjg.max_jjg AS max_jjg, " & _
								"gad.linenokey " & _
								"FROM/*attendance_gad*/ " & _
								"(SELECT " & _
								"fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"jobcode, " & _
								"attendance, " & _
								"locationcode, " & _
								"linenokey, " & _
								"CASE WHEN Upper('" &Default_Basis & "')='JJG' THEN unit ELSE bunch_weight END unit " & _
								"FROM " & _
								"attendance_gad " & _
								"WHERE jobcode in('6502101','1321410') " & _
								"AND attendance IN('KJ','KK','WH','WS') " & _
								"AND fcba='"& session("fcba") &"' " & _
								"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy')  " & _
								") " & _
								"gad " & _
								"INNER JOIN " & _
								"(SELECT fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"jobcode, " & _
								"MAX(CASE WHEN Upper('" &Default_Basis & "')='JJG' THEN unit ELSE bunch_weight END)AS max_jjg " & _
								"FROM attendance_gad " & _
								"WHERE jobcode in('6502101','1321410') " & _
								"AND attendance IN('KJ','KK','WH','WS') " & _
								"AND fcba='"& session("fcba") &"' " & _
								"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy')  " & _
								"GROUP BY fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"jobcode " & _
								") " & _
								"gad_max_jjg " & _
								"ON(gad.fcba=gad_max_jjg.fcba " & _
								"AND gad.fddate=gad_max_jjg.fddate " & _
								"AND gad.employeecode=gad_max_jjg.employeecode " & _
								"AND gad.jobcode=gad_max_jjg.jobcode " & _
								"AND gad.unit=gad_max_jjg.max_jjg) " & _
								"WHERE gad.jobcode in('6502101','1321410') " & _
								"AND gad.attendance IN('KJ','KK','WH','WS') " & _
								"AND gad.fcba='"& session("fcba") &"' " & _
								"AND gad.fddate = to_date('" & tglnow & "','mm/dd/yyyy')  " & _
								") " & _
								"gad)gad2 " & _
								"INNER JOIN " & _
								"(SELECT  " & _
								"field.fcba, " & _
								"field.fccode, " & _
								"field.division, " & _
								"TO_CHAR( " & _
								"field.plantingdate, " & _
								"'yyyy') " & _
								"AS plantingyear, " & _
								"field.terraintype " & _
								"FROM field " & _
								"WHERE " & _
								"field.fcba='"& session("fcba") &"' " & _
								"AND field.activation='Y')ff " & _
								"ON(gad2.fcba=ff.fcba " & _
								"AND gad2.locationcode=ff.fccode) " & _
								"WHERE gad2.minimal=1)gad_max_jjg " & _
								"ON(gad.fcba=gad_max_jjg.fcba " & _
								"AND gad.fddate=gad_max_jjg.fddate " & _
								"AND gad.employeecode=gad_max_jjg.employeecode " & _
								"AND gad.jobcode=gad_max_jjg.jobcode) " & _
								"INNER JOIN " & _
								"(SELECT gad2.fcba, " & _
								"gad2.fddate, " & _
								"gad2.employeecode, " & _
								"gad2.jobcode, " & _
								"gad2.attendance, " & _
								"gad2.locationcode, " & _
								"ff.division AS division_max_ha, " & _
								"ff.plantingyear AS plantingyear_max_ha, " & _
								"ff.terraintype AS terraintype_max_ha, " & _
								"gad2.max_ha, " & _
								"gad2.linenokey " & _
								"FROM(SELECT gad.fcba, " & _
								"gad.fddate, " & _
								"gad.employeecode, " & _
								"gad.jobcode, " & _
								"gad.attendance, " & _
								"gad.locationcode, " & _
								"gad.max_ha AS max_ha, " & _
								"gad.linenokey linenokey, " & _
								"ROW_NUMBER() " & _
								"OVER( " & _
								"PARTITION BY gad.fcba " & _
								"||gad.fddate " & _
								"||gad.employeecode " & _
								"||gad.max_ha " & _
								"ORDER BY " & _
								"gad.linenokey ASC) " & _
								"AS minimal " & _
								"FROM(SELECT gad.fcba, " & _
								"gad.fddate, " & _
								"gad.employeecode, " & _
								"gad.jobcode, " & _
								"gad.attendance, " & _
								"gad.locationcode, " & _
								"gad_max_ha.max_ha AS max_ha, " & _
								"gad.linenokey " & _
								"FROM attendance_gad " & _
								"gad " & _
								"INNER JOIN " & _
								"(SELECT fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"jobcode, " & _
								"MAX(output)AS max_ha " & _
								"FROM attendance_gad " & _
								"WHERE jobcode in('6502101','1321410') " & _
								"AND attendance IN('KJ','KK','WH','WS') " & _
								"AND fcba='"& session("fcba") &"' " & _
								"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy')  " & _
								"GROUP BY fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"jobcode) " & _
								"gad_max_ha " & _
								"ON(gad.fcba=gad_max_ha.fcba " & _
								"AND gad.fddate=gad_max_ha.fddate " & _
								"AND gad.employeecode=gad_max_ha.employeecode " & _
								"AND gad.jobcode=gad_max_ha.jobcode " & _
								"AND gad.output=gad_max_ha.max_ha) " & _
								"WHERE gad.jobcode in('6502101','1321410') " & _
								"AND gad.attendance IN('KJ','KK','WH','WS') " & _
								"AND gad.fcba='"& session("fcba") &"' " & _
								"AND gad.fddate = to_date('" & tglnow & "','mm/dd/yyyy')  " & _
								") " & _
								"gad)gad2 " & _
								"INNER JOIN " & _
								"(SELECT field.fcba, " & _
								"field.fccode, " & _
								"field.division, " & _
								"TO_CHAR( " & _
								"field.plantingdate, " & _
								"'yyyy') " & _
								"AS plantingyear, " & _
								"field.terraintype " & _
								"FROM field " & _
								"WHERE field.fcba='"& session("fcba") &"' " & _
								"AND field.activation='Y')ff " & _
								"ON(gad2.fcba=ff.fcba " & _
								"AND gad2.locationcode=ff.fccode) " & _
								"WHERE gad2.minimal=1)gad_max_ha " & _
								"ON(gad.fcba=gad_max_ha.fcba " & _
								"AND gad.fddate=gad_max_ha.fddate " & _
								"AND gad.employeecode=gad_max_ha.employeecode " & _
								"AND gad.jobcode=gad_max_ha.jobcode) " & _
								"INNER JOIN " & _
								"(SELECT DISTINCT " & _
								"multilevel_harvesting_rate.fcba, " & _
								"multilevel_harvesting_rate.employeetype, " & _
								"multilevel_harvesting_rate.attendance, " & _
								"multilevel_harvesting_rate.planting_year, " & _
								"multilevel_harvesting_rate.terrain_type, " & _
								"multilevel_harvesting_rate.basis AS basis_jjg_max_jjg, " & _
								"multilevel_harvesting_rate.basis_ha AS basis_ha_max_jjg " & _
								"FROM multilevel_harvesting_rate " & _
								"WHERE multilevel_harvesting_rate.fcba='"& session("fcba") &"' " & _
								" AND " & year(session("v_Employee")) & "* 100 +" & month(session("v_Employee")) & " BETWEEN To_Number(To_Char(valid_from,'yyyymm')) AND To_Number(To_Char(valid_to,'yyyymm'))  " & _
								") mhr_max_jjg  " & _
								"ON(gad.fcba=mhr_max_jjg.fcba " & _
								"AND emp.presentgrade= " & _
								"mhr_max_jjg.employeetype " & _
								"AND gad.attendance= " & _
								"mhr_max_jjg.attendance " & _
								"AND gad_max_jjg.plantingyear_max_jjg= " & _
								"mhr_max_jjg.planting_year " & _
								"AND gad_max_jjg.terraintype_max_jjg= " & _
								"mhr_max_jjg.terrain_type) " & _
								"INNER JOIN " & _
								"(SELECT DISTINCT " & _
								"multilevel_harvesting_rate.fcba, " & _
								"multilevel_harvesting_rate.employeetype, " & _
								"multilevel_harvesting_rate.attendance, " & _
								"multilevel_harvesting_rate.planting_year, " & _
								"multilevel_harvesting_rate.terrain_type, " & _
								"multilevel_harvesting_rate.basis AS basis_jjg_max_ha, " & _
								"multilevel_harvesting_rate.basis_ha AS basis_ha_max_ha " & _
								"FROM multilevel_harvesting_rate " & _
								"WHERE multilevel_harvesting_rate.fcba='"& session("fcba") &"' " & _
								" AND " & year(session("v_Employee")) & "*100+" & month(session("v_Employee")) & " BETWEEN To_Number(To_Char(valid_from,'yyyymm')) AND To_Number(To_Char(valid_to,'yyyymm'))  " & _
								") mhr_max_ha  " & _
								"ON(gad_max_ha.fcba=mhr_max_ha.fcba " & _
								"AND emp.presentgrade=mhr_max_ha.employeetype " & _
								"AND gad.attendance=mhr_max_ha.attendance " & _
								"AND gad_max_ha.plantingyear_max_ha=mhr_max_ha.planting_year " & _
								"AND gad_max_ha.terraintype_max_ha=mhr_max_ha.terrain_type)) " & _
								") " & _
								"gad_extra_basis " & _
								"LEFT JOIN " & _
								"(SELECT fcba, " & _
								"employeetype, " & _
								"attendance, " & _
								"division, " & _
								"planting_year, " & _
								"terrain_type, " & _
								"SUM(DECODE(levelno,1,min_bunch,0)) " & _
								"min_bunch_level_1, " & _
								"SUM(DECODE(levelno,2,min_bunch,0)) " & _
								"min_bunch_level_2, " & _
								"SUM(DECODE(levelno,3,min_bunch,0)) " & _
								"min_bunch_level_3, " & _
								"SUM(DECODE(levelno,4,min_bunch,0)) " & _
								"min_bunch_level_4, " & _
								"SUM(DECODE(levelno,1,max_bunch,0)) " & _
								"max_bunch_level_1, " & _
								"SUM(DECODE(levelno,2,max_bunch,0)) " & _
								"max_bunch_level_2, " & _
								"SUM(DECODE(levelno,3,max_bunch,0)) " & _
								"max_bunch_level_3, " & _
								"SUM(DECODE(levelno,4,max_bunch,0)) " & _
								"max_bunch_level_4, " & _
								"SUM(DECODE(levelno,1,standart,0)) " & _
								"standart_level_1, " & _
								"SUM(DECODE(levelno,2,standart,0)) " & _
								"standart_level_2, " & _
								"SUM(DECODE(levelno,3,standart,0)) " & _
								"standart_level_3, " & _
								"SUM(DECODE(levelno,4,standart,0)) " & _
								"standart_level_4, " & _
								"SUM(DECODE(levelno,1,rate,0)) " & _
								"rate_level_1, " & _
								"SUM(DECODE(levelno,2,rate,0)) " & _
								"rate_level_2, " & _
								"SUM(DECODE(levelno,3,rate,0)) " & _
								"rate_level_3, " & _
								"SUM(DECODE(levelno,4,rate,0)) " & _
								"rate_level_4, " & _
								"SUM(DECODE(levelno,1,rate_brondolan,0)) " & _
								"rate_brondolan_level_1, " & _
								"SUM(DECODE(levelno,2,rate_brondolan,0)) " & _
								"rate_brondolan_level_2, " & _
								"SUM(DECODE(levelno,3,rate_brondolan,0)) " & _
								"rate_brondolan_level_3, " & _
								"SUM(DECODE(levelno,4,rate_brondolan,0)) " & _
								"rate_brondolan_level_4, " & _
								"SUM(DECODE(levelno,1,rate_nobasis,0)) " & _
								"rate_nobasis_level_1, " & _
								"SUM(DECODE(levelno,2,rate_nobasis,0)) " & _
								"rate_nobasis_level_2, " & _
								"SUM(DECODE(levelno,3,rate_nobasis,0)) " & _
								"rate_nobasis_level_3, " & _
								"SUM(DECODE(levelno,4,rate_nobasis,0)) " & _
								"rate_nobasis_level_4, " & _
								"bonus_basis,bonus_premi " & _
								"FROM(SELECT iplasprod.multilevel_harvesting_rate.fcba, " & _
								"iplasprod.multilevel_harvesting_rate.levelno, " & _
								"iplasprod.multilevel_harvesting_rate.employeetype, " & _
								"iplasprod.multilevel_harvesting_rate.attendance, " & _
								"iplasprod.multilevel_harvesting_rate.division, " & _
								"iplasprod.multilevel_harvesting_rate.planting_year, " & _
								"iplasprod.multilevel_harvesting_rate.terrain_type, " & _
								"iplasprod.multilevel_harvesting_rate.min_bunch, " & _
								"iplasprod.multilevel_harvesting_rate.max_bunch, " & _
								"iplasprod.multilevel_harvesting_rate.standart, " & _
								"iplasprod.multilevel_harvesting_rate.rate, " & _
								"iplasprod.multilevel_harvesting_rate.rate_brondolan, " & _
								"iplasprod.multilevel_harvesting_rate.rate_nobasis, " & _
								"iplasprod.multilevel_harvesting_rate.bonus_basis, " & _
								"iplasprod.multilevel_harvesting_rate.bonus_premi " & _
								"FROM iplasprod.multilevel_harvesting_rate " & _
								"WHERE fcba='"& session("fcba") &"' " & _
								" AND (" & cdbl(year(session("v_Employee"))) & "*100+" & cdbl(month(session("v_Employee"))) & ") BETWEEN To_Number(To_Char(valid_from,'yyyymm')) AND To_Number(To_Char(valid_to,'yyyymm'))  " & _
								") " & _
								"GROUP BY fcba, " & _
								"employeetype, " & _
								"attendance, " & _
								"division, " & _
								"planting_year, " & _
								"terrain_type,bonus_basis,bonus_premi)mhr_level " & _
								"ON(gad_extra_basis.fcba=mhr_level.fcba " & _
								"AND gad_extra_basis.presentgrade=mhr_level.employeetype " & _
								"AND gad_extra_basis.attendance=mhr_level.attendance " & _
								"AND gad_extra_basis.division_final=mhr_level.division " & _
								"AND gad_extra_basis.planting_year_final=mhr_level.planting_year " & _
								"AND gad_extra_basis.terraintype_final=mhr_level.terrain_type) " & _
								"LEFT JOIN " & _
								"(SELECT fcba, " & _
								"fddate, " & _
								"employeecode, " & _
								"attendance, " & _
								"jobcode, " & _
								"division, " & _
								"locationcode, " & _
								"/*plantingyear,*/ " & _
								"/*terraintype,*/ " & _
								"sum(kg_brondolan) over(partition by fcba||fddate||employeecode||attendance||locationcode)AS kg_brondolan " & _
								"FROM(SELECT  " & _
								"attendance_gad.fcba, " & _
								"attendance_gad.fddate, " & _
								"attendance_gad.employeecode, " & _
								"attendance_gad.jobcode, " & _
								"attendance_gad.attendance, " & _
								"attendance_gad.locationcode, " & _
								"ff.division, " & _
								"/*ff.plantingyear,*/ " & _
								"/*ff.terraintype,*/ " & _
								"SUM(attendance_gad.kg_brondolan) AS kg_brondolan " & _
								"FROM attendance_gad " & _
								"INNER JOIN " & _
								"(SELECT field.fcba, " & _
								"field.fccode, " & _
								"field.division, " & _
								"TO_CHAR(field.plantingdate,'yyyy')AS plantingyear, " & _
								"field.terraintype " & _
								"FROM field " & _
								"WHERE field.fcba='"& session("fcba") &"' " & _
								"AND field.activation='Y')ff " & _
								"ON(attendance_gad.fcba=ff.fcba " & _
								"AND attendance_gad.locationcode=ff.fccode) " & _
								"WHERE attendance_gad.jobcode in('6502101','1321410') " & _
								"AND attendance_gad.attendance IN('KJ','KK','WH','WS') " & _
								"AND attendance_gad.fcba='"& session("fcba") &"' " & _
								"AND fddate =to_date('" & tglnow & "','mm/dd/yyyy')  " & _
								"GROUP BY attendance_gad.fcba, " & _
								"attendance_gad.fddate, " & _
								"attendance_gad.employeecode, " & _
								"attendance_gad.jobcode, " & _
								"attendance_gad.attendance, " & _
								"attendance_gad.locationcode, " & _
								"ff.division " & _
								"/*ff.plantingyear,*/ " & _
								"/*ff.terraintype*/) " & _
								"WHERE kg_brondolan>0)brd " & _
								"ON(gad_extra_basis.fcba=brd.fcba " & _
								"AND gad_extra_basis.fddate=brd.fddate " & _
								"AND gad_extra_basis.employeecode=brd.employeecode " & _
								"AND gad_extra_basis.attendance=brd.attendance " & _
								"AND gad_extra_basis.jobcode=brd.jobcode " & _
								"AND gad_extra_basis.division_final=brd.division " & _
								" " & _
								"AND gad_extra_basis.blok=brd.locationcode " & _
								"/*UpdateAMHJan21-PremiBrondolan*/ " & _
								"/*ANDgad_extra_basis.planting_year_final=brd.plantingyear*/ " & _
								"/*ANDgad_extra_basis.terraintype_final=brd.terraintype*/) " & _
								") " & _
								"WHERE       " & _
								" fddate =to_date('" & tglnow & "','mm/dd/yyyy')   " & _
								"ORDER BY fcba,fddate,employeecode " & _
								") " & _
								"ORDER BY employeecode,fddate" 
'response.Write (qNormalisasiDetail)
			
		    connstr.execute(qNormalisasiDetail)
			
			
            If Err.Number = 0 Then
					connstr.CommitTrans
			Else
					connstr.RollbackTrans
					Response.write("Error Delete ["&qNormalisasiDetail&"] ("&Err.Description&")")
					Response.Flush
			End If
			
			response.Write " Start Premi Insentif ..." & "<br>" 
			response.Flush()
			
			response.Write " Calculate HK Panen..." & "<br>" 
			response.Flush()
			
			qjournal = "MERGE INTO attendance_gad_temp a " & _
								" USING  " & _
								" (  " & _
								" SELECT fcba,FDDATE,employeecode,MANDAYS,LINENOKEY FROM employee_premi_panen_tab " & _
								" WHERE fcba = '" & Session("fcBA") & "' " & _
								" AND fddate =to_date('" & tglnow & "','mm/dd/yyyy')   " & _
								" ORDER BY employeecode,fddate " & _
								" )b ON (a.fcba=b.fcba AND a.FDDATE=b.FDDATE AND a.employeecode=b.employeecode AND a.LINENOKEY=b.LINENOKEY)  " & _
								" WHEN MATCHED THEN  " & _
								" UPDATE SET a.MANDAYS = b.MANDAYS " 
			connstr.execute(qjournal)
			
			qPremiInsentifBasis = "SELECT FCBA,EMPLOYEECODE,FDDATE,SUM(PREMI_CAPAI_BASIS) PREMI_CAPAI_BASIS,SUM(PREMI_LEBIH_BASIS) PREMI_LEBIH_BASIS,SUM(premi_bonus) bonus_premi,SUM(BRONDOLAN_TOTAL) PREMI_BRONDOLAN  " & _ 
			                      "FROM employee_premi_panen_tab  " & _
								  "WHERE fcba = '" & Session("fcBA") & "' " & _
								  "AND fddate =to_date('" & tglnow & "','mm/dd/yyyy') " & _
								  "GROUP BY FCBA,EMPLOYEECODE,FDDATE "
								  
			
			Set rsPremiInsentifBasis = Server.CreateObject("ADODB.Recordset")
			rsPremiInsentifBasis.CursorType = 3
			rsPremiInsentifBasis.CursorLocation = 3
			rsPremiInsentifBasis.Open qPremiInsentifBasis, connstr
			if not rsPremiInsentifBasis.eof then
			  Do While Not rsPremiInsentifBasis.eof
			  'response.write "Masuk Premi capai basis"
				 '**A27 -> Premi_capai_basis**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("PREMI_CAPAI_BASIS"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values   " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "A27" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "',' " & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 'response.Write "A27 : " & qAllowance & "<br>"
					 connstr.execute(qAllowance)
				  end if
				'** End A27 -> Premi_capai_basis**'
				
				'**A28 -> premi_lebih_basis**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("PREMI_LEBIH_BASIS"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values   " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "A24" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "',' " & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 'response.Write "A24 : " & qAllowance & "<br>"
					 connstr.execute(qAllowance)
				  end if
				'** End A24 -> premi_lebih_basis**'

				'**A28 -> premi_kerajinan (AMH Apr 21)**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("bonus_premi"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values   " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "A28" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "',' " & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 'response.Write "A24 : " & qAllowance & "<br>"
					 connstr.execute(qAllowance)
				  end if
				'** End A24 -> premi_lebih_basis**'
				
				'**D19 -> potongan_tidak_dapat_basis**'
				'** AMH Mar-22 Non Aktive, dipindahkan ke Proposi HK */
				  ' nRpPremi=cdbl(rsPremiInsentifBasis("POTONGAN_TIDAK_DAPAT_BASIS"))
				  ' if nRpPremi > 0 then
					 ' maxLineNo = maxLineNo + 1
					 ' qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values   " & _
								' "(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "D19" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								' ",0,'" & Session("CodeName") & "','" & Session("CodeName") & "',' " & _
								' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					 ' 'response.Write "D19 : " & qAllowance & "<br>"
					 ' connstr.execute(qAllowance)
				  ' end if
				'**D19 -> potongan_tidak_dapat_basis**'
				
				
				
				'**A26 -> premi_brondolan**'
				  nRpPremi=cdbl(rsPremiInsentifBasis("PREMI_BRONDOLAN"))
				  if nRpPremi > 0 then
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values   " & _
								"(" & maxLineNo & ",'" & rsPremiInsentifBasis("EMPLOYEECODE") & "','" & "A26" & "'," &  " to_date('" & rsPremiInsentifBasis("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "',' " & _
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