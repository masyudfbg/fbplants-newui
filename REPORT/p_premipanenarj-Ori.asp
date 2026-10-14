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

	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")
	v_txtFcba = Request("txtFcba")
	
	txttdate = txtMonth & "" & txtYear
	v_txtPOCode		=	trim(request("txtPOCode"))
	compare1 = trim(request("compare1"))

	varDefaultBa = Session("fcBA")

	v_txtsectionname = request("txtsectionname")
	v_txtvehiclecode = request("txtvehiclecode")
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	v_txtown = Request("txtown")
	v_txtempcode = Request("txtempcode")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	

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
	  
	strGrnInv		=	" SELECT fcba," & _
						"        fddate," & _
						"        employeecode," & _
						"        employee_api.get_employeename(fcba, employeecode)employeename," & _
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
						"        kg_brondolan_fix AS kg_brondolan," & _
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
						"                                  AND (  gad_extra_basis.total_ha" & _
						"                                       / gad_extra_basis.max_ha <" & _
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
						"                                  AND (  gad_extra_basis.total_ha" & _
						"                                       / gad_extra_basis.max_ha =" & _
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
						"                                  AND (  gad_extra_basis.total_ha" & _
						"                                       / gad_extra_basis.max_ha =" & _
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
						"                                                      WHERE     field.fcba = '" & v_txtFcba & "'" & _
						"                                                            AND field.activation =" & _
						"                                                                   'Y') ff" & _
						"                                                       ON (    attendance_gad.fcba =" & _
						"                                                                  ff.fcba" & _
						"                                                           AND attendance_gad.locationcode =" & _
						"                                                                  ff.fccode)" & _
						"                                              WHERE     attendance_gad.jobcode =" & _
						"                                                           '6502101'" & _
						"                                                    AND attendance_gad.attendance IN" & _
						"                                                           ('KJ', 'KK','WH','WS')" & _
						"                                                    AND attendance_gad.fcba =" & _
						"                                                           '" & v_txtFcba & "'" & _
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
						"                                            WHERE     employee.fcba = '" & v_txtFcba & "'" & _
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
						"                                                                                   ('KJ', 'KK','WH','WS')" & _
						"                                                                            AND fcba =" & _
						"                                                                                   '" & v_txtFcba & "'" & _
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
						"                                                    WHERE     field.fcba = '" & v_txtFcba & "'" & _
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
						"                                                                                   ('KJ', 'KK','WH','WS')" & _
						"                                                                            AND fcba =" & _
						"                                                                                   '" & v_txtFcba & "'" & _
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
						"                                                    WHERE     field.fcba = '" & v_txtFcba & "'" & _
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
						"                                                     '" & v_txtFcba & "') mhr_max_jjg" & _
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
						"                                                     '" & v_txtFcba & "') mhr_max_ha" & _
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
						"                              WHERE fcba = '" & v_txtFcba & "'" & _
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
						"                                  plantingyear," & _
						"                                  terraintype," & _
						"                                  sum(kg_brondolan) over (partition by fcba||fddate||employeecode||attendance) as kg_brondolan" & _
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
						"                                              WHERE     field.fcba = '" & v_txtFcba & "'" & _
						"                                                    AND field.activation = 'Y') ff" & _
						"                                               ON (    attendance_gad.fcba =" & _
						"                                                          ff.fcba" & _
						"                                                   AND attendance_gad.locationcode =" & _
						"                                                          ff.fccode)" & _
						"                                      WHERE     attendance_gad.jobcode = '6502101'" & _
						"                                            AND attendance_gad.attendance IN" & _
						"                                                   ('KJ', 'KK','WH','WS')" & _
						"                                            AND attendance_gad.fcba = '" & v_txtFcba & "'" & _
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
						"            WHERE     fddate >= TO_DATE ('"& v_txtdatefrom &"', 'mm/dd/yyyy')" & _
						"                  AND fddate <= TO_DATE ('"& v_txtdateuntil &"', 'mm/dd/yyyy')" 

   
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
	end if 
								
	strGrnInv	=	strGrnInv & "  ORDER BY fcba, fddate, employeecode)" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Premi Panen</title>
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
							<font face="verdana" size="1" style="font-weight: bold;">DARI</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">SAMPAI</font>
						</td> 
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">EMPLOYEE CODE</font>
						</td>
					</tr>
					<tr>						
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   	" select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL and fctype= 'E' "
							' response.write strsql			
							rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsFcba.eof then
								DO while not rsFcba.eof
									if v_txtFcba = rsFcba("fcba") then
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
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td> 
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
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
	    	<th colspan="23" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>PREMI PANEN</h2>
		                <br/>
					</div>
				</div>
			</th>
		</tr>
		<tr> 
			<th style="background-color:#00A300;color:#fff;">Date</th>
			<th style="background-color:#00A300;color:#fff;">Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">Attendance</th>
			<th style="background-color:#00A300;color:#fff;">Jobcode</th>
			<th style="background-color:#00A300;color:#fff;">Division Final</th>
			<th style="background-color:#00A300;color:#fff;">Planting Year Final</th>
			<th style="background-color:#00A300;color:#fff;">Terrain Type Final</th>
			<th style="background-color:#00A300;color:#fff;">Present Grade</th>
			<th style="background-color:#00A300;color:#fff;">Present Rate</th>
			<th style="background-color:#00A300;color:#fff;">Potongan Tidak Dapat Basis</th>
			<th style="background-color:#00A300;color:#fff;">Total Janjang</th>
			<th style="background-color:#00A300;color:#fff;">Total HA</th>
			<th style="background-color:#00A300;color:#fff;">Premi Capai Basis</th>
			<th style="background-color:#00A300;color:#fff;">Janjang Lebih Basis</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 1</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 2</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 3</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 4</th>
			<th style="background-color:#00A300;color:#fff;">Premi Lebih Basis</th>
			<th style="background-color:#00A300;color:#fff;">KG Brondolan</th>
			<th style="background-color:#00A300;color:#fff;">Premi Brondolan</th>
			<th style="background-color:#00A300;color:#fff;">Total Pendapatan Panen</th>
		</tr>
	</thead>
	<tbody>
	<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
			<tr>
				<td colspan="23" align="left">NO DATA AVAILABLE</td>
			</tr>
	<%end if
			NO = 1
			v_kgbrondolan = 0
			v_premi_brondolan = 0
			v_total_pendapatan_panen = 0
			v_presentrate = 0
			v_potongan_tidak_dapat_basis = 0
			v_total_janjang = 0	
			v_total_ha = 0
			v_jjg_lebih_basis = 0
			v_premi_level_1 = 0
			v_premi_level_2 = 0	
			v_premi_level_3 = 0	
			v_premi_level_4 = 0
			v_premi_lebih_basis = 0
			v_premi_capai_basis = 0
			while not rsGrnInv.eof
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
	%>
			<tr class="<%=cClass%>">
				<td align="center"><%=rsGrnInv("fddate")%></td>
				<td align="center"><%=rsGrnInv("employeecode")%></td>
				<td align="center"><%=rsGrnInv("employeename")%></td>
				<td align="center"><%=rsGrnInv("attendance")%></td>
				<td align="center"><%=rsGrnInv("jobcode")%></td>
				<td align="center"><%=rsGrnInv("division_final")%></td>
				<td align="center"><%=rsGrnInv("planting_year_final")%></td>
				<td align="center"><%=rsGrnInv("terraintype_final")%></td>
				<td align="center"><%=rsGrnInv("presentgrade")%></td> 
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("presentrate"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("potongan_tidak_dapat_basis"))),2)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_janjang"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_ha"))),2)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_capai_basis"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("jjg_lebih_basis"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_1"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_2"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_3"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_4"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_lebih_basis"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_brondolan"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_brondolan"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_pendapatan_panen"))),2)%></td>
			</tr>
	<%			v_kgbrondolan = v_kgbrondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))
				v_premi_brondolan = v_premi_brondolan + cdbl(to_nol(rsGrnInv("premi_brondolan")))
				v_total_pendapatan_panen = v_total_pendapatan_panen + cdbl(to_nol(rsGrnInv("total_pendapatan_panen")))
				v_presentrate = v_presentrate + cdbl(to_nol(rsGrnInv("presentrate")))
				v_potongan_tidak_dapat_basis = v_potongan_tidak_dapat_basis + cdbl(to_nol(rsGrnInv("potongan_tidak_dapat_basis")))
				v_total_janjang = v_total_janjang + cdbl(to_nol(rsGrnInv("total_janjang")))	
				v_total_ha = v_total_ha + cdbl(to_nol(rsGrnInv("total_ha")))
				v_jjg_lebih_basis = v_jjg_lebih_basis + cdbl(to_nol(rsGrnInv("jjg_lebih_basis")))
				v_premi_level_1 = v_premi_level_1 + cdbl(to_nol(rsGrnInv("premi_level_1")))	
				v_premi_level_2 = v_premi_level_2 + cdbl(to_nol(rsGrnInv("premi_level_2")))	
				v_premi_level_3 = v_premi_level_3 + cdbl(to_nol(rsGrnInv("premi_level_3")))	
				v_premi_level_4 = v_premi_level_4 + cdbl(to_nol(rsGrnInv("premi_level_4")))	
				v_premi_lebih_basis = v_premi_lebih_basis + cdbl(to_nol(rsGrnInv("premi_lebih_basis")))
				v_premi_capai_basis = v_premi_capai_basis + cdbl(to_nol(rsGrnInv("premi_capai_basis")))	

				rsGrnInv.MoveNext
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%>
			<tr bgcolor="#F4F4A0"> 
				<td align="center" colspan="9"><font style="font-weight: bold;">TOTAL</font></td>				
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_presentrate)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_potongan_tidak_dapat_basis)),2)%></font></td>
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_janjang)),0)%></font></td>  
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_ha)),2)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_capai_basis)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_jjg_lebih_basis)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_level_1)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_level_2)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_level_3)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_level_4)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_lebih_basis)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_brondolan)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_brondolan)),0)%></font></td>
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_pendapatan_panen)),2)%></font></td> 
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