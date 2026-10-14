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

	  
	strGrnInv		=	"   SELECT fcba," & _
						"          fddate," & _
						"          employeecode," & _
						"          employee_api.get_employeename(employeecode, fcba)employeename," & _
						"          attendance," & _
						"          jobcode," & _
						"          division_final," & _
						"          planting_year_final," & _
						"          terraintype_final," & _
						"          presentgrade," & _
						"          presentrate_fix," & _
						"          potongan_tidak_dapat_basis," & _
						"          total_janjang," & _
						"          total_ha," & _
						"          jjg_lebih_basis," & _
						"          premi_level_1," & _
						"          premi_level_2," & _
						"          premi_level_3," & _
						"          premi_level_4," & _
						"          (premi_level_1 + premi_level_2 + premi_level_3 + premi_level_4)" & _
						"             AS total_premi_panen" & _
						"     FROM (SELECT gad_extra_basis.fcba," & _
						"                  gad_extra_basis.fddate," & _
						"                  gad_extra_basis.employeecode," & _
						"                  gad_extra_basis.attendance," & _
						"                  gad_extra_basis.jobcode," & _
						"                  gad_extra_basis.division_final," & _
						"                  gad_extra_basis.planting_year_final," & _
						"                  gad_extra_basis.terraintype_final," & _
						"                  gad_extra_basis.presentgrade," & _
						"                  gad_extra_basis.presentrate_fix," & _
						"                  CASE" & _
						"                     WHEN gad_extra_basis.presentrate_original =" & _
						"                             gad_extra_basis.presentrate_fix" & _
						"                     THEN" & _
						" 0" & _
						"                     ELSE" & _
						"                          gad_extra_basis.presentrate_original" & _
						"                        - gad_extra_basis.presentrate_fix" & _
						"                  END" & _
						"                     AS potongan_tidak_dapat_basis," & _
						"                  gad_extra_basis.total_janjang," & _
						"                  gad_extra_basis.total_ha," & _
						"                  gad_extra_basis.jjg_lebih_basis," & _
						"                  CASE" & _
						"                     WHEN (gad_extra_basis.jjg_lebih_basis >" & _
						"                              mhr_level.max_bunch_level_1)" & _
						"                     THEN" & _
						"                        mhr_level.max_bunch_level_1 * mhr_level.rate_level_1" & _
						"                     WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
						"                                  mhr_level.min_bunch_level_1)" & _
						"                          AND (gad_extra_basis.jjg_lebih_basis <=" & _
						"                                  mhr_level.max_bunch_level_1)" & _
						"                     THEN" & _
						"                        gad_extra_basis.jjg_lebih_basis * mhr_level.rate_level_1" & _
						"                     ELSE" & _
						" 0" & _
						"                  END" & _
						"                     AS premi_level_1," & _
						"                  CASE" & _
						"                     WHEN (gad_extra_basis.jjg_lebih_basis >" & _
						"                              mhr_level.max_bunch_level_2)" & _
						"                     THEN" & _
						"                        mhr_level.max_bunch_level_1 * mhr_level.rate_level_2" & _
						"                     WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
						"                                  mhr_level.min_bunch_level_2)" & _
						"                          AND (gad_extra_basis.jjg_lebih_basis <=" & _
						"                                  mhr_level.max_bunch_level_2)" & _
						"                     THEN" & _
						"                          (  gad_extra_basis.jjg_lebih_basis" & _
						"                           - mhr_level.max_bunch_level_1)" & _
						"                        * mhr_level.rate_level_2" & _
						"                     ELSE" & _
						" 0" & _
						"                  END" & _
						"                     AS premi_level_2," & _
						"                  CASE" & _
						"                     WHEN (gad_extra_basis.jjg_lebih_basis >" & _
						"                              mhr_level.max_bunch_level_3)" & _
						"                     THEN" & _
						"                        mhr_level.max_bunch_level_1 * mhr_level.rate_level_3" & _
						"                     WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
						"                                  mhr_level.min_bunch_level_3)" & _
						"                          AND (gad_extra_basis.jjg_lebih_basis <=" & _
						"                                  mhr_level.max_bunch_level_3)" & _
						"                     THEN" & _
						"                          (  gad_extra_basis.jjg_lebih_basis" & _
						"                           - mhr_level.max_bunch_level_2)" & _
						"                        * mhr_level.rate_level_3" & _
						"                     ELSE" & _
						" 0" & _
						"                  END" & _
						"                     AS premi_level_3," & _
						"                  CASE" & _
						"                     WHEN     (gad_extra_basis.jjg_lebih_basis >=" & _
						"                                  mhr_level.min_bunch_level_4)" & _
						"                          AND (gad_extra_basis.jjg_lebih_basis <=" & _
						"                                  mhr_level.max_bunch_level_4)" & _
						"                     THEN" & _
						"                          (  gad_extra_basis.jjg_lebih_basis" & _
						"                           - mhr_level.max_bunch_level_3)" & _
						"                        * mhr_level.rate_level_4" & _
						"                     ELSE" & _
						" 0" & _
						"                  END" & _
						"                     AS premi_level_4" & _
						"             FROM (SELECT DISTINCT fcba," & _
						"                                   fddate," & _
						"                                   employeecode," & _
						"                                   jobcode," & _
						"                                   jml_baris," & _
						"                                   presentgrade," & _
						"                                   presentrate_original," & _
						"                                   presentrate_fix," & _
						"                                   attendance," & _
						"                                   total_janjang," & _
						"                                   total_ha," & _
						"                                   division_final," & _
						"                                   planting_year_final," & _
						"                                   terraintype_final," & _
						"                                   jjg_lebih_basis" & _
						"                     FROM (SELECT gad.fcba," & _
						"                                  gad.fddate," & _
						"                                  gad.employeecode," & _
						"                                  gad.jobcode," & _
						"                                  gad.jml_baris," & _
						"                                  emp.presentgrade," & _
						"                                  emp.presentrate AS presentrate_original," & _
						"                                  CASE" & _
						"                                     WHEN     (gad.total_ha > gad_max_ha.max_ha)" & _
						"                                          AND (gad.total_ha / gad_max_ha.max_ha <" & _
						"                                                  gad.jml_baris)" & _
						"                                          AND (gad.jml_baris > 1)" & _
						"                                     THEN" & _
						"                                        (CASE" & _
						"                                            WHEN gad.total_ha >" & _
						"                                                    mhr_max_ha.basis_ha_max_ha" & _
						"                                            THEN" & _
						"                                               emp.presentrate" & _
						"                                            ELSE" & _
						"                                                 (  gad.total_ha" & _
						"                                                  / mhr_max_ha.basis_ha_max_ha)" & _
						"                                               * emp.presentrate" & _
						"                                         END)" & _
						"                                     WHEN     (gad.total_ha > gad_max_ha.max_ha)" & _
						"                                          AND (gad.total_ha / gad_max_ha.max_ha =" & _
						"                                                  gad.jml_baris)" & _
						"                                          AND (gad.jml_baris > 1)" & _
						"                                     THEN" & _
						"                                        (CASE" & _
						"                                            WHEN gad.total_ha >" & _
						"                                                    mhr_max_jjg.basis_ha_max_jjg" & _
						"                                            THEN" & _
						"                                               emp.presentrate" & _
						"                                            ELSE" & _
						"                                                 (  gad.total_ha" & _
						"                                                  / mhr_max_jjg.basis_ha_max_jjg)" & _
						"                                               * emp.presentrate" & _
						"                                         END)" & _
						"                                     WHEN     (gad.total_ha = gad_max_ha.max_ha)" & _
						"                                          AND (gad.total_ha / gad_max_ha.max_ha =" & _
						"                                                  gad.jml_baris)" & _
						"                                          AND (gad.jml_baris = 1)" & _
						"                                     THEN" & _
						"                                        emp.presentrate" & _
						"                                     ELSE" & _
						"                                          (  gad.total_ha" & _
						"                                           / mhr_max_ha.basis_ha_max_ha)" & _
						"                                        * emp.presentrate" & _
						"                                  END" & _
						"                                     AS presentrate_fix," & _
						"                                  gad.attendance," & _
						"                                  gad.total_janjang," & _
						"                                  gad.total_ha," & _
						"                                  CASE" & _
						"                                     WHEN gad.total_ha > gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        gad_max_ha.division_max_ha" & _
						"                                     WHEN gad.total_ha = gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        (CASE" & _
						"                                            WHEN gad.total_janjang >" & _
						"                                                    gad_max_jjg.max_janjang" & _
						"                                            THEN" & _
						"                                               gad_max_jjg.division_max_jjg" & _
						"                                            ELSE" & _
						"                                               ''" & _
						"                                         END)" & _
						"                                     ELSE" & _
						"                                        ''" & _
						"                                  END" & _
						"                                     AS division_final," & _
						"                                  CASE" & _
						"                                     WHEN gad.total_ha > gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        gad_max_ha.plantingyear_max_ha" & _
						"                                     WHEN gad.total_ha = gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        (CASE" & _
						"                                            WHEN gad.total_janjang >" & _
						"                                                    gad_max_jjg.max_janjang" & _
						"                                            THEN" & _
						"                                               gad_max_jjg.plantingyear_max_jjg" & _
						"                                            ELSE" & _
						"                                               ''" & _
						"                                         END)" & _
						"                                     ELSE" & _
						"                                        ''" & _
						"                                  END" & _
						"                                     AS planting_year_final," & _
						"                                  CASE" & _
						"                                     WHEN gad.total_ha > gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        gad_max_ha.terraintype_max_ha" & _
						"                                     WHEN gad.total_ha = gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        (CASE" & _
						"                                            WHEN gad.total_janjang >" & _
						"                                                    gad_max_jjg.max_janjang" & _
						"                                            THEN" & _
						"                                               gad_max_jjg.terraintype_max_jjg" & _
						"                                            ELSE" & _
						"                                               ''" & _
						"                                         END)" & _
						"                                     ELSE" & _
						"                                        ''" & _
						"                                  END" & _
						"                                     AS terraintype_final," & _
						"                                  CASE" & _
						"                                     WHEN gad.total_ha > gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        (CASE" & _
						"                                            WHEN gad.total_janjang >" & _
						"                                                    mhr_max_ha.basis_jjg_max_ha" & _
						"                                            THEN" & _
						"                                                 gad.total_janjang" & _
						"                                               - mhr_max_ha.basis_jjg_max_ha" & _
						"                                            ELSE" & _
						" 0" & _
						"                                         END)" & _
						"                                     WHEN gad.total_ha = gad_max_ha.max_ha" & _
						"                                     THEN" & _
						"                                        (CASE" & _
						"                                            WHEN gad.total_janjang >" & _
						"                                                    gad_max_jjg.max_janjang" & _
						"                                            THEN" & _
						"                                               (CASE" & _
						"                                                   WHEN gad.total_janjang >" & _
						"                                                           mhr_max_jjg.basis_jjg_max_jjg" & _
						"                                                   THEN" & _
						"                                                        gad.total_janjang" & _
						"                                                      - mhr_max_jjg.basis_jjg_max_jjg" & _
						"                                                   ELSE" & _
						" 0" & _
						"                                                END)" & _
						"                                            ELSE" & _
						" 0" & _
						"                                         END)" & _
						"                                     ELSE" & _
						" 0" & _
						"                                  END" & _
						"                                     AS jjg_lebih_basis" & _
						"                             FROM (  SELECT attendance_gad.fcba," & _
						"                                            attendance_gad.fddate," & _
						"                                            attendance_gad.employeecode," & _
						"                                            attendance_gad.jobcode," & _
						"                                            attendance_gad.attendance," & _
						"                                            ff.division," & _
						"                                            ff.plantingyear," & _
						"                                            ff.terraintype," & _
						"                                               attendance_gad.fcba" & _
						"                                            || attendance_gad.fddate" & _
						"                                            || attendance_gad.employeecode" & _
						"                                            || attendance_gad.jobcode" & _
						"                                            || attendance_gad.attendance" & _
						"                                               AS n," & _
						"                                            SUM (" & _
						"                                               attendance_gad.unit)" & _
						"                                            OVER (" & _
						"                                               PARTITION BY    attendance_gad.fcba" & _
						"                                                            || attendance_gad.fddate" & _
						"                                                            || attendance_gad.employeecode" & _
						"                                                            || attendance_gad.jobcode" & _
						"                                                            || attendance_gad.attendance)" & _
						"                                               AS total_janjang," & _
						"                                            SUM (" & _
						"                                               attendance_gad.output)" & _
						"                                            OVER (" & _
						"                                               PARTITION BY    attendance_gad.fcba" & _
						"                                                            || attendance_gad.fddate" & _
						"                                                            || attendance_gad.employeecode" & _
						"                                                            || attendance_gad.jobcode" & _
						"                                                            || attendance_gad.attendance)" & _
						"                                               AS total_ha," & _
						"                                            COUNT (" & _
						"                                               attendance_gad.fcba)" & _
						"                                            OVER (" & _
						"                                               PARTITION BY    attendance_gad.fcba" & _
						"                                                            || attendance_gad.fddate" & _
						"                                                            || attendance_gad.employeecode" & _
						"                                                            || attendance_gad.jobcode" & _
						"                                                            || attendance_gad.attendance)" & _
						"                                               AS jml_baris" & _
						"                                       FROM attendance_gad" & _
						"                                            INNER JOIN" & _
						"                                            (SELECT field.fcba," & _
						"                                                    field.fccode," & _
						"                                                    field.division," & _
						"                                                    TO_CHAR (field.plantingdate, 'yyyy') AS plantingyear," & _
						"                                                    field.terraintype" & _
						"                                               FROM field" & _
						"                                              WHERE     field.fcba = '"& v_txtfcba &"'" & _
						"                                                    AND field.activation = 'Y') ff" & _
						"                                               ON (    attendance_gad.fcba =" & _
						"                                                          ff.fcba" & _
						"                                                   AND attendance_gad.locationcode =" & _
						"                                                          ff.fccode)" & _
						"                                      WHERE     attendance_gad.jobcode = '6502101'" & _
						"                                            AND attendance_gad.attendance IN" & _
						"                                                   ('KJ', 'KK')" & _
						"                                            AND attendance_gad.fddate >= to_date('"& v_txtdatefrom &"','mm/dd/yyyy')" & _
						"                                            AND attendance_gad.fddate <= to_date('"& v_txtdateuntil &"','mm/dd/yyyy')" & _
						"                                            AND attendance_gad.fcba = '"& v_txtfcba &"'" & _
						"                                   ORDER BY attendance_gad.fcba," & _
						"                                            attendance_gad.fddate," & _
						"                                            attendance_gad.employeecode," & _
						"                                            attendance_gad.jobcode," & _
						"                                            attendance_gad.attendance) gad" & _
						"                                  INNER JOIN" & _
						"                                  (SELECT employee.fcba," & _
						"                                          employee.fccode," & _
						"                                          employee.presentgrade," & _
						"                                          CASE" & _
						"                                             WHEN employee.presentgrade IN" & _
						"                                                     ('KHT', 'KBT')" & _
						"                                             THEN" & _
						"                                                employee.presentrate / 25" & _
						"                                             ELSE" & _
						"                                                employee.presentrate" & _
						"                                          END" & _
						"                                             AS presentrate" & _
						"                                     FROM employee" & _
						"                                    WHERE     employee.fcba = '"& v_txtfcba &"'" & _
						"                                          AND employee.dateterminate IS NULL)" & _
						"                                  emp" & _
						"                                     ON (    gad.fcba = emp.fcba" & _
						"                                         AND gad.employeecode = emp.fccode)" & _
						"                                  INNER JOIN" & _
						"                                  (SELECT gad2.fcba," & _
						"                                          gad2.fddate," & _
						"                                          gad2.employeecode," & _
						"                                          gad2.jobcode," & _
						"                                          gad2.attendance," & _
						"                                          gad2.locationcode," & _
						"                                          ff.division AS division_max_jjg," & _
						"                                          ff.plantingyear AS plantingyear_max_jjg," & _
						"                                          ff.terraintype AS terraintype_max_jjg," & _
						"                                          gad2.max_janjang," & _
						"                                          gad2.linenokey" & _
						"                                     FROM (SELECT gad.fcba," & _
						"                                                  gad.fddate," & _
						"                                                  gad.employeecode," & _
						"                                                  gad.jobcode," & _
						"                                                  gad.attendance," & _
						"                                                  gad.locationcode," & _
						"                                                  gad.max_jjg AS max_janjang," & _
						"                                                  gad.linenokey linenokey," & _
						"                                                  ROW_NUMBER ()" & _
						"                                                  OVER (" & _
						"                                                     PARTITION BY gad.max_jjg" & _
						"                                                     ORDER BY gad.linenokey ASC)" & _
						"                                                     AS minimal" & _
						"                                             FROM (SELECT gad.fcba," & _
						"                                                          gad.fddate," & _
						"                                                          gad.employeecode," & _
						"                                                          gad.jobcode," & _
						"                                                          gad.attendance," & _
						"                                                          gad.locationcode," & _
						"                                                          gad_max_jjg.max_jjg AS max_jjg," & _
						"                                                          gad.linenokey" & _
						"                                                     FROM attendance_gad gad" & _
						"                                                          INNER JOIN" & _
						"                                                          (  SELECT fcba," & _
						"                                                                    fddate," & _
						"                                                                    employeecode," & _
						"                                                                    jobcode," & _
						"                                                                    MAX (unit)" & _
						"                                                                       AS max_jjg" & _
						"                                                               FROM attendance_gad" & _
						"                                                              WHERE     jobcode = '6502101'" & _
						"                                                                    AND attendance IN" & _
						"                                                                           ('KJ'," & _
						"                                                                            'KK')" & _
						"                                                                    AND fddate >= to_date('"& v_txtdatefrom &"', 'mm/dd/yyyy')" & _
						"                                                                    AND fddate <= to_date('"& v_txtdateuntil &"', 'mm/dd/yyyy')" & _
						"                                                                    AND fcba = '"& v_txtfcba &"'" & _
						"                                                           GROUP BY fcba," & _
						"                                                                    fddate," & _
						"                                                                    employeecode," & _
						"                                                                    jobcode)" & _
						"                                                          gad_max_jjg" & _
						"                                                             ON (    gad.fcba =" & _
						"                                                                        gad_max_jjg.fcba" & _
						"                                                                 AND gad.fddate =" & _
						"                                                                        gad_max_jjg.fddate" & _
						"                                                                 AND gad.employeecode =" & _
						"                                                                        gad_max_jjg.employeecode" & _
						"                                                                 AND gad.jobcode =" & _
						"                                                                        gad_max_jjg.jobcode" & _
						"                                                                 AND gad.unit =" & _
						"                                                                        gad_max_jjg.max_jjg))" & _
						"                                                  gad) gad2" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT field.fcba," & _
						"                                                  field.fccode," & _
						"                                                  field.division," & _
						"                                                  TO_CHAR (field.plantingdate," & _
						"                                                           'yyyy')" & _
						"                                                     AS plantingyear," & _
						"                                                  field.terraintype" & _
						"                                             FROM field" & _
						"                                            WHERE     field.fcba = '"& v_txtfcba &"'" & _
						"                                                  AND field.activation = 'Y') ff" & _
						"                                             ON (    gad2.fcba = ff.fcba" & _
						"                                                 AND gad2.locationcode =" & _
						"                                                        ff.fccode)" & _
						"                                    WHERE gad2.minimal = 1) gad_max_jjg" & _
						"                                     ON (    gad.fcba = gad_max_jjg.fcba" & _
						"                                         AND gad.fddate = gad_max_jjg.fddate" & _
						"                                         AND gad.employeecode =" & _
						"                                                gad_max_jjg.employeecode" & _
						"                                         AND gad.jobcode = gad_max_jjg.jobcode)" & _
						"                                  INNER JOIN" & _
						"                                  (SELECT gad2.fcba," & _
						"                                          gad2.fddate," & _
						"                                          gad2.employeecode," & _
						"                                          gad2.jobcode," & _
						"                                          gad2.attendance," & _
						"                                          gad2.locationcode," & _
						"                                          ff.division AS division_max_ha," & _
						"                                          ff.plantingyear AS plantingyear_max_ha," & _
						"                                          ff.terraintype AS terraintype_max_ha," & _
						"                                          gad2.max_ha," & _
						"                                          gad2.linenokey" & _
						"                                     FROM (SELECT gad.fcba," & _
						"                                                  gad.fddate," & _
						"                                                  gad.employeecode," & _
						"                                                  gad.jobcode," & _
						"                                                  gad.attendance," & _
						"                                                  gad.locationcode," & _
						"                                                  gad.max_ha AS max_ha," & _
						"                                                  gad.linenokey linenokey," & _
						"                                                  ROW_NUMBER ()" & _
						"                                                  OVER (" & _
						"                                                     PARTITION BY gad.max_ha" & _
						"                                                     ORDER BY gad.linenokey ASC)" & _
						"                                                     AS minimal" & _
						"                                             FROM (SELECT gad.fcba," & _
						"                                                          gad.fddate," & _
						"                                                          gad.employeecode," & _
						"                                                          gad.jobcode," & _
						"                                                          gad.attendance," & _
						"                                                          gad.locationcode," & _
						"                                                          gad_max_ha.max_ha" & _
						"                                                             AS max_ha," & _
						"                                                          gad.linenokey" & _
						"                                                     FROM attendance_gad gad" & _
						"                                                          INNER JOIN" & _
						"                                                          (  SELECT fcba," & _
						"                                                                    fddate," & _
						"                                                                    employeecode," & _
						"                                                                    jobcode," & _
						"                                                                    MAX (output)" & _
						"                                                                       AS max_ha" & _
						"                                                               FROM attendance_gad" & _
						"                                                              WHERE     jobcode = '6502101'" & _
						"                                                                    AND attendance IN ('KJ', 'KK')" & _
						"                                                                    AND fddate >= to_date('"& v_txtdatefrom &"', 'mm/dd/yyyy')" & _
						"                                                                    AND fddate <= to_date('"& v_txtdateuntil &"', 'mm/dd/yyyy')" & _
						"                                                                    AND fcba = '"& v_txtfcba &"'" & _
						"                                                           GROUP BY fcba," & _
						"                                                                    fddate," & _
						"                                                                    employeecode," & _
						"                                                                    jobcode)" & _
						"                                                          gad_max_ha" & _
						"                                                             ON (    gad.fcba =" & _
						"                                                                        gad_max_ha.fcba" & _
						"                                                                 AND gad.fddate =" & _
						"                                                                        gad_max_ha.fddate" & _
						"                                                                 AND gad.employeecode =" & _
						"                                                                        gad_max_ha.employeecode" & _
						"                                                                 AND gad.jobcode =" & _
						"                                                                        gad_max_ha.jobcode" & _
						"                                                                 AND gad.output =" & _
						"                                                                        gad_max_ha.max_ha))" & _
						"                                                  gad) gad2" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT field.fcba," & _
						"                                                  field.fccode," & _
						"                                                  field.division," & _
						"                                                  TO_CHAR (field.plantingdate," & _
						"                                                           'yyyy')" & _
						"                                                     AS plantingyear," & _
						"                                                  field.terraintype" & _
						"                                             FROM field" & _
						"                                            WHERE     field.fcba = '"& v_txtfcba &"'" & _
						"                                                  AND field.activation = 'Y') ff" & _
						"                                             ON (    gad2.fcba = ff.fcba" & _
						"                                                 AND gad2.locationcode =" & _
						"                                                        ff.fccode)" & _
						"                                    WHERE gad2.minimal = 1) gad_max_ha" & _
						"                                     ON (    gad.fcba = gad_max_ha.fcba" & _
						"                                         AND gad.fddate = gad_max_ha.fddate" & _
						"                                         AND gad.employeecode =" & _
						"                                                gad_max_ha.employeecode" & _
						"                                         AND gad.jobcode = gad_max_ha.jobcode)" & _
						"                                  INNER JOIN" & _
						"                                  (SELECT DISTINCT" & _
						"                                          multilevel_harvesting_rate.fcba," & _
						"                                          multilevel_harvesting_rate.employeetype," & _
						"                                          multilevel_harvesting_rate.attendance," & _
						"                                          multilevel_harvesting_rate.planting_year," & _
						"                                          multilevel_harvesting_rate.terrain_type," & _
						"                                          multilevel_harvesting_rate.basis" & _
						"                                             AS basis_jjg_max_jjg," & _
						"                                          multilevel_harvesting_rate.basis_ha" & _
						"                                             AS basis_ha_max_jjg" & _
						"                                     FROM multilevel_harvesting_rate" & _
						"                                    WHERE multilevel_harvesting_rate.fcba =" & _
						"                                             '"& v_txtfcba &"') mhr_max_jjg" & _
						"                                     ON (    gad.fcba = mhr_max_jjg.fcba" & _
						"                                         AND emp.presentgrade =" & _
						"                                                mhr_max_jjg.employeetype" & _
						"                                         AND gad.attendance =" & _
						"                                                mhr_max_jjg.attendance" & _
						"                                         AND gad_max_jjg.plantingyear_max_jjg =" & _
						"                                                mhr_max_jjg.planting_year" & _
						"                                         AND gad_max_jjg.terraintype_max_jjg =" & _
						"                                                mhr_max_jjg.terrain_type)" & _
						"                                  INNER JOIN" & _
						"                                  (SELECT DISTINCT" & _
						"                                          multilevel_harvesting_rate.fcba," & _
						"                                          multilevel_harvesting_rate.employeetype," & _
						"                                          multilevel_harvesting_rate.attendance," & _
						"                                          multilevel_harvesting_rate.planting_year," & _
						"                                          multilevel_harvesting_rate.terrain_type," & _
						"                                          multilevel_harvesting_rate.basis" & _
						"                                             AS basis_jjg_max_ha," & _
						"                                          multilevel_harvesting_rate.basis_ha" & _
						"                                             AS basis_ha_max_ha" & _
						"                                     FROM multilevel_harvesting_rate" & _
						"                                    WHERE multilevel_harvesting_rate.fcba =" & _
						"                                             '"& v_txtfcba &"') mhr_max_ha" & _
						"                                     ON (    gad_max_ha.fcba = mhr_max_ha.fcba" & _
						"                                         AND emp.presentgrade =" & _
						"                                                mhr_max_ha.employeetype" & _
						"                                         AND gad.attendance =" & _
						"                                                mhr_max_ha.attendance" & _
						"                                         AND gad_max_ha.plantingyear_max_ha =" & _
						"                                                mhr_max_ha.planting_year" & _
						"                                         AND gad_max_ha.terraintype_max_ha =" & _
						"                                                mhr_max_ha.terrain_type)))" & _
						"                  gad_extra_basis" & _
						"                  INNER JOIN" & _
						"                  (  SELECT fcba," & _
						"                            employeetype," & _
						"                            attendance," & _
						"                            division," & _
						"                            planting_year," & _
						"                            terrain_type," & _
						"                            SUM (DECODE (levelno, 1, min_bunch, 0))" & _
						"                               min_bunch_level_1," & _
						"                            SUM (DECODE (levelno, 2, min_bunch, 0))" & _
						"                               min_bunch_level_2," & _
						"                            SUM (DECODE (levelno, 3, min_bunch, 0))" & _
						"                               min_bunch_level_3," & _
						"                            SUM (DECODE (levelno, 4, min_bunch, 0))" & _
						"                               min_bunch_level_4," & _
						"                            SUM (DECODE (levelno, 1, max_bunch, 0))" & _
						"                               max_bunch_level_1," & _
						"                            SUM (DECODE (levelno, 2, max_bunch, 0))" & _
						"                               max_bunch_level_2," & _
						"                            SUM (DECODE (levelno, 3, max_bunch, 0))" & _
						"                               max_bunch_level_3," & _
						"                            SUM (DECODE (levelno, 4, max_bunch, 0))" & _
						"                               max_bunch_level_4," & _
						"                            SUM (DECODE (levelno, 1, rate, 0)) rate_level_1," & _
						"                            SUM (DECODE (levelno, 2, rate, 0)) rate_level_2," & _
						"                            SUM (DECODE (levelno, 3, rate, 0)) rate_level_3," & _
						"                            SUM (DECODE (levelno, 4, rate, 0)) rate_level_4" & _
						"                       FROM (SELECT iplasprod.multilevel_harvesting_rate.fcba," & _
						"                                    iplasprod.multilevel_harvesting_rate.levelno," & _
						"                                    iplasprod.multilevel_harvesting_rate.employeetype," & _
						"                                    iplasprod.multilevel_harvesting_rate.attendance," & _
						"                                    iplasprod.multilevel_harvesting_rate.division," & _
						"                                    iplasprod.multilevel_harvesting_rate.planting_year," & _
						"                                    iplasprod.multilevel_harvesting_rate.terrain_type," & _
						"                                    iplasprod.multilevel_harvesting_rate.min_bunch," & _
						"                                    iplasprod.multilevel_harvesting_rate.max_bunch," & _
						"                                    iplasprod.multilevel_harvesting_rate.rate" & _
						"                               FROM iplasprod.multilevel_harvesting_rate)" & _
						"                      WHERE fcba = '"& v_txtfcba &"'"
   
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
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

	strGrnInv	=	strGrnInv &	"                   GROUP BY fcba," & _
								"                            employeetype," & _
								"                            attendance," & _
								"                            division," & _
								"                            planting_year," & _
								"                            terrain_type) mhr_level" & _
								"                     ON (    gad_extra_basis.fcba = mhr_level.fcba" & _
								"                         AND gad_extra_basis.presentgrade =" & _
								"                                mhr_level.employeetype" & _
								"                         AND gad_extra_basis.attendance = mhr_level.attendance" & _
								"                         AND gad_extra_basis.division_final = mhr_level.division" & _
								"                         AND gad_extra_basis.planting_year_final =" & _
								"                                mhr_level.planting_year" & _
								"                         AND gad_extra_basis.terraintype_final =" & _
								"                                mhr_level.terrain_type))" 
								
								
	strGrnInv	=	strGrnInv & " ORDER BY fcba, fddate, employeecode" 

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
	    	<th colspan="21" style="background-color:#FFFFFF; border:none;">  
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
			<th style="background-color:#00A300;color:#fff;">Attendance</th>
			<th style="background-color:#00A300;color:#fff;">Jobcode</th>
			<th style="background-color:#00A300;color:#fff;">Division Final</th>
			<th style="background-color:#00A300;color:#fff;">Planting Year Final</th>
			<th style="background-color:#00A300;color:#fff;">Terrain Type Final</th>
			<th style="background-color:#00A300;color:#fff;">Present Grade</th>
			<th style="background-color:#00A300;color:#fff;">Present Rate Fix</th>
			<th style="background-color:#00A300;color:#fff;">Potongan Tidak Dapat Basis</th>
			<th style="background-color:#00A300;color:#fff;">Total Janjang</th>
			<th style="background-color:#00A300;color:#fff;">Total HA</th>
			<th style="background-color:#00A300;color:#fff;">Janjang Lebih Basis</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 1</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 2</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 3</th>
			<th style="background-color:#00A300;color:#fff;">Premi Level 4</th>
			<th style="background-color:#00A300;color:#fff;">Total Premi Panen</th>
		</tr>
	</thead>
	<tbody>
	<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
			<tr>
				<td colspan="18" align="left">NO DATA AVAILABLE</td>
			</tr>
	<%end if
			NO = 1
			v_estate_bunches = 0
			v_mill_bunches = 0
			v_bunch_mill_weight_bruto = 0
			v_kg_potongan = 0
			v_bunch_mill_weight_netto = 0
			v_bjr_per_chitno = 0
			v_total_bucket = 0
			v_kg_brondolan = 0
			v_kg_per_bucket = 0
			while not rsGrnInv.eof
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
	%>
			<tr class="<%=cClass%>">
				<td align="center"><%=rsGrnInv("fddate")%></td>
				<td align="center"><%=rsGrnInv("employeecode")%> - <%=rsGrnInv("employeename")%></td>
				<td align="left"><%=rsGrnInv("attendance")%></td>
				<td align="left"><%=rsGrnInv("jobcode")%></td>
				<td align="center"><%=rsGrnInv("division_final")%></td>
				<td align="center"><%=rsGrnInv("planting_year_final")%></td>
				<td align="center"><%=rsGrnInv("terraintype_final")%></td>
				<td align="center"><%=rsGrnInv("presentgrade")%></td> 
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("presentrate_fix"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("potongan_tidak_dapat_basis"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_janjang"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_ha"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("jjg_lebih_basis"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_1"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_2"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_3"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_level_4"))),0)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_premi_panen"))),0)%></td>
			</tr>
	<%			
				v_estate_bunches = v_estate_bunches + cdbl(to_nol(rsGrnInv("potongan_tidak_dapat_basis")))
				v_mill_bunches = v_mill_bunches + cdbl(to_nol(rsGrnInv("total_janjang")))	
				v_bunch_mill_weight_bruto = v_bunch_mill_weight_bruto + cdbl(to_nol(rsGrnInv("total_ha")))
				v_kg_potongan = v_kg_potongan + cdbl(to_nol(rsGrnInv("jjg_lebih_basis")))
				v_bunch_mill_weight_netto = v_bunch_mill_weight_netto + cdbl(to_nol(rsGrnInv("premi_level_1")))	
				v_bjr_per_chitno =v_bjr_per_chitno + cdbl(to_nol(rsGrnInv("premi_level_2")))	
				v_total_bucket = v_total_bucket + cdbl(to_nol(rsGrnInv("premi_level_3")))	
				v_kg_brondolan = v_kg_brondolan + cdbl(to_nol(rsGrnInv("premi_level_4")))	
				v_kg_per_bucket = v_kg_per_bucket + cdbl(to_nol(rsGrnInv("total_premi_panen")))	

				rsGrnInv.MoveNext
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%>
			<tr bgcolor="#F4F4A0"> 
				<td align="center" colspan="9"><font style="font-weight: bold;">TOTAL</font></td>				
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_estate_bunches)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_mill_bunches)),0)%></font></td>  
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bunch_mill_weight_bruto)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_potongan)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bunch_mill_weight_netto)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_bjr_per_chitno)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_bucket)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_brondolan)),0)%></font></td> 
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_per_bucket)),0)%></font></td> 
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