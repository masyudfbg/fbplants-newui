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

	v_txtFcba = Request("txtFcba")
	v_txtMonth = Request("txtMonth")
	v_txtYear = Request("txtYear")
	v_txtsectionname = Request("txtsectionname")
	v_txtempcode = Request("txtempcode")

	txttdate = txtMonth & "" & txtYear
	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")

	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
	nSpanTotal = 0
    ' selectCons = selectCons & ",fcba"
    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
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

	q = "SELECT type_2 FROM allowded WHERE fccode = 'A62' AND fcba = '"& v_txtfcba &"'"

	' response.Write q
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.Open q, connstr
	 
	If not rsCallCombo.eof then
	   	type_report = rsCallCombo("type_2")
	End if
	
	If 	type_report = "MIN" then 
		strGrnInv	=	"   SELECT premi.fcba," & _
						"          premi.transdate," & _
						"          premi.operatorcode," & _
						"          emp.fcname AS operatorname," & _
						"          premi.attendance," & _
						"          premi.kg_tbs," & _
						"          premi.kg_brondolan," & _
						"          premi.kg_pupuk," & _
						"          premi.kg_lain_lain," & _
						"          premi.total_kg," & _
						"          premi.basis," & _
						"          premi.kg_premi_lebih_basis," & _
						"          premi.total_premi," & _
						"          CASE" & _
						"             WHEN premi.kg_premi_lebih_basis > 0 THEN 1" & _
						"             WHEN premi.kg_lain_lain > 0 THEN 1" & _
						"             ELSE 0" & _
						"          END AS HK_Efektif" & _
						"     FROM (SELECT fcba," & _
						"                  receptiondate AS transdate," & _
						"                  Operator AS operatorcode," & _
						"                  attendance," & _
						"                  kg_tbs," & _
						"                  kg_brondolan," & _
						"                  Kg_Pupuk," & _
						"                  Kg_Lain_Lain," & _
						"                  total_kg," & _
						"                  basis," & _
						"                  kg_premi_lebih_basis," & _
						"                  kg_premi_brondolan," & _
						"                  premi_level_1," & _
						"                  premi_level_2," & _
						"                  premi_level_3," & _
						"                  premi_level_4," & _
						"                  extra_rate_level_1," & _
						"                  extra_rate_level_2," & _
						"                  extra_rate_level_3," & _
						"                  extra_rate_level_4," & _
						"                  brondolan_rate," & _
						"                  lain_lain_rate," & _
						"                  (  (  (premi_level_1 * extra_rate_level_1)" & _
						"                      + (premi_level_2 * extra_rate_level_2)" & _
						"                      + (premi_level_3 * extra_rate_level_3)" & _
						"                      + (premi_level_4 * extra_rate_level_4))" & _
						"                   + (kg_premi_brondolan * brondolan_rate)" & _
						"                   + (kg_lain_lain * lain_lain_rate))" & _
						"                     AS total_premi" & _
						"             FROM (SELECT fcba," & _
						"                          receptiondate," & _
						"                          Operator," & _
						"                          attendance," & _
						"                          kg_tbs," & _
						"                          kg_brondolan," & _
						"                          Kg_Pupuk," & _
						"                          Kg_Lain_Lain," & _
						"                          total_kg," & _
						"                          basis," & _
						"                          kg_premi_lebih_basis," & _
						"                          CASE" & _
						"                             WHEN kg_premi_lebih_basis > 0" & _
						"                             THEN kg_brondolan" & _
						"                             WHEN kg_premi_lebih_basis = 0" & _
						"                                  AND (total_kg + kg_brondolan) < basis" & _
						"                             THEN 0" & _
						"                             ELSE (total_kg + kg_brondolan) - basis" & _
						"                          END AS kg_premi_brondolan," & _
						"                          CASE" & _
						"                             WHEN kg_premi_lebih_basis > basis_extra_level_1" & _
						"                             THEN basis_extra_level_1" & _
						"                             ELSE kg_premi_lebih_basis" & _
						"                          END AS premi_level_1," & _
						"                          CASE" & _
						"                             WHEN (kg_premi_lebih_basis - basis_extra_level_1) <= 0" & _
						"                             THEN 0" & _
						"                             WHEN (kg_premi_lebih_basis - basis_extra_level_1) > basis_extra_level_2" & _
						"                             THEN basis_extra_level_2" & _
						"                             ELSE (kg_premi_lebih_basis - basis_extra_level_1)" & _
						"                          END AS premi_level_2," & _
						"                          CASE WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2) <= 0" & _
						"                             THEN 0" & _
						"                             WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2) > basis_extra_level_3" & _
						"                             THEN basis_extra_level_3" & _
						"                             ELSE (  kg_premi_lebih_basis" & _
						"                                 - basis_extra_level_1" & _
						"                                 - basis_extra_level_2)" & _
						"                          END AS premi_level_3," & _
						"                          CASE WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2" & _
						"                                   - basis_extra_level_3) <= 0" & _
						"                             THEN 0" & _
						"                             WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2" & _
						"                                   - basis_extra_level_3) > basis_extra_level_4" & _
						"                             THEN" & _
						"                                (  kg_premi_lebih_basis" & _
						"                                 - basis_extra_level_1" & _
						"                                 - basis_extra_level_2" & _
						"                                 - basis_extra_level_3)" & _
						"                             ELSE" & _
						"                                (  kg_premi_lebih_basis" & _
						"                                 - basis_extra_level_1" & _
						"                                 - basis_extra_level_2" & _
						"                                 - basis_extra_level_3)" & _
						"                          END AS premi_level_4," & _
						"                          extra_rate_level_1," & _
						"                          extra_rate_level_2," & _
						"                          extra_rate_level_3," & _
						"                          extra_rate_level_4," & _
						"                          brondolan_rate," & _
						"                          lain_lain_rate" & _
						"                     FROM (  SELECT fcba," & _
						"                                    receptiondate," & _
						"                                    Operator," & _
						"                                    attendance," & _
						"                                    SUM (kg_tbs) AS kg_tbs," & _
						"                                    SUM (kg_brondolan) AS kg_brondolan," & _
						"                                    SUM (Kg_Pupuk) AS Kg_Pupuk," & _
						"                                    SUM (Kg_Lain_Lain) AS Kg_Lain_Lain," & _
						"                                    (SUM (kg_tbs) + (0.5 * SUM (Kg_Pupuk)))" & _
						"                                       AS total_kg," & _
						"                                    MAX (basis) AS basis," & _
						"                                    CASE" & _
						"                                       WHEN (SUM (kg_tbs) + (0.5 * SUM (Kg_Pupuk))) > MAX (basis)" & _
						"                                       THEN (SUM (kg_tbs) + (0.5 * SUM (Kg_Pupuk))) - MAX (basis)" & _
						"                                       ELSE 0" & _
						"                                    END AS kg_premi_lebih_basis," & _
						"                                    MAX (basis_extra_level_1) AS basis_extra_level_1," & _
						"                                    MAX (basis_extra_level_2) AS basis_extra_level_2," & _
						"                                    MAX (basis_extra_level_3) AS basis_extra_level_3," & _
						"                                    MAX (basis_extra_level_4) AS basis_extra_level_4," & _
						"                                    MAX (extra_rate_level_1) AS extra_rate_level_1," & _
						"                                    MAX (extra_rate_level_2) AS extra_rate_level_2," & _
						"                                    MAX (extra_rate_level_3) AS extra_rate_level_3," & _
						"                                    MAX (extra_rate_level_4) AS extra_rate_level_4," & _
						"                                    MAX (brondolan_rate) AS brondolan_rate," & _
						"                                    MAX (lain_lain_rate) AS lain_lain_rate" & _
						"                               FROM (SELECT fcba," & _
						"                                            receptiondate," & _
						"                                            Operator," & _
						"                                            attendance," & _
						"                                            kg_tbs," & _
						"                                            kg_brondolan," & _
						"                                            0 AS Kg_Pupuk," & _
						"                                            0 AS Kg_Lain_Lain," & _
						"                                            basis," & _
						"                                            basis_extra_level_1," & _
						"                                            basis_extra_level_2," & _
						"                                            basis_extra_level_3," & _
						"                                            basis_extra_level_4," & _
						"                                            extra_rate_level_1," & _
						"                                            extra_rate_level_2," & _
						"                                            extra_rate_level_3," & _
						"                                            extra_rate_level_4," & _
						"                                            brondolan_rate," & _
						"                                            lain_lain_rate" & _
						"                                       FROM (  SELECT hasiltbsbrd.fcba," & _
						"                                                      hasiltbsbrd.receptiondate," & _
						"                                                      hasiltbsbrd.Operator," & _
						"                                                      hasiltbsbrd.attendance," & _
						"                                                      SUM (hasiltbsbrd.Kg_TBS) AS kg_tbs," & _
						"                                                      SUM (hasiltbsbrd.Kg_Brondolan) AS kg_brondolan," & _
						"                                                      hasiltbsbrd.basis," & _
						"                                                      hasiltbsbrd.basis_extra_level_1," & _
						"                                                      hasiltbsbrd.basis_extra_level_2," & _
						"                                                      hasiltbsbrd.basis_extra_level_3," & _
						"                                                      hasiltbsbrd.basis_extra_level_4," & _
						"                                                      hasiltbsbrd.extra_rate_level_1," & _
						"                                                      hasiltbsbrd.extra_rate_level_2," & _
						"                                                      hasiltbsbrd.extra_rate_level_3," & _
						"                                                      hasiltbsbrd.extra_rate_level_4," & _
						"                                                      hasiltbsbrd.brondolan_rate," & _
						"                                                      hasiltbsbrd.lain_lain_rate" & _
						"                                                 FROM (SELECT kg_tbs_brd.fcba," & _
						"                                                              kg_tbs_brd.receptiondate," & _
						"                                                              kg_tbs_brd.Operator," & _
						"                                                              kg_tbs_brd.attendance," & _
						"                                                              kg_tbs_brd.division," & _
						"                                                                 kg_tbs_brd.fcba" & _
						"                                                              || kg_tbs_brd.receptiondate" & _
						"                                                              || kg_tbs_brd.operator," & _
						"                                                              kg_tbs_brd.Kg_TBS," & _
						"                                                              kg_tbs_brd.Kg_Brondolan," & _
						"											 case " & _
                        "                                            when kg_tbs_brd.division = 'AFD-13' then " & _
                        "                                            min(mor.basis) over (partition by kg_tbs_brd.fcba||kg_tbs_brd.receptiondate||kg_tbs_brd.operator)" & _
                        "                                            else avg(mor.basis) over (partition by kg_tbs_brd.fcba||kg_tbs_brd.receptiondate||kg_tbs_brd.operator)" & _ 
                        "                                            end as basis, " & _
						"                                                              MIN ( mor_level.basis_extra_level_1)" & _
						"                                                              OVER ( PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_1," & _
						"                                                              MIN ( mor_level.basis_extra_level_2)" & _
						"                                                              OVER ( PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_2," & _
						"                                                              MIN (" & _
						"                                                                 mor_level.basis_extra_level_3)" & _
						"                                                              OVER (" & _
						"                                                                 PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_3," & _
						"                                                              MIN (" & _
						"                                                                 mor_level.basis_extra_level_4)" & _
						"                                                              OVER (" & _
						"                                                                 PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_4," & _
						"                                                              mor_level.extra_rate_level_1," & _
						"                                                              mor_level.extra_rate_level_2," & _
						"                                                              mor_level.extra_rate_level_3," & _
						"                                                              mor_level.extra_rate_level_4," & _
						"                                                              mor.brondolan_rate," & _
						"                                                              mor.lain_lain_rate " & _
						"                                                         FROM (  SELECT fcba," & _
						"                                                                        receptiondate," & _
						"                                                                        driver" & _
						"                                                                           AS Operator," & _
						"                                                                        attendance," & _
						"                                                                        division," & _
						"                                                                        SUM (kg_bunch)" & _
						"                                                                           AS Kg_TBS," & _
						"                                                                        SUM (" & _
						"                                                                           kg_bucket)" & _
						"                                                                           AS Kg_Brondolan" & _
						"                                                                   FROM (SELECT bunchbucket.fcba," & _
						"                                                                                bunchbucket.spbno," & _
						"                                                                                bunchbucket.receptiondate," & _
						"                                                                                bunchbucket.vehicle," & _
						"                                                                                bunchbucket.driver," & _
						"                                                                                bunchbucket.attendance," & _
						"                                                                                bunchbucket.division," & _
						"                                                                                bunchbucket.total_bunch," & _
						"                                                                                iplasprod.wb_harvesting.bjr_per_chitno," & _
						"                                                                                (  bunchbucket.total_bunch" & _
						"                                                                                 * iplasprod.wb_harvesting.bjr_per_chitno)" & _
						"                                                                                   AS Kg_Bunch," & _
						"                                                                                bunchbucket.total_bucket," & _
						"                                                                                iplasprod.wb_harvesting.kg_per_bucket," & _
						"                                                                                (  bunchbucket.total_bucket" & _
						"                                                                                 * iplasprod.wb_harvesting.kg_per_bucket)" & _
						"                                                                                   AS Kg_Bucket" & _
						"                                                                           FROM (  SELECT fcba," & _
						"                                                                                          spbno," & _
						"                                                                                          receptiondate," & _
						"                                                                                          vehicle," & _
						"                                                                                          driver," & _
						"                                                                                          attendance," & _
						"                                                                                          division," & _
						"                                                                                          SUM (" & _
						"                                                                                             total_bunch)" & _
						"                                                                                             AS total_bunch," & _
						"                                                                                          SUM (" & _
						"                                                                                             total_bucket)" & _
						"                                                                                             AS total_bucket" & _
						"                                                                                     FROM (  SELECT iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division," & _
						"                                                                                                    NVL (" & _
						"                                                                                                       SUM (" & _
						"                                                                                                          iplasprod.harvestingspb.bunch)," & _
						"                                                                                                       0)" & _
						"                                                                                                       AS total_bunch, 0" & _
						"                                                                                                       AS total_bucket" & _
						"                                                                                               FROM iplasprod.harvestingspb" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.attendance_gad" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.attendance_gad.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.receptiondate =" & _
						"                                                                                                                  iplasprod.attendance_gad.fddate" & _
						"                                                                                                           AND iplasprod.harvestingspb.driver =" & _
						"                                                                                                                  iplasprod.attendance_gad.employeecode" & _
						"                                                                                                           AND iplasprod.harvestingspb.vehicle =" & _
						"                                                                                                                  iplasprod.attendance_gad.locationcode)" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.field" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.field.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.fieldcode =" & _
						"                                                                                                                  iplasprod.field.fccode)" & _
						"                                                                                              WHERE     iplasprod.harvestingspb.bunch > 0" & _
						"                                                                                                    AND iplasprod.attendance_gad.attendance IN" & _
						"                                                                                                           ('KJ'," & _
						"                                                                                                            'WH'," & _
						"                                                                                                            'WS'," & _
						"                                                                                                            'KK')" & _
						"                                                                                                    AND iplasprod.attendance_gad.locationtype =" & _
						"                                                                                                           'VH'" & _
						"                                                                                                    AND iplasprod.attendance_gad.locationcode LIKE" & _
						"                                                                                                           'DT%'" & _
						"                                                                                           GROUP BY iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division" & _
						"                                                                                           UNION ALL" & _
						"                                                                                             SELECT iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division, 0" & _
						"                                                                                                       AS total_bunch," & _
						"                                                                                                    NVL (" & _
						"                                                                                                       SUM (" & _
						"                                                                                                          iplasprod.harvestingspb.bucket)," & _
						"                                                                                                       0)" & _
						"                                                                                                       AS total_bucket" & _
						"                                                                                               FROM iplasprod.harvestingspb" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.attendance_gad" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.attendance_gad.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.receptiondate =" & _
						"                                                                                                                  iplasprod.attendance_gad.fddate" & _
						"                                                                                                           AND iplasprod.harvestingspb.driver =" & _
						"                                                                                                                  iplasprod.attendance_gad.employeecode" & _
						"                                                                                                           AND iplasprod.harvestingspb.vehicle =" & _
						"                                                                                                                  iplasprod.attendance_gad.locationcode)" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.field" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.field.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.fieldcode =" & _
						"                                                                                                                  iplasprod.field.fccode)" & _
						"                                                                                             WHERE     iplasprod.harvestingspb.bucket > 0" & _
						"                                                                                             AND iplasprod.attendance_gad.attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
						"                                                                                             AND iplasprod.attendance_gad.locationtype = 'VH'" & _
						"                                                                                             AND iplasprod.attendance_gad.locationcode LIKE 'DT%'" & _
						"                                                                                             AND iplasprod.attendance_gad.fcba = '"& v_txtfcba &"'" & _
						"                                                                                           GROUP BY iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division)" & _
						"                                                                                 GROUP BY fcba," & _
						"                                                                                          spbno," & _
						"                                                                                          receptiondate," & _
						"                                                                                          vehicle," & _
						"                                                                                          driver," & _
						"                                                                                          attendance," & _
						"                                                                                          division)" & _
						"                                                                                bunchbucket" & _
						"                                                                                INNER JOIN" & _
						"                                                                                iplasprod.wb_harvesting" & _
						"                                                                                   ON (    bunchbucket.fcba =" & _
						"                                                                                              iplasprod.wb_harvesting.fcba_estate" & _
						"                                                                                       AND bunchbucket.spbno =" & _
						"                                                                                              iplasprod.wb_harvesting.spbno" & _
						"                                                                                       AND bunchbucket.vehicle =" & _
						"                                                                                              iplasprod.wb_harvesting.vehiclecode)" & _
						"                                                                                              WHERE fcba = '"& v_txtfcba &"')" & _
						"                                                                                       WHERE fcba = '"& v_txtfcba &"'" & _
						"                                                               GROUP BY fcba," & _
						"                                                                        receptiondate," & _
						"                                                                        driver," & _
						"                                                                        attendance," & _
						"                                                                        division)" & _
						"                                                              kg_tbs_brd" & _
						"                                                              INNER JOIN" & _
						"                                                              (SELECT DISTINCT" & _
						"                                                                      iplasprod.multilevel_operator_rate.fcba," & _
						"                                                                      iplasprod.multilevel_operator_rate.attendance," & _
						"                                                                      iplasprod.multilevel_operator_rate.division," & _
						"                                                                      iplasprod.multilevel_operator_rate.basis," & _
						"                                                                      iplasprod.multilevel_operator_rate.brondolan_rate," & _
						"                                                                      iplasprod.multilevel_operator_rate.lain_lain_rate" & _
						"                                                                 FROM iplasprod.multilevel_operator_rate" & _
						"                                                                 WHERE iplasprod.multilevel_operator_rate.fcba = '"& v_txtfcba &"')" & _
						"                                                              mor" & _
						"                                                                 ON (    kg_tbs_brd.fcba =" & _
						"                                                                            mor.fcba" & _
						"                                                                     AND kg_tbs_brd.division =" & _
						"                                                                            mor.division" & _
						"                                                                     AND kg_tbs_brd.attendance =" & _
						"                                                                            mor.attendance)" & _
						"                                                              INNER JOIN" & _
						"                                                              (  SELECT fcba," & _
						"                                                                        attendance," & _
						"                                                                        division," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              1, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_1," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              2, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_2," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              3, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_3," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              4, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_4," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              1, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_1," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              2, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_2," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              3, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_3," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              4, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_4" & _
						"                                                                   FROM (SELECT iplasprod.multilevel_operator_rate.fcba," & _
						"                                                                                iplasprod.multilevel_operator_rate.attendance," & _
						"                                                                                iplasprod.multilevel_operator_rate.levelno," & _
						"                                                                                iplasprod.multilevel_operator_rate.division," & _
						"                                                                                iplasprod.multilevel_operator_rate.basis_extra," & _
						"                                                                                iplasprod.multilevel_operator_rate.extra_rate" & _
						"                                                                           FROM iplasprod.multilevel_operator_rate" & _
						"                                                                           WHERE iplasprod.multilevel_operator_rate.fcba = '"& v_txtfcba &"')" & _
						"                                                               GROUP BY fcba," & _
						"                                                                        attendance," & _
						"                                                                        division)" & _
						"                                                              mor_level" & _
						"                                                                 ON (    kg_tbs_brd.fcba =" & _
						"                                                                            mor_level.fcba" & _
						"                                                                     AND kg_tbs_brd.division =" & _
						"                                                                            mor_level.division" & _
						"                                                                     AND kg_tbs_brd.attendance =" & _
						"                                                                            mor_level.attendance))" & _
						"                                                      hasiltbsbrd" & _
						"                                             GROUP BY hasiltbsbrd.fcba," & _
						"                                                      hasiltbsbrd.receptiondate," & _
						"                                                      hasiltbsbrd.Operator," & _
						"                                                      hasiltbsbrd.attendance," & _
						"                                                      hasiltbsbrd.basis," & _
						"                                                      hasiltbsbrd.basis_extra_level_1," & _
						"                                                      hasiltbsbrd.basis_extra_level_2," & _
						"                                                      hasiltbsbrd.basis_extra_level_3," & _
						"                                                      hasiltbsbrd.basis_extra_level_4," & _
						"                                                      hasiltbsbrd.extra_rate_level_1," & _
						"                                                      hasiltbsbrd.extra_rate_level_2," & _
						"                                                      hasiltbsbrd.extra_rate_level_3," & _
						"                                                      hasiltbsbrd.extra_rate_level_4," & _
						"                                                      hasiltbsbrd.brondolan_rate," & _
						"                                                      hasiltbsbrd.lain_lain_rate)" & _
						"                                     UNION ALL" & _
						"                                     SELECT fcba," & _
						"                                            receptiondate," & _
						"                                            Operator," & _
						"                                            attendance," & _
						"                                            0 AS kg_tbs," & _
						"                                            0 AS kg_brondolan," & _
						"                                            Kg_Pupuk," & _
						"                                            0 AS Kg_Lain_Lain," & _
						"                                            0 AS basis," & _
						"                                            0 AS basis_extra_level_1," & _
						"                                            0 AS basis_extra_level_2," & _
						"                                            0 AS basis_extra_level_3," & _
						"                                            0 AS basis_extra_level_4," & _
						"                                            0 AS extra_rate_level_1," & _
						"                                            0 AS extra_rate_level_2," & _
						"                                            0 AS extra_rate_level_3," & _
						"                                            0 AS extra_rate_level_4," & _
						"                                            0 AS brondolan_rate," & _
						"                                            0 AS lain_lain_rate" & _
						"                                       FROM (  SELECT fcba," & _
						"                                                      fddate AS receptiondate," & _
						"                                                      Operator," & _
						"                                                      attendance," & _
						"                                                      SUM (" & _
						"                                                         DECODE (" & _
						"                                                            categorytransportdesc," & _
						"                                                            'Angkut Pupuk', output," & _
						"                                                            0))" & _
						"                                                         AS Kg_Pupuk" & _
						"                                                 FROM (  SELECT iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode" & _
						"                                                                   AS Operator," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                iplasprod.vehicleactivity.categorytransport," & _
						"                                                                iplasprod.vehicleactivitycategory.fcname" & _
						"                                                                   AS categorytransportdesc," & _
						"                                                                NVL (" & _
						"                                                                   SUM (" & _
						"                                                                      iplasprod.vehicleactivity.output)," & _
						"                                                                   0)" & _
						"                                                                   AS output" & _
						"                                                           FROM iplasprod.attendance_gad" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivity" & _
						"                                                                   ON (    iplasprod.attendance_gad.fcba =" & _
						"                                                                              iplasprod.vehicleactivity.fcba" & _
						"                                                                       AND iplasprod.attendance_gad.fddate =" & _
						"                                                                              iplasprod.vehicleactivity.vehicledate" & _
						"                                                                       AND iplasprod.attendance_gad.employeecode =" & _
						"                                                                              iplasprod.vehicleactivity.employeecode" & _
						"                                                                       AND iplasprod.attendance_gad.locationcode =" & _
						"                                                                              iplasprod.vehicleactivity.vehiclecode)" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivitycategory" & _
						"                                                                   ON (iplasprod.vehicleactivity.categorytransport =" & _
						"                                                                          iplasprod.vehicleactivitycategory.fccode)" & _
						"                                                          WHERE     iplasprod.attendance_gad.attendance IN" & _
						"                                                                       ('KJ'," & _
						"                                                                        'WH'," & _
						"                                                                        'WS'," & _
						"                                                                        'KK')" & _
						"                                                                AND iplasprod.attendance_gad.fcba = '"& v_txtfcba &"'" & _
						"                                                                AND iplasprod.attendance_gad.locationtype = 'VH'" & _
						"                                                                AND iplasprod.attendance_gad.locationcode LIKE 'DT%'" & _
						"                                                                AND iplasprod.vehicleactivitycategory.fcname = 'Angkut Pupuk'" & _
						"                                                       GROUP BY iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                iplasprod.vehicleactivity.categorytransport," & _
						"                                                                iplasprod.vehicleactivitycategory.fcname)" & _
						"                                                       WHERE fcba = '"& v_txtfcba &"'" & _
						"                                             GROUP BY fcba," & _
						"                                                      fddate," & _
						"                                                      Operator," & _
						"                                                      attendance)" & _
						"                                             WHERE fcba = '"& v_txtfcba &"'" & _
						"                                     UNION ALL" & _
						"                                     SELECT fcba," & _
						"                                            receptiondate," & _
						"                                            Operator," & _
						"                                            attendance," & _
						"                                            0 AS kg_tbs," & _
						"                                            0 AS kg_brondolan," & _
						"                                            0 AS Kg_Pupuk," & _
						"                                            Kg_Lain_Lain," & _
						"                                            0 AS basis," & _
						"                                            0 AS basis_extra_level_1," & _
						"                                            0 AS basis_extra_level_2," & _
						"                                            0 AS basis_extra_level_3," & _
						"                                            0 AS basis_extra_level_4," & _
						"                                            0 AS extra_rate_level_1," & _
						"                                            0 AS extra_rate_level_2," & _
						"                                            0 AS extra_rate_level_3," & _
						"                                            0 AS extra_rate_level_4," & _
						"                                            0 AS brondolan_rate," & _
						"                                            0 AS lain_lain_rate" & _
						"                                       FROM (  SELECT fcba," & _
						"                                                      fddate AS receptiondate," & _
						"                                                      Operator," & _
						"                                                      attendance," & _
						"                                                      SUM (" & _
						"                                                         DECODE (" & _
						"                                                            categorytransportdesc," & _
						"                                                            'Angkut Lain-Lain', output," & _
						"                                                            0))" & _
						"                                                         AS Kg_Lain_Lain" & _
						"                                                 FROM (  SELECT iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode" & _
						"                                                                   AS Operator," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                'Angkut Lain-Lain'" & _
						"                                                                   AS categorytransportdesc," & _
						"                                                                NVL (" & _
						"                                                                   SUM (" & _
						"                                                                      iplasprod.vehicleactivity.output)," & _
						"                                                                   0)" & _
						"                                                                   AS output" & _
						"                                                           FROM iplasprod.attendance_gad" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivity" & _
						"                                                                   ON (    iplasprod.attendance_gad.fcba =" & _
						"                                                                              iplasprod.vehicleactivity.fcba" & _
						"                                                                       AND iplasprod.attendance_gad.fddate =" & _
						"                                                                              iplasprod.vehicleactivity.vehicledate" & _
						"                                                                       AND iplasprod.attendance_gad.employeecode =" & _
						"                                                                              iplasprod.vehicleactivity.employeecode" & _
						"                                                                       AND iplasprod.attendance_gad.locationcode =" & _
						"                                                                              iplasprod.vehicleactivity.vehiclecode)" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivitycategory" & _
						"                                                                   ON (iplasprod.vehicleactivity.categorytransport =" & _
						"                                                                          iplasprod.vehicleactivitycategory.fccode)" & _
						"                                                          WHERE     iplasprod.attendance_gad.attendance IN" & _
						"                                                                       ('KJ'," & _
						"                                                                        'WH'," & _
						"                                                                        'WS'," & _
						"                                                                        'KK')" & _
						"                                                                AND iplasprod.attendance_gad.fcba = '"& v_txtfcba &"'" & _
						"                                                                AND iplasprod.attendance_gad.locationtype = 'VH'" & _
						"                                                                AND iplasprod.attendance_gad.locationcode LIKE 'DT%'" & _
						"                                                                AND iplasprod.vehicleactivitycategory.fcname NOT IN" & _
						"                                                                       ('Angkut TBS Langsung'," & _
						"                                                                        'Angkut TBS Tak Langsung'," & _
						"                                                                        'Angkut Brondolan'," & _
						"                                                                        'Angkut Pupuk')" & _
						"                                                       GROUP BY iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                'Angkut Lain-Lain')" & _
						"                                                       WHERE fcba = '"& v_txtfcba &"'" & _
						"                                             GROUP BY fcba," & _
						"                                                      fddate," & _
						"                                                      Operator," & _
						"                                                      attendance)" & _
						"                                             WHERE fcba = '"& v_txtfcba &"')" & _
						"                           GROUP BY fcba," & _
						"                                    receptiondate," & _
						"                                    Operator," & _
						"                                    attendance))) premi" & _
						"          INNER JOIN iplasprod.employee emp" & _
						"             ON (premi.fcba = emp.fcba AND premi.operatorcode = emp.fccode)" & _
						"    WHERE premi.fcba = '"& v_txtfcba &"'" 
						 

		if  len(trim(v_txtdatefrom)) > 0 then
			strGrnInv = strGrnInv & " AND premi.transdate >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
			txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
		end if
		if  len(trim(v_txtdateuntil)) > 0 then
			strGrnInv = strGrnInv & " AND premi.transdate <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
			txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
		end if  
		if  len(trim(v_txtempcode)) > 0 then
		    strGrnInv	=	strGrnInv & " AND premi.operatorcode = '"& v_txtempcode &"'" 
		    txtSelect = txtSelect & " Employee Code " & v_txtempcode
		end if 
									
		strGrnInv	=	strGrnInv & " ORDER BY premi.fcba," & _
									"          premi.operatorcode," & _
									"          premi.transdate"

		' response.write strGrnInv	
		Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
		rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	Else
		strGrnInv	=	"   SELECT premi.fcba," & _
						"          premi.transdate," & _
						"          premi.operatorcode," & _
						"          emp.fcname AS operatorname," & _
						"          premi.attendance," & _
						"          premi.kg_tbs," & _
						"          premi.kg_brondolan," & _
						"          premi.kg_pupuk," & _
						"          premi.kg_lain_lain," & _
						"          premi.total_kg," & _
						"          premi.basis," & _
						"          premi.kg_premi_lebih_basis," & _
						"          premi.total_premi," & _
						"          CASE" & _
						"             WHEN premi.kg_premi_lebih_basis > 0 THEN 1" & _
						"             WHEN premi.kg_lain_lain > 0 THEN 1" & _
						"             ELSE 0" & _
						"          END" & _
						"             AS HK_Efektif" & _
						"     FROM (SELECT fcba," & _
						"                  receptiondate AS transdate," & _
						"                  Operator AS operatorcode," & _
						"                  attendance," & _
						"                  kg_tbs," & _
						"                  kg_brondolan," & _
						"                  Kg_Pupuk," & _
						"                  Kg_Lain_Lain," & _
						"                  total_kg," & _
						"                  basis," & _
						"                  kg_premi_lebih_basis," & _
						"                  kg_premi_brondolan," & _
						"                  premi_level_1," & _
						"                  premi_level_2," & _
						"                  premi_level_3," & _
						"                  premi_level_4," & _
						"                  extra_rate_level_1," & _
						"                  extra_rate_level_2," & _
						"                  extra_rate_level_3," & _
						"                  extra_rate_level_4," & _
						"                  brondolan_rate," & _
						"                  lain_lain_rate," & _
						"                  (  (premi_level_1 * extra_rate_level_1)" & _
						"                   + (premi_level_2 * extra_rate_level_2)" & _
						"                   + (premi_level_3 * extra_rate_level_3)" & _
						"                   + (premi_level_4 * extra_rate_level_4))" & _
						"                     AS total_premi" & _
						"             FROM (SELECT fcba," & _
						"                          receptiondate," & _
						"                          Operator," & _
						"                          attendance," & _
						"                          kg_tbs," & _
						"                          kg_brondolan," & _
						"                          Kg_Pupuk," & _
						"                          Kg_Lain_Lain," & _
						"                          total_kg," & _
						"                          basis," & _
						"                          kg_premi_lebih_basis," & _
						"                          0 AS kg_premi_brondolan," & _
						"                          CASE" & _
						"                             WHEN kg_premi_lebih_basis > basis_extra_level_1" & _
						"                             THEN" & _
						"                                basis_extra_level_1" & _
						"                             ELSE" & _
						"                                kg_premi_lebih_basis" & _
						"                          END" & _
						"                             AS premi_level_1," & _
						"                          CASE" & _
						"                             WHEN (kg_premi_lebih_basis - basis_extra_level_1) <= 0" & _
						"                             THEN 0" & _
						"                             WHEN (kg_premi_lebih_basis - basis_extra_level_1) >" & _
						"                                     basis_extra_level_2" & _
						"                             THEN" & _
						"                                basis_extra_level_2" & _
						"                             ELSE" & _
						"                                (kg_premi_lebih_basis - basis_extra_level_1)" & _
						"                          END" & _
						"                             AS premi_level_2," & _
						"                          CASE" & _
						"                             WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2) <= 0" & _
						"                             THEN  0" & _
						"                             WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2) > basis_extra_level_3" & _
						"                             THEN" & _
						"                                basis_extra_level_3" & _
						"                             ELSE" & _
						"                                (  kg_premi_lebih_basis" & _
						"                                 - basis_extra_level_1" & _
						"                                 - basis_extra_level_2)" & _
						"                          END" & _
						"                             AS premi_level_3," & _
						"                          CASE" & _
						"                             WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2" & _
						"                                   - basis_extra_level_3) <= 0" & _
						"                             THEN 0" & _
						"                             WHEN (  kg_premi_lebih_basis" & _
						"                                   - basis_extra_level_1" & _
						"                                   - basis_extra_level_2" & _
						"                                   - basis_extra_level_3) > basis_extra_level_4" & _
						"                             THEN" & _
						"                                (  kg_premi_lebih_basis" & _
						"                                 - basis_extra_level_1" & _
						"                                 - basis_extra_level_2" & _
						"                                 - basis_extra_level_3)" & _
						"                             ELSE" & _
						"                                (  kg_premi_lebih_basis" & _
						"                                 - basis_extra_level_1" & _
						"                                 - basis_extra_level_2" & _
						"                                 - basis_extra_level_3)" & _
						"                          END" & _
						"                             AS premi_level_4," & _
						"                          extra_rate_level_1," & _
						"                          extra_rate_level_2," & _
						"                          extra_rate_level_3," & _
						"                          extra_rate_level_4," & _
						"                          brondolan_rate," & _
						"                          lain_lain_rate" & _
						"                     FROM (  SELECT fcba," & _
						"                                    receptiondate," & _
						"                                    Operator," & _
						"                                    attendance," & _
						"                                    SUM (kg_tbs) AS kg_tbs," & _
						"                                    SUM (kg_brondolan) AS kg_brondolan," & _
						"                                    SUM (Kg_Pupuk) AS Kg_Pupuk," & _
						"                                    SUM (Kg_Lain_Lain) AS Kg_Lain_Lain," & _
						"                                    (  SUM (kg_tbs)" & _
						"                                     + SUM (kg_brondolan)" & _
						"                                     + SUM (Kg_Pupuk)" & _
						"                                     + SUM (Kg_Lain_Lain))" & _
						"                                       AS total_kg," & _
						"                                    MAX (basis) AS basis," & _
						"                                    CASE" & _
						"                                       WHEN (  SUM (kg_tbs)" & _
						"                                             + SUM (kg_brondolan)" & _
						"                                             + SUM (Kg_Pupuk)" & _
						"                                             + SUM (Kg_Lain_Lain)) > MAX (basis)" & _
						"                                       THEN" & _
						"                                            (  SUM (kg_tbs)" & _
						"                                             + SUM (kg_brondolan)" & _
						"                                             + SUM (Kg_Pupuk)" & _
						"                                             + SUM (Kg_Lain_Lain))" & _
						"                                          - MAX (basis)" & _
						"                                       ELSE  0" & _
						"                                    END" & _
						"                                       AS kg_premi_lebih_basis," & _
						"                                    MAX (basis_extra_level_1)" & _
						"                                       AS basis_extra_level_1," & _
						"                                    MAX (basis_extra_level_2)" & _
						"                                       AS basis_extra_level_2," & _
						"                                    MAX (basis_extra_level_3)" & _
						"                                       AS basis_extra_level_3," & _
						"                                    MAX (basis_extra_level_4)" & _
						"                                       AS basis_extra_level_4," & _
						"                                    MAX (extra_rate_level_1) AS extra_rate_level_1," & _
						"                                    MAX (extra_rate_level_2) AS extra_rate_level_2," & _
						"                                    MAX (extra_rate_level_3) AS extra_rate_level_3," & _
						"                                    MAX (extra_rate_level_4) AS extra_rate_level_4," & _
						"                                    MAX (brondolan_rate) AS brondolan_rate," & _
						"                                    MAX (lain_lain_rate) AS lain_lain_rate" & _
						"                               FROM (SELECT fcba," & _
						"                                            receptiondate," & _
						"                                            Operator," & _
						"                                            attendance," & _
						"                                            kg_tbs," & _
						"                                            kg_brondolan," & _
						"                                            0 AS Kg_Pupuk," & _
						"                                            0 AS Kg_Lain_Lain," & _
						"                                            basis," & _
						"                                            basis_extra_level_1," & _
						"                                            basis_extra_level_2," & _
						"                                            basis_extra_level_3," & _
						"                                            basis_extra_level_4," & _
						"                                            extra_rate_level_1," & _
						"                                            extra_rate_level_2," & _
						"                                            extra_rate_level_3," & _
						"                                            extra_rate_level_4," & _
						"                                            brondolan_rate," & _
						"                                            lain_lain_rate" & _
						"                                       FROM (  SELECT hasiltbsbrd.fcba," & _
						"                                                      hasiltbsbrd.receptiondate," & _
						"                                                      hasiltbsbrd.Operator," & _
						"                                                      hasiltbsbrd.attendance," & _
						"                                                      SUM (hasiltbsbrd.Kg_TBS)" & _
						"                                                         AS kg_tbs," & _
						"                                                      SUM (hasiltbsbrd.Kg_Brondolan)" & _
						"                                                         AS kg_brondolan," & _
						"                                                      hasiltbsbrd.basis," & _
						"                                                      hasiltbsbrd.basis_extra_level_1," & _
						"                                                      hasiltbsbrd.basis_extra_level_2," & _
						"                                                      hasiltbsbrd.basis_extra_level_3," & _
						"                                                      hasiltbsbrd.basis_extra_level_4," & _
						"                                                      hasiltbsbrd.extra_rate_level_1," & _
						"                                                      hasiltbsbrd.extra_rate_level_2," & _
						"                                                      hasiltbsbrd.extra_rate_level_3," & _
						"                                                      hasiltbsbrd.extra_rate_level_4," & _
						"                                                      hasiltbsbrd.brondolan_rate," & _
						"                                                      hasiltbsbrd.lain_lain_rate" & _
						"                                                 FROM (SELECT kg_tbs_brd.fcba," & _
						"                                                              kg_tbs_brd.receptiondate," & _
						"                                                              kg_tbs_brd.Operator," & _
						"                                                              kg_tbs_brd.attendance," & _
						"                                                              kg_tbs_brd.division," & _
						"                                                                 kg_tbs_brd.fcba" & _
						"                                                              || kg_tbs_brd.receptiondate" & _
						"                                                              || kg_tbs_brd.operator," & _
						"                                                              kg_tbs_brd.Kg_TBS," & _
						"                                                              kg_tbs_brd.Kg_Brondolan," & _
						"                                                              AVG (" & _
						"                                                                 mor.basis)" & _
						"                                                              OVER (" & _
						"                                                                 PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis," & _
						"                                                              AVG (" & _
						"                                                                 mor_level.basis_extra_level_1)" & _
						"                                                              OVER (" & _
						"                                                                 PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_1," & _
						"                                                              AVG (" & _
						"                                                                 mor_level.basis_extra_level_2)" & _
						"                                                              OVER (" & _
						"                                                                 PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_2," & _
						"                                                              AVG (" & _
						"                                                                 mor_level.basis_extra_level_3)" & _
						"                                                              OVER (" & _
						"                                                                 PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_3," & _
						"                                                              AVG (" & _
						"                                                                 mor_level.basis_extra_level_4)" & _
						"                                                              OVER (" & _
						"                                                                 PARTITION BY    kg_tbs_brd.fcba" & _
						"                                                                              || kg_tbs_brd.receptiondate" & _
						"                                                                              || kg_tbs_brd.operator)" & _
						"                                                                 AS basis_extra_level_4," & _
						"                                                              mor_level.extra_rate_level_1," & _
						"                                                              mor_level.extra_rate_level_2," & _
						"                                                              mor_level.extra_rate_level_3," & _
						"                                                              mor_level.extra_rate_level_4," & _
						"                                                              mor.brondolan_rate," & _
						"                                                              mor.lain_lain_rate" & _
						"                                                         FROM (  SELECT fcba," & _
						"                                                                        receptiondate," & _
						"                                                                        driver" & _
						"                                                                           AS Operator," & _
						"                                                                        attendance," & _
						"                                                                        division," & _
						"                                                                        SUM (kg_bunch)" & _
						"                                                                           AS Kg_TBS," & _
						"                                                                        SUM (" & _
						"                                                                           kg_bucket)" & _
						"                                                                           AS Kg_Brondolan" & _
						"                                                                   FROM (SELECT bunchbucket.fcba," & _
						"                                                                                bunchbucket.spbno," & _
						"                                                                                bunchbucket.receptiondate," & _
						"                                                                                bunchbucket.vehicle," & _
						"                                                                                bunchbucket.driver," & _
						"                                                                                bunchbucket.attendance," & _
						"                                                                                bunchbucket.division," & _
						"                                                                                bunchbucket.total_bunch," & _
						"                                                                                iplasprod.wb_harvesting.bjr_per_chitno," & _
						"                                                                                (  bunchbucket.total_bunch" & _
						"                                                                                 * iplasprod.wb_harvesting.bjr_per_chitno)" & _
						"                                                                                   AS Kg_Bunch," & _
						"                                                                                bunchbucket.total_bucket," & _
						"                                                                                iplasprod.wb_harvesting.kg_per_bucket," & _
						"                                                                                (  bunchbucket.total_bucket" & _
						"                                                                                 * iplasprod.wb_harvesting.kg_per_bucket)" & _
						"                                                                                   AS Kg_Bucket" & _
						"                                                                           FROM (  SELECT fcba," & _
						"                                                                                          spbno," & _
						"                                                                                          receptiondate," & _
						"                                                                                          vehicle," & _
						"                                                                                          driver," & _
						"                                                                                          attendance," & _
						"                                                                                          division," & _
						"                                                                                          SUM (" & _
						"                                                                                             total_bunch)" & _
						"                                                                                             AS total_bunch," & _
						"                                                                                          SUM (" & _
						"                                                                                             total_bucket)" & _
						"                                                                                             AS total_bucket" & _
						"                                                                                     FROM (  SELECT iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division," & _
						"                                                                                                    NVL (" & _
						"                                                                                                       SUM (" & _
						"                                                                                                          iplasprod.harvestingspb.bunch)," & _
						"                                                                                                       0)" & _
						"                                                                                                       AS total_bunch,  0" & _
						"                                                                                                       AS total_bucket" & _
						"                                                                                               FROM iplasprod.harvestingspb" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.attendance_gad" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.attendance_gad.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.receptiondate =" & _
						"                                                                                                                  iplasprod.attendance_gad.fddate" & _
						"                                                                                                           AND iplasprod.harvestingspb.driver =" & _
						"                                                                                                                  iplasprod.attendance_gad.employeecode" & _
						"                                                                                                           AND iplasprod.harvestingspb.vehicle =" & _
						"                                                                                                                  iplasprod.attendance_gad.locationcode)" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.field" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.field.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.fieldcode =" & _
						"                                                                                                                  iplasprod.field.fccode)" & _
						"                                                                                              WHERE     iplasprod.harvestingspb.bunch > 0" & _
						"                                                                                                    AND iplasprod.attendance_gad.attendance IN" & _
						"                                                                                                           ('KJ'," & _
						"                                                                                                            'WH'," & _
						"                                                                                                            'WS'," & _
						"                                                                                                            'KK')" & _
						"                                                                                                    AND iplasprod.attendance_gad.fcba = '"& v_txtFcba &"'" & _
						"                                                                                                    AND iplasprod.attendance_gad.locationtype = 'VH'" & _
						"                                                                                                    AND iplasprod.attendance_gad.locationcode LIKE 'DT%'" & _
						"                                                                                           GROUP BY iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division" & _
						"                                                                                           UNION ALL" & _
						"                                                                                             SELECT iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division,  0" & _
						"                                                                                                       AS total_bunch," & _
						"                                                                                                    NVL (" & _
						"                                                                                                       SUM (" & _
						"                                                                                                          iplasprod.harvestingspb.bucket)," & _
						"                                                                                                       0)" & _
						"                                                                                                       AS total_bucket" & _
						"                                                                                               FROM iplasprod.harvestingspb" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.attendance_gad" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.attendance_gad.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.receptiondate =" & _
						"                                                                                                                  iplasprod.attendance_gad.fddate" & _
						"                                                                                                           AND iplasprod.harvestingspb.driver =" & _
						"                                                                                                                  iplasprod.attendance_gad.employeecode" & _
						"                                                                                                           AND iplasprod.harvestingspb.vehicle =" & _
						"                                                                                                                  iplasprod.attendance_gad.locationcode)" & _
						"                                                                                                    INNER JOIN" & _
						"                                                                                                    iplasprod.field" & _
						"                                                                                                       ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                                                                                  iplasprod.field.fcba" & _
						"                                                                                                           AND iplasprod.harvestingspb.fieldcode =" & _
						"                                                                                                                  iplasprod.field.fccode)" & _
						"                                                                                              WHERE     iplasprod.harvestingspb.bucket > 0" & _
						"                                                                                                    AND iplasprod.attendance_gad.attendance IN" & _
						"                                                                                                           ('KJ'," & _
						"                                                                                                            'WH'," & _
						"                                                                                                            'WS'," & _
						"                                                                                                            'KK')" & _
						"                                                                                                    AND iplasprod.attendance_gad.locationtype =" & _
						"                                                                                                           'VH'" & _
						"                                                                                                    AND iplasprod.attendance_gad.locationcode LIKE" & _
						"                                                                                                           'DT%'" & _
						"                                                                                           GROUP BY iplasprod.harvestingspb.fcba," & _
						"                                                                                                    iplasprod.harvestingspb.spbno," & _
						"                                                                                                    iplasprod.harvestingspb.receptiondate," & _
						"                                                                                                    iplasprod.harvestingspb.vehicle," & _
						"                                                                                                    iplasprod.harvestingspb.driver," & _
						"                                                                                                    iplasprod.attendance_gad.attendance," & _
						"                                                                                                    iplasprod.field.division)" & _
						"                                                                                 GROUP BY fcba," & _
						"                                                                                          spbno," & _
						"                                                                                          receptiondate," & _
						"                                                                                          vehicle," & _
						"                                                                                          driver," & _
						"                                                                                          attendance," & _
						"                                                                                          division)" & _
						"                                                                                bunchbucket" & _
						"                                                                                INNER JOIN" & _
						"                                                                                iplasprod.wb_harvesting" & _
						"                                                                                   ON (    bunchbucket.fcba =" & _
						"                                                                                              iplasprod.wb_harvesting.fcba_estate" & _
						"                                                                                       AND bunchbucket.spbno =" & _
						"                                                                                              iplasprod.wb_harvesting.spbno" & _
						"                                                                                       AND bunchbucket.vehicle =" & _
						"                                                                                              iplasprod.wb_harvesting.vehiclecode))" & _
						"                                                               GROUP BY fcba," & _
						"                                                                        receptiondate," & _
						"                                                                        driver," & _
						"                                                                        attendance," & _
						"                                                                        division)" & _
						"                                                              kg_tbs_brd" & _
						"                                                              INNER JOIN" & _
						"                                                              (SELECT DISTINCT" & _
						"                                                                      iplasprod.multilevel_operator_rate.fcba," & _
						"                                                                      iplasprod.multilevel_operator_rate.attendance," & _
						"                                                                      iplasprod.multilevel_operator_rate.division," & _
						"                                                                      iplasprod.multilevel_operator_rate.basis," & _
						"                                                                      iplasprod.multilevel_operator_rate.brondolan_rate," & _
						"                                                                      iplasprod.multilevel_operator_rate.lain_lain_rate" & _
						"                                                                 FROM iplasprod.multilevel_operator_rate" & _
						"                                                                 WHERE iplasprod.multilevel_operator_rate.fcba = '"& v_txtfcba &"')" & _
						"                                                              mor" & _
						"                                                                 ON (    kg_tbs_brd.fcba =" & _
						"                                                                            mor.fcba" & _
						"                                                                     AND kg_tbs_brd.division =" & _
						"                                                                            mor.division" & _
						"                                                                     AND kg_tbs_brd.attendance =" & _
						"                                                                            mor.attendance)" & _
						"                                                              INNER JOIN" & _
						"                                                              (  SELECT fcba," & _
						"                                                                        attendance," & _
						"                                                                        division," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              1, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_1," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              2, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_2," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              3, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_3," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              4, basis_extra," & _
						"                                                                              0))" & _
						"                                                                           basis_extra_level_4," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              1, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_1," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              2, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_2," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              3, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_3," & _
						"                                                                        SUM (" & _
						"                                                                           DECODE (" & _
						"                                                                              levelno," & _
						"                                                                              4, extra_rate," & _
						"                                                                              0))" & _
						"                                                                           extra_rate_level_4" & _
						"                                                                   FROM (SELECT iplasprod.multilevel_operator_rate.fcba," & _
						"                                                                                iplasprod.multilevel_operator_rate.attendance," & _
						"                                                                                iplasprod.multilevel_operator_rate.levelno," & _
						"                                                                                iplasprod.multilevel_operator_rate.division," & _
						"                                                                                iplasprod.multilevel_operator_rate.basis_extra," & _
						"                                                                                iplasprod.multilevel_operator_rate.extra_rate" & _
						"                                                                           FROM iplasprod.multilevel_operator_rate" & _
						"                                                                           WHERE iplasprod.multilevel_operator_rate.fcba = '"& v_txtfcba &"')" & _
						"                                                               GROUP BY fcba," & _
						"                                                                        attendance," & _
						"                                                                        division)" & _
						"                                                              mor_level" & _
						"                                                                 ON (    kg_tbs_brd.fcba =" & _
						"                                                                            mor_level.fcba" & _
						"                                                                     AND kg_tbs_brd.division =" & _
						"                                                                            mor_level.division" & _
						"                                                                     AND kg_tbs_brd.attendance =" & _
						"                                                                            mor_level.attendance))" & _
						"                                                      hasiltbsbrd" & _
						"                                             GROUP BY hasiltbsbrd.fcba," & _
						"                                                      hasiltbsbrd.receptiondate," & _
						"                                                      hasiltbsbrd.Operator," & _
						"                                                      hasiltbsbrd.attendance," & _
						"                                                      hasiltbsbrd.basis," & _
						"                                                      hasiltbsbrd.basis_extra_level_1," & _
						"                                                      hasiltbsbrd.basis_extra_level_2," & _
						"                                                      hasiltbsbrd.basis_extra_level_3," & _
						"                                                      hasiltbsbrd.basis_extra_level_4," & _
						"                                                      hasiltbsbrd.extra_rate_level_1," & _
						"                                                      hasiltbsbrd.extra_rate_level_2," & _
						"                                                      hasiltbsbrd.extra_rate_level_3," & _
						"                                                      hasiltbsbrd.extra_rate_level_4," & _
						"                                                      hasiltbsbrd.brondolan_rate," & _
						"                                                      hasiltbsbrd.lain_lain_rate)" & _
						"                                     UNION ALL" & _
						"                                     SELECT fcba," & _
						"                                            receptiondate," & _
						"                                            Operator," & _
						"                                            attendance," & _
						"                                            0 AS kg_tbs," & _
						"                                            0 AS kg_brondolan," & _
						"                                            Kg_Pupuk," & _
						"                                            0 AS Kg_Lain_Lain," & _
						"                                            3500 AS basis," & _
						"                                            3500 AS basis_extra_level_1," & _
						"                                            3500 AS basis_extra_level_2," & _
						"                                            3500 AS basis_extra_level_3," & _
						"                                            3500 AS basis_extra_level_4," & _
						"                                            4 AS extra_rate_level_1," & _
						"                                            5 AS extra_rate_level_2," & _
						"                                            6.75 AS extra_rate_level_3," & _
						"                                            10.125 AS extra_rate_level_4," & _
						"                                            1 AS brondolan_rate," & _
						"                                            1 AS lain_lain_rate" & _
						"                                       FROM (  SELECT fcba," & _
						"                                                      fddate AS receptiondate," & _
						"                                                      Operator," & _
						"                                                      attendance," & _
						"                                                      SUM (" & _
						"                                                         DECODE (" & _
						"                                                            categorytransportdesc," & _
						"                                                            'Angkut Pupuk', output," & _
						"                                                            0))" & _
						"                                                         AS Kg_Pupuk" & _
						"                                                 FROM (  SELECT iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode" & _
						"                                                                   AS Operator," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                iplasprod.vehicleactivity.categorytransport," & _
						"                                                                iplasprod.vehicleactivitycategory.fcname" & _
						"                                                                   AS categorytransportdesc," & _
						"                                                                NVL (" & _
						"                                                                   SUM (" & _
						"                                                                      iplasprod.vehicleactivity.output)," & _
						"                                                                   0)" & _
						"                                                                   AS output" & _
						"                                                           FROM iplasprod.attendance_gad" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivity" & _
						"                                                                   ON (    iplasprod.attendance_gad.fcba =" & _
						"                                                                              iplasprod.vehicleactivity.fcba" & _
						"                                                                       AND iplasprod.attendance_gad.fddate =" & _
						"                                                                              iplasprod.vehicleactivity.vehicledate" & _
						"                                                                       AND iplasprod.attendance_gad.employeecode =" & _
						"                                                                              iplasprod.vehicleactivity.employeecode" & _
						"                                                                       AND iplasprod.attendance_gad.locationcode =" & _
						"                                                                              iplasprod.vehicleactivity.vehiclecode)" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivitycategory" & _
						"                                                                   ON (iplasprod.vehicleactivity.categorytransport =" & _
						"                                                                          iplasprod.vehicleactivitycategory.fccode)" & _
						"                                                          WHERE     iplasprod.attendance_gad.attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
						"                                                                AND iplasprod.attendance_gad.fcba = '"& v_txtfcba &"'" & _
						"                                                                AND iplasprod.attendance_gad.locationtype = 'VH'" & _
						"                                                                AND iplasprod.attendance_gad.locationcode LIKE 'DT%'" & _
						"                                                                AND iplasprod.vehicleactivitycategory.fcname = 'Angkut Pupuk'" & _
						"                                                       GROUP BY iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                iplasprod.vehicleactivity.categorytransport," & _
						"                                                                iplasprod.vehicleactivitycategory.fcname)" & _
						"                                                       WHERE fcba = '"& v_txtfcba &"'" & _
						"                                             GROUP BY fcba," & _
						"                                                      fddate," & _
						"                                                      Operator," & _
						"                                                      attendance)" & _
						"                                             WHERE fcba = '"& v_txtfcba &"'" & _
						"                                     UNION ALL" & _
						"                                     SELECT fcba," & _
						"                                            receptiondate," & _
						"                                            Operator," & _
						"                                            attendance," & _
						"                                            0 AS kg_tbs," & _
						"                                            0 AS kg_brondolan," & _
						"                                            0 AS Kg_Pupuk," & _
						"                                            Kg_Lain_Lain," & _
						"                                            3500 AS basis," & _
						"                                            3500 AS basis_extra_level_1," & _
						"                                            3500 AS basis_extra_level_2," & _
						"                                            3500 AS basis_extra_level_3," & _
						"                                            3500 AS basis_extra_level_4," & _
						"                                            4 AS extra_rate_level_1," & _
						"                                            5 AS extra_rate_level_2," & _
						"                                            6.75 AS extra_rate_level_3," & _
						"                                            10.125 AS extra_rate_level_4," & _
						"                                            1 AS brondolan_rate," & _
						"                                            1 AS lain_lain_rate" & _
						"                                       FROM (  SELECT fcba," & _
						"                                                      fddate AS receptiondate," & _
						"                                                      Operator," & _
						"                                                      attendance," & _
						"                                                      SUM (" & _
						"                                                         DECODE (" & _
						"                                                            categorytransportdesc," & _
						"                                                            'Angkut Lain-Lain', output," & _
						"                                                            0))" & _
						"                                                         AS Kg_Lain_Lain" & _
						"                                                 FROM (  SELECT iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode" & _
						"                                                                   AS Operator," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                'Angkut Lain-Lain'" & _
						"                                                                   AS categorytransportdesc," & _
						"                                                                NVL (" & _
						"                                                                   SUM (" & _
						"                                                                      iplasprod.vehicleactivity.output)," & _
						"                                                                   0)" & _
						"                                                                   AS output" & _
						"                                                           FROM iplasprod.attendance_gad" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivity" & _
						"                                                                   ON (    iplasprod.attendance_gad.fcba =" & _
						"                                                                              iplasprod.vehicleactivity.fcba" & _
						"                                                                       AND iplasprod.attendance_gad.fddate =" & _
						"                                                                              iplasprod.vehicleactivity.vehicledate" & _
						"                                                                       AND iplasprod.attendance_gad.employeecode =" & _
						"                                                                              iplasprod.vehicleactivity.employeecode" & _
						"                                                                       AND iplasprod.attendance_gad.locationcode =" & _
						"                                                                              iplasprod.vehicleactivity.vehiclecode)" & _
						"                                                                INNER JOIN" & _
						"                                                                iplasprod.vehicleactivitycategory" & _
						"                                                                   ON (iplasprod.vehicleactivity.categorytransport =" & _
						"                                                                          iplasprod.vehicleactivitycategory.fccode)" & _
						"                                                          WHERE     iplasprod.attendance_gad.attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
						"                                                                AND iplasprod.attendance_gad.fcba = '"& v_txtfcba &"'" & _
						"                                                                AND iplasprod.attendance_gad.locationtype = 'VH'" & _
						"                                                                AND iplasprod.attendance_gad.locationcode LIKE 'DT%'" & _
						"                                                                AND iplasprod.vehicleactivitycategory.fcname NOT IN" & _
						"                                                                       ('Angkut TBS Langsung'," & _
						"                                                                        'Angkut TBS Tak Langsung'," & _
						"                                                                        'Angkut Brondolan'," & _
						"                                                                        'Angkut Pupuk')" & _
						"                                                       GROUP BY iplasprod.attendance_gad.fcba," & _
						"                                                                iplasprod.attendance_gad.fddate," & _
						"                                                                iplasprod.attendance_gad.employeecode," & _
						"                                                                iplasprod.attendance_gad.attendance," & _
						"                                                                'Angkut Lain-Lain')" & _
						"                                                       WHERE fcba = '"& v_txtfcba &"'" & _
						"                                             GROUP BY fcba," & _
						"                                                      fddate," & _
						"                                                      Operator," & _
						"                                                      attendance))" & _
						"                           GROUP BY fcba," & _
						"                                    receptiondate," & _
						"                                    Operator," & _
						"                                    attendance))) premi" & _
						"          INNER JOIN iplasprod.employee emp" & _
						"             ON (premi.fcba = emp.fcba AND premi.operatorcode = emp.fccode)" & _
						"    WHERE     premi.fcba = '"& v_txtfcba &"'" 
						
		if  len(trim(v_txtdatefrom)) > 0 then
			strGrnInv = strGrnInv & " AND premi.transdate >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
			txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
		end if
		if  len(trim(v_txtdateuntil)) > 0 then
			strGrnInv = strGrnInv & " AND premi.transdate <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
			txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
		end if 
		if  len(trim(v_txtempcode)) > 0 then
		    strGrnInv	=	strGrnInv & " AND premi.operatorcode = '"& v_txtempcode &"'" 
		    txtSelect = txtSelect & " Employee Code " & v_txtempcode
		end if 
									
		strGrnInv	=	strGrnInv & " ORDER BY premi.fcba," & _
									"          premi.operatorcode," & _
									"          premi.transdate"

		' response.write strGrnInv	
		Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
		rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	End If
%>


<html>

<head>
<title>ePlaS Reporting - Premi Operator DT</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
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
						<td>
							<font face="verdana" size="1">KEBUN</font>
						</td>
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
							  
							
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
						<td>
							<font face="verdana" size="1">DARI</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">SAMPAI</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td> 
						<td>
							<font face="verdana" size="1">EMPLOYEE</font>
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
    		<th colspan="13" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>LAPORAN PREMI OPERATOR DT</h2>
		                <br/>
						<h4>Selected By :<%=txtSelect%></h4>
					</div> 
				</div>
			</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">OPERATOR CODE</th> 
			<th style="background-color:#00A300;color:#fff;">OPERATOR NAME</th>
			<th style="background-color:#00A300;color:#fff;">DATE</th>
			<th style="background-color:#00A300;color:#fff;">ATTENDANCE</th>
			<th style="background-color:#00A300;color:#fff;">KG TBS</th>
			<th style="background-color:#00A300;color:#fff;">KG BRONDOLAN</th>
			<th style="background-color:#00A300;color:#fff;">KG PUPUK</th>
			<th style="background-color:#00A300;color:#fff;">KG LAIN - LAIN</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL KG</th>		
			<th style="background-color:#00A300;color:#fff;">BASIS</th>		
			<th style="background-color:#00A300;color:#fff;">KG PREMI LEBIH BASIS</th>		
			<th style="background-color:#00A300;color:#fff;">TOTAL PREMI</th>		
			<th style="background-color:#00A300;color:#fff;">HK EFEKTIF</th>						
		</tr>   
		    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="13" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		rows = 1
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<% if rows = 1 then
					sqlrow = "  select count(operatorcode) rowsd  from ( " & strGrnInv & " ) where operatorcode = '" & rsGrnInv("operatorcode") & "'"
					' response.write sqlrow
					Set rsrow = Server.CreateObject("ADODB.Recordset")
					rsrow.Open sqlrow, connstr, adOpenStatic
					if not rsrow.EOF then
						temprows = CDbl(rsrow("rowsd"))
					else
						temprows = 0
					end if
				end if 
			%>
			<%if rows = 1 then%>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsGrnInv("operatorcode")%></td>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsGrnInv("operatorname")%></td>
			<%end if%>
			<td align="center"><%=rsGrnInv("transdate")%></td>
			<td align="center"><%=rsGrnInv("attendance")%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_tbs"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_brondolan"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_pupuk"))),2)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_lain_lain"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("total_kg"))),2)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("basis"))),2)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("total_premi"))),2)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("hk_efektif"))),2)%></td>
		</tr>
<%	if rows = temprows then
			nkg_tbs = nkg_tbs + cdbl(to_nol(rsGrnInv("kg_tbs")))
			nkg_brondolan = nkg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))	
			nkg_pupuk = nkg_pupuk + cdbl(to_nol(rsGrnInv("kg_pupuk")))
			nkg_lain_lain = nkg_lain_lain + cdbl(to_nol(rsGrnInv("kg_lain_lain")))
			ntotal_kg = ntotal_kg + cdbl(to_nol(rsGrnInv("total_kg")))	 
			nbasis = nbasis + cdbl(to_nol(rsGrnInv("basis")))
			nkg_premi_lebih_basis = nkg_premi_lebih_basis + cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis")))
			ntotal_premi = ntotal_premi + cdbl(to_nol(rsGrnInv("total_premi")))
			nhk_efektif = nhk_efektif + cdbl(to_nol(rsGrnInv("hk_efektif"))) 
			rows = 1
			
	%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="4"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_tbs)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_brondolan)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_pupuk)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_lain_lain)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(ntotal_kg)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nbasis)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nkg_premi_lebih_basis)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(ntotal_premi)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(nhk_efektif)),2)%></font></td>   
		</tr>
		<tr>
			<td align="center" colspan="<%= 13+ nSpanTotal %>" ></td>
		</tr>
		<%
			nkg_tbs = cdbl("0")
			nkg_brondolan = cdbl("0")
			nkg_pupuk = cdbl("0")
			nkg_lain_lain = cdbl("0")
			ntotal_kg = cdbl("0") 
			nbasis = cdbl("0")
			nkg_premi_lebih_basis = cdbl("0") 
			ntotal_premi = cdbl("0") 
			nhk_efektif = cdbl("0")  
		else
		' cari sub total
			nkg_tbs = nkg_tbs + cdbl(to_nol(rsGrnInv("kg_tbs")))
			nkg_brondolan = nkg_brondolan + cdbl(to_nol(rsGrnInv("kg_brondolan")))	
			nkg_pupuk = nkg_pupuk + cdbl(to_nol(rsGrnInv("kg_pupuk")))
			nkg_lain_lain = nkg_lain_lain + cdbl(to_nol(rsGrnInv("kg_lain_lain")))
			ntotal_kg = ntotal_kg + cdbl(to_nol(rsGrnInv("total_kg")))
			nbasis = nbasis + cdbl(to_nol(rsGrnInv("basis")))
			nkg_premi_lebih_basis = nkg_premi_lebih_basis + cdbl(to_nol(rsGrnInv("kg_premi_lebih_basis")))
			ntotal_premi = ntotal_premi + cdbl(to_nol(rsGrnInv("total_premi")))
			nhk_efektif = nhk_efektif + cdbl(to_nol(rsGrnInv("hk_efektif")))
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_kg_tbs = cdbl(gt_kg_tbs) + cdbl(rsGrnInv("kg_tbs"))
				gt_kg_brondolan = cdbl(gt_kg_brondolan) + cdbl(rsGrnInv("kg_brondolan"))
				gt_kg_pupuk = cdbl(gt_kg_pupuk) + cdbl(rsGrnInv("kg_pupuk"))
				gt_kg_lain_lain = cdbl(gt_kg_lain_lain) + cdbl(rsGrnInv("kg_lain_lain"))
				gt_total_kg = cdbl(gt_total_kg) + cdbl(rsGrnInv("total_kg")) 
				gt_basis = cdbl(gt_basis) + cdbl(rsGrnInv("basis"))
				gt_kg_premi_lebih_basis = cdbl(gt_kg_premi_lebih_basis) + cdbl(rsGrnInv("kg_premi_lebih_basis"))
				gt_total_premi = cdbl(gt_total_premi) + cdbl(rsGrnInv("total_premi")) 
				gt_hk_efektif = cdbl(gt_hk_efektif) + cdbl(rsGrnInv("hk_efektif"))  

			rsGrnInv.MoveNext			
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%> 
		<tr bgcolor="#FFFF00">	 
			<td align="center" colspan="4"><b>GRAND TOTAL</b></td>
			<td align="right"><b><%=formatnumber(gt_kg_tbs,2)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_brondolan,2)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_pupuk,2)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_lain_lain,2)%></td>
			<td align="right"><b><%=formatnumber(gt_total_kg,2)%></td>
			<td align="right"><b><%=formatnumber(gt_basis,2)%></td>
			<td align="right"><b><%=formatnumber(gt_kg_premi_lebih_basis,2)%></td>
			<td align="right"><b><%=formatnumber(gt_total_premi,2)%></td>
			<td align="right"><b><%=formatnumber(gt_hk_efektif,2)%></td> 
			<!--<td align="right"></td>-->
		</tr>
	</tbody>
</table>
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

<!-- 251 -254 -->
<!-- 	"                                                              AVG ( mor.basis)" & _
		"                                                              OVER ( PARTITION BY    kg_tbs_brd.fcba" & _
		"                                                                              || kg_tbs_brd.receptiondate" & _
		"                                                                              || kg_tbs_brd.operator)" & _
		"                                                                 AS basis," & _ -->