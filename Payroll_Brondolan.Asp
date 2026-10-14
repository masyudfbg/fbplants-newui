
<%
	'Amandemen Premi Brondolan Hari Libur
	'1. Code Allowance Panen = A26
	'   Rate1 = Rupiah Per Brondolan
	
	' Cari Type Calculate (M/D)
			qBA="select TYPE_1 from ALLOWDED where FCCODE='A26' and fcBA='" & Session("fcBA") & "'"
			Set rsBA = Server.CreateObject("ADODB.Recordset")
			'response.Write(qba)
			response.Flush()
			rsBA.Open qBA, connstr,3,3
			if not rsBA.eof then
			   calcType= rsBA("TYPE_1")
			else
			   calcType= "D"
			   Response.Write  "<font face=Verdana size=1 color=red>" & "TYPE 1 di allowded A26 belum di setting"
			   err.number=10001
			end if
			rsBA.close
			set rsBA = Nothing
			set qBA = Nothing	
	' End Cari Type Calculate (M/D)
	if calcType = "M" then
	
		qReceive = 	" SELECT FCBA," & _
					" LAST_DAY(TO_DATE(YEAR||MONTH,'yyyymm'))FDDATE," & _
					" FCCODE," & _
					" AMOUNT" & _
					" FROM" & _
					" (" & _
					" select" & _ 
					" ag.FCBA, TO_CHAR(ag.FDDATE,'YYYY') YEAR, TO_CHAR(ag.FDDATE,'mm')MONTH," & _
					" ag.EMPLOYEECODE FCCODE, ff.DIVISION, " & _
					" SUM(nvl(ag.BUCKET,0)) * MAX(RATE.RATE_BRONDOLAN) * MAX(BBR.BBR) AMOUNT" & _
					" from attendance_gad ag" & _
					" inner join field ff on ff.fccode = ag.LOCATIONCODE and ff.FCBA =  ag.FCBA and ff.ACTIVATION = 'Y'" & _
					" inner join (" & _
					"     select FCBA, DIVISION, MAX(RATE_BRONDOLAN) RATE_BRONDOLAN" & _
					"     FROM MULTILEVEL_HARVESTING_RATE " & _
					"     where " & _
					"     fcba='"& session("fcba") &"'" & _
					"     GROUP BY " & _
					"     FCBA, DIVISION" & _
					" ) RATE ON RATE.fcba = ag.fcba and ff.DIVISION = RATE.DIVISION" & _
					" inner join (" & _
					"     SELECT" & _
					"     FCBA," & _
					"     TO_CHAR(HARVESTDATE,'YYYY') YEARSX," & _
					"     TO_CHAR(HARVESTDATE,'mm') MONTHX," & _
					"     DIVISION," & _
					"     SUM(KGPERBLOK)/ SUM(BUCKET) BBR" & _
					"     FROM" & _
					"     (" & _
					"         SELECT " & _
					"         spb.FCBA," & _
					"         spb.HARVESTDATE," & _
					"         spb.SPBNO," & _
					"         spb.FIELDCODE," & _
					"         ff.DIVISION," & _
					"         spb.BUCKET," & _
					"         wh.total_bucket," & _
					"         wh.Kg_Brondolan," & _
					"         spb.BUCKET / wh.total_bucket * wh.Kg_Brondolan KGPERBLOK" & _
					"         FROM " & _
					"         wb_harvesting wh" & _
					"         inner join harvestingspb spb on wh.SPBNO = spb.SPBNO and wh.FCBA_ESTATE =  spb.FCBA and spb.FCBA ='"& session("fcba") &"'" & _
					"         inner join field ff on ff.fccode = spb.FIELDCODE and ff.FCBA =  spb.FCBA and ff.ACTIVATION = 'Y'" & _
					"         WHERE " & _
					"         fcba_estate = '"& session("fcba") &"'" & _
					"         and TO_NUMBER(TO_CHAR(spb.HARVESTDATE,'mm')) = " & Month(Session("v_Employee")) & _
					"         and TO_NUMBER(TO_CHAR(spb.HARVESTDATE,'yyyy')) = " & Year(Session("v_Employee")) & _
					"         and wh.TBSTYPE = 'TBSB'" & _
					"     )" & _
					"     GROUP BY" & _
					"     FCBA," & _
					"     TO_CHAR(HARVESTDATE,'YYYY')," & _
					"     TO_CHAR(HARVESTDATE,'mm')," & _
					"     DIVISION" & _
					" ) BBR ON BBR.fcba = ag.fcba and ff.DIVISION = BBR.DIVISION " & _
					"         and TO_CHAR(ag.FDDATE,'yyyy') = BBR.YEARSX and TO_CHAR(ag.FDDATE,'mm') = BBR.MONTHX" & _
					" where " & _
					" ag.fcba='"& session("fcba") &"'" & _
					" AND TO_NUMBER(TO_CHAR(ag.FDDATE,'mm')) = " & Month(Session("v_Employee")) & _
					" AND TO_NUMBER(TO_CHAR(ag.FDDATE,'yyyy')) = " & Year(Session("v_Employee")) & _
					" AND nvl(ag.BUCKET,0) > 0 " & _
					" AND ag.ATTENDANCE IN ('KJ','WH','WS','KK')" & _
					" GROUP BY" & _
					" ag.FCBA, TO_CHAR(ag.FDDATE,'YYYY'), TO_CHAR(ag.FDDATE,'mm'), ag.EMPLOYEECODE, ff.DIVISION" & _
					" )"
	else
	qReceive =		" SELECT emp.fcba," & _
					"        emp.FCCODE," & _ 
					"        mwn.harvestdate FDDATE," & _ 
					"        ROUND ( (att_gad.bucket * MWN.BBR * mhr.rate_brondolan), 2) AS AMOUNT " & _
					"   FROM attendance_gad att_gad" & _
					"        INNER JOIN Employee Emp" & _
					"           ON     att_gad.EMPLOYEECODE = emp.FCCODE" & _
					"              AND att_gad.FCBA = emp.FCBA" & _
					"              AND emp.DATETERMINATE IS NULL" & _
					"        INNER JOIN" & _
					"        (SELECT FCBA_ESTATE," & _
					"                harvestdate," & _
					"                DIVISION," & _
					"                fieldcode," & _
					"                total_bunch AS bunch," & _
					"                kg_per_block," & _
					"                Kg_Brondolan," & _
					"                bucket_per_block," & _
					"                total_bucket," & _
					"                BBR" & _
					"           FROM (  SELECT FCBA_ESTATE," & _
					"                          harvestdate," & _
					"                          DIVISION," & _
					"                          fieldcode," & _
					"                          SUM (bunch) AS bunch," & _
					"                          SUM (total_bunch) AS total_bunch," & _
					"                          SUM (Bunch_Mill_Weight_Netto)" & _
					"                             AS Bunch_Mill_Weight_Netto," & _
					"                          SUM (kg_per_block) AS kg_per_block," & _
					"                          SUM (Kg_Brondolan) AS Kg_Brondolan," & _
					"                          SUM (bucket_per_block) AS bucket_per_block," & _
					"                          SUM (total_bucket) AS total_bucket," & _
					"                          NVL (" & _
					"                             DECODE (SUM (total_bucket)," & _
					"                                     0, 0," & _
					"                                     SUM (Kg_Brondolan) / SUM (total_bucket))," & _
					"                             0)" & _
					"                             AS BBR" & _
					"                     FROM (SELECT iplasprod.harvestingspb.fcba AS FCBA_ESTATE," & _
					"                                  iplasprod.harvestingspb.spbno," & _
					"                                  iplasprod.harvestingspb.harvestdate," & _
					"                                  iplasprod.harvestingspb.fieldcode," & _
					"                                  FIELD.DIVISION," & _
					"                                  NVL (iplasprod.harvestingspb.bunch, 0)" & _
					"                                     AS bunch," & _
					"                                  NVL (total_bunch.total_bunch, 0)" & _
					"                                     AS total_bunch," & _
					"                                  NVL (wb_harv.Bunch_Mill_Weight_Netto, 0)" & _
					"                                     AS Bunch_Mill_Weight_Netto," & _
					"                                  DECODE (" & _
					"                                     total_bunch.total_bunch," & _
					"                                     0, 0," & _
					"                                     (  (  iplasprod.harvestingspb.bunch" & _
					"                                         / total_bunch.total_bunch)" & _
					"                                      * wb_harv.Bunch_Mill_Weight_Netto))" & _
					"                                     AS kg_per_block," & _
					"                                  Kg_Brondolan," & _
					"                                  NVL (iplasprod.harvestingspb.bucket, 0)" & _
					"                                     AS bucket_per_block," & _
					"                                  total_bucket" & _
					"                             FROM iplasprod.harvestingspb" & _
					"                                  INNER JOIN FIELD" & _
					"                                     ON     iplasprod.harvestingspb.fieldcode =" & _
					"                                               field.fccode" & _
					"                                        AND iplasprod.harvestingspb.fcba =" & _
					"                                               field.fcba" & _
					"                                        AND field.ACTIVATION = 'Y'" & _
					"                                  INNER JOIN" & _
					"                                  (  SELECT fcba," & _
					"                                            spbno," & _
					"                                            SUM (bunch) AS total_bunch" & _
					"                                       FROM iplasprod.harvestingspb" & _
					"                                   GROUP BY fcba, spbno) total_bunch" & _
					"                                     ON (    iplasprod.harvestingspb.fcba =" & _
					"                                                total_bunch.fcba" & _
					"                                         AND iplasprod.harvestingspb.spbno =" & _
					"                                                total_bunch.spbno)" & _
					"                                  INNER JOIN" & _
					"                                  (  SELECT iplasprod.wb_harvesting.fcba_estate," & _
					"                                            iplasprod.wb_harvesting.spbno," & _
					"                                            SUM (" & _
					"                                               iplasprod.wb_harvesting.bunch_mill_weight_netto)" & _
					"                                               AS Bunch_Mill_Weight_Netto," & _
					"                                            SUM (" & _
					"                                               iplasprod.wb_harvesting.Kg_Brondolan)" & _
					"                                               Kg_Brondolan," & _
					"                                            SUM (" & _
					"                                               iplasprod.wb_harvesting.total_bucket)" & _
					"                                               total_bucket" & _
					"                                       FROM iplasprod.wb_harvesting" & _
					"                                   GROUP BY iplasprod.wb_harvesting.fcba_estate," & _
					"                                            iplasprod.wb_harvesting.spbno)" & _
					"                                  wb_harv" & _
					"                                     ON (    iplasprod.harvestingspb.fcba =" & _
					"                                                wb_harv.fcba_estate" & _
					"                                         AND iplasprod.harvestingspb.spbno =" & _
					"                                                wb_harv.spbno))" & _
					"                 GROUP BY FCBA_ESTATE," & _
					"                          harvestdate," & _
					"                          DIVISION," & _
					"                          fieldcode)) MWN" & _
					"           ON     (    att_gad.fcba = MWN.FCBA_ESTATE" & _
					"                   AND att_gad.LOCATIONCODE = MWN.FIELDCODE)" & _
					"              AND TO_CHAR (att_gad.FDDATE, 'MM/DD/YYYY') =" & _
					"                     TO_CHAR (MWN.harvestdate, 'MM/DD/YYYY')" & _
					"        INNER JOIN FIELD ff" & _
					"           ON     att_gad.LOCATIONCODE = ff.FCCODE" & _
					"              AND att_gad.fcba = ff.fcba" & _
					"              AND ff.ACTIVATION = 'Y'" & _
					"        INNER JOIN (SELECT DISTINCT fcba," & _
					"                                    division," & _
					"                                    planting_year," & _
					"                                    rate_brondolan" & _
					"                      FROM MULTILEVEL_HARVESTING_RATE" & _
					"                     WHERE fcba = '"& session("fcba") &"') mhr" & _
					"           ON (    att_gad.fcba = mhr.fcba" & _
					"               AND ff.DIVISION = mhr.DIVISION" & _
					"               AND TO_CHAR (ff.PLANTINGDATE, 'yyyy') = mhr.PLANTING_YEAR)" & _
					"  WHERE     att_gad.Attendance IN ('KJ'," & _
					"                                   'WH'," & _
					"                                   'WS'," & _
					"                                   'KK')" & _
					"        AND att_gad.Jobcode IN ('112030101', '505030101')" & _
					"        AND att_gad.LOCATIONTYPE = 'FF'" & _
					"        AND att_gad.MANDAYS > 0" & _
					"        AND att_gad.BUCKET > 0" & _
					"        AND att_gad.rate = 0" & _
					"        AND att_gad.fcBA = '"& session("fcba") &"'" &_  
					" 		 AND TO_NUMBER(TO_CHAR(MWN.harvestdate,'mm')) = " & Month(Session("v_Employee")) & _
					" 		 AND TO_NUMBER(TO_CHAR(MWN.harvestdate,'yyyy')) = " & Year(Session("v_Employee")) & _
					" ORDER BY fcba," & _
					"          employeecode," & _
 					"         fddate"
	end if
	
	' response.Write " sql " & qReceive
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate Premi Kutip Brondolan ..." & "<br>" 
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof
         
		  nRpPremi=cdbl(rsEmpAllDed("AMOUNT"))
		  if nRpPremi > 0 then
		     maxLineNo = maxLineNo + 1
	         qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
			            "(" & maxLineNo & ",'" & rsEmpAllDed("FCCODE") & "','" & "A26" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremi & _
			            ",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			            Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
		     'response.Write "qAllowance : " & qAllowance
			 connstr.execute(qAllowance)
		  end if

	      rsEmpAllDed.MoveNext
	   loop
	End if
	rsEmpAllDed.Close
	set rsEmpAllDed = Nothing
%>