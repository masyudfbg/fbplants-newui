
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade, v_view
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsKJ, strKJ, status_payroll
Dim NO, cClass, nday

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtGangCode	=	Request("txtGangCode")
v_txtGrade		=	Request("txtGrade")
v_txtsectionname		=	Request("txtsectionname")
v_view = Request("RPT")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if "" & v_txtYear = "" then
   v_txtYear = Year(Session("v_Employee"))
end if
if "" & v_txtMonth = "" then
   v_txtMonth = Month(Session("v_Employee"))
end if

if "" & v_view = "" then
    v_view = "HK"
end if

strKJ = " SELECT NVL(payroll_executed, 'FALSE') detail" & _
        " FROM empalldedstatus" & _
		" WHERE fcba = '" & Session ("FCBA") & "'" & _
        "       AND year = " & v_txtYear & "" & _
        "       AND month = " & v_txtMonth & ""

Set rsKJ = Server.CreateObject("ADODB.Recordset")
rsKJ.Open strKJ, connstr, adOpenStatic

IF NOT rsKJ.EOF THEN
	status_payroll = rsKJ("detail")
ELSE
	status_payroll = "FALSE"
END IF
rsKJ.Close
Set rsKJ = Nothing
'response.write status_payroll
' IF status_payroll = "TRUE" THEN
	' strEmployee	=	" SELECT fcba, employeecode, employeename, status, gangcode" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 1, attendance, NULL)) day_1" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 2, attendance, NULL)) day_2" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 3, attendance, NULL)) day_3" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 4, attendance, NULL)) day_4" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 5, attendance, NULL)) day_5" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 6, attendance, NULL)) day_6" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 7, attendance, NULL)) day_7" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 8, attendance, NULL)) day_8" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 9, attendance, NULL)) day_9" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 10, attendance, NULL)) day_10" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 11, attendance, NULL)) day_11" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 12, attendance, NULL)) day_12" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 13, attendance, NULL)) day_13" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 14, attendance, NULL)) day_14" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 15, attendance, NULL)) day_15" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 16, attendance, NULL)) day_16" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 17, attendance, NULL)) day_17" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 18, attendance, NULL)) day_18" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 19, attendance, NULL)) day_19" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 20, attendance, NULL)) day_20" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 21, attendance, NULL)) day_21" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 22, attendance, NULL)) day_22" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 23, attendance, NULL)) day_23" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 24, attendance, NULL)) day_24" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 25, attendance, NULL)) day_25" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 26, attendance, NULL)) day_26" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 27, attendance, NULL)) day_27" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 28, attendance, NULL)) day_28" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 29, attendance, NULL)) day_29" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 30, attendance, NULL)) day_30" & _
					' "        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 31, attendance, NULL)) day_31" & _
					' " FROM (" & _
					' "     SELECT DISTINCT ag.fcba, ag.fddate, ag.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode, ag.attendance" & _
					' "     FROM attendance_gad_temp ag" & _
					' "     INNER JOIN employee_hist_tab em" & _
					' "     ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode AND TO_NUMBER(TO_CHAR(ag.fddate, 'yyyy')) = em.year AND TO_NUMBER(TO_CHAR(ag.fddate, 'mm')) = em.month)" & _
					' "     WHERE ag.fcba = '" & Session ("FCBA") & "'" & _
					' "           AND TO_CHAR(ag.fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
					' "           AND ag.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" & _
					' "           AND em.presentgrade IN ('KHT', 'KBT')"
	' if LEN(TRIM(v_txtGangCode)) > 0 then
		' strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
	' end if
	' if  LEN(TRIM(v_txtGrade)) > 0 then
		' strEmployee	=	strEmployee & " AND em.presentgrade = '" & v_txtGrade & "'"
	' end if
	' if  LEN(TRIM(v_txtsectionname)) > 0 then
		' strEmployee	=	strEmployee & " AND em.sectionname = '" & v_txtsectionname & "'"
	' end if
	' strEmployee	= strEmployee & " )" & _
								' " GROUP BY fcba, employeecode, employeename, status, gangcode" & _
								' " ORDER BY employeecode ASC"
' ELSE

IF TRIM(v_view) <> "ABSEN" THEN
	strEmployee	=	" SELECT fcba, employeecode, employeename, status, gangcode,position" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 1, mandays, 0)) day_1" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 2, mandays, 0)) day_2" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 3, mandays, 0)) day_3" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 4, mandays, 0)) day_4" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 5, mandays, 0)) day_5" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 6, mandays, 0)) day_6" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 7, mandays, 0)) day_7" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 8, mandays, 0)) day_8" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 9, mandays, 0)) day_9" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 10, mandays, 0)) day_10" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 11, mandays, 0)) day_11" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 12, mandays, 0)) day_12" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 13, mandays, 0)) day_13" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 14, mandays, 0)) day_14" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 15, mandays, 0)) day_15" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 16, mandays, 0)) day_16" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 17, mandays, 0)) day_17" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 18, mandays, 0)) day_18" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 19, mandays, 0)) day_19" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 20, mandays, 0)) day_20" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 21, mandays, 0)) day_21" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 22, mandays, 0)) day_22" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 23, mandays, 0)) day_23" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 24, mandays, 0)) day_24" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 25, mandays, 0)) day_25" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 26, mandays, 0)) day_26" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 27, mandays, 0)) day_27" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 28, mandays, 0)) day_28" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 29, mandays, 0)) day_29" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 30, mandays, 0)) day_30" & _
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 31, mandays, 0)) day_31" & _
					" FROM (" & _
					"     	SELECT ag.fcba, ag.fddate, ag.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode,ep.description position, ag.attendance, " 
					
					IF TRIM(v_view) = "HK" THEN
					strEmployee	= strEmployee +  "     SUM(ag.mandays) mandays" & _
												 "     FROM attendance_gad_temp ag " 
					END IF

					IF TRIM(v_view) = "BORONGAN" THEN
					strEmployee	= strEmployee +  "     SUM(ag.unit * ag.rate) mandays" & _
												 "     FROM attendance_gad_temp ag " 
					END IF

					IF TRIM(v_view) = "PREMI" THEN
					strEmployee	= strEmployee +  "     SUM(ed.amount) mandays" & _
												 "     FROM attendance_gad_temp ag " & _
												 "     INNER JOIN attendance_empallded ed on (ag.fcba=ed.fcba and ag.linenokey=ed.linenokey) "
					END IF
					strEmployee	= strEmployee + "     INNER JOIN employee em" & _
												"     ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode)" & _
												" 	   LEFT JOIN EMPPOSITION ep " & _
												" ON ep.position = em.POSITION AND ep.fcba = em.fcba" & _
												"     WHERE ag.fcba = '" & Session ("FCBA") & "'" & _
												"           AND TO_CHAR(ag.fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" 
					IF TRIM(v_view) = "HK" THEN
					strEmployee	= strEmployee + "           AND ag.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" 
					END IF

					IF TRIM(v_view) = "BORONGAN" THEN
					strEmployee	= strEmployee + "           AND ag.attendance IN ('KB')" 
					END IF	

					IF TRIM(v_view) = "PREMI" THEN
					strEmployee	= strEmployee + "           AND ag.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" 
					END IF
												
					strEmployee	= strEmployee +	"           AND em.presentgrade IN ('KHT', 'KBT')"
	
	
	if LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
	end if
	if  LEN(TRIM(v_txtGrade)) > 0 then
		strEmployee	=	strEmployee & " AND em.presentgrade = '" & v_txtGrade & "'"
	end if
	if  LEN(TRIM(v_txtsectionname)) > 0 then
		strEmployee	=	strEmployee & " AND em.sectionname = '" & v_txtsectionname & "'"
	end if
	strEmployee	= strEmployee & " GROUP BY ag.fcba, ag.fddate, ag.employeecode, em.fcname , em.presentgrade , em.gangcode,ep.description, ag.attendance "
	
	strEmployee	= strEmployee & "   UNION ALL " & _ 
			    "   SELECT a.fcba, b.bkmdate fddate, a.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode,ep.description,a.attendance," 

					IF TRIM(v_view) = "HK" THEN
					strEmployee	= strEmployee &  "     SUM(a.mandays) mandays" & _
												 "     FROM BKMUPAHDTL a INNER JOIN BKMHDR b ON (a.bkmno=b.bkmno and a.fcba=b.fcba and b.postingflg='1' )  " 
												 
					END IF

					IF TRIM(v_view) = "BORONGAN" THEN
					strEmployee	= strEmployee &  "     SUM(a.unit * a.rate) mandays" & _
												 "     FROM BKMUPAHDTL a INNER JOIN BKMHDR b ON (a.bkmno=b.bkmno and a.fcba=b.fcba and b.postingflg='1' )  " 
					END IF

					IF TRIM(v_view) = "PREMI" THEN
					strEmployee	= strEmployee &  "     SUM(a.allowance_amt) mandays" & _
												 "     FROM BKMUPAHDTL a INNER JOIN BKMHDR b ON (a.bkmno=b.bkmno and a.fcba=b.fcba and b.postingflg='1' )   " 
					END IF

					strEmployee	= strEmployee &	"     INNER JOIN employee em" & _
												"     ON (a.fcba = em.fcba AND a.employeecode = em.fccode)" & _
												" 	   LEFT JOIN EMPPOSITION ep " & _
												" ON ep.position = em.POSITION AND ep.fcba = em.fcba" & _

												"     WHERE a.fcba = '" & Session ("FCBA") & "'" & _
												"           AND TO_CHAR(b.bkmdate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" 

					IF TRIM(v_view) = "HK" THEN
					strEmployee	= strEmployee & "           AND a.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" 
					END IF

					IF TRIM(v_view) = "BORONGAN" THEN
					strEmployee	= strEmployee & "           AND a.attendance IN ('KB')" 
					END IF

                    IF TRIM(v_view) = "PREMI" THEN
					strEmployee	= strEmployee & "           AND a.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" 
					END IF
					
					strEmployee	= strEmployee &	"           AND em.presentgrade IN ('KHT', 'KBT')"
	

					if LEN(TRIM(v_txtGangCode)) > 0 then
						strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
					end if
					strEmployee	= strEmployee & "     GROUP BY a.fcba, b.bkmdate, a.employeecode, em.fcname, em.presentgrade, em.gangcode,ep.description, a.attendance" & _
												
	
	                            ")" & _
								" GROUP BY fcba, employeecode, employeename, status, gangcode,position" & _
								" ORDER BY employeecode ASC"

ELSE
strEmployee	=	" SELECT fcba, employeecode, employeename, status, gangcode, position" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 1, attendance, NULL)) day_1" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 2, attendance, NULL)) day_2" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 3, attendance, NULL)) day_3" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 4, attendance, NULL)) day_4" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 5, attendance, NULL)) day_5" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 6, attendance, NULL)) day_6" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 7, attendance, NULL)) day_7" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 8, attendance, NULL)) day_8" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 9, attendance, NULL)) day_9" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 10, attendance, NULL)) day_10" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 11, attendance, NULL)) day_11" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 12, attendance, NULL)) day_12" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 13, attendance, NULL)) day_13" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 14, attendance, NULL)) day_14" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 15, attendance, NULL)) day_15" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 16, attendance, NULL)) day_16" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 17, attendance, NULL)) day_17" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 18, attendance, NULL)) day_18" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 19, attendance, NULL)) day_19" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 20, attendance, NULL)) day_20" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 21, attendance, NULL)) day_21" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 22, attendance, NULL)) day_22" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 23, attendance, NULL)) day_23" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 24, attendance, NULL)) day_24" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 25, attendance, NULL)) day_25" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 26, attendance, NULL)) day_26" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 27, attendance, NULL)) day_27" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 28, attendance, NULL)) day_28" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 29, attendance, NULL)) day_29" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 30, attendance, NULL)) day_30" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 31, attendance, NULL)) day_31" & _
					" FROM (" & _
					"     SELECT DISTINCT ag.fcba, ag.fddate, ag.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode,ep.description position, ag.attendance" & _
					"     FROM attendance_gad_temp ag" & _
					"     INNER JOIN employee em" & _
					"     ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode)" & _
					" 	   LEFT JOIN EMPPOSITION ep " & _
               		" ON ep.position = em.POSITION AND ep.fcba = em.fcba" & _
					"     WHERE ag.fcba = '" & Session ("FCBA") & "'" & _
					"           AND TO_CHAR(ag.fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
					"           AND ag.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" & _
					"           AND em.presentgrade IN ('KHT', 'KBT')"
	if LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
	end if
	if  LEN(TRIM(v_txtGrade)) > 0 then
		strEmployee	=	strEmployee & " AND em.presentgrade = '" & v_txtGrade & "'"
	end if
	if  LEN(TRIM(v_txtsectionname)) > 0 then
		strEmployee	=	strEmployee & " AND em.sectionname = '" & v_txtsectionname & "'"
	end if
	strEmployee	= strEmployee & " UNION ALL" & _    
					"     SELECT DISTINCT a.fcba, b.bkmdate fddate, a.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode,ep.description, a.attendance" & _
					"     FROM bkmupahdtl a INNER JOIN bkmhdr b ON (a.bkmno=b.bkmno and a.fcba=b.fcba) " & _
					"     INNER JOIN employee em" & _
					"     ON (a.fcba = em.fcba AND a.employeecode = em.fccode)" & _
					" 	   LEFT JOIN EMPPOSITION ep " & _
               		" ON ep.position = em.POSITION AND ep.fcba = em.fcba" & _
					"     WHERE b.postingflg='1' and a.fcba = '" & Session ("FCBA") & "'" & _
					"           AND TO_CHAR(b.bkmdate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
					"           AND a.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" & _
					"           AND em.presentgrade IN ('KHT', 'KBT')"
	if LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
	end if
	if  LEN(TRIM(v_txtGrade)) > 0 then
		strEmployee	=	strEmployee & " AND em.presentgrade = '" & v_txtGrade & "'"
	end if
	if  LEN(TRIM(v_txtsectionname)) > 0 then
		strEmployee	=	strEmployee & " AND em.sectionname = '" & v_txtsectionname & "'"
	end if
	
	
	strEmployee	= strEmployee & " )" & _
								" GROUP BY fcba, employeecode, employeename, status, gangcode,position" & _
								" ORDER BY employeecode ASC"

END IF
	
''	 response.write strEmployee

'END IF
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Absensi KHT dan KBT</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
    <table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
		<td><font face="verdana" size="1">GANGCODE</font></td><td><select name="txtGangCode" ><option value=""></option>
		<%
		Dim rsGangCode, strGangCode
		strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
		
		Set rsGangCode = Server.CreateObject("ADODB.Recordset")
		rsGangCode.Open strGangCode, connstr, adOpenStatic 
		
		if not rsGangCode.eof then
		Do While Not rsGangCode.eof
		if rsGangCode("FCCODE") = v_txtGangCode then%>
		        <option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option> 
		     <%Else%>
		        <option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option>  
		    <%End if
			rsGangCode.movenext
		loop
		end if
		rsGangCode.Close
		set rsGangCode = Nothing
		%>
		</select></td>
		<td><font face="verdana" size="1">VIEW</font></td><td>
		<select name="RPT" title="VIEW">
					<option value="ABSEN" <% If "ABSEN"=v_view Then %> Selected <%End if%>>ABSEN</option> 
					<option value="HK" <% If "HK"=v_view Then %> Selected <%End if%>>HK</option> 
					<option value="PREMI" <% If "PREMI"=v_view Then %> Selected <%End if%>>PREMI</option> 
					<option value="BORONGAN" <% If "BORONGAN"=v_view Then %> Selected <%End if%>>BORONGAN</option> 
		</select>
		</td>
		<td><font face="Verdana" size="1">Section Name</font></td>
		<td>
			<select name="txtsectionname">
				<option value=""></option>
				<%
				Dim rssectionname, strsectionname
				strsectionname	=	"select distinct sectionname as fccode, sectionname_api.get_fcname(sectionname) fcname from employee where fcba= '"& session("fcba") &"' order by sectionname"
				
				Set rssectionname = Server.CreateObject("ADODB.Recordset")
				rssectionname.Open strsectionname, connstr, adOpenStatic 
				
				if not rssectionname.eof then
				Do While Not rssectionname.eof
				if rssectionname("fccode") = v_txtsectionname then%>
						<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
					 <%Else%>
						<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
					<%End if
					rssectionname.movenext
				loop
				end if
				rssectionname.Close
				set rssectionname = Nothing
				%>
			</select>
		</td>
		<td><font face="verdana" size="1">STATUS</font></td><td><select name="txtGrade" ><option value=""></option>
		<%
		Dim rsGrade, strGrade
		strGrade	=	" SELECT 'KHT' FCCODE, 'Karyawan Harian Tetap' FCNAME FROM DUAL" & _
						" UNION" & _
						" SELECT 'KBT' FCCODE, 'Karyawan Bulanan Tetap' FCNAME FROM DUAL"
		
		Set rsGrade = Server.CreateObject("ADODB.Recordset")
		rsGrade.Open strGrade, connstr, adOpenStatic 
		
		if not rsGrade.eof then
		Do While Not rsGrade.eof
		if rsGrade("FCCODE") = v_txtGrade then%>
		        <option value="<%=rsGrade("FCCODE")%>" Selected><%=rsGrade("FCCODE")%>-<%=rsGrade("FCNAME")%></option> 
		     <%Else%>
		        <option value="<%=rsGrade("FCCODE")%>"><%=rsGrade("FCCODE")%>-<%=rsGrade("FCNAME")%></option>  
		    <%End if
			rsGrade.movenext
		loop
		end if
		rsGrade.Close
		set rsGrade = Nothing
		%>
		</select></td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
            <h4>PT. <%=Session("COMPANY")%></h4>
            <h5><%=Session("BA")%></h5>
		</div>
        <div id="subheader" class="centered" style="width:50%; margin:auto;">
            <h2>ABSENSI KHT dan KBT</h2>
            <br/>
            <h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
        </div>
	</div>
       </th>
  </tr>
	<tr>
		<th>NO</th>
		<th>NO PEKERJA</th>
		<th>NAMA PEKERJA</th>
		<th>STATUS</th>
		<th>GANG</th>
		<th>JABATAN</th>
<%for nday = 1 to 31%>
<th><%=nday%></th>
<%next%>
	</tr>
 </thead>
 <tbody>

<%
' response.write strEmployee
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic
if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="37">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		while not rsEmployee.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="center"><%=rsEmployee("EMPLOYEECODE")%></td>
				<td align="left"><%=rsEmployee("EMPLOYEENAME")%></td>
				<td align="center"><%=rsEmployee("STATUS")%></td>
				<td align="center"><%=rsEmployee("GANGCODE")%></td>
				<td align="left"><%=rsEmployee("POSITION")%></td>
				<td align="center"><%=rsEmployee("day_1")%></td>
				<td align="center"><%=rsEmployee("day_2")%></td>
				<td align="center"><%=rsEmployee("day_3")%></td>
				<td align="center"><%=rsEmployee("day_4")%></td>
				<td align="center"><%=rsEmployee("day_5")%></td>
				<td align="center"><%=rsEmployee("day_6")%></td>
				<td align="center"><%=rsEmployee("day_7")%></td>
				<td align="center"><%=rsEmployee("day_8")%></td>
				<td align="center"><%=rsEmployee("day_9")%></td>
				<td align="center"><%=rsEmployee("day_10")%></td>
				<td align="center"><%=rsEmployee("day_11")%></td>
				<td align="center"><%=rsEmployee("day_12")%></td>
				<td align="center"><%=rsEmployee("day_13")%></td>
				<td align="center"><%=rsEmployee("day_14")%></td>
				<td align="center"><%=rsEmployee("day_15")%></td>
				<td align="center"><%=rsEmployee("day_16")%></td>
				<td align="center"><%=rsEmployee("day_17")%></td>
				<td align="center"><%=rsEmployee("day_18")%></td>
				<td align="center"><%=rsEmployee("day_19")%></td>
				<td align="center"><%=rsEmployee("day_20")%></td>
				<td align="center"><%=rsEmployee("day_21")%></td>
				<td align="center"><%=rsEmployee("day_22")%></td>
				<td align="center"><%=rsEmployee("day_23")%></td>
				<td align="center"><%=rsEmployee("day_24")%></td>
				<td align="center"><%=rsEmployee("day_25")%></td>
				<td align="center"><%=rsEmployee("day_26")%></td>
				<td align="center"><%=rsEmployee("day_27")%></td>
				<td align="center"><%=rsEmployee("day_28")%></td>
				<td align="center"><%=rsEmployee("day_29")%></td>
				<td align="center"><%=rsEmployee("day_30")%></td>
				<td align="center"><%=rsEmployee("day_31")%></td>
<%
			NO = NO + 1			
			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>
</tbody>
</table>

   <table width="90%" class="data">
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
