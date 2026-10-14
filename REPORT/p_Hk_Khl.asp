<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade,v_view
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsKJ, strKJ, status_payroll
Dim NO, cClass, nday

v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtGangCode = Request("txtGangCode")
v_txtGrade = Request("txtGrade")
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
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 1, mandays, 0)) day_1" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 2, mandays, 0)) day_2" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 3, mandays, 0)) day_3" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 4, mandays, 0)) day_4" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 5, mandays, 0)) day_5" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 6, mandays, 0)) day_6" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 7, mandays, 0)) day_7" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 8, mandays, 0)) day_8" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 9, mandays, 0)) day_9" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 10, mandays, 0)) day_10" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 11, mandays, 0)) day_11" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 12, mandays, 0)) day_12" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 13, mandays, 0)) day_13" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 14, mandays, 0)) day_14" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 15, mandays, 0)) day_15" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 16, mandays, 0)) day_16" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 17, mandays, 0)) day_17" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 18, mandays, 0)) day_18" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 19, mandays, 0)) day_19" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 20, mandays, 0)) day_20" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 21, mandays, 0)) day_21" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 22, mandays, 0)) day_22" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 23, mandays, 0)) day_23" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 24, mandays, 0)) day_24" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 25, mandays, 0)) day_25" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 26, mandays, 0)) day_26" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 27, mandays, 0)) day_27" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 28, mandays, 0)) day_28" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 29, mandays, 0)) day_29" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 30, mandays, 0)) day_30" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 31, mandays, 0)) day_31" & _
					' "        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'yyyy')), " & v_txtYear & ", mandays" & _
					' "                     , TO_NUMBER(TO_CHAR(fddate, 'mm')), " & v_txtMonth & ", mandays" & _
					' "                     , 0)) total" & _
					' " FROM (" & _
					' "     SELECT ag.fcba, ag.fddate, ag.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode, "
					
					' IF TRIM(v_view) = "HK" THEN
					' strEmployee	= strEmployee +  "     SUM(ag.mandays) mandays" & _
												 ' "     FROM attendance_gad_temp ag " 
					' END IF

					' IF TRIM(v_view) = "BORONGAN" THEN
					' strEmployee	= strEmployee +  "     SUM(ag.unit * ag.rate) mandays" & _
												 ' "     FROM attendance_gad_temp ag " 
					' END IF

					' IF TRIM(v_view) = "PREMI" THEN
					' strEmployee	= strEmployee +  "     SUM(ed.amount) mandays" & _
												 ' "     FROM attendance_gad_temp ag " & _
												 ' "     INNER JOIN attendance_empallded ed on (ag.fcba=ed.fcba and ag.linenokey=ed.linenokey) "
					' END IF					
					' strEmployee	= strEmployee +	"     INNER JOIN employee_hist_tab em" & _
												' "     ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode AND TO_NUMBER(TO_CHAR(ag.fddate, 'yyyy')) = em.year AND TO_NUMBER(TO_CHAR(ag.fddate, 'mm')) = em.month)" & _
												' "     WHERE ag.fcba = '" & Session ("FCBA") & "'" & _
												' "           AND TO_CHAR(ag.fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" 
												
					' IF TRIM(v_view) = "HK" THEN
					' strEmployee	= strEmployee + "           AND ag.attendance IN ('KJ', 'DL','KK')" 
					' END IF

					' IF TRIM(v_view) = "HK" THEN
					' strEmployee	= strEmployee + "           AND ag.attendance IN ('KB')" 
					' END IF

					' strEmployee	= strEmployee +	"           AND em.presentgrade = 'KHL'"
												
												
						' if LEN(TRIM(v_txtGangCode)) > 0 then
							' strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
						' end if
						' strEmployee	= strEmployee & "     GROUP BY ag.fcba, ag.fddate, ag.employeecode, em.fcname, em.presentgrade, em.gangcode" & _
													' " )" & _
													' " GROUP BY fcba, employeecode, employeename, status, gangcode" & _
													' " ORDER BY employeecode ASC"
' ELSE
IF TRIM(v_view) <> "ABSEN" THEN
	strEmployee	=	" SELECT fcba, employeecode, employeename, status, gangcode" & _
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
					"        , SUM(DECODE(TO_NUMBER(TO_CHAR(fddate, 'yyyy')), " & v_txtYear & ", mandays" & _
					"                     , TO_NUMBER(TO_CHAR(fddate, 'mm')), " & v_txtMonth & ", mandays" & _
					"                     , 0)) total" & _
					" FROM (" & _
					"     SELECT ag.fcba, ag.fddate, ag.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode, " 

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

					strEmployee	= strEmployee +	"     INNER JOIN employee em" & _
												"     ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode)" & _
												"     WHERE ag.fcba = '" & Session ("FCBA") & "'" & _
												"           AND TO_CHAR(ag.fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" 

					IF TRIM(v_view) = "HK" THEN
					strEmployee	= strEmployee + "           AND ag.attendance IN ('KJ','DL','KK','TG')" 
					END IF

					IF TRIM(v_view) = "BORONGAN" THEN
					strEmployee	= strEmployee + "           AND ag.attendance IN ('KB')" 
					END IF

					strEmployee	= strEmployee +	"           AND em.presentgrade = 'KHL'"
	

					if LEN(TRIM(v_txtGangCode)) > 0 then
						strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
					end if
					strEmployee	= strEmployee & "     GROUP BY ag.fcba, ag.fddate, ag.employeecode, em.fcname, em.presentgrade, em.gangcode" 
					
					strEmployee	= strEmployee & "   UNION ALL " & _ 
					                            "   SELECT a.fcba, b.bkmdate fddate, a.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode, " 

					IF TRIM(v_view) = "HK" THEN
					strEmployee	= strEmployee &  "     SUM(a.mandays) mandays" & _
												 "     FROM BKMUPAHDTL a INNER JOIN BKMHDR b ON (a.bkmno=b.bkmno and a.fcba=b.fcba and b.postingflg='1'  )  " 
												 
					END IF

					IF TRIM(v_view) = "BORONGAN" THEN
					strEmployee	= strEmployee &  "     SUM(a.unit * a.rate) mandays" & _
												 "     FROM BKMUPAHDTL a INNER JOIN BKMHDR b ON (a.bkmno=b.bkmno and a.fcba=b.fcba and b.postingflg='1' )  " 
					END IF

					IF TRIM(v_view) = "PREMI" THEN
					strEmployee	= strEmployee &  "     SUM(a.allowance_amt) mandays" & _
												 "     FROM BKMUPAHDTL a INNER JOIN BKMHDR b ON (a.bkmno=b.bkmno and a.fcba=b.fcba and b.postingflg='1' )  " 
					END IF

					strEmployee	= strEmployee &	"     INNER JOIN employee em" & _
												"     ON (a.fcba = em.fcba AND a.employeecode = em.fccode)" & _
												"     WHERE a.fcba = '" & Session ("FCBA") & "'" & _
												"           AND TO_CHAR(b.bkmdate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" 

					IF TRIM(v_view) = "HK" THEN
					strEmployee	= strEmployee & "           AND a.attendance IN ('KJ','DL','KK','TG')" 
					END IF

					IF TRIM(v_view) = "BORONGAN" THEN
					strEmployee	= strEmployee & "           AND a.attendance IN ('KB')" 
					END IF

					strEmployee	= strEmployee &	"           AND em.presentgrade = 'KHL'"
	

					if LEN(TRIM(v_txtGangCode)) > 0 then
						strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
					end if
					strEmployee	= strEmployee & "     GROUP BY a.fcba, b.bkmdate, a.employeecode, em.fcname, em.presentgrade, em.gangcode" & _
												
												" )" & _
												" GROUP BY fcba, employeecode, employeename, status, gangcode" & _
												" ORDER BY employeecode ASC"
ELSE 
strEmployee	=	" SELECT fcba, employeecode, employeename, status, gangcode" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 1, attendance, '-')) day_1" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 2, attendance, '-')) day_2" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 3, attendance, '-')) day_3" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 4, attendance, '-')) day_4" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 5, attendance, '-')) day_5" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 6, attendance, '-')) day_6" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 7, attendance, '-')) day_7" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 8, attendance, '-')) day_8" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 9, attendance, '-')) day_9" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 10, attendance, '-')) day_10" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 11, attendance, '-')) day_11" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 12, attendance, '-')) day_12" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 13, attendance, '-')) day_13" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 14, attendance, '-')) day_14" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 15, attendance, '-')) day_15" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 16, attendance, '-')) day_16" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 17, attendance, '-')) day_17" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 18, attendance, '-')) day_18" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 19, attendance, '-')) day_19" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 20, attendance, '-')) day_20" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 21, attendance, '-')) day_21" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 22, attendance, '-')) day_22" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 23, attendance, '-')) day_23" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 24, attendance, '-')) day_24" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 25, attendance, '-')) day_25" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 26, attendance, '-')) day_26" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 27, attendance, '-')) day_27" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 28, attendance, '-')) day_28" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 29, attendance, '-')) day_29" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 30, attendance, '-')) day_30" & _
					"        , MAX(DECODE(TO_NUMBER(TO_CHAR(fddate, 'dd')), 31, attendance, '-')) day_31" & _
					"        , COUNT(DECODE(TO_NUMBER(TO_CHAR(fddate, 'yyyy')), " & v_txtYear & ", attendance" & _
					"                     , TO_NUMBER(TO_CHAR(fddate, 'mm')), " & v_txtMonth & ", attendance" & _
					"                     , NVL(attendance,'-'), 'KJ', attendance" & _
					"                     , 0)) total" & _
					" FROM (" & _
					"     SELECT DISTINCT ag.fcba, ag.fddate, ag.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode, ag.attendance" & _
					"     FROM attendance_gad_temp ag" & _
					"     INNER JOIN employee em" & _
					"     ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode)" & _
					"     WHERE ag.fcba = '" & Session ("FCBA") & "'" & _
					"           AND TO_CHAR(ag.fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
					"           AND ag.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" & _
					"           AND em.presentgrade IN ('KHL')"
	if LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
	end if
	strEmployee	= strEmployee & " UNION ALL" & _    
					"     SELECT DISTINCT a.fcba, b.bkmdate fddate, a.employeecode, em.fcname employeename, em.presentgrade status, em.gangcode gangcode, a.attendance" & _
					"     FROM bkmupahdtl a INNER JOIN bkmhdr b ON (a.bkmno=b.bkmno and a.fcba=b.fcba) " & _
					"     INNER JOIN employee em" & _
					"     ON (a.fcba = em.fcba AND a.employeecode = em.fccode)" & _
					"     WHERE a.fcba = '" & Session ("FCBA") & "'" & _
					"           AND TO_CHAR(b.bkmdate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
					"           AND a.attendance IN (SELECT fccode FROM parameter WHERE fcba = '" & Session ("FCBA") & "' AND parcode = 'E0102' AND upper(param_7) = upper('TRUE'))" & _
					"           AND em.presentgrade IN ('KHL')"
	if LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & " AND em.gangcode = '" & v_txtGangCode & "'"
	end if
	
	
	strEmployee	= strEmployee & " )" & _
								" GROUP BY fcba, employeecode, employeename, status, gangcode" & _
								" ORDER BY employeecode ASC"

END IF
'END IF								
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - HK/Premi/Borongan KHL</title>
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
		
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="37" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;"> 
		<h2>LAPORAN HARI KERJA KHL</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
       </th>
  </tr>
  
	<tr>
		<th>No</th>
		<th>No Pekerja</th>
		<th>Nama Pekerja</th>
		<th>Status</th>
		<th>Gang</th>
<%for nday = 1 to 31%>
<th><%=nday%></th>
<%next%>
		<th>Total</th>
	</tr>
    </thead>
    <tbody>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
'response.write strEmployee
rsEmployee.Open strEmployee, connstr, adOpenStatic

Dim nday_1
Dim nday_2
Dim nday_3
Dim nday_4
Dim nday_5
Dim nday_6
Dim nday_7
Dim nday_8
Dim nday_9
Dim nday_10
Dim nday_11
Dim nday_12
Dim nday_13
Dim nday_14
Dim nday_15
Dim nday_16
Dim nday_17
Dim nday_18
Dim nday_19
Dim nday_20
Dim nday_21
Dim nday_22
Dim nday_23
Dim nday_24
Dim nday_25
Dim nday_26
Dim nday_27
Dim nday_28
Dim nday_29
Dim nday_30
Dim nday_31
Dim nTOTAL

if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="37">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		
		nday_1 = 0
		nday_2 = 0
		nday_3 = 0
		nday_4 = 0
		nday_5 = 0
		nday_6 = 0
		nday_7 = 0
		nday_8 = 0
		nday_9 = 0
		nday_10 = 0
		nday_11 = 0
		nday_12 = 0
		nday_13 = 0
		nday_14 = 0
		nday_15 = 0
		nday_16 = 0
		nday_17 = 0
		nday_18 = 0
		nday_19 = 0
		nday_20 = 0
		nday_21 = 0
		nday_22 = 0
		nday_23 = 0
		nday_24 = 0
		nday_25 = 0
		nday_26 = 0
		nday_27 = 0
		nday_28 = 0
		nday_29 = 0
		nday_30 = 0
		nday_31 = 0
		nTOTAL = 0
		
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
				<td align="right"><%=rsEmployee("day_1")%></td>
				<td align="right"><%=rsEmployee("day_2")%></td>
				<td align="right"><%=rsEmployee("day_3")%></td>
				<td align="right"><%=rsEmployee("day_4")%></td>
				<td align="right"><%=rsEmployee("day_5")%></td>
				<td align="right"><%=rsEmployee("day_6")%></td>
				<td align="right"><%=rsEmployee("day_7")%></td>
				<td align="right"><%=rsEmployee("day_8")%></td>
				<td align="right"><%=rsEmployee("day_9")%></td>
				<td align="right"><%=rsEmployee("day_10")%></td>
				<td align="right"><%=rsEmployee("day_11")%></td>
				<td align="right"><%=rsEmployee("day_12")%></td>
				<td align="right"><%=rsEmployee("day_13")%></td>
				<td align="right"><%=rsEmployee("day_14")%></td>
				<td align="right"><%=rsEmployee("day_15")%></td>
				<td align="right"><%=rsEmployee("day_16")%></td>
				<td align="right"><%=rsEmployee("day_17")%></td>
				<td align="right"><%=rsEmployee("day_18")%></td>
				<td align="right"><%=rsEmployee("day_19")%></td>
				<td align="right"><%=rsEmployee("day_20")%></td>
				<td align="right"><%=rsEmployee("day_21")%></td>
				<td align="right"><%=rsEmployee("day_22")%></td>
				<td align="right"><%=rsEmployee("day_23")%></td>
				<td align="right"><%=rsEmployee("day_24")%></td>
				<td align="right"><%=rsEmployee("day_25")%></td>
				<td align="right"><%=rsEmployee("day_26")%></td>
				<td align="right"><%=rsEmployee("day_27")%></td>
				<td align="right"><%=rsEmployee("day_28")%></td>
				<td align="right"><%=rsEmployee("day_29")%></td>
				<td align="right"><%=rsEmployee("day_30")%></td>
				<td align="right"><%=rsEmployee("day_31")%></td>
				<td align="right"><%=rsEmployee("TOTAL")%></td>
<%
			NO = NO + 1

IF TRIM(v_view) <> "ABSEN" THEN 			
			nday_1 =  nday_1 + Cdbl(rsEmployee("day_1"))
			nday_2 =  nday_2 + Cdbl(rsEmployee("day_2"))
			nday_3 =  nday_3 + Cdbl(rsEmployee("day_3"))
			nday_4 =  nday_4 + Cdbl(rsEmployee("day_4"))
			nday_5 =  nday_5 + Cdbl(rsEmployee("day_5"))
			nday_6 =  nday_6 + Cdbl(rsEmployee("day_6"))
			nday_7 =  nday_7 + Cdbl(rsEmployee("day_7"))
			nday_8 =  nday_8 + Cdbl(rsEmployee("day_8"))
			nday_9 =  nday_9 + Cdbl(rsEmployee("day_9"))
			nday_10 =  nday_10 + Cdbl(rsEmployee("day_10"))
			nday_11 =  nday_11 + Cdbl(rsEmployee("day_11"))
			nday_12 =  nday_12 + Cdbl(rsEmployee("day_12"))
			nday_13 =  nday_13 + Cdbl(rsEmployee("day_13"))
			nday_14 =  nday_14 + Cdbl(rsEmployee("day_14"))
			nday_15 =  nday_15 + Cdbl(rsEmployee("day_15"))
			nday_16 =  nday_16 + Cdbl(rsEmployee("day_16"))
			nday_17 =  nday_17 + Cdbl(rsEmployee("day_17"))
			nday_18 =  nday_18 + Cdbl(rsEmployee("day_18"))
			nday_19 =  nday_19 + Cdbl(rsEmployee("day_19"))
			nday_20 =  nday_20 + Cdbl(rsEmployee("day_20"))
			nday_21 =  nday_21 + Cdbl(rsEmployee("day_21"))
			nday_22 =  nday_22 + Cdbl(rsEmployee("day_22"))
			nday_23 =  nday_23 + Cdbl(rsEmployee("day_23"))
			nday_24 =  nday_24 + Cdbl(rsEmployee("day_24"))
			nday_25 =  nday_25 + Cdbl(rsEmployee("day_25"))
			nday_26 =  nday_26 + Cdbl(rsEmployee("day_26"))
			nday_27 =  nday_27 + Cdbl(rsEmployee("day_27"))
			nday_28 =  nday_28 + Cdbl(rsEmployee("day_28"))
			nday_29 =  nday_29 + Cdbl(rsEmployee("day_29"))
			nday_30 =  nday_30 + Cdbl(rsEmployee("day_30"))
			nday_31 =  nday_31 + Cdbl(rsEmployee("day_31"))
			nTOTAL =  nTOTAL + Cdbl(rsEmployee("TOTAL"))
ELSE
			nday_1 =  nday_1 + InStr(rsEmployee("day_1"),"KJ")
			nday_2 =  nday_2 + InStr(rsEmployee("day_2"),"KJ")
			nday_3 =  nday_3 + InStr(rsEmployee("day_3"),"KJ")
			nday_4 =  nday_4 + InStr(rsEmployee("day_4"),"KJ")
			nday_5 =  nday_5 + InStr(rsEmployee("day_5"),"KJ")
			nday_6 =  nday_6 + InStr(rsEmployee("day_6"),"KJ")
			nday_7 =  nday_7 + InStr(rsEmployee("day_7"),"KJ") 
			nday_8 =  nday_8 + InStr(rsEmployee("day_8"),"KJ")
			nday_9 =  nday_9 + InStr(rsEmployee("day_9"),"KJ")
			nday_10 =  nday_10 + InStr(rsEmployee("day_10"),"KJ")
			nday_11 =  nday_11 + InStr(rsEmployee("day_11"),"KJ")
			nday_12 =  nday_12 + InStr(rsEmployee("day_12"),"KJ")
			nday_13 =  nday_13 + InStr(rsEmployee("day_13"),"KJ")
			nday_14 =  nday_14 + InStr(rsEmployee("day_14"),"KJ")
			nday_15 =  nday_15 + InStr(rsEmployee("day_15"),"KJ")
			nday_16 =  nday_16 + InStr(rsEmployee("day_16"),"KJ")
			nday_17 =  nday_17 + InStr(rsEmployee("day_17"),"KJ")
			nday_18 =  nday_18 + InStr(rsEmployee("day_18"),"KJ")
			nday_19 =  nday_19 + InStr(rsEmployee("day_19"),"KJ")
			nday_20 =  nday_20 + InStr(rsEmployee("day_20"),"KJ")
			nday_21 =  nday_21 + InStr(rsEmployee("day_21"),"KJ")
			nday_22 =  nday_22 + InStr(rsEmployee("day_22"),"KJ")
			nday_23 =  nday_23 + InStr(rsEmployee("day_23"),"KJ")
			nday_24 =  nday_24 + InStr(rsEmployee("day_24"),"KJ")
			nday_25 =  nday_25 + InStr(rsEmployee("day_25"),"KJ")
			nday_26 =  nday_26 + InStr(rsEmployee("day_26"),"KJ")
			nday_27 =  nday_27 + InStr(rsEmployee("day_27"),"KJ")
			nday_28 =  nday_28 + InStr(rsEmployee("day_28"),"KJ")
			nday_29 =  nday_29 + InStr(rsEmployee("day_29"),"KJ")
			nday_30 =  nday_30 + InStr(rsEmployee("day_30"),"KJ")
			nday_31 =  nday_31 + InStr(rsEmployee("day_31"),"KJ")
			nTOTAL =  nTOTAL + Cdbl(rsEmployee("TOTAL"))
END IF			
			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>

	<tr bgcolor="#F4F4A0">
		<td colspan="5" align="center"><strong>Total</strong></td>
		<td align="right"><strong><%=nday_1%></strong></td>
		<td align="right"><strong><%=nday_2%></strong></td>
		<td align="right"><strong><%=nday_3%></strong></td>
		<td align="right"><strong><%=nday_4%></strong></td>
		<td align="right"><strong><%=nday_5%></strong></td>
		<td align="right"><strong><%=nday_6%></strong></td>
		<td align="right"><strong><%=nday_7%></strong></td>
		<td align="right"><strong><%=nday_8%></strong></td>
		<td align="right"><strong><%=nday_9%></strong></td>
		<td align="right"><strong><%=nday_10%></strong></td>
		<td align="right"><strong><%=nday_11%></strong></td>
		<td align="right"><strong><%=nday_12%></strong></td>
		<td align="right"><strong><%=nday_13%></strong></td>
		<td align="right"><strong><%=nday_14%></strong></td>
		<td align="right"><strong><%=nday_15%></strong></td>
		<td align="right"><strong><%=nday_16%></strong></td>
		<td align="right"><strong><%=nday_17%></strong></td>
		<td align="right"><strong><%=nday_18%></strong></td>
		<td align="right"><strong><%=nday_19%></strong></td>
		<td align="right"><strong><%=nday_20%></strong></td>
		<td align="right"><strong><%=nday_21%></strong></td>
		<td align="right"><strong><%=nday_22%></strong></td>
		<td align="right"><strong><%=nday_23%></strong></td>
		<td align="right"><strong><%=nday_24%></strong></td>
		<td align="right"><strong><%=nday_25%></strong></td>
		<td align="right"><strong><%=nday_26%></strong></td>
		<td align="right"><strong><%=nday_27%></strong></td>
		<td align="right"><strong><%=nday_28%></strong></td>
		<td align="right"><strong><%=nday_29%></strong></td>
		<td align="right"><strong><%=nday_30%></strong></td>
		<td align="right"><strong><%=nday_31%></strong></td>
		<td align="right"><strong><%=nTOTAL%></strong></td>
	</tr>
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
