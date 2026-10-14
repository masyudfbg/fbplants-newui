<%
Option Explicit
response.buffer=True
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
	Const adOpenForwardOnly = 0
	Const adLockReadOnly = 1
	Const adCmdTableDirect = &H0200
	Const adUseClient = 3
	Dim TimeAwal,TimeAkhir,Hitung,Lama
	TimeAwal=Time()
%>
<div id="status_running">
	<img src="ICON/working.gif"/>
</div>
<br/><br/>
<font face="Verdana" size="1">
Please Wait...<br/><br/></font>
<%
Response.Write "Start ..." & Time() & "<br/>" 
response.Flush()

' ADD AKBAR 14 JANUARI 1013
Dim Bulan, Tahun
Dim rsPERIOD,NormalisasiARTMTBM, SelectARTMTBM, rsARTMTBM, InsertARTMTBM
Dim SelectPERIOD,NormalisasiARMill, SelectARMill, rsARMill, InsertARMill

SelectPERIOD = " SELECT TO_CHAR(LASTPERIODPOSTED,'MM') BULAN, TO_CHAR(LASTPERIODPOSTED,'YYYY')  TAHUN  FROM  PERIODCONTROL  " & _
               " WHERE FCBA= '" & Session("fcBA") & "' AND SYSTEM ='Costbook' AND " & _
      		   " LASTPERIODPOSTED = ( SELECT  MAX(LASTPERIODPOSTED)  FROM  PERIODCONTROL WHERE SYSTEM ='Costbook' AND fcba='" & Session("fcBA") & "'  ) " 
			   
Set rsPERIOD= Server.CreateObject("ADODB.Recordset")
rsPERIOD.CursorType = 3
rsPERIOD.CursorLocation = 3
rsPERIOD.Open SelectPERIOD, connstr

if not rsPERIOD.eof then
	Do While Not rsPERIOD.eof
       Bulan = rsPERIOD("BULAN")
       Tahun = rsPERIOD("TAHUN") 
	   rsPERIOD.MoveNext
	Loop
end if	
	Response.Write "Create Generate Account Report for Month = " & Bulan & " Year = " & Tahun & "<br/>" 
    response.Flush()   
' END ADD AKBAR 14 JANUARI 1013	   

'----BEGIN ACCOUNT REPORT TM TBM----

'Normalisasi Account Report TM TBM
NormalisasiARTMTBM =	" DELETE FROM SAC_ACCOUNTREPORT_TM_TBM" & _
						" WHERE" & _
						" PERIOD_MONTH = " & Bulan & " AND" & _
						" PERIOD_YEAR = " & Tahun & " AND" & _
						" BA ='" & Session("fcBA") & "'"
connstr.execute(NormalisasiARTMTBM)

'Collecting Data for Account Report TM TBM
SelectARTMTBM =	" select" & _
				" 	 BA," & _
				"    CROPTYPE," & _
				"    OWNER," & _
				"    STATUS," & _
				"    DIVISION," & _
				"    PLANTINGDATE," & _
				"    ACTIVITY_HEADING," & _
				"    JOBGROUPCODE," & _
				"    JOBGROUPNAME," & _
				"    JOBCODE," & _
				"    JOBNAME," & _
				"    sum(balanceamount) balanceamount," & _
				"    sum(balanceqty) balanceqty," & _
				"    sum(amount) amount," & _
				"    sum(qty_ha_or_kg) qty_ha_or_kg" & _
				" from (" & _
				" SELECT a.BA," & _
				"       a.CROPTYPE," & _
				"       a.OWNER," & _
				"       a.STATUS," & _
				"       a.DIVISION," & _
				"       a.PLANTINGDATE," & _
				"       a.ACTIVITY_HEADING," & _
				"       a.JOBGROUPCODE," & _
				"       a.JOBGROUPNAME," & _
				"       a.JOBCODE," & _
				"       a.JOBNAME," & _
				"       sum(a.amount) balanceamount," & _
				"       sum(a.QTY_HA_OR_KG) balanceqty," & _
				"       0 amount," & _
				"       0 qty_ha_or_kg" & _
				" from view_accountreport_tm_tbm a" & _
				" where" & _
				" a.period_month < " & Bulan & " AND a.period_year = " & Tahun & " AND a.ba = '" & Session("fcBA") & "'" & _
				" group by a.BA," & _
				"       a.CROPTYPE," & _
				"       a.OWNER," & _
				"       a.STATUS," & _
				"       a.DIVISION," & _
				"       a.PLANTINGDATE," & _
				"       a.ACTIVITY_HEADING," & _
				"       a.JOBGROUPCODE," & _
				"       a.JOBGROUPNAME," & _
				"       a.JOBCODE," & _
				"       a.JOBNAME" & _
				" union" & _
				" SELECT a.BA," & _
				"       a.CROPTYPE," & _
				"       a.OWNER," & _
				"       a.STATUS," & _
				"       a.DIVISION," & _
				"       a.PLANTINGDATE," & _
				"       a.ACTIVITY_HEADING," & _
				"       a.JOBGROUPCODE," & _
				"       a.JOBGROUPNAME," & _
				"       a.JOBCODE," & _
				"       a.JOBNAME," & _
				"       0 BalanceAmount," & _
				"       0 BalanceQty," & _
				"       sum(AMOUNT) amount," & _
				"       sum(QTY_HA_OR_KG) QTY_HA_OR_KG" & _
				" FROM view_accountreport_tm_tbm a" & _
				" WHERE a.period_month = " & Bulan & " AND a.period_year = " & Tahun & " AND a.ba = '" & Session("fcBA") & "'" & _
				" group by a.BA," & _
				"       a.CROPTYPE," & _
				"       a.OWNER," & _
				"       a.STATUS," & _
				"       a.DIVISION," & _
				"       a.PLANTINGDATE," & _
				"       a.ACTIVITY_HEADING," & _
				"       a.JOBGROUPCODE," & _
				"       a.JOBGROUPNAME," & _
				"       a.JOBCODE," & _
				"       a.JOBNAME" & _
				" )" & _      
				" where ACTIVITY_HEADING is not null  group by" & _
				"    BA," & _
				"    CROPTYPE," & _
				"    OWNER," & _
				"    STATUS," & _
				"    DIVISION," & _
				"    PLANTINGDATE," & _
				"    ACTIVITY_HEADING," & _
				"    JOBGROUPCODE," & _
				"    JOBGROUPNAME," & _
				"    JOBCODE," & _
				"    JOBNAME"

Set rsARTMTBM = Server.CreateObject("ADODB.Recordset")
rsARTMTBM.CursorType = 3
rsARTMTBM.CursorLocation = 3
rsARTMTBM.Open SelectARTMTBM, connstr

if not rsARTMTBM.eof then
	Do While Not rsARTMTBM.eof
	'Inserting Data To Account Report Tm Tbm
	InsertARTMTBM =	" INSERT INTO SAC_ACCOUNTREPORT_TM_TBM (" & _
					" PERIOD_YEAR, PERIOD_MONTH, BA, CROPTYPE," & _
					" OWNERSHIP, STATUS, DIVISION," & _
					" PLANTINGDATE, ACTIVITY_HEADING, ACCOUNTGROUPCODE," & _
					" ACCOUNTGROUP, ACCOUNTCODE, DESCRIPTION,OPENBALANCE_AMT_ACTUAL," & _
					" OPENBALANCE_HA_KG_ACTUAL, TM_AMT_ACTUAL, TM_HA_KG_ACTUAL)" & _
					" VALUES (" & _
					" " & Tahun & "," & _
					" " & Bulan & "," & _
					" '" & Session("fcBA") & "'," & _
					" '" & rsARTMTBM("CROPTYPE") & "'," & _
					" '" & rsARTMTBM("OWNER") & "'," & _
					" '" & rsARTMTBM("STATUS") & "'," & _
					" '" & rsARTMTBM("DIVISION") & "'," & _
					" '" & rsARTMTBM("PLANTINGDATE") & "'," & _
					" '" & rsARTMTBM("ACTIVITY_HEADING") & "'," & _
					" '" & rsARTMTBM("JOBGROUPCODE") & "'," & _
					" '" & rsARTMTBM("JOBGROUPNAME") & "'," & _
					" '" & rsARTMTBM("JOBCODE") & "'," & _
					" '" & rsARTMTBM("JOBNAME") & "'," & _
					" " & rsARTMTBM("BALANCEAMOUNT") & "," & _
					" " & rsARTMTBM("BALANCEQTY") & "," & _
					" " & rsARTMTBM("AMOUNT") & "," & _
					" " & rsARTMTBM("QTY_HA_OR_KG") & "" & _
					" )"
	connstr.execute(InsertARTMTBM)
	
	rsARTMTBM.MoveNext
	loop
end if

'----END ACCOUNT REPORT TM TBM----


'----BEGIN ACCOUNT REPORT MILL----

'Normalisasi Account Report Mill
' NormalisasiARMill =		" DELETE FROM SAC_ACCOUNTREPORT_MILL" & _
						' " WHERE" & _
						' " PERIOD_MONTH = " & Bulan & " AND" & _
						' " PERIOD_YEAR = " & Tahun & " AND" & _
						' " BA ='" & Session("fcBA") & "'"
' connstr.execute(NormalisasiARMill)

'Collecting Data for Account Report Mill
' SelectARMill =	" select" & _
				' " 	 BA," & _
				' "    CROPTYPE," & _
				' "    OWNER," & _
				' "    STATUS," & _
				' "    DIVISION," & _
				' "    PLANTINGDATE," & _
				' "    ACTIVITY_HEADING," & _
				' "    JOBGROUPCODE," & _
				' "    JOBGROUPNAME," & _
				' "    JOBCODE," & _
				' "    JOBNAME," & _
				' "    sum(balanceamount) balanceamount," & _
				' "    sum(balanceqty) balanceqty," & _
				' "    sum(amount) amount," & _
				' "    sum(qty_ha_or_kg) qty_ha_or_kg" & _
				' " from (" & _
				' " SELECT a.BA," & _
				' "       a.CROPTYPE," & _
				' "       a.OWNER," & _
				' "       a.STATUS," & _
				' "       a.DIVISION," & _
				' "       a.PLANTINGDATE," & _
				' "       a.ACTIVITY_HEADING," & _
				' "       a.JOBGROUPCODE," & _
				' "       a.JOBGROUPNAME," & _
				' "       a.JOBCODE," & _
				' "       a.JOBNAME," & _
				' "       sum(a.amount) balanceamount," & _
				' "       sum(a.QTY_HA_OR_KG) balanceqty," & _
				' "       0 amount," & _
				' "       0 qty_ha_or_kg" & _
				' " from view_accountreport_tm_tbm a" & _
				' " where" & _
				' " a.period_month < " & Bulan & " AND a.period_year = " & Tahun & " AND a.ba = '" & Session("fcBA") & "'" & _
				' " group by a.BA," & _
				' "       a.CROPTYPE," & _
				' "       a.OWNER," & _
				' "       a.STATUS," & _
				' "       a.DIVISION," & _
				' "       a.PLANTINGDATE," & _
				' "       a.ACTIVITY_HEADING," & _
				' "       a.JOBGROUPCODE," & _
				' "       a.JOBGROUPNAME," & _
				' "       a.JOBCODE," & _
				' "       a.JOBNAME" & _
				' " union" & _
				' " SELECT a.BA," & _
				' "       a.CROPTYPE," & _
				' "       a.OWNER," & _
				' "       a.STATUS," & _
				' "       a.DIVISION," & _
				' "       a.PLANTINGDATE," & _
				' "       a.ACTIVITY_HEADING," & _
				' "       a.JOBGROUPCODE," & _
				' "       a.JOBGROUPNAME," & _
				' "       a.JOBCODE," & _
				' "       a.JOBNAME," & _
				' "       0 BalanceAmount," & _
				' "       0 BalanceQty," & _
				' "       sum(AMOUNT) amount," & _
				' "       sum(QTY_HA_OR_KG) QTY_HA_OR_KG" & _
				' "  FROM view_accountreport_tm_tbm a" & _
				' " WHERE a.period_month = " & Bulan & " AND a.period_year = " & Tahun & " AND a.ba = '" & Session("fcBA") & "'" & _
				' " group by a.BA," & _
				' "       a.CROPTYPE," & _
				' "       a.OWNER," & _
				' "       a.STATUS," & _
				' "       a.DIVISION," & _
				' "       a.PLANTINGDATE," & _
				' "       a.ACTIVITY_HEADING," & _
				' "       a.JOBGROUPCODE," & _
				' "       a.JOBGROUPNAME," & _
				' "       a.JOBCODE," & _
				' "       a.JOBNAME" & _
				' " )" & _      
				' " group by" & _
				' "    BA," & _
				' "    CROPTYPE," & _
				' "    OWNER," & _
				' "    STATUS," & _
				' "    DIVISION," & _
				' "    PLANTINGDATE," & _
				' "    ACTIVITY_HEADING," & _
				' "    JOBGROUPCODE," & _
				' "    JOBGROUPNAME," & _
				' "    JOBCODE," & _
				' "    JOBNAME"

' Set rsARMill = Server.CreateObject("ADODB.Recordset")
' rsARMill.CursorType = 3
' rsARMill.CursorLocation = 3
' rsARMill.Open SelectARMill, connstr

' if not rsARMill.eof then
	' Do While Not rsARMill.eof
	'Inserting Data To Account Report Mill
	' InsertARMill =	" INSERT INTO SAC_ACCOUNTREPORT_MILL (" & _
					' " PERIOD_YEAR, PERIOD_MONTH, BA, CROPTYPE," & _
					' " OWNERSHIP, STATUS, DIVISION," & _
					' " PLANTINGDATE, ACTIVITY_HEADING, ACCOUNTGROUPCODE," & _
					' " ACCOUNTGROUP, ACCOUNTCODE, DESCRIPTION,OPENBALANCE_AMT_ACTUAL," & _
					' " OPENBALANCE_HA_KG_ACTUAL, TY_AMT_ACTUAL, TY_HA_KG_ACTUAL)" & _
					' " VALUES (" & _
					' " " & Tahun & "," & _
					' " " & Bulan & "," & _
					' " '" & Session("fcBA") & "'," & _
					' " '" & rsARMill("CROPTYPE") & "'," & _
					' " '" & rsARMill("OWNER") & "'," & _
					' " '" & rsARMill("STATUS") & "'," & _
					' " '" & rsARMill("DIVISION") & "'," & _
					' " '" & rsARMill("PLANTINGDATE") & "'," & _
					' " '" & rsARMill("ACTIVITY_HEADING") & "'," & _
					' " '" & rsARMill("JOBGROUPCODE") & "'," & _
					' " '" & rsARMill("JOBGROUPNAME") & "'," & _
					' " '" & rsARMill("JOBCODE") & "'," & _
					' " '" & rsARMill("JOBNAME") & "'," & _
					' " " & rsARMill("BALANCEAMOUNT") & "," & _
					' " " & rsARMill("BALANCEQTY") & "," & _
					' " " & rsARMill("AMOUNT") & "," & _
					' " " & rsARMill("QTY_HA_OR_KG") & "" & _
					' " )"
	' connstr.execute(InsertARMill)
	' rsARTMTBM.MoveNext
	' loop
' end if

'----END ACCOUNT REPORT MILL----


TimeAkhir=Time()
Lama =TimeAkhir-TimeAwal
Response.Write "Finish ... " & time() & "<br /><br />"
Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."%>

<%if err.number <> 0 then%>
    <br /><br /><img src="images/error.gif" />
<%
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
	cmdView	= "True"
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>	
<%Else%>
    <br /><br />
	<img src="images/ok.gif" /><font color=Blue pointsize=9 face=Verdana>Transfer to SAC ACCOUNT REPORT TM TBM is succesfully...</font>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
<%End if%>