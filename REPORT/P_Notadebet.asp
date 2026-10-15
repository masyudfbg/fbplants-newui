<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim rsJob, strJob, rsSourceType, strSourceType, rsCA, strCA, v_fcname
Dim ScriptName, v_txtYear, v_txtMonth, ZeroValue, v_txtJobCode, v_txtFcba
Dim NO, cClass, nSubTotal, nGrandTotal
v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")
v_txtJobCode=	Request("txtJobCode")
v_txtFcba 	= 	Request("txtFcba")

strDefaultBa =  " SELECT " & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _ 
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

strFCBA = "SELECT BUSINESSUNIT_API.get_fcname('" & v_txtFcba & "') fcname from dual "
Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strFCBA, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	v_fcname = rsDefaultBa("fcname")
Else
	v_fcname = session("BA") 
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
ZeroValue	=	0
nGrandTotal	=	0

if Len(Trim(v_txtYear)) = 0 then
	v_txtYear	= Year(Session("v_Costbook"))
end if
if Len(Trim(v_txtMonth)) = 0 then
	v_txtMonth	= Month(Session("v_Costbook"))
end if
if Len(Trim(v_txtMonth)) = 0 then
	v_txtJobCode	= "xxxxxxx"
end if

strJob		= 	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & v_txtFcba & "' AND FCCODE = '" & v_txtJobCode & "'"
strJob		=	strJob & " ORDER BY FCCODE ASC"

Set rsJob = Server.CreateObject("ADODB.Recordset")
rsJob.Open strJob, connstr, adOpenStatic 

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Nota Debet</title>
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
	    <td>
			<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
			<%
				Dim rsFcba
				Set rsFcba = Server.CreateObject("ADODB.Recordset")
				rsFcba.CursorLocation = adUseClient
				strsql =    " SELECT  fccode fcba from businessunit " & _
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
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
		<td><select name="txtJobCode" >
		<%
		Dim rsJobCat, strJobCat
		strJobCat	=	"select B.FCCODE,b.FCNAME,A.FCCODE BU from businessunit a inner join jobmaster b on a.jobintraco=b.fccode " & _
						"where a.fccompanycode='" & Session("COMPANYCODE") & "' " & _
		                "ORDER BY FCCODE ASC"
		
		response.write strJobCat
		
		Set rsJobCat = Server.CreateObject("ADODB.Recordset")
		rsJobCat.Open strJobCat, connstr, adOpenStatic 
		
		if not rsJobCat.eof then
		Do While Not rsJobCat.eof
		if rsJobCat("FCCODE") = v_txtJobCode then%>
		        <option value="<%=rsJobCat("FCCODE")%>" Selected><%=rsJobCat("FCCODE")%>-<%=rsJobCat("FCNAME")%></option> 
		     <%Else%>
		        <option value="<%=rsJobCat("FCCODE")%>"><%=rsJobCat("BU")%>-<%=rsJobCat("FCNAME")%></option>  
		    <%End if
			rsJobCat.movenext
		loop
		end if
		rsJobCat.Close
		set rsJobCat = Nothing
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
    <th colspan="53" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=v_fcname%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>NOTA DEBET - CREDIT</h3>
		<h5>PERIODE&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h5>
	</div>
	</div>
   </th>
  </tr>
<%if not rsJob.eof then
	while not rsJob.eof%>
<tr>
<th colspan="53" align="left" style="background-color:#FFFFFF; border:none; color:#000000;"><%=rsJob("FCCODE")%>&nbsp;<%=rsJob("FCNAME")%></th>
</tr>
<%
		rsJob.MoveNext
	Wend
	rsJob.Close
	set rsJob = Nothing
end if
%>
	
	<tr>
		<th rowspan="2">NO</th>
		<th rowspan="2">TANGGAL</th>
		<th rowspan="2">NO.BUKTI</th>
		<th rowspan="2">REFERENSI</th>
		<th rowspan="2">KETERANGAN</th>
		<th rowspan="2">CHARGE JOB</th>
		<th rowspan="2">CHARGE TYPE</th>
		<th rowspan="2">CHARGE CODE</th>
		<th colspan="2">JUMLAH (Rp.)</th>
		<th colspan="5">REFERENSI</th>
	</tr>
	<tr>
		<th>DEBET</th>
		<th>CREDIT</th>
		<th>REFERENSI 01</th>
		<th>REFERENSI 02</th>
		<th>REFERENSI 03</th>
		<th>REFERENSI 04</th>
		<th>REFERENSI 05</th>
	</tr>
    </thead>
<tbody>
	<tr>
		<col span="8">
		<col span="2" style="background-color:#fffea5">
	</tr>
<%
	Dim rsOpeningBalance, strOpeningBalance
	strOpeningBalance	= 	" SELECT NVL(OPENING_BALANCE,0) OPENING_BALANCE FROM FAC_TRIALBALANCE" & _
							" WHERE FCBA = '" & v_txtFcba & "' AND FCCODE = '" & v_txtJobCode & "'" & _
							" AND PERIOD_YEAR = " & Cint(v_txtYear) & _
							" AND PERIOD_MONTH = " & Cint(v_txtMonth)

Set rsOpeningBalance = Server.CreateObject("ADODB.Recordset")
rsOpeningBalance.Open strOpeningBalance, connstr, adOpenStatic 

%>
<%if rsOpeningBalance.bof and rsOpeningBalance.eof then%>
	<tr><td align="left" colspan="15">NO DATA AVAILABLE</td></tr>
<%else
	nGrandTotal	=	nGrandTotal	+	Cdbl(rsOpeningBalance("OPENING_BALANCE"))
	while not rsOpeningBalance.eof%>
	<tr bgcolor="#9CB770">
		<td align="center" colspan="8"><strong><i>Saldo Awal</i></strong></td>
		<td align="center" colspan="5"><strong><i><%=FormatNumber(rsOpeningBalance("OPENING_BALANCE"),2)%></i></strong></td>
		<td align="center" colspan="8"></td>
	</tr>
<%
rsOpeningBalance.MoveNext
Wend
rsOpeningBalance.Close
set rsOpeningBalance = Nothing
end if
%>
<%
	strSourceType = 	" SELECT 'BA' SOURCETYPE, 'BANK' SOURCENAME, 1 SEQ FROM DUAL" & _
						" UNION ALL" & _
						" SELECT 'PR' SOURCETYPE, 'PURCHASING' SOURCENAME, 2 SEQ FROM DUAL" & _
						" UNION ALL" & _
						" SELECT 'ST' SOURCETYPE, 'MATERIAL' SOURCENAME, 3 SEQ FROM DUAL" & _
						" UNION ALL" & _
						" SELECT 'EM' SOURCETYPE, 'EMPLOYEE' SOURCENAME, 4 SEQ FROM DUAL" & _
						" UNION ALL" & _
						" SELECT 'WS' SOURCETYPE, 'WORKSHOP' SOURCENAME, 5 SEQ FROM DUAL" & _
						" UNION ALL" & _
						" SELECT 'VH' SOURCETYPE, 'VEHICLE' SOURCENAME, 6 SEQ FROM DUAL" & _
						" UNION ALL" & _
						" SELECT 'WC' SOURCETYPE, 'CONTRACT' SOURCENAME, 7 SEQ FROM DUAL" & _
						" UNION ALL" & _
						" SELECT 'MN' SOURCETYPE, 'NURSERY' SOURCENAME, 8 SEQ FROM DUAL"& _
						" UNION ALL" & _
						" SELECT 'TT' SOURCETYPE, 'TRANSPORTTASK' SOURCENAME, 9 SEQ FROM DUAL"& _
						" UNION ALL" & _
						" SELECT 'RT' SOURCETYPE, 'RECEIVEDTASK' SOURCENAME, 10 SEQ FROM DUAL"& _
						" ORDER BY SEQ ASC"

Set rsSourceType = Server.CreateObject("ADODB.Recordset")
rsSourceType.Open strSourceType, connstr, adOpenStatic 

%>
<%if rsSourceType.bof and rsSourceType.eof then%>
	<tr><td align="left" colspan="15">NO DATA AVAILABLE</td></tr>
<%else
	while not rsSourceType.eof%>
	<tr><td align="left" colspan="15"bgcolor="F0F0B0"><strong><i><%=rsSourceType("SOURCENAME")%><i></strong></td></tr>

<%
	if rsSourceType("SOURCETYPE") = Cstr("BA") then
	strCA	=	"   SELECT dataca.fcba fcba," & _
				"          dataca.tdate tdate," & _
				"          dataca.accountingyear," & _
				"          dataca.accountingperiod," & _
				"          dataca.jobcode jobcode," & _
				"          dataca.sourcetype sourcetype," & _
				"          dataca.TITLE_REFERENCE_01 TITLE_REFERENCE_01," & _
				"          dataca.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				"          dataca.TITLE_REFERENCE_02 TITLE_REFERENCE_02," & _
				"          dataca.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				"          dataca.TITLE_REFERENCE_03 TITLE_REFERENCE_03," & _
				"          dataca.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				"          dataca.TITLE_REFERENCE_04 TITLE_REFERENCE_04," & _
				"          dataca.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				"          dataca.TITLE_REFERENCE_05 TITLE_REFERENCE_05," & _
				"          dataca.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				"          ba.fcname reference," & _
				"          dataca.referenceno referenceno," & _
				"          pvd.remarks remarks," & _
				"          pvd.chargejob," & _
				"          jb.fcname jobname," & _
				"          pvd.chargetype," & _
				"          pvd.chargecode," & _
				"          NVL (dataca.quantity, 0) quantity," & _
				"          ROUND (NVL (dataca.amount, 0), 0) amount" & _
				"          FROM cost_activity dataca  " & _
				"		   INNER JOIN paymentvoucherdetail pvd on (dataca.fcba = pvd.fcba " & _
				"          										  	AND dataca.referenceno = pvd.vouchercode " & _
				"          											AND dataca.jobcode = pvd.jobcode " & _
				"											        AND dataca.targettype = pvd.locationtype " & _
				"											        AND dataca.targetcode = pvd.locationcode " & _
                "											        AND dataca.auto_reference_05 = pvd.lineno " & _
                "          										   ) " & _
				"			INNER JOIN paymentvoucher pv ON ( pv.fcba = pvd.fcba AND pv.vouchercode = pvd.vouchercode) " & _
				"			INNER JOIN bank ba ON (pv.fcba = ba.fcba AND pv.bankcode = ba.fccode) " & _
				"			LEFT JOIN job jb ON (pvd.chargejob = jb.fccode AND pvd.fcba = jb.fcba) " & _				     
				"    		WHERE " & _
				"           dataca.sourcetype = 'BA'" & _
				"           AND dataca.fcba = '"& v_txtFcba &"'" & _
				"           AND dataca.jobcode = '"& v_txtJobCode &"'"

	' strCA	=	"   SELECT dataca.fcba fcba," & _
				' "          dataca.tdate tdate," & _
				' "          dataca.accountingyear," & _
				' "          dataca.accountingperiod," & _
				' "          dataca.jobcode jobcode," & _
				' "          dataca.sourcetype sourcetype," & _
				' "          dataca.TITLE_REFERENCE_01 TITLE_REFERENCE_01," & _
				' "          dataca.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				' "          dataca.TITLE_REFERENCE_02 TITLE_REFERENCE_02," & _
				' "          dataca.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				' "          dataca.TITLE_REFERENCE_03 TITLE_REFERENCE_03," & _
				' "          dataca.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				' "          dataca.TITLE_REFERENCE_04 TITLE_REFERENCE_04," & _
				' "          dataca.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				' "          dataca.TITLE_REFERENCE_05 TITLE_REFERENCE_05," & _
				' "          dataca.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				' "          ba.fcname reference," & _
				' "          dataca.referenceno referenceno," & _
				' "          pvd.remarks remarks," & _
				' "          pvd.chargejob," & _
				' "          jb.fcname jobname," & _
				' "          pvd.chargetype," & _
				' "          pvd.chargecode," & _
				' "          NVL (dataca.quantity, 0) quantity," & _
				' "          ROUND (NVL (dataca.amount, 0), 0) amount" & _
				' "     FROM cost_activity dataca," & _
				' "          paymentvoucherdetail pvd," & _
				' "          paymentvoucher pv," & _
				' "          bank ba," & _
				' "          job jb" & _
				' "    WHERE     dataca.fcba = pvd.fcba" & _
				' "          AND dataca.referenceno = pvd.vouchercode" & _
				' "          AND dataca.jobcode = pvd.jobcode" & _
				' "          AND dataca.targettype = pvd.locationtype" & _
				' "          AND dataca.targetcode = pvd.locationcode" & _
				' "          AND dataca.fcba = pv.fcba" & _
				' "          AND dataca.referenceno = pv.vouchercode" & _
				' "          AND dataca.tdate = pv.datecreated" & _
				' "          AND dataca.auto_reference_05 = pvd.lineno" & _
				' "          AND pv.fcba = pvd.fcba" & _
				' "          AND pv.vouchercode = pvd.vouchercode" & _
				' "          AND pv.fcba = ba.fcba" & _
				' "          AND pvd.chargejob = jb.fccode" & _
				' "          AND pvd.fcba = jb.fcba" & _
				' "          AND pv.bankcode = ba.fccode" & _
				' "          AND dataca.sourcetype = 'BA'" & _
				' "          AND dataca.fcba = '"& v_txtFcba &"'" & _
				' "          AND dataca.jobcode = '"& v_txtJobCode &"'"

				' " select" & _
				' "     dataca.fcba fcba" & _
				' "     , dataca.tdate tdate" & _
				' "     , dataca.accountingyear" & _
    '             "     , dataca.accountingperiod" & _
				' "     , dataca.jobcode jobcode" & _
				' "     , dataca.sourcetype sourcetype" & _
				' "	  , dataca.TITLE_REFERENCE_01 TITLE_REFERENCE_01, dataca.AUTO_REFERENCE_01 AUTO_REFERENCE_01" & _
				' " 	  , dataca.TITLE_REFERENCE_02 TITLE_REFERENCE_02, dataca.AUTO_REFERENCE_02 AUTO_REFERENCE_02" & _
				' " 	  , dataca.TITLE_REFERENCE_03 TITLE_REFERENCE_03, dataca.AUTO_REFERENCE_03 AUTO_REFERENCE_03" & _
				' " 	  , dataca.TITLE_REFERENCE_04 TITLE_REFERENCE_04, dataca.AUTO_REFERENCE_04 AUTO_REFERENCE_04" & _
				' " 	  , dataca.TITLE_REFERENCE_05 TITLE_REFERENCE_05, dataca.AUTO_REFERENCE_05 AUTO_REFERENCE_05" & _
				' "     , ba.fcname reference" & _
				' "     , dataca.referenceno referenceno" & _
				' "     , pvd.remarks remarks" & _
				' "     , nvl (dataca.quantity, 0) quantity" & _
				' "     , round(nvl (dataca.amount, 0), 0) amount" & _
				' " from    cost_activity dataca" & _
				' "         , paymentvoucherdetail pvd" & _
				' "         , paymentvoucher pv" & _
				' "         , bank ba" & _
				' " where   dataca.fcba = pvd.fcba" & _
				' "         and dataca.referenceno = pvd.vouchercode" & _
				' "         and dataca.jobcode = pvd.jobcode" & _
				' "         and dataca.targettype = pvd.locationtype" & _
				' "         and dataca.targetcode = pvd.locationcode" & _
				' "         and dataca.fcba = pv.fcba" & _
				' "         and dataca.referenceno = pv.vouchercode" & _
				' "         and dataca.tdate = pv.datecreated" & _
				' "         and dataca.auto_reference_05 = pvd.lineno" & _
				' "         and pv.fcba = pvd.fcba" & _
				' "         and pv.vouchercode = pvd.vouchercode" & _
				' "         and pv.fcba = ba.fcba" & _
				' "         and pv.bankcode = ba.fccode" & _
				' "         and dataca.sourcetype = 'BA'" & _
				' "         and dataca.fcba = '" & v_txtFcba & "' and dataca.jobcode = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("PR") then
	strCA	=	" SELECT" & _
				" DATACA.FCBA FCBA," & _
				" DATACA.TDATE TDATE," & _
				" dataca.accountingyear, " & _
                " dataca.accountingperiod, " & _
				" DATACA.JOBCODE JOBCODE," & _
				" DATACA.SOURCETYPE SOURCETYPE," & _
				" DATACA.SOURCECODE_STATUS REFERENCE," & _
				" DATACA.SOURCECODE REFERENCENO," & _
				" DATACA.REFERENCENO REMARKS," & _
				" DATACA.TITLE_REFERENCE_01 TITLE_REFERENCE_01, DATACA.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				" DATACA.TITLE_REFERENCE_02 TITLE_REFERENCE_02, DATACA.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				" DATACA.TITLE_REFERENCE_03 TITLE_REFERENCE_03, DATACA.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				" DATACA.TITLE_REFERENCE_04 TITLE_REFERENCE_04, DATACA.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				" DATACA.TITLE_REFERENCE_05 TITLE_REFERENCE_05, DATACA.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				" NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				" NVL (DATACA.AMOUNT, 0) AMOUNT" & _
				" FROM" & _
				" COST_ACTIVITY DATACA" & _
				" WHERE" & _
				" DATACA.SOURCETYPE = 'PR'" & _
				" AND DATACA.FCBA = '" & v_txtFcba & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("ST") then
	strCA	=	" Select fcba" & _
				"        , tdate" & _
				"        , accountingyear " & _
                "        , accountingperiod " & _
				"        , jobcode" & _
				"        , sourcetype" & _
				"        , chargejob" & _
				"        , chargetype" & _
				"        , chargecode" & _
				"        , reference" & _
				"        , referenceno" & _
				"        , remarks" & _
				"        , quantity" & _
				"        , amount" & _
				" from (" & _
				"     Select isud.fcba" & _
				"            , isu.sivdate tdate" & _
				"            , to_char(isu.sivdate,'YYYY') accountingyear " & _
                "            , to_char(isu.sivdate,'MM') accountingperiod " & _
				"            , isud.jobcode" & _
				"            , '' chargejob" & _
				"            , '' chargetype" & _
				"            , '' chargecode" & _
				"            , 'ST' sourcetype" & _
				"            , stk.fcname reference" & _
				"            , isud.sivcode referenceno" & _
				"            , isud.remarks remarks" & _
				"            , nvl(sum(nvl(isud.quantity, 0)), 0) quantity" & _
				"            , nvl(sum(nvl(isud.amount, 0)), 0) amount" & _
				"     from Issue isu" & _
				"          , IssueDetails isud" & _
				"          , stock stk" & _
				"     Where isu.fcba = isud.fcba" & _
				"           and isu.sivcode = isud.sivcode" & _
				"           and isud.fcba = stk.fcba" & _
				"           and isud.stockcode = stk.fccode" & _
				"     group by isud.fcba" & _
				"              , isu.sivdate" & _
				"              , to_char(isu.sivdate,'YYYY') " & _
                "              , to_char(isu.sivdate,'MM')  " & _
				"              , isud.jobcode" & _
				"              , stk.fcname" & _
				"              , isud.sivcode" & _
				"              , isud.remarks" & _
				"       ) dataca" & _
				" where dataca.sourcetype = 'ST' " & _
				"       and dataca.fcba = '" & v_txtFcba & "' and dataca.jobcode = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("WS") then 
	strCA	=	"   SELECT DATACA.FCBA FCBA," & _
				"          DATACA.TDATE TDATE," & _
				"          dataca.accountingyear," & _
				"          dataca.accountingperiod," & _
				"          DATACA.JOBCODE JOBCODE," & _
				"          DATACA.SOURCETYPE SOURCETYPE," & _
				"          DATACA.TITLE_REFERENCE_01 TITLE_REFERENCE_01," & _
				"          DATACA.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				"          DATACA.TITLE_REFERENCE_02 TITLE_REFERENCE_02," & _
				"          DATACA.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				"          DATACA.TITLE_REFERENCE_03 TITLE_REFERENCE_03," & _
				"          DATACA.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				"          DATACA.TITLE_REFERENCE_04 TITLE_REFERENCE_04," & _
				"          DATACA.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				"          DATACA.TITLE_REFERENCE_05 TITLE_REFERENCE_05," & _
				"          DATACA.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				"          WA.TIME REFERENCENO," & _
				"          WA.CHARGEJOB," & _
				"          WA.CHARGETYPE," & _
				"          WA.CHARGECODE," & _
				"          EM.FCNAME REFERENCE," & _
				"          WA.REMARKS REMARKS," & _
				"          NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				"          ROUND (" & _
				"             DECODE (" & _
				"                NVL (DATACA.QUANTITY, 0)," & _
				"                0, 0," & _
				"                  (NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0))" & _
				"                * NVL (WA.UNIT, 0))," & _
				"             0)" & _
				"             AMOUNT" & _
				"     FROM COST_ACTIVITY DATACA, WORKSHOPACTIVITY WA, WORKSHOP EM, JOB JB" & _
				"    WHERE     (    DATACA.FCBA = WA.FCBA" & _
				"               AND DATACA.SOURCECODE = WA.WORKSHOPCODE" & _
				"               AND DATACA.TDATE = WA.WORKSHOPDATE" & _
				"               AND DATACA.JOBCODE = WA.JOB)" & _
				"          AND (WA.FCBA = EM.FCBA AND WA.WORKSHOPCODE = EM.FCCODE)" & _
				"          AND WA.CHARGEJOB = JB.FCCODE" & _
				"          AND WA.FCBA = JB.FCBA" & _
				"          AND DATACA.SOURCETYPE = 'WS'" & _
				"          AND DATACA.FCBA = '"& v_txtFcba &"'" & _
				"          AND DATACA.JOBCODE = '"& v_txtJobCode &"'" 

				' " SELECT" & _
				' " DATACA.FCBA FCBA," & _
				' " DATACA.TDATE TDATE," & _
				' " dataca.accountingyear, " & _
    '             " dataca.accountingperiod, " & _
				' " DATACA.JOBCODE JOBCODE," & _
				' " DATACA.SOURCETYPE SOURCETYPE," & _
				' " DATACA.TITLE_REFERENCE_01 TITLE_REFERENCE_01, DATACA.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				' " DATACA.TITLE_REFERENCE_02 TITLE_REFERENCE_02, DATACA.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				' " DATACA.TITLE_REFERENCE_03 TITLE_REFERENCE_03, DATACA.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				' " DATACA.TITLE_REFERENCE_04 TITLE_REFERENCE_04, DATACA.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				' " DATACA.TITLE_REFERENCE_05 TITLE_REFERENCE_05, DATACA.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				' " WA.TIME REFERENCENO," & _
				' " EM.FCNAME REFERENCE," & _
				' " WA.REMARKS REMARKS," & _
				' " NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				' " round(DECODE(NVL(DATACA.QUANTITY, 0),0,0,(NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0)) * NVL (WA.UNIT, 0)),0) AMOUNT " & _
				' " FROM COST_ACTIVITY DATACA," & _
				' " WORKSHOPACTIVITY WA," & _
				' " WORKSHOP EM" & _
				' " WHERE" & _
				' " (" & _
				' " DATACA.FCBA = WA.FCBA" & _
				' " AND DATACA.SOURCECODE = WA.WORKSHOPCODE" & _
				' " AND DATACA.TDATE = WA.WORKSHOPDATE" & _
				' " AND DATACA.JOBCODE = WA.JOB" & _
				' " )" & _
				' " AND" & _
				' " (WA.FCBA = EM.FCBA AND WA.WORKSHOPCODE = EM.FCCODE)" & _
				' " AND DATACA.SOURCETYPE = 'WS'" & _
				' " AND DATACA.FCBA = '" & v_txtFcba & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("VH") then 
	strCA	=	"   SELECT DATACA.FCBA FCBA," & _
				"          DATACA.TDATE TDATE," & _
				"          dataca.accountingyear," & _
				"          dataca.accountingperiod," & _
				"          DATACA.JOBCODE JOBCODE," & _
				"          DATACA.SOURCETYPE SOURCETYPE," & _
				"          DATACA.TITLE_REFERENCE_01 TITLE_REFERENCE_01," & _
				"          DATACA.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				"          DATACA.TITLE_REFERENCE_02 TITLE_REFERENCE_02," & _
				"          DATACA.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				"          DATACA.TITLE_REFERENCE_03 TITLE_REFERENCE_03," & _
				"          DATACA.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				"          DATACA.TITLE_REFERENCE_04 TITLE_REFERENCE_04," & _
				"          DATACA.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				"          DATACA.TITLE_REFERENCE_05 TITLE_REFERENCE_05," & _
				"          DATACA.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				"          VA.TIME REFERENCENO," & _
				"          VH.FCNAME REFERENCE," & _
				"          VA.REMARKS REMARKS," & _
				"          VA.CHARGEJOB," & _
				"          JB.FCNAME JOBNAME," & _
				"          VA.CHARGETYPE," & _
				"          VA.CHARGECODE," & _
				"          NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				"          ROUND (" & _
				"             DECODE (" & _
				"                NVL (DATACA.QUANTITY, 0)," & _
				"                0, 0," & _
				"                  (NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0))" & _
				"                * NVL (VA.UNIT, 0))," & _
				"             0)" & _
				"             AMOUNT" & _
				"     FROM COST_ACTIVITY DATACA, VEHICLEACTIVITY VA, VEHICLE VH, JOB JB" & _
				"    WHERE     (    DATACA.FCBA = VA.FCBA" & _
				"               AND DATACA.SOURCECODE = VA.VEHICLECODE" & _
				"               AND DATACA.TDATE = VA.VEHICLEDATE" & _
				"               AND DATACA.JOBCODE = VA.JOBCODE)" & _
				"          AND (VA.FCBA = VH.FCBA AND VA.VEHICLECODE = VH.FCCODE)" & _
				"          AND VA.CHARGEJOB = JB.FCCODE" & _
				"          AND VA.FCBA = JB.FCBA" & _
				"          AND DATACA.SOURCETYPE = 'VH'" & _
				"          AND DATACA.FCBA = '"& v_txtFcba &"'" & _
				"          AND DATACA.JOBCODE = '"& v_txtJobCode &"'" 

				' " SELECT" & _
				' " DATACA.FCBA FCBA," & _
				' " DATACA.TDATE TDATE," & _
				' " dataca.accountingyear, " & _
    '             " dataca.accountingperiod, " & _
				' " DATACA.JOBCODE JOBCODE," & _
				' " DATACA.SOURCETYPE SOURCETYPE," & _
				' " DATACA.TITLE_REFERENCE_01 TITLE_REFERENCE_01, DATACA.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				' " DATACA.TITLE_REFERENCE_02 TITLE_REFERENCE_02, DATACA.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				' " DATACA.TITLE_REFERENCE_03 TITLE_REFERENCE_03, DATACA.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				' " DATACA.TITLE_REFERENCE_04 TITLE_REFERENCE_04, DATACA.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				' " DATACA.TITLE_REFERENCE_05 TITLE_REFERENCE_05, DATACA.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				' " VA.TIME REFERENCENO," & _
				' " VH.FCNAME REFERENCE," & _
				' " VA.REMARKS REMARKS," & _
				' " NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				' " ROUND(DECODE(NVL (DATACA.QUANTITY, 0),0,0,(NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0)) * NVL (VA.UNIT, 0)),0) AMOUNT" & _
				' " FROM COST_ACTIVITY DATACA," & _
				' " VEHICLEACTIVITY VA," & _
				' " VEHICLE VH" & _
				' " WHERE" & _
				' " (" & _
				' " DATACA.FCBA = VA.FCBA" & _
				' " AND DATACA.SOURCECODE = VA.VEHICLECODE" & _
				' " AND DATACA.TDATE = VA.VEHICLEDATE" & _
				' " AND DATACA.JOBCODE = VA.JOBCODE" & _
				' " )" & _
				' " AND" & _
				' " (VA.FCBA = VH.FCBA AND VA.VEHICLECODE = VH.FCCODE)" & _
				' " AND DATACA.SOURCETYPE = 'VH'" & _
				' " AND DATACA.FCBA = '" & v_txtFcba & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("MN") then 
	' strCA	=	" SELECT" & _
				' " DATACA.FCBA FCBA," & _
				' " DATACA.TDATE TDATE," & _
				' " DATACA.JOBCODE JOBCODE," & _
				' " DATACA.SOURCETYPE SOURCETYPE," & _
				' " DATACA.SOURCECODE REFERENCENO," & _
				' " DATACA.TITLE_REFERENCE_01 TITLE_REFERENCE_01, DATACA.AUTO_REFERENCE_01 AUTO_REFERENCE_01," & _
				' " DATACA.TITLE_REFERENCE_02 TITLE_REFERENCE_02, DATACA.AUTO_REFERENCE_02 AUTO_REFERENCE_02," & _
				' " DATACA.TITLE_REFERENCE_03 TITLE_REFERENCE_03, DATACA.AUTO_REFERENCE_03 AUTO_REFERENCE_03," & _
				' " DATACA.TITLE_REFERENCE_04 TITLE_REFERENCE_04, DATACA.AUTO_REFERENCE_04 AUTO_REFERENCE_04," & _
				' " DATACA.TITLE_REFERENCE_05 TITLE_REFERENCE_05, DATACA.AUTO_REFERENCE_05 AUTO_REFERENCE_05," & _
				' " NR.DESCRIPTION REFERENCE," & _
				' " NI.REMARKS REMARKS," & _
				' " NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				' " ROUND(DECODE(NVL (DATACA.QUANTITY, 0),0,0, (NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0)) * NVL(NI.QUANTITY,0)),0) AMOUNT " & _
				' " FROM COST_ACTIVITY DATACA," & _
				' " NURSERYISSUE NI," & _
				' " NURSERY NR" & _
				' " WHERE" & _
				' " (" & _
				' " DATACA.FCBA = NI.FCBA" & _
				' " AND DATACA.SOURCECODE = NI.NURSERYCODE" & _
				' " AND DATACA.TDATE = NI.ISSUEDATE" & _
				' " AND DATACA.JOBCODE = NI.JOBCODE" & _
				' " AND DATACA.REFERENCENO = NI.REFERENCE" & _
				' " )" & _
				' " AND" & _
				' " (NI.FCBA = NR.FCBA AND NI.NURSERYCODE = NR.NURSERYCODE)" & _
				' " AND DATACA.SOURCETYPE = 'MN'" & _
				' " AND DATACA.FCBA = '" & v_txtFcba & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
				
	strCA	=	" Select fcba" & _
				"        , tdate" & _
				"        , accountingyear " & _
                "        , accountingperiod " & _
				"        , jobcode" & _
				"        , sourcetype" & _
				"        , reference" & _
				"        , referenceno" & _
				"        , remarks" & _
				"        , quantity" & _
				"        , amount" & _
				" from (" & _
				"     select nid.fcba" & _
				"            , ni.issuedate tdate" & _
				"            , to_char(ni.issuedate,'YYYY') accountingyear" & _
                "            , to_char(ni.issuedate,'MM')   accountingperiod" & _
				"            , nid.jobcode" & _
				"            , 'MN' sourcetype" & _
				"            , nursery_api.get_description(nid.nurserycode, nid.fcba) reference" & _
				"            , nid.documentno referenceno" & _
				"            , nid.remarks" & _
				"            , nvl(sum(nvl(nid.quantity, 0)), 0) quantity" & _
				"            , nvl(sum(nvl(nid.amount, 0)), 0) amount" & _
				"     from nurseryissuehdr ni" & _
				"          , nurseryissuedtl nid" & _
				"     where ni.documentno = nid.documentno" & _
				"           and ni.fcba = nid.fcba" & _
				"     group by nid.fcba" & _
				"              , ni.issuedate" & _
				"              , nid.jobcode              " & _
				"              , nursery_api.get_description(nid.nurserycode, nid.fcba)" & _
				"              , nid.documentno" & _
				"              , nid.remarks" & _
				"       ) dataca" & _
				" where dataca.sourcetype = 'MN' " & _
				"       and dataca.fcba = '" & v_txtFcba & "' and dataca.jobcode = '" & v_txtJobCode & "'"
				
	elseif rsSourceType("SOURCETYPE") = Cstr("EM") then
	  strCA	= " Select * from ( " & _
	          " Select fcba,tdate, accountingyear , accountingperiod, " & _
			  "		   referenceno, sourcetype||' - '||afd REFERENCE, job_api.get_jobcode_name(chargejob,fcba) REMARKS, " & _
			  "	       chargejob, chargetype, chargecode,(mandays/TotalMandays) * amount  amount," & _
              "		   'total Mandays : ' TITLE_REFERENCE_01 , TotalMandays Auto_REFERENCE_01, 'Mandays : ' TITLE_REFERENCE_02 , Mandays Auto_REFERENCE_02, " & _
			  "        'Total Cost : ' TITLE_REFERENCE_03 , amount Auto_REFERENCE_03,'' TITLE_REFERENCE_04 , '' Auto_REFERENCE_04,'' TITLE_REFERENCE_05 , '' Auto_REFERENCE_05 from ( " & _
			  "	Select  cost_activity_api.get_AmountSocLoc( bb.fcba,to_char(bb.fddate,'MMYYYY'), '" & v_txtJobCode & "', bb.Sourcetype, bb.Afd)  Amount, bb.* , " & _
			  "	        sourcetype referenceno," & _
			  "     case  bb.Sourcetype " & _
			  "		 when  'LOC'  then  ATTENDANCE_GAD_API.get_MandayslocPerAfd (bb.fcba,to_char(bb.fddate,'MMYYYY'),bb.Afd,'" & v_txtJobCode & "') " & _
			  "		 when  'SOC'  then  ATTENDANCE_GAD_API.get_MandaysSocPerAfd (bb.fcba,to_char(bb.fddate,'MMYYYY'),bb.Afd,'" & v_txtJobCode & "')  " & _
			  "		 else  0 " & _
			  "	   end    TotalMandays ,bb.fddate tdate, to_char(bb.fddate,'YYYY') accountingyear , to_char(bb.fddate,'MM') accountingperiod " & _
			  "	From  ( " & _
			  "	select a.fcba, a.fddate,'LOC' Sourcetype, employee_api.get_sectionname(a.fcba,a.employeecode) afd, a.chargejob,a.chargetype, a.chargecode, sum(nvl(a.mandays, 0)) mandays " & _
			  "	  from attendance_gad a " & _
			  "	 where a.fcba = '" & v_txtFcba & "' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   and a.jobcode = '" & v_txtJobCode & "' " & _
			  "	group by a.fcba, a.fddate, employee_api.get_sectionname(a.fcba,a.employeecode) ,a.chargejob, a.chargetype, a.chargecode " & _
			  "	union all " & _
			  "	select a.fcba, a.fddate,'SOC' Sourcetype, employee_api.get_sectionname(a.fcba,a.supervision_1) afd, a.chargejob, a.chargetype, a.chargecode, sum(nvl(a.mandays, 0)) mandays " & _
			  "	  from attendance_gad a " & _
			  "	 where a.fcba = '" & v_txtFcba & "' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   and a.jobcode = '" & v_txtJobCode & "' " & _
			  "	   and employee_api.get_sectionname(a.fcba,a.supervision_1) is not null " & _
			  " group by  a.fcba, a.fddate, employee_api.get_sectionname(a.fcba,a.supervision_1) ,a.chargejob, a.chargetype, a.chargecode " & _
			  "	union all " & _
			  "	select a.fcba, a.fddate, 'SOC' Sourcetype, employee_api.get_sectionname(a.fcba,a.supervision_2) afd, a.chargejob, a.chargetype, a.chargecode, sum(nvl(a.mandays, 0)) mandays " & _
			  "	  from attendance_gad a " & _
			  "	 where a.fcba = '" & v_txtFcba & "' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   and a.jobcode = '" & v_txtJobCode & "' " & _
			  "	   and employee_api.get_sectionname(a.fcba,a.supervision_2) is not null " & _
			  "	group by  a.fcba, a.fddate, employee_api.get_sectionname(a.fcba,a.supervision_2) ,a.chargejob, a.chargetype, a.chargecode " & _
			  "	union all " & _
			  "	select  a.fcba, a.fddate, 'SOC' Sourcetype, employee_api.get_sectionname(a.fcba,a.supervision_3) afd, a.chargejob, a.chargetype, a.chargecode, sum(nvl(a.mandays, 0)) mandays " & _
			  "	  from attendance_gad a " & _
			  "	 where a.fcba = '" & v_txtFcba & "' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   and a.jobcode = '" & v_txtJobCode & "' " & _
			  "	   and employee_api.get_sectionname(a.fcba,a.supervision_3) is not null " & _
			  "	group by  a.fcba, a.fddate, employee_api.get_sectionname(a.fcba,a.supervision_3) ,a.chargejob, a.chargetype, a.chargecode " & _
			  "	union all " & _
			  "	select  a.fcba, a.fddate, 'SOC' Sourcetype, employee_api.get_sectionname(a.fcba,a.supervision_4) afd, a.chargejob, a.chargetype, a.chargecode, sum(nvl(a.mandays, 0)) mandays " & _
			  "	  from attendance_gad a " & _
			  "	 where a.fcba = '" & v_txtFcba & "' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   and a.jobcode = '" & v_txtJobCode & "' " & _
			  "	   and employee_api.get_sectionname(a.fcba,a.supervision_4) is not null " & _
			  "	group by a.fcba, a.fddate, employee_api.get_sectionname(a.fcba,a.supervision_4) ,a.chargejob, a.chargetype, a.chargecode ) bb ) where Amount > 0 " & _ 
		      " Union All " & _
        	  " Select  fcba ,fddate tdate, accountingyear , accountingperiod,REFERENCENO,REFERENCE, REMARKS, " & _
			  "		    chargejob, chargetype, chargecode,(mandays/TotalMandays) * amount  amount, " & _
			  "			'total Mandays : ' TITLE_REFERENCE_01 , TotalMandays Auto_REFERENCE_01, 'Mandays : ' TITLE_REFERENCE_02 , Mandays Auto_REFERENCE_02, " & _
			  "			'Total Cost ('||REFERENCE||') : ' TITLE_REFERENCE_03 , amount Auto_REFERENCE_03,'' TITLE_REFERENCE_04 , '' Auto_REFERENCE_04,'' TITLE_REFERENCE_05 , '' Auto_REFERENCE_05 from ( " & _
			  "	   Select fcba ,fddate, to_char(fddate,'YYYY') accountingyear , to_char(fddate,'MM') accountingperiod, " & _
			  "			  REFERENCENO,REFERENCE,job_api.get_jobcode_name(chargejob,fcba) REMARKS,chargejob,chargetype,chargecode , mandays, " & _
			  "			  case  REFERENCENO " & _
			  "				   when  'EM'  then  ATTENDANCE_GAD_API.get_MandaysEMNonAFD(fcba,to_char(fddate,'MMYYYY'),'" & v_txtJobCode & "',REFERENCE) " & _
			  "				   when  'ES'  then  ATTENDANCE_GAD_API.get_MandaysESNonAFD(fcba,to_char(fddate,'MMYYYY'),'" & v_txtJobCode & "',REFERENCE) " & _
			  "				   else  0 " & _
			  "			   end    TotalMandays,amount " & _
			  "	   From    " & _  
			  "	   ( " & _
			  " 	   Select bb.fcba , bb.fddate, " & _
			  "			  cost_activity_api.get_AmountESEM(bb.fcba,'" & v_txtJobCode & "',to_char(bb.fddate,'MMYYYY'),bb.sourcetype,bb.sourcecode) Amount, " & _
			  "			  bb.sourcetype REFERENCENO, sourcecode REFERENCE,bb.chargejob, bb.chargetype, bb.chargecode , sum(bb.mandays) mandays " & _
			  "	   From     " & _  
			  "	   (   " & _
			  "	   select a.fcba, a.fddate, 'EM' sourcetype,employee_api.get_presentgrade(a.fcba,a.employeecode) sourcecode, a.chargejob,a.chargetype, a.chargecode,a.mandays " & _
			  "	   from attendance_gad a  " & _                                            
			  "	   where a.fcba ='" & v_txtFcba & "'  " & _                                                                                                                                                                                                                                      
			  "	   And   a.jobcode =  '" & v_txtJobCode & "' " & _
			  "	   and   employee_api.get_sectionname(a.fcba,a.employeecode) not like 'AFD%' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   union all " & _
			  "	   select a.fcba, a.fddate, 'ES' sourcetype,employee_api.get_presentgrade(a.fcba,a.supervision_1) sourcecode, a.chargejob,a.chargetype, a.chargecode,a.mandays " & _
			  "	   from attendance_gad a     " & _
			  "	   where a.fcba ='" & v_txtFcba & "' " & _
			  "	   And   a.jobcode =  '" & v_txtJobCode & "'" & _
			  "	   and   employee_api.get_sectionname(a.fcba,a.supervision_1) not like 'AFD%' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   union all " & _
			  "	   select a.fcba, a.fddate, 'ES' sourcetype,employee_api.get_presentgrade(a.fcba,a.supervision_2) sourcecode, a.chargejob,a.chargetype, a.chargecode,a.mandays " & _
			  "	   from attendance_gad a   " & _  
			  "	   where a.fcba ='" & v_txtFcba & "' " & _
			  "	   And   a.jobcode =  '" & v_txtJobCode & "' " & _
			  "	   and   employee_api.get_sectionname(a.fcba,a.supervision_2) not like 'AFD%' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   union all " & _
			  "	   select a.fcba, a.fddate, 'ES' sourcetype,employee_api.get_presentgrade(a.fcba,a.supervision_3) sourcecode, a.chargejob,a.chargetype, a.chargecode,a.mandays " & _
			  "	   from attendance_gad a   " & _ 
			  "	   where a.fcba ='" & v_txtFcba & "' " & _
			  "	   And   a.jobcode =  '" & v_txtJobCode & "' " & _
			  "	   and   employee_api.get_sectionname(a.fcba,a.supervision_3) not like 'AFD%' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   union all " & _
			  "	   select a.fcba, a.fddate, 'ES' sourcetype,employee_api.get_presentgrade(a.fcba,a.supervision_4) sourcecode, a.chargejob,a.chargetype, a.chargecode,a.mandays " & _
			  "	   from attendance_gad a  " & _   
			  "	   where a.fcba ='" & v_txtFcba & "' " & _
			  "	   And   a.jobcode =  '" & v_txtJobCode & "' " & _
			  "	   and   employee_api.get_sectionname(a.fcba,a.supervision_4) not like 'AFD%' " & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
			  "    and TO_NUMBER (TO_CHAR (a.fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
			  "	   ) bb  group by bb.fcba, bb.fddate," & _
			  "					  cost_activity_api.get_AmountESEM(bb.fcba,'" & v_txtJobCode & "',to_char(bb.fddate,'MMYYYY'),bb.sourcetype,bb.sourcecode)," & _
			  " 				  bb.sourcetype ,sourcecode ,bb.chargejob, bb.chargetype, bb.chargecode " & _
			  "	  )  ) where Amount > 0  ) dataca where 1=1 "  
			  
			  
	
	' strCA	=	" SELECT fcba," & _
				' "        tdate," & _
				' "        accountingyear," & _
				' "        accountingperiod," & _
				' "        jobcode," & _
				' "        sourcetype," & _
				' "        referenceno," & _
				' "        reference," & _
				' "        remarks," & _
				' "        chargejob," & _
				' "        chargetype," & _
				' "        chargecode," & _
				' "        amount" & _
				' "   FROM (  SELECT ec.fcba," & _
				' "                  TRUNC (transdate, 'mm') tdate," & _
				' "                  TO_CHAR (transdate, 'YYYY') accountingyear," & _
				' "                  TO_CHAR (transdate, 'MM') accountingperiod," & _
				' "                  jobcode," & _
				' "                  sourcetype," & _
				' "                  em.presentgrade referenceno," & _
				' "                  c.PresentGradeDesc reference," & _
				' "                  'ALOKASI SUPERVISI' remarks," & _
				' "                  '' chargejob," & _
				' "                  '' chargetype," & _
				' "                  '' chargecode," & _
				' "                  ROUND (NVL (SUM (NVL (amount, 0)), 0), 0) amount" & _
				' "             FROM employeecb ec," & _
				' "                  employee em," & _
				' "                  (SELECT Seq, PresentGradeCode, PresentGradeDesc" & _
				' "                     FROM (SELECT 1 Seq," & _
				' "                                  'KHL' PresentGradeCode," & _
				' "                                  'KARYAWAN HARIAN LEPAS' PresentGradeDesc" & _
				' "                             FROM DUAL" & _
				' "                           UNION ALL" & _
				' "                           SELECT 2 Seq," & _
				' "                                  'KHT' PresentGradeCode," & _
				' "                                  'KARYAWAN HARIAN TETAP' PresentGradeDesc" & _
				' "                             FROM DUAL" & _
				' "                           UNION ALL" & _
				' "                           SELECT 3 Seq," & _
				' "                                  'KBT' PresentGradeCode," & _
				' "                                  'KARYAWAN BULANAN TETAP' PresentGradeDesc" & _
				' "                             FROM DUAL)) c" & _
				' "            WHERE     ec.fcba = em.fcba" & _
				' "                  AND ec.empcode = em.fccode" & _
				' "                  AND em.PresentGrade = c.PresentGradeCode" & _
				' "                  AND ec.fcba = '"& v_txtFcba &"'" & _
				' "                  AND TO_NUMBER (TO_CHAR (transdate, 'mm')) = "& Cint(v_txtMonth) &"" & _
				' "                  AND TO_NUMBER (TO_CHAR (transdate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
				' "                  AND jobcode = '"& v_txtJobCode &"'" & _
				' "                  AND sourcetype = 'ES'" & _
				' "         GROUP BY ec.fcba," & _
				' "                  TRUNC (transdate, 'mm')," & _
				' "                  TO_CHAR (transdate, 'YYYY')," & _
				' "                  TO_CHAR (transdate, 'MM')," & _
				' "                  em.presentgrade," & _
				' "                  c.PresentGradeDesc," & _
				' "                  jobcode," & _
				' "                  sourcetype" & _
				' "         UNION ALL" & _
				' "         SELECT fcba," & _
				' "                tdate," & _
				' "                accountingyear," & _
				' "                accountingperiod," & _
				' "                jobcode," & _
				' "                sourcetype," & _
				' "                referenceno," & _
				' "                reference," & _
				' "                remarks," & _
				' "                chargejob," & _
				' "                chargetype," & _
				' "                chargecode," & _
				' "                ROUND ( (amount / Unit) * quantity, 0) amount" & _
				' "           FROM (  SELECT ag.fcba," & _
				' "                          TRUNC (fddate, 'mm') tdate," & _
				' "                          TO_CHAR (fddate, 'YYYY') accountingyear," & _
				' "                          TO_CHAR (fddate, 'MM') accountingperiod," & _
				' "                          ag.jobcode," & _
				' "                          'EM' Sourcetype," & _
				' "                          em.presentgrade referenceno," & _
				' "                          c.presentgradedesc reference," & _
				' "                          remarks," & _
				' "                          ag.chargejob," & _
				' "                          ag.chargetype," & _
				' "                          ag.chargecode," & _
				' "                          NVL (SUM (NVL (Mandays, 0)), 0) quantity," & _
				' "                          (SELECT NVL (SUM (NVL (Mandays, 0)), 0)" & _
				' "                             FROM attendance_gad" & _
				' "                            WHERE     fcba = '"& v_txtFcba &"'" & _
				' "                                  AND TO_NUMBER (TO_CHAR (fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
				' "                                  AND TO_NUMBER (TO_CHAR (fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
				' "                                  AND jobcode = '"& v_txtJobCode&"'" & _
				' "                                  AND attendance IN ('KJ', 'WH', 'WS'))" & _
				' "                             Unit," & _
				' "                          (SELECT NVL (SUM (NVL (Amount, 0)), 0)" & _
				' "                             FROM EmployeeCb" & _
				' "                            WHERE     fcba = '"& v_txtFcba &"'" & _
				' "                                  AND TO_NUMBER (TO_CHAR (transdate, 'mm')) = "& Cint(v_txtMonth) &"" & _
				' "                                  AND TO_NUMBER (TO_CHAR (transdate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
				' "                                  AND jobcode = '"& v_txtJobCode &"'" & _
				' "                                  AND SUBSTR (referenceno, 1, 3) NOT IN" & _
				' "                                         ('A45', 'A26')" & _
				' "                                  AND sourcetype = 'EM')" & _
				' "                             amount" & _
				' "                     FROM attendance_gad ag," & _
				' "                          employee em," & _
				' "                          (SELECT Seq, PresentGradeCode, PresentGradeDesc" & _
				' "                             FROM (SELECT 1 Seq," & _
				' "                                          'KHL' PresentGradeCode," & _
				' "                                          'KARYAWAN HARIAN LEPAS'" & _
				' "                                             PresentGradeDesc" & _
				' "                                     FROM DUAL" & _
				' "                                   UNION ALL" & _
				' "                                   SELECT 2 Seq," & _
				' "                                          'KHT' PresentGradeCode," & _
				' "                                          'KARYAWAN HARIAN TETAP'" & _
				' "                                             PresentGradeDesc" & _
				' "                                     FROM DUAL" & _
				' "                                   UNION ALL" & _
				' "                                   SELECT 3 Seq," & _
				' "                                          'KBT' PresentGradeCode," & _
				' "                                          'KARYAWAN BULANAN TETAP'" & _
				' "                                             PresentGradeDesc" & _
				' "                                     FROM DUAL)) c" & _
				' "                    WHERE     ag.fcba = em.fcba" & _
				' "                          AND ag.employeecode = em.fccode" & _
				' "                          AND em.PresentGrade = c.PresentGradeCode" & _
				' "                          AND ag.fcba = '"& v_txtFcba &"'" & _
				' "                          AND TO_NUMBER (TO_CHAR (fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
				' "                          AND TO_NUMBER (TO_CHAR (fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
				' "                          AND jobcode = '"& v_txtJobCode &"'" & _
				' "                          AND attendance IN ('KJ', 'WH', 'WS')" & _
				' "                 GROUP BY ag.fcba," & _
				' "                          TRUNC (fddate, 'mm')," & _
				' "                          TO_CHAR (fddate, 'YYYY')," & _
				' "                          TO_CHAR (fddate, 'MM')," & _
				' "                          ag.jobcode," & _
				' "                          remarks," & _
				' "                          em.presentgrade," & _
				' "                          c.presentgradedesc," & _
				' "                          ag.chargejob," & _
				' "                          ag.chargetype," & _
				' "                          ag.chargecode" & _
				' "                 UNION ALL" & _
				' "                   SELECT ag.fcba," & _
				' "                          TRUNC (fddate, 'mm') tdate," & _
				' "                          TO_CHAR (fddate, 'YYYY') accountingyear," & _
				' "                          TO_CHAR (fddate, 'MM') accountingperiod," & _
				' "                          ag.jobcode," & _
				' "                          'EM' Sourcetype," & _
				' "                          em.presentgrade referenceno," & _
				' "                          c.presentgradedesc reference," & _
				' "                          remarks," & _
				' "                          ag.chargejob," & _
				' "                          ag.chargetype," & _
				' "                          ag.chargecode," & _
				' "                          NVL (SUM (NVL (Unit, 0)), 0) quantity," & _
				' "                          (SELECT NVL (SUM (NVL (Unit, 0)), 0)" & _
				' "                             FROM attendance_gad" & _
				' "                            WHERE     fcba = '"& v_txtFcba &"'" & _
				' "                                  AND TO_NUMBER (TO_CHAR (fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
				' "                                  AND TO_NUMBER (TO_CHAR (fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
				' "                                  AND jobcode = '"& v_txtJobCode &"'" & _
				' "                                  AND attendance IN ('KB'))" & _
				' "                             Unit," & _
				' "                          (SELECT NVL (SUM (NVL (Amount, 0)), 0)" & _
				' "                             FROM EmployeeCb" & _
				' "                            WHERE     fcba = '"& v_txtFcba &"'" & _
				' "                                  AND TO_NUMBER (TO_CHAR (transdate, 'mm')) = "& Cint(v_txtMonth) &"" & _
				' "                                  AND TO_NUMBER (TO_CHAR (transdate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
				' "                                  AND jobcode = '"& v_txtJobCode &"'" & _
				' "                                  AND SUBSTR (referenceno, 1, 3) IN" & _
				' "                                         ('A45', 'A26')" & _
				' "                                  AND sourcetype = 'EM')" & _
				' "                             amount" & _
				' "                     FROM attendance_gad ag," & _
				' "                          employee em," & _
				' "                          (SELECT Seq, PresentGradeCode, PresentGradeDesc" & _
				' "                             FROM (SELECT 1 Seq," & _
				' "                                          'KHL' PresentGradeCode," & _
				' "                                          'KARYAWAN HARIAN LEPAS'" & _
				' "                                             PresentGradeDesc" & _
				' "                                     FROM DUAL" & _
				' "                                   UNION ALL" & _
				' "                                   SELECT 2 Seq," & _
				' "                                          'KHT' PresentGradeCode," & _
				' "                                          'KARYAWAN HARIAN TETAP'" & _
				' "                                             PresentGradeDesc" & _
				' "                                     FROM DUAL" & _
				' "                                   UNION ALL" & _
				' "                                   SELECT 3 Seq," & _
				' "                                          'KBT' PresentGradeCode," & _
				' "                                          'KARYAWAN BULANAN TETAP'" & _
				' "                                             PresentGradeDesc" & _
				' "                                     FROM DUAL)) c" & _
				' "                    WHERE     ag.fcba = em.fcba" & _
				' "                          AND ag.employeecode = em.fccode" & _
				' "                          AND em.PresentGrade = c.PresentGradeCode" & _
				' "                          AND ag.fcba = '"& v_txtFcba &"'" & _
				' "                          AND TO_NUMBER (TO_CHAR (fddate, 'mm')) = "& Cint(v_txtMonth) &"" & _
				' "                          AND TO_NUMBER (TO_CHAR (fddate, 'yyyy')) = "& Cint(v_txtYear) &"" & _
				' "                          AND jobcode = '"& v_txtJobCode &"'" & _
				' "                          AND attendance IN ('KB')" & _
				' "                 GROUP BY ag.fcba," & _
				' "                          TRUNC (fddate, 'mm')," & _
				' "                          TO_CHAR (fddate, 'YYYY')," & _
				' "                          TO_CHAR (fddate, 'MM')," & _
				' "                          ag.jobcode," & _
				' "                          remarks," & _
				' "                          em.presentgrade," & _
				' "                          c.presentgradedesc," & _
				' "                          ag.chargejob," & _
				' "                          ag.chargetype," & _
				' "                          ag.chargecode)) dataca" & _
				' "  WHERE     dataca.sourcetype IN ('EM', 'ES')" & _
				' "        AND dataca.fcba = '"& v_txtFcba &"'" & _
				' "        AND dataca.jobcode = '"& v_txtJobCode &"'" 

				' " select fcba" & _
				' "        , tdate" & _
				' "        , accountingyear " & _
    '             "        , accountingperiod " & _
				' "        , jobcode" & _
				' "        , sourcetype" & _
				' "        , referenceno" & _
				' "        , reference" & _
				' "        , remarks" & _
				' "        , amount" & _
				' " from (" & _
				' "     select ec.fcba" & _
				' "            , trunc(transdate,'mm') tdate" & _
			 '    "            , to_char(transdate,'YYYY') accountingyear" & _
    '             "            , to_char(transdate,'MM')   accountingperiod" & _
				' "            , jobcode" & _
				' "            , sourcetype" & _
				' "            , em.presentgrade referenceno" & _
				' "            , c.PresentGradeDesc reference       " & _
				' "            , 'ALOKASI SUPERVISI' remarks" & _
				' "            , round(nvl(sum(nvl(amount, 0)), 0), 0) amount" & _
				' "     from employeecb ec" & _
				' "             , employee em" & _
				' "             , (Select Seq, PresentGradeCode,  PresentGradeDesc From(" & _
				' "                 Select 1 Seq, 'KHL' PresentGradeCode, 'KARYAWAN HARIAN LEPAS' PresentGradeDesc From Dual" & _
				' "                 Union All" & _
				' "                 Select 2 Seq, 'KHT' PresentGradeCode, 'KARYAWAN HARIAN TETAP' PresentGradeDesc From Dual" & _
				' "                 Union All" & _
				' "                 Select 3 Seq, 'KBT' PresentGradeCode, 'KARYAWAN BULANAN TETAP' PresentGradeDesc From Dual" & _
				' "                )) c" & _
				' "     where ec.fcba = em.fcba" & _
				' "               and ec.empcode = em.fccode" & _
				' "               and em.PresentGrade = c.PresentGradeCode" & _
				' "               and ec.fcba = '" & v_txtFcba & "'" & _
				' "               and to_number(to_char(transdate, 'mm')) = " & Cint(v_txtMonth) & "" & _
				' "               and to_number(to_char(transdate, 'yyyy')) = " & Cint(v_txtYear) & "" & _
				' "               and jobcode = '" & v_txtJobCode & "'" & _
				' "               and sourcetype = 'ES'" & _
				' "     group by ec.fcba" & _
				' "              , trunc(transdate,'mm')" & _
				' "              , to_char(transdate,'YYYY') " & _
    '             "              , to_char(transdate,'MM') " & _
				' "              , em.presentgrade" & _
				' "              , c.PresentGradeDesc" & _
				' "              , jobcode" & _
				' "              , sourcetype" & _
				' "     Union All" & _
				' "     Select fcba" & _
				' "            , tdate" & _
				' "            , accountingyear" & _
    '             "            , accountingyear" & _
				' "            , jobcode" & _
				' "            , sourcetype" & _
				' "            , referenceno" & _
				' "            , reference" & _
				' "            , remarks" & _
				' "            , round((amount/Unit)*quantity, 0) amount" & _
				' "     from (" & _
				' "         Select " & _
				' "             ag.fcba" & _
				' "             , trunc(fddate, 'mm') tdate" & _
				' "             , to_char(fddate,'YYYY') accountingyear" & _
    '             "             , to_char(fddate,'MM')   accountingperiod" & _
				' "             , ag.jobcode" & _
				' "             , 'EM' Sourcetype" & _
				' "             , em.presentgrade referenceno" & _
				' "             , c.presentgradedesc reference" & _
				' "             , remarks" & _
				' "             , nvl(sum(nvl(Mandays, 0)), 0) quantity" & _
				' "             , (Select nvl(Sum(nvl(Mandays, 0)), 0) From attendance_gad " & _
				' "                 where fcba = '" & v_txtFcba & "'" & _
				' "                           and to_number(to_char(fddate, 'mm')) = " & Cint(v_txtMonth) & "" & _
				' "                           and to_number(to_char(fddate, 'yyyy')) = " & Cint(v_txtYear) & "" & _
				' "                           and jobcode = '" & v_txtJobCode & "'" & _
				' "                           and attendance in ('KJ', 'WH', 'WS')) Unit" & _
				' "             , (Select nvl(Sum(nvl(Amount, 0)), 0) From EmployeeCb " & _
				' "                 where fcba = '" & v_txtFcba & "'" & _
				' "                           and to_number(to_char(transdate, 'mm')) = " & Cint(v_txtMonth) & "" & _
				' "                           and to_number(to_char(transdate, 'yyyy')) = " & Cint(v_txtYear) & "" & _
				' "                           and jobcode = '" & v_txtJobCode & "'" & _
				' "                           and substr(referenceno, 1, 3) not in ('A45', 'A26') and sourcetype = 'EM') amount" & _
				' "         from attendance_gad ag" & _
				' "                 , employee em" & _
				' "                 , (Select Seq, PresentGradeCode,  PresentGradeDesc From(" & _
				' "                     Select 1 Seq, 'KHL' PresentGradeCode, 'KARYAWAN HARIAN LEPAS' PresentGradeDesc From Dual" & _
				' "                     Union All" & _
				' "                     Select 2 Seq, 'KHT' PresentGradeCode, 'KARYAWAN HARIAN TETAP' PresentGradeDesc From Dual" & _
				' "                     Union All" & _
				' "                     Select 3 Seq, 'KBT' PresentGradeCode, 'KARYAWAN BULANAN TETAP' PresentGradeDesc From Dual" & _
				' "                    )) c" & _
				' "         where ag.fcba = em.fcba" & _
				' "                   and ag.employeecode = em.fccode" & _
				' "                   and em.PresentGrade = c.PresentGradeCode" & _
				' "                   and ag.fcba = '" & v_txtFcba & "'" & _
				' "                   and to_number(to_char(fddate, 'mm')) = " & Cint(v_txtMonth) & "" & _
				' "                   and to_number(to_char(fddate, 'yyyy')) = " & Cint(v_txtYear) & "" & _
				' "                   and jobcode = '" & v_txtJobCode & "'" & _
				' "                   and attendance in ('KJ', 'WH', 'WS')" & _
				' "         Group by ag.fcba" & _
				' "                  , trunc(fddate, 'mm')" & _
				' "             , to_char(fddate,'YYYY')" & _
    '             "             , to_char(fddate,'MM')" & _
				' "                  , ag.jobcode" & _
				' "                  , remarks" & _
				' "                  , em.presentgrade" & _
				' "                  , c.presentgradedesc" & _
				' "         Union All" & _
				' "         Select " & _
				' "             ag.fcba" & _
				' "             , trunc(fddate, 'mm') tdate" & _
				' "             , to_char(fddate,'YYYY') accountingyear" & _
    '             "             , to_char(fddate,'MM')   accountingperiod" & _
				' "             , ag.jobcode" & _
				' "             , 'EM' Sourcetype" & _
				' "             , em.presentgrade referenceno" & _
				' "             , c.presentgradedesc reference" & _
				' "             , remarks" & _
				' "             , nvl(sum(nvl(Unit, 0)), 0) quantity" & _
				' "             , (Select nvl(Sum(nvl(Unit, 0)), 0) From attendance_gad " & _
				' "                 where fcba = '" & v_txtFcba & "'" & _
				' "                           and to_number(to_char(fddate, 'mm')) = " & Cint(v_txtMonth) & "" & _
				' "                           and to_number(to_char(fddate, 'yyyy')) = " & Cint(v_txtYear) & "" & _
				' "                           and jobcode = '" & v_txtJobCode & "'" & _
				' "                           and attendance in ('KB')) Unit" & _
				' "             , (Select nvl(Sum(nvl(Amount, 0)), 0) From EmployeeCb " & _
				' "                 where fcba = '" & v_txtFcba & "'" & _
				' "                           and to_number(to_char(transdate, 'mm')) = " & Cint(v_txtMonth) & "" & _
				' "                           and to_number(to_char(transdate, 'yyyy')) = " & Cint(v_txtYear) & "" & _
				' "                           and jobcode = '" & v_txtJobCode & "'" & _
				' "                           and substr(referenceno, 1, 3) in ('A45', 'A26') and sourcetype = 'EM') amount" & _
				' "         from attendance_gad ag" & _
				' "                 , employee em" & _
				' "                 , (Select Seq, PresentGradeCode,  PresentGradeDesc From(" & _
				' "                     Select 1 Seq, 'KHL' PresentGradeCode, 'KARYAWAN HARIAN LEPAS' PresentGradeDesc From Dual" & _
				' "                     Union All" & _
				' "                     Select 2 Seq, 'KHT' PresentGradeCode, 'KARYAWAN HARIAN TETAP' PresentGradeDesc From Dual" & _
				' "                     Union All" & _
				' "                     Select 3 Seq, 'KBT' PresentGradeCode, 'KARYAWAN BULANAN TETAP' PresentGradeDesc From Dual" & _
				' "                    )) c" & _
				' "         where ag.fcba = em.fcba" & _
				' "                   and ag.employeecode = em.fccode" & _
				' "                   and em.PresentGrade = c.PresentGradeCode" & _
				' "                   and ag.fcba = '" & v_txtFcba & "'" & _
				' "                   and to_number(to_char(fddate, 'mm')) = " & Cint(v_txtMonth) & "" & _
				' "                   and to_number(to_char(fddate, 'yyyy')) = " & Cint(v_txtYear) & "" & _
				' "                   and jobcode = '" & v_txtJobCode & "'" & _
				' "                   and attendance in ('KB')" & _
				' "         Group by ag.fcba" & _
				' "                  , trunc(fddate, 'mm')" & _
				' "                  , to_char(fddate,'YYYY') " & _
    '             "                  , to_char(fddate,'MM') " & _
				' "                  , ag.jobcode" & _
				' "                  , remarks" & _
				' "                  , em.presentgrade" & _
				' "                  , c.presentgradedesc" & _
				' "     )" & _
				' " ) dataca" & _
				' " where dataca.sourcetype in ('EM', 'ES') " & _
				' "       and dataca.fcba = '" & v_txtFcba & "' and dataca.jobcode = '" & v_txtJobCode & "'"
				
	elseif rsSourceType("SOURCETYPE") = Cstr("WC") then
	strCA	=	" select fcba" & _
				"        , tdate" & _
				"        , accountingyear " & _
                "        , accountingperiod " & _
				"        , jobcode" & _
				"        , sourcetype" & _
				"        , referenceno" & _
				"        , reference" & _
				"        , remarks" & _
				"        , amount" & _
				" from (" & _
				"     select wip.fcba" & _
				"            , wip.wipdate tdate" & _
				"            , to_char(wip.wipdate,'YYYY') accountingyear " & _
                "            , to_char(wip.wipdate,'MM') accountingperiod " & _
				"            , ad.job jobcode" & _
				"            , 'WC' sourcetype" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba) reference" & _
				"            , wip.agreementcode referenceno" & _
				"            , wip.remarks" & _
				"            , nvl(sum(nvl(wip.qty, 0)), 0) quantity" & _
				"            , round(nvl(sum(nvl(ad.rate*wip.qty, 0)), 0), 0) amount" & _
				"     from workinprogress wip" & _
				"          , agreementdetail ad" & _
				"     where wip.fcba = ad.fcba" & _
				"           and wip.agreementcode = ad.agreementcode" & _
				"           and wip.agreementlineno = ad.lineno" & _
				"     group by wip.fcba" & _
				"            , wip.wipdate" & _
				"            , to_char(wip.wipdate,'YYYY') " & _
                "            , to_char(wip.wipdate,'MM') " & _
				"            , ad.job" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba)" & _
				"            , wip.agreementcode" & _
				"            , wip.remarks" & _
				"     union all" & _
				"     select wip.fcba" & _
				"            , wip.wipdate tdate" & _
				"            , to_char(wip.wipdate,'YYYY') accountingyear " & _
                "            , to_char(wip.wipdate,'MM') accountingperiod " & _
				"            , Contractor_api.get_controljob(ad.contractorcode, wip.fcba)  jobcode" & _
				"            , 'WC' sourcetype" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba) reference" & _
				"            , wip.agreementcode referenceno" & _
				"            , wip.remarks" & _
				"            , 0 quantity" & _
				"            , round(nvl(sum(nvl(ad.rate*wip.qty, 0)), 0), 0)* -1 amount" & _
				"     from workinprogress wip" & _
				"          , agreementdetail ad" & _
				"     where wip.fcba = ad.fcba" & _
				"           and wip.agreementcode = ad.agreementcode" & _
				"           and wip.agreementlineno = ad.lineno" & _
				"     group by wip.fcba" & _
				"            , wip.wipdate" & _
				"            , to_char(wip.wipdate,'YYYY') " & _
                "            , to_char(wip.wipdate,'MM') " & _
				"            , Contractor_api.get_controljob(ad.contractorcode, wip.fcba)" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba)" & _
				"            , wip.agreementcode" & _
				"            , wip.remarks" & _
				" ) dataca" & _
				" where dataca.sourcetype in ('WC') " & _
				"       and dataca.fcba = '" & v_txtFcba & "' and dataca.jobcode = '" & v_txtJobCode & "'"

	elseif rsSourceType("SOURCETYPE") = Cstr("TT") then
	strCA	=	" SELECT referenceno," & _
				"        chargejob," & _
				"        chargetype," & _
				"        chargecode," & _
				"        tdate," & _
				"        fcba," & _
				"        accountingyear," & _
				"        accountingperiod," & _
				"        remarks," & _
				"        TITLE_REFERENCE_01," & _
				"        TITLE_REFERENCE_02," & _
				"        TITLE_REFERENCE_03," & _
				"        TITLE_REFERENCE_04," & _
				"        TITLE_REFERENCE_05," & _
				"        AUTO_REFERENCE_01," & _
				"        AUTO_REFERENCE_02," & _
				"        AUTO_REFERENCE_03," & _
				"        AUTO_REFERENCE_04," & _
				"        AUTO_REFERENCE_05," & _
				"        jobcode," & _
				"        (stockcode||' : '||stockname) reference," & _
				"        amount," & _
				"        sourcetype FROM (" & _
				"   SELECT tt.orderno referenceno," & _
				"          tt.orderdate tdate," & _
				"          '' chargejob," & _
				"          '' chargetype," & _
				"          '' chargecode," & _
				"          '' TITLE_REFERENCE_01," & _
				"          '' TITLE_REFERENCE_02," & _
				"          '' TITLE_REFERENCE_03," & _
				"          '' TITLE_REFERENCE_04," & _
				"          '' TITLE_REFERENCE_05," & _
				"          '' AUTO_REFERENCE_01," & _
				"          '' AUTO_REFERENCE_02," & _
				"          '' AUTO_REFERENCE_03," & _
				"          '' AUTO_REFERENCE_04," & _
				"          '' AUTO_REFERENCE_05," & _
				"          TO_CHAR (tt.orderdate, 'mm') accountingperiod," & _
				"          TO_CHAR (tt.orderdate, 'yyyy') accountingyear," & _
				"          tt.remarks," & _
				"          td.jobcode," & _
				"          td.fcba," & _
				"          td.stockcode," & _
				"          stock_api.get_stockname(td.stockcode, td.fcba) stockname," & _
				"          td.amount," & _
				"          'TT' sourcetype" & _
				"     FROM transporttaskdtl_tab td" & _
				"          INNER JOIN transporttask_tab tt" & _
				"             ON (tt.orderno = td.orderno AND tt.fcba = td.fcba)" & _
				"    WHERE     td.fcba = '"& v_txtFcba &"'" & _
				"          AND td.jobcode = '"& v_txtJobCode &"' ) dataca" & _
				"  WHERE fcba = '"& v_txtFcba &"'" & _
				"  AND jobcode = '"& v_txtJobCode &"'"

	elseif rsSourceType("SOURCETYPE") = Cstr("RT") then
		strCA	=	" select fcba" & _
				"        , tdate" & _
				"        , accountingyear " & _
                "        , accountingperiod " & _
				"        , jobcode" & _
				"        , sourcetype" & _
				"        , referenceno" & _
				"        , reference" & _
				"        , remarks" & _
				"        , amount" & _
				" from (" & _
				"     select wip.fcba" & _
				"            , wip.wipdate tdate" & _
				"            , to_char(wip.wipdate,'YYYY') accountingyear " & _
                "            , to_char(wip.wipdate,'MM') accountingperiod " & _
				"            , ad.job jobcode" & _
				"            , 'WC' sourcetype" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba) reference" & _
				"            , wip.agreementcode referenceno" & _
				"            , wip.remarks" & _
				"            , nvl(sum(nvl(wip.qty, 0)), 0) quantity" & _
				"            , round(nvl(sum(nvl(ad.rate*wip.qty, 0)), 0), 0) amount" & _
				"     from workinprogress wip" & _
				"          , agreementdetail ad" & _
				"     where wip.fcba = ad.fcba" & _
				"           and wip.agreementcode = ad.agreementcode" & _
				"           and wip.agreementlineno = ad.lineno" & _
				"     group by wip.fcba" & _
				"            , wip.wipdate" & _
				"            , to_char(wip.wipdate,'YYYY') " & _
                "            , to_char(wip.wipdate,'MM') " & _
				"            , ad.job" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba)" & _
				"            , wip.agreementcode" & _
				"            , wip.remarks" & _
				"     union all" & _
				"     select wip.fcba" & _
				"            , wip.wipdate tdate" & _
				"            , to_char(wip.wipdate,'YYYY') accountingyear " & _
                "            , to_char(wip.wipdate,'MM') accountingperiod " & _
				"            , Contractor_api.get_controljob(ad.contractorcode, wip.fcba)  jobcode" & _
				"            , 'WC' sourcetype" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba) reference" & _
				"            , wip.agreementcode referenceno" & _
				"            , wip.remarks" & _
				"            , 0 quantity" & _
				"            , round(nvl(sum(nvl(ad.rate*wip.qty, 0)), 0), 0)* -1 amount" & _
				"     from workinprogress wip" & _
				"          , agreementdetail ad" & _
				"     where wip.fcba = ad.fcba" & _
				"           and wip.agreementcode = ad.agreementcode" & _
				"           and wip.agreementlineno = ad.lineno" & _
				"     group by wip.fcba" & _
				"            , wip.wipdate" & _
				"            , to_char(wip.wipdate,'YYYY') " & _
                "            , to_char(wip.wipdate,'MM') " & _
				"            , Contractor_api.get_controljob(ad.contractorcode, wip.fcba)" & _
				"            , contractor_api.Get_ContractorName(ad.contractorcode, wip.fcba)" & _
				"            , wip.agreementcode" & _
				"            , wip.remarks" & _
				" ) dataca" & _
				" where dataca.sourcetype in ('WC') " & _
				"       and dataca.fcba = '" & v_txtFcba & "' and dataca.jobcode = '" & v_txtJobCode & "'"
	else	
	strCA	=	" SELECT * FROM" & _
				" (SELECT" & _
				" 'DUMMY' FCBA," & _
				" TO_DATE('11/1/2012','MM/DD/YYYY') TDATE," & _
				"  2012 accountingyear, " & _
                "  11  accountingperiod, " & _
				" 'DUMMY' JOBCODE," & _
				" 'DUMMY' SOURCETYPE," & _
				" 'DUMMY' REFERENCENO," & _
				" 'DUMMY' REFERENCENO," & _
				" 'DUMMY' REMARKS," & _
				" 0 QUANTITY," & _
				" round(0,2) AMOUNT" & _
				" FROM DUAL" & _
				" ) DATACA" & _
				" WHERE" & _
				" DATACA.FCBA = '" & v_txtFcba & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	end if
'response.write(strCA)	
	if  LEN(TRIM(v_txtYear)) > 0 then
		strCA	=	strCA & " AND dataca.accountingyear =" & Cint(v_txtYear)
	end if
	if  LEN(TRIM(v_txtMonth)) > 0 then
		strCA	=	strCA &  " AND dataca.accountingperiod =" & Cint(v_txtMonth)
	end if
	'response.write strCA
	'strCA	=	strCA & " ORDER BY  DATACA.REFERENCENO"
	strCA	=	strCA & " ORDER BY DATACA.TDATE, DATACA.REFERENCENO"
Set rsCA = Server.CreateObject("ADODB.Recordset")
 'response.write(strCA)
rsCA.Open strCA, connstr, adOpenStatic 
%>
<%if rsCA.bof and rsCA.eof then%>
	<tr><td align="left" colspan="15">NO DATA AVAILABLE</td></tr>
<%else
	if Not rsCA.eof then
	NO = 1
	nSubTotal = 0
	while not rsCA.eof
	if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
	%>
	<tr class="<%=cClass%>">
		<td align="left"><%=NO%></td>
		<td align="left"><%=rsCA("TDATE")%></td>
		<td align="left"><%=rsCA("REFERENCENO")%></td>
		<td align="left"><%=rsCA("REFERENCE")%></td>
		<td align="left"><%=rsCA("REMARKS")%></td>	
		<% if trim(rsCA("CHARGEJOB")) <> "" then %> 
			<td align="left"><%=rsCA("CHARGEJOB")%> </td>	
			<td align="left"><%=rsCA("CHARGETYPE")%></td>	
			<td align="left"><%=rsCA("CHARGECODE")%></td>		
		<%else%>
			<td align="left"></td>	
			<td align="left"></td>	
			<td align="left"></td>
		<%end if%>	
		<%if trim(rsCA("AMOUNT")) <> "" then %>
			<%if Cdbl(rsCA("AMOUNT")) > 0 then%>
			<td align="right"><%=FormatNumber(rsCA("AMOUNT"),2)%></td>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<%elseif Cdbl(rsCA("AMOUNT")) < 0 then%>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<td align="right"><%=FormatNumber(rsCA("AMOUNT"),2)%></td>
			<%else%>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<%end if%>	
		<%else%>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<%end if%>	
	<% if rsSourceType("SOURCETYPE") = Cstr("EM") OR rsSourceType("SOURCETYPE") = Cstr("BA") OR rsSourceType("SOURCETYPE") = Cstr("PR") OR rsSourceType("SOURCETYPE") = Cstr("WS") OR rsSourceType("SOURCETYPE") = Cstr("VH") OR rsSourceType("SOURCETYPE") = Cstr("TT") OR rsSourceType("SOURCETYPE") = Cstr("RT") then %>
		<td align="left"><%=rsCA("TITLE_REFERENCE_01")%>&nbsp;<%=rsCA("AUTO_REFERENCE_01")%></td>	
		<td align="left"><%=rsCA("TITLE_REFERENCE_02")%>&nbsp;<%=rsCA("AUTO_REFERENCE_02")%></td>	
		<td align="left"><%=rsCA("TITLE_REFERENCE_03")%>&nbsp;<%=rsCA("AUTO_REFERENCE_03")%></td>	
		<td align="left"><%=rsCA("TITLE_REFERENCE_04")%>&nbsp;<%=rsCA("AUTO_REFERENCE_04")%></td>	
		<td align="left"><%=rsCA("TITLE_REFERENCE_05")%>&nbsp;<%=rsCA("AUTO_REFERENCE_05")%></td>	
	<%else%>
		<td align="left"></td>
		<td align="left"></td>
		<td align="left"></td>
		<td align="left"></td>
		<td align="left"></td>
	<%end if%>	
	</tr>
<%
		NO = NO + 1 

		if rsCA("AMOUNT") <> "" then 
			nSubTotal = nSubTotal + Cdbl(rsCA("AMOUNT"))
		else
			nSubTotal = nSubTotal + 0
		end if			
		
		rsCA.MoveNext
	Wend
	nGrandTotal = nGrandTotal + nSubTotal	
%>
	<tr bgcolor="#9CB770">
		<td align="center" colspan="8"><strong><i>Sub Total</i></strong></td>
		<td align="center" colspan="5"><strong><i><%=FormatNumber(nSubTotal,2)%></i></strong></td>
		<td align="center" colspan="8"></td>
	</tr>
<%
end if
	rsCA.Close
	set rsCA = Nothing
end if
		rsSourceType.MoveNext
	Wend
	rsSourceType.Close
	set rsSourceType = Nothing
end if
%>
	<tr bgcolor="#9CB770">
		<td align="center" colspan="8"><strong><i>Grand Total</i></strong></td>
		<td align="center" colspan="5"><strong><i><%=FormatNumber(nGrandTotal,2)%></i></strong></td>
		<td align="center" colspan="8"></td>
	</tr>
  </tbody>
</table>
<!--#include file="../Bottom_Report.asp"-->