<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
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
<title>RW Plant Reporting - Cost Activity</title>
</head>
<body>
<div id="wrapper">
	<!--#include file="print_excel.asp"-->
	</div>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF;"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF;"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF;"

v_txtyear=TRIM(request("txtyear"))
v_txtmonth=TRIM(request("txtmonth"))
v_txtjobgroup=TRIM(request("txtjobgroup"))
v_txtjob=TRIM(request("txtjob"))
v_txtST=TRIM(request("txtST"))
v_txtSC=TRIM(request("txtSC"))
v_txtTT=TRIM(request("txtTT"))
v_txtTC=TRIM(request("txtTC"))
v_txtReference=TRIM(request("txtReference"))
v_txtReference01=TRIM(request("txtReference01"))
v_txtReference02=TRIM(request("txtReference02"))
v_txtReference03=TRIM(request("txtReference03"))
v_txtReference04=TRIM(request("txtReference04"))
v_txtReference05=TRIM(request("txtReference05"))
v_txtDebitCredit=TRIM(request("txtDebitCredit"))
tdate_from=TRIM(request("tdate_from"))
tdate_to=TRIM(request("tdate_to"))

txtSQL = " SELECT docdate Tdate,a.jobcode,c.fcname,b.sourcecode source,decode(amountdebit,0,'Cr','Db') DEBETCREDIT, " & _ 
         " Locationtype TARGETTYPE, Locationcode TARGETCODE," & _ 
		 " bkmno REFERENCENO,amountdebit+amountkredit as AMOUNT,nvl(unit,0) as QUANTITY, nvl(OUTPUT,0) as OUTPUT," & _ 
		 " decode(nvl(a.chargeunit,'-'),'-',a.description,a.description||' Charge:'||a.chargeunit||'-'||a.chargejob||'-'||a.chargecode) AUTO_REFERENCE" & _ 
		 " FROM bkmjournaldtl a " & _
		 " INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba)" & _
		 " INNER JOIN jobmaster c ON (a.jobcode=c.fccode)" & _
		 " WHERE b.postingflg='1' and a.fcba='" & session("fcba") & "' " 
		 
		 
if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " and extract (year from docdate) = " & v_txtyear
	txtSelect = txtSelect & " Year : " & v_txtyear & " |"
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL =txtSQL &  " and extract (month from docdate)  = " & v_txtmonth
	txtSelect = txtSelect & " Month : " & v_txtmonth & " |"
end if
if  len(trim(tdate_from)) > 0 then
    txtSQL = txtSQL & " and docdate >= to_date('"& tdate_from &"','mm/dd/yyyy')" 
	txtSelect = txtSelect & " From " & tdate_from & " |"
end if
if  len(trim(tdate_to)) > 0 then
    txtSQL =txtSQL &  " and docdate <= to_date('"& tdate_to &"','mm/dd/yyyy')"
	txtSelect = txtSelect & " To " & tdate_to & " |"
end if
if  len(trim(v_txtjobgroup)) > 0 then
    txtSQL = txtSQL & " and a.JobCode in (select fccode from jobmaster where jobGROUPCODE=" & v_txtjobgroup & ")"
	txtSelect = txtSelect & " Job Group Code : " & v_txtjobgroup & "|"
end if
if  len(trim(v_txtjob)) > 0 then
    txtSQL =txtSQL &  " and a.JobCode='" & v_txtjob & "'"
	txtSelect = txtSelect & " Job Code : " & v_txtjob & " |"
end if 

txtsql=(txtsql) & " Order by a.jobcode,docdate,SOURCE,a.docno "

'response.Write(txtsql)
response.Flush()
%>
<!--
 <table width="471">
  <tr> 
    <td rowspan="5" width="100"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">REPORT</td><td  width="300"><b><FONT color="#000099">ACTIVITY BASE COSTING</FONT></td></tr>
 </table>
 -->
<table border="0" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="19" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>ACTIVITY BASE COSTING</h2>
				</div>
			</div>
		</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="19" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
                        
<!--                    
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	<tr>
    <th colspan="13" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" >Selected by : <%=txtSelect%></th>
	</tr>-->
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	
	<tr>
	     <th style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</th>
		<th align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Date</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Code</th>
		<th align="center" style="<%=cmdstyleTOP%> min-width:200px;"> <font face="Verdana" size="1">Activity</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Debet/Credit</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Cost Element</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TT</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TC</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Referenceno</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Qty</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Output</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Amount</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Description</th>
		/tr>
    </thead>
    <tbody>
	<tr>
		<col span="13">
		<col style="background-color:#fffea5">
	</tr>
	<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo
	v_Opening= 0
	v_Debit  = 0
	v_Credit = 0
	v_Saldo  = 0
	While Not rsCat.EOF
	    
	    %>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TDATE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOBCODE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DEBETCREDIT")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("source")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TARGETTYPE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TARGETCODE")%></td>
		<%If rscat("REFERENCENO")="" or isnull(rscat("REFERENCENO"))  THEN%>
		     <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1">&nbsp;</td>
		<%Else%>
		     <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REFERENCENO")%></td>
		<%End if%>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QUANTITY"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OUTPUT"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("AMOUNT"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("AUTO_REFERENCE")%></td>
		
		</tr>
	  <%
	  v_Saldo  =v_Saldo+trim(cdbl(rsCat("AMOUNT")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="11"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Saldo,2)%></td>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="1"> <font face="Verdana" size="1"></td>
	</tr>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->