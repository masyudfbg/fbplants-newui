<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>iPS Reporting - Trial Balance</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
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
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtyear=TRIM(request("txtyear"))
v_txtmonth=TRIM(request("txtmonth"))
v_txtjobgroup=TRIM(request("txtjobgroup"))
v_txtjob=TRIM(request("txtjob"))
v_txtnonull=request("no-null")

txtSQL = "select FCCODE,fcname,OPENING_BALANCE,DEBET,CREDIT,Saldo from(" & _
		 "Select sl.FCCODE, s.fcname,nvl(sum(sl.OPENING_BALANCE),0) as OPENING_BALANCE,nvl(sum(sl.DEBET),0) as DEBET,nvl(sum(sl.CREDIT),0) as CREDIT, " & _
         "nvl((sum(sl.OPENING_BALANCE)+sum(sl.DEBET)+sum(sl.CREDIT)),0) AS Saldo " & _
		  " from TRIALBALANCE SL,job S where s.fccode=sl.FCCODE and s.fcba=sl.fcba  and sl.fcba='" & session("fcBa") & "'"
if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " and PERIOD_YEAR=" & v_txtyear
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL =txtSQL &  " and PERIOD_MONTH=" & v_txtmonth
end if
if  len(trim(v_txtjobgroup)) > 0 then
    txtSQL = txtSQL & " and sl.FCCODE in (select fccode from job where jobGROUPCODE=" & v_txtjobgroup & ")"
end if
if  len(trim(v_txtjob)) > 0 then
    txtSQL =txtSQL &  " and sl.FCCODE='" & v_txtjob & "'"
end if 
if v_txtnonull = "full" then 
txtsql=(txtsql) & " group by sl.FCCODE, s.FCNAME) order by FCCODE"
elseif v_txtnonull = "no-null" then
txtsql=(txtsql) & " group by sl.FCCODE, s.FCNAME) WHERE" & _
				  " (OPENING_BALANCE <> 0 OR DEBET <> 0 OR CREDIT <> 0 OR Saldo <> 0) order by FCCODE"
elseif v_txtnonull = "" then
txtsql=(txtsql) & " group by sl.FCCODE, s.FCNAME) order by FCCODE"
end if
'response.Write(txtsql)
response.Flush()
%>
<div id="wrapper">
	<!--#include file="print_excel.asp"-->
	</div>
<div id="header-wrapper">
	<!--#include file="report/menutop.asp"-->
</div>
<!-- <table width="575">
  <tr> 
    <td rowspan="5" width="99"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="123">Company</td>
	<td  width="337"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="123">Business Unit</td>
	<td  width="337"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="123">Manager</td>
	<td  width="337"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="123">REPORT</td>
	<td  width="337"><b><FONT color="#000099">TRIAL BALANCE</FONT></td>
	</tr>
 </table>
<BR><BR>
<table Cellpadding="2" cellspacing="0" id="product-table">-->
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>TRIAL BALANCE</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	
	</div>
	<%=v_txtnonull%>
	<form name="formnull" method="post" action="p_GLMasterRegisterMonthly.asp?txtKind=TrialBalance" >
	<select name="no-null" style="float:right;margin-top:-30px;width:100px;" onChange="this.form.submit();">
	   <% if v_txtnonull = "full" then %>
		<option value="full" selected>Detail</option>
		<option value="no-null">No Detail</option>
		<%elseif v_txtnonull = "no-null" then %>
		<option value="full">Detail</option>
		<option value="no-null" selected>No Detail</option>
		<%else%>
		<option value="full" selected>Detail</option>
		<option value="no-null">No Detail</option>
		<%end if%>
	</select>
	<input type="hidden" name="txtyear" value="<%=v_txtyear%>">
	<input type="hidden" name="txtmonth" value="<%=v_txtmonth%>">
	<input type="hidden" name="txtjobgroup" value="<%=v_txtjobgroup%>">
	<input type="hidden" name="txtjob" value="<%=v_txtjob%>">
	</form>
</th>
  </tr>
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField = rsCat.fields.count - 1
	end if
	NO=1%>
	</tr>
	<tr>
	    <th style="color:#FFFFFF">NO</th>
		<th style="color:#FFFFFF">Account<br>Code</th>
		<th style="color:#FFFFFF">Account<br>Description</th>
		<th style="color:#FFFFFF">Opening<br>Balance</th>
		<th style="color:#FFFFFF">Debet</th>
		<th style="color:#FFFFFF">Credit</th>
		<th style="color:#FFFFFF">Saldo</th>
	</tr>
</thead>
<tbody>
	<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo
	v_Opening= 0
	v_Debit  = 0
	v_Credit = 0
	v_Saldo  = 0
	While Not rsCat.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if	
%>
		<tr class="<%=cClass%>">
		<td align="Center" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCCODE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCNAME")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("OPENING_BALANCE"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Debet"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Credit"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Saldo"),2)%></td>
		</tr>
	  <%
	  v_Opening=v_Opening+cdbl(rsCat("OPENING_BALANCE"))
	  v_Debit  =v_Debit+(cdbl(rsCat("Debet")))
	  v_Credit =v_Credit+(cdbl(rsCat("Credit")))
	  v_Saldo  =v_Saldo+trim((rsCat("Saldo")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="3"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Opening,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Debit,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Credit,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Saldo,2)%></td>
	</tr>
 <tbody>
</table>
<!--#include file="Bottom_Report.asp"-->