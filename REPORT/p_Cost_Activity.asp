<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
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
v_txtST=TRIM(request("txtST"))
v_txtSC=TRIM(request("txtSC"))
v_txtTT=TRIM(request("txtTT"))
v_txtTC=TRIM(request("txtTC"))
v_txtReference=TRIM(request("txtReference"))
tdate_from=TRIM(request("tdate_from"))
tdate_to=TRIM(request("tdate_to"))

txtSQL = " Select TDATE, ca.JOBCODE, s.fcname,SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
         " TARGETTYPE, TARGETCODE, REFERENCENO, nvl(sum(AMOUNT),0) as AMOUNT, nvl(sum(QUANTITY),0) as QUANTITY, nvl(sum(OUTPUT),0) as OUTPUT " & _
		 " from COST_ACTIVITY CA,job S where CA.jobcode=s.fccode and ca.fcba=s.fcba " & _
		 " and ca.fcBA='" & session("fcBA") & "'"
if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " and to_number(to_char(TDATE,'yyyy')) = " & v_txtyear
	txtSelect = txtSelect & " Year " & v_txtyear
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL =txtSQL &  " and to_number(to_char(TDATE,'mm')) = " & v_txtmonth
	txtSelect = txtSelect & " Month " & v_txtmonth
end if
if  len(trim(tdate_from)) > 0 then
    txtSQL = txtSQL & " and TDATE => to_date('"& tdate_from &"','mm/dd/yyyy')" 
	txtSelect = txtSelect & " From " & tdate_from
end if
if  len(trim(tdate_to)) > 0 then
    txtSQL =txtSQL &  " and TDATE <= to_date('"& tdate_to &"','mm/dd/yyyy')"
	txtSelect = txtSelect & " To " & tdate_to
end if
if  len(trim(v_txtjobgroup)) > 0 then
    txtSQL = txtSQL & " and CA.JobCode in (select fccode from job where jobGROUPCODE=" & v_txtjobgroup & ")"
	txtSelect = txtSelect & " Job Group Code " & v_txtjobgroup
end if
if  len(trim(v_txtjob)) > 0 then
    txtSQL =txtSQL &  " and CA.JobCode='" & v_txtjob & "'"
	txtSelect = txtSelect & " Job Code " & v_txtjob
end if 
if  len(trim(v_txtST)) > 0 then
    txtSQL =txtSQL &  " and CA.SourceType='" & v_txtST & "'"
	txtSelect = txtSelect & " Source Type " & v_txtST
end if 
if  len(trim(v_txtSC)) > 0 then
    txtSQL =txtSQL &  " and CA.SourceCode LIKE '%" & v_txtSC & "%'"
	txtSelect = txtSelect & " Source Code " & v_txtSC
end if 
if  len(trim(v_txtTT)) > 0 then
    txtSQL =txtSQL &  " and CA.TargetType='" & v_txtTT & "'"
	txtSelect = txtSelect & " Target Type " & v_txtTT
end if 
if  len(trim(v_txtTC)) > 0 then
    txtSQL =txtSQL &  " and CA.TargetCode LIKE '%" & v_txtTC & "%'"
	txtSelect = txtSelect & " Target Code " & v_txtTC
end if 

if  len(trim(v_txtReference)) > 0 then
    txtSQL =txtSQL &  " and CA.Referenceno LIKE '%" & v_txtReference & "%'"
	txtSelect = txtSelect & " Reference " & v_txtReference
end if 

txtsql=(txtsql) & " group by TDATE, ca.JOBCODE, s.fcname, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, REFERENCENO" & _
                  " Order by TDATE,SOURCE, SOURCETYPE, SOURCECODE,TARGETTYPE, TARGETCODE,REFERENCENO"
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
<table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
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
							<td colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
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
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Activity</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Cost Element</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">ST</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">SC</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TT</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TC</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Referenceno</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Qty</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Output</th>
		<th align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Amount</th>
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
	    Dim rsCostElement,qCostElement,txtCostElement
		if rsCat("SOURCETYPE") ="EM" or rsCat("SOURCETYPE") ="ES" then
		   Set rsCostElement = connStr.Execute("Select ITEM from LOCATIONTYPE where LOCATIONTYPECODE='" & rscat("SOURCETYPE") & "'")
		else
		   Set rsCostElement = connStr.Execute("Select ITEM from LOCATIONTYPE where LOCATIONTYPECODE='" & rscat("SOURCE") & "'")
		end if   
		txtCostElement=""
	    if Not rsCostElement.eof then
	       txtCostElement=rsCostElement("ITEM")
	    end if
		rsCostElement.Close
		set rsCostElement = Nothing
	    %>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TDATE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOBCODE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=txtCostElement%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SOURCETYPE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SOURCECODE")%></td>
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
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="12"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Saldo,2)%></td>
	</tr>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->