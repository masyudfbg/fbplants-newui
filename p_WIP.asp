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
<title>ePlaS Reporting - Contract Work In Progress</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

DATE_FROM		= Request("DATE_FROM")
DATE_TO			= Request("DATE_TO")
AGREEMENTCODE	= Request("AGREEMENTCODE")
CONTRACTORCODE	= Request("CONTRACTORCODE")
LINENO			= Request("LINENO")
JOBCODE			= Request("JOBCODE")
LOCATIONTYPE	= Request("LOCATIONTYPE")
LOCATIONCODE	= Request("LOCATIONCODE")

	
txtSQL =	" SELECT WIP.AGREEMENTCODE" & _
			"        , TO_CHAR(WIP.WIPDATE,'MM/DD/YYYY') AS WIPDATE" & _
			"        , AG.CONTRACTORCODE" & _
			"        , C.FCNAME CONTRACTORNAME" & _
			"        , WIP.AGREEMENTLINENO" & _
			"        , AG.JOB JOBCODE" & _
			"        , JOB.FCNAME JOBNAME" & _
			"        , AG.LOCATIONTYPE" & _
			"        , AG.LOCATIONCODE" & _
			"        , AG.REMARKS" & _
			"        , AG.RATE" & _	
			"        , NVL(SUM(NVL(WIP.QTY, 0)), 0) AS QTY" & _
			"        , ROUND(NVL(SUM(NVL(WIP.QTY, 0)), 0)*AG.RATE, 0) AS TOTAL" & _
			" FROM WORKINPROGRESS WIP" & _
			"      , AGREEMENTDETAIL AG" & _
			"      , JOB" & _
			"      , CONTRACTOR C" & _
			" WHERE (WIP.AGREEMENTCODE = AG.AGREEMENTCODE AND WIP.AGREEMENTLINENO = AG.LINENO AND WIP.FCBA = AG.FCBA)" & _
			"       AND (AG.JOB = JOB.FCCODE AND AG.FCBA = JOB.FCBA)" & _
			"       AND (AG.CONTRACTORCODE = C.FCCODE AND AG.FCBA = C.FCBA)" & _
			"       AND WIP.FCBA = '" & session("fcBA") & "'"
if  len(trim(DATE_FROM)) > 0 then
	txtSQL = txtSQL & " AND WIP.WIPDATE >= TO_DATE('" & DATE_FROM & "', 'MM/DD/YYYY')"
end if
if  len(trim(DATE_TO)) > 0 then
	txtSQL = txtSQL & " AND WIP.WIPDATE <= TO_DATE('" & DATE_TO & "', 'MM/DD/YYYY')"
end if
if  len(trim(AGREEMENTCODE)) > 0 then
	txtSQL = txtSQL & " AND AG.AGREEMENTCODE='" & AGREEMENTCODE & "'"
end if
if  len(trim(CONTRACTORCODE)) > 0 then
	txtSQL = txtSQL & " AND AG.CONTRACTORCODE='" & CONTRACTORCODE & "'"
end if
if  len(trim(LINENO)) > 0 then
	txtSQL = txtSQL & " AND WIP.AGREEMENTLINENO='" & LINENO & "'"
end if
if  len(trim(JOBCODE)) > 0 then
	txtSQL = txtSQL & " AND AG.JOB='" & JOBCODE & "'"
end if
if  len(trim(LOCATIONTYPE)) > 0 then
	txtSQL = txtSQL & " and AG.LOCATIONTYPE='" & LOCATIONTYPE & "'"
end if 
if  len(trim(LOCATIONCODE)) > 0 then
	txtSQL = txtSQL & " and AG.LOCATIONCODE='" & LOCATIONCODE & "'"
END IF	

txtsql	= txtsql & 	" GROUP BY WIP.AGREEMENTCODE" & _
					"          , TO_CHAR(WIP.WIPDATE,'MM/DD/YYYY')" & _
					"          , AG.CONTRACTORCODE" & _
					"          , C.FCNAME" & _
					"          , WIP.AGREEMENTLINENO" & _
					"          , AG.JOB" & _
					"          , JOB.FCNAME" & _
					"          , AG.LOCATIONTYPE" & _
					"          , AG.LOCATIONCODE" & _
					"          , AG.REMARKS" & _
					"          , AG.RATE"
					
txtsql	= txtsql & 	" ORDER BY WIP.AGREEMENTCODE" & _
					"          , AG.CONTRACTORCODE" & _
					"          , WIP.AGREEMENTLINENO" & _
					"          , TO_CHAR(WIP.WIPDATE,'MM/DD/YYYY')"
					
	'response.Write(txtsql)
	'response.Flush()
%>
 <body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
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
					<h2>WORK IN PROGRESS</h2>
				</div>
			</div>
		</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="12" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
   <tr>   
	   <th>Agreement Code</th>
	   <th>Contractor Code</th>
	   <th>Contractor Name</th>
	   <th>Line No</th>
	   <th>WIP Date</th>	   
	   <th>Target Type</th>
	   <th>Target Code</th>
	   <th>Job Code</th>
	   <th>Job Name</th>
	   <th>Rate</th>
	   <th>Qty</th>	   
	   <th>Total</th>
	   <th>Remarks</th>
   </tr>
    </thead>
    <tbody>
	<%
	Dim v_Quantity : v_Quantity = 0
	Dim v_Total : v_Total = 0
	
	While Not rsCat.EOF%>
		<tr bgcolor="<%=cColor%>">		   
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("AGREEMENTCODE")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("CONTRACTORCODE")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("CONTRACTORNAME")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("AGREEMENTLINENO")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("WIPDATE")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("LOCATIONTYPE")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("LOCATIONCODE")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("JOBCODE")%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("JOBNAME")%></td>
			<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("RATE"), 2)%></td>
			<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTY"), 2)%></td>			
			<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("TOTAL"), 0)%></td>
			<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("REMARKS")%></td>
		</tr>
	  <%
	  v_Quantity = v_Quantity + trim(cdbl(rsCat("QTY")))
	  v_Rate = v_Rate + trim(cdbl(rsCat("RATE")))
	  v_Total = v_Total+ trim(cdbl(rsCat("TOTAL")))
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
			<td align="Center" style="<%=cmdstyleLeft%>" colspan="10"> <font face="Verdana" size="1"><strong>T o t a l</strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Quantity, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Total, 0)%></strong></td>
			<td></td>
		</tr>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>