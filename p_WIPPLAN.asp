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
<title>ePlaS Reporting - Contract Work In Progress Plan</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

DATE_FROM = Request("DATE_FROM")
DATE_TO = Request("DATE_TO")
AGREEMENTCODE = Request("AGREEMENTCODE")
CONTRACTORCODE = Request("CONTRACTORCODE")
LINENO = Request("LINENO")
JOBCODE = Request("JOBCODE")
LOCATIONTYPE = Request("LOCATIONTYPE")
LOCATIONCODE = Request("LOCATIONCODE")
businessunit= trim(request("businessunit"))

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if
	
txtSQL= "  SELECT wip.agreementcode," & _
		"         TO_CHAR (wip.wipdate, 'mm/dd/yyyy') AS wipdate," & _
		"         ca.contractorcode," & _
		"         cr.fcname contractorname," & _
		"         wip.agreementlineno," & _
		"         cad.jobcode," & _
		"         jb.fcname jobname," & _
		"         wip.locationcode," & _
		"         wip.locationtype," & _
		"         wip.REFERENCE ||'-'||cad.remarks remarks," & _
		"         nvl (cad.rate,0) rate," & _
		"         NVL (SUM (NVL (wip.qty, 0)), 0) AS qty," & _
		"         NVL (SUM (NVL (wip.unit, 0)), 0) AS unit," & _
		"         NVL (SUM (NVL (wip.output, 0)), 0) AS output," & _
		"         NVL (SUM (NVL (wip.qty, 0)), 0) * nvl (cad.rate,0) AS total" & _
		"    FROM workinprogressplan wip" & _
		"         LEFT JOIN contractagreement ca" & _
		"            ON (wip.agreementcode = ca.agreementcode AND wip.fcba = ca.fcba)" & _
		"         LEFT JOIN contractagreementdetail cad" & _
		"            ON (cad.agreementcode = ca.agreementcode AND cad.fcba = ca.fcba AND wip.agreementlineno = cad.lineno)" & _
		"         LEFT JOIN jobmaster jb ON (jb.fccode = cad.jobcode )" & _
		"         LEFT JOIN contractor cr" & _
		"            ON (cr.fccode = ca.contractorcode AND cr.fcba = ca.fcba)" & _
		"   WHERE  /*   (    wip.agreementcode = ca.agreementcode */" & _
		"          /*    AND wip.agreementlineno = cad.lineno */" & _
		"          /*    AND wip.fcba = ca.fcba) */" & _
		"          /* AND (cad.jobcode = jb.fccode AND cad.fcba = jb.fcba) */" & _
		"          /* AND (ca.contractorcode = cr.fccode AND ca.fcba = cr.fcba) */" & _
		"          wip.fcba = '"& session("fcba") &"' AND wip.wipamt > 0" 

if  len(trim(DATE_FROM)) > 0 then
	txtSQL = txtSQL & " and wip.wipdate >= to_date('" & DATE_FROM & "', 'mm/dd/yyyy')"
end if
if  len(trim(DATE_TO)) > 0 then
	txtSQL = txtSQL & " and wip.wipdate <= to_date('" & DATE_TO & "', 'mm/dd/yyyy')"
end if
if  len(trim(AGREEMENTCODE)) > 0 then
	txtSQL = txtSQL & " and ca.agreementcode = '" & AGREEMENTCODE & "'"
end if
if  len(trim(CONTRACTORCODE)) > 0 then
	txtSQL = txtSQL & " and (UPPER(ca.contractorcode) like UPPER ('%" & CONTRACTORCODE & "&') OR cr.fcname like UPPER('%"& CONTRACTORCODE &"%')) "
end if
if  len(trim(LINENO)) > 0 then
	txtSQL = txtSQL & " and wip.agreementlineno = '" & LINENO & "'"
end if
if  len(trim(JOBCODE)) > 0 then
	txtSQL = txtSQL & " and UPPER(cad.jobcode) like UPPER ('%" & JOBCODE & "%')"
end if
if  len(trim(LOCATIONTYPE)) > 0 then
	txtSQL = txtSQL & " and wip.locationtype = '" & LOCATIONTYPE & "'"
end if 
if  len(trim(LOCATIONCODE)) > 0 then
	txtSQL = txtSQL & " and wip.locationcode = '" & LOCATIONCODE & "'"
END IF	

txtsql = txtsql & 	"	GROUP BY wip.agreementcode," & _
					"         TO_CHAR (wip.wipdate, 'mm/dd/yyyy')," & _
					"         ca.contractorcode," & _
					"         cr.fcname," & _
					"         wip.agreementlineno," & _
					"         cad.jobcode," & _
					"         jb.fcname," & _
					"         wip.REFERENCE ||'-'||cad.remarks," & _
					"         cad.rate," & _
					"         wip.locationcode," & _
					"         wip.locationtype" 
					
txtsql = txtsql &	"	ORDER BY wip.agreementcode," & _
					"         ca.contractorcode," & _
					"         wip.agreementlineno," & _
					"         TO_CHAR (wip.wipdate, 'mm/dd/yyyy')" 

					
	' response.Write(txtsql)
	'response.Flush()
%>
<body>
	<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
			<tr>
				<th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
						<div id="head_judul" style="float:left; text-align:left;">
							<h4><%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h2>WORK IN PROGRESS PLAN</h2>
						</div>
					</div>
				</th>
			</tr>
			<tr>
				<td colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
			</tr>
<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField = rsCat.fields.count-1
	end if
	NO = 1
%>
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
				<th>Quantity</th>
				<th>Unit</th>
				<th>Output</th>
				<th>Estimated WIP</th>
				<th>Remarks</th>
			</tr>
			</thead>
			<tbody>
	<%
	Dim v_Quantity : v_Quantity = 0
	Dim v_Unit : v_Unit = 0
	Dim v_Output : v_Output = 0
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
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("UNIT"), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("OUTPUT"), 2)%></td>
				<td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("TOTAL"), 0)%></td>
				<td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("REMARKS")%></td>
			</tr>
<%
	  v_Quantity = v_Quantity + trim(cdbl(rsCat("QTY")))
	  v_Unit = v_Unit + trim(cdbl(rsCat("UNIT")))
	  v_Output = v_Output + trim(cdbl(rsCat("OUTPUT")))
	  v_Total = v_Total+ trim(cdbl(rsCat("TOTAL")))
	  
	  rsCat.MoveNext
	  
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyleLeft%>" colspan="10"> <font face="Verdana" size="1"><strong>T o t a l</strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Quantity, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Unit, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Output, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Total, 0)%></strong></td>
			<td></td>
		</tr>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>