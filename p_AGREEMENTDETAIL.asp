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
<title>ePlaS Reporting - Contract Agreement</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
    Dim txtNoWS,txtDateWS,txtRemarks,q,nCountofField,no,i
    Dim Remarks
    Dim rsWS,intRecordsPerPage,cColor,txtAgreementDetail
	'txtAgreementDetail =cmdFindTitle("AgreementDetail","fcCode",request("AgreementDetailCode"))
  
	AGREEMENTCODE	= Request("AGREEMENTCODE")
	CONTRACTORCODE	= Request("CONTRACTORCODE")
	DATE_FROM		= Request("DATE_FROM")
	DATE_TO			= Request("DATE_TO")
	LINENO			= Request("LINENO")
	JOBCODE			= Request("JOBCODE")
	LOCATIONTYPE	= Request("LOCATIONTYPE")
	LOCATIONCODE	= Request("LOCATIONCODE")
	businessunit = request("businessunit")

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

	
if request("txtKind") = "ContractAgreement" Then	
	txtSQL = "select AGREEMENTCODE, LINENO, JOB,b.fcname, LOCATIONTYPE, LOCATIONCODE, to_char(DUEDATE,'mm/dd/yyyy') as DUEDATE, " & _
     "RETENTION, QTY, RATE from Agreementdetail a ,job b  where a.job=b.fcCode and a.fcBA='" & businessunit & "'"
Else
	txtSQL =	" SELECT ca.agreementcode," & _
				"       ca.agreementdate," & _
				"       ca.contractorcode," & _
				"       ca.description," & _
				"       ca.retention," & _
				"       ca.startdate," & _
				"       ca.finishdate," & _
				"       ca.tax," & _
				"       cad.lineno," & _
				"       cad.jobcode," & _
				"       cad.qty," & _
				"       cad.rate," & _
				"       cad.reference," & _
				"       cad.remarks," & _
				"       jb.fcname jobname," & _
				"       cr.fcname contractorname," & _
				"       jb.uom_unit" & _
				"  FROM contractagreement ca" & _
				"       INNER JOIN contractagreementdetail cad" & _
				"          ON (ca.agreementcode = cad.agreementcode AND ca.fcba = cad.fcba)" & _
				"       LEFT JOIN jobmaster jb ON (jb.fccode = cad.jobcode )" & _
				"       LEFT JOIN contractor cr" & _
				"          ON (cr.fccode = ca.contractorcode AND cr.fcba = ca.fcba)" & _
				" WHERE     ca.fcba = '"& session("fcba") &"'"

End if
if  len(trim(AGREEMENTCODE)) > 0 then
	txtSQL = txtSQL & " AND ca.AGREEMENTCODE = '" & AGREEMENTCODE & "'"
end if
if  len(trim(CONTRACTORCODE)) > 0 then
	txtSQL = txtSQL & " AND (UPPER(ca.CONTRACTORCODE) like UPPER ('%" & CONTRACTORCODE & "%') OR c.fcname like UPPER('%"& CONTRACTORCODE &"%'))"
end if
if  len(trim(DATE_FROM)) > 0 then
	txtSQL = txtSQL & " AND ca.startdate >= TO_DATE('" & DATE_FROM & "', 'MM/DD/YYYY')"
end if
if  len(trim(DATE_TO)) > 0 then
	txtSQL = txtSQL & " AND ca.startdate <= TO_DATE('" & DATE_TO & "', 'MM/DD/YYYY')"
end if
if  len(trim(JOBCODE)) > 0 then
	txtSQL = txtSQL & " AND UPPER(cad.jobcode) like UPPER ('%" & JOBCODE & "%')"
end if
	
txtSQL = txtSQL &  " ORDER BY ca.AGREEMENTCODE, ca.CONTRACTORCODE, cad.LINENO"

' response.write txtSQL

Set rsAgreementDetail = Server.CreateObject("ADODB.Recordset")
rsAgreementDetail.CursorType = 3
rsAgreementDetail.CursorLocation = 3
rsAgreementDetail.Open txtSQL, connstr
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
					<h4><%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>AGREEMENT DETAIL</h2>
				</div>
			</div>
	</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
   
	<tr>
		<th>Agreement Code</th>	   
		<th>Contractor Code</th>
		<th>Contractor Name</th>
		<th>Line</th>
		<th>Job Code</th>
		<th>Job Name</th>
		<th>Job Uom</th>
		<th>Due Date</th>
		<th>Retention</th>
		<th>Quantity</th>
		<th>Rate</th>
		<th>Tax</th>
		<th>Description</th>
		<th>Remarks</th>
   </tr>
</thead>
<tbody>
   <%
   if Not rsAgreementDetail.eof then
      nCountofField=rsAgreementDetail.fields.count-1
	  no=1
	  nRetention = 0
	  QTY = 0
	  RATE = 0
	  While Not rsAgreementDetail.EOF

	      nRetention = nRetention+cdbl(to_nol(rsAgreementDetail("Retention")))
	      QTY = QTY+cdbl(rsAgreementDetail("QTY"))
	      RATE = RATE+cdbl(to_nol(rsAgreementDetail("RATE")))
		  ' if no  mod 2 = 1 then
             ' cColor = "#EEEEEE"
		  ' else
		     ' cColor = "#E5E5E5"
		  ' end if
		  %>
		  <tr bgcolor="<%=cColor%>">
          <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("AGREEMENTCODE")%></td>
		  <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("CONTRACTORCODE")%></td>
		  <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("CONTRACTORNAME")%></td>
	      <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LINENO")%></td>
	      <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("JOBCODE")%></td>
	      <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("JOBNAME")%></td>
		  <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("uom_unit")%></td>
		  <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("startdate")%> - <%=rsAgreementDetail("finishdate")%></td>
	      <td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsAgreementDetail("RETENTION"))),2)%></td>
	      <td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsAgreementDetail("QTY"))),2)%></td>
	      <td align="Right" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsAgreementDetail("RATE"))),2)%></td>
	      <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("tax")%></td>
	      <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("description")%></td>
		  <td align="left" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("remarks")%></td>
		  </tr>
		  <%
		  rsAgreementDetail.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend %>
	  
	<tr bgcolor="#F4F4A0">
		<td colspan="8" align="center" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
		<td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(nRetention)),2)%></strong></td>
		<td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(QTY)),2)%></strong></td>
		<td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(RATE)),2)%></strong></td>
		<td align="right"style="<%=cmdstyle%>" colspan = "3"><font face="Verdana" size="1"><strong></strong></td>
	</tr>
</tbody>
<%Else%>
	<tr>
		<td colspan="13" align="Left">NO DATA AVAILABLE</td>
	</tr>
</tbody>
<%
   end if 
   rsAgreementDetail.Close
   set rsAgreementDetail = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>