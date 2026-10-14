<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Contract Agreement</title>
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
    Dim txtNoWS,txtDateWS,txtRemarks,q,nCountofField,no,i
    Dim Remarks
    Dim rsWS,intRecordsPerPage,cColor,txtAgreementDetail
	'txtAgreementDetail =cmdFindTitle("AgreementDetail","fcCode",request("AgreementDetailCode"))
    %>

	 <body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>AGREEMENT DETAIL</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="13"><%=txtSelect%></td>
						</tr>
						</div>
   <%
	v_Contractor_From=trim(request("DATEACTIVITY_From"))
	v_Contractor_To=trim(request("DATEACTIVITY_To"))
	v_JOBCODE=trim(request("JOBCODE"))
	v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
	v_LocationCode=trim(request("LocationCode"))
	
if request("txtKind") = "ContractAgreement" Then	
	txtSQL = "select AGREEMENTCODE, LINENO, JOB,b.fcname, LOCATIONTYPE, LOCATIONCODE, to_char(DUEDATE,'mm/dd/yyyy') as DUEDATE, " & _
     "RETENTION, QTY, RATE from Agreementdetail a ,job b  where a.job=b.fcCode and a.fcBA='" & Session("fcBA") & "'"
Else
	txtSQL = "select AGREEMENTCODE, LINENO, JOB,b.fcname, LOCATIONTYPE, LOCATIONCODE, to_char(DUEDATE,'mm/dd/yyyy') as DUEDATE, " & _
     "RETENTION, QTY, RATE from Agreementdetail a ,job b  where a.job=b.fcCode and a.fcBA='" & Session("fcBA") & "'"
End if
	
	if  len(trim(v_Contractor_From)) > 0 then
		txtSQL = txtSQL & " and a.DUEDATE >=to_Date('" & v_Contractor_From & "','MM/DD/YYYY')"
	end if
	if  len(trim(v_Contractor_To)) > 0 then
		txtSQL =txtSQL &  " and a.DUEDATE <=to_Date('" & v_Contractor_To & "','MM/DD/YYYY')"
	end if
	if  len(trim(v_JOBCODE)) > 0 then
		txtSQL = txtSQL & " and a.JOB='" & v_JOBCODE & "'"
	end if
	if  len(trim(v_LOCATIONTYPE)) > 0 then
		txtSQL =txtSQL &  " and a.LOCATIONTYPE='" & v_LOCATIONTYPE & "'"
	end if 
	if  len(trim(v_LocationCode)) > 0 then
		txtSQL =txtSQL &  " and a.LOCATIONCODE='" & v_LocationCode & "'"
	end if 
	
   'response.Write q	
   Set rsAgreementDetail = Server.CreateObject("ADODB.Recordset")
   rsAgreementDetail.CursorType = 3
   rsAgreementDetail.CursorLocation = 3
   rsAgreementDetail.Open txtSQL, connstr%>
   
  <tr>
	<th>Code</th>	   
    <th>Cont. No</th>
    <th>Line</th>
    <th>Job Code</th>
    <th>Job Desc</th>
    <th>Target Type</th>
    <th>Target Code</th>
    <th>Due Date</th>
    <th>Retention</th>
    <th>Quantity</th>
    <th>Rate</th>
   </tr>
   <%
   if Not rsAgreementDetail.eof then
      nCountofField=rsAgreementDetail.fields.count-1
	  no=1
	  nRetention = 0
	  QTY = 0
	  RATE = 0
	  While Not rsAgreementDetail.EOF

	      nRetention = nRetention+cdbl(rsAgreementDetail("Retention"))
	      QTY = QTY+cdbl(rsAgreementDetail("QTY"))
	      RATE = RATE+cdbl(rsAgreementDetail("RATE"))
		  ' if no  mod 2 = 1 then
             ' cColor = "#EEEEEE"
		  ' else
		     ' cColor = "#E5E5E5"
		  ' end if
		  %>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
          <td align="left" title="WS.AGREEMENTCODE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("AGREEMENTCODE")%></td>
	      <td align="left" title="WS.LINENO" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LINENO")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("JOB")%></td>
	      <td align="left" title="WS.fcname" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("fcname")%></td>
	      <td align="left" title="WS.LOCATIONTYPE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LOCATIONTYPE")%></td>
	      <td align="left" title="WS.LOCATIONCODE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LOCATIONCODE")%></td>
	      <td align="left" title="WS.DUEDATE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("DUEDATE")%></td>
	      <td align="Right" title="WS.RETENTION" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsAgreementDetail("RETENTION"),2)%></td>
	      <td align="Right" title="WS.QTY" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsAgreementDetail("QTY"),2)%></td>
	      <td align="Right" title="WS.RATE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsAgreementDetail("RATE"),2)%></td>
		  </tr>
		  <%
		  rsAgreementDetail.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend %>
	  
	 <tr bgcolor="#F4F4A0">
	  <td colspan="8" align="center" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1">T o t a l</td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nRetention,2)%></td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(QTY,2)%></td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(RATE,2)%></td>
	  </tr>
	  <%
   else
      response.Write("Data is empty")	  
   end if 
   rsAgreementDetail.Close
   set rsAgreementDetail = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>