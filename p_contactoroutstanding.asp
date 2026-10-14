<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Contract Outstanding</title>
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


</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

' txtSQL = "	SELECT DATA.agreementcode,DATA.agreementlineno, DATA.wipdate, DATA.contractorcode, " &_
		 ' "	contractor.fcname, DATA.PAID, DATA.totalContract, DATA.netbalance " &_
		 ' "	FROM (SELECT a.agreementcode, a.agreementlineno, b.contractorcode, a.wipdate," &_
		 ' "				a.PAID," &_
		 ' "				b.totalContract," &_
		 ' "				(b.totalContract - a.PAID) netbalance" &_
		 ' "	FROM (SELECT agreementcode, agreementlineno, wipdate, nvl(PAID,0) paid" &_
		 ' "	FROM workinprogress) a," &_
		 ' "		 (SELECT contractorcode, agreementcode, lineno, SUM (qty * rate) totalContract" &_
		 ' "	FROM agreementdetail" &_
		 ' "	GROUP BY  contractorcode, agreementcode, lineno) b" &_
		 ' "	WHERE" &_
		 ' "			a.agreementcode = b.agreementcode and" &_
		 ' "			a.AGREEMENTLINENO = b.lineno and" &_
		 ' "			(b.totalContract - a.PAID) > 0) DATA," &_
		 ' "			contractor" &_
		 ' "	WHERE DATA.contractorcode = contractor.fccode"'"
		 
		 
txtSQL	=	" SELECT DATA.FCBA," & _
			"        DATA.AGREEMENTCODE," & _
			"        DATA.AGREEMENTLINENO," & _
			"        DATA.WIPDATE," & _
			"        DATA.CONTRACTORCODE," & _
			"        CONTRACTOR.FCNAME," & _
			"        DATA.PAID," & _
			"        DATA.TOTALCONTRACT," & _
			"        DATA.NETBALANCE" & _
			"   FROM (SELECT A.FCBA," & _
			"                A.AGREEMENTCODE," & _
			"                A.AGREEMENTLINENO," & _
			"                B.CONTRACTORCODE," & _
			"                A.WIPDATE," & _
			"                A.PAID," & _
			"                B.TOTALCONTRACT," & _
			"                (B.TOTALCONTRACT - A.PAID) NETBALANCE" & _
			"           FROM (SELECT FCBA," & _
			"                        AGREEMENTCODE," & _
			"                        AGREEMENTLINENO," & _
			"                        WIPDATE," & _
			"                        NVL (PAID, 0) PAID" & _
			"                   FROM WORKINPROGRESS) A," & _
			"                (  SELECT FCBA," & _
			"                          CONTRACTORCODE," & _
			"                          AGREEMENTCODE," & _
			"                          LINENO," & _
			"                          SUM (QTY * RATE) TOTALCONTRACT" & _
			"                     FROM AGREEMENTDETAIL" & _
			"                 GROUP BY FCBA," & _
			"                          CONTRACTORCODE," & _
			"                          AGREEMENTCODE," & _
			"                          LINENO) B" & _
			"          WHERE     A.AGREEMENTCODE = B.AGREEMENTCODE" & _
			"                AND A.AGREEMENTLINENO = B.LINENO" & _
			"                AND A.FCBA = B.FCBA" & _
			"                AND (B.TOTALCONTRACT - A.PAID) <> 0) DATA," & _
			"        CONTRACTOR" & _
			"  WHERE     DATA.CONTRACTORCODE = CONTRACTOR.FCCODE" & _
			"        AND DATA.FCBA = CONTRACTOR.FCBA" & _
			"        AND DATA.FCBA = '" & Session("Fcba") & "'" & _
			" ORDER BY DATA.AGREEMENTCODE, DATA.AGREEMENTLINENO"

%>
	<body>
		<div id="wrapper">
			<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<th colspan="8" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
						<div id="head_judul" style="float:left; text-align:left;">
							<h4>PT. <%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h1 class="judul">INVOICE OUTSTANDING</h1>
						</div>
						<div style="clear:both;"></div>
					</div>
					</th>
				</tr>
				<tr>
					<th>AGREEMENT CODE</th>
					<th>AGREEMENT LINENO</th>
					<th>WIP DATE</th>
					<th>CONTRACTOR CODE</th>	
					<th>CONTRACTOR NAME</th>		
					<th>PAID</th>
					<th>TOTAL CONTRACT</th>
					<th>NET BALANCE</th>					
				</tr>
			</thead>
			<tbody>
				
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
				   nCountofField=rsCat.fields.count-1
				end if
				NO = 1
				%>
				<%If (rsCat.BOF and rsCat.EOF) Then%>
					<tr>
						<td Colspan = "8" align = "Left">NO DATA AVAILABLE</td>
					</tr>
				<%Else%>
					<%	While Not rsCat.EOF %>
					<tr>
						<td align="Left"><%=rscat("AgreementCode")%></td>
						<td align="Right"><%=rscat("AgreementLineNo")%></td>
						<td align="Right"><%=rscat("WipDate")%></td>
						<td align="Left"><%=rscat("ContractorCode")%></td>
						<td align="Left"><%=rscat("fcname")%></td>
						<td align="Right"><%=FormatNumber(rscat("PAID"),2)%></td>
						<td align="Right"><%=FormatNumber(rscat("totalContract"),2)%></td>
						<td align="Right"><%=FormatNumber(rscat("netbalance"),2)%></td>
					</tr>
					<%
						nRptotalContract  = nRptotalContract+cdbl(rsCat("totalContract"))
						nRpNetbalance  = nRpNetbalance+cdbl(rsCat("netbalance"))
						
						rsCat.MoveNext
						'response.Flush()
						'no=no+1
					Wend
					rsCat.close()
					rsCat=null
					%>
					<tr bgcolor="#F4F4A0">
						<td align="Center" colspan="6"><strong><i>T o t a l</i></strong></td>
						<td align="Right"><strong><i><%=FormatNumber(nRptotalContract,2)%></i></strong></td>
						<td align="Right"><strong><i><%=FormatNumber(nRpNetbalance,2)%></i></strong></td>
					</tr>
				<%End If%>
			</tbody>
			</table>
		</div>
	</body>
</html>