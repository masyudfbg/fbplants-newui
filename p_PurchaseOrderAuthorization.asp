<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Purchase Order Approval</title>
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
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPODATE_From=trim(request("PODATE_From"))
v_txtPODATE_To=trim(request("PODATE_To"))
v_txtPOCODE=trim(request("POCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtSUPPLIERCODE=trim(request("SUPPLIERCODE"))
ROWSTATEID	=	trim(request("ROWSTATEID")) 'ditambah ukie tgl 9 okt 2013 untuk filter status
businessunit=trim(request("businessunit"))

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

	txtSQL	=	" SELECT hdr.fcba, hdr.pocode, hdr.podate, hdr.purchaseby, hdr.grandtotal, hdr.remarks, aus.authorizelevelcode, NVL(hist.currentauthorizelevel,0) currentauthorizelevel" & _
			    " FROM purchaseorder hdr" & _
			    " LEFT JOIN authorizetemplate at ON (at.document_type = 'PO' AND at.document_classification = '0')" & _
			    " LEFT JOIN ( SELECT fcba, pocode, NVL(MAX(currentauthorizelevel),0) currentauthorizelevel FROM purchaseorderdetails_hist group by fcba, pocode) hist ON (hist.fcba = hdr.fcba AND hist.pocode = hdr.pocode)" & _
				" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "' AND authorizestructure_api.get_nextlevel('PO', '0', hdr.fcba, NVL(hist.currentauthorizelevel,0)) = aus.authorizelevelcode )" & _
			    " WHERE hdr.fcba = '"& businessunit &"' AND hdr.rowstateid = 'AuthorizeOnProgress' AND hdr.grandtotal BETWEEN at.amountrangefrom AND at.amountrangeuntil" 


if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and hdr.podate >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtPODATE_From,1)
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and hdr.podate <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtPODATE_To,1)
end if
if  len(trim(v_txtPOCODE)) > 0 then
    txtSQL = txtSQL & " and hdr.pocode='" & v_txtPOCODE & "'"
	txtSelect = txtSelect & " PO Code " & v_txtPOCODE
end if

	txtsql	=	(txtsql)	&		" ORDER BY hdr.fcba, hdr.pocode"

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="10" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div  id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyNAme%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PURCHASE ORDER AUTHORIZATION</h2>
				</div>
			</div>
			 </th>
			</tr>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="10" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						</div>
						<tr>
						<%	
						Set rsCat = connStr.Execute(txtSQL)
						if Not rsCat.eof then
						   nCountofField=rsCat.fields.count-1
						end if
						NO=1%>
						</tr>
						<tr>
							 
							<th>No</th>
							<th>PO Code</th>
							<th>PO Date</th><!--'ditambah ukie tgl 9 okt 2013 untuk tampilkan status-->
							<th>Purchase By</th>
							<th>Remarks</th>
							<th>Status</th>
							<th>Amount</th>
						</tr>
                        </thead>
                <tbody>
							<%
								Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity, v_Received
								v_Quantity=0
								v_Received = 0
								v_Amount = 0
								While Not rsCat.EOF
								no=no+1
							%>


						<tr>
							 
							<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=NO%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%> </td>
							<!--'ditambah ukie tgl 9 okt 2013 untuk tampilkan status-->
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("podate")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("purchaseby")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstateid")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("Amount")),2)%> </td>
						</tr>
						  <%
						  	rsCat.MoveNext
						  'response.Flush()
							Wend
							rsCat.close()
							rsCat=null
						  %>
						<input  type="hidden" value="<%=txtTable%>" name="txtTable">
						<input  type="hidden" value="<%=txtModule%>" name="txtModule">
						</form>
                        </tbody>
					</table>
	<!--#include file="Bottom_Report.asp"-->
		</DIV>			
	</BODY>
</HTML>	
