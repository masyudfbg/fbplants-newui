<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Purchase Orders Transaction</title>
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

v_txtPODATE_From=trim(request("PODATE_From"))
v_txtPODATE_To=trim(request("PODATE_To"))
v_txtPOCODE=trim(request("POCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtRequestorID=trim(request("REQUESTORID"))
v_txtRowStateId=trim(request("ROWSTATEID"))
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

txtSQL = 	"  SELECT purchaseorder.fcba AS BUSINESSUNIT," & _
			"         purchaseorder.pocode AS POCODE," & _
			"         purchaseorder.podate AS PODATE," & _
			"         purchaseorderdetails.itemcode AS ITEMCODE," & _
			"         stock.fcname AS ITEMNAME," & _
			"         purchaseorderdetails.prcode||'-'||purchaseorderdetails.prlineid AS procdeid," & _
			"         purchaseorderdetails.quantity AS POQUANTITY," & _
			"         purchaseorderdetails.amount AS POAMOUNT," & _
			"         authorizelevel.fccode AS AUTHORIZELEVELCODE," & _
			"         authorizelevel.fcname AS AUTHORIZELEVELNAME," & _
			"         purchaseorderdetails_hist.fcedit AS LASTAUTHORIZER," & _
			"         purchaseorderdetails_hist.lastupdate AS LASTACTIONDATE" & _
			"    FROM purchaseorderdetails_hist" & _
			"         INNER JOIN purchaseorderdetails" & _
			"            ON (    purchaseorderdetails_hist.fcba = purchaseorderdetails.fcba" & _
			"                AND purchaseorderdetails_hist.pocode =" & _
			"                       purchaseorderdetails.pocode)" & _
			"         INNER JOIN purchaseorder" & _
			"            ON (    purchaseorderdetails.fcba = purchaseorder.fcba" & _
			"                AND purchaseorderdetails.pocode = purchaseorder.pocode)" & _
			"         INNER JOIN authorizelevel" & _
			"            ON purchaseorderdetails_hist.currentauthorizelevel =" & _
			"                  authorizelevel.fccode" & _
			"         INNER JOIN" & _
			"         (  SELECT purchaseorderdetails_hist.pocode," & _
			"                   MAX (purchaseorderdetails_hist.currentauthorizelevel) alevel" & _
			"              FROM purchaseorderdetails_hist" & _
			"          GROUP BY purchaseorderdetails_hist.pocode) aa" & _
			"            ON (    purchaseorderdetails_hist.currentauthorizelevel = aa.alevel" & _
			"                AND purchaseorderdetails_hist.pocode = aa.pocode)" & _
			"         INNER JOIN stock" & _
			"            ON (    purchaseorderdetails.fcba = stock.fcba" & _
			"                AND purchaseorderdetails.itemcode = stock.fccode)" & _
			" WHERE purchaseorder.fcba = '"& businessunit &"'" 

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and purchaseorder.podate >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtPODATE_From,1)	
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and purchaseorder.podate <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtPODATE_To,1)	
end if
if  len(trim(v_txtPOCODE)) > 0 then
    txtSQL = txtSQL & " and purchaseorder.pocode='" & v_txtPOCODE & "'"
	txtSelect = txtSelect & " PO Code " & v_txtPOCODE	
end if
	txtSQL = txtSQL &	" GROUP BY purchaseorder.fcba," & _
						"         purchaseorder.pocode," & _
						"         purchaseorder.podate," & _
						"         purchaseorderdetails.itemcode," & _
						"         stock.fcname," & _
						"         purchaseorderdetails.quantity," & _
						"         purchaseorderdetails.amount," & _
						"         authorizelevel.fccode," & _
						"         authorizelevel.fcname," & _
						"         purchaseorderdetails.prcode||'-'||purchaseorderdetails.prlineid," & _
						"         purchaseorderdetails_hist.fcedit," & _
						"         purchaseorderdetails_hist.lastupdate" 

	txtSQL = txtSQL & 	" ORDER BY purchaseorder.fcba," & _
						"         purchaseorder.pocode," & _
						"         purchaseorderdetails.itemcode"


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PURCHASE ORDERS TRANSACTION</h2>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="11"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							<%	
							
							Set rsCat = Server.CreateObject("ADODB.Recordset")
							rsCat.CursorType = 3
							rsCat.CursorLocation = 3
							rsCat.Open txtSQL, connstr

							
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							<!--<th>No.</th>-->
							<th width="200">PO Code</th>
							<th>Date</th>
							<th>PR Code</th>
							<th>Item Code</th>
							<th>Item Name</th>
							<th>PO Quantity</th>
							<th>PO Amount</th>
							<th>Last Authorize Level Code</th>
							<th>Last Authorizer level Name</th>
							<th>Last Authorizer</th>
							<th>Last Action Date</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
							v_Quantity=0
							v_Amount = 0
							While Not rsCat.EOF%>


							<tr>
								<!--<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%> </td>-->
								<td align="Center" style="<%=cmdstyle%>">
									<a href="javascript:launchwin('p_AuthorizeTransactionPODet.Asp?pocode=<%=rscat("pocode")%>&podate=<%=rscat("podate")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1"><%=rscat("pocode")%></font> 
									</a> 
								</td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("podate")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("procdeid")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("itemcode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("itemname")%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("poquantity"))),0)%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("poamount"))),0)%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizelevelcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizelevelname")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("lastauthorizer")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("lastactiondate")%> </td>
							</tr>
							  <%

							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							rsCat.close()
							rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	