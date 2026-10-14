<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Delivery Order</title>
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
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtDELIVERYORDERDATE_From=trim(request("DELIVERYORDERDATE_From"))
v_txtDELIVERYORDERDATE_To=trim(request("DELIVERYORDERDATE_To"))
v_txtSALESCONTRACTNO=trim(request("SALESCONTRACTNO"))
v_txtDELIVERYORDERNO=trim(request("DELIVERYORDERNO"))
' v_txtITEMCODE=trim(request("ITEMCODE"))
' v_txtRequestorID=trim(request("REQUESTORID"))
' v_txtRowStateId=trim(request("ROWSTATEID"))
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


txtSQL =  " SELECT sad.salescontractno, sad.lineno, sad.productcode, sad.unitofmeasure, nvl(do.quantity_do,0)quantity_do, do.deliverycondition, do.remarks, do.deliveryorderdate, do.deliveryorderno, " & _
		  " sad.salesprices, do.startdate, do.enddate, sad.amount, sad.amountvat, sad.totalamountvat, sad.quantity, sad.vat, " & _
		  " sah.customercode, sah.creationdate, sah.deliverydate, sah.deliverytype, sad.lineno," & _
		  " prod.productdescription as productname, prod.unitofmeasure, " & _
		  " tc.transportercontractno, tc.transporter, tr.fcname transportername, tc.losstolerance, do.transportationcost, sad.fcba" & _
		  " FROM salesagreementdet sad" & _
		  " INNER JOIN salesagreementhdr sah ON (sad.salescontractno = sah.salescontractno AND sad.fcba = sah.fcba)" & _
		  " LEFT JOIN deliveryorder do ON (sad.salescontractno = do.salescontractno AND sad.fcba = do.fcba)" & _
 		  " LEFT JOIN product prod ON (prod.productcode = sad.productcode AND sad.fcba = prod.fcba)" & _
		  " LEFT JOIN transportercontract tc ON (do.transportercontractno = tc.transportercontractno AND do.fcba = tc.fcba AND do.transporterline = tc.lineno)" & _
		  " INNER JOIN transporter tr ON (tc.fcba = tr.fcba AND tc.transporter = tr.fcCode)" & _
		  " WHERE sad.fcba = '" & businessunit & "'" 


' response.write(txtSQL)
if  len(trim(v_txtDELIVERYORDERDATE_From)) > 0 then
    txtSQL = txtSQL & " and DELIVERYORDERDATE >=to_Date('" & v_txtDELIVERYORDERDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtDELIVERYORDERDATE_From,1)	
end if
if  len(trim(v_txtDELIVERYORDERDATE_To)) > 0 then
    txtSQL =txtSQL &  " and DELIVERYORDERDATE <=to_Date('" & v_txtDELIVERYORDERDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtDELIVERYORDERDATE_To,1)	
end if
if  len(trim(v_txtSALESCONTRACTNO)) > 0 then
    txtSQL = txtSQL & " and sad.SALESCONTRACTNO='" & v_txtSALESCONTRACTNO & "'"
	txtSelect = txtSelect & " Sales Contract No. " & v_txtSALESCONTRACTNO	
end if
if  len(trim(v_txtDELIVERYORDERNO)) > 0 then
    txtSQL = txtSQL & " and do.DELIVERYORDERNO='" & v_txtDELIVERYORDERNO & "'"
	txtSelect = txtSelect & " Delivery Order No. " & v_txtDELIVERYORDERNO	
end if
' if  len(trim(v_txtITEMCODE)) > 0 then
'     txtSQL =txtSQL &  " and b.itemcode='" & v_txtITEMCODE & "'"
' 	txtSelect = txtSelect & " Item Code " & v_txtITEMCODE	
' end if 
' if  len(trim(v_txtRequestorID)) > 0 then
'     txtSQL =txtSQL &  " and a.RequestorId='" & v_txtRequestorID & "'"
' 	txtSelect = txtSelect & " Requestor Code " & v_txtRequestorID
' end if 

' if  len(trim(v_txtRowStateId)) > 0 then
'     txtSQL =txtSQL &  " and a.ROWSTATEID='" & v_txtRowStateId & "'"
' 	txtSelect = txtSelect & " Status  " & v_txtRowStateId	
' end if

	txtSQL = txtSQL & " Order by SALESCONTRACTNO " 'ditambah ukie tgl 9 okt 2013 untuk Order by


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
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>DELIVERY ORDER</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none"><%=txtSelect%></td>
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
								<th>Sales Contract No</th>
								<th>Delivery Order No</th>
								<th>Delivery Order Date</th>
								<th>Product</th>
								<th>Quantity</th>
								<th>UOM</th>
								<th>Start</th>
								<th>End</th>
								<th>Quantity Delivery Order</th>
								<th>Transporter Contract</th>
								<th>Transporter</th>
								<th>Transportation Cost</th>
								<th>Loss Tolerance (%)</th>
							</tr>
                            </thead>
                            <tbody>
								<%
								Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
								v_Quantity=0
								v_Amount = 0
								While Not rsCat.EOF%>


									<tr>
									</td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("salescontractno")%> </td>
										<td align="center" style="<%=cmdstyle%>">
										<a href="javascript:launchwin('printDeliveryOrder.Asp?DONo=<%=rscat("deliveryorderno")%>&fcba=<%=rscat("fcba")%>&transcode=<%=rscat("transporter")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
										<font face="Verdana" size="1"><%=rscat("deliveryorderno")%></font>
										</a>

										</td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("deliveryorderdate")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("productcode")%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("quantity")),0)%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("unitofmeasure")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("startdate")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("enddate")%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("quantity_do")),0)%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportercontractno")%> </td>		
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transporter")%> - <%=rscat("transportername")%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportationcost")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("losstolerance")%> </td>
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
	</body>	
</HTML>