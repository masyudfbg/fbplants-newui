<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Purchase Shipment</title>
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

v_txtCREATIONDATE_From=trim(request("CREATIONDATE_From"))
v_txtCREATIONDATE_To=trim(request("CREATIONDATE_To"))
v_txtSALESCONTRACTNO=trim(request("SALESCONTRACTNO"))
v_txtDELIVERYORDERNO=trim(request("DELIVERYORDERNO"))
v_txtSALESSHIPMENTNO=trim(request("purchaseshipmentNO"))
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

txtSQL =  " SELECT ssd.wbt_chitno, ssd.postingdate, ssh.creationdate, ssd.vehicleno, ssd.transporter, tr.fcname as transportername, ssd.quantity, ssd.acceptquantity, " & _
		  " ssh.purchaseshipmentno, ssh.purchasecontractno, ssh.deliveryorderno, ssd.customer_chitno, ssd.acceptquantitybruto, (NVL(ssd.acceptquantity,0) - NVL(ssd.quantity,0)) variance" & _
		  " FROM purchaseshipmentdtl ssd" & _
		  " left JOIN purchaseshipmenthdr ssh"& _
		  " ON (ssd.purchaseshipmentno = ssh.purchaseshipmentno AND ssd.fcba = ssh.fcba)" & _
		  " left JOIN transporter tr ON (ssd.transporter = tr.fccode AND ssd.fcba = tr.fcba)" & _
		  " WHERE ssh.fcba = '" & businessunit & "'"  

' response.write(txtSQL)
if  len(trim(v_txtCREATIONDATE_From)) > 0 then
    txtSQL = txtSQL & " and CREATIONDATE >=to_Date('" & v_txtCREATIONDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtCREATIONDATE_From,1)	
end if
if  len(trim(v_txtCREATIONDATE_To)) > 0 then
    txtSQL =txtSQL &  " and CREATIONDATE <=to_Date('" & v_txtCREATIONDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtCREATIONDATE_To,1)	
end if
if  len(trim(v_txtSALESCONTRACTNO)) > 0 then
    txtSQL = txtSQL & " and ssh.purchasecontractno='" & v_txtSALESCONTRACTNO & "'"
	txtSelect = txtSelect & " Purchase Contract No. " & v_txtSALESCONTRACTNO	
end if
if  len(trim(v_txtDELIVERYORDERNO)) > 0 then
    txtSQL =txtSQL &  " and ssh.deliveryorderno='" & v_txtDELIVERYORDERNO & "'"
	txtSelect = txtSelect & " Delivery Order No. " & v_txtDELIVERYORDERNO	
end if 
if  len(trim(v_txtSALESSHIPMENTNO)) > 0 then
    txtSQL =txtSQL &  " and ssh.purchaseshipmentno='" & v_txtSALESSHIPMENTNO & "'"
	txtSelect = txtSelect & " Shipment No. " & v_txtSALESSHIPMENTNO
end if 

' if  len(trim(v_txtRowStateId)) > 0 then
'     txtSQL =txtSQL &  " and a.ROWSTATEID='" & v_txtRowStateId & "'"
' 	txtSelect = txtSelect & " Status  " & v_txtRowStateId	
' end if

	txtSQL = txtSQL & " Order by purchaseshipmentNO " 'ditambah ukie tgl 9 okt 2013 untuk Order by


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
					<h2>PURCHASE SHIPMENT</h2>
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
								<th>Shipment No</th>
								<th>Creation Date</th>
								<th>Purchase Contract No</th>
								<th>Delivery Order No</th>
								<th>Quantity Delivery Order</th>
								<th>Ticket No</th>
								<th>Date</th>
								<th>Transporter</th>
								<th>Transporter Name</th>
								<th>Vehicle No</th> 
								<th>Customer Ticket</th>
								<th>Customer Accept Quantity</th>
								<th>Variance</th>
							</tr>
                            </thead>
                            <tbody>
								<%
								Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
								v_Quantity=0
								v_Amount = 0
								v_acceptquantity = 0
								v_acceptquantitybruto = 0
								var= 0
								While Not rsCat.EOF%>


									<tr>
									</td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("purchaseshipmentno")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("creationdate")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("purchasecontractno")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("deliveryorderno")%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("quantity")),2)%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("wbt_chitno")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("postingdate")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transporter")%></td>
										<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportername")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vehicleno")%></td> 
										<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("customer_chitno")%></td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("acceptquantity")),2)%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("variance")),2)%> </td>
									</tr>
								  <%
								  v_Quantity = v_Quantity + cdbl(rscat("quantity"))
								  v_acceptquantity = v_acceptquantity +  cdbl(to_nol(rscat("acceptquantity")))
								  v_acceptquantitybruto = v_acceptquantitybruto +  cdbl(to_nol(rscat("acceptquantitybruto")))
								  var = var + cdbl(to_nol(rscat("variance")))
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
								<tr bgcolor="#F4F4A0" style="font-weight: bold;">
									<td colspan="4" align="center">TOTAL</td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(v_Quantity)),2)%></td>
									<td colspan="6" align="center"></td>	
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(v_acceptquantity)),2)%></td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(var)),2)%></td>
								</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</body>	
</HTML>