<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice</title>
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
<script language="JavaScript">
$(document).ready(function(){
    $(".pesan").fadeIn().delay(3000).fadeOut();
	alert("tes");
});
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

v_txtINVOICEDATE_From= trim(request("TRANSDATE_From"))
v_txtINVOICEDATE_To= trim(request("TRANSDATE_To"))
v_txtSUPPLIERCODE= trim(request("SUPPLIERCODE"))
v_txtINVOICECODE= trim(request("INVOICECODE"))
v_txtITEMCODE= trim(request("ITEMCODE"))

businessunit= trim(request("businessunit"))
txtjenis= trim(request("txtjenis"))
cmdView= trim(request("cmdView"))

dim total,cmdView
Dim sqlview,rsview ,q
if txtjenis = "Delete" then 
	q = "select count(vouchercode) as total  from paymentvoucherdetail where invoicecode='" & v_txtINVOICECODE & "' "
	'Response.Write q
	'Response.End
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   total = rsview("total")
	end if
	if cint(total)> 0 then 
		err.number = 1001
		err.description = "No Invoice """ & v_txtINVOICECODE & """, tidak bisa dihapus karena sudah ada pembayaran."
		Response.Write "<div style='text-align:center; font-weight: bold' class = ""pesan"">" & Err.Description & "</div>"
		cmdView	= "True"
		
	else 
		q = "delete from INVOICEDPDETAIL where invoicecode='" & v_txtINVOICECODE & "' "
		set rsview = connStr.Execute(q)
		q = "delete from invoicedphdr where invoicecode='" & v_txtINVOICECODE & "' "
		set rsview = connStr.Execute(q)
		err.description = "No Invoice """ & v_txtINVOICECODE & """, sukses dihapus"
		Response.Write "<div style='text-align:center; font-weight: bold' class = ""pesan"">" & Err.Description & "</div>"
		'response.redirect("p_paymentrequest.asp")
		cmdView	= "True"
	end if
else 
	cmdView	= request("cmdView")
end if

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   fcbaname = rsview("fcName")
	   CompanyName=rsview("FCCOMPANYNAME")
	   MgrName = rsview("FCMANAGER")
	   KTUName = rsview("FCKTU")
	   CropName=rsview("FCCROP")
	   FcbaType=rsview("fcType")
	end if


sqlview = "select i.invoicedate,i.invoicecode,i.pocode,s.fcname,i.amount,i.totalppn,i.OTHERCHARGESTOCK,i.OTHERCHARGE,i.totalamount " & _
" from invoicedphdr i,supplier s " & _
" where s.FCCODE=i.SUPPLIERCODE and i.invoicedate between TO_DATE('" & v_txtINVOICEDATE_From & "','mm/dd/yyyy') and to_date('" & v_txtINVOICEDATE_To & "','mm/dd/yyyy') " & _
" AND i.fcba=s.fcba and i.fcba='" & businessunit & "'"
set rsview = connStr.Execute(sqlview)


 'response.Write(sqlview)
'response.Flush()
%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
		<input Type="Hidden" Value="True" Name="cmdView">
         <thead>
         <tr>
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>PAYMENT REQUEST</h2>
				</div>
			</div>
				</th>
			</tr>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
						
						
						</tr> 
						<tr>
							<th colspan="3">Action</th>
							<th>Date</th>
							<th>Invoice Code</th>
							<th>PO Code</th>
							<th>Supplier</th>	  
							<th>Amount</th>
							<th>PPn</th>
							<th>Other Charge Stock</th>
							<th>Other Charge Non Stock</th>
							<th>Total Amount Invoice</th>
						</tr>
                         </thead>

						 <%
							While Not rsview.EOF

						 %>

                         <tbody>
						<tr>
							<td>
								<a href="paymentrequestedit.asp?fcba=<%=businessunit%>&invoicecode=<%=rsview("invoicecode")%>">
								<img src="IMAGES/edit.gif" alt="Print Data : <%=rsview(1)%>" border="0">
								</a>
								
							</td>
							<td>
								
								<a href="p_paymentrequest.asp?fcba=<%=businessunit%>&invoicecode=<%=rsview("invoicecode")%>&txtjenis=Delete&cmdView=True">
								<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsview(1)%>" border="0">
								</a>
							</td>
							<td>
								
								<a href="prinpaymentreq.asp?fcba=<%=businessunit%>&invoicecode=<%=rsview("invoicecode")%>">
								<img src="IMAGES/print.gif" alt="Edit Data : <%=rsview(1)%>" border="0">
								</a>
							</td>
							 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("invoicedate")%></td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("invoicecode")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("POCODE")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("fcname")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("amount"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("totalppn"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("OTHERCHARGESTOCK"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("OTHERCHARGE"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("totalamount"))),2)%></td> 
						</tr>
						</tbody>
						<%
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	