<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice TBS</title>
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

v_txtINVOICEDATE_From= trim(request("INVOICEDATE_From"))
v_txtINVOICEDATE_To= trim(request("INVOICEDATE_To"))
v_txtSUPPLIERCODE= trim(request("SUPPLIERCODE"))
v_txtINVOICECODE= trim(request("INVOICECODE"))
v_txtITEMCODE= trim(request("ITEMCODE"))

businessunit= trim(request("businessunit"))
company     = trim(request("company"))

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

	txtSQL = "SELECT spo.invoiceno,spo.pocode,spo.suppliername as suppliercode,customer_api.get_description(spo.suppliername, spo.fcba) suppliername, " & _
	" bu.fccompanycode,spo.fcba,spo.invoice_by,spo.valid_from,spo.valid_until,spo.transdate,spo.no_faktur, " & _
	" spo.fakturdate,spo.pocode,spo.invoice_type, spo.ppn,spo.pph,spo.totalppn,(spo.totalpph+spo.grandtotal) as total," & _
	" spo.totalpph,spo.grandtotal " & _
	" FROM invoiceservicehdr spo " & _
	" INNER JOIN businessunit bu ON (bu.fccode = spo.fcba) " & _
	" WHERE spo.fcba in (SELECT fccode from BusinessUnit where fccompanycode='" & company & "' and fccode like '%"& businessunit &"%')" & _
	" AND invoice_type = 'TBS'" 
			

if  len(trim(v_txtINVOICEDATE_From)) > 0 then
    txtSQL = txtSQL & " AND spo.transdate >= to_date('" & v_txtINVOICEDATE_From & "','mm/dd/yyyy')"
	txtSelect = " From " & FormatDateTime(v_txtINVOICEDATE_From,1)
end if
if  len(trim(v_txtINVOICEDATE_To)) > 0 then
    txtSQL =txtSQL &  " AND spo.transdate <= to_date('" & v_txtINVOICEDATE_To & "','mm/dd/yyyy')"
	txtSelect = txtSelect & " to " & FormatDateTime(v_txtINVOICEDATE_To,1)
end if
if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    txtSQL = txtSQL & " and spo.suppliername='" & v_txtSUPPLIERCODE & "'"
	txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
end if
if  len(trim(v_txtINVOICECODE)) > 0 then
    txtSQL =txtSQL &  " and spo.invoiceno='" & v_txtINVOICECODE & "'"
	txtSelect = txtSelect & " Inv. Code " & v_txtINVOICECODE
end if 
					
txtsql=(txtsql) & 	" ORDER BY spo.transdate," & _
					"          spo.invoiceno "  
					
					
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
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>INVOICE TBS</h2>
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
							<%	
							Set rsCat = connStr.Execute(txtSQL)
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr> 
						<tr>
							<th>Date</th>
							<th>Invoice Code</th>
							<th>Invoice By</th>
							<th>Supplier Code</th>
							<th>Supplier Name</th>
							<th>Valid From</th>
							<th>Valid Until</th>
							<th>Ppn (%)</th>	  
							<th>Total PPn</th>
							<th>Total Harga</th>
							<th>PPh (%)</th>
							<th>Total PPh</th>
							<th>Grand Total</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim v_Quantity,v_price,v_amount
							v_totalharga=0 
							v_totalppn=0 
							v_totalpph=0 
							v_grandtotal=0 

							While Not rsCat.EOF%>
						<tr>
							 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
								<a href="javascript:launchwin('printinvoicetbs.asp?invoiceno=<%=rscat("invoiceno")%>&transdate=<%=rscat("transdate")%>&fcba=<%=rscat("fcba")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1"><%=rscat("invoiceno")%></font> 
								</a>
						 	</td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoice_by")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliercode")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliername")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("valid_from")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("valid_until")%></td>  
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("ppn"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>">
								<font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("totalppn"))),2)%>
							</td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1">
								<%=FormatNumber(cdbl(to_nol(rscat("total"))),0)%>
							</td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pph"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1">
								<%=FormatNumber(cdbl(to_nol(rscat("totalpph"))),0)%>
							</td>
							<td align="Right" style="<%=cmdstyle%>">
								<font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("grandtotal"))),2)%>
							</td> 
						</tr>
						  <%
						  ' v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
						v_totalharga= v_totalharga+Trim(cdbl(to_nol(rsCat("total"))))
						v_totalppn= v_totalppn+Trim(cdbl(to_nol(rsCat("totalppn"))))
						v_totalpph = v_totalpph+Trim(cdbl(to_nol(rsCat("totalpph"))))
						v_grandtotal =v_grandtotal +Trim(cdbl(to_nol(rsCat("grandtotal")))) 
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
							<tr bgcolor="#F4F4A0" style="font-weight: bold;">
								<td align="Center" style="<%=cmdstyleLeft%>" colspan="8"> <font face="Verdana" size="1">T o t a l</td> 
								<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_totalppn)),2)%></td>
								<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_totalharga)),2)%></td>
								<td></td>
								<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl((v_totalpph)),2)%></td>
								<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1">
									<%=FormatNumber(cdbl(to_nol(v_grandtotal)),0)%>
								</td>
							</tr>
                           </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	