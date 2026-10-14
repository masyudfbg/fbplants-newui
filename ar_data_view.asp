<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim fccode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName  
fccode= trim(request("fccode"))
bulan= trim(request("month"))
tahun= trim(request("tahun"))
jobcode= trim(request("jobcode"))
fcba = trim(request("fcba"))


q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME") 
end if

' get data header  to_date('" & IssueDate & "','mm/dd/yyyy')
dim sqlheader,rsHeader,suppliername 
sqlheader = "select customercode, description from customer where customercode='" & fccode & "' AND fcba='" & fcba & "' "
set rsHeader = connStr.Execute(sqlheader)
'response.write sqlheader
if not rsHeader.eof then 
	
	suppliername = rsHeader("description")
	
end if


%>

<html>
	<head>
		<title>FBPlants - SLS VIEW</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
	</head>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: 20px;
	}
	.posisiheader{
		position: relative;
		left: 100px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarabsen{
		width: 60px;
	}
	.lebaremployee{
		width: 230px;
	}
	.lebaractivity{
		width: 250px;
	}
	.lebaracllowedcode{
		width: 210px;
	}
	.lebarupah{
		width: 120px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}

	
</style>

<script>

</script>

<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : AR SALES</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" >
	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >CUSTOMER CODE</td>
			<td align="left">
				<input type="text" name="fccode" readonly id="fccode" size="20" value="<%=fccode%>" >
			</td>
			<td align="left" >PERIODE</td>
			<td align="left">
				<%=bulan%>/<%=tahun%> 
			</td>
			
		</tr>
		
		<tr>
			<td align="left" >CUSTOMER NAME</td>
			<td align="left" colspan="">
				<input type="text" name="suppliername" readonly id="suppliername" size="50" value="<%=suppliername%>" >				
			</td>						
		</tr>
		
		<tr>
		<td colspan="4" align="left">
		<!--<a class="button_menu_top" onclick="createData('PYV')" ><img src="ICON/Bank.png" border="0" ><span> New Payment </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData('RCV')" ><img src="ICON/Bank.png" border="0" ><span> New Receive </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData()" ><img  src="ICON/Bank.png" border="0" ><span> New Pindah Buku</span></a>&nbsp;	-->	
		<input type=button onclick="window.close();" name="btnclose" value="CLOSE">
		
		
		</td>
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='800' id='table1'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">INV.SLS</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

<!-- PURCHASING -->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No.Invoice</th>
									<th>Date</th>
									<th>No.Contract</th>
									<th>Supplier</th>
									<th>Amount</th>
									<th>PPN</th>
									<th>Other Charge</th>
									<th>Total Amount</th>
									<th>Payment</th>
									<th>Posting</th>
									<th>Action</th>	
									
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								sqldtlupah = " SELECT inv.invoicecode, to_char(inv.transdate,'dd/mm/yyyy')invoicedate,inv.customercode, inv.customercode || ' - ' || s.description as supplier, " & _
											 " inv.salescontractno pocode, nvl(inv.amount_beforevat,0)amount, nvl(inv.vat,0)totalppn, nvl(inv.amountvat,0)amountvat, 0 othercharge, " & _ 
											 " nvl(inv.grandtotal,0)totalamount,inv.fcba,nvl(inv.receiveamt,0)payment,'0' postingflg " & _
											 " FROM salesinvoicehdr inv INNER JOIN customer s ON (inv.customercode=s.customercode and inv.fcba=s.fcba) " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND to_number(to_char(inv.transdate,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(inv.transdate,'MM'))="& bulan & " " & _
											 " AND inv.customercode='" & fccode & "' " & _
											 " ORDER BY transdate ASC, invoicecode DESC "
                           	
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
								no = 0
								totalamount = 0
								While Not rsdtlupah.EOF 
									'postingflg = rsdtlupah("postingflg")
									'if postingflg = "0" then 
									'	postingflg="Tidak"
									'else 
									'	postingflg="YA"
									'end if
								%>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("invoicecode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("invoicedate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("pocode")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("supplier")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("amount"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("amountvat"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("othercharge"),0)%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("totalamount"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("payment"),0)%>
									</td>
			                        <td>
										<%if rsdtlupah("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%>
									</td>
																		
									<td colspan=2 class="ratatengah">
									<a href="salesinvoice.Asp?invoicecode=<%=rsdtlupah("invoicecode")%>&fcba=<%=rsdtlupah("fcba")%>&txtJenis=Edit">
									<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
									</a>&nbsp;&nbsp;&nbsp;<a href="salesinv_jurnal_view.Asp?invoicecode=<%=rsdtlupah("invoicecode")%>&tglinvoice=<%=rsdtlupah("invoicedate")%>&fcba=<%=rsdtlupah("fcba")%>&tipetran=SLS" target="_blank">
									<img src="IMAGES/view.png" target="_blank"  alt="Posting Data : <%=rsdtlupah("invoicecode")%>" border="0">
									</a>
									</td>
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("totalamount"))
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
										<td colspan="7" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
								</tr>
							</table>
						</div>
			
					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">
			
		
			 
			
		</tr>
	</table>

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript" LANGUAGE="VBScript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

$(function() {	
	$("#jurnaldate").datepicker();
	var postingflg =  $("#POSTINGFLG").val();
	if(postingflg==0){
		$("#btnsave").show();
		$("#btnunposting").hide();
	}else {
		$("#btnunposting").show();
		$("#btnsave").hide();
	}
	
		
});


function numformat(val) {
	var nStr = val.toString();
	nStr = nStr.replace(/\,/g, "");
	var x = nStr.split('.');
	var x1 = x[0];
	var x2 = x.length > 1 ? '.' + x[1] : '';
	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + '.' + '$2');
	}
	return x1 + x2;
}


</script>
