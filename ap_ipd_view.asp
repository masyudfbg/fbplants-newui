<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim fccode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName ,txtJenis, cmdview  
fccode= trim(request("fccode"))
bulan= trim(request("month"))
tahun= trim(request("tahun"))
jobcode= trim(request("jobcode"))
fcba = trim(request("fcba"))
cmdview = "TRUE"

txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME") 
end if

' get data header  to_date('" & IssueDate & "','mm/dd/yyyy')
dim sqlheader,rsHeader,suppliername,n,sqlexec
sqlheader = "select fccode, fcname from supplier where fccode='" & fccode & "' AND fcba='" & fcba & "' "
set rsHeader = connStr.Execute(sqlheader)
'response.write sqlheader
if not rsHeader.eof then 
	
	suppliername = rsHeader("fcname")
	
end if

if Request("txtJenis")= "POSTING" then
	For n=0 to (cdbl(request("no"))-1)
	'response.write "Y-"& n & "-" & request("cb_"&TRIM(n)) & "<br>"

			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				sqlexec     = "UPDATE purchaseinvoicehdr SET postingflg='1' WHERE  invoicecode='" & request("invoiceno_"&TRIM(n)) & "' AND fcba='" & request("fcba_"& TRIM(n))  & "'"
				response.write "1-"& sqlexec & "<br>"
				connStr.execute(sqlexec)

				sqlexec 	= "BEGIN fac_journal_inv_ipd_api.prosesbyNotran('" & request("invoiceno_"&TRIM(n)) & "', '" & request("fcba_"&TRIM(n)) & "','SYSTEM'); END;"
                response.write "2-"& sqlexec & "<br>"
				connStr.execute(sqlexec)
				
				sqlexec = "update BKMJOURNALHDR set postingflg=1 where fcba='" & request("fcba_"&TRIM(n))  & "' AND docno='" & request("invoiceno_"&TRIM(n)) & "'"
				response.write "3-"& sqlexec & "<br>"
				connStr.execute(sqlexec)
				
			
			end if
	next
		sqlexec = "BEGIN fac_closing_accounting_api.ProsesClosing(" & tahun& "," & bulan & ",'" & fcba & "','System'); " & _ 
     		     "END;"													 
				connStr.execute(sqlexec)

end if

%>

<html>
	<head>
		<title>FBPlants - IPD VIEW</title>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : AP PURCHASING PRODUCT</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form action="ap_ipd_view.Asp" method="post" name="eForm" >
	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >SUPPLIER CODE</td>
			<td align="left">
				<input type="text" name="fccode" readonly id="fccode" size="20" value="<%=fccode%>" >
				<input type="hidden" name="fcba" id="fcba" size="20" value="<%=fcba%>" >
				<input type="hidden" name="month" id="month" size="20" value="<%=bulan%>" >
				<input type="hidden" name="tahun" id="tahun" size="20" value="<%=tahun%>" >
			</td>
			<td align="left" >PERIODE</td>
			<td align="left">
				<%=bulan%>/<%=tahun%> 
			</td>
			
		</tr>
		
		<tr>
			<td align="left" >SUPPLIER NAME</td>
			<td align="left" colspan="">
				<input type="text" name="suppliername" readonly id="suppliername" size="50" value="<%=suppliername%>" >				
			</td>						
		</tr>
		
		<tr>
		<td colspan="4" align="left">
		<!--<a class="button_menu_top" onclick="createData('PYV')" ><img src="ICON/Bank.png" border="0" ><span> New Payment </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData('RCV')" ><img src="ICON/Bank.png" border="0" ><span> New Receive </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData()" ><img  src="ICON/Bank.png" border="0" ><span> New Pindah Buku</span></a>&nbsp;	-->	
		<input type=Submit name="txtJenis" id="btnsave" value="POSTING" onClick="return confirm('Anda yakin untuk melakukan Posting terhadap Invoice yang sudah dipilih?')">&nbsp;&nbsp;
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
						<li class="TabbedPanelsTab">INV PRODUCT BULAN INI</li>
						<li class="TabbedPanelsTab">PAYMENT LIST BULAN INI</li>
						<li class="TabbedPanelsTab">OUTSTANDING PAYMENT LIST S.D BULAN INI</li>
						
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
									<th>Total Amount</th>
									<th><input type='checkbox' name='select-all' id='select-all'/>Posting</th>
									<th>Action</th>	
									
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = "  SELECT inv.invoicecode, to_char(inv.transdate,'dd/mm/yyyy')invoicedate,inv.suppliercode, inv.suppliercode || ' - ' || s.fcname as supplier, " & _
											 " inv.suppliercontractno pocode, nvl(inv.amount_beforevat,0)amount_beforevat,  nvl(inv.amountvat,0)amountvat, " & _ 
											 " nvl(inv.grandtotal,0)grandtotal,inv.fcba,inv.postingflg " & _
											 " FROM purchaseinvoicehdr inv INNER JOIN supplier s ON (inv.suppliercode=s.fccode and inv.fcba=s.fcba) " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND to_number(to_char(inv.transdate,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(inv.transdate,'MM'))="& bulan & " " & _
											 " AND inv.suppliercode='" & fccode & "' " & _
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
										<%=formatnumber(rsdtlupah("amount_beforevat"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("amountvat"),0)%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("grandtotal"),0)%>
									</td>
									<td>
										<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
										<%if rsdtlupah("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%>
										<input type="hidden" name="fcba_<%=NO%>" value="<%=rsdtlupah("fcba")%>">
										<input type="hidden" name="invoiceno_<%=NO%>" value="<%=rsdtlupah("invoicecode")%>">
									</td>
																		
									<td colspan=2 class="ratatengah">
									<a href="purchaseinvoice.Asp?invoicecode=<%=rsdtlupah("invoicecode")%>&fcba=<%=rsdtlupah("fcba")%>&txtJenis=Edit">
									<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
									</a>&nbsp;&nbsp;&nbsp;<a href="bkm_jurnal_view.Asp?bkmno=<%=rsdtlupah("invoicecode")%>&tglbkm=<%=rsdtlupah("invoicedate")%>&fcba=<%=rsdtlupah("fcba")%>&tipetran=IPD" target="_blank">
									<img src="IMAGES/view.png" target="_blank"  alt="Posting Data : <%=rsdtlupah("invoicecode")%>" border="0">
									</a>
									</td>
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("grandtotal"))
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								        <input Type="Hidden" Value="<%=no%>" Name="no">
										<td colspan="5" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
								</tr>
								<%end if%>
							</table>
						</div>
						
						<!-- PAYMENT LIST-->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No.Voucher</th>
									<th>Date</th>
									<th>No.Invoice</th>
									<th>Supplier</th>
									<th>Amount</th>
								</tr>
								<%
								'get data detail upah
								'dim sqldtlupah,rsdtlupah,no,totalamount
								sqldtlupah = " select t.vouchercode,t.datecreated,t.remarks,t.invoicecode,sum(t.totalamount)totalamount " & _ 
								" from ( " & _
								
								"		SELECT a.vouchercode,to_char(a.datecreated,'dd/mm/yyyy')datecreated, " & _
								"		a.remarks,d.invoicecode,case when d.amount>0 then d.amount else 0 end totalamount " & _
								"		FROM paymentvoucher a   " & _
								"		LEFT JOIN paymentvoucherdetail d on (a.vouchercode=d.vouchercode and a.fcba=d.fcba)  " & _
								"		left join purchaseinvoicehdr i on (d.invoicecode=i.invoicecode and i.invoice_by=a.fcba) " & _
								"		where i.fcba ='" & fcba & "' AND (d.locationcode = '" & fccode & "' or d.CHARGECODE  ='" & fccode & "' ) " & _
								"		AND to_number(to_char(i.transdate,'YYYY'))= " & tahun & "   " & _
								"		AND to_number(to_char(i.transdate,'MM'))=" & bulan & " and a.totalamount>0 " & _
								" )t " & _
								" GROUP BY t.vouchercode,t.datecreated,t.remarks,t.invoicecode " & _
								" ORDER BY datecreated ASC, vouchercode DESC "

                          	'Response.write sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
								no = 0
								totalamount = 0
								While Not rsdtlupah.EOF 
									
								%>
								<tr id="tru<%=no%>">
									<td class="ratakiri">
										<%=rsdtlupah("vouchercode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("datecreated")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("invoicecode")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("remarks")%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("totalamount"),0)%>
									</td>
								</tr>
									<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("totalamount"))
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
										<td colspan="4" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
								</tr>

							</table>
						</div>
						<!-- OUSTANDING LIST-->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id="product-table">
								<tr id="trheader3">									
									<th>No. Invoice</th>
									<th>Date</th>
									<th>PPN</th>
									<th>Grand Total</th>
									<th>Payment</th>
									<th>Outstanding</th>
								</tr>

								<% 
								
								'get data detail upah
							'	dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = "SELECT spo.invoicecode,spo.suppliercode,spo.postingflg, " & _
											" bu.fccompanycode, spo.fcba, spo.invoice_by, spo.valid_from, spo.valid_until, " & _
											" TO_CHAR(spo.TRANSDATE, 'dd/mm/yyyy') AS transdate, spo.no_faktur," & _
											" spo.fakturdate, spo.vat, spo.amountvat, (spo.amountvat + spo.grandtotal) AS total, spo.grandtotal, " & _ 
											" NVL(( " & _ 
											"	SELECT SUM(dtl.amount) " & _ 
											"	FROM paymentvoucher hdr " & _ 
											"	INNER JOIN paymentvoucherdetail dtl ON hdr.vouchercode = dtl.vouchercode " & _ 
											"	WHERE dtl.invoicecode = spo.invoicecode  AND hdr.rowstateid in ('Confirmed','AuthorizeOnProgress','Approved') " & _ 
											"	AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'YYYYMM')) <= " & tahun&bulan & "), 0) AS paymentview " & _
										
											" FROM purchaseinvoicehdr spo " & _
											" INNER JOIN businessunit bu ON (bu.fccode = spo.fcba) " & _
											" INNER JOIN supplier spod on (spod.FCCODE = spo.SUPPLIERCODE) " & _
											" WHERE spo.fcba = '" & fcba & "'" & _ 
											" AND to_number(to_char(spo.transdate,'YYYYMM'))<="& tahun&bulan & " " & _ 
											" AND spo.SUPPLIERCODE='" & fccode & "' " & _
											" AND (spo.grandtotal - NVL(( " & _
                                            "                SELECT SUM(dtl.amount) " & _
                                            "                FROM paymentvoucher hdr " & _
                                            "                INNER JOIN paymentvoucherdetail dtl ON hdr.vouchercode = dtl.vouchercode " & _
                                            "                WHERE dtl.invoicecode = spo.invoicecode  AND hdr.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
                                            "                AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'YYYYMM')) <= " & tahun&bulan & " " & _
                                            "               ), 0)) > 1 " & _
											" GROUP BY " & _
											" spo.invoicecode, spo.suppliercode, spo.postingflg, bu.fccompanycode, spo.fcba, spo.invoice_by, " & _
											" spo.valid_from, spo.valid_until, spo.TRANSDATE, spo.no_faktur, spo.fakturdate, spo.vat, spo.amountvat, spo.grandtotal " & _
						     				" ORDER BY TRANSDATE ASC, invoicecode ASC "
								
								'Response.Write sqldtlupah
                           	
								set rsdtlupah = connStr.Execute(sqldtlupah)
								dim grnd,pay,totalpayment,txtoustanding,totaloustanding
																
								no = 0
								totalamount = 0
								totalpayment = 0
								totaloustanding = 0
								While Not rsdtlupah.EOF 
									grnd = rsdtlupah("grandtotal")
									pay = rsdtlupah("paymentview")
									txtoustanding = cdbl(grnd)-cdbl(pay)
									
								%>
								<tr id="tru3<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("invoicecode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("TRANSDATE")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTVAT"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("total"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("paymentview"),0)%>
									</td>
									<td>
										<%=formatnumber(txtoustanding,0)%>
									</td>
			                        
															
									
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("GRANDTOTAL"))
									totalpayment = cdbl(totalpayment) + cdbl(rsdtlupah("paymentview"))
									totaloustanding = cdbl(totaloustanding) + cdbl(txtoustanding)
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								       
										<td colspan="3" class="tebal">Total</td>
										<td class="tebal">
											<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/>
										</td>
										<td class="tebal">
											<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalpayment,0)%>"/>
										</td>
										<td class="tebal">
											<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totaloustanding,0)%>"/>
										</td>
								</tr>
							</table>
							<%end if%>
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

<script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 1){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
	}
	

</script>
<script type="text/javascript">	
		$('#select-all').click( function () {
			$('.item').prop('checked', this.checked);
			$('.item').val(this.checked ? 1 : 0);
		})
</script>
	
<script type="text/javascript" LANGUAGE="VBScript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
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
