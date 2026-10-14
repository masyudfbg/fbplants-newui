<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim fccode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName ,txtJenis, cmdview,rsdtlaging
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
				sqlexec     = "UPDATE invoice SET postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE  invoicecode='" & request("invoiceno_"&TRIM(n)) & "' AND fcba='" & request("fcba_"& TRIM(n))  & "'"
				'response.write "1-"& sqlexec & "<br>"
				connStr.execute(sqlexec)
				
			 	sqlexec 	= "BEGIN fac_journal_invpurch_api.prosesbyNotran('" & request("invoiceno_"&TRIM(n)) & "', '" & request("fcba_"&TRIM(n)) & "','SYSTEM'); END;"
            '     'response.write "2-"& sqlexec & "<br>"
			 	connStr.execute(sqlexec)
				
			 	sqlexec = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & request("fcba_"&TRIM(n))  & "' AND docno='" & request("invoiceno_"&TRIM(n)) & "'"
			' 	'response.write "3-"& sqlexec & "<br>"
			 	connStr.execute(sqlexec)
				
				'response.write "Y-"& n & "-" & request("cb_"&TRIM(n)) & "<br>"
			 end if
	next
		sqlexec = "BEGIN fac_closing_accounting_api.ProsesClosing(" & tahun& "," & bulan & ",'" & fcba & "','System'); " & _ 
     		     "END;"													 
				connStr.execute(sqlexec)

end if

%>

<html>
	<head>
		<title>FBPlants - IPO VIEW</title>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : AP PURCHASING</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form action="p_detail_agingap.Asp" method="post" name="eForm" >
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
		
		<input type=Submit name="txtJenis" id="btnsave" value="POSTING" onClick="return confirm('Anda yakin untuk melakukan Posting terhadap Invoice yang sudah dipilih?')">&nbsp;&nbsp;
		<a class="button_menu_top" onclick="getExport()" ><img  src="Images/butt_export_xls.png" border="0" ><span> Export XLS</span></a>&nbsp;&nbsp;
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
						<li class="TabbedPanelsTab">INV.PO BULAN INI</li>
						<li class="TabbedPanelsTab">PAYMENT BULAN INI</li>
						<li class="TabbedPanelsTab">PAYMENT ALL</li>
						<li class="TabbedPanelsTab">OUTSTANDING LIST S.D BULAN INI</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

<!-- PURCHASING -->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">		
									<th>Invoice By</th>							
									<th>No.Invoice</th>
									<th>Tgl.Invoice</th>
									<th>No.PO</th>
									<th>Supplier</th>
									<th>Amount</th>
									<th>PPN</th>
									<th>Other Charge</th>
									<th>Total Amount</th>
									<th>Payment</th>
									<th><input type='checkbox' name='select-all' id='select-all'/>Posting</th>
									<th>Action</th>	
									
								</tr>

								<%							
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = "  SELECT inv.invoice_by, inv.invoicecode, to_char(inv.invoicedate,'dd/mm/yyyy')invoicedate,inv.suppliercode, inv.suppliercode || ' - ' || s.fcname as supplier, " & _
											 " inv.pocode, nvl(inv.amount,0)amount, nvl(inv.totalppn,0)totalppn, nvl(inv.amountvat,0)amountvat, nvl(inv.othercharge+inv.otherchargestock,0) othercharge, " & _ 
											 " nvl(inv.amountvat+inv.otherchargestock,0)totalamount,inv.fcba,nvl(inv.payment,0)payment,inv.postingflg " & _
											 " FROM invoice inv INNER JOIN supplier s ON (inv.suppliercode=s.fccode and inv.fcba=s.fcba) " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND to_number(to_char(inv.invoicedate,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(inv.invoicedate,'MM'))="& bulan & " " & _
											 " AND inv.suppliercode='" & fccode & "' " & _
											 " ORDER BY datecreated ASC, invoicecode DESC "
								' response.write sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
								no = 0
								totalamount = 0
								While Not rsdtlupah.EOF 
								
								%>
								<tr id="tru<%=no%>">
									
									<td class="ratatengah">
										<%=rsdtlupah("invoice_by")%>
									</td>
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
										<%=formatnumber(rsdtlupah("totalppn"),0)%>
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
										<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value="" onclick='handleClick(this);'>
										<%if rsdtlupah("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%>
										<input type="hidden" name="fcba_<%=NO%>" value="<%=rsdtlupah("fcba")%>">
										<input type="hidden" name="invoiceno_<%=NO%>" value="<%=rsdtlupah("invoicecode")%>">
									</td>
																		
									<td colspan=2 class="ratatengah">
									<a href="Invoice.Asp?fcCode=<%=rsdtlupah("invoicecode")%>&fcba=<%=rsdtlupah("fcba")%>&txtJenis=Edit">
									<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
									</a>&nbsp;&nbsp;&nbsp;<a href="bkm_jurnal_view.Asp?bkmno=<%=rsdtlupah("invoicecode")%>&tglbkm=<%=rsdtlupah("invoicedate")%>&fcba=<%=rsdtlupah("fcba")%>&tipetran=IPO" target="_blank">
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
								        <input Type="Hidden" Value="<%=no%>" Name="no">
										<td colspan="8" class="tebal">Total</td>
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
									<th>Tgl.Voucher</th>
									<th>No.Invoice</th>
									<th>Tgl.Invoice</th>
									<th>Supplier</th>
									<th>Amount</th>
								</tr>
								<%
								'get data detail upah
								'dim sqldtlupah,rsdtlupah,no,totalamount
								sqldtlupah = " select t.vouchercode,t.datecreated,t.invdate,t.remarks,t.invoicecode,sum(t.totalamount)totalamount " & _ 
								" from ( " & _
								
								"		SELECT a.vouchercode,to_char(a.datecreated,'dd/mm/yyyy')datecreated, to_char(i.invoicedate,'dd/mm/yyyy') invdate," & _
								"		a.remarks,d.invoicecode,case when d.amount>0 then d.amount else 0 end totalamount " & _
								"		FROM paymentvoucher a   " & _
								"		LEFT JOIN paymentvoucherdetail d on (a.vouchercode=d.vouchercode and a.fcba=d.fcba)  " & _
								"		left join invoice i on (d.invoicecode=i.invoicecode and i.invoice_by=a.fcba) " & _
								"		where i.fcba ='" & fcba & "' AND (d.locationcode = '" & fccode & "' or d.CHARGECODE  ='" & fccode & "' ) " & _
								"		AND to_number(to_char(a.datecreated,'YYYY'))= " & tahun & "   " & _
								"		AND to_number(to_char(a.datecreated,'MM'))=" & bulan & " and a.totalamount>0 and a.Rowstateid in('Confirmed','AuthorizeOnProgress','Approved')" & _
								" )t " & _
								" GROUP BY t.vouchercode,t.datecreated,t.invdate,t.remarks,t.invoicecode " & _
								" UNION ALL " & _
								" select t.vouchercode,t.datecreated,t.invdate,t.remarks,t.invoicecode,sum(t.totalamount)totalamount " & _ 
								" from ( " & _
								
								"		SELECT a.vouchercode,to_char(a.datecreated,'dd/mm/yyyy')datecreated, to_char(i.invoicedate,'dd/mm/yyyy') invdate, " & _
								"		a.remarks,d.invoicecode,case when d.amount>0 then d.amount else 0 end totalamount " & _
								"		FROM paymentvoucher a   " & _
								"		LEFT JOIN paymentvoucherdetail d on (a.vouchercode=d.vouchercode and a.fcba=d.fcba)  " & _
								"		left join invoicedphdr i on (d.invoicecode=i.invoicecode and i.invoice_by=a.fcba) " & _
								"		where i.fcba ='" & fcba & "' AND i.suppliercode='" & fccode & "' " & _
								"		AND to_number(to_char(a.datecreated,'YYYY'))= " & tahun & "   " & _
								"		AND to_number(to_char(a.datecreated,'MM'))=" & bulan & " and a.totalamount>0 AND a.Rowstateid in('Confirmed','AuthorizeOnProgress','Approved')" & _
								" )t " & _
								"GROUP BY t.vouchercode,t.datecreated,t.invdate,t.remarks,t.invoicecode " & _
								"ORDER BY datecreated ASC, vouchercode DESC "

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
										<%=rsdtlupah("invdate")%>
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

						<!-- PAYMENT ALL-->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id="product-table">
								<tr id="trheader3">									
									<th>No.Invoice</th>
									<th>Tgl.Invoice</th>
									<th>No.PO</th>
									<th>Supplier</th>
									<th>Amount</th>
									<th>PPN</th>
									<th>Other Charge</th>
									<th>Total Amount</th>
									<th>Payment</th>
									<th>Outstanding</th>
								</tr>
                                <% 
								
								
								'get data detail upah
								dim paymentAllSQl,rsdtlpaymentAll,no_,totalamount_,postingflg_
								if cmdview = "TRUE" then
								paymentAllSQl = "SELECT a.INVOICECODE, a.INVOICEDATE tgl, a.POCODE,a.fcba, " &_
													" a.suppliercode supp_code,supplier_api.get_supp_name(a.suppliercode,a.fcba) supp_name, ''  FCTYPE, " &_
													" a.amount, a.totalppn ppn, a.OTHERCHARGESTOCK othercharge, (a.AMOUNTVAT+a.OTHERCHARGESTOCK) totalamount, a.payment, 0 outstanding " &_
													" FROM invoice a " &_
													" inner join supplier b on a.suppliercode = b.fccode " &_
													" where a.fcba='" & fcba & "'" & _
													" AND a.suppliercode = '" & fccode & "' " &_
													" GROUP BY a.INVOICECODE, a.INVOICEDATE, a.POCODE, a.fcba, a.suppliercode, supplier_api.get_supp_name(a.suppliercode,a.fcba), FCTYPE, a.amount, "&_
													" a.totalppn, a.OTHERCHARGESTOCK, (a.AMOUNTVAT+a.OTHERCHARGESTOCK), a.payment " &_

													" UNION ALL " &_

													" SELECT a.INVOICENO, a.transdate tgl, '' POCODE, a.fcba, " &_
													" a.suppliername supp_code,supplier_api.get_supp_name(a.suppliername,a.fcba) supp_name,'' FCTYPE, " &_
													" 0 amount, 0 ppn, 0 othercharge, a.grandtotal totalamount, a.payment, 0 outstanding " &_
													" FROM invoiceservicehdr a " &_
													" inner join supplier b on a.suppliername = b.fccode " &_
													" WHERE a.fcba = '" & fcba & "' AND a.suppliername = '" & fccode & "'" &_
													" GROUP BY a.INVOICENO, a.transdate, POCODE, a.fcba, a.suppliername, supplier_api.get_supp_name(a.suppliername,a.fcba),FCTYPE, 0, 0, 0, a.grandtotal, a.payment " &_
													" ORDER BY tgl DESC"
                           	    ' Response.Write paymentAllSQl
								set rsdtlpaymentAll = connStr.Execute(paymentAllSQl)
							
																
								no_ = 0
								totalamount_ = 0
								While Not rsdtlpaymentAll.EOF 
								
								%>
								<tr id="tru<%=no_%>">
									
									<td class="ratatengah">
										<%=rsdtlpaymentAll("invoicecode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlpaymentAll("tgl")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlpaymentAll("pocode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlpaymentAll("supp_name")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlpaymentAll("amount"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlpaymentAll("ppn"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlpaymentAll("othercharge"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlpaymentAll("totalamount"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlpaymentAll("payment"),0)%>
									</td>
									<td>
										<%=formatnumber(formatnumber(rsdtlpaymentAll("totalamount"),0) - formatnumber(rsdtlpaymentAll("payment"),0), 0)%>
									</td>
								</tr>
								<%

									no = no_ +1
									totalamount_ = cdbl(totalamount_) + cdbl(rsdtlpaymentAll("totalamount"))
									rsdtlpaymentAll.MoveNext
									Wend
								%>
								<tr class="ratakanan">
									<td colspan="7" class="tebal">Total</td>
									<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
								</tr>
								<%end if%>
							</table>
						</div>

						<!-- OUSTANDING LIST-->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id="product-table">
								<tr id="trheader3">									
									<th>No.Invoice</th>
									<th>Tgl.Invoice</th>
									<th>No.PO</th>
									<th>Supplier</th>
									<th>Amount</th>
									<th>PPN</th>
									<th>Other Charge</th>
									<th>Total Amount</th>
									<th>Payment</th>
									<th>Outstanding</th>
								</tr>

								<% 
																
								'get data detail upah
								
								if cmdview = "TRUE" then
								sqldtlupah = "  SELECT inv.invoicecode, to_char(inv.invoicedate,'dd/mm/yyyy')invoicedate,inv.suppliercode, inv.suppliercode || ' - ' || s.fcname as supplier, " & _
											 " inv.pocode, nvl(inv.amount,0)amount, nvl(inv.totalppn,0)totalppn, nvl(inv.amountvat,0)amountvat, nvl(inv.othercharge,0) othercharge, " & _ 
											 " nvl(inv.totalamount,0)totalamount,inv.fcba,inv.postingflg, " & _
											 " NVL(( " & _
											"	SELECT SUM(dtl.amount) " & _
											"	FROM paymentvoucher hdr " & _
											"	INNER JOIN paymentvoucherdetail dtl ON hdr.vouchercode = dtl.vouchercode  " & _
											"	WHERE dtl.invoicecode = inv.invoicecode AND inv.suppliercode='" & fccode & "' " & _
											"	AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'YYYYMM')) <= "& tahun&bulan & " AND hdr.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
											"	), 0) AS payment " & _
											 " FROM invoice inv INNER JOIN supplier s ON (inv.suppliercode=s.fccode and inv.fcba=s.fcba) " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND to_number(to_char(inv.invoicedate,'YYYY'))<="& tahun & " " & _ 
											 " AND to_number(to_char(inv.invoicedate,'MM'))<="& bulan & " " & _
											 " AND inv.suppliercode='" & fccode & "' " & _
											 "  AND (inv.totalamount -   NVL((SELECT SUM(dtl.amount) " & _
											 "	FROM paymentvoucher hdr " & _
											 "  INNER JOIN paymentvoucherdetail dtl ON hdr.vouchercode = dtl.vouchercode " & _
											 "	WHERE dtl.invoicecode = inv.invoicecode  AND inv.suppliercode='" & fccode & "' AND hdr.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
											 "	AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'YYYYMM')) <= " & tahun&bulan & " " & _
											 "	), 0)) > 1" & _
											 " ORDER BY datecreated ASC, invoicecode DESC "
						'		Response.Write sqldtlupah
								set rsdtlaging = connStr.Execute(sqldtlupah)
								DIM grnd,pay,txtoustanding
								dim totalpayment,totalending
																
								no = 0
								totalamount = 0
								totalpayment = 0
								totalending = 0
								While Not rsdtlaging.EOF 
									grnd = rsdtlaging("totalamount")
									pay = rsdtlaging("payment")
									txtoustanding = cdbl(grnd)-cdbl(pay)
								%>
								<tr id="tru3<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlaging("invoicecode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlaging("invoicedate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlaging("pocode")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlaging("supplier")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlaging("amount"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlaging("totalppn"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlaging("othercharge"),0)%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlaging("totalamount"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlaging("payment"),0)%>
									</td>
									<td>
										<%=formatnumber(txtoustanding,0)%>
									</td>
			                       
																		
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlaging("totalamount"))
									totalpayment = cdbl(totalpayment) + cdbl(rsdtlaging("payment"))
									totalending = cdbl(totalending) + cdbl(txtoustanding)
									rsdtlaging.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								      
										<td colspan="7" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalpayment,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalending,0)%>"/></td>
								</tr>
								<%end if%>
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
<script src="/distexcel/tableToExcel.js"></script>
<script src="/distexcel/jquery.tableToExcel.js"></script>
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

function getExport()
	{
		

		var periode 	= "<%=bulan& " - " &tahun%>";
		var supname 	= $("#suppliername").val();
		var nm = 'Ap_Pch_ ' + periode + ' - ' + supname + ".xlsx";  
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "Aging Purchasing"
            }
          });

	}

</script>
