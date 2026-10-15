<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim fccode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName ,txtJenis, cmdview
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
 
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
dim sqlheader,rsHeader,transportername ,n,sqlexec
sqlheader = "select fccode, fcname from transporter where fccode='" & fccode & "' AND fcba='" & fcba & "' "
set rsHeader = connStr.Execute(sqlheader)
'response.write sqlheader
if not rsHeader.eof then 
	
	transportername = rsHeader("fcname")
	
end if

if Request("txtJenis")= "POSTING" then
	For n=0 to (cdbl(request("no"))-1)
	'response.write "Y-"& n & "-" & request("cb_"&TRIM(n)) & "<br>"
		if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
			sqlexec     = "UPDATE transporterinvoice SET postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE invoicecode='" & request("invoiceno_"&TRIM(n)) & "' AND fcba='" & request("fcba_"& TRIM(n))  & "'"
			'response.write "1-"& sqlexec & "<br>"
			'connStr.execute(sqlexec)
			ExecuteAndLog ScriptName, sqlexec
			
			sqlexec 	= "BEGIN fac_journal_inv_itr_api.prosesbyNotran('" & request("invoiceno_"&TRIM(n)) & "', '" & request("fcba_"&TRIM(n)) & "','SYSTEM'); END;"
			'response.write "2-"& sqlexec & "<br>"
			'connStr.execute(sqlexec)
			ExecuteAndLog ScriptName, sqlexec
			
			sqlexec = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & request("fcba_"&TRIM(n))  & "' AND docno='" & request("invoiceno_"&TRIM(n)) & "'"
			'response.write "3-"& sqlexec & "<br>"
			'connStr.execute(sqlexec)
			ExecuteAndLog ScriptName, sqlexec			
		end if
	next
		sqlexec = "BEGIN fac_closing_accounting_api.ProsesClosing(" & tahun& "," & bulan & ",'" & fcba & "','System'); " & _ 
     		     "END;"													 
				'connStr.execute(sqlexec)
				ExecuteAndLog ScriptName, sqlexec

end if

%>

<html>
	<head>
		<title>FBPlants - ITR VIEW</title>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : AP INVOICE TRANSPORT</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form action="ap_itr_view.Asp" method="post" name="eForm" >
	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >TRANSPORTER CODE</td>
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
			<td align="left" >TRANSPORTER NAME</td>
			<td align="left" colspan="">
				<input type="text" name="transportername" readonly id="transportername" size="50" value="<%=transportername%>" >				
			</td>						
		</tr>
		
		<tr>
		<td colspan="4" align="left">
		<!--<a class="button_menu_top" onclick="createData('PYV')" ><img src="ICON/Bank.png" border="0" ><span> New Payment </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData('RCV')" ><img src="ICON/Bank.png" border="0" ><span> New Receive </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData()" ><img  src="ICON/Bank.png" border="0" ><span> New Pindah Buku</span></a>&nbsp;	-->	
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
						<li class="TabbedPanelsTab">INV.TRANSPORTER  BULAN INI</li>
						<li class="TabbedPanelsTab">PAYMENT LIST  BULAN INI</li>
						<li class="TabbedPanelsTab">OUTSTANDING LIST S.D BULAN INI</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

<!-- INV TRANSPORTER -->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No. Invoice</th>
									<th>Date</th>
									<th>Transporter Contract</th>
									<th>Transporter Name</th>
									<th>Amount</th>
									<th>VAT</th>
									<th>PPh</th>
									<th>Grand Total</th>
									<th>Payment</th>
									<th><input type='checkbox' name='select-all' id='select-all'/>Posting</th>
									<th>Action</th>	
									
								</tr>

								<% 
								
								'get data detail upah
								Dim sqldtlupah, rsdtlupah, dictPay, dictShown, invcode, pay
								Dim AMOUNT, PPH, GRANDTOTAL, VAT
								Dim ttlamt, ttlpph, ttlgrand, ttlpayment, no, totalamount
								if cmdview = "TRUE" then
								sqldtlupah = "  SELECT distinct inv.invoicecode,inv.fcba, to_char(inv.CREATIONDATE,'dd/mm/yyyy') CREATIONDATE,inv.salesshipmentno,inv.postingflg, " & _
											 " inv.transportercontractno,tr.fccode || ' - ' || tr.fcname AS transportername, " & _ 
											 " nvl(inv.amount,0)amount,nvl(inv.vat,0)vat,nvl(inv.pph,0)pph,nvl(inv.amountafterpph,0)amountafterpph, case when pvd.amount>0 then pvd.amount else 0 end payment " & _
											 " FROM transporterinvoice inv  " & _
											 " LEFT JOIN transportercontract tc ON (tc.fcba = inv.fcba AND tc.transportercontractno = inv.transportercontractno) " & _
											 " LEFT JOIN paymentvoucherdetail pvd ON inv.invoicecode=pvd.invoicecode " & _
											 " LEFT JOIN paymentvoucher pv ON pv.VOUCHERCODE = pvd.VOUCHERCODE " & _
 											 " INNER JOIN transporter tr  ON (tr.fccode = tc.transporter AND tr.fcba = tc.fcba) " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND to_number(to_char(inv.CREATIONDATE,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(inv.CREATIONDATE,'MM'))="& bulan & " " & _
											 " AND tr.fccode='" & fccode & "' " & _
											 " GROUP BY inv.invoicecode, inv.fcba, TO_CHAR(inv.CREATIONDATE,'dd/mm/yyyy'),inv.salesshipmentno, inv.postingflg," & _
         									 " inv.transportercontractno,tr.fccode, tr.fcname,inv.amount,inv.vat,inv.pph,inv.amountafterpph, case when pvd.amount>0 then pvd.amount else 0 end " & _
											 " ORDER BY to_char(inv.CREATIONDATE,'dd/mm/yyyy') ASC, inv.invoicecode ASC "
								' response.write sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)

								' --- Dictionary untuk simpan total payment per invoice ---
								Set dictPay = Server.CreateObject("Scripting.Dictionary")
								Set dictShown = Server.CreateObject("Scripting.Dictionary")

								' --- Loop pertama: hitung payment per invoice ---
								Do While Not rsdtlupah.EOF
									invcode = rsdtlupah("invoicecode")
									pay     = CDbl(rsdtlupah("payment"))

									If dictPay.Exists(invcode) Then
										dictPay(invcode) = dictPay(invcode) + pay
									Else
										dictPay.Add invcode, pay
									End If

									rsdtlupah.MoveNext
								Loop

								' --- Reset recordset biar bisa dipakai render tabel ---
								rsdtlupah.Close
								Set rsdtlupah = connStr.Execute(sqldtlupah)

								ttlamt = 0
								ttlpph = 0
								ttlgrand = 0
								ttlpayment = 0
								no = 0

								While Not rsdtlupah.EOF 

									AMOUNT      = rsdtlupah("amount")
									PPH         = rsdtlupah("pph")
									GRANDTOTAL  = rsdtlupah("amountafterpph")
									VAT         = rsdtlupah("vat")
									invcode    = rsdtlupah("invoicecode")

								%>
								<% If Not dictShown.Exists(invcode) Then %>
									<tr id="tru<%=no%>">
										<td class="ratakiri"><%=rsdtlupah("invoicecode")%></td>
										<td class="ratakiri"><%=rsdtlupah("CREATIONDATE")%></td>
										<td class="ratakiri"><%=rsdtlupah("transportercontractno")%></td>
										<td class="ratakiri"><%=rsdtlupah("transportername")%></td>
										<td class="ratakanan"><%=formatnumber(AMOUNT,2)%></td>
										<td class="ratakanan"><%=VAT%></td>
										<td class="ratakanan"><%=formatnumber(PPH,2)%></td>
										<td class="ratakanan"><%=formatnumber(GRANDTOTAL,2)%></td>
										<td class="ratakanan">
											<% If Not dictShown.Exists(invcode) Then %>
												<%=FormatNumber(dictPay(invcode),0)%>
												<% dictShown.Add invcode, True %>
											<% Else %>
												
											<% End If %>
										</td>
										<td>
											<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
											<%if rsdtlupah("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%>
											<input type="hidden" name="fcba_<%=NO%>" value="<%=rsdtlupah("fcba")%>">
											<input type="hidden" name="invoiceno_<%=NO%>" value="<%=rsdtlupah("invoicecode")%>">
										</td>
										<td colspan=2 class="ratatengah">
											<a href="transporterinvoice.Asp?invoicecode=<%=rsdtlupah("invoicecode")%>&fcba=<%=rsdtlupah("fcba")%>&transportercontractno=<%=rsdtlupah("transportercontractno")%>&salesshipmentno=<%=rsdtlupah("salesshipmentno")%>&txtJenis=Edit">
											<img src="IMAGES/edit.gif" alt="Edit Data" border="0"></a>
											&nbsp;&nbsp;&nbsp;
											<a href="bkm_jurnal_view.Asp?bkmno=<%=rsdtlupah("invoicecode")%>&tglbkm=<%=rsdtlupah("CREATIONDATE")%>&fcba=<%=rsdtlupah("fcba")%>&tipetran=ITR" target="_blank">
											<img src="IMAGES/view.png" target="_blank" alt="Posting Data" border="0"></a>
										</td>
									</tr>
								<% Else %>
									&nbsp;
								<% End If %>
								
								<%
									no = no + 1
									ttlamt     = cdbl(ttlamt) + cdbl(AMOUNT)
									ttlpph     = cdbl(ttlpph) + cdbl(PPH)
									ttlgrand   = cdbl(ttlgrand) + cdbl(GRANDTOTAL)
									ttlpayment = cdbl(ttlpayment) + cdbl(dictPay(rsdtlupah("invoicecode")))
									rsdtlupah.MoveNext
								Wend
								%>

								<tr class="ratakanan">
									<input Type="Hidden" Value="<%=no%>" Name="no">
									<td colspan="4" class="tebal">Total</td>
									<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly style="background-color:#eee;" value="<%=FormatNumber(ttlamt,0)%>"/></td>
									<td class="tebal"></td>
									<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly style="background-color:#eee;" value="<%=FormatNumber(ttlpph,0)%>"/></td>
									<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly style="background-color:#eee;" value="<%=FormatNumber(ttlgrand,0)%>"/></td>
									<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly style="background-color:#eee;" value="<%=FormatNumber(ttlpayment,0)%>"/></td>
								</tr>
								<%end if%>
							</table>
						</div>
			            
						<!-- PAYMENT LIST TBS-->
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
								' dim sqldtlupah,rsdtlupah,no,totalamount
								sqldtlupah = " select t.vouchercode,t.datecreated,t.remarks,t.invoicecode,sum(t.totalamount)totalamount " & _ 
								" from ( " & _
								
								"		SELECT a.vouchercode,to_char(a.datecreated,'dd/mm/yyyy')datecreated, " & _
								"		a.remarks,d.invoicecode,case when d.amount>0 then d.amount else 0 end totalamount " & _
								"		FROM paymentvoucher a   " & _
								"		LEFT JOIN paymentvoucherdetail d on (a.vouchercode=d.vouchercode and a.fcba=d.fcba)  " & _
								"		left join transporterinvoice i on (d.invoicecode=i.invoicecode) " & _
								"		where d.CHARGECODE  ='" & fccode & "'  " & _
								"		AND to_number(to_char(i.CREATIONDATE,'YYYY'))= " & tahun & "   " & _
								"		AND to_number(to_char(i.CREATIONDATE,'MM'))=" & bulan & " and totalamount>0 " & _
								"		AND a.Rowstateid in('Confirmed','AuthorizeOnProgress','Approved')" & _
								" )t " & _
								"GROUP BY t.vouchercode,t.datecreated,t.remarks,t.invoicecode " & _
								"ORDER BY  t.invoicecode ASC "
                        '  	Response.write sqldtlupah
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
									<th>Amount</th>
									<th>VAT</th>
									<th>PPh</th>
									<th>Grand Total</th>
									<th>Payment</th>
									<th>Outstanding</th>
																</tr>
								<% 
								
								'get data detail upah
								if cmdview = "TRUE" then
								sqldtlupah = "SELECT tri.INVOICECODE,tri.postingflg,tri.VAT, " & _
											"  bu.fccompanycode,tri.fcba,tri.invoice_by,to_char(tri.CREATIONDATE,'dd/mm/yyyy') CREATIONDATE,tri.no_faktur, " & _
											"  tri.fakturdate,tri.pph,(tri.AMOUNTPPH+tri.AMOUNTAFTERPPH) as total," & _
											"  tri.AMOUNTPPH*-1 AMOUNTPPH,tri.AMOUNTAFTERPPH, " & _ 
											" NVL(( " & _ 
											"	SELECT SUM(pvd.amount) " & _ 
											"	FROM paymentvoucher pv " & _ 
											"	INNER JOIN paymentvoucherdetail pvd ON pv.vouchercode = pvd.vouchercode " & _ 
											"	WHERE pvd.invoicecode = tri.INVOICECODE AND pv.rowstateid in ('Confirmed','AuthorizeOnProgress','Approved') " & _ 
											"	AND TO_NUMBER(TO_CHAR(pv.datecreated, 'YYYYMM')) <= " & tahun&bulan & "), 0) AS paymentview " & _
											"  FROM transporterinvoice tri " & _
											"	INNER JOIN businessunit bu ON (bu.fccode = tri.fcba) " & _
											"	WHERE tri.fcba = '" & fcba & "' " & _
											"	AND to_number(to_char(tri.CREATIONDATE,'YYYYMM'))<=" & tahun&bulan & " " & _
											"	AND (tri.AMOUNTAFTERPPH - NVL(( " & _
											"					SELECT SUM(pvd.amount) " & _
											"					FROM paymentvoucher pv " & _
											"					INNER JOIN paymentvoucherdetail pvd ON pv.vouchercode = pvd.vouchercode " & _
											"					WHERE pvd.invoicecode = tri.INVOICECODE AND pv.rowstateid in ('Confirmed','AuthorizeOnProgress','Approved') " & _
											"					AND TO_NUMBER(TO_CHAR(pv.datecreated, 'YYYYMM')) <= " & tahun&bulan & " " & _
											"				), 0)) > 1 " & _
											"	ORDER BY CREATIONDATE ASC, INVOICECODE ASC "
                      		   '  Response.Write sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
							    dim outstanding,ttloutstanding,payment,totalpph
								no = 0
								totalamount = 0
								ttlpayment = 0
								ttloutstanding = 0
								While Not rsdtlupah.EOF 
									
									PPH = rsdtlupah("AMOUNTPPH")
									GRANDTOTAL = rsdtlupah("AMOUNTAFTERPPH")
									payment = cdbl(rsdtlupah("paymentview"))
									outstanding = cdbl(GRANDTOTAL)-cdbl(payment)

								%>
								<tr id="tru3<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("INVOICECODE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("CREATIONDATE")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("total"),0)%>
									</td>
									<td class="ratakanan">
										<%=rsdtlupah("VAT")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(PPH,0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTAFTERPPH"),0)%>
									</td>
								
									<td>
										<%=formatnumber(payment,0)%>
									</td>
									<td>
										<%=formatnumber(outstanding,0)%>
									</td>

								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("AMOUNTAFTERPPH"))
									totalpph = cdbl(totalpph) + cdbl(pph)
									ttlpayment = cdbl(ttlpayment) + cdbl(payment)
									ttloutstanding = cdbl(ttloutstanding) + cdbl(outstanding)
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								       
										<td colspan="4" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalpph,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlpayment,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttloutstanding,0)%>"/></td>
								</tr>
								<%end if%>
								
							</table>
						</div>
			
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
		var nm = 'Ap_InvoiceTransporter_ ' + periode + ' - ' + supname + ".xlsx";  
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "Aging InvoiceSpecial"
            }
          });

	}

</script>
