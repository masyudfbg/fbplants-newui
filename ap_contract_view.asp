<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim fccode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName ,txtJenis, cmdview,txtoustanding
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
dim sqlheader,rsHeader,contractorname ,n,sqlexec
sqlheader = "select fccode, fcname from contractor where fccode='" & fccode & "' AND fcba='" & fcba & "' "
set rsHeader = connStr.Execute(sqlheader)
'response.write sqlheader
if not rsHeader.eof then 
	
	contractorname = rsHeader("fcname")
	
end if

if Request("txtJenis")= "POSTING" then
	For n=0 to (cdbl(request("no"))-1)
	'response.write "Y-"& n & "-" & request("cb_"&TRIM(n)) & "<br>"

			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				sqlexec     = "UPDATE CONTRACTORINVOICEHDR SET postingflg='1',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE  invoicecode='" & request("invoiceno_"&TRIM(n)) & "' AND fcba='" & request("fcba_"& TRIM(n))  & "'"
				'response.write "1-"& sqlexec & "<br>"
				connStr.execute(sqlexec)
				
				sqlexec 	= "BEGIN fac_journal_inv_ico_api.ProsesGenerate('" & request("invoiceno_"&TRIM(n)) & "', '" & request("fcba_"&TRIM(n)) & "','SYSTEM'); END;"
                'response.write "2-"& sqlexec & "<br>"
				connStr.execute(sqlexec)
				
				sqlexec = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & request("fcba_"&TRIM(n))  & "' AND docno='" & request("invoiceno_"&TRIM(n)) & "'"
				'response.write "3-"& sqlexec & "<br>"
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
		<title>FBPlants - ICO VIEW</title>
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
		left: 0px;
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : AP CONTRACT</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form action="ap_contract_view.Asp" method="post" name="eForm" >
	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >CONTRACTOR CODE</td>
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
			<td align="left" >CONTRACTOR NAME</td>
			<td align="left" colspan="">
				<input type="text" name="contractorname" readonly id="contractorname" size="50" value="<%=contractorname%>" >				
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
		
	<table align=center border="0" class='data posisiheader'  id='table1'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">INVOICE BULAN INI </li>
						<li class="TabbedPanelsTab">PAYMENT BULAN INI</li>
						<li class="TabbedPanelsTab">OUTSTANDING S.D BULAN INI</li>
						<li class="TabbedPanelsTab">OUTSTANDING DI SEMUA PT. S.D BULAN INI</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

<!-- CONTRACTOR -->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No.Invoice</th>
									<th>Invoice Date</th>
									<th>Agreement</th>
									<th>Retention</th>
									<th>Amount</th>
									<th>PPN</th>
									<th>PPH</th>
									<th>Grand Total</th>
									<th>Payment</th>
									<th><input type='checkbox' name='select-all' id='select-all'/>Posting</th>
									<th>Action</th>	
									
								</tr>

								<% 
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = "  SELECT inv.invoicecode, to_char(inv.TRANSDATE,'dd/mm/yyyy')TRANSDATE,inv.CONTRACTORCODE, inv.CONTRACTORCODE || ' - ' || s.fcname as contractor,inv.postingflg, " & _
											 " inv.AGREEMENTCODE,nvl(inv.RETENTION,0)RETENTION, nvl(inv.TOTALAMOUNT_AFTERRETENTION,0)TOTALAMOUNT_AFTERRETENTION, " & _ 
											 " nvl(inv.AMOUNTPPN,0)AMOUNTPPN, nvl(inv.AMOUNTPPH,0)AMOUNTPPH, " & _ 
											 " nvl(inv.GRANDTOTAL_AFTERTAX,0)GRANDTOTAL_AFTERTAX,inv.fcba,nvl(inv.payment,0)payment " & _
											 " FROM CONTRACTORINVOICEHDR inv INNER JOIN CONTRACTOR s ON (inv.CONTRACTORCODE=s.fccode and inv.fcba=s.fcba) " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND to_number(to_char(inv.TRANSDATE,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(inv.TRANSDATE,'MM'))="& bulan & " " & _
											 " AND inv.CONTRACTORCODE='" & fccode & "' " & _
											 " ORDER BY TRANSDATE ASC, invoicecode ASC "
								'Response.Write sqldtlupah
                           	
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
								<tr id="tru<%=no%>" bgcolor="#EEEEEE">
									
									<td class="ratakiri">
										<%=rsdtlupah("invoicecode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("TRANSDATE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("AGREEMENTCODE")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("RETENTION")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("TOTALAMOUNT_AFTERRETENTION"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTPPN"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTPPH"),0)%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("GRANDTOTAL_AFTERTAX"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("payment"),0)%>
									</td>
			                        
									<td>
										<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
										<%if rsdtlupah("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%>
										<input type="hidden" name="fcba_<%=NO%>" value="<%=rsdtlupah("fcba")%>">
										<input type="hidden" name="invoiceno_<%=NO%>" value="<%=rsdtlupah("invoicecode")%>">
									</td>
																		
									<td colspan=2 class="ratatengah">
									<a href="contractorinvoice.Asp?invoicecode=<%=rsdtlupah("invoicecode")%>&fcba=<%=rsdtlupah("fcba")%>&txtJenis=Edit">
									<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
									</a>&nbsp;&nbsp;&nbsp;<a href="bkm_jurnal_view.Asp?bkmno=<%=rsdtlupah("invoicecode")%>&tglbkm=<%=rsdtlupah("TRANSDATE")%>&fcba=<%=rsdtlupah("fcba")%>&tipetran=ICO" target="_blank">
									<img src="IMAGES/view.png" target="_blank"  alt="Posting Data : <%=rsdtlupah("invoicecode")%>" border="0">
									</a>
									</td>
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("GRANDTOTAL_AFTERTAX"))
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								        <input Type="Hidden" Value="<%=no%>" Name="no">
										<td colspan="7" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
								</tr>
							</table>
							<%end if%>
						</div>
						
						<!-- PAYMENT LIST-->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No.Voucher</th>
									<th>Payment Date</th>
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
								"		left join CONTRACTORINVOICEHDR i on (d.invoicecode=i.invoicecode and i.invoice_by=a.fcba) " & _
								"		where (d.locationcode = '" & fccode & "' or d.CHARGECODE  ='" & fccode & "' ) " & _
								"		AND to_number(to_char(a.DATECREATED,'YYYY'))= " & tahun & "   " & _
								"		AND to_number(to_char(a.DATECREATED,'MM'))=" & bulan & " and a.totalamount>0 AND a.fcba='" & fcba & "' AND a.Rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
								" )t " & _
								"GROUP BY t.vouchercode,t.datecreated,t.remarks,t.invoicecode " & _
								"ORDER BY  t.invoicecode ASC "
                          		'Response.write sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
								no = 0
								totalamount = 0
								While Not rsdtlupah.EOF 
									
								%>
								<tr id="tru<%=no%>" bgcolor="#EEEEEE">
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

						<!-- OUTSTANDING AGING -->
							<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id="product-table">
								<tr id="trheader2" >									
									<th>No.Invoice</th>
									<th>Invoice Date</th>
									<th>Agreement</th>
									<th>Retention</th>
									<th>Amount</th>
									<th>PPN</th>
									<th>PPH</th>
									<th>Grand Total</th>
									<th>Payment</th>
									<th>Outstanding</th>
									<th>Remarks</th>
									
								</tr>

								<% 
								
								'get data detail upah
							'	dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = "  SELECT inv.invoicecode, to_char(inv.TRANSDATE,'dd/mm/yyyy')TRANSDATE,inv.CONTRACTORCODE, inv.CONTRACTORCODE || ' - ' || s.fcname as contractor,inv.postingflg, " & _
											 " inv.AGREEMENTCODE,nvl(inv.RETENTION,0)RETENTION, nvl(inv.TOTALAMOUNT_AFTERRETENTION,0)TOTALAMOUNT_AFTERRETENTION, " & _ 
											 " nvl(inv.AMOUNTPPN,0)AMOUNTPPN, nvl(inv.AMOUNTPPH,0)AMOUNTPPH, " & _ 
											 " nvl(inv.GRANDTOTAL_AFTERTAX,0)GRANDTOTAL_AFTERTAX,inv.fcba,inv.DESCRIPTION, " & _
											 " NVL((SELECT SUM(dtl.amount) " & _
											 "	FROM paymentvoucher hdr " & _
											 "  INNER JOIN paymentvoucherdetail dtl ON hdr.vouchercode = dtl.vouchercode " & _
											 "	WHERE dtl.invoicecode = inv.invoicecode  AND inv.CONTRACTORCODE='" & fccode & "' AND hdr.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
											 "	AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'YYYY')) <= " & tahun & " AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'MM')) <= " & bulan & " AND hdr.Rowstateid in('Confirmed','AuthorizeOnProgress','Approved')" & _
											 "	), 0) AS payment"  & _
											 " FROM CONTRACTORINVOICEHDR inv INNER JOIN CONTRACTOR s ON (inv.CONTRACTORCODE=s.fccode and inv.fcba=s.fcba) " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND to_number(to_char(inv.TRANSDATE,'YYYY'))<="& tahun  & " AND TO_NUMBER(TO_CHAR(inv.TRANSDATE, 'MM')) <= " & bulan & " " & _ 
											 " AND inv.CONTRACTORCODE='" & fccode & "' " & _
											 " AND (inv.GRANDTOTAL_AFTERTAX -   NVL((SELECT SUM(dtl.amount) " & _
											 "	FROM paymentvoucher hdr " & _
											 "  INNER JOIN paymentvoucherdetail dtl ON hdr.vouchercode = dtl.vouchercode " & _
											 "	WHERE dtl.invoicecode = inv.invoicecode  AND inv.CONTRACTORCODE='" & fccode & "' AND hdr.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
											 "	AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'YYYY')) <= " & tahun & " AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'MM')) <= " & bulan & " AND hdr.Rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
											 "	), 0)) > 1000"  & _
											 " ORDER BY inv.TRANSDATE ASC "
								
							'	Response.Write sqldtlupah
                           	
								set rsdtlupah = connStr.Execute(sqldtlupah)
								dim grnd,pay,totalpayment,totaloustanding
																
								no = 0
								totalamount = 0
								totalpayment = 0
								totaloustanding = 0
								While Not rsdtlupah.EOF 
									grnd = rsdtlupah("GRANDTOTAL_AFTERTAX")
									pay = rsdtlupah("payment")
									txtoustanding = cdbl(grnd)-cdbl(pay)
									
								%>
								<tr id="tru3<%=no%>" bgcolor="#EEEEEE">
									
									<td class="ratakiri">
										<%=rsdtlupah("invoicecode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("TRANSDATE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("AGREEMENTCODE")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("RETENTION")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("TOTALAMOUNT_AFTERRETENTION"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTPPN"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTPPH"),0)%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("GRANDTOTAL_AFTERTAX"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("payment"),0)%>
									</td>
									<td>
										<%=formatnumber(txtoustanding,0)%>
									</td>
									<td>
										<%=rsdtlupah("DESCRIPTION")%>
									</td>
			                        
															
									
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("GRANDTOTAL_AFTERTAX"))
									totalpayment = cdbl(totalpayment) + cdbl(rsdtlupah("payment"))
									totaloustanding = cdbl(totaloustanding) + cdbl(txtoustanding)
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								       
										<td colspan="7" class="tebal">Total</td>
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


						<!-- OUTSTANDING AGING DI SEMUA PT -->
							<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id="product-table">
								<tr id="trheader4" >									
									<th>Company</th>
									<th>Unit</th>
									<th>No.Invoice</th>
									<th>Invoice Date</th>
									<th>Agreement</th>
									<th>Retention</th>
									<th>Amount</th>
									<th>PPN</th>
									<th>PPH</th>
									<th>Grand Total</th>
									<th>Payment</th>
									<th>Outstanding</th>
									<th>Remarks</th>
									
								</tr>

								<% 
								dim codeconmp
								'get data detail upah
							'	dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = "SELECT FCCODEMAPPING FROM CONTRACTOR WHERE FCCODE='" & fccode & "' AND fcba='" & fcba & "'"
								set rsHeader = connStr.Execute(sqldtlupah)
								'response.write sqlheader
								if not rsHeader.eof then 
									codeconmp = rsHeader("FCCODEMAPPING")
								end if

								sqldtlupah = "  SELECT b.FCCOMPANYCODE,inv.FCBA, inv.invoicecode, to_char(inv.TRANSDATE,'dd/mm/yyyy')TRANSDATE,inv.CONTRACTORCODE, inv.CONTRACTORCODE || ' - ' || s.fcname as contractor,inv.postingflg, " & _
											 " inv.AGREEMENTCODE,nvl(inv.RETENTION,0)RETENTION, nvl(inv.TOTALAMOUNT_AFTERRETENTION,0)TOTALAMOUNT_AFTERRETENTION, " & _ 
											 " nvl(inv.AMOUNTPPN,0)AMOUNTPPN, nvl(inv.AMOUNTPPH,0)AMOUNTPPH, " & _ 
											 " nvl(inv.GRANDTOTAL_AFTERTAX,0)GRANDTOTAL_AFTERTAX,inv.fcba,inv.DESCRIPTION, " & _
											 " NVL((SELECT SUM(dtl.amount) " & _
											 "	FROM paymentvoucher hdr " & _
											 "  INNER JOIN paymentvoucherdetail dtl ON hdr.vouchercode = dtl.vouchercode " & _
											 "	WHERE dtl.invoicecode = inv.invoicecode  AND s.FCCODEMAPPING='" & codeconmp & "' AND hdr.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
											 "	AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'YYYY')) <= " & tahun & " AND TO_NUMBER(TO_CHAR(hdr.datecreated, 'MM')) <= " & bulan & " AND hdr.Rowstateid in('Confirmed','AuthorizeOnProgress','Approved')" & _
											 "	), 0) AS payment"  & _
											 " FROM CONTRACTORINVOICEHDR inv INNER JOIN CONTRACTOR s ON (inv.CONTRACTORCODE=s.fccode and inv.fcba=s.fcba) " & _
											 " INNER JOIN BUSINESSUNIT b ON (b.fccode=inv.fcba) " & _
											 " where to_number(to_char(inv.TRANSDATE,'YYYY'))<="& tahun  & " AND TO_NUMBER(TO_CHAR(inv.TRANSDATE, 'MM')) <= " & bulan & " " & _ 
											 " AND s.FCCODEMAPPING='" & codeconmp & "' " & _
											 " AND inv.GRANDTOTAL_AFTERTAX - inv.payment   > 1"  & _
											 " ORDER BY b.FCCOMPANYCODE,inv.TRANSDATE ASC "
								
								'Response.Write sqldtlupah
                           	
								set rsdtlupah = connStr.Execute(sqldtlupah)
							'	dim grnd,pay,totalpayment,totaloustanding
																
								no = 0
								totalamount = 0
								totalpayment = 0
								totaloustanding = 0
								While Not rsdtlupah.EOF 
									grnd = rsdtlupah("GRANDTOTAL_AFTERTAX")
									pay = rsdtlupah("payment")
									txtoustanding = cdbl(grnd)-cdbl(pay)
									
								%>
								<tr id="tru4<%=no%>" bgcolor="#EEEEEE">
									
									<td class="ratakiri">
										<%=rsdtlupah("FCCOMPANYCODE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("FCBA")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("invoicecode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("TRANSDATE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("AGREEMENTCODE")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("RETENTION")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("TOTALAMOUNT_AFTERRETENTION"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTPPN"),0)%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("AMOUNTPPH"),0)%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("GRANDTOTAL_AFTERTAX"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("payment"),0)%>
									</td>
									<td>
										<%=formatnumber(txtoustanding,0)%>
									</td>
									<td>
										<%=rsdtlupah("DESCRIPTION")%>
									</td>
			                        
															
									
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("GRANDTOTAL_AFTERTAX"))
									totalpayment = cdbl(totalpayment) + cdbl(rsdtlupah("payment"))
									totaloustanding = cdbl(totaloustanding) + cdbl(txtoustanding)
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								       
										<td colspan="9" class="tebal">Total</td>
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
		var contractorname 	= $("#contractorname").val();
		var nm = 'AP_Contractor_' + periode + ' - ' + contractorname + ".xlsx";  
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "Aging Contractor"
            }
          });

	}


</script>
