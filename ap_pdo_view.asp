<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim fccode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName ,txtJenis, cmdview,txtname
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
		<title>FBPlants - PDO VIEW</title>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : AP PDO</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form action="ap_ipd_view.Asp" method="post" name="eForm" >
	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >CODE</td>
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
		<%
			if fccode=1 then 
				txtname = "PDO UPAH" 
			ELSE 
				txtname = "PDO NON UPAH" 
			end if
		%>
		<tr>
			<td align="left" >NAME</td>
			<td align="left" colspan="">
				<input type="text" name="suppliername" readonly id="suppliername" size="50" value="<%=txtname%>" >				
			</td>						
		</tr>
		
		<tr>
		<td colspan="4" align="left">

		
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
						<li class="TabbedPanelsTab">INV PDO BULAN INI</li>
					
					
					</ul>

					<div class="TabbedPanelsContentGroup">

<!-- PURCHASING -->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="1050" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No.Invoice</th>
									<th>Date</th>
									<th>Sumber</th>
									<th>Estate</th>
									<th>Amount</th>

								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,postingflg
								if cmdview = "TRUE" then
								sqldtlupah = "  SELECT inv.noinvoice, to_char(inv.CREATEDATE	,'dd/mm/yyyy')invoicedate, " & _
											 " CASE WHEN inv.sumber = 1 THEN 'PDO UPAH' ELSE 'PDO NON UPAH' END as Sumber, " & _
											 " inv.fcba, nvl(inv.amount,0)amount " & _
											 " FROM OPR_INV_PDO inv " & _
											 " where inv.fcba='" & fcba & "'" & _ 
											 " AND inv.tahun="& tahun & " " & _ 
											 " AND inv.bulan="& bulan & " " & _
											 " AND inv.sumber='" & fccode & "' " & _
											 " ORDER BY CREATEDATE ASC, noinvoice DESC "
                           	
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
								no = 0
								totalamount = 0
								While Not rsdtlupah.EOF 
									
								%>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("noinvoice")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("invoicedate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("Sumber")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("fcba")%>
									</td>
									<td class="ratakanan">
										<%=formatnumber(rsdtlupah("amount"),0)%>
									</td>
									
						
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("amount"))
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
								        <input Type="Hidden" Value="<%=no%>" Name="no">
										<td colspan="4" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
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
