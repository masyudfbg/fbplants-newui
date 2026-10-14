<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,rkbno,sc
dim rsCompany,company
rkbno= trim(request("rkbno"))
sc= trim(request("sectioncode"))
businessunit= trim(session("fcba"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

dim sqlview,rsview


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID,Clasification,Clasificationname,ROWSTATE
DIM sqlhdr,rshdr 
DIM RKBNOHDR,PERIOD_YEAR,PERIOD_MONTH,SECTIONCODE,CREATEDATE,NOKONTRAK,CHARGEUNIT,FCCODE,FCNAME,COMPANYCODE,FCBA


sqlheader = "select rkbno,period_month,period_year,sectioncode,document_classification,rowstate,totupah,totbahan,totvra,totwip,classification_api.get_fcname(document_classification) docname " & _
            " from rkbhdr where rkbno='" & rkbno & "' AND fcba='" & businessunit & "'"
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	bkmnohdr = rsHeader("rkbno")
	period_month = rsHeader("period_month")
	period_year = rsHeader("period_year")
	SECTIONNAME = rsHeader("sectioncode")
	Clasification = rsHeader("document_classification")
	Clasificationname = rsHeader("docname")
	
	ROWSTATE = rsHeader("rowstate")
end if

sqlhdr = "SELECT H.RKBNO,H.PERIOD_YEAR,H.PERIOD_MONTH,H.SECTIONCODE,H.CREATEDATE,H.NOKONTRAK,H.FCBA,H.CHARGEUNIT,C.FCCODE,C.FCNAME,B.FCCOMPANYCODE " & _
" FROM RKBHDR H " & _
" LEFT JOIN CONTRACTAGREEMENT CR ON H.NOKONTRAK=CR.AGREEMENTCODE AND H.FCBA=CR.FCBA " & _
" LEFT JOIN CONTRACTOR  C ON C.FCCODE=CR.CONTRACTORCODE AND C.FCBA=CR.FCBA " & _
" LEFT JOIN BUSINESSUNIT B ON B.FCCODE = H.FCBA " & _
" WHERE H.RKBNO='" & rkbno & "' and H.SECTIONCODE = '"&sc&"'"
set rshdr = connStr.Execute(sqlhdr)
if not rshdr.eof then 
	RKBNOHDR 		= rshdr("RKBNO")
	PERIOD_YEAR = rshdr("PERIOD_YEAR")
	PERIOD_MONTH = rshdr("PERIOD_MONTH")
	SECTIONCODE = rshdr("SECTIONCODE")
	CREATEDATE 	= rshdr("CREATEDATE")
	NOKONTRAK 	= rshdr("NOKONTRAK")
	CHARGEUNIT 	= rshdr("CHARGEUNIT")
	FCCODE 		= rshdr("FCCODE")
	FCNAME 		= rshdr("FCNAME")
	COMPANYCODE = rshdr("FCCOMPANYCODE")
	FCBA 		= rshdr("FCBA")
	
end if
%>

<html>
	<head>
		<title>FBPlants - RKB WIP</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
		<link href="/js/spry/widgets/selectvalidation/SpryValidationSelect.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/selectvalidation/SpryValidationSelect.js" type="text/javascript"></script>
	</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: 15px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarhk{
		width: 45px;
	}
	.lebarabsen{
		width: 48px;
	}
	.lebarlokasi{
		width: 60px;
	}
	.lebartc{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		/*width: 135px;*/
	}
	.lebarallow{
		width: 130px;
	}
	.lebarupah{
		width: 80px;
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
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 200px;
	}
	.lebarvehicle{
		width: 200px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	
</style>



<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : RKB WIP</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="rkb_wip_input.asp">
	
	<table class="data posisikiri" border="1" width="80%" align="center">		
		<tr>
			<td  align="left">
				<font face="Verdana" size="1" align="LEFT">Company</font>
			</td>
			<td  align="left" colspan="5">
				<font face="Verdana" size="1" align="LEFT"><%=COMPANYCODE%></font>
			</td>
		</tr>	
		<tr>
			<td  align="left">
				<font face="Verdana" size="1" align="LEFT">RKB NO.</font>
			</td>
			<td  align="left">
				<%=rkbno%>
			</td align="left">
				<td align="left" >FCBA Kontrak</td>
				<td align="left" colspan="3">
					<%=FCBA%>
			</td>
		</tr>
		<tr>
			<td align="left" >No Kontrak</td>
			<td align="left">
				<%=NOKONTRAK%>
			</td>
			<td align="left" >Contractor Name</td>
			<td align="left" colspan="3">
				<%=FCCODE%> - <%=FCNAME%>
			</td>
					
		</tr>
		<tr>
			<td align="left" >Charge Unit</td>
			<td align="left">
				<%=CHARGEUNIT%>
			</td>
			
			<td align="left" >Period</td>
			<td align="left">
			   <%=PERIOD_YEAR%>	- <%=PERIOD_MONTH%>				
			</td>
			<td align="left" >Tanggal Input</td>
			<td align="left">
				<%=CREATEDATE%>
			</td>
		</tr>

	</table>
	
	<br>
		<input type="hidden" name="jmlrkb" id="jmlrkb" value="0" />
	<table align=center border="0" class='data posisikiri' width='80%' id='tableall'>
		<tr>
			<td>
			   		<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB WIP</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm wip-->
						<div class="TabbedPanelsContent">
							<table align=center border="1" width="900" class='data' id='table2'>
								<tr id="trheader2">
									<th>Aktivitas</th>
									<th>Qty Kontrak</th>
									<th>Harga</th>
									<th>Blok</th>
									<th>RKB Qty</th>
									<th>Total</th>
								</tr>
								<%
								dim sqldtl,rsdtl 
								sqldtl = "Select H.RKBNO, H.ACTCODE,J.FCNAME, H.LOCATIONCODE, H.QTY, H.HARGA, (H.QTY*H.HARGA) SUBTOTAL, " & _
								" H.LINENOKEY, H.FCBA " & _
								" From   RKBWIPDTL H " & _
								" INNER JOIN RKBHDR I ON H.RKBNO=I.RKBNO " & _
								" INNER JOIN JOBMASTER J ON H.ACTCODE=J.FCCODE " & _
								" WHERE H.RKBNO='" & rkbno & "' and I.SECTIONCODE = '"&sc&"'"
								' response.write sqldtl
								Set rsdtl = connStr.Execute(sqldtl)		
								DIM TOTALRKBQTY,TOTALALL	
								TOTALRKBQTY = 0
								TOTALALL	= 0				
								While Not rsdtl.EOF
								
								%>
								<tr>
									
									<td align="left"><%=rsdtl("ACTCODE")%> - <%=rsdtl("FCNAME")%></td>
									<td align="left"></td>
									<td><%=FormatNumber(rsdtl("HARGA"))%></td>
									<td align="center"><%=rsdtl("LOCATIONCODE")%></td>
									<td align="center"><%=rsdtl("QTY")%></td>
									<td><%=FormatNumber(rsdtl("SUBTOTAL"))%></td>
								</tr>
							<%
								TOTALRKBQTY = cdbl(TOTALRKBQTY) + cdbl(rsdtl("QTY"))
								TOTALALL	= cdbl(TOTALALL) + cdbl(rsdtl("SUBTOTAL"))
								rsdtl.MoveNext
								
								Wend
							%>
								<tr class="ratakanan">
										<td colspan="4" class="tebal">Total</td>
										<td class="tebal">
											<%=FormatNumber(TOTALRKBQTY)%>
										</td>
										<td class="tebal">
											<%=FormatNumber(TOTALALL)%>
										</td>
								</tr>
								<tr>
									<td  colspan="6" class="ratatengah"><input type=Submit name="simpan" value="KEMBALI"></td>
								</tr>
							</table>
						</div>

					</div>
				</div>
			</td>
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
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
	
	showKontrak();
	
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


	function getSubtotal(i)
	{
		var harga 		= parseInt($('#rate' + i).val());
		var rkbqty 		= parseInt($('#rkbqty' + i).val());
		//alert(harga);
		var total 		= harga*rkbqty;
		var totalfrm    = numformat(total);
		$('#subtotalfrm' + i).val((totalfrm));
		$('#subtotal' + i).val((total));
	}
	
	function getPuttotal() {
		var ttlrkb	= 0;
		var jmlrkb	= parseInt($('#jmlrkb').val());
		for(var i=0; i<jmlrkb; i++) {
			ttlrkb	    += parseFloat($('#rkbqty'+i).val());
		  
			
		}	
		ttlrkbfrm     = numformat(ttlrkb);
		$('#amountrkbqtyfrm').val((ttlrkbfrm));
		$('#amountrkbqty').val((ttlrkb));
		
	}

	function getPuttotalrp() {
		var subtotal	= 0;
		var jmlrkb	= parseInt($('#jmlrkb').val());
		for(var i=0; i<jmlrkb; i++) {
			subtotal	    += parseFloat($('#subtotal'+i).val());
		}	
		var subtotalfrm     = numformat(subtotal);
		$('#amountrkbrpfrm').val((subtotalfrm));
		$('#amountrkbrp').val((subtotal));
		
	}



</script>
