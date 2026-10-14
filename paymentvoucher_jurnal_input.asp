<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,vouchercode,tglvoucher,datecreated,bankcode,fcba,docno,postingflg
dim sqljurnal,rsjurnal,sqlfac,NoOfRows,tipetran

docno	= trim(request("fccode"))
bankcode= trim(request("xcode"))
fcba = trim(request("fcba"))
datecreated = trim(request("datecreated"))
tipetran = trim(request("tipetran"))



q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

' 'cek transaksi pada tabel jurnal, jika tidak ada buat jurnal dulu
' sqljurnal = "SELECT Count (1) TOTALJR from BKMJOURNALHDR H INNER JOIN BKMJOURNALDTL D ON (H.DOCNO=D.DOCNO AND H.FCBA=D.FCBA)" & _ 
' " where D.DOCNO='" & docno & "' AND H.FCBA='" & fcba& "'"
	
' set rsjurnal = connStr.Execute(sqljurnal)
' if not rsjurnal.eof then
  ' NoOfRows = rsjurnal("TOTALJR")
  ' 'response.write NoOfRows
' end if

' if NoOfRows = "0" then
' if tipetran = "k" then
' sqlfac = "BEGIN FAC_JOURNAL_FAC_BNK_API.FAC_BNK_PYV_DOCNO('" & docno & "',TO_NUMBER(TO_CHAR(TO_DATE('" & datecreated & "','DD/MM/YYYY'),'YYYY')), " & _ 
																		' "TO_NUMBER(TO_CHAR(TO_DATE('" & datecreated & "','DD/MM/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
		' "END;"
		' 'response.write sqlfac													 															 
' connStr.execute(sqlfac)
' end if 
' if tipetran = "m" then
' sqlfac = "BEGIN FAC_JOURNAL_FAC_BNK_API.FAC_BNK_RCV_DOCNO('" & docno & "',TO_NUMBER(TO_CHAR(TO_DATE('" & datecreated & "','DD/MM/YYYY'),'YYYY')), " & _ 
																		' "TO_NUMBER(TO_CHAR(TO_DATE('" & datecreated & "','DD/MM/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
		' "END;"
		' 'response.write sqlfac													 															 
' connStr.execute(sqlfac)
' end if
' end if

' get data bank 
dim sqlheader,rsHeader,bankname
sqlheader = " select B.FCCODE || ' - ' || B.FCNAME BANKCODE " & _
			" from BANK B " & _
			" where B.FCCODE='" & bankcode & "' AND B.FCBA='" & fcba & "'"
			'response.write sqlheader
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	bankname = rsHeader("BANKCODE")
end if

sqlheader = " select postingflg " & _
			" from  BKMJOURNALHDR" & _
			" where DOCNO='" & docno & "' AND FCBA='" & fcba & "'"
			'response.write sqlheader
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	postingflg = rsHeader("postingflg")
end if


%>

<html>
	<head>
		<title>FBPlants - VOUCHER JURNAL </title>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : JURNAL VOUCHER KAS DAN BANK </strong><br>		
			
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="paymentvoucher_jurnal_input_proses.asp">
	<table class="data posisiheader" border="0" width="80%">			
		<tr>
			<td align="left" >Jurnal NO</td>
			<td align="left">
				<input type="text" name="jurnalno" readonly id="jurnalno" size="20" value="<%=docno%>" >
			</td>


			<td align="left" >Jurnal Date</td>
			<td align="left">
				<input type="text" name="jurnaldate" readonly id="jurnaldate" size="20" value="<%=datecreated%>" >
			</td>
			<td align="left" >Remark</td>
			<td align="left" colspan="">
				<textarea  name="remarks" id="remarks" cols="30" rows=30></textarea>
				
			</td>
		</tr>
		
		<tr>
			<td align="left" >Kas/Bank</td>
			<td align="left" colspan="">
				<input type="text" name="bank" id="bank" value="<%=bankname%>" ></input>
				
			</td>
			<td align="left" >Posting</td>
			<td align="left" colspan="">
			
				<%
					if postingflg= "0" then 
						response.write "TIDAK"
					elseif postingflg= "" then
						response.write "TIDAK"
					else 
						response.write "YA"
					end if
				%>
				
				
			</td>
			
		</tr>
		
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data ' width='80%' id='table1'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">JURNAL VOUCHER</li>
					</ul>

					<div class="TabbedPanelsContentGroup">
						<!-- PYMENT VOUCHER-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="900" class='data' id='table1'>
									<tr id="trheader2">
										<th>JOBNAME</th>
										<th>REMARKS</th>
										<th>TARGET CODE</th>
										<th>DEBIT</th>
										<th>KREDIT</th>
										
									</tr>
								<%
										'get data detail payment voucher
									'dim sqljrnlbhnx,rsjrnlbhn,nobhn,ttlamountbhn,ttlamountbhnkredit
									dim sqlvc,rsvc,novc,ttlamountvc,ttlamountvckredit
									sqlvc = "select d.BKMNO,j.fccode, j.fccode || '-' ||  j.fcname as jobcode,d.locationcode,d.description,sum(d.amountdebit) amountdebit,sum(d.amountkredit)amountkredit " & _ 
									" from BKMJOURNALHDR h " & _
									" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba) " & _
									" inner join job j on (j.fccode=d.JOBCODE and j.fcba=d.fcba) " & _
									" where h.DOCNO='" & docno & "' and h.FCBA='" & fcba & "' " & _
									" group by d.BKMNO,j.fccode,d.locationcode,j.fccode || '-' ||  j.fcname,lineno,d.description order by lineno" 
									'Response.Write sqlvc
									set rsvc = connStr.Execute(sqlvc)

								novc = 0
								ttlamountvc = 0
								ttlamountvckredit = 0 
								While Not rsvc.EOF %>
								<tr id="trb<%=novc%>">
									
									
									<td class="ratakiri">
										<input class="ratakiri lebaractivity" readonly type="text" name="activitybhn<%=novc%>" id="activitybhn<%=novc%>" value="<%=rsvc("jobcode")%>"/>
										<input class="ratakiri lebaractivity" readonly type="hidden" name="activitybhn<%=novc%>" id="activitybhn<%=novc%>" value="<%=rsvc("fccode")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="bkmnobhn<%=novc%>" id="bkmnobhn<%=novc%>" value="<%=rsvc("description")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebaremployee" readonly type="text" name="tcbhn<%=novc%>" id="tcbhn<%=novc%>" value="<%=rsvc("locationcode")%>"/>
									</td>
									
									<td>
										<input class="ratakanan lebarupah"  type="text" name="debitbhn<%=novc%>" id="debitbhn<%=novc%>" value="<%=formatnumber(rsvc("amountdebit"),0)%>"/>
									</td>
									<td>
										<input class="ratakanan lebarupah" type="text" name="kreditbhn<%=novc%>" id="kreditbhn<%=novc%>" value="<%=formatnumber(rsvc("amountkredit"),0)%>"/>
									</td>
								
									<td></td>
									
								</tr>

								
								<%

									novc = novc +1
									
									ttlamountvc = cdbl(ttlamountvc) + cdbl(rsvc("amountdebit"))
									ttlamountvckredit = cdbl(ttlamountvckredit) + cdbl(rsvc("amountkredit"))
								
									rsvc.MoveNext
									
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="3" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountbhn" id="amountbhn<%=novc%>" value="<%=FormatNumber(ttlamountvc,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amountbhn" id="amountbhn<%=novc%>" value="<%=FormatNumber(ttlamountvckredit,0)%>"/></td>
									</tr>
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>
						
					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">
			<td  colspan="5">
			                
			                  <input type=Submit name="btnsave" id="btnsave" value="POSTING">&nbsp;&nbsp;
			                  <input type=Submit name="btnunposting" id="btnunposting" value="UNPOSTING">&nbsp;&nbsp;
							  <input type=button onclick="window.close();" name="btnclose" value="CLOSE">
							 
							 
			</td>
			<input type="hidden" name="POSTINGFLG" id="POSTINGFLG" value="<%=postingflg%>"/>
			<input type="hidden" name="fcba" id="fcba" value="<%=fcba%>"/>
			<input type="hidden" name="view" id="view" value="0"/>
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
