<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno,tglbkm,tglsplit,bulansplit,thnsplit,sqlexec,tipetran,sqlexec1
bkmno= request("bkmno")
tglbkm= request("tglbkm")
businessunit= request("fcba")
tipetran = request("tipetran")
tglsplit  = left(tglbkm,2)
bulansplit  = mid(tglbkm,4,2)
thnsplit    = right(tglbkm,4)

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	'response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

if tipetran="IPO" then
sqlexec 	= "BEGIN FAC_JOURNAL_INVPURCH_API.prosesbyNotran('" & bkmno & "', '" & businessunit & "','SYSTEM'); END;"
End if

if tipetran="ISP" then
sqlexec 	= "BEGIN fac_journal_inv_isp_api.prosesbyNotran('" & bkmno & "', '" & businessunit & "','SYSTEM'); END;"
End if

if tipetran="IFB" then
sqlexec 	= "BEGIN fac_journal_inv_ifb_api.prosesbyNotran('" & bkmno & "', '" & businessunit & "','SYSTEM'); END;"
End if

if tipetran="ITR" then
sqlexec 	= "BEGIN fac_journal_inv_itr_api.prosesbyNotran('" & bkmno & "', '" & businessunit & "','SYSTEM'); END;"
End if

if tipetran="ICO" then
sqlexec	= "BEGIN fac_journal_inv_ico_api.ProsesGenerate('" & bkmno & "', '" & businessunit & "','SYSTEM'); END;"
'response.write sqlexec
End if


if tipetran="IST" then
sqlexec 	= "BEGIN fac_journal_inv_ist_api.prosesbyNotran('" & bkmno & "', '" & businessunit & "','SYSTEM'); END;"
End if

'response.write sqlexec
connStr.Execute(sqlexec)

' get data header  to_date('" & IssueDate & "','mm/dd/yyyy')
dim sqlheader,rsHeader,jurnalno,jurnaldate,remarks,sourcecode
sqlheader = "select docno, to_char(docdate,'dd/mm/yyyy')docdate,remarks,sourcecode from bkmjournalhdr where fcba='" & businessunit & "' AND docno='" & bkmno & "' AND DOCDATE=to_date('" & tglbkm & "','dd/mm/yyyy') "

'response.write sqlheader

set rsHeader = connStr.Execute(sqlheader)
if Not rsHeader.EOF then 
	jurnalno = rsHeader("docno")
	jurnaldate = rsHeader("docdate")
	remarks = rsHeader("remarks")	
	sourcecode = rsHeader("sourcecode")
end if

'response.write sqlexec

Dim rsAvailable, strAvailable,Access_Post
 	strAvailable =	" select authorizestructure_api.get_access_post('ACC-JOURNAL','" & businessunit & "','" & Session("CodeName") & "') result" & _
 					" from dual"
 	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
 	rsAvailable.Open strAvailable, connstr, 3
	
 	If Not rsAvailable.EOF Then
 		Access_Post = cdbl(rsAvailable("result"))
 	Else
 		Access_Post = 0
 	End If
 	rsAvailable.Close
 	set rsAvailable = Nothing


'count detailupah
Dim SQL_Rows, NoOfRows,RSCount

SQL_Rows = "SELECT Count (docno) from bkmjournaldtl where docno='" & jurnalno & "' and fcba='" & businessunit & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close


'get value button
Dim sqlbuton,rsbuton,POSTINGFLG
sqlbuton = "SELECT POSTINGFLG FROM GLEDGERDTL WHERE JOURNALNO='" & jurnalno & "' AND FCBA='" & businessunit & "' "
set rsbuton = connStr.Execute(sqlbuton)
if not rsbuton.eof then 
	POSTINGFLG = rsbuton("POSTINGFLG")
else 
	POSTINGFLG = 0
end if
rsbuton.close

%>

<html>
	<head>
		<title>FBPlants - JURNAL VIEW</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014b"></script>
<% end if %>
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

<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-bkm"<% end if %>>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM JURNAL</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_jurnal_input_proses.asp">
<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<table class="data posisiheader" border="0" width="50%">			
		<tr>
			<td align="left" >Jurnal NO</td>
			<td align="left">
				<input type="text" name="jurnalno" readonly id="jurnalno" size="20" value="<%=jurnalno%>" >
				<input type="text" name="sourcecode" readonly id="sourcecode" size="5" value="<%=sourcecode%>" >
				
			</td>


			<td align="left" >Jurnal Date</td>
			<td align="left">
				<input type="text" name="jurnaldate" readonly id="jurnaldate" size="20" value="<%=jurnaldate%>" >
			</td>
			
		</tr>
		
		<tr>
			<td align="left" >Remark</td>
			<td align="left" colspan="">
			  <textarea  name="remarks" id="remarks" cols="30" rows=30><%=remarks%></textarea>
				
			</td>
			<td align="left" >Posting</td>
			<td align="left" colspan="">
				<%
				
					if POSTINGFLG= "0" then 
						response.write "TIDAK"
					elseif POSTINGFLG= "" then
						response.write "TIDAK"
					else 
						response.write "YA"
					end if
				%>
				
				
			</td>
			
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='800' id='table1'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">JURNAL</li>				
					</ul>

					<div class="TabbedPanelsContentGroup">

						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>NO REFERENCE </th>
									<th>DESCRIPTION </th>
									<th>JOBNAME</th>
									<th>TARGET CODE</th>
									<th>DEBIT</th>
									<th>KREDIT</th>

								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,totalamountkredit
								sqldtlupah = "select d.BKMNO,d.description, j.fccode,d.locationcode, j.fccode || '-' ||  j.fcname as jobcode,sum(d.amountdebit) amountdebit,sum(d.amountkredit)amountkredit " & _ 
								" from BKMJOURNALHDR h " & _
								" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba) " & _
								" inner join jobmaster j on (j.fccode=d.JOBCODE) " & _
								" where h.fcba='" & businessunit & "' AND h.docno = '" & bkmno & "' AND h.DOCDATE=to_date('" & tglbkm & "','dd/mm/yyyy') " & _
								" group by h.fcba,d.BKMNO,d.description,j.fccode,d.locationcode,j.fccode || '-' ||  j.fcname,lineno order by lineno " 
								set rsdtlupah = connStr.Execute(sqldtlupah)
								'response.write sqldtlupah
								

																
								no = 0
								totalamount = 0
								totalamountkredit = 0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("BKMNO")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("description")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("jobcode")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("locationcode")%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("amountdebit"),0)%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("amountkredit"),0)%>
									</td>
								
									<td></td>
									
								</tr>

								
								<%

									

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("amountdebit"))
									totalamountkredit = cdbl(totalamountkredit) + cdbl(rsdtlupah("amountkredit"))
									rsdtlupah.MoveNext
									
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="4" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamountkredit,0)%>"/></td>
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
			                 <% if Access_Post <> 0 Then %>
							 <input type=Submit name="btnsave" id="btnsave" value="POSTING">&nbsp;&nbsp;
			                 <input type=Submit name="btnunposting" id="btnunposting" value="UNPOSTING">&nbsp;&nbsp;
							 <%end if%>
			                <input type=button onclick="window.close();" name="btnclose" value="CLOSE">
							 
							 
			</td>
			
			<!--<td  colspan="5">
			 <input type=Submit name="btnsave" id="btnsave" value="POSTING">&nbsp;&nbsp;<input type=Submit name="btncancel" value="CANCEL">&nbsp;&nbsp;<input type=Submit name="btnunposting" id="btnunposting" value="UNPOSTING">
			 </td> -->
			 
			<input type="hidden" name="POSTINGFLG" id="POSTINGFLG" value="<%=POSTINGFLG%>"/>
			<input type="hidden" name="fcba" id="fcba" value="<%=businessunit%>"/>
			<input type="hidden" name="view" id="view" value="2"/>
			<input type="hidden" name="tipetran" id="tipetran" value="<%=tipetran%>"/>
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
