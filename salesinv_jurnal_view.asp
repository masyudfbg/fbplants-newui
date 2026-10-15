<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,invoicecode,tglinvoice,tglsplit,bulansplit,thnsplit
invoicecode	= request("invoicecode")
tglinvoice		= request("tglinvoice")
businessunit = request("fcba")


q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	'response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

'count detailupah
Dim SQL_Rows, NoOfRows,RSCount,SQLJurnal

SQL_Rows = "SELECT Count (docno) from bkmjournalhdr where docno='" & invoicecode & "' and fcba='" & businessunit & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close

 if TRIM(NoOfRows) = 0 then
 SQLJurnal = "BEGIN fac_journal_com_api.COM_SLS_INV_DOCNO('" & invoicecode & "',TO_NUMBER(TO_CHAR(TO_DATE('" & tglinvoice & "','DD/MM/YYYY'),'YYYY')), " & _ 
																		 "TO_NUMBER(TO_CHAR(TO_DATE('" & tglinvoice & "','DD/MM/YYYY'),'MM')) , '" & businessunit & "','System'); " & _ 
		 "END;"
		 'response.write sqlfac													 															 
 connStr.execute(SQLJurnal)
 end if 



dim sqlheader,rsHeader,jurnalno,jurnaldate,remarks,sourcecode,POSTINGFLG
sqlheader = "select docno, to_char(docdate,'dd/mm/yyyy')docdate,remarks,sourcecode,POSTINGFLG from bkmjournalhdr where docno='" & invoicecode & "'"
set rsHeader = connStr.Execute(sqlheader)
'response.write sqlheader
if not rsHeader.eof then 
	jurnalno = rsHeader("docno")
	jurnaldate = rsHeader("docdate")
	remarks = rsHeader("remarks")	
	sourcecode = rsHeader("sourcecode")
	POSTINGFLG = rsHeader("POSTINGFLG")
end if

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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : SALES INVOICE JURNAL</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="salesinv_jurnal_proses.asp">

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
									
									<th>DESCRIPTION </th>
									<th>JOBNAME</th>
									<th>TARGET CODE</th>
									<th>DEBIT</th>
									<th>KREDIT</th>

								</tr>

								<% 
								
								
								'get data detail sales inv
								dim sqldtlupah,rsdtlupah,no,totalamount,totalamountkredit
								sqldtlupah = "select d.BKMNO,d.description, j.fccode,d.locationcode, j.fccode || '-' ||  j.fcname as jobcode,sum(d.amountdebit) amountdebit,sum(d.amountkredit)amountkredit " & _ 
								" from BKMJOURNALHDR h " & _
								" inner join BKMJOURNALDTL d on  (h.DOCNO=d.DOCNO and h.fcba=d.fcba) " & _
								" inner join jobmaster j on (j.fccode=d.JOBCODE) " & _
								" where h.docno = '" & invoicecode & "' " & _
								" group by h.fcba,d.BKMNO,d.description,j.fccode,d.locationcode,j.fccode || '-' ||  j.fcname,lineno order by lineno " 
								set rsdtlupah = connStr.Execute(sqldtlupah)
								'response.write sqldtlupah
								

																
								no = 0
								totalamount = 0
								totalamountkredit = 0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>">
									
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
										<td colspan="3" class="tebal">Total</td>
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
			                
							 <input type=Submit name="btnsave" id="btnsave" value="POSTING">&nbsp;&nbsp;
			                 <input type=Submit name="btnunposting" id="btnunposting" value="UNPOSTING">&nbsp;&nbsp;
							
			                
							 <input type=button onclick="window.close();" name="btnclose" value="CLOSE">
							 
							 
			</td>
			
			<!--<td  colspan="5">
			 <input type=Submit name="btnsave" id="btnsave" value="POSTING">&nbsp;&nbsp;<input type=Submit name="btncancel" value="CANCEL">&nbsp;&nbsp;<input type=Submit name="btnunposting" id="btnunposting" value="UNPOSTING">
			 </td> -->
			 
			<input type="hidden" name="POSTINGFLG" id="POSTINGFLG" value="<%=POSTINGFLG%>"/>
			<input type="hidden" name="fcba" id="fcba" value="<%=businessunit%>"/>
			<input type="hidden" name="view" id="view" value="1"/>
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


// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
function numformat2(val) {
	var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
	if (isNaN(n)) return val == null ? '' : val;
	var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
	return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
}

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


function actionBtn()
	{
		var bkmfrom = $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var jumlinv = parseInt($('#jumldata').val());
		
		if (bkmfrom != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?bkmfrom=" + bkmfrom + "&bkmto=" + bkmto,function(data) {
				$('#bodytable').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var x = json.results[i].fccode
					//alert(x);
					var docketlist = '';
						docketlist += '<tr id="tru'+i+'">';
						docketlist += '<td align="left">' + json.results[i].bkmno + '</td>';
						docketlist += '<td align="left">' + json.results[i].createdate + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].totalbkmupah) + '</td>';
						docketlist += '<td align="right">' + json.results[i].totalbkmmaterial + '</td>';				
						docketlist += '<td align="right">' + json.results[i].totalbkmvehicle + '</td>';
						docketlist += '<td align="right">' + json.results[i].totalbkmwip + '</td>';
						docketlist += '<td align="center"><a target=_blank href="bkm_operasional_edit.Asp?bkmno='+ json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data border="0"></a></td>';
						docketlist += '<td align="center"><a target=_blank href="bkm_jurnal_input.Asp?tglbkm='+ json.results[i].createdate + '"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
						docketlist += '</tr>';
						$('#bodytable').append(docketlist);
					no++;
					$("#jumldata").val(jumlview);
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				
			});
		 
		} else {
			$('#bodytable').empty();
		}
	}



</script>
