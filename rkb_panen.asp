<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno 
bkmno= trim(request("rkbno"))
businessunit= trim(session("fcba"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID
dim SUPERVISI1,SUPERVISI2,SUPERVISI3,SUPERVISI4,SUPERVISINAME1,SUPERVISINAME2,SUPERVISINAME3,SUPERVISINAME4,Clasification,Clasificationname,ROWSTATE 
dim period_month,period_year 

sqlheader = "select rkbno,period_month,period_year,sectioncode,document_classification,rowstate,totupah,totbahan,totvra,totwip,classification_api.get_fcname(document_classification) docname " & _
            " from rkbhdr where rkbno='" & bkmno & "' AND fcba='" & businessunit & "'"
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


'count detailupah
Dim SQL_Rows, NoOfRows,RSCount

SQL_Rows = "SELECT Count (linenokey) from rkbupahdtl where rkbno='" & bkmno & "' and fcba='" & businessunit & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close

'count detailbahan
Dim SQL_Rowsitem, NoOfRowsitem,RSCountitem

SQL_Rowsitem = "SELECT Count (linenokey) from rkbbahandtl where rkbno='" & bkmno & "' and fcba='" & businessunit & "'"
set RSCountitem = connStr.Execute(SQL_Rowsitem)
if not RSCountitem.eof then 
	NoOfRowsitem = RSCountitem(0).value
end if
RSCountitem.close




%>

<html>
	<head>
		<title>FBPlants - RKB Panen</title>
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
		width: 300px;
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

	.lebarhaactual{
		width: 70px;
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : RKB PANEN</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="<%=NoOfRowsitem%>" />

	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >RKB NO</td>
			<td align="left">
				<input type="text" name="bkmno" readonly id="bkmno" size="30" value="<%=bkmnohdr%>" title="BKM No">
			</td>


			<td align="left" >Section</td>
			<td align="left">
				<input type="text" name="sectionname" readonly id="sectionname" size="20" value="<%=SECTIONNAME%>">
				
			</td>
			
		</tr>
		<tr>
			<td align="left" >Period</td>
			<td align="left">
			    <select name="tahun" id="tahun" style="width:100px">
				<%
					dim sqltahun7,rstahun7
					sqltahun7 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
					Set rstahun7 = Server.CreateObject("ADODB.Recordset")
					rstahun7.CursorType = 3
					rstahun7.CursorLocation = 3
					rstahun7.Open sqltahun7, connstr
					While Not rstahun7.EOF 
						if cint(period_year)=cint(rstahun7("tahun")) then %>
						<option value="<%=rstahun7("tahun")%>" Selected ><%=rstahun7("tahun")%></option>
						<%else%>
						<option value="<%=rstahun7("tahun")%>"><%=rstahun7("tahun")%></option>
						<%end if%>					
					<%rstahun7.movenext
					wend 
					rstahun7.close 
					

				%>	
				</select>
				<select name="bulan" id="bulan" style="width:120px">
				<%
					dim sqlbulan7,rsbulan7
					sqlbulan7 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
					Set rsbulan7 = Server.CreateObject("ADODB.Recordset")
					rsbulan7.CursorType = 3
					rsbulan7.CursorLocation = 3
					rsbulan7.Open sqlbulan7, connstr
					While Not rsbulan7.EOF 
						if cint(period_month)=cint(rsbulan7("bulan")) then %>
						<option value="<%=rsbulan7("bulan")%>" Selected><%=rsbulan7("description")%></option>
						<%else%>
						<option value="<%=rsbulan7("bulan")%>"><%=rsbulan7("description")%></option>
						<%end if%>
					<%rsbulan7.movenext
					wend 
					rsbulan7.close 
					

				%>	
				</select>
				
				
			</td>
			<td align="left" >Type</td>
			<td align="left">
				
				<select name="clasification" id="clasification" class="lebarsection">
					<OPTION value="<%=Clasification%>"><%=Clasificationname%></OPTION>
				</select>
			</td>
			
			
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='900' id='tableall'>
		<tr>
			<td>
			   		<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB Panen</li>
					
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm upah-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Tahun Tanam</th>
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Rate(Rp)</th>
									<th>Bjr</th>
									<th>KG/JJG</th>
									<th>SB(Rp.)</th>
									<th>LB(Rp.)</th>
									<th>Rate Borongan</th>
									<th>Total Borongan</th>
									
								</tr>

							<%
									DIM ttlluas,ttlsemuahk,ttlrotasi,ttlsemuaamount,sqldtl,rsdtl,no,totalkg
								
									sqldtl = "Select R.TAHUNTANAM, R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE, NVL(R.RATE,0)RATE, " & _ 
									" NVL(R.ABW_BUDGET,0)ABW_BUDGET, NVL(R.KG_BUDGET,0)KG_BUDGET, NVL(R.RATE_SB,0)RATE_SB,NVL(R.RATE_LB,0)RATE_LB, " & _ 
									" NVL(R.RATEBORONGAN,0)RATEBORONGAN,NVL(R.TOTALAMOUNT_BORONGAN,0)TOTALAMOUNT_BORONGAN From RKBUPAHDTL R " & _ 
									" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE " & _ 
									" WHERE R.RKBNO='" & bkmno & "'"
									set rsdtl = connStr.Execute(sqldtl)

																		
									no = 1
									ttlluas = 0
									ttlsemuahk = 0
									ttlrotasi = 0
									ttlsemuaamount = 0
									While Not rsdtl.EOF %>
								<tr id='trb<%=no%>'>
									<td class="ratakiri">
										<select class="lebarhaactual " name="tahuntanam<%=no%>" id="tahuntanam<%=no%>" >
											<option value=<%=rsdtl("TAHUNTANAM")%>><%=rsdtl("TAHUNTANAM")%></option>
										</select>
									</td>
									<td class="ratakiri">
										<select class="lebarsection " name="activityrwt<%=no%>" id="activityrwt<%=no%>" class='lebaractbarang' onchange="getTCrwt(<%=no%>)">
											<option value=<%=rsdtl("ACTCODE")%>><%=rsdtl("fcname")%></option>
										</select>
									</td>

									<td class="ratatengah">
										<select class='lebartc ratatengah' name="tcrwt<%=no%>" id="tcrwt<%=no%>" >
											<option value=<%=rsdtl("LOCATIONCODE")%>><%=rsdtl("LOCATIONCODE")%></option>
										</select>
									</td>

									<td>
										<input class='lebarhaactual ratatengah' readonly type="text"  name="ratedpn<%=no%>" id="ratedpn<%=no%>" value="<%=formatnumber(rsdtl("RATE"),0)%>"/>
										<input class='lebarhaactual ratatengah' readonly type="hidden" style='background-color:#eee;' name="rate<%=no%>" id="rate<%=no%>" value="<%=rsdtl("RATE")%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarlokasi"  type="text"  name="ABW_BUDGETDPN<%=no%>" id="ABW_BUDGETDPN<%=no%>" value="<%=formatnumber(rsdtl("ABW_BUDGET"),2)%>"/>
										<input class="ratatengah lebarlokasi" readonly type="hidden" style='background-color:#eee;' name="bjrbudget<%=no%>" id="bjrbudget<%=no%>" value="<%=formatnumber(rsdtl("ABW_BUDGET"),2)%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarhaactual" type="text" onkeyup="hanyaAngka(this);getKgbudget(<%=no%>)" name="kgbudgetdpn<%=no%>" id="kgbudgetdpn<%=no%>" value="<%=formatnumber(rsdtl("KG_BUDGET"),0)%>"/>
										<input class="ratatengah lebarhaactual" type="hidden" name="kgbudget<%=no%>" id="kgbudget<%=no%>" value="<%=rsdtl("KG_BUDGET")%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual"  type="text"  name="sbdepan<%=no%>" id="sbdepan<%=no%>" value="<%=formatnumber(rsdtl("RATE_SB"),0)%>"/>
										<input class="ratatengah lebarhaactual" readonly type="hidden" style='background-color:#eee;' name="sb<%=no%>" id="sb<%=no%>" value="<%=rsdtl("RATE_SB")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebartc" type="text" name="lb<%=no%>" id="lb<%=no%>" value="<%=rsdtl("RATE_LB")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text" name="rateborongan<%=no%>" id="rateborongan<%=no%>" value="<%=rsdtl("RATEBORONGAN")%>" onkeyup="getAmountborongan('<%=no%>')"/>
									</td>

									<td class="ratakiri">
										<input class="ratakanan lebarlokasi" readonly type="text" name="amountborongan<%=no%>" id="amountborongan<%=no%>" value="<%=formatnumber(rsdtl("TOTALAMOUNT_BORONGAN"),0)%>"/>
										<input class="ratakanan lebarlokasi" readonly type="hidden" name="amountboronganbck<%=no%>" id="amountboronganbck<%=no%>" value="<%=rsdtl("TOTALAMOUNT_BORONGAN")%>"/>
									</td>

								
								</tr>
								<%
									totalkg = cdbl(totalkg)+cdbl(rsdtl("KG_BUDGET"))
									no = no + 1 
									rsdtl.MoveNext
									Wend
								
								%>
								
								<tr class="ratakanan">
										<td colspan="5" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalkg,0)%>"/></td>
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

$(function() {	
	
	showClasification();
	
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



//get clasification 
function showClasification() {
	//$('#clasification').empty();
	$('#clasification').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_clasification.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


</script>
