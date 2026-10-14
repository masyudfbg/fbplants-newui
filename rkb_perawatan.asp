<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno 
bkmno= trim(request("rkbno"))
businessunit= trim(request("fcba"))

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
		<title>FBPlants - RKB Perawatan</title>
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
	
</style>



<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : RKB PERAWATAN</strong><br>					
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
						<li class="TabbedPanelsTab">RKB UPAH</li>
						<li class="TabbedPanelsTab">RKB MATERIAL</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm upah-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Aktivitas</th>
									<th>Upah(Rp)</th>
									<th>Norma(Hk)</th>
									<th>Blok</th>
									<th>HA</th>
									<th>Rencana(Ha)</th>
									<th>Rotasi</th>
									<th>Total.HK</th>
									<th>Total.Upah</th>
									<th colspan="2">&nbsp;</th>
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupahrw,rsdtlupahrw,norw,totalamountrw
								sqldtlupahrw = "SELECT ACTCODE,jobmaster_api.get_fcname(actcode) activityname,nvl(rate,0)rate,nvl(normamandays,0)normamandays,locationcode,field_api.get_ha(locationcode,fcba)field_ha,nvl(volume,0)volume,nvl(rotasi,0)rotasi,nvl(mandays,0)mandays,nvl(totalamount_hk,0)totalamount_hk " & _
											   " FROM RKBUPAHDTL  " & _
								               " WHERE rkbno='" & bkmno & "' and fcba='" & businessunit & "'" 
											   
								set rsdtlupahrw = connStr.Execute(sqldtlupahrw)
								'response.write sqldtlupah
								norw = 0
								totalamountrw = 0
								While Not rsdtlupahrw.EOF %>
								<tr id="tru<%=norw%>">
									
									<td class="ratakiri">
										<select class="lebaractivity ratakiri" name="actcode<%=norw%>" id="actcode<%=norw%>">
											<option value=<%=rsdtlupahrw("ACTCODE")%>><%=rsdtlupahrw("activityname")%></option>
										</select>
									</td>
									
									<td class="ratakanan">
										<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="rate<%=norw%>" id="rate<%=norw%>" value="<%=FormatNumber(rsdtlupahrw("rate"),0)%>"/>
									</td>
									
									<td class="ratakanan">
										<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="normamandays<%=norw%>" id="normamandays<%=norw%>" value="<%=rsdtlupahrw("normamandays")%>"/>
									</td>
									
									<td class="ratatengah">
										<select class="lebartc ratakiri" name="locationcode<%=norw%>" id="locationcode<%=norw%>">
											<option value=<%=rsdtlupahrw("LOCATIONCODE")%>><%=rsdtlupahrw("LOCATIONCODE")%></option>
										</select>
									</td>
									
									<td class="ratakanan">
										<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="field_ha<%=norw%>" id="field_ha<%=norw%>" value="<%=FormatNumber(rsdtlupahrw("field_ha"),0)%>"/>
									</td>
									
									<td class="ratakanan">
										<input class="ratakanan lebarupah" type="text"  name="volume<%=norw%>" id="volume<%=norw%>" value="<%=FormatNumber(rsdtlupahrw("volume"),0)%>"/>
									</td>
									
									<td class="ratakanan">
										<input class="ratakanan lebarupah" type="text"  name="rotasi<%=norw%>" id="rotasi<%=norw%>" value="<%=FormatNumber(rsdtlupahrw("rotasi"),0)%>"/>
									</td>
									
									<td class="ratakanan">
										<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="mandays<%=norw%>" id="mandays<%=norw%>" value="<%=FormatNumber(rsdtlupahrw("mandays"),0)%>"/>
									</td>
									
									<td class="ratakanan">
										<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="totalamount_hk<%=norw%>" id="totalamount_hk<%=norw%>" value="<%=FormatNumber(rsdtlupahrw("totalamount_hk"),0)%>"/>
									</td>
									
									
								
								</tr>
								<%
									norw = norw +1
									totalamountrw = cdbl(totalamountrw) + cdbl(rsdtlupahrw("totalamount_hk"))
									rsdtlupahrw.MoveNext
									Wend
								%>
								
								<tr class="ratakanan">
										<td colspan="8" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=norw%>" value="<%=FormatNumber(totalamountrw,0)%>"/></td>
								</tr>
							</table>
						</div>

						<!-- view bkm bahan-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="850" class='data' id='table1'>
									<tr id="trheader1">
									    <th>Aktivitas</th>
										<th>Blok</th>
										<th>HA</th>
										<th>Material</th>
										<th>Satuan</th>
										<th>Norma</th>
										<th>Qty</th>
										<th>Rate (Rp)</th>
										<th>Total Material(rp) </th>
										<th colspan="2">&nbsp;</th>
									</tr>

									<%
									
										dim sqlbarang,rsbarang,totalamount_br,nobr
										sqlbarang = " SELECT ACTCODE,jobmaster_api.get_fcname(actcode) activityname,locationcode,field_api.get_ha(locationcode,fcba)field_ha,stockcode,c.fcname stockname, uom,activematerialnorm,qty,ratestock,totalamount,linenokey " & _
													" from RKBBAHANDTL b inner join activematerial c on b.stockcode=c.fccode" & _
													" where b.rkbno='" & bkmno & "' and b.fcba='" & businessunit & "' order by b.linenokey asc"
										 set rsbarang = connStr.Execute(sqlbarang)
										 'response.write sqlbarang
									%>

									<% 
										totalamount_br = 0
										nobr = 0
										While Not rsbarang.EOF 
									%>
										<tr id='trb<%=nobr%>'>
										    <td class="ratakiri">
											<select class="lebaractivity ratakiri" name="actcodebr<%=nobr%>" id="actcodebr<%=nobr%>">
												<option value=<%=rsbarang("ACTCODE")%>><%=rsbarang("activityname")%></option>
											</select>
											</td>
											
											<td class="ratatengah">
											<select class="lebartc ratakiri" name="locationcodebr<%=nobr%>" id="locationcode<%=nobr%>">
												<option value=<%=rsbarang("LOCATIONCODE")%>><%=rsbarang("LOCATIONCODE")%></option>
											</select>
											</td>
										
											<td class="ratakanan">
												<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="field_habr<%=nobr%>" id="field_habr<%=nobr%>" value="<%=FormatNumber(rsbarang("field_ha"),0)%>"/>
											</td>
											
											<td class="ratakiri">
												<select class="lebaractivity " name="stockcode" id="stockcode<%=nobr%>">
													<option value=<%=rsbarang("stockcode")%>><%=rsbarang("stockname")%></option>
												</select>
											</td>
											
											<td class="ratatengah">
												<input class="ratatengah lebarabsen" type="text" readonly  style="background-color:#eee;" name="uom" id="uom<%=nobr%>" value="<%=rsbarang("uom")%>"/>
											</td>
											
											<td class="ratatengah">
												<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="normabr" id="normabr<%=nobr%>" value="<%=FormatNumber(rsbarang("activematerialnorm"),0)%>"/>
											</td>
											
											<td class="ratatengah">
												<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="qtyitem" id="qtyitem<%=nobr%>" value="<%=FormatNumber(rsbarang("qty"),0)%>"/>
											</td>

											<td class="ratatengah">
												<input class="ratakanan lebarupah" type="text" name="ratestock" id="ratestock<%=nobr%>" value="<%=FormatNumber(rsbarang("ratestock"),0)%>"/>
											</td>

											<td class="ratatengah">
												<input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="totalamountbr" id="totalamountbr<%=nobr%>" value="<%=FormatNumber(rsbarang("totalamount"),0)%>"/>
											</td>



										</tr>

											
										
									<%
										nobr = nobr + 1
										totalamount_br = cdbl(totalamount_br) + cdbl(rsbarang("totalamount"))
										rsbarang.MoveNext
										Wend
									
									%>
									
									<tr class="ratakanan">
										<td colspan="8" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=norw%>" value="<%=FormatNumber(totalamount_br,0)%>"/></td>
									</tr>
									
									
									
								</table>
							</div>
							
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
