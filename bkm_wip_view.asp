<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,gangcode,bkmno,totalpremi,bkmdate
dim ScriptName,btnApprove,bkmno1,SQL
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
businessunit= trim(session("fcba"))


q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

bkmno		= trim(request("bkmno"))


function Akses_User
Dim rsUsrPosting, strUsrPosting,Akses_Usr
    Akses_Usr  =  0
 	strUsrPosting =	" SELECT 1 result " & _
					" FROM bkmhdr hdr " & _
					" INNER JOIN authorizetemplate at ON (at.document_type = 'AA' AND at.document_classification = hdr.document_classification)" & _
					" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "')" & _
					" WHERE hdr.bkmno='" & trim(request("bkmno")) & "' " & _
					" AND hdr.rowstate = 'Planned'" & _
					" AND hdr.postingflg = 0" & _
					" AND at.fccode = '504'" & _
					" AND hdr.fcba='" & trim(session("fcba")) & "' " & _
					" AND authorizelevelcode = 900" & _
				    " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' AND sectionname='" & SECTIONNAME & "') " 
 				
 	Set rsUsrPosting = Server.CreateObject("ADODB.Recordset")
 	rsUsrPosting.Open strUsrPosting, connstr, 3
	
 	If Not rsUsrPosting.EOF Then
 		Akses_Usr = cdbl(rsUsrPosting("result"))
 	Else
 		Akses_Usr = 0
 	End If
 	rsUsrPosting.Close
 	set rsUsrPosting = Nothing
 
	Akses_User = Akses_Usr
end function 

function Akses_Posting
Dim rsAvailable, strAvailable,Access_Post
    Access_Post  =  0
 	strAvailable =	" select authorizestructure_api.get_access_post('504','" & businessunit & "','" & Session("CodeName") & "') result" & _
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
 
	Akses_Posting = Access_Post
end function


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,CONTRACTORCODE,POSTINGFLG,CONTRACTORNAME,ROWSTATE
sqlheader = "select h.BKMNO,TO_CHAR(h.BKMDATE,'dd/mm/yyyy')BKMDATE,h.CONTRACTORCODE,c.FCNAME || ' - ' || c.FCCODE CONTRACTORNAME, " & _
" h.DOCUMENT_CLASSIFICATION,h.POSTINGFLG,h.ROWSTATE " & _
" from bkmhdr h INNER JOIN CONTRACTOR c on (c.fccode=h.CONTRACTORCODE and c.fcba=h.fcba) where h.BKMNO='" & bkmno & "'"
'response.write sqlheader
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	bkmnohdr = rsHeader("BKMNO")
	DATECREATED = rsHeader("BKMDATE")
	CONTRACTORCODE = rsHeader("CONTRACTORCODE")
	POSTINGFLG = rsHeader("POSTINGFLG")
	CONTRACTORNAME = rsHeader("CONTRACTORNAME")
	ROWSTATE = rsHeader("ROWSTATE")
	
	
end if

IF POSTINGFLG = "0" THEN 
   POSTINGFLG = "TIDAK"
ELSE
   POSTINGFLG	= "YA"
END IF

'count detailwip
Dim SQL_Rowswip, NoOfRowswip,RSCountwip

SQL_Rowswip = "SELECT Count (*) from BKMWIPDTL where BKMNO='" & bkmno & "'"
set RSCountwip = connStr.Execute(SQL_Rowswip)
if not RSCountwip.eof then 
	NoOfRowswip = RSCountwip(0).value
end if
RSCountwip.close

'proses dibawah ini digunakan untuk approved bkm dan create jurnal
btnApprove 	= request("btnsave")
bkmno1 		= request("bkmno1")
bkmdate 	= request("bkmdate")
dim strQueryStatus,modulFile,executeQry,modulFile1,executeQry1
if btnApprove = "Approve" then 

	strQueryStatus = "UPDATE bkmhdr SET rowstate = 'Approved', postingflg='1',fcedit='" & Session("CodeName") & "',lastupdate=sysdate WHERE fcba = '" & session("fcba") & "' AND bkmno = '" & bkmno1 & "'"
	connStr.execute(strQueryStatus)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & strQueryStatus
	ExecuteAndLog  modulFile,executeQry

	SQL = "begin bkmjournal_api.Update_bkm(TO_DATE('" & bkmdate & "', 'dd/mm/yyyy'),'" & session("fcba") & "','" & Session("CodeName") & "',1); end;"
	connStr.execute(SQL)
	modulFile1 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry1	= "Query executed:" & SQL
	ExecuteAndLog modulFile1,executeQry1
	response.redirect("bkm_wip_view.asp?bkmno="&bkmno1)

end if



%>

<html>
	<head>
		<title>FBPlants - View BKM WIP</title>
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
		left: -15px;
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
		width: 45px;
	}
	.lebarlokasi{
		width: 50px;
	}
	.lebartc{
		width: 60px;
	}
	.lebarovtime{
		width: 70px;
	}
	.lebarrate{
		width: 90px;
	}
	.lebaractivity{
		width: 190px;
	}
	.lebarallow{
		width: 90px;
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
		width: 230px;
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
		<strong class="">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM WIP VIEW</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_wip_view.asp">
	<input type="hidden" name="jumldatawip" id="jumldatawip" value="<%=NoOfRowswip%>" />
	<input type="hidden" name="bkmno1" id="bkmno1" value="<%=bkmno%>" />
	<input type="hidden" name="bkmdate" id="bkmdate" value="<%=DATECREATED%>" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >BKM No</td>
			<td align="left">
				<%=bkmno%>
			</td>
			<td align="center" >&nbsp;&nbsp;&nbsp;Tanggal</td>
			<td align="left">
				<%=DATECREATED%>
			</td>
			<td align="left" >Contractor Name</td>
			<td align="left">
				<%=CONTRACTORCODE%>&nbsp;<%=CONTRACTORNAME%>
			</td>
			
		</tr>
		<tr>
			
			<td align="left" >Posting</td>
			<td align="left"><%=POSTINGFLG%>-<%=ROWSTATE%></td>
			
			
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='850' id='tableall'>
		<tr>
			<td colspan=>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM WIP</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm WIP-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="850" class='data' id='table4'>
									<tr id="trheader4">
										<th>Lokasi</th>
										<th>SPK No.</th>
										<th>Activity</th>
										<th>Uom</th>
										<th>LineNo</th>
										<th>Rate</th>
										<th>TC</th>
										<th>Qty</th>
										<th>Unit</th>
										<th>Output</th>
										<th colspan="2">&nbsp;</th>
									</tr>
									<%
									
										dim sqlwip,rswip
										sqlwip = "Select wi.CONTRACTORCODE,ct.FCCODE || ' - ' || ct. FCNAME CONTRACTORNAME,wi.AGREEMENTCODE,wi.ACTCODE,c.FCCODE || ' - ' || c.FCNAME COANAME, " & _
										" wi.LOCATIONCODE, wi.QTY, wi.OUTPUT, wi.UNIT,wi.CHARGEUNIT,wi.BKMAGREEMENTLINENO,wi.RATE ,c.uom_unit" & _
										" From   BKMWIPDTL wi " & _
										" INNER JOIN Jobmaster c on (wi.ACTCODE=c.FCCODE) " & _
										" INNER JOIN CONTRACTOR  ct on (wi.CONTRACTORCODE=ct.FCCODE AND wi.FCBA=ct.FCBA) WHERE wi.BKMNO='" & bkmno & "' order by wi.AGREEMENTCODE asc "
										'' response.write sqlwip
										 set rswip = connStr.Execute(sqlwip)
									%>

									<% 
										dim nowip
										nowip = 0
										While Not rswip.EOF 
									%>
										<tr id='trw<%=nowip%>'>
											<td class="ratakiri">
												<%=rswip("CHARGEUNIT")%>
											</td>

											<td class="ratakiri">
												<%=rswip("AGREEMENTCODE")%>
											</td>

																					
											<td class="ratakiri">
												<%=rswip("ACTCODE")%>&nbsp;<%=rswip("COANAME")%>
											</td>
											
											<td class="ratakiri">
												<%=rswip("uom_unit")%>
											</td>

											<td class="ratakiri">
												<%=rswip("BKMAGREEMENTLINENO")%>
											</td>

											<td class="ratatengah">
												<%=rswip("RATE")%>
											</td>


											<td class="ratatengah">
												<%=rswip("LOCATIONCODE")%>
											</td>

											<td class="ratatengah">
												<%=rswip("QTY")%>
											</td>

											<td class="ratatengah">
												<%=rswip("UNIT")%>
											</td>

											<td class="ratatengah">
												<%=rswip("OUTPUT")%>
											</td>

											
											<td></td>
										</tr>

											
										
									<%
										nowip = nowip + 1
										rswip.MoveNext
										Wend
									
									%>
								</table>
							</div>
						</div>


					</div>
				</div>
			</td>
		</tr>
		<tr>
		<td colspan="9" align="center" >
		<% if Akses_Posting <> 0 and ROWSTATE="Planned" and POSTINGFLG="TIDAK" Then %>
			<input type=Submit name="btnsave" id="btnsave" value="Approve">&nbsp;&nbsp;
		<%end if%>
		    <input type=button onclick="window.close();" name="btnclose" value="CLOSE">
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
	//window.onload = getcountEmployee();
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
	$("#tglbkm").datepicker();
	showContractor();
	getLoadWIP();
	
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



//get contractor
function showContractor() {
	//$('#contractorcode').empty();
	var jmlitem = 7;
	$('#contractorcode').select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem,
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

function getActivity(i)
{
	var spkno = $('#spkno' + i).val();

	$('#activitywip' + i).empty();
	$('#activitywip' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 3 + '&spkno=' + spkno,
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


function getLine(i)
{
	var spkno = $('#spkno' + i).val();
	var activityiwip = $('#activitywip' + i).val();

	$('#linenowip' + i).empty();
	$('#linenowip' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_show_item.asp?jmlitem=' + 8 + '&activitywip=' + activityiwip + '&agreement=' + spkno,
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


//getEmployee
function getRate(i) {
	
	var spkno = $('#spkno' + i).val();
	var activityiwip = $('#activitywip' + i).val();
	var linenowip = $('#linenowip' + i).val();
	
	if (linenowip != "") {
		$('#ratewip' + i).empty();
		$.post('bkm_show_item.asp?jmlitem=' + 9 + '&activitywip=' + activityiwip + '&agreement=' + spkno + '&linenowip=' + linenowip, function(data) {
			var json  = JSON.parse(data);
			var ratewip = json.results[0].rate;
			$("#ratewip" + i).val(ratewip);

		});

	}	
	
}

function getTC(i)
{
	var lokasi   = $('#lokasi' + i).val();
	$('#tcwip' + i).empty();
	$('#tcwip' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_gangcode.asp?gancode=' + 3 +'&lokasi=' + lokasi,
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



var jumlwip = $('#jumldatawip').val();
var StartIterationwip =  jumlwip + 1
function addrowWIP(trid)
{

	var lstwip = this.StartIterationwip++;
	if (lstwip < 2) {
		var row = $("#trheader4").after("<tr id='trw" + lstwip + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trw' + trid).after("<tr id='trw" + lstwip + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowWIP(" + lstwip + ")' />";
	}

	var lokasi = "<select name='lokasi" + lstwip + "' id='lokasi" + lstwip + "'  onchange='getTC(" + lstwip + ")' class='lebarlokasi'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + lokasi + "</td>");

	var spkno = "<select name='spkno" + lstwip + "' id='spkno" + lstwip + "' class='lebarspkno'  onchange='getActivity(" + lstwip + ")'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + spkno + "</td>");

	var activitywip = "<select name='activitywip" + lstwip + "' id='activitywip" + lstwip + "' onchange='getLine(" + lstwip + ")' class='lebaractivity'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + activitywip + "</td>");

	var linenowip = "<select name='linenowip" + lstwip + "' id='linenowip" + lstwip + "' onchange='getRate(" + lstwip + ")' class='lebartc ratatengah'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + linenowip + "</td>");

	var ratewip = "<input type='text' name='ratewip" + lstwip + "' id='ratewip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' readonly class='lebarrate ratakanan'>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + ratewip + "</td>");

	var tcwip = "<select name='tcwip" + lstwip + "' id='tcwip" + lstwip + "' class='lebartc ratatengah'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + tcwip + "</td>");

	var qtywip = "<input type='text' name='qtywip" + lstwip + "' id='qtywip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' class='lebarabsen ratatengah'>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + qtywip + "</td>");

	var unitwip = "<input type='text' name='unitwip" + lstwip + "' id='unitwip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' class='lebarabsen ratatengah'>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + unitwip + "</td>");

	var outputwip = "<input type='text' name='outputwip" + lstwip + "' id='outputwip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' class='lebarabsen ratatengah'>";
	$('#trw' + lstwip).append("<td class='ratatengah'>" + outputwip + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowWIP(" + lstwip + ")' />";
	$('#trw' + lstwip).append("<td>" + tambahbutton + "</td>");

	$('#trw' + lstwip).append("<td>" + kurangbutton + "</td>");

	$("#jumldatawip").val(lstwip);
	showFCBA("lokasi" + lstwip);
	showSPK("spkno" + lstwip);


	
}

// DELETE ROW WIP
function deleterowWIP(i) {
	Spry.Widget.Utils.destroyWidgets('spkno' + i);
	Spry.Widget.Utils.destroyWidgets('activitywip' + i);
	Spry.Widget.Utils.destroyWidgets('linenowip' + i);
	Spry.Widget.Utils.destroyWidgets('ratewip' + i);
	Spry.Widget.Utils.destroyWidgets('tcwip' + i);
	Spry.Widget.Utils.destroyWidgets('qtywip' + i);
	Spry.Widget.Utils.destroyWidgets('unitwip' + i);
	Spry.Widget.Utils.destroyWidgets('outputwip' + i);
	$('#trw' + i).remove();
	
}





//getLoadwip
function getLoadWIP() {
	
	var bkmno = $("#bkmno").val();
	var jmlitem = 6;
	$.post("bkm_show_item.asp?jmlitem=" + jmlitem +"&bkmno=" + bkmno, function(data) {
		var json  = JSON.parse(data);
		var jumlview  = json.results.length;
		
		for (var i = 0; i < jumlview; i++) {

			showFCBA("lokasi" + i);
			showSPK("spkno" + i);
			
		}
	});
	
	
}


//get lokasi upah
function showFCBA(selid) {
	var fcba = 'fcba';
	$('#lokasi').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp?viewdata=' + fcba,
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

//get WIP
function showSPK(selid) {
	//$('#spkno').empty();

	var jmlitem = 5;
	var tglbkm = $('#tglbkm').val();
	var contractorcode = $('#contractorcode').val();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem + "&contractorcode=" + contractorcode + "&tglbkm=" + tglbkm,
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



//new Spry.Widget.ValidationTextField("DATECREATED", "none");
//new Spry.Widget.ValidationSelect("clasification");
</script>