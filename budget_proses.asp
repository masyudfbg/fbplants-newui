<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>
	<head>
		
		<title>FBPlants - Budget Proses</title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>

 <script>
$(function() {
	//getFCBA();
});

</script>

<%


%>


</head>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiriisi{
		position: relative;
		left: 0px;
	}
	.lebarkolomsatu{
		width: 10px;
	}
	.lebarkolomdua{
		width: 10px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
	.lebarselect{
		width: 200px;
	}
</style>
<body>
<!--#include file="top.jsp"-->

<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
<div class="judul">
	   Module : PROSES BUDGET
</div>
<div class="box_both" style="overflow:scroll;">




<form name="frmEntry" method="post" action="p_budgetproses.asp">
	
	<table class="data" border="0" width="600px">
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="2" >BUDGET PROSES</td>
		</tr>
			
		<tr>
			<td align="left" class="">PT.</td>
			<td align="left" class="" colspan="">
				<select name="company" id="company" class="lebarselect" onchange="getFCBA()">
					<option value=""></option>
					<% 	
							Dim q,rsCompany,company
							q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" class="">FCBA</td>
			<td align="left" class="" colspan="">
				<select name="fcba" id="fcba" class="lebarselect" onchange="showKegiatan();showKegiatan2()">
					<option value=""></option>
				</select>
			</td>
		</tr>
		
		<tr>
			<td align="left" class="">Tahun</td>
			<td align="left" class="">
				<select name="tahun" id="tahun" class="lebarselect">
					<option value=""></option>
					<% 	
							Dim qq,rsTahun,tahun
							qq = " SELECT PERIOD_YEAR " & _
							 	" FROM MONTHLYWORKPLAN " & _
							    " GROUP BY PERIOD_YEAR order by PERIOD_YEAR ASC"
							Set rsTahun = Server.CreateObject("ADODB.Recordset")
							rsTahun.CursorType = 3
							rsTahun.CursorLocation = 3
							rsTahun.Open qq, connstr
							if not rsTahun.eof then%>
							<%
							Do While Not rsTahun.eof
							%>
								<option value="<%=rsTahun("PERIOD_YEAR")%>" 
									<%if cstr(""&tahun)=cstr(""&rsTahun("PERIOD_YEAR")) then %> selected 
									<%End If%>><%=rsTahun("PERIOD_YEAR")%>
								</option> 
							<%rsTahun.movenext
							loop
							End if
							rsTahun.Close
							set rsTahun = Nothing
							qq=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" class="">Kegiatan</td>
			<td align="left" class="">
				<select name="kegiatan1" id="kegiatan1" class="lebarselect">
					<option value=""></option>
				</select>&nbsp;&nbsp;s/d&nbsp;&nbsp;<select name="kegiatan2" id="kegiatan2" class="lebarselect">
					<option value=""></option>
				</select>
			</td>
		</tr>
		
		<tr>
			<td align="center" colspan="">
		        <input Type="submit"  name="proses" value="PROSES">
		        
			</td>
		</tr>
	</table>
</div>
</div>
</form>
</div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="plugins/chosen/chosen.jquery.min.js"></script>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script language="JavaScript">

function getFCBA()
	{
		var company = $("#company").val();
		showFCBA(company);
		

	}

	function showFCBA(company) {
				$('#fcba').empty();
				$('#fcba').select2({
					ajax: {
						delay: 1000,
						url : 'getlistfcbaajax.asp?fccompanycode=' + company,
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

		function showKegiatan() {
				var fcba = $('#fcba').val();
				$('#kegiatan1').empty();
				$('#kegiatan1').select2({
					ajax: {
						delay: 1000,
						url : 'getListCOA.asp?fcba=' + fcba,
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

		function showKegiatan2() {
				var fcba = $('#fcba').val();
				$('#kegiatan2').empty();
				$('#kegiatan2').select2({
					ajax: {
						delay: 1000,
						url : 'getListCOA.asp?fcba=' + fcba,
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