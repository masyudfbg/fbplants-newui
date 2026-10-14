<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth,rowspan
Dim ScriptName, NO, cClass
Dim nSALDOAWAL, nDEBET, nKREDIT, nSALDOAKHIR,v_txtFCBA,objRSArray,arrRS,strsql
Function GetRSArray(strSQL)
			set objRSArray = connStr.Execute(strSQL)
			If Not objRSArray.EOF Then
				arrRS = objRSArray.GetRows()
				GetRSArray = arrRS
			Else
                Exit Function
			End If
			objRSArray.Close()
			Set objRSArray = Nothing
End Function
v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtFCBA		=	Request("txtFCBA")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if
if request("txtFCBA")="" then
   v_txtFCBA	=	Session("FBCA")
end if

Dim result, strquery

'--New Query--
strquery	=		" SELECT" & _
					"     s.itemcode" & _
					"     , stock_api.get_stockName(s.itemcode,s.fcba) stockName" & _
					"     , stock_api.get_StockGroup(s.itemcode,s.fcba) StockGroup" & _
					"     , stockgroup_api.Get_StockGrpName(stock_api.get_StockGroup(s.itemcode,s.fcba),s.fcba) StockGroupName" & _
					"     , (s.openingbalanceqty+s.qtyreceived+s.qtyreturned-s.qtywriteoff-s.qtytransfered-s.qtyissued) SOH" & _
					"     , Stocktranshist_api.GetTypeMovingPart(s.fcba,s.itemcode,last_day(to_date('01/'||lpad(s.month,2,0)||'/'||s.year||'','DD/MM/YYYY')) + 1) TypeMovingPart" & _
					"     , round(s.amtperunit,0) amtperunit" & _
					"     , round(s.amtperunit,0) * (s.openingbalanceqty+s.qtyreceived+s.qtyreturned-s.qtywriteoff-s.qtytransfered-s.qtyissued) amttot" & _
					" from stockledger s" & _
					" WHERE fcba='" & v_txtFCBA & "'" & _
					" AND year=" & v_txtYear & _
					" AND month=" & v_txtMonth & _
					" AND (s.openingbalanceqty+s.qtyreceived+s.qtyreturned-s.qtywriteoff-s.qtytransfered-s.qtyissued) > 0 " & _
					" AND storecode='ST1'"

strquery	=	strquery & " ORDER BY StockGroup,s.itemcode"

'Response.Write strquery
arrRS = GetRSArray(strquery)
'Set result = Server.CreateObject("ADODB.Recordset")
'result.Open strquery, connstr, adOpenStatic
%>
<%
Function StripSpecialChar(strInput)
	Dim objRegExp
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "[\']"
	'//Replace all character matches with an Empty String
	StripSpecialChar = objRegExp.Replace(strInput, "")  
	Set objRegExp = Nothing
End Function
%>
<script>

var result_array = {<% If IsArray(arrRS) Then %>
'group_code':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%=arrRS(2, i)%>',<%Next%>],
'kode_stock':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%=arrRS(0, i)%>',<%Next%>],
'nama_stock':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%=StripSpecialChar(arrRS(1, i))%>',<%Next%>],
'soh':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%=FormatNumber(arrRS(4, i),0)%>',<%Next%>],
'type':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%=StripSpecialChar(arrRS(5, i))%>',<%Next%>],
'group_name':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%=StripSpecialChar(arrRS(3, i))%>',<%Next%>],
'amtperunit':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%=FormatNumber(arrRS(6, i),0)%>',<%Next%>],
'amttot':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%= arrRS(7, i)%>',<%Next%>],
'amttotFormat':[<% For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>'<%= FormatNumber(arrRS(7, i),0)%>',<%Next%>]
<%End If %>};
					
function onlyUnique(value, index, self) { 
	return self.indexOf(value) === index;
}

function data_description(view){
	Number.prototype.format = function(n, x, s, c) {
		var re = '\\d(?=(\\d{' + (x || 3) + '})+' + (n > 0 ? '\\D' : '$') + ')',
			num = this.toFixed(Math.max(0, ~~n));

		return (c ? num.replace('.', c) : num).replace(new RegExp(re, 'g'), '$&' + (s || ','));
	};
	var group = document.getElementById("category").value;
	var stock = document.getElementById("stock").value;
	var jml_data 	= result_array.group_code.length;
	var group_filter = result_array.group_code.filter(onlyUnique); 
	var qty_filter 	= result_array.type.filter(onlyUnique);
	var group_name_filter = result_array.group_name.filter(onlyUnique); 
	var rowspanno 	= [];
	var rowspanid 	= [];
	var rowspanname = [];
	var rspn = 0;
	
	var counts = {};

	var split = view.split(":");
	var view = split[0];
	var qty = split[1];
	option_category = "<option value='ALL'> -- ALL Kategori -- </option>";
	for (f = 0; f < group_filter.length; f++) {
		if(group_filter[f] == group){
			option_category += "<option value='"+group_filter[f]+"' selected>"+group_filter[f]+"&nbsp;:&nbsp;"+group_name_filter[f]+"</option>";
		}else{
		    //result_array.group_name[result_array.group_code.indexOf(view)]
			option_category += "<option value='"+group_filter[f]+"'>"+group_filter[f]+"&nbsp;:&nbsp;"+group_name_filter[f]+"</option>";
		}
		//style_qty += ".qty_"+rowspanid[c]+"{display:none;}";
	}
	option_qtycategory = "<option value='"+view+":ALL'> -- ALL Stok Kategori -- </option>";
	for (f = 0; f < qty_filter.length; f++) {
		if(qty_filter[f].replace(/ /g, "_") == qty){
			option_qtycategory += "<option value='"+view+":"+qty_filter[f].replace(/ /g, "_")+"' selected>"+qty_filter[f]+"</option>";
		}else{
			option_qtycategory += "<option value='"+view+":"+qty_filter[f].replace(/ /g, "_")+"'>"+qty_filter[f]+"</option>";
		}	
		//style_qty += ".qty_"+rowspanid[c]+"{display:none;}";
	}
	document.getElementById("category").innerHTML = option_category;
	document.getElementById("stock").innerHTML = option_qtycategory;
	if (jml_data > 0 ){	
		var no = 1;
		var total = 0;
		var html_table = "";
		for (i = 0; i < jml_data; i++) {
			next = Number(i + 1);
			prev = Number(i - 1);
			
			if(view == "ALL"){	
					//no = next;
					if(qty && qty != "ALL" && result_array.type[i].replace(/ /g, "_") != qty){
					
					}else{
						html_table += '<tr>';
							html_table += '<td align="center">'+no+'</td>';
						if(no == 1){	
							rspn = 1;
							rowspanid.push(result_array.group_code[i]);
							html_table += '<td valign="top" id="row'+result_array.group_code[i]+'" align="center">'+result_array.group_name[i]+'</td>';
						}else if(result_array.group_code[i] != code_group){
							//alert(code_group+"/"+rspn);
							rowspanid.push(result_array.group_code[i]);
							html_table += '<td valign="top" id="row'+result_array.group_code[i]+'" align="center">'+result_array.group_name[i]+'</td>';
							rowspanno.push(rspn);
							rspn = 1;
							//rowspanid.push(result_array.group_code[i]);
						}else{
							rspn = rspn + 1;
							html_table += '';
						}
						
						    html_table += '<td align="left">'+result_array.kode_stock[i]+'</td>';
							html_table += '<td align="left">'+result_array.nama_stock[i]+'</td>';
							html_table += '<td align="right">'+result_array.soh[i]+'</td>';
						if(result_array.type[i] == "FAST MOVING"){
							html_table += '<td align="center"><img src="/images/check.gif"/></td>';
						}else{
							html_table += '<td align="center"></td>';
						}
						if(result_array.type[i] == "SLOW MOVING"){
							html_table += '<td align="center"><img src="/images/check.gif"/></td>';
						}else{
							html_table += '<td align="center"></td>';
						}
						if(result_array.type[i] == "DEAD STOCK"){
							html_table += '<td align="center"><img src="/images/check.gif"/></td>';
						}else{
							html_table += '<td align="center"></td>';
						}
						html_table += '<td align="right">'+result_array.amtperunit[i]+'</td>';
						html_table += '<td align="right">'+result_array.amttotFormat[i] +'</td>';
						html_table += '</tr>';
						total = parseInt(total) + parseInt(result_array.amttot[i]);
						code_group = result_array.group_code[i];
						no = no + 1;
					}	
					
				}else{
					if(qty && qty != "ALL"){
						if(result_array.group_code[i] == view && result_array.type[i].replace(/ /g, "_") == qty){
							html_table += '<tr>';
								html_table += '<td align="center">'+no+'</td>';
							if(no == 1){	
								rspn = 1;
								html_table += '<td valign="top" id="row'+result_array.group_code[i]+'" align="center">'+result_array.group_name[i]+'</td>';
							}else{
								rspn = rspn + 1;
								html_table += '';
							}
								html_table += '<td align="left">'+result_array.kode_stock[i]+'</td>';
								html_table += '<td align="left">'+result_array.nama_stock[i]+'</td>';
								html_table += '<td align="right">'+result_array.soh[i]+'</td>';
							if(result_array.type[i] == "FAST MOVING"){
								html_table += '<td align="center"><img src="/images/check.gif"/></td>';
							}else{
								html_table += '<td align="center"></td>';
							}
							if(result_array.type[i] == "SLOW MOVING"){
								html_table += '<td align="center"><img src="/images/check.gif"/></td>';
							}else{
								html_table += '<td align="center"></td>';
							}
							if(result_array.type[i] == "DEAD STOCK"){
								html_table += '<td align="center"><img src="/images/check.gif"/></td>';
							}else{
								html_table += '<td align="center"></td>';
							}
							html_table += '<td align="right">'+result_array.amtperunit[i]+'</td>';
							html_table += '<td align="right">'+result_array.amttotFormat[i]+'</td>';
							html_table += '</tr>';
							total = parseInt(total) + parseInt(result_array.amttot[i]);
							no = no + 1;
							code_group = result_array.group_code[i];
						}
					}else{
							if(result_array.group_code[i] == view){
								html_table += '<tr>';
									html_table += '<td align="center">'+no+'</td>';
								if(i == 0){	
									rspn = 1;
									html_table += '<td valign="top" id="row'+result_array.group_code[i]+'" align="center">'+result_array.group_name[i]+'</td>';
								}else if(result_array.group_code[i] != result_array.group_code[prev]){
									rspn = 1;
									html_table += '<td valign="top" id="row'+result_array.group_code[i]+'" align="center">'+result_array.group_name[i]+'</td>';
								}else{
									rspn = rspn + 1;
									html_table += '';
								}
									html_table += '<td align="left">'+result_array.kode_stock[i]+'</td>';
									html_table += '<td align="left">'+result_array.nama_stock[i]+'</td>';
									html_table += '<td align="right">'+result_array.soh[i]+'</td>';
								if(result_array.type[i] == "FAST MOVING"){
									html_table += '<td align="center"><img src="/images/check.gif"/></td>';
								}else{
									html_table += '<td align="center"></td>';
								}
								if(result_array.type[i] == "SLOW MOVING"){
									html_table += '<td align="center"><img src="/images/check.gif"/></td>';
								}else{
									html_table += '<td align="center"></td>';
								}
								if(result_array.type[i] == "DEAD STOCK"){
									html_table += '<td align="center"><img src="/images/check.gif"/></td>';
								}else{
									html_table += '<td align="center"></td>';
								}
								html_table += '<td align="right">'+result_array.amtperunit[i]+'</td>';
								html_table += '<td align="right">'+result_array.amttotFormat[i]+'</td>';
								html_table += '</tr>';
								total = parseInt(total) + parseInt(result_array.amttot[i]);
								no = no + 1;
								code_group = result_array.group_code[i];
							}
						}
						
					}
					
				}
	
				//alert(total);
		html_table += '<tr><td colspan="9" align="right"  >TOTAL</td><td align="right"  >'+total.format(0, 3, ',', '.')+'</td></tr>';		
		document.getElementById("data_table").innerHTML =  html_table;
		if(view == "ALL"){	
			for (c = 0; c < rowspanid.length; c++) {
				if(c == (rowspanid.length -1)){
					document.getElementById("row"+rowspanid[c]).rowSpan = rspn;
				}else{
				//alert(rowspanno[c]+","+rowspanid[c])
					document.getElementById("row"+rowspanid[c]).rowSpan = rowspanno[c];
				}
			}
		}else{
			document.getElementById("row"+code_group).rowSpan = no;
		}
	}else{
		html_table = '<tr><td colspan="9">NO DATA AVAILABLE</td></tr>';
		document.getElementById("data_table").innerHTML =  html_table;
	}
}
</script>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Category Stock</title>
<link rel="stylesheet" href="/css/font/font-awesome.min.css" type="text/css">
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
var stop = document.getElementById('stop');
xAddEventListener(window, 'load',header_fixed, false);
function header_fixed(){
  new xTableHeaderFixed('xthf-site', window);
}
function header_fixed_refres(){
	xTableHeaderFixed('xthf-site', window);
}

</script>

<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>


<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<table id="form-table" class="centered-margin">	
					<tr>
						<td><font face="verdana" size="1">YEAR</font></td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, i, MM
					mon = Month(Session("v_Costbook"))
					If len(v_txtMonth)=0 then v_txtMonth=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = Yr - 2
					MaxYr = Yr  + 2
					If len(v_txtYear)= 0 then v_txtYear = Yr
				%>
						<td width="4">
							<select name="txtYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
							</select>
						</td>
						<td><font face="verdana" size="1">MONTH</font></td>
						<td width="2">
							<select name="txtMonth">
				<%
					For i = 1 to 12
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
				<%
					Next
				%>
							</select>
						</td>
					<td><font face="verdana" size="1">FCBA</font></td>
					<td><select name="txtFcba">
					<%
								Dim rsFcba
								Set rsFcba = Server.CreateObject("ADODB.Recordset")
								rsFcba.CursorLocation = adUseClient
								strsql =    " SELECT fccode fcba from businessunit " & _
											" where " & _
											" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
								
								rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsFcba.eof then
									DO while not rsFcba.eof
										if v_txtFcba = rsFcba("fcba") then
										%>			
											<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
										<% Else %>	
											<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
										<% End If
										
										rsFcba.movenext
									loop
								End if
								rsFcba.Close
								set rsFcba = Nothing
								%></select></td>
					<td><input type="submit" value="View" /></td>
					</tr>	
				</table>
			</div>
			</form>
			<div class="subheader">
					<table class="centered-margin" style="width:40%;border:0px;">	
						<tr>
							<td><font face="verdana" size="1">Category</font></td>
							<td>
								<select id="category" onchange="javascript:data_description(this.value);">
									<!-- Kategori Option -->	
								</select>
							</td>
							<td><font face="verdana" size="1">Stock Category</font></td>
							<td>
								<select id="stock" onchange="javascript:data_description(this.value);">
									<!-- Stock Category Option -->	
								</select>
							</td>
						</tr>	
					</table>
					
					
				</div>		
</div>

<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
			
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>Category Stock</h3>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>

				
			</div>
</th>
  </tr>
  
	<tr>
			<th rowspan='2' style="background-color:#00A300;color:#fff;">NO</th>
			<th rowspan='2' style="background-color:#00A300;color:#fff;">Group Descrption</th>
			<th rowspan='2' style="background-color:#00A300;color:#fff;">Kode Stock</th>
			<th rowspan='2' style="background-color:#00A300;color:#fff;">Nama Stock</th>
			<th rowspan='2' style="background-color:#00A300;color:#fff;">On hand </th>
			<th colspan='3' style="background-color:#00A300;color:#fff;">Stock Category (Qty)</th>
			<th rowspan='2' style="background-color:#00A300;color:#fff;">Amount Per Unit</th>
			<th rowspan='2' style="background-color:#00A300;color:#fff;">Amount Total</th>
	</tr>
	<tr>	
			<th style="background-color:#00A300;color:#fff;">Fast Moving</th>
			<th style="background-color:#00A300;color:#fff;">Slow Moving</th>
			<th style="background-color:#00A300;color:#fff;">Dead Stock</th>
		</tr>
</thead>
<tbody id="data_table">

	
</tbody>
</table>
<script>
	data_description("ALL");
</script>
<style id="category_style"><!-- Filtering category --></style>
<style id="stock_style"><!-- Filtering Stock --></style>

</div>
</body>
</html>