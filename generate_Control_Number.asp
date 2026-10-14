<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
  Dim textReportName : textReportName = "Generate Control Number"
  Dim ScriptName : ScriptName = "generate_Control_Number.asp"

  ScriptName = "p_" & ScriptName
%>
<html>
<head>
<title>ePlaS - <%=PCase(textReportName)%></title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<link rel="stylesheet" href="css/ns-window.css">
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script>

function ActionFrom(command){
	var year = document.getElementById("txtYear").value;
	var month = document.getElementById("txtMonth").value;
	
	if(command == "select"){
		CallMethod("P_Generate_Control_Number.asp","?command="+command+"&txtYear="+year+"&txtMonth="+month, onSuccess);
	}else if(command == "action"){
		CallMethod("P_Generate_Control_Number.asp","?command="+command+"&txtYear="+year+"&txtMonth="+month, onSuccess);
	}
};	

//Function For AJAX author Atwal Arifin
function CallMethod(url,param, successCallback) {
                //show loading... image
				response = "";
				response += "<tr>";
				response += "<td colspan='7' align='center'><img src='../IMAGES/loader2.gif'></td>";
				response += "</tr>";
				$("#ItemFromDataAjax").html(response);
                $.ajax({
                    type: 'GET',
                    url: url+param,
                    contentType: 'application/json;',
                    dataType: 'json',
					encode: true,
                    success: successCallback,
                    error: function(xhr, textStatus, errorThrown) {
                        console.log('error');
                    }
                });
            }
$( document ).click(function() {
  $( "#nswindow" ).effect( "shake" );
});			
function close(id,button){
var h = $( window ).height();
var w = $( window ).width();
	 $( "#nswindow" ).animate({
		height: 50,
		width: (w*20)/100,
		left: (w-((w*20)/100))/2+"px",
		top: -1000+"px"
	 }, 100, function() {
	    var elem = document.getElementById(id);
		elem.parentElement.removeChild(elem);
	 });
}
function onSuccess(data){
var new_window = '<div id="form_window" class="nswindowContainer" style="display: table-row; z-index: 1010; background-color: rgba(0, 0, 0, 0.4);"><div id="nswindow" class="yellow" style=" opacity: 1;"><div id="ctrlbar" class="controlbar"><div class="title" style="margin-top: 4.5px;"></div><a id="ButtonCloseWindow" class="close" href=javascript:close("form_window","ButtonCloseWindow") title="close"></a></div><div id="nswindowContent" style="background:#f1f1f1;top: 26px;"><div id="WARP_input" style="text-align:left;padding:20px;"></div></div></div></div>';
var year = document.getElementById("txtYear").value;
var month = document.getElementById("txtMonth").value;
var h = $( window ).height();
var w = $( window ).width();
response = "";
   if(data){
		if(data.response == "successfully" || data.response == "selected"){
			for (i = 0; i < data.item.length; i++) {
				response += "<tr>";
				response += "<td align='left'>"+data.item[i].modulename+"</td>";
				response += "<td>"+data.item[i].Accountingyear+"</td>";
				response += "<td>"+data.item[i].AccountingPeriod+"</td>";
				response += "<td>"+data.item[i].Prefix+"</td>";
				response += "<td>"+data.item[i].Suffix+"</td>";
				response += "<td>"+data.item[i].CurrentNumber+"</td>";
				response += "<td>"+data.item[i].Lengthpad+"</td>";
				response += "</tr>";
			}
			if(data.response == "successfully"){
				alert("Success");
			}
		}else if(data.response == "error"){
			error_html = "";
			if (data.queryname == "Checking"){
				response += "<tr>";
				response += "<td colspan='7' align='left'>No Data</td>";
				response += "</tr>";
				//$("HTML").after(new_window);
				$("#ItemFromDataAjax").after(new_window);
				$( "#nswindow" ).animate({
					height: 400,
					width: (w*60)/100,
					left: (w-((w*60)/100))/2+"px",
					top: "25px"
				}, 70, function() {
					$("#ctrlbar .title").html("ERROR");
					error_html += "<h1>Query Name : "+ data.queryname +"</h1>";
					error_html += "<h2> Command : "+ data.item[0].command +"</h2>";
					for (i = 0; i < data.item.length; i++) {
						if(data.item[i].description == 0 ){
							error_html += "<p>"+ data.item[i].query +" : " + data.item[i].description + "<br/>";
							error_html += "" + data.item[i].query + " ditahun "+year+" dan bulan "+month+", Belum dibuat.";
							error_html += "</p>";
						}
					}
					$("#WARP_input").html(error_html);
				});
				
			}else{
				response += "<tr>";
				response += "<td colspan='7' align='left'>No Data</td>";
				response += "</tr>";
				if(data.item[0].command != "select"){
				$("#ItemFromDataAjax").after(new_window);
					$( "#nswindow" ).animate({
						height: 400,
						width: (w*60)/100,
						left: (w-((w*60)/100))/2+"px",
						top: "25px"
					}, 70, function() {
						$("#ctrlbar .title").html("ERROR");
						error_html += "<h1>Query Name : "+ data.queryname +"</h1>";
						error_html += "<h2> Command : "+ data.item[0].command +"</h2>";
						for (i = 0; i < data.item.length; i++) {
								error_html += "<p>"+ data.item[i].query +" : " + data.item[i].description + "<br/>";
								error_html += "</p>";
							}
							
						$("#WARP_input").html(error_html);
						
					});
				}
			}
		}
	}else{
		response += "<tr>";
		response += "<td colspan='7' align='left'>"+data.response+"</td>";
		response += "</tr>";
	}
	
   $("#ItemFromDataAjax").html(response);
   
}
//alert("success");
</script>

	</head>

	<body>
		<!--#include file="top.jsp"-->
		<%
		Dim DATE_FROM : DATE_FROM = Request("DATE_FROM")
		Dim DATE_TO : DATE_TO = Request("DATE_TO")
		
		If Len(Trim(DATE_FROM)) = 0 Then
			DATE_FROM = textDate(Date)
		End if
		
		If Len(Trim(DATE_TO)) = 0 Then
			DATE_TO = textDate(Date)
		End if
		
		Dim txtTable, txtModule
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
		<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
		<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
				<div class="judul">
					Module : Generate Control Number / <%=PCase(textReportName)%>
				</div>
				<div class="box_both" style="overflow:scroll;">
				<form  name="frmEntry">
					<table class = "data">
						<thead>
						</thead>
						<tbody>
							<tr>
								<td align="left" width="150">Tahun</td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, i, MM, v_txtMonth, v_txtYear
					mon = Month(Session("v_Costbook"))
					If len(v_txtMonth)= 0 then v_txtMonth = mon

					Yr = Year(Session("v_Costbook"))
					If len(v_txtYear) = 0 then v_txtYear = Yr
					MinYr = Yr - 2
					MaxYr = Yr + 2
					
				%>
						<td align="left">
							<select id="txtYear" name="txtYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
							</select>
						</td>
							</tr>
							<tr>
								<td align="left">Bulan</td>
						<td align="left">
							<select id="txtMonth" name="txtMonth" onchange="javacript:ActionFrom('select');">
				<%
					For i = 1 to 15
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
				<%
					Next
				%>
							</select>
						</td>
							</tr>
							<tr>
								<td align="left"><a class="button_menu_top" type="Button" onclick="ActionFrom('action');">Process</a></td>
							</tr>
						</tbody>
					</table>
				</form>
				<table class="data xthf-site">
				<thead>
					<tr>
						<th>Modul Name</th><th>Acounting Year</th><th>Accounting Period</th><th>Prefix</th><th>Suffix</th><th>Current Number</th><th>Length Pad</th>
					</tr>
				<thead>
				<tbody id="ItemFromDataAjax">
					
				</tbody>
					</table>
				</div>
		</div>
		<!--#include file="bottom.jsp"-->
	</body>
</html>

