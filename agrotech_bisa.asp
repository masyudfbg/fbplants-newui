<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/mm.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
	<title>Agrotechnical BISA - iPlaS</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
</script>

<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">

<style>
.close-button {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 30px;
  height: 30px;
  position: relative;
  padding: 0 1px 0 0;
  border: none;
  -webkit-border-radius: 1em;
  border-radius: 1em;
  font: normal 8em/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #6d6d6d;
  cursor:pointer;
}

.close-button:hover {background: #f72c2c;}

.close-button::before {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(45deg)   ;
  transform: rotateZ(45deg)   ;
}

.close-button::after {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(-45deg)   ;
  transform: rotateZ(-45deg)   ;
}
</style>

</head>
<body>
	<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
<%
'--Awal--
Dim stringduedate : stringduedate = 0

Dim IsOverDuedate, varRemainDueDate, varDueDate, CompleteStringError, varErrorDueDate, stringremainqty, varRemainQty, IsOverRemainQty, varErrorRemainQty, varEditedQty

Dim qtyplanedit : qtyplanedit = Request("qtyplanedit")
Dim unitplanedit : unitplanedit = Request("unitplanedit")
Dim outputplanedit : outputplanedit = Request("outputplanedit")
Dim qtyfinaledit : qtyfinaledit = Request("qtyfinaledit")
Dim unitfinaledit : unitfinaledit = Request("unitfinaledit")
Dim outputfinaledit : outputfinaledit = Request("outputfinaledit")
Dim txtagreementlineno : txtagreementlineno = Request("txtagreementlineno")

Dim txtGRNCode : txtGRNCode = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim txtLocationCode : txtLocationCode = ""

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

txtyear = Request("txtyear")
txtmonth = Request("txtmonth")
txtblok = Request("txtblok")
txtstatus = Request("txtstatus")
fddate = Request("fddate")

if "" & txtyear = "" then txtyear = year(Session("v_Costbook"))
if "" & txtmonth = "" then txtmonth = Month(Session("v_Costbook"))
if "" & fddate = "" then fddate = txtmonth & "/1/" & txtyear

MinYr = txtyear -2
MaxYr = txtyear +2

arr_field_eof = true
oSql = " SELECT fccode" & _
       "        , field_api.getvalid_name(fccode, fcba, TO_DATE('" & fddate & "', 'mm/dd/yyyy')) fcname" & _
       "        , field_api.getvalid_status(fccode, fcba, TO_DATE('" & fddate & "', 'mm/dd/yyyy')) status" & _
       " FROM field" & _
	   " WHERE fcba = '" & Session("fcBA") & "'" & _
	   "       AND field_api.getvalid_activation(fccode, fcba, TO_DATE('" & fddate & "', 'mm/dd/yyyy')) = 'Y'"

set oRs = server.CreateObject("Adodb.recordset")
oRs.Open oSql, connstr
if Not oRs.Eof then
	arr_field = oRs.getrows()
	arr_field_eof = false
End if
oRs.Close
Set oRs = Nothing


txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Header" then
		arr_insertdata_eof = true
		oSql = " SELECT fcba, accountingyear, accountingperiod, fieldcode, fccode" & _
               " FROM trans_bisa_agrotech_dtl" & _
               " WHERE fcba = '" & Session("fcBA") & "'" & _
               "       AND accountingyear = '" & txtyear & "'" & _
               "       AND accountingperiod = '" & txtmonth & "'" & _
               "       AND fieldcode = '" & txtblok & "'"

		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		if Not oRs.Eof then
			arr_insertdata = oRs.getrows()
			arr_insertdata_eof = false
		End if
		oRs.Close
		Set oRs = Nothing
		
		if not arr_insertdata_eof then
			for nn=0 to ubound(arr_insertdata,2)
				q = " DELETE FROM trans_bisa_agrotech_dtl" & _
                    " WHERE fcba = '" & Session("fcBA") & "'" & _
                    "       AND accountingyear = '" & txtyear & "'" & _
                    "       AND accountingperiod = '" & txtmonth & "'" & _
                    "       AND fieldcode = '" & txtblok & "'" & _
					"       AND fccode = '" & arr_insertdata(4, nn) & "'"
				connStr.execute(q)
			next
		end if
		
		q = " DELETE FROM trans_bisa_agrotech_hdr" & _
			" WHERE fcba = '" & Session("fcBA") & "'" & _
			"       AND accountingyear = '" & txtyear & "'" & _
			"       AND accountingperiod = '" & txtmonth & "'" & _
			"       AND fieldcode = '" & txtblok & "'"
		connStr.execute(q)
		
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
		arr_insertdata_eof = true
		oSql = " SELECT fcba, accountingyear, accountingperiod, fieldcode, fccode" & _
               " FROM trans_bisa_agrotech_dtl" & _
               " WHERE fcba = '" & Session("fcBA") & "'" & _
               "       AND accountingyear = '" & txtyear & "'" & _
               "       AND accountingperiod = '" & txtmonth & "'" & _
               "       AND fieldcode = '" & txtblok & "'"

		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		if Not oRs.Eof then
			arr_insertdata = oRs.getrows()
			arr_insertdata_eof = false
		End if
		oRs.Close
		Set oRs = Nothing
		
		if not arr_insertdata_eof then
			for nn=0 to ubound(arr_insertdata,2)
				q = " DELETE FROM trans_bisa_agrotech_dtl" & _
                    " WHERE fcba = '" & Session("fcBA") & "'" & _
                    "       AND accountingyear = '" & txtyear & "'" & _
                    "       AND accountingperiod = '" & txtmonth & "'" & _
                    "       AND fieldcode = '" & txtblok & "'" & _
					"       AND fccode = '" & arr_insertdata(4, nn) & "'"
				connStr.execute(q)
			next
		end if
		
		arr_insertdata_eof = true
		oSql = " SELECT * FROM (" & _
		       " SELECT DISTINCT fccode" & _
               " FROM setup_bisa_agrotech_tab ch" & _
               " CONNECT BY PRIOR parent_fccode = fccode" & _
               " START WITH status = '" & txtstatus & "'" & _
			   " ) ORDER BY SUBSTR(fccode,1,2), fccode"
		
		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		if Not oRs.Eof then
			arr_insertdata = oRs.getrows()
			arr_insertdata_eof = false
		End if
		oRs.Close
		Set oRs = Nothing
		
		if not arr_insertdata_eof then
			for nn=0 to ubound(arr_insertdata,2)
				if len("" & arr_insertdata(0,nn)) = 4 then
					if cdbl(to_nol(request("cb_"&arr_insertdata(0,nn)))) = 0 then
						q = " INSERT INTO trans_bisa_agrotech_dtl ( accountingyear" & _
							"                                       , accountingperiod" & _
							"                                       , fieldcode" & _
							"                                       , fccode" & _
							"                                       , conformation" & _
							"                                       , fcentry" & _
							"                                       , fcedit" & _
							"                                       , fcip" & _
							"                                       , fcba" & _
							"                                       , lastupdate" & _
							"                                       , lasttime)" & _
							" VALUES ( '" & txtyear & "'" & _
							"          , '" & txtmonth & "'" & _
							"          , '" & txtblok & "'" & _
							"          , '" & arr_insertdata(0,nn) & "'" & _
							"          , '1'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & Session("fcBA") & "'" & _
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE, 'hh24:mi') )"
						connStr.execute(q)
					end if
				end if
			next
		end if
		
		q = " BEGIN rpt_bisa_agrotech_api.insertdata( '" & txtyear & "'" & _
		    "                                         , '" & txtmonth & "'" & _
		    "                                         , '" & txtblok & "'" & _
		    "                                         , '" & Session("CodeName") & "'" & _
		    "                                         , '" & Session("CodeName") & "'" & _
		    "                                         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
		    "                                         , '" & Session("fcBA") & "' ); END;"
		connStr.execute(q)
		
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	
		cmdView	= "True"
    else
	
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")		    
End if

val_IsTrue = "N"
oSql = " SELECT rpt_bisa_agrotech_api.bolehinput('" & txtyear & "'||LPAD('" & txtmonth & "',2,0), '" & Session("fcBA") & "') istrue" & _
       " FROM DUAL"
			
set oRs = server.CreateObject("Adodb.recordset")
oRs.Open oSql, connstr
If Not oRs.Eof then
	val_IsTrue = oRs("istrue")
Else
	val_IsTrue = "N"
End if
oRs.Close
Set oRs = Nothing


If Len(Trim(CompleteStringError)) > 0 Then
	response.Write  "<div class = ""error"">" & CompleteStringError & "</div>"
End If
if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if
 %>
<form name="frmEntry" method="post" action="agrotech_bisa.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Year/ Tahun</td>
			<td align="left">
				<select name="txtyear">
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(txtyear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Month/ Bulan</td>
			<td align="left">
				<select name="txtmonth">
				<%For i = 1 to 12%>
					<Option Value="<%=i%>" <%If cint(i) = cint(txtmonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Field/Blok</td>
			<td align="left">
				<select id="txtblok" name="txtblok" style="width:50%;" class="chosen-select">
					<option value=""></option>
				<%
					if not arr_field_eof then
						for xi=0 to ubound(arr_field,2)
							if cstr(txtblok) = cstr(arr_field(0, xi)) then
								txtstatus = cstr(arr_field(2, xi))
								response.write "<option value='"&arr_field(0, xi)&"' selected>"&arr_field(0, xi) & " : " & arr_field(1, xi)&"</option>"
							else
								response.write "<option value='"&arr_field(0, xi)&"'>"&arr_field(0, xi) & " : " & arr_field(1, xi)&"</option>"
							end if
						next
					end if
				%>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;
				<%if val_IsTrue = "Y" then%>
				<a href="agrotech_bisa.asp?txtyear=<%=txtyear%>&txtmonth=<%=txtmonth%>&txtblok=<%=txtblok%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0">
				</a>
				<%end if%>
			</td>
		</tr>
	</table>  
</form>
<%
If cmdView = "True" and "" & txtblok <> "" Then	
	q = " SELECT COUNT(1) result FROM trans_bisa_agrotech_hdr" & _
		" WHERE fcba = '" & Session("fcBA") & "'" & _
		"       AND accountingyear = '" & txtyear & "'" & _
		"       AND accountingperiod = '" & txtmonth & "'" & _
		"       AND fieldcode = '" & txtblok & "'"

	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open q, connstr
   
	if not oRs.eof then
		has_header = cdbl(oRs("result"))
	end if

	arr_data_eof = true
	q = " SELECT data.fccode, master.fcname, SUM(data.conformation) conformation, master.priority FROM (" & _
	    " SELECT fccode, 0 conformation" & _
		" FROM (" & _
		"     SELECT DISTINCT fccode, fcname" & _
		"     FROM setup_bisa_agrotech_tab ch" & _
		"     CONNECT BY PRIOR parent_fccode = fccode" & _
		"     START WITH status = '" & txtstatus & "'" & _
		" ) UNION ALL" & _
		" SELECT fccode, conformation FROM trans_bisa_agrotech_dtl" & _
		" WHERE fcba = '" & Session("fcBA") & "'" & _
		"       AND accountingyear = '" & txtyear & "'" & _
		"       AND accountingperiod = '" & txtmonth & "'" & _
		"       AND fieldcode = '" & txtblok & "'" & _
		"       AND conformation > 0" & _
		" ) data LEFT JOIN setup_bisa_agrotech_tab master ON master.fccode = data.fccode" & _
		" GROUP BY data.fccode, master.fcname, master.priority" & _
		" ORDER BY SUBSTR(data.fccode, 1,2), data.fccode"

	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open q, connstr
   
	if not oRs.eof then
		arr_data = oRs.getrows()
		arr_data_eof = false
	end if

   
%>
	<form action="agrotech_bisa.asp" method="post" name="eForm">
	<table class="entry" cellpadding="1" cellspacing="1" width="100%" style="font-size:12px;">
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		   <td align="center">Code</td>
		   <td align="center">Description</td>
		   <td align="center">Conformation</td>
		</tr>  

<%
	response.write "<tr style='background-color:#FF851B;'>"
	response.write "    <td colspan='2'>&nbsp;</td>"
	response.write "    <td align='center'><input type='checkbox' name='select-all' id='select-all'/> Check All/ Uncheck All</td>"
	response.write "<t/r>"
	if not arr_data_eof then
		no = 1
		for xi=0 to ubound(arr_data,2)
		
			if cdbl(to_nol(arr_data(3,xi))) = 1 then
				cColor = "background-color:#FF4136;color:#FFFFFF;"
			elseif cdbl(to_nol(arr_data(3,xi))) = 2 then
				cColor = "background-color:#FFDC00;color:#111111;"
			elseif cdbl(to_nol(arr_data(3,xi))) = 3 then
				cColor = "background-color:#2ECC40;color:#111111;"
			end if
			
			if len(cstr(arr_data(0,xi))) = 4 then
				add_spasi = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				add_colspan = ""
			end if
			if len(cstr(arr_data(0,xi))) = 2 then
				add_style = "font-weight:bold;background-color:#f0f0b0;color:#000000;"
				add_colspan = "colspan='2'"
				
				if xi <> 0 then
					response.write "<tr><td colspan='3' style='background-color:#FFFFFF;'></td></tr>"
				end if
			end if
			response.write "<tr style='"&cColor&"'>"
			response.write "	<td align='right' style='"&add_style&"'>"&arr_data(0,xi)&"</td>"
			response.write "	<td align='left' style='"&add_style&"' "&add_colspan&">"&add_spasi&arr_data(1,xi)&"</td>"
			if len(cstr(arr_data(0,xi))) = 4 then
				propchecked = ""
				if cdbl(to_nol(arr_data(2,xi))) = 0 then
					if has_header = 1 then
						propchecked = "checked"
					end if
				else
					propchecked = ""
				end if
				response.write "	<td align='center' style='"&add_style&"'>"
				response.write "	    <input type='checkbox' name='cb_"&arr_data(0,xi)&"' id='cb_"&arr_data(0,xi)&"' class='item' value='"&cdbl(to_nol(arr_data(2,xi)))&"' "&propchecked&" onclick='handleClick(this);'>"
				response.write "	</td>"
			end if
			response.write "</tr>"
			add_spasi = ""
			add_style = ""
			add_colspan = ""
			no = no +1
		next
	else
		response.Write("<tr><td colspan='3' align='left'>Data is empty</td></tr>")	  
	end if 
%>
		<tr style='background-color:#CCC;'>
			<td colspan="3" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="txtyear" value="<%=txtyear%>">
			<input type="hidden" name="txtmonth" value="<%=txtmonth%>">
			<input type="hidden" name="txtblok" value="<%=txtblok%>">
			<input type="hidden" name="txtstatus" value="<%=txtstatus%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<%If val_IsTrue = "Y" Then%>
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit"  name="Add" value="Update">
			<%Else%>
				<input type="hidden" name="txtJenis" value="New">
				<input Type="Submit"  name="Add" value="Submit">
			<%End if%>
			<input Type="Reset"  name="Reset" value="Reset">
			<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
			<%End if%>
		</tr>
	</table>
</div>
</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
<!--#include file="bottom.jsp"-->	
	<script src="plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
	
	<script type="text/javascript">
		function handleClick(cb){
			if(cb.checked == 1){
				cb.value = 1;
			} else {
				cb.value = 0;
			}
        }
		
	
	</script>
	<script type="text/javascript">	
		$('#select-all').click( function () {
			$('.item').prop('checked', this.checked);
			$('.item').val(this.checked ? 1 : 0);
		})
	</script>
</body>
</html>
