<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim textReportName : textReportName = "Fingerprint Validation"
Dim ScriptName : ScriptName = "fingerprint_validation.asp?txtJenis=edit"

Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim v_txtFcba 
Dim INDAY  
Dim onoffswitch  
Dim Ischecked 
Dim TrueOrFalse 

Dim orecordset, sqlquery
Dim rowdata : rowdata = 0
Dim maxlineno : maxlineno = 0

Dim InsertDataFp, InsertDataFpHist, UpdateDataFp

v_txtFcba   = Request("txtFcba")
INDAY       = Request("INDAY")
onoffswitch = Request("onoffswitch")
Ischecked   = Request("Ischecked")

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then v_txtFcba = varDefaultBa
If Len(Trim(INDAY)) = 0 Then INDAY = textDate(Date)


if Request("txtJenis")= "edit"  then

	if onoffswitch = "on" then 
		Ischecked = "checked"
		TrueOrFalse = "TRUE"
	else
		Ischecked = ""
		TrueOrFalse = "FALSE"
	end if
	
	sqlquery = " select count(1) rowdata" & _
	           " from fingervalidation" & _
	           " where fcba = '" & v_txtFcba & "'"
				
	Set orecordset = Server.CreateObject("ADODB.Recordset")
	orecordset.Open sqlquery, connstr
	
	if not orecordset.EOF then
		rowdata = 1
	else
		rowdata = 0
	end if
	orecordset.close
	Set orecordset = Nothing
	
'response.Write ("<script>alert('"&rowdata&"');</script>")
	
	if rowdata = 0 then
		
		InsertDataFp = " insert into fingervalidation ( fcba" & _
					   "                                , fddate" & _
					   "                                , isvalidation" & _
					   "                                , fcentry" & _
					   "                                , fcedit" & _
					   "                                , fcip" & _
					   "                                , lastupdate" & _
					   "                                , lasttime)" & _
					   " values ( '" & v_txtFcba & "'" & _
					   "          , to_date('" & INDAY & "', 'mm/dd/yyyy')" & _
					   "          , '" & TrueOrFalse & "'" & _
					   "          , '" & Session("CodeName") & "'" & _
					   "          , '" & Session("CodeName") & "'" & _
					   "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					   "          , sysdate" & _
					   "          , to_char(sysdate, 'hh24:mi'))"
		'response.write InsertDataFp & "<br/>"
		connStr.execute(InsertDataFp)
		
		InsertDataFpHist = " insert into fingervalidationhist ( lineno" & _
					       "                                    , fcba" & _
					       "                                    , fddate" & _
					       "                                    , isvalidation" & _
					       "                                    , fcentry" & _
					       "                                    , fcedit" & _
					       "                                    , fcip" & _
					       "                                    , lastupdate" & _
					       "                                    , lasttime)" & _
					       " values ( 1" & _
						   "          , '" & v_txtFcba & "'" & _
					       "          , to_date('" & INDAY & "', 'mm/dd/yyyy')" & _
					       "          , '" & TrueOrFalse & "'" & _
					       "          , '" & Session("CodeName") & "'" & _
					       "          , '" & Session("CodeName") & "'" & _
					       "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					       "          , sysdate" & _
					       "          , to_char(sysdate, 'hh24:mi'))"
		'response.write InsertDataFpHist & "<br/>"
		connStr.execute(InsertDataFpHist)
		
	else
		'response.write "row <> 0 <br/>"
		sqlquery = " select nvl(max(lineno), 0) maxlineno" & _
				   " from fingervalidationhist" & _
				   " where fcba = '" & v_txtFcba & "'"
					
		Set orecordset = Server.CreateObject("ADODB.Recordset")
		orecordset.Open sqlquery, connstr
		
		if not orecordset.EOF then
			maxlineno = cdbl(orecordset("maxlineno"))
		else
			maxlineno = 0
		end if
		orecordset.close
		Set orecordset = Nothing
		
		maxlineno = maxlineno + 1
		
		UpdateDataFp = " update fingervalidation" & _
                       " set fddate = to_date('" & INDAY & "', 'mm/dd/yyyy')" & _
                       "     , isvalidation = '" & TrueOrFalse & "'" & _
                       "     , fcentry = '" & Session("CodeName") & "'" & _
                       "     , fcedit = '" & Session("CodeName") & "'" & _
                       "     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
                       "     , lastupdate = sysdate" & _
                       "     , lasttime = to_char(sysdate, 'hh24:mi')" & _
                       " where fcba = '" & v_txtFcba & "'"
		
		connStr.execute(UpdateDataFp)
	
		InsertDataFpHist = " insert into fingervalidationhist ( lineno" & _
					       "                                    , fcba" & _
					       "                                    , fddate" & _
					       "                                    , isvalidation" & _
					       "                                    , fcentry" & _
					       "                                    , fcedit" & _
					       "                                    , fcip" & _
					       "                                    , lastupdate" & _
					       "                                    , lasttime)" & _
					       " values ( " & maxlineno & "" & _
						   "          , '" & v_txtFcba & "'" & _
					       "          , to_date('" & INDAY & "', 'mm/dd/yyyy')" & _
					       "          , '" & TrueOrFalse & "'" & _
					       "          , '" & Session("CodeName") & "'" & _
					       "          , '" & Session("CodeName") & "'" & _
					       "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					       "          , sysdate" & _
					       "          , to_char(sysdate, 'hh24:mi'))"
		'response.write InsertDataFpHist & "<br/>"
		connStr.execute(InsertDataFpHist)
		
		
end if

	
end if

		rowdata = 0
		INDAY = textDate(Date)
		Ischecked = ""
		TrueOrFalse = "FALSE"
		
sqlquery = " select isvalidation,to_char(fddate,'mm/dd/yyyy') fddate" & _
	           " from fingervalidation" & _
	           " where fcba = '" & v_txtFcba & "'"
				
	Set orecordset = Server.CreateObject("ADODB.Recordset")
	orecordset.Open sqlquery, connstr
	
	if not orecordset.EOF then
		rowdata = 1
		INDAY     = orecordset("fddate")
		Ischecked = orecordset("isvalidation")
		TrueOrFalse = orecordset("isvalidation")
		IF TrueOrFalse = "FALSE" then 
			Ischecked = ""
		Else	
			Ischecked = "Checked"
		END IF	
%>
<html>
	<head>
<title>ePlaS - <%=PCase(textReportName)%></title>

<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>

<style>
.onoffswitch {
    position: relative; width: 90px;
    -webkit-user-select:none; -moz-user-select:none; -ms-user-select: none;
}
.onoffswitch-checkbox {
    display: none;
}
.onoffswitch-label {
    display: block; overflow: hidden; cursor: pointer;
    border: 2px solid #999999; border-radius: 20px;
}
.onoffswitch-inner {
    display: block; width: 200%; margin-left: -100%;
    transition: margin 0.3s ease-in 0s;
}
.onoffswitch-inner:before, .onoffswitch-inner:after {
    display: block; float: left; width: 50%; height: 30px; padding: 0; line-height: 30px;
    font-size: 14px; color: white; font-family: Trebuchet, Arial, sans-serif; font-weight: bold;
    box-sizing: border-box;
}
.onoffswitch-inner:before {
    content: "ON";
    padding-left: 10px;
    background-color: #34A7C1; color: #FFFFFF;
	text-align: left;
}
.onoffswitch-inner:after {
    content: "OFF";
    padding-right: 10px;
    background-color: #EEEEEE; color: #999999;
    text-align: right;
}
.onoffswitch-switch {
    display: block; width: 18px; margin: 6px;
    background: #FFFFFF;
    position: absolute; top: 0; bottom: 0;
    right: 56px;
    border: 2px solid #999999; border-radius: 20px;
    transition: all 0.3s ease-in 0s; 
}
.onoffswitch-checkbox:checked + .onoffswitch-label .onoffswitch-inner {
    margin-left: 0;
}
.onoffswitch-checkbox:checked + .onoffswitch-label .onoffswitch-switch {
    right: 0px; 
}
</style>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261015a"></script>
<% end if %>
	</head>

	<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
		<!--#include file="top.jsp"-->
		<%
		Dim FCBA : FCBA = Request("FCBA")
		dim status : status = "No Status"
		
		Dim txtTable, txtModule
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
		<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
		<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
				<div class="judul">
					Module : Fingerprint Validation/ <%=PCase(textReportName)%>
				</div>
				<div class="box_both" style="overflow:scroll;">
					<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-check-square-o" aria-hidden="true"></i>Validasi Fingerprint</div><% end if %>					<form name="frmEntry" method="POST" action="<%=ScriptName%>">
					<table class = "data">
						<thead>
							
						</thead>
						<tbody>
							<tr>
								<td align="left" width="100">FCBA</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;<select id="fcba" name="txtFcba" style="width:100px;">

										<%
										Dim rsFcba,strsql
										Set rsFcba = Server.CreateObject("ADODB.Recordset")
										rsFcba.CursorLocation = adUseClient
										strsql =    " select fcba from user_fcba_tab " & _
													" where " & _
													" fcusercode = '" & session("USERNAME") & "' "
										
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
										%>
								</select>
								</td>
							</tr>
							<tr>
								<td align="left">DATE</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;<input id="inday" class="datepicker" type="text" maxlength="10" name="INDAY" size="10" value="<%=orecordset("fddate")%>" style="width:100px;">
								</td>
							</tr>
							<tr>
								<td align="left">STATUS</td>
								<td align="left" >
								<div class="onoffswitch">
										<input type="checkbox" name="onoffswitch" class="onoffswitch-checkbox" id="myonoffswitch" <%=Ischecked%>>
										<label class="onoffswitch-label" for="myonoffswitch">
											<span class="onoffswitch-inner"></span>
											<span class="onoffswitch-switch"></span>
										</label>
									</div>
								</td>
							</tr>
							<tr>
								<td colspan="2" align="left">
									<input type="submit" value="Save" >
								</td>
							</tr>
<%
end if
orecordset.close
Set orecordset = Nothing							
%>
						</tbody>
					</table>
					</form>
				</div>
		</div>
		<!--#include file="bottom.jsp"-->
	</body>
</html>