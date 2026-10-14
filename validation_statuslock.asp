<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim textReportName : textReportName = "Validation Statuslock"
Dim ScriptName : ScriptName = "validation_statuslock.asp?txtJenis=edit"

Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim v_txtFcba, v_modul
Dim INDAY  
Dim onoffswitch  
Dim Ischecked 
Dim TrueOrFalse 

Dim orecordset, sqlquery
Dim rowdata : rowdata = 0
Dim maxlineno : maxlineno = 0


Dim UpdateDataFp

v_txtFcba   = Request("txtFcba")
onoffswitch = Request("onoffswitch")


If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
    v_txtFcba = Session("fcba")
End If


If onoffswitch = "on" Then 
    TrueOrFalse = "1"  
Else
    TrueOrFalse = "0"   
End If

If Request("txtJenis") = "edit" Then
    UpdateDataFp = " update STATUSLOCK " & _
                   " set statuslock = '" & TrueOrFalse & "'" & _
                   " where fcba = '" & v_txtFcba & "' "

    connStr.execute(UpdateDataFp)
End If
sqlquery = " select statuslock " & _
           " from STATUSLOCK " & _
           " where fcba = '" & v_txtFcba & "'"

Set orecordset = Server.CreateObject("ADODB.Recordset")
orecordset.Open sqlquery, connstr

If Not orecordset.EOF Then
    TrueOrFalse = orecordset("statuslock")
    If TrueOrFalse = "0" Then 
        Ischecked = ""
    Else
        Ischecked = "checked"
    End If
Else
    Ischecked = ""
End If

orecordset.Close
Set orecordset = Nothing
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
	</head>

	<body>
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
					<form name="frmEntry" method="POST" action="<%=ScriptName%>">
					<table class = "data">
						<thead>
							
						</thead>
						<tbody>
							<tr>
								<td align="left" width="100">FCBA</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;
                                    <select id="fcba" name="txtFcba" style="width:100px;">

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
								<td align="left">MODUL</td>
                                <td align="left">
                                :&nbsp;&nbsp;&nbsp;&nbsp;
                                <select ID="modul" name="modul" class="chosen-select">
                                    <option></option>
                                   <%
										Dim rsFcba1,strsql1
										Set rsFcba1 = Server.CreateObject("ADODB.Recordset")
										rsFcba1.CursorLocation = adUseClient
										strsql1 =    " SELECT DISTINCT modulname FROM STATUSLOCK ORDER BY modulname "
										
										rsFcba1.Open strsql1, connStr, adOpenForwardOnly, adLockReadOnly	
										if not rsFcba1.eof then
											DO while not rsFcba1.eof
												if v_txtFcba = rsFcba1("modulname") then
												%>			
													<option value="<%=rsFcba1("modulname")%>" Selected><%=rsFcba1("modulname")%></option>
												<% Else %>	
													<option value="<%=rsFcba1("modulname")%>"><%=rsFcba1("modulname")%></option>
												<% End If
												
												rsFcba1.movenext
											loop
										End if
										rsFcba1.Close
										set rsFcba1 = Nothing
										%>
                                </select>
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
					</tbody>
					</table>
					</form>
				</div>
		</div>
		<!--#include file="bottom.jsp"-->
	</body>
</html>