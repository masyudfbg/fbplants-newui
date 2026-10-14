<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
' Dim v_txtParcode, v_txtParheadcode
maxCol = 0

v_txtParcode  =  Request("txtParcode")
v_txtParheadcode  =  Request("txtParheadcode")
v_txtFcba = Request("txtFcba")
txtYear = Request("txtYear")
txtMonth = Request("txtMonth")

txttdate = Request("txttdate")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txttdate = "" Then txttdate = textDate(date)

' txttdate = txtMonth & "/1/" & txtYear

add_col_bud_prod_qty = ""
d_span = 67
	
	if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
       whereCons = whereCons & "  WHERE fcba = '" & v_txtfcba & "' "
	end if

    ' selectCons = selectCons & ",fcba"
    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & Session("fcba") & "' "
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   fcbaname = rsCat("fcName")
	   CompanyName=rsCat("FCCOMPANYNAME")
	   MgrName = rsCat("FCMANAGER")
	   KTUName = rsCat("FCKTU")
	   CropName=rsCat("FCCROP")
	   FcbaType=rsCat("fcType")
else
    nSpanTotal = nSpanTotal - 1 
end if

strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_field_eof = true
txtSQL = 	"SELECT P.PARCODE,"& _	
			"       P.PARHEADCODE,"& _	
			"       P.FCCODE,"& _	
			"       P.FCNAME,"& _	
			"       P.RATE_1,"& _	
			"       P.RATE_2,"& _	
			"       P.RATE_3,"& _	
			"       P.RATE_4,"& _	
			"       P.RATE_5,"& _
			"       P.PARAM_1,"& _	
			"       P.PARAM_2,"& _	
			"       P.PARAM_3,"& _	
			"       P.PARAM_4,"& _	
			"       P.PARAM_5,"& _
			"       P.LASTTIME"& _	
			"  FROM IPLASPROD.PARAMETERDETAIL P"&_
			" WHERE P.FCBA = '" & SESSION("FCBA") & "'"	

if  len(trim(v_txtParcode)) > 0 then
	txtSQL	=	txtSQL &  " AND P.PARCODE = '" & v_txtParcode & "'"
end if
if  len(trim(v_txtParheadcode)) > 0 then
	txtSQL	=	txtSQL &  " AND P.PARNAME LIKE '%" & v_txtParheadcode & "%'"
end if
	txtSQL = txtSQL & 	" ORDER BY P.PARCODE," & _
						"          P.PARHEADCODE," & _
						"          P.FCCODE," & _
						"          P.FCNAME," &_ 
						"          P.FCBA" 


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
' response.write(v_txtParcode)
' response.Write(txtsql)
strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & session("fcba") & "'"
' response.write strBaInfo
Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

%>
<html>

<head>
<title>ePlaS Reporting - Parameter Detail</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
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
-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:50px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>


<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">
	 <!-- include file="../print_excel.asp" -->
		<div id="header-wrapper">
			<div class="printhide">
				<!-- include file="menutop.asp" -->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">
						<tr>
							<th><font face="verdana" size="1"> KEBUN</font></th>
							<th><font face="verdana" size="1"> KODE</font></th>
						</tr>
						<tr>
							<td>
								<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
								<%
								Dim rsFcba
								Set rsFcba = Server.CreateObject("ADODB.Recordset")
								rsFcba.CursorLocation = adUseClient
								strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
								            " select fccode fcba from businessunit " & _
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
								%>
								</select>	
							</td>
							<td>
								<input type="text" maxlength="10" name="txtParcode" size="10" value="<%=v_txtParcode%>" title="Input Code">
							</td>
							<td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div  id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>PARAMETER DETAIL</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	 <tr>
		<th>PARCODE</th>
		<th>PARNHEADCODE</th>
		<th>FCCODE</th>
		<th>FCNAME</th>
		<th>RATE 1</th>
		<th>RATE 2</th>
		<th>RATE 3</th>
		<th>RATE 4</th>
		<th>RATE 5</th>
		<th>PARAM 1</th>
		<th>PARAM 2</th>
		<th>PARAM 3</th>
		<th>PARAM 4</th>
		<th>PARAM 5</th>
	</tr>
    </thead>
	<tbody>
		<% 
		Set rsCat = connStr.Execute(txtSQL)
		
		if not rsCat.eof then
			v_Quantity=0
			v_Amount = 0
		While Not rsCat.EOF%>
		<tr>
			<!-- <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=NO%> </td> -->
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("parcode")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("parheadcode")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("fccode")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("fcname")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("rate_1")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("rate_2")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("rate_3")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("rate_4")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("rate_5")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("param_1")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("param_2")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("param_3")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("param_4")%> </td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("param_5")%> </td>
		</tr>
		  <%

		  rsCat.MoveNext
		  'response.Flush()
		  no=no+1
		Wend
	end if
		rsCat.close()
		rsCat=null%>
		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
        </tbody>
</table>


<script src="/plugins/chosen/chosen.jquery.min.js"></script>
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
	</body>
</html>