<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay,v_txtDate, v_txtsectionname, v_txtBlok, v_txtColor
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName,varCompanyCode, varCompanyName
Dim oRS, oSQL
maxCol = 0

txtYear = Request("txtYear")
txtFcba = Request("txtFcba")
txtsectionname = Request("txtsectionname")
txtpresentgrade = Request("txtpresentgrade")

If "" & txtYear = "" Then txtYear = Year(date)

arr_fcba_eof = true
strDefaultBa = " SELECT fccode AS fcba" & _
				" FROM businessunit BU" & _
				" INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
				" WHERE isdefault = 'TRUE'" & _
				"       AND fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	arr_fcba = rsDefaultBa.getrows()
	arr_fcba_eof = false
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

t_fcba = 0
if not arr_fcba_eof then
	t_fcba = 0
	for xyy=0 to ubound(arr_fcba,2)
		t_fcba = t_fcba +1
		if (t_fcba = 1 and "" & txtFcba = "") then txtFcba = arr_fcba(0,xyy)
	next
end if

d_span = 8

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_data_eof = true
strPusingan	= " SELECT fcba" & _
              "        , fccode" & _
              "        , fcname" & _
              "        , sectionname" & _
              "        , presentgrade" & _
              "        , TO_CHAR(dateofbirth, 'dd-Mon-yyyy') dateofbirth" & _
              "        , "&txtYear&" - TO_NUMBER(TO_CHAR(dateofbirth,'yyyy')) umur" & _
              " FROM employee" & _
              " WHERE"
if "" & txtFcba = "All" then
	strPusingan = strPusingan & ""
else
	strPusingan = strPusingan & " fcba = '"&txtFcba&"' AND"
end if
strPusingan = strPusingan & "       dateterminate IS NULL" & _
                            "       AND "&txtYear&" - TO_NUMBER(TO_CHAR(dateofbirth,'yyyy')) >= 55" & _
                            "       AND presentgrade <> 'STF'"
				
if "" & txtsectionname <> "" then strPusingan = strPusingan & " AND sectionname = '" & txtsectionname & "'"
if "" & txtpresentgrade <> "" then strPusingan = strPusingan & " AND presentgrade = '" & txtpresentgrade & "'"
strPusingan = strPusingan & " ORDER BY fcba, sectionname, fccode"

'response.write strPusingan

Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_data = qORS.getrows()
	arr_data_eof = false
End If
qORS.Close
set qORS = Nothing

strBaInfo = " SELECT * FROM BUSINESSUNIT WHERE FCCODE = '" & txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("fcname")
	varCompanyName = rsBaInfo("fccompanyname")
	varCompanyCode = rsBaInfo("fccompanycode")
	varFcManager = rsBaInfo("fcmanager")
	varFcKTU = rsBaInfo("fcktu")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
	varFcManager = ""
	varFcKTU = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

Function f_maxrow_fcba(data_array__, code__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if cstr("" & arr_1(0, indx)) = cstr("" & code__) then
			result_ = result_ +1
		end if
	next
	f_maxrow_fcba = result_
End Function

%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>Pensiun - iPlaS Reporting</title>
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
  min-width:40px;
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
						<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM

						Yr = txtYear
						MinYr = 2015
						MaxYr = Yr  + 2
					%>
							<td width="4">
								<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
								</select>
							</td>
							
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<option value="All">--All--</option>
									<%
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>

							<td>
								<font face="verdana" size="1">SECTION</font>
							</td>
							<td>
								<select name="txtsectionname" >
									<option value="">--All--</option>
									<%
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'E0002' AND FCBA = '" & txtFcba & "' ORDER BY FCCODE"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = txtsectionname then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
									<%End if
										rsAfdeling.movenext
									loop
									end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							
							<td>
								<font face="verdana" size="1">STATUS</font>
							</td>
							<td>
								<select name="txtpresentgrade" >
									<option value="">--All--</option>
									<%
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'E0001' AND FCBA = '" & txtFcba & "' ORDER BY FCCODE"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = txtpresentgrade then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
									<%End if
										rsAfdeling.movenext
									loop
									end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							

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
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>KARYAWAN PENSIUN TAHUN <%=txtYear%></h2>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th><%=titleSite%></th>
		<th>No</th>		
		<th>Kode Karyawan</th>
		<th>Nama Karyawan</th>
		<th>Unit Kerja</th>
		<th>Status</th>
		<th>Tanggal Lahir</th>
		<th>Umur</th>
	</tr>
    </thead>
	<tbody>
<%
	if not arr_data_eof then
		no = 0
		for y=0 to ubound(arr_data,2)
			no = no +1
			curr_fcba = arr_data(0,y)
			
			response.write "<tr>"
			if curr_fcba <> prev_fcba then
				maxST = f_maxrow_fcba(arr_data, curr_fcba)
				response.write "	<td valign='top' align='center' rowspan='"&maxST&"'>"&arr_data(0,y)&"</td>"
			end if
			
			response.write "	<td align='center'>"&no&"</td>"
			response.write "	<td align='center'>"&arr_data(1,y)&"</td>"
			response.write "	<td align='left'>"&arr_data(2,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(3,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(4,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(5,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(6,y)&"</td>"
			response.write "</tr>"
			
			if no = maxST then no = 0
			
			prev_fcba = curr_fcba
		next
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
</tbody>
</table>

   <table width="90%" class="data">
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>

	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varFcManager%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varFcKTU%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
	</body>
</html>