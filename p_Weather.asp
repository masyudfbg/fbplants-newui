<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
<title>ePlaS Reporting - Weather</title>
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtWIDATE_From= trim(request("WIDATE_From"))
v_txtWIDATE_To= trim(request("WIDATE_To"))
v_txtWIDIVISION= trim(request("WIDIVISION"))
txtSQL = "SELECT WIDATE, WIDIVISION, nvl(SUM(TOTRNINCHES),0) AS TOTRNINCHES, nvl(SUM(TOTRNDAYS),0) AS TOTRNDAYS, nvl(SUM(AVGTEMP),0) AS AVGTEMP, nvl(SUM(AVGWIND),0) AS AVGWIND FROM Weather" & _
		 " where fcba='" & session("fcBA") & "'"
if  len(trim(v_txtWIDATE_From)) > 0 then
    txtSQL = txtSQL & " and WIDATE >=to_Date('" & v_txtWIDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtWIDATE_From,1)
end if
if  len(trim(v_txtWIDATE_To)) > 0 then
    txtSQL =txtSQL &  " and WIDATE <=to_Date('" & v_txtWIDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtWIDATE_To,1)
end if
if  len(trim(v_txtWIDIVISION)) > 0 then
    txtSQL = txtSQL & " and WIDIVISION='" & v_txtWIDIVISION & "'"
	txtSelect = txtSelect & " DIVISION " & v_txtWIDIVISION
end if

txtsql=(txtsql) & " group by WIDATE, WIDIVISION"
txtsql = txtsql & " order by WIDATE, WIDIVISION"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
'response.Flush()
%>
 <body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="6" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">

					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>WEATHER INFORMATION</h2>
				</div>
			</div>
			</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="6" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Date</th>
					<th>Division</th>
					<th>Rain Inches</th>
					<th>Rain days</th>
					<th>Temperature</th>
					<th>Wind Speed</th>
					
				</tr>
                 </thead>
                            <tbody>
				<%

				Dim v_TOTRNINCHES,v_TOTRNDAYS,v_AVGTEMP,v_AVGWIND
				v_TOTRNINCHES=0
				v_TOTRNDAYS=0
				v_AVGTEMP=0
				v_AVGWIND=0				
				While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WIDATE")%></td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WIDIVISION")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("TOTRNINCHES"),2)%></td>		
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("TOTRNDAYS"), 0)%></td>		
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("AVGTEMP"),2)%></td>		
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("AVGWIND"),2)%></td>		
		</tr>
	  <%
		v_TOTRNINCHES=v_TOTRNINCHES+trim(cdbl(rsCat("TOTRNINCHES")))
		v_TOTRNDAYS=v_TOTRNDAYS+trim(cdbl(rsCat("TOTRNDAYS")))
		v_AVGTEMP=v_AVGTEMP+trim(cdbl(rsCat("AVGTEMP")))
		v_AVGWIND=v_AVGWIND+trim(cdbl(rsCat("AVGWIND")))		

	  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" colspan="2"> <strong><i> <font face="Verdana" size="1">T o t a l</i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_TOTRNINCHES,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_TOTRNDAYS, 0)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_AVGTEMP,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_AVGWIND,2)%></i></strong></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</div>
	</body>
</html>