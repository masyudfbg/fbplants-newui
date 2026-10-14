<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
 <%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade
Dim ScriptName
Dim rsErptoSap, strErptoSap, rsMandor, strMandor, NOMOR
Dim NO, cClass, nday

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
ScriptName		=	Request.ServerVariables("SCRIPT_NAME")

strErptoSap		= 	" select t.delivery_date,t.received_date,t.wb_ticket,t.delivery_qty,t.pom_bruto,t.pom_tare,t.pom_received, " & _
					" t.pom_received - t.pom_grading pom_netto,t.uom,'' fruit_type,t.vehicleno plat_no,'' delivery_note,t.total_tandan, " & _
					" '' cpo_rendemen, '' unripKG,'' unripbunch,'' halfripeKG,'' halfripebunch,'' rottenKG,'' rottenbunch, " & _
					" '' overripeKG,'' overripebunch,'' parthenKG,'' parthenbunch,'' fr3kg,'' fr3bunch,'' emptykg,'' emptybunch, " & _
					" '' longKG, '' longbunch,'' losekg,'' waterkg,'' garbagekg,'' sandkg,t.pom_grading otherKG,'' otherbunch,'' abnormalKG, " & _
					" '' abnormalbunch, '' eatenKG,'' eatenbunch, '' duraKG,'' durabunch,'' teneraKG,'' tenerabunch,t.material_code, " & _
					" t.customer_desc estate,'PKU' Discharge,t.drivername,'NO' checkgrading,'TBS' batch from ERP_REPORT_SAP_SD t " & _
					" WHERE FCBA = '" & Session("FCBA") & "' "
										
 if  LEN(TRIM(v_txtYear)) > 0 then
	 strErptoSap	=	strErptoSap & " AND substr(t.delivery_date,7,4) = " & v_txtYear
 end if
 if  LEN(TRIM(v_txtMonth)) > 0 then
	 strErptoSap	=	strErptoSap & " AND substr(t.delivery_date,4,2) = " & v_txtMonth
 end if
 strErptoSap	=	strErptoSap & " order by wb_ticket "
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - ESTATE TO OWN POM  </title>
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
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<!--#include file="menutop.asp"-->
			<div id="header">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
		<div id="subheader" class="centered">
			<h3>INTEGRASI IPLAS - SAP</h3>
			<h4>&nbsp;</h4>
			<h2>FFB ESTATE TO OWN POM (SD / SALES & DISTRIBUTION)</h2>
			<br/>
			<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
		</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td>
<% 
	Dim mon, Yr, MinYr, MaxYr, i, MM
    mon = Month(Session("v_Costbook"))
    If len(v_txtMonth)=0 then v_txtMonth = mon

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
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th  style="color:#FFFFFF" rowspan="2">DELIVERY DATE</th>
		<th  style="color:#FFFFFF" rowspan="2">RECEIVE DATE</th>
		<th  style="color:#FFFFFF" rowspan="2">WB TICKET</th>
		<th  style="color:#FFFFFF" rowspan="2">DELIVERY QTY (Kg)</th>
		<th  style="color:#FFFFFF" rowspan="2">POM BRUTO QTY (Kg)</th>
		<th  style="color:#FFFFFF" rowspan="2">POM TARE QTY (Kg)</th>
		<th  style="color:#FFFFFF" rowspan="2">POM RECEIVE QTY (Kg)</th>
		<th  style="color:#FFFFFF" rowspan="2">POM NETTO QTY (Kg)</th>
		<th  style="color:#FFFFFF" rowspan="2">UOM</th>
		<th  style="color:#FFFFFF" rowspan="2">FRUIT TYPE</th>
		<th  style="color:#FFFFFF" rowspan="2">CAR FLATE NO</th>
		<th  style="color:#FFFFFF" rowspan="2">DELIVERY NOTE (SPB)</th>
		<th  style="color:#FFFFFF" rowspan="2">NETTO BUNCHES</th>
		<th  style="color:#FFFFFF" rowspan="2">CPO RENDEMEN</th>
		<th  style="color:#FFFFFF" colspan="2">UNRIPE BUNCH</th>
		<th  style="color:#FFFFFF" colspan="2">HALF-RIPE BUNCH</th>
		<th  style="color:#FFFFFF" colspan="2">ROTTEN BUNCH</th>
		<th  style="color:#FFFFFF" colspan="2">OVER-RIPE BUNCH</th>
		<th  style="color:#FFFFFF" colspan="2">PARTHENOCARPIC</th>
		<th  style="color:#FFFFFF" colspan="2">BUNCH < 3KG </th>
		<th  style="color:#FFFFFF" colspan="2">EMPTY BUNCH</th>
	    <th  style="color:#FFFFFF" colspan="2">LONG STACK</th>
		<th  style="color:#FFFFFF" rowspan="2">LOST FRUIT (KG)</th>
		<th  style="color:#FFFFFF" rowspan="2">WATER (KG)</th>
		<th  style="color:#FFFFFF" rowspan="2">GARBAGE (KG)</th>
		<th  style="color:#FFFFFF" rowspan="2">SAND (KG)</th>
		<th  style="color:#FFFFFF" colspan="2">OTHERS</th>
		<th  style="color:#FFFFFF" colspan="2">ABNORMAL</th>
		<th  style="color:#FFFFFF" colspan="2">EATEN BY RAT</th>
		<th  style="color:#FFFFFF" colspan="2">DURA</th>
		<th  style="color:#FFFFFF" colspan="2">TENERA</th>
		<th  style="color:#FFFFFF" rowspan="2">MATRIAL CODE</th>
		<th  style="color:#FFFFFF" rowspan="2">ESTATE</th>
		<th  style="color:#FFFFFF" rowspan="2">DISCHARGE MILL</th>
		<th  style="color:#FFFFFF" rowspan="2">DRIVER NAME</th>
		<th  style="color:#FFFFFF" rowspan="2">CHECK GRADING</th>
		<th  style="color:#FFFFFF" rowspan="2">BATCH</th>
	</tr>
	<tr>
	    <th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>
		<th  style="color:#FFFFFF" >KG</th>
		<th  style="color:#FFFFFF" >BUNCH</th>


	</tr>
<%
Set rsErptoSap = Server.CreateObject("ADODB.Recordset")
rsErptoSap.Open strErptoSap, connstr, adOpenStatic
'response.write strErptoSap
if rsErptoSap.eof then
%>
	<tr><td colspan="50">NO DATA AVAILABLE</td></tr>
<%
else
	NO = 1

	while not rsErptoSap.eof
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass ="default"
		end if
		
		
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=rsErptoSap("delivery_date")%></td>
		<td align="center"><%=rsErptoSap("received_date")%></td>
		<td align="center"><%=rsErptoSap("wb_ticket")%></td>
		<td align="center"><%=rsErptoSap("delivery_qty")%></td>
		<td align="center"><%=rsErptoSap("pom_bruto")%></td>
		<td align="center"><%=rsErptoSap("pom_tare")%></td>
		<td align="center"><%=rsErptoSap("pom_received")%></td>
		<td align="center"><%=rsErptoSap("pom_netto")%></td>
		<td align="center"><%=rsErptoSap("uom")%></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("plat_no")%></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("total_tandan")%></td>
		<td align="right"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("otherKG")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("material_code")%></td>
		<td align="center"><%=rsErptoSap("estate")%></td>
		<td align="center"><%=rsErptoSap("Discharge")%></td>
		<td align="center"><%=rsErptoSap("drivername")%></td>
		<td align="center"><%=rsErptoSap("checkgrading")%></td>
        <td align="center"><%=rsErptoSap("batch")%></td>
	</tr>	
<%
	NO = NO + 1
	rsErptoSap.MoveNext
Wend
end if
rsErptoSap.Close
set rsErptoSap = Nothing
%>
	

</table>
  
<table width="90%" class="data">
    <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	</tr>
	<tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	</tr>
</table>
    </body>
    </html>
