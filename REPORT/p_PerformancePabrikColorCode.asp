<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	' Query Filter Region 
	Dim rsEPS,strMain, strSqlRow,rsSqlRow, strsql,rsCount,visibleR,nSpanTotal
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth, v_txttodaydate, NO, cClass
	Dim v_txtWilayah, txtyear, txtmonth
	Dim ColorAch
	
	strsql = " SELECT count(*) ncount from ( SELECT DISTINCT a.region fccode  " & _
			 " FROM Businessunit a, region b  " & _
			 " WHERE a.region = b.fccode  " & _
			 " AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', a.fccode) =  'TRUE'  ) " 
	Set rsCount = Server.CreateObject("ADODB.Recordset")
	rsCount.Open strsql, connStr
  if not rsCount.eof then
     visibleR = rsCount("ncount")
	 rsCount.Close
	 if cint(visibleR) = 1 then 
	 
	strsql = " SELECT DISTINCT a.region fccode  " & _
			 " FROM Businessunit a, region b  " & _
			 " WHERE a.region = b.fccode  " & _
			 " AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', a.fccode) =  'TRUE'  " 
		Set rsCount = Server.CreateObject("ADODB.Recordset")
		 rsCount.Open strsql, connStr
		 txtWilayah = rsCount("fccode")
		 rsCount.Close
	 end if	 
  else
     visibleR = 0
	 rsCount.Close
  end if  
  'response.write visibleR
%>
<%
	' Query Main Grid 	
	v_txtyear 		= Request("txtyear")
	v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	v_txtWilayah  	= Request("txtWilayah")
	nSpanTotal 		= 8
	
	if len(trim(request("txtYear"))) = 0 then
			v_txtyear = year(Now)	 
		end if

		if len(trim(request("txtmonth"))) = 0 then
			v_txtmonth = Month(Now)	 
	end if
	
	Dim wherecons, selectCons, orderCons
	wherecons = "where COMPANY is not null"
	
	if  v_txtWilayah <> "--SUMMARY--"  and v_txtWilayah <> "" then
		if v_txtWilayah <> "--ALL--" then
			whereCons = whereCons & " And CASE COMPANY WHEN 'PT.THIP' THEN 'THIP'  WHEN 'PT.CRS' THEN 'CRS' ELSE 'AMS GANDA' END = '" & v_txtWilayah & "'"
		end if
			selectCons = selectCons & " ,region"
			orderCons  = orderCons & ", REGION_API.get_OrderData(region) "
		else
			nSpanTotal = nSpanTotal - 1 
	end if
	
	strMain = "select CASE COMPANY WHEN 'PT.THIP' THEN 'THIP'  WHEN 'PT.CRS' THEN 'CRS' " &_
			  "ELSE 'AMS GANDA' END WILAYAH, "&_
			  "CASE COMPANY WHEN 'PT.THIP' THEN 1 WHEN 'PT.CRS' THEN 3 " &_
				"ELSE 2 END URUTAN, X.*  from MGT_MONTHLY_RPT_MILL_TAB X "& whereCons &" and year= '"& v_txtyear &"' and month='"& v_txtmonth &"' order by URUTAN asc"
	Set rsEPS = Server.CreateObject("ADODB.Recordset")
	rsEPS.Open strMain , connstr, adOpenStatic
%>
<html>
<head>
<title>ePlaS Reporting - Performance Pabrik</title>
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
$( ".datepicker" ).datepicker();
});
</script>

<!--<script type="text/javascript" src="http://jqueryjs.googlecode.com/files/jquery-1.3.1.min.js" > </script> -->
<script type="text/javascript">

    function PrintElem(elem)
    {
        Popup($(elem).html());
    }

    function Popup(data) 
    {
        var mywindow = window.open('', 'my div', 'height=400,width=600');
        mywindow.document.write('<html><head><title>my div</title>');
        mywindow.document.write('<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />');
        mywindow.document.write('</head><body >');
        mywindow.document.write(data);
        mywindow.document.write('</body></html>');

        mywindow.print();
        mywindow.close();

        return true;
    }

</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);
}
</script>-->
</head>
<body>
<div id="wrapper">
<!--#include file="../print_excel.asp"-->
<div class="top_select" style=" position:absolute; top:0; left:0px; width:100%;">
<form name="frmEntry" method="post" action="p_PerformancePabrikColorCode.asp">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
<table border="0" id="form-table" class="centered-margin">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %-->
		<% 
					MM=Request("txtMonth")
					yyyy=Request("txtYear")

					mon = Month(Session("v_Costbook"))
					If len(MM)=0 then MM=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = 2015
					MaxYr = Yr  + 2
					If len(yyyy)= 0 then yyyy = Yr
				%>
				<th align="left">Wilayah</th>
					<td>
						<select name="txtWilayah">
								<option value="--ALL--" <%if request("v_txtWilayah")="--ALL--" then%> selected <%End if%>>--ALL--</option>
								<%
									Dim ListWil
									ListWil = Array("THIP", "AMS GANDA", "CRS")
								For i =0 to 2
									if v_txtWilayah = ListWil(i) then
								%>
									<Option Value="<%=ListWil(i)%>" Selected><%=ListWil(i)%></option>
								<%
									ELSE
								%>
									<Option Value="<%=ListWil(i)%>"><%=ListWil(i)%></option>
								<%
									end if
									Next
								%>
						</select>
					</td>
			<th align="left">Bulan</th>
		   <td>
				<font size="2" color="#0000FF" ></font>
					<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
						<%
							Next
						%>
					</select>
		  </td>	
			<th align="left">Tahun</th>
		   <td>
				<font size="2"  color="#0000FF" ></font>
				<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
				</select>
			</td>
					
		<td align="center"><input type="submit" value="View" /></td>

	</tr>
    <tr>
	<td colspan="140" style=" ">
		<div id="header-wrapper"> 
			<div id="subheader"class="centered">
				<h2>Laporan Performance Pabrik</h2>
				<h4><i>Color Code</i></h2>&nbsp;
				<%
					Dim Periode
					IF Request.ServerVariables("REQUEST_METHOD")= "POST" then 
						Periode = MonthName(v_txtmonth) & " - " & v_txtyear
					else
						Periode = "-"
					end if
				%>
				<h5>Periode : <%=Periode %></h5>
			</div>
		</div>
	</td>
</tr>
</table>
</div> 

<div id="mydiv" style="margin-top:200px; z-index:10;">
<table border="0" width="100%" id="product-table" class="xthf-site" style="background-color:#FFFFFF;">
<thead>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">WIL</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">MILL</th>
		<th colspan="6" style="background-color:#00A300;color:#fff;">FISIK(%)</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BIAYA(%)</th>
        <th colspan="2" style="background-color:#00A300;color:#fff;">ACHIEVE(%)</th>
	</tr>
	<tr>
		<th colspan="2" style="background-color:#00A300;color:#fff;">THROUGHPUT</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">OER</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">FFA</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PANEN TM</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BI</th>
        <th rowspan="2"style="background-color:#00A300;color:#fff;">SBI</th>
	</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
	</tr>
	</thead>
	<tbody>
		<%
		Dim Ttl_TROUGHPUT_BI, Ttl_TROUGHPUT_SBI, Ttl_BUD_TROUGHPUT_BI, Ttl_BUD_TROUGHPUT_SBI, Ttl_ACH_TROUGHPUT_BI, Ttl_ACH_TROUGHPUT_SBI
		Dim Ttl_OER_BI, Ttl_OER_SBI, Ttl_BUD_OER_BI, Ttl_BUD_OER_SBI, Ttl_ACH_OER_BI, Ttl_ACH_OER_SBI
		Dim Ttl_FFA_BI, Ttl_FFA_SBI, Ttl_BUD_FFA_BI, Ttl_BUD_FFA_SBI, Ttl_ACH_FFA_BI, Ttl_ACH_FFA_SBI, Ttl_CALC_ACH_FFA_BI, Ttl_CALC_ACH_FFA_SBI 
		Dim Ttl_AMT_BI, Ttl_AMT_SBI, Ttl_BUD_AMT_BI, Ttl_BUD_AMT_SBI, Ttl_ACH_AMT_BI, Ttl_ACH_AMT_SBI, Ttl_CALC_ACH_AMT_BI, Ttl_CALC_ACH_AMT_SBI
		Dim Ttl_All_TBS_PROCESS_BI, Ttl_All_TBS_PROCESS_SBI, Ttl_All_BUD_TBS_PROCESS_BI, Ttl_All_BUD_TBS_PROCESS_SBI, Ttl_All_CPO_BI, Ttl_All_CPO_SBI, Ttl_All_BUD_CPO_BI, Ttl_All_BUD_CPO_SBI
		Dim Ttl_ACH_BI, Ttl_ACH_SBI 

		Dim Grnd_Ttl_TROUGHPUT_BI,Grnd_Ttl_TROUGHPUT_SBI, Grnd_Ttl_BUD_TROUGHPUT_BI,Grnd_Ttl_BUD_TROUGHPUT_SBI, Grnd_Ttl_ACH_TROUGHPUT_BI, Grnd_Ttl_ACH_TROUGHPUT_SBI
		Dim Grnd_Ttl_OER_BI, Grnd_Ttl_OER_SBI, Grnd_Ttl_BUD_OER_BI, Grnd_Ttl_BUD_OER_SBI, Grnd_Ttl_ACH_OER_BI, Grnd_Ttl_ACH_OER_SBI 
		Dim Grnd_Ttl_FFA_BI, Grnd_Ttl_FFA_SBI, Grnd_Ttl_BUD_FFA_BI, Grnd_Ttl_BUD_FFA_SBI, Grnd_Ttl_ACH_FFA_BI, Grnd_Ttl_ACH_FFA_SBI, Grnd_Ttl_CALC_ACH_FFA_BI, Grnd_Ttl_CALC_ACH_FFA_SBI 
		Dim Grnd_Ttl_AMT_BI, Grnd_Ttl_AMT_SBI, Grnd_Ttl_BUD_AMT_BI, Grnd_Ttl_BUD_AMT_SBI, Grnd_Ttl_ACH_AMT_BI, Grnd_Ttl_ACH_AMT_SBI, Grnd_Ttl_CALC_ACH_AMT_BI, Grnd_Ttl_CALC_ACH_AMT_SBI
		Dim	Grnd_Ttl_All_TBS_PROCESS_BI, Grnd_Ttl_All_TBS_PROCESS_SBI, Grnd_Ttl_All_BUD_TBS_PROCESS_BI, Grnd_Ttl_All_BUD_TBS_PROCESS_SBI
		Dim Grnd_Ttl_All_CPO_BI, Grnd_Ttl_All_CPO_SBI, Grnd_Ttl_All_BUD_CPO_BI, Grnd_Ttl_All_BUD_CPO_SBI
		Dim Grnd_Ttl_ACH_BI, Grnd_Ttl_ACH_SBI
		if not rsEPS.EOF then
		NO = 1
		while not rsEPS.EOF
		if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
		%>
		<tr class="<%=cClass%>">
			<%
			strSqlRow = "Select count(*) RowsCount from MGT_MONTHLY_RPT_MILL_TAB X where CASE COMPANY WHEN 'PT.THIP' THEN 'THIP'  WHEN 'PT.CRS' THEN 'CRS' ELSE 'AMS GANDA' END = '"& rsEPS("WILAYAH") &"' and year= '"& v_txtyear &"' and month='"& v_txtmonth &"'"
			Set rsSqlRow = Server.CreateObject("ADODB.Recordset")
			rsSqlRow.Open strSqlRow ,connstr, adOpenStatic		
				if NO = 1 then
			%>
				<td align="center" Rowspan="<%=CDbl(rsSqlRow("RowsCount"))+1 %>"><%=rsEPS("WILAYAH")%></td>
			<%
				end if
			%>
			<td align="center"><%=rsEPS("FCBA")%></td>	
			<%
			if rsEPS("ACH_TROUGHPUT_BI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_TROUGHPUT_BI") > 85 and rsEPS("ACH_TROUGHPUT_BI") <= 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_TROUGHPUT_BI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(rsEPS("ACH_TROUGHPUT_BI"),0)%></td>
			<%
			if rsEPS("ACH_TROUGHPUT_SBI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_TROUGHPUT_SBI") > 85 and rsEPS("ACH_TROUGHPUT_SBI") <= 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_TROUGHPUT_SBI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="Right" style="background-color:<%=ColorAch%>"><%=FormatNumber(rsEPS("ACH_TROUGHPUT_SBI"),0)%></td>
			<%
			if rsEPS("ACH_OER_BI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_OER_BI") > 85 and rsEPS("ACH_OER_BI") <= 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_OER_BI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="Right" style="background-color:<%=ColorAch%>"><%=FormatNumber(rsEPS("ACH_OER_BI"),0)%></td>
			<%
			if rsEPS("ACH_OER_SBI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_OER_SBI") > 85 and rsEPS("ACH_OER_SBI") <= 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_OER_SBI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="Right" style="background-color:<%=ColorAch%>"><%=FormatNumber(rsEPS("ACH_OER_SBI"),0)%></td>
			<%
				Dim CALC_ACH_FFA_BI
				if rsEPS("ACH_FFA_BI") > 105 then
					ColorAch = "#FF0000"
					CALC_ACH_FFA_BI = 85
				elseif (rsEPS("ACH_FFA_BI") > 100 and rsEPS("ACH_FFA_BI") < 105) then
					ColorAch = "#FFFF00"
					CALC_ACH_FFA_BI = 95
				elseif rsEPS("ACH_FFA_BI") <= 100 then
					ColorAch = "#00FF00"
					CALC_ACH_FFA_BI = 100
				end if
			%>
			<td align="Right"  style="background-color:<%=ColorAch%>"><%=FormatNumber(CALC_ACH_FFA_BI,0)%></td>
			<%
				Dim CALC_ACH_FFA_SBI
				if rsEPS("ACH_FFA_SBI") > 105 then
					ColorAch = "#FF0000"
					CALC_ACH_FFA_SBI = 85
				elseif (rsEPS("ACH_FFA_SBI") > 100 and rsEPS("ACH_FFA_SBI") < 105) then
					ColorAch = "#FFFF00"
					CALC_ACH_FFA_SBI = 95
				elseif rsEPS("ACH_FFA_SBI") <= 100 then
					ColorAch = "#00FF00"
					CALC_ACH_FFA_SBI = 100
				end if
			%>
			<td align="Right"  style="background-color:<%=ColorAch%>"><%=FormatNumber(CALC_ACH_FFA_SBI,0)%></td>
			<%
				Dim AMT_BI, AMT_SBI,BUD_AMT_BI,BUD_AMT_SBI
				IF CDbl(rsEPS("TBS_PROCESS_BI")) <> 0 Then
					AMT_BI 		= CDbl(rsEPS("AMT_BI")) / CDbl(rsEPS("TBS_PROCESS_BI"))
				Else 
					AMT_BI 		= 0
				End IF
				
				IF CDbl(rsEPS("TBS_PROCESS_SBI")) <> 0 Then
					AMT_SBI		= CDbl(rsEPS("AMT_SBI"))/CDbl(rsEPS("TBS_PROCESS_SBI"))
				ELSE
					AMT_SBI		= 0
				End IF
				
				IF CDbl(rsEPS("BUD_TBS_PROCESS_BI")) <> 0 Then
					BUD_AMT_BI	= CDbl(rsEPS("BUD_AMT_BI"))/CDbl(rsEPS("BUD_TBS_PROCESS_BI"))
				ELSE
					BUD_AMT_BI	= 0 
				End IF
				
				IF CDbl(rsEPS("BUD_TBS_PROCESS_SBI")) <> 0 Then
					BUD_AMT_SBI	= CDbl(rsEPS("BUD_AMT_SBI"))/CDbl(rsEPS("BUD_TBS_PROCESS_SBI"))
				ELSE
					BUD_AMT_SBI = 0
				End IF
				
				Dim ACH_AMT_BI, ACH_AMT_SBI
				IF BUD_AMT_BI <> 0 THEN
					ACH_AMT_BI = (AMT_BI / BUD_AMT_BI) * 100 
				ELSE
					ACH_AMT_BI = 0
				END If
				
				IF BUD_AMT_SBI <> 0 THEN
					ACH_AMT_SBI = (AMT_SBI / BUD_AMT_SBI) * 100 
				ELSE
					ACH_AMT_SBI = 0
				END If
				
				Dim CALC_ACH_AMT_BI
				if ACH_AMT_BI < 90 OR ACH_AMT_BI > 110 then
					ColorAch = "#FF0000"
					CALC_ACH_AMT_BI = 85
				elseif ((ACH_AMT_BI >= 90 and ACH_AMT_BI < 95) OR (ACH_AMT_BI > 105 and ACH_AMT_BI < 110)) then
					ColorAch = "#FFFF00"
					CALC_ACH_AMT_BI = 95
				elseif ACH_AMT_BI >= 95 and ACH_AMT_BI <= 105 then
					ColorAch = "#00FF00"
					CALC_ACH_AMT_BI = 100
				end if
			%>
			<td align="Right"  style="background-color:<%=ColorAch%>"><%=FormatNumber(CALC_ACH_AMT_BI,0)%></td>
			<%
				Dim CALC_ACH_AMT_SBI
				if ACH_AMT_SBI < 90 OR ACH_AMT_SBI > 110 then
					ColorAch = "#FF0000"
					CALC_ACH_AMT_SBI = 85
				elseif ((ACH_AMT_SBI >= 90 and ACH_AMT_SBI < 95) OR (ACH_AMT_SBI > 105 and ACH_AMT_SBI < 110)) then
					ColorAch = "#FFFF00"
					CALC_ACH_AMT_SBI = 95
				elseif ACH_AMT_SBI >= 95 and ACH_AMT_SBI <= 105 then
					ColorAch = "#00FF00"
					CALC_ACH_AMT_SBI = 100
				end if
			%>
			<td align="Right"  style="background-color:<%=ColorAch%>"><%=FormatNumber(CALC_ACH_AMT_SBI,0)%></td>
			<%
				Dim ACH_BI,ACH_SBI
				ACH_BI 	= (rsEPS("ACH_TROUGHPUT_BI")* 0.15)+ (rsEPS("ACH_OER_BI") * 0.40) + (CALC_ACH_FFA_BI * 0.15) + (CALC_ACH_AMT_BI * 0.30)
				ACH_SBI	= (rsEPS("ACH_TROUGHPUT_SBI")* 0.15)+ (rsEPS("ACH_OER_SBI") * 0.40) + (CALC_ACH_FFA_SBI * 0.15) + (CALC_ACH_AMT_SBI * 0.30)
			%>
			<td align="right" style="background-color:<%=ColorAch%>;"><b><%=FormatNumber(ACH_BI,0) %></b></td>
			<td align="right" style="background-color:<%=ColorAch%>;"><b><%=FormatNumber(ACH_SBI,0) %></b></td>
		</tr>
		<%
			Dim Ttl_All_AMT_BI, Ttl_All_AMT_SBI, Ttl_All_BUD_AMT_BI, Ttl_All_BUD_AMT_SBI
			
			Ttl_TROUGHPUT_BI		= Ttl_TROUGHPUT_BI + rsEPS("TROUGHPUT_BI")          
			Ttl_TROUGHPUT_SBI		= Ttl_TROUGHPUT_SBI + rsEPS("TROUGHPUT_SBI")        
			Ttl_BUD_TROUGHPUT_BI	= Ttl_BUD_TROUGHPUT_BI + rsEPS("BUD_TROUGHPUT_BI")  
			Ttl_BUD_TROUGHPUT_SBI	= Ttl_BUD_TROUGHPUT_SBI + rsEPS("BUD_TROUGHPUT_SBI")
		
			Ttl_FFA_BI				= Ttl_FFA_BI + rsEPS("FFA_BI")           
			Ttl_FFA_SBI				= Ttl_FFA_SBI + rsEPS("FFA_SBI")         
			Ttl_BUD_FFA_BI			= Ttl_BUD_FFA_BI + rsEPS("BUD_FFA_BI")   
			Ttl_BUD_FFA_SBI			= Ttl_BUD_FFA_SBI + rsEPS("BUD_FFA_SBI")
		
			Ttl_All_AMT_BI			= Ttl_All_AMT_BI + CDbl(rsEPS("AMT_BI"))
			Ttl_All_AMT_SBI			= Ttl_All_AMT_SBI + CDbl(rsEPS("AMT_SBI"))
			Ttl_All_BUD_AMT_BI		= Ttl_All_BUD_AMT_BI +  CDbl(rsEPS("BUD_AMT_BI"))
			Ttl_All_BUD_AMT_SBI		= Ttl_All_BUD_AMT_SBI +  CDbl(rsEPS("BUD_AMT_SBI"))
			
			Ttl_ACH_BI				= Ttl_ACH_BI + ACH_BI
			Ttl_ACH_SBI				= Ttl_ACH_SBI + ACH_SBI	
			
			Ttl_All_TBS_PROCESS_BI			= Ttl_All_TBS_PROCESS_BI + rsEPS("TBS_PROCESS_BI")
			Ttl_All_TBS_PROCESS_SBI			= Ttl_All_TBS_PROCESS_SBI + rsEPS("TBS_PROCESS_SBI")
			Ttl_All_BUD_TBS_PROCESS_BI		= Ttl_All_BUD_TBS_PROCESS_BI + rsEPS("BUD_TBS_PROCESS_BI")
			Ttl_All_BUD_TBS_PROCESS_SBI		= Ttl_All_BUD_TBS_PROCESS_SBI + rsEPS("BUD_TBS_PROCESS_SBI")
			Ttl_All_CPO_BI					= Ttl_All_CPO_BI + rsEPS("CPO_BI")
			Ttl_All_CPO_SBI					= Ttl_All_CPO_SBI + rsEPS("CPO_SBI")
			Ttl_All_BUD_CPO_BI				= Ttl_All_BUD_CPO_BI + rsEPS("BUD_CPO_BI")
			Ttl_All_BUD_CPO_SBI				= Ttl_All_BUD_CPO_SBI + rsEPS("BUD_CPO_SBI")
		
		if NO = CDbl(rsSqlRow("RowsCount"))then
			NO = 1
		%>
		<tr>
			<td align="right" style="background-color:#FFD700;"><b>Total</b></td>
			<%
				IF Ttl_BUD_TROUGHPUT_BI <> 0 Then 
					Ttl_ACH_TROUGHPUT_BI = (Ttl_TROUGHPUT_BI / Ttl_BUD_TROUGHPUT_BI) * 100
				Else 
					Ttl_ACH_TROUGHPUT_BI = 0 
				END iF
				
				IF Ttl_BUD_TROUGHPUT_SBI <> 0 Then 
					Ttl_ACH_TROUGHPUT_SBI = (Ttl_TROUGHPUT_SBI / Ttl_BUD_TROUGHPUT_SBI) * 100
				Else 
					Ttl_ACH_TROUGHPUT_SBI = 0 
				END iF
			%>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_ACH_TROUGHPUT_BI,0)%></td>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_ACH_TROUGHPUT_SBI,0)%></td>
			<%
				IF Ttl_All_TBS_PROCESS_BI <> 0 THEN
					Ttl_OER_BI		= (Ttl_All_CPO_BI / Ttl_All_TBS_PROCESS_BI) * 100
				ELSE
					Ttl_OER_BI		= 0 
				End IF
				
				IF Ttl_All_TBS_PROCESS_SBI <> 0 THEN
					Ttl_OER_SBI		= (Ttl_All_CPO_SBI / Ttl_All_TBS_PROCESS_SBI) * 100
				ELSE
					Ttl_OER_SBI		= 0 
				End IF
				
				IF Ttl_All_BUD_TBS_PROCESS_BI <> 0 THEN
					Ttl_BUD_OER_BI		= (Ttl_All_BUD_CPO_BI / Ttl_All_BUD_TBS_PROCESS_BI) * 100
				ELSE
					Ttl_BUD_OER_BI		= 0 
				End IF
				
				IF Ttl_All_BUD_TBS_PROCESS_SBI <> 0 THEN
					Ttl_BUD_OER_SBI		= (Ttl_All_BUD_CPO_SBI / Ttl_All_BUD_TBS_PROCESS_SBI) * 100
				ELSE
					Ttl_BUD_OER_SBI		= 0 
				End IF
				
				IF Ttl_BUD_OER_BI <> 0 THEN
					Ttl_ACH_OER_BI		= (Ttl_OER_BI / Ttl_BUD_OER_BI)* 100
				ELSE
					Ttl_ACH_OER_BI		= 0
				End If
				
				IF Ttl_BUD_OER_SBI <> 0 THEN
					Ttl_ACH_OER_SBI		= (Ttl_OER_SBI / Ttl_BUD_OER_SBI)* 100
				ELSE
					Ttl_ACH_OER_SBI		= 0
				End If
			%>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_ACH_OER_BI,0)%></td>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_ACH_OER_SBI,0)%></td>
			<%
				IF Ttl_BUD_FFA_BI <> 0 THEN
					Ttl_ACH_FFA_BI = (Ttl_FFA_BI / Ttl_BUD_FFA_BI) * 100 
				ELSE
					Ttl_ACH_FFA_BI = 0
				END If
				
				IF Ttl_BUD_FFA_SBI <> 0 THEN
					Ttl_ACH_FFA_SBI = (Ttl_FFA_SBI / Ttl_BUD_FFA_SBI) * 100 
				ELSE
					Ttl_ACH_FFA_SBI = 0
				END If
				
				if Ttl_ACH_FFA_BI > 105 then
					Ttl_CALC_ACH_FFA_BI = 85
				elseif (Ttl_ACH_FFA_BI > 100 and Ttl_ACH_FFA_BI < 105) then
					Ttl_CALC_ACH_FFA_BI = 95
				elseif Ttl_ACH_FFA_BI <= 100 then
					Ttl_CALC_ACH_FFA_BI = 100
				end if
				
				if Ttl_ACH_FFA_SBI > 105 then
					Ttl_CALC_ACH_FFA_SBI = 85
				elseif (Ttl_ACH_FFA_SBI > 100 and Ttl_ACH_FFA_SBI < 105) then
					Ttl_CALC_ACH_FFA_SBI = 95
				elseif Ttl_ACH_FFA_SBI <= 100 then
					Ttl_CALC_ACH_FFA_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_CALC_ACH_FFA_BI,0)%></td>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_CALC_ACH_FFA_SBI,0)%></td>
			<%
				IF Ttl_All_TBS_PROCESS_BI <> 0 THEN
					Ttl_AMT_BI = Ttl_All_AMT_BI / Ttl_All_TBS_PROCESS_BI
				ELSE 
					Ttl_AMT_BI = 0 
				END IF
				
				IF Ttl_All_TBS_PROCESS_SBI <> 0 THEN
					Ttl_AMT_SBI = Ttl_All_AMT_SBI / Ttl_All_TBS_PROCESS_SBI
				ELSE 
					Ttl_AMT_SBI = 0 
				END IF
				
				IF Ttl_All_BUD_TBS_PROCESS_BI <> 0 THEN
					Ttl_BUD_AMT_BI = Ttl_All_BUD_AMT_BI / Ttl_All_BUD_TBS_PROCESS_BI
				ELSE 
					Ttl_BUD_AMT_BI = 0 
				END IF
				
				IF Ttl_All_BUD_TBS_PROCESS_SBI <> 0 THEN
					Ttl_BUD_AMT_SBI = Ttl_All_BUD_AMT_SBI / Ttl_All_BUD_TBS_PROCESS_SBI
				ELSE 
					Ttl_BUD_AMT_SBI = 0 
				END IF
			
				IF Ttl_BUD_AMT_BI <> 0 THEN
					Ttl_ACH_AMT_BI 	= (Ttl_AMT_BI / Ttl_BUD_AMT_BI) * 100
				ELSE
					Ttl_ACH_AMT_BI 	= 0 
				END IF
				
				IF Ttl_BUD_AMT_SBI <> 0 THEN
					Ttl_ACH_AMT_SBI 	= (Ttl_AMT_SBI / Ttl_BUD_AMT_SBI) * 100
				ELSE
					Ttl_ACH_AMT_SBI 	= 0 
				END IF
				
				if Ttl_ACH_AMT_BI < 90 OR Ttl_ACH_AMT_BI > 110 then
					Ttl_CALC_ACH_AMT_BI = 85
				elseif ((Ttl_ACH_AMT_BI >= 90 and Ttl_ACH_AMT_BI < 95) OR (Ttl_ACH_AMT_BI > 105 and Ttl_ACH_AMT_BI < 110)) then
					Ttl_CALC_ACH_AMT_BI = 95
				elseif Ttl_ACH_AMT_BI <= 95 and Ttl_ACH_AMT_BI <= 105 then
					Ttl_CALC_ACH_AMT_BI = 100
				end if
			
				if Ttl_ACH_AMT_SBI < 90 OR Ttl_ACH_AMT_SBI > 110 then
					Ttl_CALC_ACH_AMT_SBI = 85
				elseif ((Ttl_ACH_AMT_SBI >= 90 and Ttl_ACH_AMT_SBI < 95) OR (Ttl_ACH_AMT_SBI > 105 and Ttl_ACH_AMT_SBI < 110)) then
					Ttl_CALC_ACH_AMT_SBI = 95
				elseif Ttl_ACH_AMT_SBI <= 95 and Ttl_ACH_AMT_SBI <= 105 then
					Ttl_CALC_ACH_AMT_SBI = 100
				end if
			%>
			
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_CALC_ACH_AMT_BI,0)%></td>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_CALC_ACH_AMT_SBI,0)%></td>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_ACH_BI,0)%></td>
			<td align="right" style="background-color:#FFD700;"><%=FormatNumber(Ttl_ACH_SBI,0)%></td>
		</tr>
		<%
			Dim Grnd_Ttl_All_AMT_BI,Grnd_Ttl_All_AMT_SBI,Grnd_Ttl_All_BUD_AMT_BI,Grnd_Ttl_All_BUD_AMT_SBI
			
			Grnd_Ttl_TROUGHPUT_BI		= Grnd_Ttl_TROUGHPUT_BI + Ttl_TROUGHPUT_BI
			Ttl_TROUGHPUT_BI			= 0
			Grnd_Ttl_TROUGHPUT_SBI		= Grnd_Ttl_TROUGHPUT_SBI + Ttl_TROUGHPUT_SBI
			Ttl_TROUGHPUT_SBI 			= 0
			Grnd_Ttl_BUD_TROUGHPUT_BI	= Grnd_Ttl_BUD_TROUGHPUT_BI + Ttl_BUD_TROUGHPUT_BI
			Ttl_BUD_TROUGHPUT_BI 		= 0
			Grnd_Ttl_BUD_TROUGHPUT_SBI 	= Grnd_Ttl_BUD_TROUGHPUT_SBI + Ttl_BUD_TROUGHPUT_SBI 	
			Ttl_BUD_TROUGHPUT_SBI 		= 0 
			Ttl_ACH_TROUGHPUT_BI		= 0
			Ttl_ACH_TROUGHPUT_SBI 		= 0
			
			Ttl_OER_BI 					= 0 						
			Ttl_OER_SBI 				= 0 				  
			Ttl_BUD_OER_BI 				= 0 			
			Ttl_BUD_OER_SBI 			= 0
			Ttl_ACH_OER_BI				= 0
			Ttl_ACH_OER_SBI				= 0
			
			Grnd_Ttl_FFA_BI 			= Grnd_Ttl_FFA_BI + Ttl_FFA_BI 						  
			Ttl_FFA_BI 					= 0
			Grnd_Ttl_FFA_SBI 			= Grnd_Ttl_FFA_SBI + Ttl_FFA_SBI 						
			Ttl_FFA_SBI 				= 0
			Grnd_Ttl_BUD_FFA_BI 		= Grnd_Ttl_BUD_FFA_BI + Ttl_BUD_FFA_BI 				  
			Ttl_BUD_FFA_BI 				= 0
			Grnd_Ttl_BUD_FFA_SBI 		= Grnd_Ttl_BUD_FFA_SBI + Ttl_BUD_FFA_SBI 				
			Ttl_BUD_FFA_SBI 			= 0		  
			Ttl_ACH_FFA_BI 				= 0
			Ttl_ACH_FFA_SBI 			= 0	
			Ttl_CALC_ACH_FFA_BI			= 0
			Ttl_CALC_ACH_FFA_SBI		= 0
			
			Ttl_AMT_BI 					= 0  						
			Ttl_AMT_SBI 				= 0 
			Ttl_BUD_AMT_BI 				= 0
			Ttl_BUD_AMT_SBI 			= 0 				  
			Ttl_ACH_AMT_BI 				= 0 				
			Ttl_ACH_AMT_SBI 			= 0
			Ttl_CALC_ACH_AMT_BI			= 0
			Ttl_CALC_ACH_AMT_SBI		= 0
			
			Grnd_Ttl_All_AMT_BI					= Grnd_Ttl_All_AMT_BI + Ttl_All_AMT_BI
			Ttl_All_AMT_BI						= 0
			Grnd_Ttl_All_AMT_SBI				= Grnd_Ttl_All_AMT_SBI + Ttl_All_AMT_SBI
			Ttl_All_AMT_SBI						= 0
			Grnd_Ttl_All_BUD_AMT_BI				= Grnd_Ttl_All_BUD_AMT_BI +  Ttl_All_BUD_AMT_BI
			Ttl_All_BUD_AMT_BI					= 0
			Grnd_Ttl_All_BUD_AMT_SBI			= Grnd_Ttl_All_BUD_AMT_SBI +  Ttl_All_BUD_AMT_SBI
			Ttl_All_BUD_AMT_SBI					= 0
			
			Grnd_Ttl_ACH_BI				= Grnd_Ttl_ACH_BI + Ttl_ACH_BI
			Ttl_ACH_BI					= 0
			Grnd_Ttl_ACH_SBI			= Grnd_Ttl_ACH_SBI + Ttl_ACH_SBI
			Ttl_ACH_SBI					= 0
			
			Grnd_Ttl_All_TBS_PROCESS_BI			= Grnd_Ttl_All_TBS_PROCESS_BI + Ttl_All_TBS_PROCESS_BI
			Ttl_All_TBS_PROCESS_BI				= 0
			Grnd_Ttl_All_TBS_PROCESS_SBI		= Grnd_Ttl_All_TBS_PROCESS_SBI + Ttl_All_TBS_PROCESS_SBI
			Ttl_All_TBS_PROCESS_SBI				= 0
			Grnd_Ttl_All_BUD_TBS_PROCESS_BI		= Grnd_Ttl_All_BUD_TBS_PROCESS_BI + Ttl_All_BUD_TBS_PROCESS_BI
			Ttl_All_BUD_TBS_PROCESS_BI			= 0
			Grnd_Ttl_All_BUD_TBS_PROCESS_SBI	= Grnd_Ttl_All_BUD_TBS_PROCESS_SBI + Ttl_All_BUD_TBS_PROCESS_SBI
			Ttl_All_BUD_TBS_PROCESS_SBI			= 0
			Grnd_Ttl_All_CPO_BI					= Grnd_Ttl_All_CPO_BI + Ttl_All_CPO_BI
			Ttl_All_CPO_BI						= 0
			Grnd_Ttl_All_CPO_SBI				= Grnd_Ttl_All_CPO_SBI + Ttl_All_CPO_SBI
			Ttl_All_CPO_SBI						= 0
			Grnd_Ttl_All_BUD_CPO_BI				= Grnd_Ttl_All_BUD_CPO_BI + Ttl_All_BUD_CPO_BI
			Ttl_All_BUD_CPO_BI					= 0
			Grnd_Ttl_All_BUD_CPO_SBI			= Grnd_Ttl_All_BUD_CPO_SBI + Ttl_All_BUD_CPO_SBI
			Ttl_All_BUD_CPO_SBI					= 0
		Else
			NO = NO+ 1
		end if
		rsEPS.MoveNext
		wend
		%>
		<tr>
			<td align="right" colspan="2" style="background-color:#FFFF00;"><b>Grand Total</b></td>
			<%
				IF Grnd_Ttl_BUD_TROUGHPUT_BI <> 0 THEN
					Grnd_Ttl_ACH_TROUGHPUT_BI = (Grnd_Ttl_TROUGHPUT_BI / Grnd_Ttl_BUD_TROUGHPUT_BI) * 100
				ELSE
					Grnd_Ttl_ACH_TROUGHPUT_BI = 0
				END IF
				
				IF Grnd_Ttl_BUD_TROUGHPUT_SBI <> 0 THEN
					Grnd_Ttl_ACH_TROUGHPUT_SBI = (Grnd_Ttl_TROUGHPUT_SBI / Grnd_Ttl_BUD_TROUGHPUT_SBI) * 100
				ELSE
					Grnd_Ttl_ACH_TROUGHPUT_SBI = 0
				END IF
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TROUGHPUT_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TROUGHPUT_SBI, 0)%></b></td>
			<%
				IF Grnd_Ttl_All_TBS_PROCESS_BI <> 0 THEN
					Grnd_Ttl_OER_BI		= (Grnd_Ttl_All_CPO_BI / Grnd_Ttl_All_TBS_PROCESS_BI) * 100
				ELSE
					Ttl_OER_BI		= 0 
				End IF
				
				IF Grnd_Ttl_All_TBS_PROCESS_SBI <> 0 THEN
					Grnd_Ttl_OER_SBI		= (Grnd_Ttl_All_CPO_SBI / Grnd_Ttl_All_TBS_PROCESS_SBI) * 100
				ELSE
					Grnd_Ttl_OER_SBI		= 0 
				End IF
				
				IF Grnd_Ttl_All_BUD_TBS_PROCESS_BI <> 0 THEN
					Grnd_Ttl_BUD_OER_BI		= (Grnd_Ttl_All_BUD_CPO_BI / Grnd_Ttl_All_BUD_TBS_PROCESS_BI) * 100
				ELSE
					Grnd_Ttl_BUD_OER_BI		= 0 
				End IF
				
				IF Grnd_Ttl_All_BUD_TBS_PROCESS_SBI <> 0 THEN
					Grnd_Ttl_BUD_OER_SBI		= (Grnd_Ttl_All_BUD_CPO_SBI / Grnd_Ttl_All_BUD_TBS_PROCESS_SBI) * 100
				ELSE
					Grnd_Ttl_BUD_OER_SBI		= 0 
				End IF
				
				IF Grnd_Ttl_BUD_OER_BI <> 0 THEN
					Grnd_Ttl_ACH_OER_BI		= (Grnd_Ttl_OER_BI / Grnd_Ttl_BUD_OER_BI)* 100
				ELSE
					Grnd_Ttl_ACH_OER_BI		= 0
				End If
				
				IF Grnd_Ttl_BUD_OER_SBI <> 0 THEN
					Grnd_Ttl_ACH_OER_SBI		= (Grnd_Ttl_OER_SBI / Grnd_Ttl_BUD_OER_SBI)* 100
				ELSE
					Grnd_Ttl_ACH_OER_SBI		= 0
				End If
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_OER_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_OER_SBI, 0)%></b></td>
			<%
				IF Grnd_Ttl_BUD_FFA_BI <> 0 THEN
					Grnd_Ttl_ACH_FFA_BI = (Grnd_Ttl_FFA_BI / Grnd_Ttl_BUD_FFA_BI) * 100 
				ELSE
					Grnd_Ttl_ACH_FFA_BI = 0
				END If
				
				IF Grnd_Ttl_BUD_FFA_SBI <> 0 THEN
					Grnd_Ttl_ACH_FFA_SBI = (Grnd_Ttl_FFA_SBI / Grnd_Ttl_BUD_FFA_SBI) * 100 
				ELSE
					Grnd_Ttl_ACH_FFA_SBI = 0
				END If
				
				if Grnd_Ttl_ACH_FFA_BI > 105 then
					Grnd_Ttl_CALC_ACH_FFA_BI = 85
				elseif (Grnd_Ttl_ACH_FFA_BI > 100 and Grnd_Ttl_ACH_FFA_BI < 105) then
					Grnd_Ttl_CALC_ACH_FFA_BI = 95
				elseif Grnd_Ttl_ACH_FFA_BI <= 100 then
					Grnd_Ttl_CALC_ACH_FFA_BI = 100
				end if
				
				if Grnd_Ttl_ACH_FFA_SBI > 105 then
					Grnd_Ttl_CALC_ACH_FFA_SBI = 85
				elseif (Grnd_Ttl_ACH_FFA_SBI > 100 and Grnd_Ttl_ACH_FFA_SBI < 105) then
					Grnd_Ttl_CALC_ACH_FFA_SBI = 95
				elseif Grnd_Ttl_ACH_FFA_SBI <= 100 then
					Grnd_Ttl_CALC_ACH_FFA_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_CALC_ACH_FFA_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_CALC_ACH_FFA_SBI, 0)%></b></td>
			<%
				IF Grnd_Ttl_All_TBS_PROCESS_BI <> 0 THEN
					Grnd_Ttl_AMT_BI = Grnd_Ttl_All_AMT_BI / Grnd_Ttl_All_TBS_PROCESS_BI
				ELSE 
					Grnd_Ttl_AMT_BI = 0 
				END IF
					
				IF Grnd_Ttl_All_TBS_PROCESS_SBI <> 0 THEN
					Grnd_Ttl_AMT_SBI = Grnd_Ttl_All_AMT_SBI / Grnd_Ttl_All_TBS_PROCESS_SBI
				ELSE 
					Grnd_Ttl_AMT_SBI = 0 
				END IF
				
				IF Grnd_Ttl_All_BUD_TBS_PROCESS_BI <> 0 THEN
					Grnd_Ttl_BUD_AMT_BI = Grnd_Ttl_All_BUD_AMT_BI / Grnd_Ttl_All_BUD_TBS_PROCESS_BI
				ELSE 
					Grnd_Ttl_BUD_AMT_BI = 0 
				END IF
					
				IF Grnd_Ttl_All_BUD_TBS_PROCESS_SBI <> 0 THEN
					Grnd_Ttl_BUD_AMT_SBI = Grnd_Ttl_All_BUD_AMT_SBI / Grnd_Ttl_All_BUD_TBS_PROCESS_SBI
				ELSE 
					Grnd_Ttl_BUD_AMT_SBI = 0 
				END IF
				
				IF Grnd_Ttl_BUD_AMT_BI <> 0 THEN
					Grnd_Ttl_ACH_AMT_BI = (Grnd_Ttl_AMT_BI / Grnd_Ttl_BUD_AMT_BI) * 100
				ELSE
					Grnd_Ttl_ACH_AMT_BI = 0
				END IF
				
				IF Grnd_Ttl_BUD_AMT_SBI <> 0 THEN
					Grnd_Ttl_ACH_AMT_SBI = (Grnd_Ttl_AMT_SBI / Grnd_Ttl_BUD_AMT_SBI) * 100
				ELSE
					Grnd_Ttl_ACH_AMT_SBI = 0
				END IF
				
				if Grnd_Ttl_ACH_AMT_BI < 90 OR Grnd_Ttl_ACH_AMT_BI > 110 then
					Grnd_Ttl_CALC_ACH_AMT_BI = 85
				elseif ((Grnd_Ttl_ACH_AMT_BI >= 90 and Grnd_Ttl_ACH_AMT_BI < 95) OR (Grnd_Ttl_ACH_AMT_BI > 105 and Grnd_Ttl_ACH_AMT_BI < 110)) then
					Grnd_Ttl_CALC_ACH_AMT_BI = 95
				elseif Grnd_Ttl_ACH_AMT_BI <= 95 and Grnd_Ttl_ACH_AMT_BI <= 105 then
					Grnd_Ttl_CALC_ACH_AMT_BI = 100
				end if
			
				if Grnd_Ttl_ACH_AMT_SBI < 90 OR Grnd_Ttl_ACH_AMT_SBI > 110 then
					Grnd_Ttl_CALC_ACH_AMT_SBI = 85
				elseif ((Grnd_Ttl_ACH_AMT_SBI >= 90 and Grnd_Ttl_ACH_AMT_SBI < 95) OR (Grnd_Ttl_ACH_AMT_SBI > 105 and Grnd_Ttl_ACH_AMT_SBI < 110)) then
					Grnd_Ttl_CALC_ACH_AMT_SBI = 95
				elseif Grnd_Ttl_ACH_AMT_SBI <= 95 and Grnd_Ttl_ACH_AMT_SBI <= 105 then
					Grnd_Ttl_CALC_ACH_AMT_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_CALC_ACH_AMT_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_CALC_ACH_AMT_SBI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_SBI, 0)%></b></td>
		</tr>
		<%
		end if
		%>
	</tbody
</table>
</form>
</div>
</div>
</body>
</html>