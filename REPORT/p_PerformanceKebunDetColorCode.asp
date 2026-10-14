<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	' Query Filter Region 
	Dim rsEPS,strMain, strSqlWilRow,rsSqlWilRow, strSqlRow,rsSqlRow, strsql,rsCount,visibleR,nSpanTotal
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth, v_txttodaydate, NO, cClass,NOWil
	Dim v_txtRegion, txtyear, txtmonth
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
		 txtRegion = rsCount("fccode")
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
	v_txtRegion  	= Request("txtRegion")
	nSpanTotal 		= 8
	
	if len(trim(request("txtYear"))) = 0 then
			v_txtyear = year(Now)	 
		end if

		if len(trim(request("txtmonth"))) = 0 then
			v_txtmonth = Month(Now)	 
	end if
		
	Dim wherecons, selectCons, orderCons
	wherecons = "where region is not null"
	
	if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
		if v_txtRegion <> "--ALL--" then
			whereCons = whereCons & " And  region = '" & v_txtRegion & "' "
		end if
			selectCons = selectCons & " ,region"
			orderCons  = orderCons & ", REGION_API.get_OrderData(region) "
		else
			nSpanTotal = nSpanTotal - 1 
	end if
	
	strMain = "select CASE COMPANY WHEN 'PT.THIP' THEN 'THIP'  WHEN 'PT.CRS' THEN 'CRS' " & _
			  "ELSE 'AMS GANDA' END WILAYAH," &_
			  "CASE COMPANY WHEN 'PT.THIP' THEN 1 WHEN 'PT.CRS' THEN 3 " &_
			  "ELSE 2 END URUTAN, X.* from MGT_MONTHLY_RPT_EST_TAB X "& whereCons &" and year= '"& v_txtyear &"' and month='"& v_txtmonth &"' order by URUTAN,region asc"
	Set rsEPS = Server.CreateObject("ADODB.Recordset")
	rsEPS.Open strMain , connstr, adOpenStatic
%>
<html>
<head>
<title>ePlaS Reporting - Detail Performance Kebun</title>
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
<form name="frmEntry" method="post" action="p_PerformanceKebunDetColorCode.asp">
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
		<th align="left">Region</th>
					<td>
						<select name="txtRegion">
								<option value="--SUMMARY--" <%if request("v_txtRegion")="--SUMMARY--" then%> selected <%End if%>>--SUMMARY--</option>
								<%
								Dim rsRegion
								Set rsRegion = Server.CreateObject("ADODB.Recordset")
								rsRegion.CursorLocation = adUseClient
								strsql =  " SELECT '--ALL--' as fccode from DUAL UNION " & _
								          " SELECT DISTINCT a.region fccode  " & _
										  " FROM Businessunit a, region b  " & _
										  " WHERE a.region = b.fccode  " & _
										  " AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', a.fccode) =  'TRUE'  " 


								rsRegion.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsRegion.eof then
									DO while not rsRegion.eof
										if v_txtRegion = rsRegion("fccode") then
										%>			
											<option value="<%=rsRegion("fccode")%>" Selected><%=rsRegion("fccode")%></option>
										<% Else %>	
											<option value="<%=rsRegion("fccode")%>"><%=rsRegion("fccode")%></option>
										<% End If
										
										rsRegion.movenext
									loop
								End if
								rsRegion.Close
								set rsRegion = Nothing
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
				<h2>Laporan Detail Performance Kebun</h2>
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
		<!--<th rowspan="3" style="background-color:#00A300;color:#fff;">X</th>-->
		<th rowspan="3" style="background-color:#00A300;color:#fff;">WIL</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">EAST</th>
		<th colspan="6" style="background-color:#00A300;color:#fff;">PRODUKSI(Ton)</th>
		<th colspan="6" style="background-color:#00A300;color:#fff;">TANAM BARU(Ha)</th>
        <th colspan="8" style="background-color:#00A300;color:#fff;">BIAYA PANEN(Rp/Kg)</th>
		<th colspan="8" style="background-color:#00A300;color:#fff;">BIAYA PEMELIHARAAN TM (Rp/Kg)</th>
		<th colspan="8" style="background-color:#00A300;color:#fff;">BIAYA PEMELIHARAAN TBM (Rp/Ha)</th>
	</tr>
	<tr>
		<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PENCAPAIAN</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PENCAPAIAN</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">VARIANCE</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PENCAPAIAN</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">VARIANCE</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PENCAPAIAN</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">AKTUAL</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">VARIANCE</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PENCAPAIAN</th>
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
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
		<th  style="background-color:#00A300;color:#fff;">BI</th>
		<th  style="background-color:#00A300;color:#fff;">SBI</th>
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
		Dim Ttl_PROD_BI, Ttl_PROD_SBI, Ttl_BUD_PROD_BI, Ttl_BUD_PROD_SBI, Ttl_ACH_PROD_BI, Ttl_ACH_PROD_SBI
		Dim Ttl_NP_BI, Ttl_NP_SBI, Ttl_BUD_NP_BI, Ttl_BUD_NP_SBI, Ttl_ACH_NP_BI, Ttl_ACH_NP_SBI 
		Dim Ttl_TM_PNN_BI, Ttl_TM_PNN_SBI, Ttl_BUD_TM_PNN_BI, Ttl_BUD_TM_PNN_SBI, Ttl_VAR_TM_PNN_BI, Ttl_VAR_TM_PNN_SBI, Ttl_ACH_TM_PNN_BI, Ttl_ACH_TM_PNN_SBI
		Dim Ttl_TM_PML_BI, Ttl_TM_PML_SBI, Ttl_BUD_TM_PML_BI, Ttl_BUD_TM_PML_SBI, Ttl_VAR_TM_PML_BI, Ttl_VAR_TM_PML_SBI, Ttl_ACH_TM_PML_BI, Ttl_ACH_TM_PML_SBI
		Dim Ttl_TBM_PML_BI, Ttl_TBM_PML_SBI, Ttl_BUD_TBM_PML_BI, Ttl_BUD_TBM_PML_SBI, Ttl_VAR_TBM_PML_BI, Ttl_VAR_TBM_PML_SBI, Ttl_ACH_TBM_PML_BI, Ttl_ACH_TBM_PML_SBI
		
		Dim Ttl_PT_PROD_BI, Ttl_PT_PROD_SBI, Ttl_PT_BUD_PROD_BI, Ttl_PT_BUD_PROD_SBI, Ttl_PT_ACH_PROD_BI, Ttl_PT_ACH_PROD_SBI
		Dim Ttl_PT_NP_BI, Ttl_PT_NP_SBI, Ttl_PT_BUD_NP_BI, Ttl_PT_BUD_NP_SBI, Ttl_PT_ACH_NP_BI, Ttl_PT_ACH_NP_SBI 
		Dim Ttl_PT_TM_PNN_BI, Ttl_PT_TM_PNN_SBI, Ttl_PT_BUD_TM_PNN_BI, Ttl_PT_BUD_TM_PNN_SBI, Ttl_PT_VAR_TM_PNN_BI, Ttl_PT_VAR_TM_PNN_SBI, Ttl_PT_ACH_TM_PNN_BI, Ttl_PT_ACH_TM_PNN_SBI
		Dim Ttl_PT_TM_PML_BI, Ttl_PT_TM_PML_SBI, Ttl_PT_BUD_TM_PML_BI, Ttl_PT_BUD_TM_PML_SBI, Ttl_PT_VAR_TM_PML_BI, Ttl_PT_VAR_TM_PML_SBI, Ttl_PT_ACH_TM_PML_BI, Ttl_PT_ACH_TM_PML_SBI
		Dim Ttl_PT_TBM_PML_BI, Ttl_PT_TBM_PML_SBI, Ttl_PT_BUD_TBM_PML_BI, Ttl_PT_BUD_TBM_PML_SBI, Ttl_PT_VAR_TBM_PML_BI, Ttl_PT_VAR_TBM_PML_SBI, Ttl_PT_ACH_TBM_PML_BI, Ttl_PT_ACH_TBM_PML_SBI
		
		DIm Grnd_Ttl_PROD_BI, Grnd_Ttl_PROD_SBI, Grnd_Ttl_BUD_PROD_BI, Grnd_Ttl_BUD_PROD_SBI, Grnd_Ttl_ACH_PROD_BI, Grnd_Ttl_ACH_PROD_SBI
		Dim Grnd_Ttl_NP_BI, Grnd_Ttl_NP_SBI, Grnd_Ttl_BUD_NP_BI, Grnd_Ttl_BUD_NP_SBI, Grnd_Ttl_ACH_NP_BI, Grnd_Ttl_ACH_NP_SBI 
		Dim Grnd_Ttl_TM_PNN_BI, Grnd_Ttl_TM_PNN_SBI, Grnd_Ttl_BUD_TM_PNN_BI, Grnd_Ttl_BUD_TM_PNN_SBI, Grnd_Ttl_VAR_TM_PNN_BI, Grnd_Ttl_VAR_TM_PNN_SBI, Grnd_Ttl_ACH_TM_PNN_BI, Grnd_Ttl_ACH_TM_PNN_SBI
		Dim Grnd_Ttl_TM_PML_BI, Grnd_Ttl_TM_PML_SBI, Grnd_Ttl_BUD_TM_PML_BI, Grnd_Ttl_BUD_TM_PML_SBI, Grnd_Ttl_VAR_TM_PML_BI, Grnd_Ttl_VAR_TM_PML_SBI, Grnd_Ttl_ACH_TM_PML_BI, Grnd_Ttl_ACH_TM_PML_SBI
		Dim Grnd_Ttl_TBM_PML_BI, Grnd_Ttl_TBM_PML_SBI, Grnd_Ttl_BUD_TBM_PML_BI, Grnd_Ttl_BUD_TBM_PML_SBI, Grnd_Ttl_VAR_TBM_PML_BI, Grnd_Ttl_VAR_TBM_PML_SBI, Grnd_Ttl_ACH_TBM_PML_BI, Grnd_Ttl_ACH_TBM_PML_SBI
		if not rsEPS.EOF then
		NOWil = 1
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
		Dim WhereConss
		WhereConss = "where region is not null"
		if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			if v_txtRegion <> "--ALL--" then
				WhereConss = WhereConss & " And region = '" & v_txtRegion & "' "
			end if
		end if
		
			strSqlWilRow = "Select Count (RegionCount) RowsCountWill from (Select Count(Region) RegionCount " &_
							"from MGT_MONTHLY_RPT_EST_TAB " & WhereConss&" and  CASE COMPANY WHEN 'PT.THIP' THEN 'THIP' WHEN 'PT.CRS' THEN 'CRS' ELSE 'AMS GANDA'" &_ 
							" END = '"& rsEPS("WILAYAH") &"' and year= '"& v_txtyear &"' and month='"& v_txtmonth &"' Group By Region)"
			Set rsSqlWilRow = Server.CreateObject("ADODB.Recordset")
			rsSqlWilRow.Open strSqlWilRow ,connstr, adOpenStatic
			
			strSqlRow = "select count(*) RowsCount from MGT_MONTHLY_RPT_EST_TAB where "&_
			" CASE COMPANY WHEN 'PT.THIP' THEN 'THIP'  WHEN 'PT.CRS' THEN 'CRS' ELSE 'AMS GANDA' END"&_ 
			"= '"& rsEPS("WILAYAH") &"' and region= '"& rsEPS("REGION") &"' and year= '"& v_txtyear &"' and month='"& v_txtmonth &"'"
			Set rsSqlRow = Server.CreateObject("ADODB.Recordset")
			rsSqlRow.Open strSqlRow ,connstr, adOpenStatic			
			if NO = 1 then
		%>			
			<!--<td align="center" Rowspan="<%=CDbl(rsSqlRow("RowsCount"))+1 %>"><%=rsEPS("WILAYAH")%></td>-->
			<td align="center" Rowspan="<%=CDbl(rsSqlRow("RowsCount"))+1 %>"><%=rsEPS("REGION")%></td>
		<%
			end if
			
			Dim PROD_BI,PROD_SBI,BUD_PROD_BI,BUD_PROD_SBI
				PROD_BI	= rsEPS("PROD_BI") / 1000
				PROD_SBI = rsEPS("PROD_SBI") / 1000
				BUD_PROD_BI = rsEPS("BUD_PROD_BI") / 1000
				BUD_PROD_SBI = rsEPS("BUD_PROD_SBI") / 1000
		%>
			<td align="center"><%=rsEPS("FCBA")%></td>
			<td align="right"><%=FormatNumber(PROD_BI,0)%></td>
			<td align="right"><%=FormatNumber(PROD_SBI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_PROD_BI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_PROD_SBI,0)%></td>
			<%
			if rsEPS("ACH_PROD_BI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_PROD_BI") > 85 and rsEPS("ACH_PROD_BI") < 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_PROD_BI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>;"><%=FormatNumber(rsEPS("ACH_PROD_BI"),0)%></td>
			<%
			if rsEPS("ACH_PROD_SBI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_PROD_SBI") > 85 and rsEPS("ACH_PROD_SBI") < 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_PROD_SBI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(rsEPS("ACH_PROD_SBI"),0)%></td>

			<td align="right"><%=FormatNumber(rsEPS("NP_BI"),0)%></td>
			<td align="right"><%=FormatNumber(rsEPS("NP_SBI"),0)%></td>
			<td align="right"><%=FormatNumber(rsEPS("BUD_NP_BI"),0)%></td>
			<td align="right"><%=FormatNumber(rsEPS("BUD_NP_SBI"),0)%></td>
			<%
			if rsEPS("ACH_NP_BI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_NP_BI") > 85 and rsEPS("ACH_NP_BI") < 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_NP_BI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(rsEPS("ACH_NP_BI"),0)%></td>
			<%
			if rsEPS("ACH_NP_SBI") < 85 then
				ColorAch = "#FF0000"
			elseif rsEPS("ACH_NP_SBI") > 85 and rsEPS("ACH_NP_SBI") < 95 then
				ColorAch = "#FFFF00"
			elseif rsEPS("ACH_NP_SBI") > 95 then
				ColorAch = "#00FF00"
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(rsEPS("ACH_NP_SBI"),0)%></td>
			
			<%
				Dim TM_PNN_BI,TM_PNN_SBI,BUD_TM_PNN_BI,BUD_TM_PNN_SBI
				IF (PROD_BI <> 0)then
					TM_PNN_BI		= CDbl(rsEPS("TM_PNN_BI"))/PROD_BI/1000
				else
					TM_PNN_BI 		= 0
				end if
				
				IF (PROD_SBI <> 0)then
					TM_PNN_SBI		= CDbl(rsEPS("TM_PNN_SBI"))/PROD_SBI/1000
				else
					TM_PNN_SBI		= 0
				end if
				
				IF (BUD_PROD_BI <> 0)then
					BUD_TM_PNN_BI	= CDbl(rsEPS("BUD_TM_PNN_BI"))/BUD_PROD_BI/1000
				else
					BUD_TM_PNN_BI	= 0
				end if
				
				IF (BUD_PROD_SBI <> 0)then
					BUD_TM_PNN_SBI	= CDbl(rsEPS("BUD_TM_PNN_SBI"))/BUD_PROD_SBI/1000
				else
					BUD_TM_PNN_SBI	= 0
				end if
			%>
			<td align="right"><%=FormatNumber(TM_PNN_BI,0)%></td>
			<td align="right"><%=FormatNumber(TM_PNN_SBI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_TM_PNN_BI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_TM_PNN_SBI,0)%></td>
			<%
				Dim VAR_TM_PNN_BI
				IF BUD_TM_PNN_BI = 0 then
					VAR_TM_PNN_BI = 0
				else
					VAR_TM_PNN_BI = TM_PNN_BI / (BUD_TM_PNN_BI/ 100)
				end if
			%>
			<td align="right"><%=FormatNumber(VAR_TM_PNN_BI,0)%></td>
			<%
				Dim VAR_TM_PNN_SBI
				IF BUD_TM_PNN_SBI = 0 then
					VAR_TM_PNN_SBI = 0
				else
					VAR_TM_PNN_SBI = TM_PNN_SBI / (BUD_TM_PNN_SBI/ 100)
				end if
			%>
			<td align="right"><%=FormatNumber(VAR_TM_PNN_SBI,0)%></td>
			<%
			Dim ACH_TM_PNN_BI
			if VAR_TM_PNN_BI < 90 OR VAR_TM_PNN_BI > 110 then
				ColorAch = "#FF0000"
				ACH_TM_PNN_BI = 85
			elseif (VAR_TM_PNN_BI > 89 and VAR_TM_PNN_BI < 95) OR  (VAR_TM_PNN_BI > 104 and VAR_TM_PNN_BI <= 110) then
				ColorAch = "#FFFF00"
				ACH_TM_PNN_BI = 95
			elseif VAR_TM_PNN_BI > 94 and VAR_TM_PNN_BI < 106 then
				ColorAch = "#00FF00"
				ACH_TM_PNN_BI = 100
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(ACH_TM_PNN_BI,0)%></td>
			<%
			Dim ACH_TM_PNN_SBI
			if VAR_TM_PNN_SBI < 90 OR VAR_TM_PNN_SBI > 110 then
				ColorAch = "#FF0000"
				ACH_TM_PNN_SBI = 85
			elseif (VAR_TM_PNN_SBI > 89 and VAR_TM_PNN_SBI< 95) OR  (VAR_TM_PNN_SBI > 104 and VAR_TM_PNN_SBI <= 110) then
				ColorAch = "#FFFF00"
				ACH_TM_PNN_SBI = 95
			elseif VAR_TM_PNN_SBI > 94 and VAR_TM_PNN_SBI < 106 then
				ColorAch = "#00FF00"
				ACH_TM_PNN_SBI = 100
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(ACH_TM_PNN_SBI,0)%></td>
			<%
				Dim TM_PML_BI,TM_PML_SBI,BUD_TM_PML_BI,BUD_TM_PML_SBI
				IF PROD_BI <> 0 then
					TM_PML_BI		= CDbl(rsEPS("TM_PML_BI"))/PROD_BI/1000
				else 
					TM_PML_BI		= 0
				end if
				
				IF PROD_SBI <> 0 then
					TM_PML_SBI		= CDbl(rsEPS("TM_PML_SBI"))/PROD_SBI/1000
				else
					TM_PML_SBI		= 0
				end if 
				
				IF BUD_PROD_BI <> 0 then
					BUD_TM_PML_BI	= CDbl(rsEPS("BUD_TM_PML_BI"))/BUD_PROD_BI/1000
				else
					BUD_TM_PML_BI	= 0
				end if 
				
				IF BUD_PROD_SBI <> 0 then
					BUD_TM_PML_SBI	= CDbl(rsEPS("BUD_TM_PML_SBI"))/BUD_PROD_SBI/1000
				else
					BUD_TM_PML_SBI	= 0
				end if 
			%>
			<td align="right"><%=FormatNumber(TM_PML_BI,0)%></td>
			<td align="right"><%=FormatNumber(TM_PML_SBI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_TM_PML_BI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_TM_PML_SBI,0)%></td>
			<%
				Dim VAR_TM_PML_BI
				IF BUD_TM_PML_BI = 0 then
					VAR_TM_PML_BI = 0
				else
					VAR_TM_PML_BI = TM_PML_BI / (BUD_TM_PML_BI/ 100)
				end if
			%>
			<td align="right"><%=FormatNumber(VAR_TM_PML_BI,0)%></td>
			<%
				Dim VAR_TM_PML_SBI
				IF BUD_TM_PML_SBI = 0 then
					VAR_TM_PML_SBI = 0
				else
					VAR_TM_PML_SBI = TM_PML_SBI / (BUD_TM_PML_SBI/ 100)
				end if
			%>
			<td align="right"><%=FormatNumber(VAR_TM_PML_SBI,0)%></td>
			<%
			Dim ACH_TM_PML_BI
			if VAR_TM_PML_BI < 90 OR VAR_TM_PML_BI > 110 then
				ColorAch = "#FF0000"
				ACH_TM_PML_BI = 85
			elseif (VAR_TM_PML_BI > 89 and VAR_TM_PML_BI < 95) OR  (VAR_TM_PML_BI > 104 and VAR_TM_PML_BI <= 110) then
				ColorAch = "#FFFF00"
				ACH_TM_PML_BI = 95
			elseif VAR_TM_PML_BI > 94 and VAR_TM_PML_BI < 106 then
				ColorAch = "#00FF00"
				ACH_TM_PML_BI = 100
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(ACH_TM_PML_BI,0)%></td>
			<%
			Dim ACH_TM_PML_SBI
			if VAR_TM_PML_SBI < 90 OR VAR_TM_PML_SBI > 110 then
				ColorAch = "#FF0000"
				ACH_TM_PML_SBI = 85
			elseif (VAR_TM_PML_SBI > 89 and VAR_TM_PML_SBI < 95) OR  (VAR_TM_PML_SBI > 104 and VAR_TM_PML_SBI <= 110) then
				ColorAch = "#FFFF00"
				ACH_TM_PML_SBI = 95
			elseif VAR_TM_PML_SBI > 94 and VAR_TM_PML_SBI < 106 then
				ColorAch = "#00FF00"
				ACH_TM_PML_SBI = 100
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(ACH_TM_PML_SBI,0)%></td>
			
			<%
				Dim TBM_PML_BI, TBM_PML_SBI, BUD_TBM_PML_BI,BUD_TBM_PML_SBI
				IF CDbl(rsEPS("HA_TBM")) <> 0 then
					TBM_PML_BI		= CDbl(rsEPS("TBM_PML_BI"))/CDbl(rsEPS("HA_TBM"))
					TBM_PML_SBI		= CDbl(rsEPS("TBM_PML_SBI"))/CDbl(rsEPS("HA_TBM"))
					BUD_TBM_PML_BI	= CDbl(rsEPS("BUD_TBM_PML_BI"))/CDbl(rsEPS("HA_TBM"))
					BUD_TBM_PML_SBI	= CDbl(rsEPS("BUD_TBM_PML_SBI"))/CDbl(rsEPS("HA_TBM"))
				else
					TBM_PML_BI		= 0
					TBM_PML_SBI		= 0
					BUD_TBM_PML_BI	= 0
					BUD_TBM_PML_SBI	= 0
				end if
			%>
			<td align="right"><%=FormatNumber(TBM_PML_BI,0)%></td>
			<td align="right"><%=FormatNumber(TBM_PML_SBI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_TBM_PML_BI,0)%></td>
			<td align="right"><%=FormatNumber(BUD_TBM_PML_SBI,0)%></td>
			<%
				Dim VAR_TBM_PML_BI
				IF BUD_TBM_PML_BI = 0 then
					VAR_TBM_PML_BI = 0
				else
					VAR_TBM_PML_BI = TBM_PML_BI / (BUD_TBM_PML_BI/ 100)
				end if
			%>
			<td align="right"><%=FormatNumber(VAR_TBM_PML_BI,0)%></td>
			<%
				Dim VAR_TBM_PML_SBI
				IF BUD_TBM_PML_SBI = 0 then
					VAR_TBM_PML_SBI = 0
				else
					VAR_TBM_PML_SBI = TBM_PML_SBI / (BUD_TBM_PML_SBI/ 100)
				end if
			%>
			<td align="right"><%=FormatNumber(VAR_TBM_PML_SBI,0)%></td>
			<%
			Dim ACH_TBM_PML_BI
			if VAR_TBM_PML_BI < 90 OR VAR_TBM_PML_BI> 110 then
				ColorAch = "#FF0000"
				ACH_TBM_PML_BI = 85
			elseif (VAR_TBM_PML_BI > 89 and VAR_TBM_PML_BI < 95) OR  (VAR_TBM_PML_BI > 104 and VAR_TBM_PML_BI <= 110) then
				ColorAch = "#FFFF00"
				ACH_TBM_PML_BI = 95
			elseif VAR_TBM_PML_BI > 94 and VAR_TBM_PML_BI< 106 then
				ColorAch = "#00FF00"
				ACH_TBM_PML_BI = 100
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(ACH_TBM_PML_BI,0)%></td>
			<%
			Dim ACH_TBM_PML_SBI
			if VAR_TBM_PML_SBI < 90 OR VAR_TBM_PML_SBI > 110 then
				ColorAch = "#FF0000"
				ACH_TBM_PML_SBI = 85
			elseif (VAR_TBM_PML_SBI > 89 and VAR_TBM_PML_SBI < 95) OR  (VAR_TBM_PML_SBI > 104 and VAR_TBM_PML_SBI <= 110) then
				ColorAch = "#FFFF00"
				ACH_TBM_PML_SBI = 95
			elseif VAR_TBM_PML_SBI > 94 and VAR_TBM_PML_SBI < 106 then
				ColorAch = "#00FF00"
				ACH_TBM_PML_SBI = 100
			end if
			%>
			<td align="right" style="background-color:<%=ColorAch%>"><%=FormatNumber(ACH_TBM_PML_SBI,0)%></td>
		</tr>
		<%
			Dim Ttl_All_TM_PNN_BI, Ttl_All_TM_PNN_SBI, Ttl_All_BUD_TM_PNN_BI,Ttl_All_BUD_TM_PNN_SBI
			Dim Ttl_All_TM_PML_BI, Ttl_All_TM_PML_SBI, Ttl_All_BUD_TM_PML_BI, Ttl_All_BUD_TM_PML_SBI
			Dim Ttl_All_TBM_PML_BI, Ttl_All_TBM_PML_SBI, Ttl_All_BUD_TBM_PML_BI, Ttl_All_BUD_TBM_PML_SBI
			Dim Ttl_All_HA_TM, Ttl_All_HA_TBM, Ttl_All_HA_NEWTBM

			Ttl_PROD_BI 			= Ttl_PROD_BI +  CDbl(rsEPS("PROD_BI"))
			Ttl_PROD_SBI 			= Ttl_PROD_SBI + CDbl(rsEPS("PROD_SBI"))
			Ttl_BUD_PROD_BI 		= Ttl_BUD_PROD_BI + CDbl(rsEPS("BUD_PROD_BI"))
			Ttl_BUD_PROD_SBI 		= Ttl_BUD_PROD_SBI + CDbl(rsEPS("BUD_PROD_SBI"))
			
			Ttl_NP_BI				= Ttl_NP_BI + rsEPS("NP_BI")
			Ttl_NP_SBI				= Ttl_NP_SBI + rsEPS("NP_SBI")
			Ttl_BUD_NP_BI			= Ttl_BUD_NP_BI + rsEPS("BUD_NP_BI")
			Ttl_BUD_NP_SBI			= Ttl_BUD_NP_SBI + rsEPS("BUD_NP_SBI")
			
			Ttl_All_TM_PNN_BI		= Ttl_All_TM_PNN_BI + CDbl(rsEPS("TM_PNN_BI"))
			Ttl_All_TM_PNN_SBI		= Ttl_All_TM_PNN_SBI + CDbl(rsEPS("TM_PNN_SBI"))
			Ttl_All_BUD_TM_PNN_BI	= Ttl_All_BUD_TM_PNN_BI + CDbl(rsEPS("BUD_TM_PNN_BI"))
			Ttl_All_BUD_TM_PNN_SBI	= Ttl_All_BUD_TM_PNN_SBI + CDbl(rsEPS("BUD_TM_PNN_SBI"))
			
			Ttl_All_TM_PML_BI		= Ttl_All_TM_PML_BI + CDbl(rsEPS("TM_PML_BI"))
			Ttl_All_TM_PML_SBI		= Ttl_All_TM_PML_SBI + CDbl(rsEPS("TM_PML_SBI"))
			Ttl_All_BUD_TM_PML_BI	= Ttl_All_BUD_TM_PML_BI + CDbl(rsEPS("BUD_TM_PML_BI"))
			Ttl_All_BUD_TM_PML_SBI	= Ttl_All_BUD_TM_PML_SBI + CDbl(rsEPS("BUD_TM_PML_SBI"))
			
			Ttl_All_TBM_PML_BI		= Ttl_All_TBM_PML_BI + CDbl(rsEPS("TBM_PML_BI"))
			Ttl_All_TBM_PML_SBI		= Ttl_All_TBM_PML_SBI + CDbl(rsEPS("TBM_PML_SBI"))
			Ttl_All_BUD_TBM_PML_BI	= Ttl_All_BUD_TBM_PML_BI + CDbl(rsEPS("BUD_TBM_PML_BI"))
			Ttl_All_BUD_TBM_PML_SBI	= Ttl_All_BUD_TBM_PML_SBI + CDbl(rsEPS("BUD_TBM_PML_SBI"))
			
			Ttl_All_HA_TM			= Ttl_All_HA_TM + CDbl(rsEPS("HA_TM"))
			Ttl_All_HA_TBM			= Ttl_All_HA_TBM + CDbl(rsEPS("HA_TBM"))
			Ttl_All_HA_NEWTBM		= Ttl_All_HA_NEWTBM + CDbl(rsEPS("HA_NEWTBM"))
		if NO = CDbl(rsSqlRow("RowsCount"))then
			NO = 1
		%>
		<tr>
			<td align="right" style="background-color:#FFA500;"><b>Total</b></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_PROD_BI / 1000,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_PROD_SBI / 1000,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_PROD_BI / 1000,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_PROD_SBI / 1000,0)%></td>
			<%
				if Ttl_PROD_BI  = 0 OR  Ttl_BUD_PROD_BI = 0 then 
					Ttl_ACH_PROD_BI = 0
				else
					Ttl_ACH_PROD_BI = Ttl_PROD_BI / Ttl_BUD_PROD_BI * 100					
				end if
				
				if Ttl_PROD_SBI = 0 OR  Ttl_BUD_PROD_SBI = 0 then 
					Ttl_ACH_PROD_SBI = 0
				else
					Ttl_ACH_PROD_SBI = Ttl_PROD_SBI / Ttl_BUD_PROD_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_PROD_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_PROD_SBI,0)%></td>
			
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_NP_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_NP_SBI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_NP_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_NP_SBI,0)%></td>
			<%
				if Ttl_NP_BI = 0 OR  Ttl_BUD_NP_BI = 0 then 
					Ttl_ACH_NP_BI = 0
				else
					Ttl_ACH_NP_BI = Ttl_NP_BI / Ttl_BUD_NP_BI * 100
				end if
				
				if Ttl_NP_SBI = 0 OR  Ttl_BUD_NP_SBI = 0 then 
					Ttl_ACH_NP_SBI = 0
				else
					Ttl_ACH_NP_SBI = Ttl_NP_SBI / Ttl_BUD_NP_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_NP_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_NP_SBI,0)%></td>
			<%
				if Ttl_PROD_BI <> 0 then
					Ttl_TM_PNN_BI 		= Ttl_All_TM_PNN_BI / Ttl_PROD_BI
				else 
					Ttl_TM_PNN_BI 		= 0
				end if 
				
				if Ttl_PROD_SBI <> 0 then
					Ttl_TM_PNN_SBI 		= Ttl_All_TM_PNN_SBI / Ttl_PROD_SBI
				else
					Ttl_TM_PNN_SBI		= 0
				end if
				
				if Ttl_BUD_PROD_BI  <> 0 then 
					Ttl_BUD_TM_PNN_BI	= Ttl_All_BUD_TM_PNN_BI / Ttl_BUD_PROD_BI
				else 
					Ttl_BUD_TM_PNN_BI 	= 0
				end if
				
				if Ttl_BUD_PROD_SBI <> 0 then
					Ttl_BUD_TM_PNN_SBI	= Ttl_All_BUD_TM_PNN_SBI / Ttl_BUD_PROD_SBI
				else
					Ttl_BUD_TM_PNN_SBI	= 0
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_TM_PNN_SBI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_TM_PNN_SBI,0)%></td>
			<%
				if Ttl_TM_PNN_BI = 0 OR  Ttl_BUD_TM_PNN_BI = 0 then 
					Ttl_VAR_TM_PNN_BI = 0
				else
					Ttl_VAR_TM_PNN_BI = Ttl_TM_PNN_BI / Ttl_BUD_TM_PNN_BI * 100
				end if
				
				if Ttl_TM_PNN_SBI = 0 OR  Ttl_BUD_TM_PNN_SBI = 0 then 
					Ttl_VAR_TM_PNN_SBI = 0
				else
					Ttl_VAR_TM_PNN_SBI = Ttl_TM_PNN_SBI / Ttl_BUD_TM_PNN_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_VAR_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_VAR_TM_PNN_SBI,0)%></td>
			<%
				if Ttl_VAR_TM_PNN_BI < 90 OR Ttl_VAR_TM_PNN_BI> 110 then
					Ttl_ACH_TM_PNN_BI = 85
				elseif (Ttl_VAR_TM_PNN_BI > 89 and Ttl_VAR_TM_PNN_BI < 95) OR  (Ttl_VAR_TM_PNN_BI > 104 and Ttl_VAR_TM_PNN_BI <= 110) then
					Ttl_ACH_TM_PNN_BI = 95
				elseif Ttl_VAR_TM_PNN_BI > 94 and Ttl_VAR_TM_PNN_BI< 106 then
					Ttl_ACH_TM_PNN_BI = 100
				end if
				
				if Ttl_VAR_TM_PNN_SBI < 90 OR Ttl_VAR_TM_PNN_SBI> 110 then
					Ttl_ACH_TM_PNN_SBI = 85
				elseif (Ttl_VAR_TM_PNN_SBI > 89 and Ttl_VAR_TM_PNN_SBI < 95) OR  (Ttl_VAR_TM_PNN_SBI > 104 and Ttl_VAR_TM_PNN_SBI <= 110) then
					Ttl_ACH_TM_PNN_SBI = 95
				elseif Ttl_VAR_TM_PNN_SBI > 94 and Ttl_VAR_TM_PNN_SBI< 106 then
					Ttl_ACH_TM_PNN_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_TM_PNN_SBI,0)%></td>
			<%
				if Ttl_PROD_BI <> 0 then
					Ttl_TM_PML_BI		= Ttl_All_TM_PML_BI / Ttl_PROD_BI
				else
					Ttl_TM_PML_BI		= 0
				end if
				
				if Ttl_PROD_SBI <> 0 then
					Ttl_TM_PML_SBI		= Ttl_All_TM_PML_SBI / Ttl_PROD_SBI
				else
					Ttl_TM_PML_SBI		= 0
				end if 
				
				if Ttl_BUD_PROD_BI <> 0 then
					Ttl_BUD_TM_PML_BI	= Ttl_All_BUD_TM_PML_BI / Ttl_BUD_PROD_BI
				else
					Ttl_BUD_TM_PML_BI	= 0
				end if
				
				if Ttl_BUD_PROD_SBI <> 0 then
					Ttl_BUD_TM_PML_SBI	= Ttl_All_BUD_TM_PML_SBI / Ttl_BUD_PROD_SBI
				else
					Ttl_BUD_TM_PML_SBI	= 0
				end if
				
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_TM_PML_SBI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_TM_PML_SBI,0)%></td>
			<%
				if Ttl_TM_PML_BI = 0 OR  Ttl_BUD_TM_PML_BI = 0 then 
					Ttl_VAR_TM_PML_BI = 0
				else
					Ttl_VAR_TM_PML_BI = Ttl_TM_PML_BI / Ttl_BUD_TM_PML_BI * 100
				end if
				
				if Ttl_TM_PML_SBI = 0 OR  Ttl_BUD_TM_PML_SBI = 0 then 
					Ttl_VAR_TM_PML_SBI = 0
				else
					Ttl_VAR_TM_PML_SBI = Ttl_TM_PML_SBI / Ttl_BUD_TM_PML_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_VAR_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_VAR_TM_PML_SBI,0)%></td>
			<%
			if Ttl_VAR_TM_PML_BI < 90 OR Ttl_VAR_TM_PML_BI> 110 then
				Ttl_ACH_TM_PML_BI = 85
			elseif (Ttl_VAR_TM_PML_BI > 89 and Ttl_VAR_TM_PML_BI < 95) OR  (Ttl_VAR_TM_PML_BI > 104 and Ttl_VAR_TM_PML_BI <= 110) then
				Ttl_ACH_TM_PML_BI = 95
			elseif Ttl_VAR_TM_PML_BI > 94 and Ttl_VAR_TM_PML_BI< 106 then
				Ttl_ACH_TM_PML_BI = 100
			end if
			
			if Ttl_VAR_TM_PML_SBI < 90 OR Ttl_VAR_TM_PML_SBI> 110 then
				Ttl_ACH_TM_PML_SBI = 85
			elseif (Ttl_VAR_TM_PML_SBI > 89 and Ttl_VAR_TM_PML_SBI < 95) OR  (Ttl_VAR_TM_PML_SBI > 104 and Ttl_VAR_TM_PML_SBI <= 110) then
				Ttl_ACH_TM_PML_SBI = 95
			elseif Ttl_VAR_TM_PML_SBI > 94 and Ttl_VAR_TM_PML_SBI< 106 then
				Ttl_ACH_TM_PML_SBI = 100
			end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_TM_PML_SBI,0)%></td>
			<%
				if Ttl_All_HA_TBM <> 0 then
					Ttl_TBM_PML_BI 		= Ttl_All_TBM_PML_BI / Ttl_All_HA_TBM
					Ttl_TBM_PML_SBI		= Ttl_All_TBM_PML_SBI / Ttl_All_HA_TBM
					Ttl_BUD_TBM_PML_BI	= Ttl_All_BUD_TBM_PML_BI / Ttl_All_HA_TBM
					Ttl_BUD_TBM_PML_SBI	= Ttl_All_BUD_TBM_PML_SBI / Ttl_All_HA_TBM
				else
					Ttl_TBM_PML_BI 		= 0
					Ttl_TBM_PML_SBI		= 0
					Ttl_BUD_TBM_PML_BI	= 0
					Ttl_BUD_TBM_PML_SBI	= 0
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_TBM_PML_SBI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_BUD_TBM_PML_SBI,0)%></td>
			<%
				if Ttl_TBM_PML_BI = 0 OR  Ttl_BUD_TBM_PML_BI = 0 then 
					Ttl_VAR_TBM_PML_BI = 0
				else
					Ttl_VAR_TBM_PML_BI = Ttl_TBM_PML_BI / Ttl_BUD_TBM_PML_BI * 100
				end if
				
				if Ttl_TBM_PML_SBI = 0 OR  Ttl_BUD_TBM_PML_SBI = 0 then 
					Ttl_VAR_TBM_PML_SBI = 0
				else
					Ttl_VAR_TBM_PML_SBI = Ttl_TBM_PML_SBI / Ttl_BUD_TBM_PML_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_VAR_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_VAR_TBM_PML_SBI,0)%></td>
			<%
			if Ttl_VAR_TBM_PML_BI < 90 OR Ttl_VAR_TBM_PML_BI> 110 then
				Ttl_ACH_TBM_PML_BI = 85
			elseif (Ttl_VAR_TBM_PML_BI > 89 and Ttl_VAR_TBM_PML_BI < 95) OR  (Ttl_VAR_TBM_PML_BI > 104 and Ttl_VAR_TBM_PML_BI <= 110) then
				Ttl_ACH_TBM_PML_BI = 95
			elseif Ttl_VAR_TBM_PML_BI > 94 and Ttl_VAR_TBM_PML_BI< 106 then
				Ttl_ACH_TBM_PML_BI = 100
			end if
			
			if Ttl_VAR_TBM_PML_SBI < 90 OR Ttl_VAR_TBM_PML_SBI> 110 then
				Ttl_ACH_TBM_PML_SBI = 85
			elseif (Ttl_VAR_TBM_PML_SBI > 89 and Ttl_VAR_TBM_PML_SBI < 95) OR  (Ttl_VAR_TBM_PML_SBI > 104 and Ttl_VAR_TBM_PML_SBI <= 110) then
				Ttl_ACH_TBM_PML_SBI = 95
			elseif Ttl_VAR_TBM_PML_SBI > 94 and Ttl_VAR_TBM_PML_SBI< 106 then
				Ttl_ACH_TBM_PML_SBI = 100
			end if
			%>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFA500;"><%=FormatNumber(Ttl_ACH_TBM_PML_SBI,0)%></td>
		</tr>
		<%
			Dim Ttl_All_PT_TM_PNN_BI, Ttl_All_PT_TM_PNN_SBI, Ttl_All_PT_BUD_TM_PNN_BI, Ttl_All_PT_BUD_TM_PNN_SBI
			Dim Ttl_All_PT_TM_PML_BI, Ttl_All_PT_TM_PML_SBI, Ttl_All_PT_BUD_TM_PML_BI, Ttl_All_PT_BUD_TM_PML_SBI
			Dim Ttl_All_PT_TBM_PML_BI, Ttl_All_PT_TBM_PML_SBI, Ttl_All_PT_BUD_TBM_PML_BI, Ttl_All_PT_BUD_TBM_PML_SBI
			Dim Ttl_All_PT_HA_TM, Ttl_All_PT_HA_TBM, Ttl_All_PT_HA_NEWTBM
			
			Ttl_PT_PROD_BI			= Ttl_PT_PROD_BI +  Ttl_PROD_BI
			Ttl_PT_PROD_SBI 		= Ttl_PT_PROD_SBI + Ttl_PROD_SBI 					
			Ttl_PT_BUD_PROD_BI 		= Ttl_PT_BUD_PROD_BI + Ttl_BUD_PROD_BI 			  
			Ttl_PT_BUD_PROD_SBI 	= Ttl_PT_BUD_PROD_SBI + Ttl_BUD_PROD_SBI 			
		  
			Ttl_PT_NP_BI 			= Ttl_PT_NP_BI + Ttl_NP_BI 						  
			Ttl_PT_NP_SBI 			= Ttl_PT_NP_SBI + Ttl_NP_SBI 						
			Ttl_PT_BUD_NP_BI 		= Ttl_PT_BUD_NP_BI + Ttl_BUD_NP_BI 				  
			Ttl_PT_BUD_NP_SBI 		= Ttl_PT_BUD_NP_SBI + Ttl_BUD_NP_SBI 				
			
			Ttl_All_PT_TM_PNN_BI 		= Ttl_All_PT_TM_PNN_BI + Ttl_All_TM_PNN_BI 				  
			Ttl_All_PT_TM_PNN_SBI 		= Ttl_All_PT_TM_PNN_SBI + Ttl_All_TM_PNN_SBI 				
			Ttl_All_PT_BUD_TM_PNN_BI 	= Ttl_All_PT_BUD_TM_PNN_BI + Ttl_All_BUD_TM_PNN_BI 		  
			Ttl_All_PT_BUD_TM_PNN_SBI 	= Ttl_All_PT_BUD_TM_PNN_SBI + Ttl_All_BUD_TM_PNN_SBI 	
			
			Ttl_All_PT_TM_PML_BI 		= Ttl_All_PT_TM_PML_BI + Ttl_All_TM_PML_BI 				  
			Ttl_All_PT_TM_PML_SBI 		= Ttl_All_PT_TM_PML_SBI + Ttl_All_TM_PML_SBI 				
			Ttl_All_PT_BUD_TM_PML_BI 	= Ttl_All_PT_BUD_TM_PML_BI + Ttl_All_BUD_TM_PML_BI 		  
			Ttl_All_PT_BUD_TM_PML_SBI 	= Ttl_All_PT_BUD_TM_PML_SBI + Ttl_All_BUD_TM_PML_SBI 		
			
			Ttl_All_PT_TBM_PML_BI 		= Ttl_All_PT_TBM_PML_BI + Ttl_All_TBM_PML_BI 				
			Ttl_All_PT_TBM_PML_SBI 		= Ttl_All_PT_TBM_PML_SBI + Ttl_All_TBM_PML_SBI 			  
			Ttl_All_PT_BUD_TBM_PML_BI 	= Ttl_All_PT_BUD_TBM_PML_BI + Ttl_All_BUD_TBM_PML_BI 		
			Ttl_All_PT_BUD_TBM_PML_SBI 	= Ttl_All_PT_BUD_TBM_PML_SBI + Ttl_All_BUD_TBM_PML_SBI 	  
			
			Ttl_All_PT_HA_TM			= Ttl_All_PT_HA_TM + Ttl_All_HA_TM
			Ttl_All_PT_HA_TBM			= Ttl_All_PT_HA_TBM + Ttl_All_HA_TBM
			Ttl_All_PT_HA_NEWTBM		= Ttl_All_PT_HA_NEWTBM + Ttl_All_HA_NEWTBM
			
			if NOWil = CDbl(rsSqlWilRow("RowsCountWill"))then
		%>
			<tr>
			<td align="right" Colspan="2" style="background-color:#FFFF00;"><b>Total <%=rsEPS("WILAYAH")%></b></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_PROD_BI / 1000,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_PROD_SBI / 1000,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_PROD_BI / 1000,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_PROD_SBI / 1000,0)%></td>
			<%
				if Ttl_PT_PROD_BI = 0 OR  Ttl_PT_BUD_PROD_BI = 0 then 
					Ttl_PT_ACH_PROD_BI = 0
				else
					Ttl_PT_ACH_PROD_BI = Ttl_PT_PROD_BI / Ttl_PT_BUD_PROD_BI * 100					
				end if
				
				if Ttl_PT_PROD_SBI = 0 OR  Ttl_PT_BUD_PROD_SBI = 0 then 
					Ttl_PT_ACH_PROD_SBI = 0
				else
					Ttl_PT_ACH_PROD_SBI = Ttl_PT_PROD_SBI / Ttl_PT_BUD_PROD_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_PROD_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_PROD_SBI,0)%></td>
			
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_NP_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_NP_SBI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_NP_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_NP_SBI,0)%></td>
			<%
				if Ttl_PT_NP_BI = 0 OR  Ttl_PT_BUD_NP_BI = 0 then 
					Ttl_PT_ACH_NP_BI = 0
				else
					Ttl_PT_ACH_NP_BI = Ttl_PT_NP_BI / Ttl_PT_BUD_NP_BI * 100					
				end if
				
				if Ttl_PT_NP_SBI = 0 OR  Ttl_PT_BUD_NP_SBI = 0 then 
					Ttl_PT_ACH_NP_SBI = 0
				else
					Ttl_PT_ACH_NP_SBI = Ttl_PT_NP_SBI / Ttl_PT_BUD_NP_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_NP_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_NP_SBI,0)%></td>
			<%
				if Ttl_PT_PROD_BI <> 0 then
					Ttl_PT_TM_PNN_BI 		= Ttl_All_PT_TM_PNN_BI / Ttl_PT_PROD_BI
				else 
					Ttl_PT_TM_PNN_BI 		= 0
				end if 
				
				if Ttl_PT_PROD_SBI <> 0 then
					Ttl_PT_TM_PNN_SBI 		= Ttl_All_PT_TM_PNN_SBI / Ttl_PT_PROD_SBI
				else
					Ttl_PT_TM_PNN_SBI		= 0
				end if
				
				if Ttl_PT_BUD_PROD_BI  <> 0 then 
					Ttl_PT_BUD_TM_PNN_BI	= Ttl_All_PT_BUD_TM_PNN_BI / Ttl_PT_BUD_PROD_BI
				else 
					Ttl_PT_BUD_TM_PNN_BI 	= 0
				end if
				
				if Ttl_PT_BUD_PROD_SBI <> 0 then
					Ttl_PT_BUD_TM_PNN_SBI	= Ttl_All_PT_BUD_TM_PNN_SBI / Ttl_PT_BUD_PROD_SBI
				else
					Ttl_PT_BUD_TM_PNN_SBI	= 0
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_TM_PNN_SBI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_TM_PNN_SBI,0)%></td>
			<%
				if Ttl_PT_TM_PNN_BI = 0 OR  Ttl_PT_BUD_TM_PNN_BI = 0 then 
					Ttl_PT_VAR_TM_PNN_BI = 0
				else
					Ttl_PT_VAR_TM_PNN_BI = Ttl_PT_TM_PNN_BI / Ttl_PT_BUD_TM_PNN_BI * 100
				end if
				
				if Ttl_PT_TM_PNN_SBI = 0 OR  Ttl_PT_BUD_TM_PNN_SBI = 0 then 
					Ttl_PT_VAR_TM_PNN_SBI = 0
				else
					Ttl_PT_VAR_TM_PNN_SBI = Ttl_PT_TM_PNN_SBI / Ttl_PT_BUD_TM_PNN_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_VAR_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_VAR_TM_PNN_SBI,0)%></td>
			<%
				if Ttl_PT_VAR_TM_PNN_BI < 90 OR Ttl_PT_VAR_TM_PNN_BI> 110 then
					Ttl_PT_ACH_TM_PNN_BI = 85
				elseif (Ttl_PT_VAR_TM_PNN_BI > 89 and Ttl_PT_VAR_TM_PNN_BI < 95) OR  (Ttl_PT_VAR_TM_PNN_BI > 104 and Ttl_PT_VAR_TM_PNN_BI <= 110) then
					Ttl_PT_ACH_TM_PNN_BI = 95
				elseif Ttl_PT_VAR_TM_PNN_BI > 94 and Ttl_PT_VAR_TM_PNN_BI< 106 then
					Ttl_PT_ACH_TM_PNN_BI = 100
				end if
				
				if Ttl_PT_VAR_TM_PNN_SBI < 90 OR Ttl_PT_VAR_TM_PNN_SBI> 110 then
					Ttl_PT_ACH_TM_PNN_SBI = 85
				elseif (Ttl_PT_VAR_TM_PNN_SBI > 89 and Ttl_PT_VAR_TM_PNN_SBI < 95) OR  (Ttl_PT_VAR_TM_PNN_SBI > 104 and Ttl_PT_VAR_TM_PNN_SBI <= 110) then
					Ttl_PT_ACH_TM_PNN_SBI = 95
				elseif Ttl_PT_VAR_TM_PNN_SBI > 94 and Ttl_PT_VAR_TM_PNN_SBI< 106 then
					Ttl_PT_ACH_TM_PNN_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_TM_PNN_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_TM_PNN_SBI,0)%></td>
			<%
				if Ttl_PT_PROD_BI <> 0 then
					Ttl_PT_TM_PML_BI		= Ttl_All_PT_TM_PML_BI / Ttl_PT_PROD_BI
				else
					Ttl_PT_TM_PML_BI		= 0
				end if
				
				if Ttl_PT_PROD_SBI <> 0 then
					Ttl_PT_TM_PML_SBI		= Ttl_All_PT_TM_PML_SBI / Ttl_PT_PROD_SBI
				else
					Ttl_PT_TM_PML_SBI		= 0
				end if 
				
				if Ttl_PT_BUD_PROD_BI <> 0 then
					Ttl_PT_BUD_TM_PML_BI	= Ttl_All_PT_BUD_TM_PML_BI / Ttl_PT_BUD_PROD_BI
				else
					Ttl_PT_BUD_TM_PML_BI	= 0
				end if
				
				if Ttl_PT_BUD_PROD_SBI <> 0 then
					Ttl_PT_BUD_TM_PML_SBI	= Ttl_All_PT_BUD_TM_PML_SBI / Ttl_PT_BUD_PROD_SBI
				else
					Ttl_PT_BUD_TM_PML_SBI	= 0
				end if
				
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_TM_PML_SBI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_TM_PML_SBI,0)%></td>
			<%
				if Ttl_PT_TM_PML_BI = 0 OR  Ttl_PT_BUD_TM_PML_BI = 0 then 
					Ttl_PT_VAR_TM_PML_BI = 0
				else
					Ttl_PT_VAR_TM_PML_BI = Ttl_PT_TM_PML_BI / Ttl_PT_BUD_TM_PML_BI * 100
				end if
				
				if Ttl_PT_TM_PML_SBI = 0 OR  Ttl_PT_BUD_TM_PML_SBI = 0 then 
					Ttl_PT_VAR_TM_PML_SBI = 0
				else
					Ttl_PT_VAR_TM_PML_SBI = Ttl_PT_TM_PML_SBI / Ttl_PT_BUD_TM_PML_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_VAR_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_VAR_TM_PML_SBI,0)%></td>
			<%
				if Ttl_PT_VAR_TM_PML_BI < 90 OR Ttl_PT_VAR_TM_PML_BI> 110 then
					Ttl_PT_ACH_TM_PML_BI = 85
				elseif (Ttl_PT_VAR_TM_PML_BI > 89 and Ttl_PT_VAR_TM_PML_BI < 95) OR  (Ttl_PT_VAR_TM_PML_BI > 104 and Ttl_PT_VAR_TM_PML_BI <= 110) then
					Ttl_PT_ACH_TM_PML_BI = 95
				elseif Ttl_PT_VAR_TM_PML_BI > 94 and Ttl_PT_VAR_TM_PML_BI< 106 then
					Ttl_PT_ACH_TM_PML_BI = 100
				end if
				
				if Ttl_PT_VAR_TM_PML_SBI < 90 OR Ttl_PT_VAR_TM_PML_SBI> 110 then
					Ttl_PT_ACH_TM_PML_SBI = 85
				elseif (Ttl_PT_VAR_TM_PML_SBI > 89 and Ttl_PT_VAR_TM_PML_SBI < 95) OR  (Ttl_PT_VAR_TM_PML_SBI > 104 and Ttl_PT_VAR_TM_PML_SBI <= 110) then
					Ttl_PT_ACH_TM_PML_SBI = 95
				elseif Ttl_PT_VAR_TM_PML_SBI > 94 and Ttl_PT_VAR_TM_PML_SBI< 106 then
					Ttl_PT_ACH_TM_PML_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_TM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_TM_PML_SBI,0)%></td>
			<%
			if Ttl_All_PT_HA_TBM <> 0 then
					Ttl_PT_TBM_PML_BI 		= Ttl_All_PT_TBM_PML_BI / Ttl_All_PT_HA_TBM
					Ttl_PT_TBM_PML_SBI		= Ttl_All_PT_TBM_PML_SBI / Ttl_All_PT_HA_TBM
					Ttl_PT_BUD_TBM_PML_BI	= Ttl_All_PT_BUD_TBM_PML_BI / Ttl_All_PT_HA_TBM
					Ttl_PT_BUD_TBM_PML_SBI	= Ttl_All_PT_BUD_TBM_PML_SBI / Ttl_All_PT_HA_TBM
				else
					Ttl_PT_TBM_PML_BI 		= 0
					Ttl_PT_TBM_PML_SBI		= 0
					Ttl_PT_BUD_TBM_PML_BI	= 0
					Ttl_PT_BUD_TBM_PML_SBI	= 0
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_TBM_PML_SBI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_BUD_TBM_PML_SBI,0)%></td>
			<%
				if Ttl_PT_TBM_PML_BI = 0 OR  Ttl_PT_BUD_TBM_PML_BI = 0 then 
					Ttl_PT_VAR_TBM_PML_BI = 0
				else
					Ttl_PT_VAR_TBM_PML_BI = Ttl_PT_TBM_PML_BI / Ttl_PT_BUD_TBM_PML_BI * 100
				end if
				
				if Ttl_PT_TBM_PML_SBI = 0 OR  Ttl_PT_BUD_TBM_PML_SBI = 0 then 
					Ttl_PT_VAR_TBM_PML_SBI = 0
				else
					Ttl_PT_VAR_TBM_PML_SBI = Ttl_PT_TBM_PML_SBI / Ttl_PT_BUD_TBM_PML_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_VAR_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_VAR_TBM_PML_SBI,0)%></td>
			<%
				if Ttl_PT_VAR_TBM_PML_BI < 90 OR Ttl_PT_VAR_TBM_PML_BI> 110 then
					Ttl_PT_ACH_TBM_PML_BI = 85
				elseif (Ttl_PT_VAR_TBM_PML_BI > 89 and Ttl_PT_VAR_TBM_PML_BI < 95) OR  (Ttl_PT_VAR_TBM_PML_BI > 104 and Ttl_PT_VAR_TBM_PML_BI <= 110) then
					Ttl_PT_ACH_TBM_PML_BI = 95
				elseif Ttl_PT_VAR_TBM_PML_BI > 94 and Ttl_PT_VAR_TBM_PML_BI< 106 then
					Ttl_PT_ACH_TBM_PML_BI = 100
				end if
				
				if Ttl_PT_VAR_TBM_PML_SBI < 90 OR Ttl_PT_VAR_TBM_PML_SBI> 110 then
					Ttl_PT_ACH_TBM_PML_SBI = 85
				elseif (Ttl_PT_VAR_TBM_PML_SBI > 89 and Ttl_PT_VAR_TBM_PML_SBI < 95) OR  (Ttl_PT_VAR_TBM_PML_SBI > 104 and Ttl_PT_VAR_TBM_PML_SBI <= 110) then
					Ttl_PT_ACH_TBM_PML_SBI = 95
				elseif Ttl_PT_VAR_TBM_PML_SBI > 94 and Ttl_PT_VAR_TBM_PML_SBI< 106 then
					Ttl_PT_ACH_TBM_PML_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_TBM_PML_BI,0)%></td>
			<td align="right" style="background-color:#FFFF00;"><%=FormatNumber(Ttl_PT_ACH_TBM_PML_SBI,0)%></td>
		</tr>
		<%
			NOWil = 1
			Ttl_PT_PROD_BI 			= 0
			Ttl_PT_PROD_SBI 		= 0
			Ttl_PT_BUD_PROD_BI 		= 0
			Ttl_PT_BUD_PROD_SBI 	= 0
			Ttl_PT_ACH_PROD_BI 		= 0
			Ttl_PT_ACH_PROD_SBI		= 0
			
			Ttl_PT_NP_BI 			= 0
			Ttl_PT_NP_SBI 			= 0
			Ttl_PT_BUD_NP_BI 		= 0
			Ttl_PT_BUD_NP_SBI 		= 0
			Ttl_PT_ACH_NP_BI 		= 0
			Ttl_PT_ACH_NP_SBI 		= 0
			
			Ttl_PT_TM_PNN_BI 		= 0
			Ttl_PT_TM_PNN_SBI 		= 0
			Ttl_PT_BUD_TM_PNN_BI 	= 0
			Ttl_PT_BUD_TM_PNN_SBI 	= 0
			Ttl_PT_VAR_TM_PNN_BI 	= 0
			Ttl_PT_VAR_TM_PNN_SBI 	= 0
			Ttl_PT_ACH_TM_PNN_BI 	= 0
			Ttl_PT_ACH_TM_PNN_SBI	= 0
			
			Ttl_All_PT_TM_PNN_BI 		= 0 				  
			Ttl_All_PT_TM_PNN_SBI 		= 0 				
			Ttl_All_PT_BUD_TM_PNN_BI 	= 0 		  
			Ttl_All_PT_BUD_TM_PNN_SBI 	= 0 	
			
			Ttl_PT_TM_PML_BI 			= 0
			Ttl_PT_TM_PML_SBI 			= 0
			Ttl_PT_BUD_TM_PML_BI 		= 0
			Ttl_PT_BUD_TM_PML_SBI 		= 0
			Ttl_PT_VAR_TM_PML_BI 		= 0
			Ttl_PT_VAR_TM_PML_SBI 		= 0
			Ttl_PT_ACH_TM_PML_BI 		= 0
			Ttl_PT_ACH_TM_PML_SBI		= 0
			
			Ttl_All_PT_TM_PML_BI 		= 0 				  
			Ttl_All_PT_TM_PML_SBI 		= 0 				
			Ttl_All_PT_BUD_TM_PML_BI 	= 0 		  
			Ttl_All_PT_BUD_TM_PML_SBI 	= 0 	
			
			Ttl_PT_TBM_PML_BI 		= 0
			Ttl_PT_TBM_PML_SBI 		= 0
			Ttl_PT_BUD_TBM_PML_BI 	= 0
			Ttl_PT_BUD_TBM_PML_SBI 	= 0
			Ttl_PT_VAR_TBM_PML_BI 	= 0
			Ttl_PT_VAR_TBM_PML_SBI 	= 0
			Ttl_PT_ACH_TBM_PML_BI 	= 0
			Ttl_PT_ACH_TBM_PML_SBI 	= 0
			
			Ttl_All_PT_TBM_PML_BI 		= 0 				
			Ttl_All_PT_TBM_PML_SBI 		= 0 			  
			Ttl_All_PT_BUD_TBM_PML_BI 	= 0 		
			Ttl_All_PT_BUD_TBM_PML_SBI 	= 0 	  
			
			Ttl_All_PT_HA_TM			= 0
			Ttl_All_PT_HA_TBM			= 0
			Ttl_All_PT_HA_NEWTBM		= 0
			
			else
				NOWil = NOWil + 1
			end if
			Dim Grnd_All_Ttl_TM_PNN_BI, Grnd_All_Ttl_TM_PNN_SBI,Grnd_All_Ttl_BUD_TM_PNN_BI,Grnd_All_Ttl_BUD_TM_PNN_SBI
			Dim Grnd_All_Ttl_TBM_PML_BI, Grnd_All_Ttl_TBM_PML_SBI, Grnd_All_Ttl_BUD_TBM_PML_BI, Grnd_All_Ttl_BUD_TBM_PML_SBI
			Dim Grnd_All_Ttl_TM_PML_BI, Grnd_All_Ttl_TM_PML_SBI,Grnd_All_Ttl_BUD_TM_PML_BI,Grnd_All_Ttl_BUD_TM_PML_SBI
			Dim Grand_Ttl_HA_TM, Grand_Ttl_HA_TBM,Grand_Ttl_HA_NEWTBM
			Grnd_Ttl_PROD_BI		= Grnd_Ttl_PROD_BI +  Ttl_PROD_BI
			Grnd_Ttl_PROD_SBI		= Grnd_Ttl_PROD_SBI + Ttl_PROD_SBI
			Grnd_Ttl_BUD_PROD_BI	= Grnd_Ttl_BUD_PROD_BI + Ttl_BUD_PROD_BI
			Grnd_Ttl_BUD_PROD_SBI	= Grnd_Ttl_BUD_PROD_SBI + Ttl_BUD_PROD_SBI	
			Ttl_PROD_BI				= 0
			Ttl_PROD_SBI 			= 0
			Ttl_BUD_PROD_BI 		= 0
			Ttl_BUD_PROD_SBI 		= 0
			Ttl_ACH_PROD_BI 		= 0
			Ttl_ACH_PROD_SBI 		= 0
			
			Grnd_Ttl_NP_BI			= Grnd_Ttl_NP_BI + Ttl_NP_BI
			Grnd_Ttl_NP_SBI			= Grnd_Ttl_NP_SBI + Ttl_NP_SBI
			Grnd_Ttl_BUD_NP_BI		= Grnd_Ttl_BUD_NP_BI + Ttl_BUD_NP_BI
			Grnd_Ttl_BUD_NP_SBI		= Grnd_Ttl_BUD_NP_SBI + Ttl_BUD_NP_SBI
			Ttl_NP_BI				= 0
			Ttl_NP_SBI				= 0
			Ttl_BUD_NP_BI			= 0
			Ttl_BUD_NP_SBI			= 0
			Ttl_ACH_NP_BI			= 0
			Ttl_ACH_NP_SBI			= 0
			
			Grnd_All_Ttl_TM_PNN_BI		= Grnd_All_Ttl_TM_PNN_BI +  CDbl(rsEPS("TM_PNN_BI"))		
			Grnd_All_Ttl_TM_PNN_SBI		= Grnd_All_Ttl_TM_PNN_SBI + CDbl(rsEPS("TM_PNN_SBI"))
			Grnd_All_Ttl_BUD_TM_PNN_BI	= Grnd_All_Ttl_BUD_TM_PNN_BI + CDbl(rsEPS("BUD_TM_PNN_BI"))
			Grnd_All_Ttl_BUD_TM_PNN_SBI	= Grnd_All_Ttl_BUD_TM_PNN_SBI + CDbl(rsEPS("BUD_TM_PNN_SBI"))
			Ttl_TM_PNN_BI			= 0
			Ttl_TM_PNN_SBI			= 0
			Ttl_BUD_TM_PNN_BI		= 0
			Ttl_BUD_TM_PNN_SBI		= 0
			Ttl_VAR_TM_PNN_BI		= 0
			Ttl_VAR_TM_PNN_SBI		= 0
			Ttl_ACH_TM_PNN_BI		= 0
			Ttl_ACH_TM_PNN_SBI		= 0
			
			Ttl_All_TM_PNN_BI		= 0
			Ttl_All_TM_PNN_SBI		= 0
			Ttl_All_BUD_TM_PNN_BI	= 0
			Ttl_All_BUD_TM_PNN_SBI	= 0
			
			Grnd_All_Ttl_TM_PML_BI		= Grnd_All_Ttl_TM_PML_BI + CDbl(rsEPS("TM_PML_BI"))
			Grnd_All_Ttl_TM_PML_SBI		= Grnd_All_Ttl_TM_PML_SBI + CDbl(rsEPS("TM_PML_SBI"))
			Grnd_All_Ttl_BUD_TM_PML_BI	= Grnd_All_Ttl_BUD_TM_PML_BI + CDbl(rsEPS("BUD_TM_PML_BI"))			
			Grnd_All_Ttl_BUD_TM_PML_SBI	= Grnd_All_Ttl_BUD_TM_PML_SBI + CDbl(rsEPS("BUD_TM_PML_SBI"))
			Ttl_TM_PML_BI			= 0
			Ttl_TM_PML_SBI			= 0
			Ttl_BUD_TM_PML_BI		= 0
			Ttl_BUD_TM_PML_SBI		= 0
			Ttl_VAR_TM_PML_BI		= 0
			Ttl_VAR_TM_PML_SBI		= 0
			Ttl_ACH_TM_PML_BI		= 0
			Ttl_ACH_TM_PML_SBI		= 0
			
			Ttl_All_TM_PML_BI		= 0
			Ttl_All_TM_PML_SBI		= 0
			Ttl_All_BUD_TM_PML_BI	= 0
			Ttl_All_BUD_TM_PML_SBI	= 0
			
			Grnd_All_Ttl_TBM_PML_BI		= Grnd_All_Ttl_TBM_PML_BI + CDbl(rsEPS("TBM_PML_BI"))
			Grnd_All_Ttl_TBM_PML_SBI	= Grnd_All_Ttl_TBM_PML_SBI + CDbl(rsEPS("TBM_PML_SBI"))
			Grnd_All_Ttl_BUD_TBM_PML_BI	= Grnd_All_Ttl_BUD_TBM_PML_BI + CDbl(rsEPS("BUD_TBM_PML_BI"))
			Grnd_All_Ttl_BUD_TBM_PML_SBI= Grnd_All_Ttl_BUD_TBM_PML_SBI + CDbl(rsEPS("BUD_TBM_PML_SBI"))
			Ttl_TBM_PML_BI			= 0
			Ttl_TBM_PML_SBI			= 0
			Ttl_BUD_TBM_PML_BI		= 0
			Ttl_BUD_TBM_PML_SBI		= 0
			Ttl_VAR_TBM_PML_BI		= 0
			Ttl_VAR_TBM_PML_SBI		= 0
			Ttl_ACH_TBM_PML_BI		= 0
			Ttl_ACH_TBM_PML_SBI		= 0
			
			Ttl_All_TBM_PML_BI		= 0
			Ttl_All_TBM_PML_SBI		= 0
			Ttl_All_BUD_TBM_PML_BI	= 0
			Ttl_All_BUD_TBM_PML_SBI	= 0
			
			Ttl_All_HA_TM			= 0
			Ttl_All_HA_TBM			= 0
			Ttl_All_HA_NEWTBM		= 0
			
			Grand_Ttl_HA_TM			= Grand_Ttl_HA_TM + CDbl(rsEPS("HA_TM"))
			Grand_Ttl_HA_TBM		= Grand_Ttl_HA_TBM + CDbl(rsEPS("HA_TBM"))
			Grand_Ttl_HA_NEWTBM		= Grand_Ttl_HA_NEWTBM + CDbl(rsEPS("HA_NEWTBM"))
		else 
			NO = NO+ 1
		end if
			rsEPS.MoveNext	
		wend
		%>
		<tr>
			<td align="right" colspan="2" style="background-color:#FFFF00;"><b>Grand Total</b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_PROD_BI / 1000, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_PROD_SBI / 1000, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_PROD_BI / 1000, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_PROD_SBI / 1000, 0)%></b></td>
			<%
				if Grnd_Ttl_PROD_BI = 0 OR  Grnd_Ttl_BUD_PROD_BI = 0 then 
					Grnd_Ttl_ACH_PROD_BI = 0
				else
					Grnd_Ttl_ACH_PROD_BI = Grnd_Ttl_PROD_BI / Grnd_Ttl_BUD_PROD_BI * 100					
				end if
				
				if Grnd_Ttl_PROD_SBI = 0 OR  Grnd_Ttl_BUD_PROD_SBI = 0 then 
					Grnd_Ttl_ACH_PROD_SBI = 0
				else
					Grnd_Ttl_ACH_PROD_SBI = Grnd_Ttl_PROD_SBI / Grnd_Ttl_BUD_PROD_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_PROD_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_PROD_SBI, 0)%></b></td>		
			
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_NP_BI, 0)%></b></td>	
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_NP_SBI, 0)%></b></td>	
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_NP_BI, 0)%></b></td>	
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_NP_SBI, 0)%></b></td>
			<%
				if Grnd_Ttl_NP_BI = 0 OR  Grnd_Ttl_BUD_NP_BI = 0 then 
					Grnd_Ttl_ACH_NP_BI = 0
				else
					Grnd_Ttl_ACH_NP_BI = Grnd_Ttl_NP_BI / Grnd_Ttl_BUD_NP_BI * 100					
				end if
				
				if Grnd_Ttl_NP_SBI = 0 OR  Grnd_Ttl_BUD_NP_SBI = 0 then 
					Grnd_Ttl_ACH_NP_SBI = 0
				else
					Grnd_Ttl_ACH_NP_SBI = Grnd_Ttl_NP_SBI / Grnd_Ttl_BUD_NP_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_NP_BI, 0)%></b></td>	
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_NP_SBI, 0)%></b></td>	
			<%
				if Grnd_Ttl_PROD_BI <> 0 then
					Grnd_Ttl_TM_PNN_BI 			= Grnd_All_Ttl_TM_PNN_BI / Grnd_Ttl_PROD_BI
				else 
					Grnd_Ttl_TM_PNN_BI 			= 0
				end if 
				
				if Grnd_Ttl_PROD_SBI <> 0 then
					Grnd_Ttl_TM_PNN_SBI 		= Grnd_All_Ttl_TM_PNN_SBI / Grnd_Ttl_PROD_SBI
				else
					Grnd_Ttl_TM_PNN_SBI			= 0
				end if
				
				if Grnd_Ttl_BUD_PROD_BI  <> 0 then 
					Grnd_Ttl_BUD_TM_PNN_BI		= Grnd_All_Ttl_BUD_TM_PNN_BI / Grnd_Ttl_BUD_PROD_BI
				else 
					Grnd_Ttl_BUD_TM_PNN_BI 		= 0
				end if
				
				if Grnd_Ttl_BUD_PROD_SBI <> 0 then
					Grnd_Ttl_BUD_TM_PNN_SBI		= Grnd_All_Ttl_BUD_TM_PNN_SBI / Grnd_Ttl_BUD_PROD_SBI
				else
					Grnd_Ttl_BUD_TM_PNN_SBI		= 0
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_TM_PNN_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_TM_PNN_SBI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_TM_PNN_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_TM_PNN_SBI, 0)%></b></td>
			<%
				if Grnd_Ttl_TM_PNN_BI = 0 OR  Grnd_Ttl_BUD_TM_PNN_BI = 0 then 
					Grnd_Ttl_VAR_TM_PNN_BI = 0
				else
					Grnd_Ttl_VAR_TM_PNN_BI = Grnd_Ttl_TM_PNN_BI / Grnd_Ttl_BUD_TM_PNN_BI * 100
				end if
				
				if Grnd_Ttl_TM_PNN_SBI = 0 OR  Grnd_Ttl_BUD_TM_PNN_SBI = 0 then 
					Grnd_Ttl_VAR_TM_PNN_SBI = 0
				else
					Grnd_Ttl_VAR_TM_PNN_SBI = Grnd_Ttl_TM_PNN_SBI / Grnd_Ttl_BUD_TM_PNN_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_VAR_TM_PNN_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_VAR_TM_PNN_SBI, 0)%></b></td>
			<%
				if Grnd_Ttl_VAR_TM_PNN_BI < 90 OR Grnd_Ttl_VAR_TM_PNN_BI> 110 then
					Grnd_Ttl_ACH_TM_PNN_BI = 85
				elseif (Grnd_Ttl_VAR_TM_PNN_BI > 89 and Grnd_Ttl_VAR_TM_PNN_BI < 95) OR  (Grnd_Ttl_VAR_TM_PNN_BI > 104 and Grnd_Ttl_VAR_TM_PNN_BI <= 110) then
					Grnd_Ttl_ACH_TM_PNN_BI = 95
				elseif Grnd_Ttl_VAR_TM_PNN_BI > 94 and Grnd_Ttl_VAR_TM_PNN_BI< 106 then
					Grnd_Ttl_ACH_TM_PNN_BI = 100
				end if
				
				if Grnd_Ttl_VAR_TM_PNN_SBI < 90 OR Grnd_Ttl_VAR_TM_PNN_SBI> 110 then
					Grnd_Ttl_ACH_TM_PNN_SBI = 85
				elseif (Grnd_Ttl_VAR_TM_PNN_SBI > 89 and Grnd_Ttl_VAR_TM_PNN_SBI < 95) OR  (Grnd_Ttl_VAR_TM_PNN_SBI > 104 and Grnd_Ttl_VAR_TM_PNN_SBI <= 110) then
					Grnd_Ttl_ACH_TM_PNN_SBI = 95
				elseif Grnd_Ttl_VAR_TM_PNN_SBI > 94 and Grnd_Ttl_VAR_TM_PNN_SBI< 106 then
					Grnd_Ttl_ACH_TM_PNN_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TM_PNN_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TM_PNN_SBI, 0)%></b></td>			
			<%
				if Grnd_Ttl_PROD_BI <> 0 then
					Grnd_Ttl_TM_PML_BI 			= Grnd_All_Ttl_TM_PML_BI / Grnd_Ttl_PROD_BI
				else 
					Grnd_Ttl_TM_PML_BI 			= 0
				end if 
				
				if Grnd_Ttl_PROD_SBI <> 0 then
					Grnd_Ttl_TM_PML_SBI 		= Grnd_All_Ttl_TM_PML_SBI / Grnd_Ttl_PROD_SBI
				else
					Grnd_Ttl_TM_PML_SBI			= 0
				end if
				
				if Grnd_Ttl_BUD_PROD_BI  <> 0 then 
					Grnd_Ttl_BUD_TM_PML_BI		= Grnd_All_Ttl_BUD_TM_PML_BI / Grnd_Ttl_BUD_PROD_BI
				else 
					Grnd_Ttl_BUD_TM_PML_BI 		= 0
				end if
				
				if Grnd_Ttl_BUD_PROD_SBI <> 0 then
					Grnd_Ttl_BUD_TM_PML_SBI		= Grnd_All_Ttl_BUD_TM_PML_SBI / Grnd_Ttl_BUD_PROD_SBI
				else
					Grnd_Ttl_BUD_TM_PML_SBI		= 0
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_TM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_TM_PML_SBI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_TM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_TM_PML_SBI, 0)%></b></td>
			<%
				if Grnd_Ttl_TM_PML_BI = 0 OR  Grnd_Ttl_BUD_TM_PML_BI = 0 then 
					Grnd_Ttl_VAR_TM_PML_BI = 0
				else
					Grnd_Ttl_VAR_TM_PML_BI = Grnd_Ttl_TM_PML_BI / Grnd_Ttl_BUD_TM_PML_BI * 100
				end if
				
				if Grnd_Ttl_TM_PML_SBI = 0 OR  Grnd_Ttl_BUD_TM_PML_SBI = 0 then 
					Grnd_Ttl_VAR_TM_PML_SBI = 0
				else
					Grnd_Ttl_VAR_TM_PML_SBI = Grnd_Ttl_TM_PML_SBI / Grnd_Ttl_BUD_TM_PML_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_VAR_TM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_VAR_TM_PML_SBI, 0)%></b></td>
			<%
				if Grnd_Ttl_VAR_TM_PML_BI < 90 OR Grnd_Ttl_VAR_TM_PML_BI> 110 then
					Grnd_Ttl_ACH_TM_PML_BI = 85
				elseif (Grnd_Ttl_VAR_TM_PML_BI > 89 and Grnd_Ttl_VAR_TM_PML_BI < 95) OR  (Grnd_Ttl_VAR_TM_PML_BI > 104 and Grnd_Ttl_VAR_TM_PML_BI <= 110) then
					Grnd_Ttl_ACH_TM_PML_BI = 95
				elseif Grnd_Ttl_VAR_TM_PML_BI > 94 and Grnd_Ttl_VAR_TM_PML_BI< 106 then
					Grnd_Ttl_ACH_TM_PML_BI = 100
				end if
				
				if Grnd_Ttl_VAR_TM_PML_SBI < 90 OR Grnd_Ttl_VAR_TM_PML_SBI> 110 then
					Grnd_Ttl_ACH_TM_PML_SBI = 85
				elseif (Grnd_Ttl_VAR_TM_PML_SBI > 89 and Grnd_Ttl_VAR_TM_PML_SBI < 95) OR  (Grnd_Ttl_VAR_TM_PML_SBI > 104 and Grnd_Ttl_VAR_TM_PML_SBI <= 110) then
					Grnd_Ttl_ACH_TM_PML_SBI = 95
				elseif Grnd_Ttl_VAR_TM_PML_SBI > 94 and Grnd_Ttl_VAR_TM_PML_SBI< 106 then
					Grnd_Ttl_ACH_TM_PML_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TM_PML_SBI, 0)%></b></td>
			<%
			if Grand_Ttl_HA_TBM <> 0 then
					Grnd_Ttl_TBM_PML_BI 		= Grnd_All_Ttl_TBM_PML_BI / Grand_Ttl_HA_TBM
					Grnd_Ttl_TBM_PML_SBI		= Grnd_All_Ttl_TBM_PML_SBI / Grand_Ttl_HA_TBM
					Grnd_Ttl_BUD_TBM_PML_BI		= Grnd_All_Ttl_BUD_TBM_PML_BI / Grand_Ttl_HA_TBM
					Grnd_Ttl_BUD_TBM_PML_SBI	= Grnd_All_Ttl_BUD_TBM_PML_SBI / Grand_Ttl_HA_TBM
				else
					Grnd_Ttl_TBM_PML_BI 		= 0
					Grnd_Ttl_TBM_PML_SBI		= 0
					Grnd_Ttl_BUD_TBM_PML_BI		= 0
					Grnd_Ttl_BUD_TBM_PML_SBI	= 0
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_TBM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_TBM_PML_SBI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_TBM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_TBM_PML_SBI, 0)%></b></td>
			<%
				if Grnd_Ttl_TBM_PML_BI = 0 OR  Grnd_Ttl_BUD_TBM_PML_BI = 0 then 
					Grnd_Ttl_VAR_TBM_PML_BI = 0
				else
					Grnd_Ttl_VAR_TBM_PML_BI = Grnd_Ttl_TBM_PML_BI / Grnd_Ttl_BUD_TBM_PML_BI * 100
				end if
				
				if Grnd_Ttl_TBM_PML_SBI = 0 OR  Grnd_Ttl_BUD_TBM_PML_SBI = 0 then 
					Grnd_Ttl_VAR_TBM_PML_SBI = 0
				else
					Grnd_Ttl_VAR_TBM_PML_SBI = Grnd_Ttl_TBM_PML_SBI / Grnd_Ttl_BUD_TBM_PML_SBI * 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_VAR_TBM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_VAR_TBM_PML_SBI, 0)%></b></td>
			<%
				if Grnd_Ttl_VAR_TBM_PML_BI < 90 OR Grnd_Ttl_VAR_TBM_PML_BI> 110 then
					Grnd_Ttl_ACH_TBM_PML_BI = 85
				elseif (Grnd_Ttl_VAR_TBM_PML_BI > 89 and Grnd_Ttl_VAR_TBM_PML_BI < 95) OR  (Grnd_Ttl_VAR_TBM_PML_BI > 104 and Grnd_Ttl_VAR_TBM_PML_BI <= 110) then
					Grnd_Ttl_ACH_TBM_PML_BI = 95
				elseif Grnd_Ttl_VAR_TBM_PML_BI > 94 and Grnd_Ttl_VAR_TBM_PML_BI< 106 then
					Grnd_Ttl_ACH_TBM_PML_BI = 100
				end if
				
				if Grnd_Ttl_VAR_TBM_PML_SBI < 90 OR Grnd_Ttl_VAR_TBM_PML_SBI> 110 then
					Grnd_Ttl_ACH_TBM_PML_SBI = 85
				elseif (Grnd_Ttl_VAR_TBM_PML_SBI > 89 and Grnd_Ttl_VAR_TBM_PML_SBI < 95) OR  (Grnd_Ttl_VAR_TBM_PML_SBI > 104 and Grnd_Ttl_VAR_TBM_PML_SBI <= 110) then
					Grnd_Ttl_ACH_TBM_PML_SBI = 95
				elseif Grnd_Ttl_VAR_TBM_PML_SBI > 94 and Grnd_Ttl_VAR_TBM_PML_SBI< 106 then
					Grnd_Ttl_ACH_TBM_PML_SBI = 100
				end if
			%>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TBM_PML_BI, 0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACH_TBM_PML_SBI, 0)%></b></td>
		</tr>
		<%
		end if
		%>
	</tbody>
</table>
</form>
</div>
</div>
</body>
</html>