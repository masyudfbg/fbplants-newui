<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockGudang, strStockGudang, strGroup, strOrder, strQuery, strmain
	Dim v_txttodaydate, v_txtStock, v_txtStockGroup, NO, cClass, txtStock
	Dim rsAccountingPeriod, v_AccountingYear, v_AccountingPeriod
	Dim sprevFcba,SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	Dim TempStockGroup, TempStock,nSpanTotal
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	
	v_txttodaydate		   = request("txttodaydate")
	v_txtStockGroup		   = request("txtStockGroup")
	v_txtStock			   = request("txtStock")
    v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
	v_txtfcba 			   = Request("txtfcba")

	if v_txttodaydate ="" then
	   v_txttodaydate = textdate(Now())
    end if	
	if v_txtStockGroup	 = "" then
	   v_txtStockGroup = "--ALL--" 
	end if
	if v_txtStockGroup	 = "FUEL " then
	   v_txtStockGroup = "FUEL & LUBCRICANT"  
	   v_txtStock	   = request("txtStock")
	end if
    nSpanTotal= 7
	
	 'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba 
	 'response.write  "  StockGroup " &  v_txtStockGroup	 & "  Stockcode " & v_txtStock
	 
	        Dim selectFields,strFieldCons,selectCons,whereCons
	        strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',a.fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',a.fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',a.fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',a.fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',a.fcba) company, BUSINESSUNIT_API.get_OrderData(a.fcba) norut, a.fcba, " 

						   
			wherecons = " where company is not null AND STOCKGROUPNAME IN ('FUEL & LUBCRICANT','CHEMICAL','FERTILIZER') AND nominal > 0 "
			
			    if v_txtStockGroup <> "--ALL--" then
				   whereCons = whereCons & " And  StockGroupname = '" & v_txtStockGroup & "' "
				end if
				selectCons = selectCons & " ,StockGroupcode, StockGroupname "
		

			
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
	        if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",norut,fcba"
			    
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   fcbaname = rsCat("fcName")
				   CompanyName=rsCat("FCCOMPANYNAME")
				   MgrName = rsCat("FCMANAGER")
				   KTUName = rsCat("FCKTU")
				   CropName=rsCat("FCCROP")
				   FcbaType=rsCat("fcType")
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			
			if v_txtStock <> "--SUMMARY--" and v_txtStock <> "" then
			    if v_txtStock <> "--ALL--" then
					whereCons = whereCons & " And  stockcode = '" & v_txtStock  & "' "
				end if
				selectCons = selectCons & " , stockcode, fcname, unitofmeasurement "    
			else
			    nSpanTotal = nSpanTotal - 3 
			end if

			
			
	
	strmain  =  " SELECT " & strFieldCons & " c.fccode  StockGroupcode,c.fcname stockgroupname, a.stockcode, b.fcname, " & _
					   " b.unitofmeasurement, SUM (a.opening_balance_day) opening_balance_day, " & _
					   " SUM (a.received) received, SUM (issued) issued, " & _
					   " SUM (a.adjustment) adjustment, " & _
					   " SUM (a.opening_balance_day) " & _
					   " + SUM (a.received) " & _
					   " + SUM (a.issued) " & _
					   " + SUM (a.adjustment) closing_balance_day, abs(SUM(a.opening_balance_day)) + abs(SUM(a.received)) + abs(SUM(a.issued)) nominal " & _
				  " FROM (SELECT   fcba, stockcode, SUM (quantity) opening_balance_day, " & _
								 " 0 received, 0 issued, 0 adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate < to_date(to_char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'),'MM') ||'/01/'|| to_char( TO_DATE ('" & _ 
						   v_txttodaydate & "',  " & _
						   " 'MM/DD/YYYY'),'YYYY'), 'MM/DD/YYYY') " & _
						" GROUP BY fcba, stockcode " & _
						" UNION " & _
						" SELECT   fcba, stockcode, 0 opening_balance_day, " & _
								 " SUM (quantity) received, 0 issued, 0 adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
							 " AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
							 " AND transtype IN ('RCVDIRECT', 'RECEIVED') " & _
						" GROUP BY fcba, stockcode " & _
						" UNION " & _
						" SELECT   fcba, stockcode, 0 opening_balance_day, 0 received, " & _
								 " SUM (quantity) issued, 0 adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
							 " AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
							 " AND transtype IN ('ST_ISSUED', 'MN_ISSUE') " & _
						" GROUP BY fcba, stockcode " & _
						" UNION " & _
						" SELECT   fcba, stockcode, 0 opening_balance_day, 0 received, 0 issued, " & _
								 " SUM (quantity) adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
							 " AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
							 " AND transtype IN ('ST_WRITEOFF') " & _
						" GROUP BY fcba, stockcode) a, " & _
					   " stock b, " & _
					   " stockgroup c " & _
				 " WHERE a.stockcode = b.fccode " & _
				   " AND b.stockgroupcode = c.fccode " & _
				   " AND a.fcba = b.fcba " & _
				   " AND a.fcba = c.fcba " & _
				   " GROUP BY a.fcba, a.stockcode,c.fccode, b.fcname, c.fcname, b.unitofmeasurement"
	
		
	' strStockGudang = " select fcba"
	
		' if len(trim(v_txtStockGroup)) > 0 then 
			' if v_txtStockGroup <> "--SUMMARY--" then
			' strStockGudang = strStockGudang & " , StockGroupname "
			' end if
		' else
			' strStockGudang = strStockGudang & " , StockGroupname "
		' end if

		' if len(trim(v_txtStock)) > 0 then 
			' if v_txtStock <> "--SUMMARY--" then
				' strStockGudang = strStockGudang & " , stockcode, fcname, unitofmeasurement "
			' end if
		' else
			' strStockGudang = strStockGudang & " , stockcode , fcname, unitofmeasurement"	
		' end if
		
		strStockGudang =", SUM (opening_balance_day) opening_balance_day," & _
						" SUM (received) received," & _
						" SUM (issued) issued," & _
						" SUM (adjustment) adjustment," & _
						" SUM (closing_balance_day) closing_balance_day"
						
		strStockGudang = "select GroupPlantation " & selectCons & strStockGudang & " FROM ( " & strMain & " ) " & whereCons & "  group by  GroupPlantation" & selectCons & "  Order by GroupPlantation" & selectCons 
		
		
		' if len(trim(v_txtFcba)) > 0 then 
			' strStockGudang = strStockGudang & " WHERE fcba = '" & v_txtFcba & "'" 
		' else
			' strStockGudang = strStockGudang & " WHERE fcba like '%' " 
		' end if
		
		' ' where condition
		
		' if len(trim(v_txtStockGroup)) > 0 then 
			' if v_txtStockGroup <> "--SUMMARY--" then
				' strStockGudang = strStockGudang & " AND stockgroupname = '" & v_txtStockGroup & "'"
			' end if
		' end if
		
		' if len(trim(v_txtStock)) > 0 then 
			' if  v_txtStock <> "--SUMMARY--" then
				' strStockGudang = strStockGudang & " AND stockcode = '" & v_txtStock & "'"
			' end if
		' end if
			
		' ' Group By
		
		' strStockGudang = strStockGudang & " GROUP BY FCBA "
		' if len(trim(v_txtStockGroup)) > 0 then 
			' if v_txtStockGroup <> "--SUMMARY--" then
				' strStockGudang = strStockGudang & " , StockGroupname "
			' end if
		' else
				' strStockGudang = strStockGudang & " , StockGroupname "
		' end if

		' if len(trim(v_txtStock)) > 0 then 
			' if v_txtStock <> "--SUMMARY--" then
				' strStockGudang = strStockGudang & " , stockcode, fcname, unitofmeasurement "
			' end if
		' else
				' strStockGudang = strStockGudang & " , stockcode, fcname, unitofmeasurement "	
		' end if
		
		' ' Order by
		
		' strStockGudang = strStockGudang & " order BY FCBA "
		' if len(trim(v_txtStockGroup)) > 0 then 
			' if v_txtStockGroup <> "--SUMMARY--" then
				' strStockGudang = strStockGudang & " , StockGroupname "
			' end if
		' else
			' strStockGudang = strStockGudang & " , StockGroupname "
		' end if

		' if len(trim(v_txtStock)) > 0 then 
			' if v_txtStock <> "--SUMMARY--" then
				' strStockGudang = strStockGudang & " , stockcode , fcname, unitofmeasurement"
			' end if
		' else
			' strStockGudang = strStockGudang & " , stockcode , fcname, unitofmeasurement"	
		' end if
		
    'response.write strStockGudang
			
	Set rsStockGudang = Server.CreateObject("ADODB.Recordset")
	rsStockGudang.Open strStockGudang, connstr, adOpenStatic 
	
	
	Set rsAccountingPeriod = Server.CreateObject("ADODB.Recordset")

	strsql =  " SELECT to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYY') as ACCOUNTINGYEAR, " & _
                      "to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'MM') as ACCOUNTINGPERIOD	FROM DUAL"

	rsAccountingPeriod.Open strsql, connstr, adOpenStatic 

	if not rsAccountingPeriod.eof then
	    v_AccountingYear = rsAccountingPeriod("ACCOUNTINGYEAR")
		v_AccountingPeriod = rsAccountingPeriod("ACCOUNTINGPERIOD")
	End if
	rsAccountingPeriod.Close
	set rsAccountingPeriod = nothing
	
	
	
	
	
	
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Laporan Harian StockLedger</title>
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
<form name="frmEntry" method="post" action="p_StockGudangMonthly_BA_NAF.asp">

	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	
	</div>
<!--#include file="FormSelectNewSum.html"-->		
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>
		<th>STOCK GROUP</th>
		<th>STOCK CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
       <td align="center">
			<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
					<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
			<% else %>
					<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
			<% end if %>
				<a href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
				</a>
		</td>
		
	
		<td align="center">
		<select name="txtStockGroup">	
		<option value="--ALL--">--ALL--</option>
		<%
			Dim rsStockGroup
			Set rsStockGroup = Server.CreateObject("ADODB.Recordset")
			rsStockGroup.CursorLocation = adUseClient

				strsql =  " SELECT DISTINCT Fcname FROM StockGroup WHERE Fcname  IN ('FUEL & LUBCRICANT','CHEMICAL','FERTILIZER')"
			
			rsStockGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly
			if not rsStockGroup.eof then
				DO while not rsStockGroup.eof
				if v_txtStockGroup = rsStockGroup("Fcname") then
		%>			<option value="<%=rsStockGroup("Fcname")%>" Selected><%=rsStockGroup("Fcname")%></option>
				<%else%>
					<option value="<%=rsStockGroup("Fcname")%>"><%=rsStockGroup("Fcname")%></option>
				<%end if%>
		<%			rsStockGroup.movenext
				loop
			End if
			rsStockGroup.Close
			set rsStockGroup = nothing
		%>
		</select>		
		</td>
		
		<td align="center">
		<select name="txtStock">	
		<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsStock
			Set rsStock = Server.CreateObject("ADODB.Recordset")
			rsStock.CursorLocation = adUseClient
				strsql =  " SELECT '--ALL--' as STOCKCODE,'' FCNAME  from DUAL UNION SELECT DISTINCT FCcode as StockCODE ,fcname FROM Stock WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null and STOCKGROUPCODE in ('300','301','302')"
			rsStock.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStock.eof then
				DO while not rsStock.eof
				if v_txtStock = rsStock("STOCKCODE") then
		%>			<option value="<%=rsStock("STOCKCODE")%>" Selected><%=rsStock("STOCKCODE")%> - <%=rsStock("fcname")%></option>
				<%else%>
					<option value="<%=rsStock("STOCKCODE")%>"><%=rsStock("STOCKCODE")%> - <%=rsStock("fcname")%></option>
				<%end if%>
		<%			rsStock.movenext
				loop
			End if
			rsStock.Close
			set rsStock = nothing
		%>
		</select>		
		</td>
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>
<div id="header-wrapper">
	<div id="header">
			 <% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
			 <% End if%>
			 <% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
			 <% End if%>
	</div>
	<div id="subheader" class="centered">
			<h2>LAPORAN HARIAN STOCKLEDGER</h2>
			<% if v_txttodaydate <> "" then  %>
		   <h5> BULAN - TAHUN : <%= v_AccountingPeriod  %> - <%= v_AccountingYear %>, KEBUN : <%=v_txtFcba%> </h5>
			<% end if %>	
	</div>
</div>
<table border="0" width="100%" id="product-table">
	<tr>
	
		
		<% if v_txtStockGroup <> "--SUMMARY--"  And v_txtStockGroup <> "" then %>
			<th  style="background-color:#00A300;color:#fff;" >STOCK GROUP</th>
		<%			
	    end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th  style="background-color:#00A300;color:#fff;" >PLANTATION</th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th  style="background-color:#00A300;color:#fff;" >GROUP REGION</th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th  style="background-color:#00A300;color:#fff;" >REGION</th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th  style="background-color:#00A300;color:#fff;" >COMPANY</th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th  style="background-color:#00A300;color:#fff;" >SITE</th>
		<% end if %>
		
		<% if v_txtStock <> "--SUMMARY--"  And v_txtStock <> "" then %>
			 <th  style="background-color:#00A300;color:#fff;" >STOCK CODE</th>
			 <th  style="background-color:#00A300;color:#fff;" >STOCK DESCRIPTION</th>
			 <th  style="background-color:#00A300;color:#fff;" >SATUAN</th>
		<%			
	    end if %>
		<th  style="background-color:#00A300;color:#fff;" >STOCK AWAL</th>
		<th  style="background-color:#00A300;color:#fff;" >TERIMA</th>
		<th  style="background-color:#00A300;color:#fff;" >KELUAR</th>
		<th  style="background-color:#00A300;color:#fff;" >ADJUSTMENT</th>
		<th  style="background-color:#00A300;color:#fff;" >STOCK AKHIR</th>
	</tr>
<%
Dim nsubOpening, nsubReceive, nsubIssue, nsubAdj, nsubTotal, ngrandTotal
Dim sgfcba,sgStockGroup,sgStock,sgstockname,sgUOM

	if not rsStockGudang.EOF then
		NO = 1
		nsubOpening	= 0
		nsubReceive = 0
		nsubIssue 	= 0
		nsubAdj 	= 0
		nsubTotal 	= 0
		ngrandTotal = 0
		
		sgfcba=""
		sgStockGroup =""
		sgStock=""
		
		
		while not rsStockGudang.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
			
				
				<% if v_txtStockGroup <> "--SUMMARY--"   AND v_txtStockGroup <> "" then %>
				    <% tempStockGroup = rsStockGudang("stockgroupname") 
					   tempPlantation = "--SUMMARY--" 
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--" 
					   tempStock = "--SUMMARY--" 
					 %> 
					
						<% if sgStockGroup <> rsStockGudang("stockgroupname") then %>
							<td align="left">
								<a href="p_StockGudangMonthly_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
								<%=rsStockGudang("stockgroupname")%>
								</a>
							</td>
							<% sgStockGroup = rsStockGudang("stockgroupname") %>
						<% else	%>
							<td align="left"></td>					
						<% end if %>							 
				<% end if %>
				
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     TempStockGroup = rsStockGudang("stockgroupname") 					   
						   tempPlantation = rsStockGudang("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempStock = "--SUMMARY--" 
					 %>
					<% if SprevPlantation <> rsStockGudang("Plantation") then %>
						<td align="center">
							<a href="p_StockGudangMonthly_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							<%=rsStockGudang("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsStockGudang("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  TempStockGroup = rsStockGudang("stockgroupname") 
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsStockGudang("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsStockGudang("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempStock = "--SUMMARY--" 
					  
					%>
					
					<% if SprevGroupRegion <> rsStockGudang("GroupRegion") then %>
						<td align="center">
							<a href="p_StockGudangMonthly_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							<%=rsStockGudang("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsStockGudang("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  TempStockGroup = rsStockGudang("stockgroupname") 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsStockGudang("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsStockGudang("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsStockGudang("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempStock = "--SUMMARY--" 
					%>
					
					<% if SprevRegion <> rsStockGudang("Region") then %>
						<td align="center">
							<a href="p_StockGudangMonthly_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							<%=rsStockGudang("region")%>
							</a>
						</td>
						<% SprevRegion = rsStockGudang("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					    TempStockGroup = rsStockGudang("stockgroupname")
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsStockGudang("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsStockGudang("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsStockGudang("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsStockGudang("Company")
					   tempfcba = "--ALL--"
					   tempStock = "--SUMMARY--" 				   
					 %>
					
					<% if SprevCompany <> rsStockGudang("Company") then %>
						<td align="center">
							<a href="p_StockGudangMonthly_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							<%=rsStockGudang("company")%>
							</a>
						</td>
						<% SprevCompany = rsStockGudang("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% TempStockGroup = rsStockGudang("stockgroupname") 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsStockGudang("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsStockGudang("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsStockGudang("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsStockGudang("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsStockGudang("fcba")
					   tempStock = "--ALL--" 		   
					 %>
					
					<% if sprevFcba <> rsStockGudang("fcba") then %>
						<td align="center">
						      <a href="p_StockGudangMonthly_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							  <%=rsStockGudang("fcba")%>
							  </a>
						</td>
						<% sprevFcba = rsStockGudang("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then %>
				    <% TempStockGroup = rsStockGudang("stockgroupname") 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsStockGudang("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsStockGudang("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsStockGudang("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsStockGudang("Company")
					   else
					       tempCompany =""
					   end if
					   if v_txtfcba <> "--SUMMARY--"   and v_txtfcba <> "" then
					      tempfcba = rsStockGudang("fcba")
					   else
					      tempfcba =""
                       end if					   
					   tempStock = rsStockGudang ("stockcode")	
					 %>  			   
						<% if sgStock  <> rsStockGudang ("stockcode") then %>
								<td align="left">
									<a href="p_StockGudangMonthly_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
									<%=rsStockGudang("stockcode")%>
									</a>
								</td>
								<% sgStock= rsStockGudang("stockcode") %>
						<% else	%>
							   <td align="center"></td>					
						<% end if %>	
						<% if sgstockname <> rsStockGudang("fcname") then %>
                              <td align="left">	
								  <%=rsStockGudang("fcname")%>	
						      </td>
							  <% sgstockname = rsStockGudang("fcname") %>
					     <% Else %>
						      <td align="center"></td>	
                         <% End if %>  
                        				 
								<td align="center">
									<%=rsStockGudang("unitofmeasurement")%>
								</td>	    						
				<% end if %>
				
				<td align="right"><%=FormatNumber(rsStockGudang("opening_balance_day"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("received"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("issued"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("adjustment"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("closing_balance_day"),2)%></td>
			</tr>
<%
		NO = NO + 1
			nsubOpening = nsubOpening + Cdbl(rsStockGudang("opening_balance_day"))
			nsubReceive = nsubReceive + Cdbl(rsStockGudang("received"))
			nsubIssue = nsubIssue + Cdbl(rsStockGudang("issued"))
			nsubAdj = nsubAdj + Cdbl(rsStockGudang("adjustment"))
			nsubTotal = nsubTotal + Cdbl(rsStockGudang("closing_balance_day"))

			rsStockGudang.MoveNext
		Wend
	end if
	rsStockGudang.Close
	set rsStockGudang = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 2 + nSpanTotal %>">TOTAL</td>	
				<td align="right"><%=FormatNumber(nsubOpening,2)%></td>
				<td align="right"><%=FormatNumber(nsubReceive,2)%></td>
				<td align="right"><%=FormatNumber(nsubIssue,2)%></td>
				<td align="right"><%=FormatNumber(nsubAdj,2)%></td>
				<td align="right"><%=FormatNumber(nsubTotal,2)%></td>
			</tr>
</table>
</form>


