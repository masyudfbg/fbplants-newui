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
    nSpanTotal= 6
	
	 'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba 
	 'response.write  "  StockGroup " &  v_txtStockGroup	 & "  G " & v_txtStock
	 
	        Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
	        strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, BUSINESSUNIT_API.get_OrderData(fcba) norut, fcba, " 

						   
			wherecons = " where company is not null  "
			    
				
				
				
			   
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end if
			
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
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
			    selectCons = selectCons & ",fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba) "
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
						whereCons = whereCons & " And  batchcode = '" & v_txtStock  & "' "
					end if
					selectCons = selectCons & " ,batchcode "  
                    orderCons  = orderCons & ",batchcode "					
				else
					nSpanTotal = nSpanTotal - 1 
				end if

			
			
	
	' strmain  =  " select * from BIBITAN_HARIAN_TAB t " & _
			    ' " where t.fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') and user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null" 
	  
	  
	  strmain  =  " select  " & strFieldCons & "  batchcode,sum(qty_begining) qty_begining, sum(qty_receive) qty_receive,sum(qty_issue)  qty_issue,sum(qty_culling) qty_culling,(sum(qty_begining) + sum(qty_receive)-sum(qty_issue)-sum(qty_culling)) qty_closing " & _
				  " From (  select fcba,receive_note_detail_api.Get_ReceiveDate(RECEIVENOTECODE,fcba)  fddate,LOCATIONCODE  batchcode,0 qty_begining,QUANTITY qty_receive,0 qty_issue,0 qty_culling " & _
				  "			from receivenotedetail where locationtype ='MN' " & _
				  "			union all " & _
				  "			select fcba,issuedate   fddate,nurserycode batchcode,0 qty_begining,0 qty_receive,quantity qty_issue,0 qty_culling from nurseryissue " & _
				  "			union all " & _
				  "			select fcba,cullingdate fddate,nurserycode batchcode,0 qty_begining,0 qty_receive,0 qty_issue,quantity  qty_culling from nurseryculling" & _
				  "			union all " & _
				  "			select n.fcba,to_date('" & v_txttodaydate & "','MM/DD/YYYY') fddate, n.nurserycode batchcode,n.openingqty qty_begining,0 qty_receive,0 qty_issue,0 qty_culling " & _
				  "			from nurseryledger n " & _
				  "			where n.year = to_char(to_date('" & v_txttodaydate & "','MM/DD/YYYY'),'YYYY') " & _
				  "			and  n.month = to_char(to_date('" & v_txttodaydate & "','MM/DD/YYYY'),'MM') " & _
				  "		  )  " & _
				  " where fddate <= to_date('" & v_txttodaydate & "','MM/DD/YYYY') and  to_char(fddate,'MMYYYY') = to_char(to_date('" & v_txttodaydate & "','MM/DD/YYYY'),'MMYYYY') " & _
				  " group by  fcba,batchcode "

		
		strStockGudang =", SUM (qty_begining) opening_balance_day," & _
						" SUM (qty_receive) received," & _
						" SUM (qty_issue) issued," & _
						" SUM (qty_culling) adjustment," & _
						" SUM (qty_closing) closing_balance_day"
						
		strStockGudang = "select GroupPlantation " & selectCons & strStockGudang & " FROM ( " & strMain & " ) " & whereCons & "  group by  GroupPlantation" & selectCons & "  Order by GroupPlantation" & orderCons 
		
		
	
    ' response.write strStockGudang
			
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
<title>ePLAS Reporting - Laporan Harian StockLedger</title>
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
<form name="frmEntry" method="post" action="p_StockbibitDaily.asp">

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
		<th>BATCH CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
       <td align="center">
			<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
					<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
			<% else %>
					<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
			<% end if %>
				<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
				</a>-->
		</td>
		
	
		
		
		<td align="center">
		<select name="txtStock">	
		<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsStock
			Set rsStock = Server.CreateObject("ADODB.Recordset")
			rsStock.CursorLocation = adUseClient
				strsql =  " SELECT '--ALL--' as nurserycode   from DUAL UNION select distinct nurserycode from Nursery WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null"
			rsStock.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStock.eof then
				DO while not rsStock.eof
				if v_txtStock = rsStock("nurserycode") then
		%>			<option value="<%=rsStock("nurserycode")%>" Selected><%=rsStock("nurserycode")%></option>
				<%else%>
					<option value="<%=rsStock("nurserycode")%>"><%=rsStock("nurserycode")%></option>
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
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
         <th colspan="14" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
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
			<h2>LAPORAN HARIAN STOCK BIBIT</h2>
			<% if v_txttodaydate <> "" then  %>
		   <h5>TANGGAL: <%=v_txttodaydate%>, KEBUN : <%=v_txtFcba%> </h5>
			<% end if %>	
	</div>
</div>
</th>
        </tr>
    
	<tr>
	     <th  style="background-color:#00A300;color:#fff;" ></th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th  style="background-color:#00A300;color:#fff;" ><%=titlePlantation%></th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th  style="background-color:#00A300;color:#fff;" ><%=titleGroupRegion%></th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th  style="background-color:#00A300;color:#fff;" ><%=titleRegion%></th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th  style="background-color:#00A300;color:#fff;" ><%=titleCompany%></th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th  style="background-color:#00A300;color:#fff;" ><%=titleSite%></th>
		<% end if %>
		
		<% if v_txtStock <> "--SUMMARY--"  And v_txtStock <> "" then %>
			 <th  style="background-color:#00A300;color:#fff;" >BATCH CODE</th>
		<%			
	    end if %>
		
		<th  style="background-color:#00A300;color:#fff;" >STOCK AWAL</th>
		<th  style="background-color:#00A300;color:#fff;" >TERIMA</th>
		<th  style="background-color:#00A300;color:#fff;" >KELUAR</th>
		<th  style="background-color:#00A300;color:#fff;" >ADJUSTMENT</th>
		<th  style="background-color:#00A300;color:#fff;" >STOCK AKHIR</th>
	</tr>
</thead>
<tbody>
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
			
				
			<td> </td>
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     					   
						   tempPlantation = rsStockGudang("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempStock = "--SUMMARY--"
					 %>
					<% if SprevPlantation <> rsStockGudang("Plantation") then %>
						<td align="center">
							<a href="p_StockbibitDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							<%=rsStockGudang("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsStockGudang("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  
					  
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
							<a href="p_StockbibitDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							<%=rsStockGudang("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsStockGudang("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% 
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
							<a href="p_StockbibitDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
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
							<a href="p_StockbibitDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
							<%=rsStockGudang("company")%>
							</a>
						</td>
						<% SprevCompany = rsStockGudang("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% 
					   
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
						      <a href="p_StockbibitDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txttodaydate=<%=v_txttodaydate%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
						      		<%=rsStockGudang("fcba")%>
							  </a>
						</td>
						<% sprevFcba = rsStockGudang("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
					<% if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then %>
				    <% 
					   tempStock = rsStockGudang ("batchcode")
					   tempPlantation = "--ALL--" 
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--" 
					   	
					 %>  			   
						<% if sgStock  <> rsStockGudang ("batchcode") then %>
								<td align="Right">
									<%=rsStockGudang("batchcode")%>
								</td>
								<% sgStock= rsStockGudang("batchcode") %>
						<% else	%>
							   <td align="center"></td>					
						<% end if %>	
						    						
				<% end if %>
				
				
				<td align="right"><%=FormatNumber(rsStockGudang("opening_balance_day"),0)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("received"),0)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("issued"),0)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("adjustment"),0)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("closing_balance_day"),0)%></td>
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
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= 1 + nSpanTotal %>">TOTAL</td>	
				<td align="right"><%=FormatNumber(nsubOpening,0)%></td>
				<td align="right"><%=FormatNumber(nsubReceive,0)%></td>
				<td align="right"><%=FormatNumber(nsubIssue,0)%></td>
				<td align="right"><%=FormatNumber(nsubAdj,0)%></td>
				<td align="right"><%=FormatNumber(nsubTotal,0)%></td>
			</tr>
    </tbody> 
</table>
</form>

</div>
</body>
</html>
