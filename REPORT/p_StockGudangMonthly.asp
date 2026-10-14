<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockGudang, strStockGudang, strGroup, strOrder, strQuery, strmain
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth,v_txttodaydate, v_txtStock, v_txtStockGroup, NO, cClass, txtStock
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
	
	 v_txtyear 		= Request("txtyear")
	 v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	 
	 if request("txtYear")="" then
	   v_txtYear	=	Year(Session("v_Costbook"))
	end if
	if request("txtMonth")="" then
	   v_txtMonth	=	Month(Session("v_Costbook"))
	end if
	
	 'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba 
	 'response.write  "  StockGroup " &  v_txtStockGroup	 & "  Stockcode " & v_txtStock
	 
	        Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
	        strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',a.fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',a.fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',a.fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',a.fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',a.fcba) company, BUSINESSUNIT_API.get_OrderData(a.fcba) norut, a.fcba, " 

						   
			wherecons = " where company is not null AND stockgroupcode IN (select distinct fccode from stockgroup where grouptype= 'STOCK')  "
			    
				
				
				
			    if v_txtStockGroup <> "--ALL--" then
				   whereCons = whereCons & " And  StockGroupname = '" & v_txtStockGroup & "' "
				end if
				selectCons = selectCons & " ,StockGroupcode, StockGroupname "
		        orderCons  = orderCons & " ,StockGroupcode"
                if v_txtStock <> "--SUMMARY--" and v_txtStock <> "" then
					if v_txtStock <> "--ALL--" then
						whereCons = whereCons & " And  stockcode = '" & v_txtStock  & "' "
					end if
					selectCons = selectCons & " ,ActMat, stockcode, fcname, unitofmeasurement "  
                    orderCons  = orderCons & ",ActMat,stockcode "					
				else
					nSpanTotal = nSpanTotal - 3 
				end if
			
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
			    selectCons = selectCons & ",norut,fcba"
			    orderCons  = orderCons & ",norut "
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
			
			

			
			
	
	strmain  =  " Select " & strFieldCons & " a.year,a.itemcode STOCKCODE,stock_api.get_stockName(a.itemcode,a.fcba) FCNAME,stock_api.get_ActiveMatrial(a.itemcode,a.fcba) ActMat,stock_api.get_UOM(a.itemcode,a.fcba)  unitofmeasurement, " & _
				"		stock_api.get_StockGroup(a.itemcode,a.fcba) StockGroupcode,STOCKGROUP_API.Get_StockGrpName(substr(a.itemcode,1,3),a.fcba) stockgroupname, " & _
				"		nvl(sum(a.openingbalanceqty),0) openingbalanceqty,sum(a.qtyreceived) qtyreceived,sum(a.qtyissued) qtyissued,sum(a.qtywriteoff) adjustment, " & _
				"		( nvl(sum(a.openingbalanceqty),0) + sum(a.qtyreceived) - sum(a.qtyissued) - sum(a.qtywriteoff) ) closing_balance " & _
				"  from stockledger a " & _
				"		where a.year = '" & v_txtyear & "'  and a.month='"& v_txtmonth &"' " & _
				"		GROUP BY a.fcba,a.year,a.month, a.itemcode"
	
		
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
		
		strStockGudang =", SUM(openingbalanceqty) openingbalanceqty," & _
						" SUM (qtyreceived) qtyreceived," & _
						" SUM (qtyissued) qtyissued," & _
						" SUM (adjustment) adjustment," & _
						" SUM (closing_balance) closing_balance"
						
		strStockGudang = "select GroupPlantation " & selectCons & strStockGudang & " FROM ( " & strMain & " ) " & whereCons & "  group by  GroupPlantation" & selectCons & "  Order by GroupPlantation" & orderCons 
		
		
    ' response.write strStockGudang
			
	Set rsStockGudang = Server.CreateObject("ADODB.Recordset")
	'response.write strStockGudang
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
<title>ePlaS Reporting - Laporan Bulanan StockLedger</title>
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
	</div>
<form name="frmEntry" method="post" action="p_StockGudangMonthly.asp">

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
		<th>&nbsp;</th>
		<th>BULAN</th>
		<th>TAHUN</th>
		<th>STOCK GROUP</th>
		<th>STOCK CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
       <td align="center">
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
		
		
	
		<td align="center">
		<select name="txtStockGroup">	
		<option value="--ALL--">--ALL--</option>
		<%
			Dim rsStockGroup
			Set rsStockGroup = Server.CreateObject("ADODB.Recordset")
			rsStockGroup.CursorLocation = adUseClient

				strsql =  "SELECT DISTINCT fcname FROM stockgroup WHERE grouptype = 'STOCK'"
			
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
				strsql =  " SELECT '--ALL--' as STOCKCODE,'' FCNAME  FROM STOCK WHERE fcba = '"& session("fcba") &"' UNION SELECT DISTINCT FCcode as StockCODE ,fcname FROM Stock " & _
						  " WHERE fcba = '"& v_txtFcba &"' AND user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null and STOCKGROUPCODE in (select distinct fccode from stockgroup where grouptype= 'STOCK')"
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
			<h2>LAPORAN BULANAN STOCKLEDGER</h2>
			<% if v_txttodaydate <> "" then  %>
		   <h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> </h5>
			<% end if %>	
	</div>
</div>
</th>
        </tr>
    
	<tr>
	
		
		<% if v_txtStockGroup <> "--SUMMARY--"  And v_txtStockGroup <> "" then %>
			<th  style="background-color:#00A300;color:#fff;" >STOCK GROUP</th>
		<%			
	    end if %>
		<% if v_txtStock <> "--SUMMARY--"  And v_txtStock <> "" then %>
			 <th  style="background-color:#00A300;color:#fff;" >STOCK CODE</th>
			 <th  style="background-color:#00A300;color:#fff;" >STOCK DESCRIPTION</th>
			 <th  style="background-color:#00A300;color:#fff;" >SATUAN</th>
		<%			
	    end if %>
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
			
				
				<% if v_txtStockGroup <> "--SUMMARY--"   AND v_txtStockGroup <> "" then %>
				    <% tempStockGroup = rsStockGudang("stockgroupname") 
					   tempStock = "--ALL--" 
					   tempPlantation = "--SUMMARY--" 
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--" 
					   
					 %> 
					
						<% if sgStockGroup <> rsStockGudang("stockgroupname") then %>
							<td align="left">
								<a href="p_StockGudangMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
								<%=rsStockGudang("stockgroupname")%>
								</a>
							</td>
							<% sgStockGroup = rsStockGudang("stockgroupname") %>
						<% else	%>
							<td align="left"></td>					
						<% end if %>							 
				<% end if %>
				
				<% if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then %>
				    <% TempStockGroup = rsStockGudang("stockgroupname") 
					   tempStock = rsStockGudang ("stockcode")
					   tempPlantation = "--ALL--" 
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--" 
					   	
					 %>  			   
						<% if sgStock  <> rsStockGudang ("stockcode") then %>
								<td align="left">
									<a href="p_StockGudangMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
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
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     TempStockGroup = rsStockGudang("stockgroupname") 
					       if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then 
                              tempStock = rsStockGudang ("stockcode")	
						   else
						      tempStock = ""
                           end if						   
						   tempPlantation = rsStockGudang("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
					 %>
					<% if SprevPlantation <> rsStockGudang("Plantation") then %>
						<td align="center">
							<a href="p_StockGudangMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
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
					    if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then 
                              tempStock = rsStockGudang ("stockcode")	
						else
						      tempStock = ""
                        end if	
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsStockGudang("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsStockGudang("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   
					  
					%>
					
					<% if SprevGroupRegion <> rsStockGudang("GroupRegion") then %>
						<td align="center">
							<a href="p_StockGudangMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
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
					   if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then 
                              tempStock = rsStockGudang ("stockcode")	
						else
						      tempStock = ""
                        end if	
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
					%>
					
					<% if SprevRegion <> rsStockGudang("Region") then %>
						<td align="center">
							<a href="p_StockGudangMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
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
						if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then 
                              tempStock = rsStockGudang ("stockcode")	
						else
						      tempStock = ""
                        end if	
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
					 %>
					
					<% if SprevCompany <> rsStockGudang("Company") then %>
						<td align="center">
							<a href="p_StockGudangMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtStock=<%=tempStock%>&txtStockGroup=<%=tempStockGroup%>">
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
					   if v_txtStock <> "--SUMMARY--" AND v_txtStock <> "" then 
                              tempStock = rsStockGudang ("stockcode")	
						else
						      tempStock = ""
                        end if	 
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
						      			  <%=rsStockGudang("fcba")%>
							  </a>
						</td>
						<% sprevFcba = rsStockGudang("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				
				
				<td align="right"><%=FormatNumber(rsStockGudang("openingbalanceqty"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("qtyreceived"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("qtyissued"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("adjustment"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("closing_balance"),2)%></td>
			</tr>
<%
		NO = NO + 1
			nsubOpening = nsubOpening + Cdbl(rsStockGudang("openingbalanceqty"))
			nsubReceive = nsubReceive + Cdbl(rsStockGudang("qtyreceived"))
			nsubIssue = nsubIssue + Cdbl(rsStockGudang("qtyissued"))
			nsubAdj = nsubAdj + Cdbl(rsStockGudang("adjustment"))
			nsubTotal = nsubTotal + Cdbl(rsStockGudang("closing_balance"))

			rsStockGudang.MoveNext
		Wend
	end if
	rsStockGudang.Close
	set rsStockGudang = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= 2 + nSpanTotal %>">TOTAL</td>	
				<td align="right"><%=FormatNumber(nsubOpening,2)%></td>
				<td align="right"><%=FormatNumber(nsubReceive,2)%></td>
				<td align="right"><%=FormatNumber(nsubIssue,2)%></td>
				<td align="right"><%=FormatNumber(nsubAdj,2)%></td>
				<td align="right"><%=FormatNumber(nsubTotal,2)%></td>
			</tr>
    </tbody> 
</table>
</form>

</body>
</html>
