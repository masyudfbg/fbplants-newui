<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
	Dim strMain, nSpanTotal
	Dim txtyear, txtmonth,txtday, txttodaydate, txtDivision, txtBlok, txtPlantingDate,txtComp,v_txtComp
    Dim tempDivision, tempBlok, tempPlantingDate ,tempownership
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempGroupPlantation
	Dim nSubMandays_Amount_hi, nSubMandays_Amount_shi
	Dim nOutput_Mandays_Amount_hi, nOutput_Mandays_Amount_shi
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode,SprevPlantingDate
	Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,Sprevownership,SprevBlok
	Dim v_txtACCOUNTINGYEARPERIOD


v_txtDATECREATED_From= trim(request("DATECREATED_From"))
v_txtDATECREATED_To=trim(request("DATECREATED_To"))
v_txtACCOUNTINGYEARPERIOD=trim(request("ACCOUNTINGYEARPERIOD"))
v_txtBANKCODE=trim(request("BANKCODE"))
txtFCNAME=trim(request("FCNAME"))
txtControlJob=""
nOpening=0

		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		Sprevownership = ""
		SprevPlantingDate = ""
		SprevBlok = ""
		
		
	        v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	= Request("txtPlantation")
			v_txtGroupRegion  	= Request("txtGroupRegion")
			v_txtRegion  		= Request("txtRegion")
			v_txtcompany  		= Request("txtCompany")
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtPlantingDate	= Request("txtPlantingDate")
			v_txttodaydate		= Request("txttodaydate")
			
			txtyear 		= YEAR(Request("txttodaydate"))
			txtmonth 		= MONTH(Request("txttodaydate"))
			txtday			= DAY(Request("txttodaydate")) 
			txttodaydate 	= Request("txttodaydate")
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 8
			v_txtComp           =   txtComp
			v_txtyear			=	txtyear
			v_txtmonth			=	txtmonth
			
			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
			
			Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons,orderBlok,orderPlantingDate
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
			wherecons = "  "
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
			       whereCons = whereCons & " And  ff.fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",ff.fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
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
			
			
			
			   if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & ",DIVISION "
			   else
			   nSpanTotal = nSpanTotal - 1  
			   end if   


				if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderPlantingDate  = ",PLANTINGYEAR "
				else
				   nSpanTotal = nSpanTotal - 1 
				end if
		
			
				if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND ff.FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , ff.FIELDCODE,FF.FCNAME "
				   orderBlok  = ",FIELDCODE "
				else
					nSpanTotal = nSpanTotal - 1 
				end if
				
				
   		 strMain = 	 " paymentvoucherdetail pay, job job,paymentvoucher p" & _
					 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
					 " ((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" 
				
	strEPS2  =  " p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
				" pay.locationcode, nvl(sum(pay.amount)*-1,0) as amountjurnal, nvl(sum(pay.amount),0) as amountjob, pay.REFERENCE, pay.remarks, pay.fcba"

	strEPS =  " select " & selectCons & strEPS2 &"  From " & strMain & " " & whereCons & " "& _
			  " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',p.fcba) = 'E' " & _
			  " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
	  		  " pay.locationcode,pay.REFERENCE, pay.remarks, pay.fcba order by fcba,bankcode,datecreated,vouchercode"
			  ' "group by GroupPlantation, status, hectarageplanted, totalstand" & selectCons & _
			  ' ",ownership " & _
			  ' "Order by GroupPlantation" & orderCons & ",ownership, status, hectarageplanted, totalstand" & orderPlantingDate & orderBlok 


response.write strEPS				
Set rsEPS = Server.CreateObject("ADODB.Recordset")


rsEPS.Open strEPS , connstr, adOpenStatic 


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Areal dan Pokok</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<style>
.janjang{font-size:10px;background-color:#00A300;color:#fff; }
</style>
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
<form name="frmEntry" method="post" action="p_ArusCash.asp">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	

<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>TANGGAL</th>
		<th>ACCOUNTING PERIOD</th>
		<th>&nbsp;</th>
	</tr>
	<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %-->
		
		<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
		<% else%>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>-->
		</td>
		
		<td align="center">
			<% 	Dim rsComboA,Data_ACCOUNTINGYEARPERIOD
				q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where fcBA='" &  Session("fcBA") & "' order by accountingyear||lpad(accountingperiod,2,0)"

				Set rsComboA = Server.CreateObject("ADODB.Recordset")
				'response.write q
				rsComboA.CursorType = 3
				rsComboA.CursorLocation = 3
				rsComboA.Open q, connstr
				if not rsComboA.eof then%>
				<select name="ACCOUNTINGYEARPERIOD">    
					<% Do While Not rsComboA.eof %>
					<% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then 
						Data_ACCOUNTINGYEARPERIOD = "true"
					else
						Data_ACCOUNTINGYEARPERIOD = "false"
					End if%>
					<option value="<%=rsComboA("ACCOUNTINGYEARPERIOD")%>" <% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then %> Selected <%End if%>><%=rsComboA("Description")%></option> 
					<%  rsComboA.movenext
					   loop
				if Data_ACCOUNTINGYEARPERIOD = "false" then %>
					<option value="<%=ACCOUNTINGYEAR&ACCOUNTINGPERIOD%>" Selected>Tahun <%=ACCOUNTINGYEAR%> Bulan <%=ACCOUNTINGPERIOD%></option>
				<% end if %>
				
				<%
					End if
					rsComboA.Close
					set rsComboA = Nothing
					q=""
					%>
			</select> 			
		</td>
		<td align="center"><input type="submit" value="View" /></td>

	</tr>
</table>



<table border="1" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="132" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
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
			<h2>LAPORAN AREAL POKOK</h2>
			<% if v_txttodaydate <> "" and v_txtFcba <> "" then %>
				<h5>PERIODE : <%=v_txttodaydate%>, KEBUN : <%=v_txtFcba%></h5>
			<% end if %>
			<h3></br></h3>
			<h3></br></h3>
		</div>
	</div>

	<%
		if len(trim(request("txtYear"))) = 0 then
			v_txtyear = year(Now)	 
		end if

		if len(trim(request("txtmonth"))) = 0 then
			v_txtmonth = Month(Now)	 
		end if
	%>

</th>
</tr>
	<tr>
		
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="6" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="6" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="6" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="6" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="6" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="6" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>

		<!-- <th style="background-color:#00A300;color:#fff;" rowspan="10">BLOK</th> -->
		<!-- <th style="background-color:#00A300;color:#fff;">BLOK NAME</th> -->
		
		<%if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then%>
			<th rowspan="6" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<%End if %>
		
	</tr>
<!-- Pengaturan Baris 1-->
			<tr>
				<!-- <th rowspan="4" style="background-color:#00A300;color:#fff;">TANGGAL</th> -->
				<th>Date</th>
				<th> Voucher Code </th>
				<th> Activity <br> Code</th>
				<th> Activity<br>Name</th>
				<th> Target<br>Type</th>
				<th> Target<br>Code</th>
				<th> Reference</th>
				<th> Remarks</th>
				<th> Amount</th>
			</tr>	
			
			<!-- Pengaturan Baris 2-->
	    </thead>
 <tbody>
<%
Set rsEPS = Server.CreateObject("ADODB.Recordset")
rsEPS.Open strEPS , connstr, adOpenStatic 

Dim xNum

if rsEPS.bof and rsEPS.eof then
%>
	<tr>
		<td colspan="<%=132+GetLastDay(v_txtMonth, v_txtYear)%>">NO DATA AVAILABLE</td>
	</tr>
<%end if
	 dim newcolum,tempownershipnewcolum,tempDivisionnewcolum, Total
	
     newcolum = 0
	 Total = 0
	 
		while not rsEPS.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if

%>
	<tr class="<%=cClass%>">
		

		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
		
			<%if tempGroupPlantation <> rsEPS("GroupPlantation") then %>	
				<td align="center" style=" border-bottom:none;"><%=rsEPS("GroupPlantation")%></td>
				<% tempGroupPlantation = rsEPS("GroupPlantation")%>
				
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
			
			<%end if %>
			
		<%End If%>

		
		<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
			<% tempPlantation = rsEPS("Plantation")
			   tempGroupRegion = "--ALL--"
			   tempRegion = "--SUMMARY--"
			   tempCompany ="--SUMMARY--"
			   tempfcba = "--SUMMARY--"
			   tempDivision = "--SUMMARY--"
			   tempPlantingDate = "--SUMMARY--"
			   tempBlok = "--SUMMARY--"
			%>
		
			<%if SprevPlantation <> rsEPS("Plantation") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_ArusCash.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Plantation")%>
							</a>
				</td>
				<% SprevPlantation = rsEPS("Plantation")%>
	
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
		
			<%end if %>
		
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			 <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsEPS("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsEPS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
			<%if SprevGroupRegion <> rsEPS("groupregion") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_ArusCash.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("groupregion")%>
							</a></td>
				<% SprevGroupRegion = rsEPS("groupregion")%>
			
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
		
			<%end if %>
			
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%> 
		<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsEPS("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
		
			<%if SprevRegion <> rsEPS("Region") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_ArusCash.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Region")%>
							</a>
				</td>
				<% SprevRegion = rsEPS("Region")%>
			
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
		
			<%end if %>
				
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then

					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   tempCompany = rsEPS("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   

		%>
			<%if SprevCompany <> rsEPS("company") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_ArusCash.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("company")%>
							</a>
				</td>
				<% SprevCompany = rsEPS("company")%>

			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
	
			<%end if %>
		
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
					<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsEPS("fcba")
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
			<%if Sprevfcba <> rsEPS("fcba") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_ArusCash.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("fcba")%>
							</a>
				</td>
				<% Sprevfcba = rsEPS("fcba")%>
				<% newcolum = 1%>
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
				<% newcolum = 0%>
			<%end if %>
			
		<%End If%>
		
		
		<% if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then%>
			 <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsEPS("fcba")
					   else
					       tempfcba = ""
					   end if
					   tempDivision = rsEPS("division")
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--ALL--"					   
					 %>
			<%if SprevDivision <> rsEPS("DIVISION") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_ArusCash.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("DIVISION")%>
							</a></td>
				<% SprevDivision = rsEPS("DIVISION")%>
				<% tempDivisionnewcolum = 1%>
			<% elseif tempDivision = rsEPS("DIVISION") and newcolum = 1 then%>
				<td align="center" style=" border-bottom:none;"><%=rsEPS("DIVISION")%></td>
					<% SprevDivision = rsEPS("DIVISION")%>
					<% tempDivisionnewcolum = 1%>
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
				<% tempDivisionnewcolum = 0%>
			<%end if %>
			
		<%End if %>
		<%

		Set rsEPS = Server.CreateObject("ADODB.Recordset")
		rsEPS.Open strEPS , connstr, adOpenStatic 

		' Dim xNum
		Dim v_Opening,v_Debit,v_Credit,v_Saldo	

		if rsEPS.bof and rsEPS.eof then
			 nCountofField=rsEPS.fields.count-1
			 end if
	NO=1
    v_Amount = 0%>
		
	<tr>
	<td align="left" style="<%=cmdstyleLeft%>" colspan="8"><font face="Verdana" size="1">Opening Balance.........</td>
	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nOpening,2)%></td>
	
	<%
	v_Amount=cdbl(nOpening)
	While Not rsEPS.EOF%>
		<tr>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Datecreated")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Vouchercode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Jobcode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Fcname")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Locationtype")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Locationcode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Reference")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEPS("Remarks")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsEPS("amountjurnal"),2)%></td>
		</tr>
	  <%
	  v_Amount=v_Amount+trim(cdbl(rsEPS("Amountjurnal")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsEPS.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsEPS.close()
	rsEPS=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="left" style="<%=cmdstyleLeft%>" colspan="8"> <font face="Verdana" size="1">Ending Balance.........</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Amount,2)%></td>
	</tr>

		    
	</tr>
<%
			rsEPS.MoveNext
		Wend
		
rsEPS.Close
set rsEPS = Nothing

%>
</tbody>
</table>
</form>
</div>
</body>
</html>

