<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	ScriptName = Request.ServerVariables("SCRIPT_NAME")
	add_colspan = 0
	
	txtparam = request("txtparam")
	
	q = " SELECT pr.fccode, pr.fcname, prd.fccode jobcode" & _
		" FROM parameter pr" & _
		" INNER JOIN parameterdetail prd ON (pr.fcba = prd.fcba AND pr.parcode = prd.parcode AND pr.fccode = prd.parheadcode)" & _
		" WHERE pr.fcba = 'KSR'" & _
		"       AND pr.parcode = 'RPT_APLIKASIJANJANG'"
	
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.Open q , connstr, adOpenStatic
	
	arr_param_eof = true
	if not oRS.eof then
		arr_param = oRS.getrows()
		arr_param_eof = false
	end if
	oRS.Close
	set oRS = nothing
	
	if not arr_param_eof then
		for xn=0 to ubound(arr_param,2)
			if xn=0 then
				if "" & txtparam = "" then txtparam = arr_param(0,xn)
			end if
		next
	end if
	
	function getjobcode(arr__, where__)
		result__ = ""
		arr1 = arr__
		rows = 0
		if not arr_param_eof then
			for xm=0 to ubound(arr1,2)
				if cstr(""&where__) = cstr(""&arr1(0,xm)) then
					rows = rows +1
					if rows = 1 then
						result__ = result__ & "'" & arr1(2,xm) & "'"
					else
						result__ = result__ & ",'" & arr1(2,xm) & "'"
					end if
				end if
			next
		end if
		
		getjobcode = result__
	end function
	
	function gettitle(arr__, where__)
		result__ = ""
		arr1 = arr__
		if not arr_param_eof then
			for xm=0 to ubound(arr1,2)
				currx = cstr(""&arr1(0,xm))
				if currx <> prevx then
					if cstr(""&where__) = cstr(""&arr1(0,xm)) then
						result__ = "" & arr1(1,xm)
					end if
				end if
			next
		end if
		
		gettitle = result__
	end function
	
	        v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	= Request("txtPlantation")
			v_txtGroupRegion  	= Request("txtGroupRegion")
			v_txtRegion  		= Request("txtRegion")
			v_txtcompany  		= Request("txtCompany")
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision 		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtdatefrom		= Request("txtdatefrom")
			v_txtdateto			= Request("txtdateto")
			
			if "" & v_txtdatefrom = "" then v_txtdatefrom = textDate(date)
			if "" & v_txtdateto = "" then v_txtdateto = textDate(date)
			
			nSpanTotal = 8
			
			
			Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
			               " , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
			               " , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
			               " , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
			               " , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
			               " , fcba" & _
			               " , BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
			wherecons = " WHERE company IS NOT NULL "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
				add_colspan = add_colspan +1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
				add_colspan = add_colspan +1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
				add_colspan = add_colspan +1
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
				add_colspan = add_colspan +1
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
	        if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",fcba"
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
				add_colspan = add_colspan +1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			
			   if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & ",DIVISION "
				  add_colspan = add_colspan +1
			   else
			    nSpanTotal = nSpanTotal - 1 	  
			   end if 
		
			
				if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , FIELDCODE "
				   orderCons  = orderCons & ",FIELDCODE "
				   add_colspan = add_colspan +1
				else
			    nSpanTotal = nSpanTotal - 1    
				end if
	
	strMain = " SELECT " & strFieldCons & " fddate, field_api.getValid_Afdeling(fieldcode, fcba, fddate) division, fieldcode, SUM(ha) ha, SUM(tonase) tonase" & _
              " FROM( SELECT fcba, fddate, locationcode fieldcode, SUM(ROUND(output, 2)) ha, SUM(ROUND(unit,0)) tonase" & _
              "       FROM attendance_gad" & _
              "       WHERE fddate BETWEEN TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy') AND TO_DATE('" & v_txtdateto & "', 'mm/dd/yyyy')" & _
              "             AND jobcode IN ("&getjobcode(arr_param, txtparam)&")" & _
              "       GROUP BY fcba, fddate, locationcode )" & _
              " WHERE USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', fcba) = 'TRUE'" & _
              " GROUP BY fcba, fddate, fieldcode"
		
	 strEPS2  =  " , SUM(ha) ha" & _
				 " , SUM(tonase) tonase"
				 
	strEPS =  " SELECT groupplantation " & selectCons & strEPS2 &", fddate" & _
	          " FROM ( " & strMain & " )" & _
			  " " & whereCons & "" & _
			  "       AND USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E'" & _
			  " GROUP BY groupplantation" & selectCons & "" & _
			  "          , fddate" & _
			  " ORDER BY groupplantation" & orderCons & "" & _
			  "          , fddate"

'response.write strEPS

Set rsEPS = Server.CreateObject("ADODB.Recordset")
rsEPS.Open strEPS , connstr, adOpenStatic 


%>

<html>

<head>
<title>Aplikasi <%=gettitle(arr_param, txtparam)%> - ePlaS Report</title>
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
<form name="frmEntry" method="post" action="<%=ScriptName%>">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	

<!--#include file="FormSelectNewSumE.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>DARI</th>
		<th>SAMPAI</th>
		<th>AFDELING</th>
		<th>BLOK</th>
		<th>LAPORAN</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<td align="center">
			<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>" title="Type Date mm/dd/yyyy" >
		</td>
		<td align="center">
			<input class="datepicker" type="text" maxlength="10" name="txtdateto" size="10" value="<%=v_txtdateto%>" title="Type Date mm/dd/yyyy" >
		</td>
		
		<td align="center">
			<select name="txtDivision">	
				<option value="--SUMMARY--">--SUMMARY--</option>
				<%
					Dim rsDivision
					Set rsDivision = Server.CreateObject("ADODB.Recordset")
					rsDivision.CursorLocation = adUseClient
					

			
					if len(trim(v_txtDivision)) > 0 then 
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELDDETAIL WHERE FCBA = '" & Session("fcBA") & "'"
					end if
					
					
					rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsDivision.eof then
						DO while not rsDivision.eof
						if v_txtDivision = rsDivision("Division") then
				%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
						<%else%>
							<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
						<%end if%>
				<%			rsDivision.movenext
						loop
					End if
					rsDivision.Close
					set rsDivision = nothing
				%>
			</select>		
		</td>


		<td align="center">
			<select name="txtBlok">	
				<option value="--SUMMARY--">--SUMMARY--</option>
				<%
					Dim rsBlok
					Set rsBlok = Server.CreateObject("ADODB.Recordset")
					rsBlok.CursorLocation = adUseClient

			
					if len(trim(v_txtBlok)) > 0 then 
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELDDETAIL WHERE FCBA = '" & Session("fcBA") & "'"
					end if
					
					rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsBlok.eof then
						DO while not rsBlok.eof
						if v_txtBlok = rsBlok("Blok") then
				%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%></option>
						<%else%>
							<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%></option>
						<%end if%>
				<%			rsBlok.movenext
						loop
					End if
					rsBlok.Close
					set rsBlok = nothing
				%>
			</select>		
		</td>
		
		<td align="center">
			<select name="txtparam">	
				<%
						if not arr_param_eof then
							for xn=0 to ubound(arr_param,2)
								currparam = arr_param(0,xn)
								if currparam <> prevparam then
				%>
									<option value="<%=currparam%>" <%if currparam = txtparam then%> selected <%end if%>><%=arr_param(1,xn)%></option>
				<%				end if
								prevparam = currparam
							next
						end if
				%>
			</select>		
		</td>
		
		
		<td align="center"><input type="submit" value="View" /></td>

	</tr>
</table>



<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="<%=3 +add_colspan%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
		</div>
		<div id="subheader"class="centered">
			<h2>LAPORAN APLIKASI <%=UCASE(gettitle(arr_param, txtparam))%></h2>
				<h5>PERIODE : <%if v_txtdatefrom = v_txtdateto then response.write v_txtdatefrom else response.write v_txtdatefrom & " - " & v_txtdateto%></h5>
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
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% 			
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% 			
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% 		
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% 		
	    end if %>

		
		<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
		    <th style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% 		
	    end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" And v_txtPlantingDate <> "" then %>
		    <th style="background-color:#00A300;color:#fff;">TT</th>
		<% 		
	    end if %>
		
		<th style="background-color:#00A300;color:#fff;" >Tanggal</th>
		
		<% if v_txtBlok <> "--SUMMARY--" And v_txtBlok <> "" then %>
		    <th style="background-color:#00A300;color:#fff;">BLOK</th>
		<% end if %>
		
		<th style="background-color:#00A300;color:#fff;" >Luas (Ha)</th>
		<th style="background-color:#00A300;color:#fff;" >Aplikasi (Kg)</th>
	</tr>
        </thead>
 <tbody>
<%
Dim nsubhi, nsubshi, nsubbi, nsubsbi, nsubtonase_hi
Dim nsubtonase_shi, nsubtonase_bi, nsubtonase_sbi, nsubpom_hi, nsubpom_shi
Dim nsubpom_bi, nsubpom_sbi, nsubpom_tonase_hi, nsubpom_tonase_shi, nsubpom_tonase_bi, nsubpom_tonase_sbi, nsubrestan
Dim nsubbudjjg_shi,nsubbudjjg_sbi,nsubbudqty_shi,nsubbudqty_sbi,nsubbudmandays_shi,nsubbudmandays_sbi,nsubbudOutput_shi,nsubbudOutput_sbi
Dim nsubmandays_hi, nsubmandays_shi, nsubmandays_bi, nsubmandays_sbi
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim nsubOutput_hi, nsubOutput_shi, nsubOutput_bi, nsubOutput_sbi
Dim nOutput_hi, nOutput_shi, nOutput_bi, nOutput_sbi ,nbudOutput_shi,nbudOutput_sbi
Dim nsubwo_plus_hi, nsubwo_plus_shi, nsubwo_plus_sbi,nsubwo_min_hi, nsubwo_min_shi, nsubwo_min_sbi


	if not rsEPS.EOF then
		NO = 1
		gt_ha = 0
		gt_tonase = 0
		
		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		while not rsEPS.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% tempPlantation = rsEPS("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsEPS("Plantation") then %>
						<td align="center">
							<a href="p_harianaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsEPS("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsEPS("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsEPS("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsEPS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsEPS("GroupRegion") then %>
						<td align="center">
							<a href="p_harianaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsEPS("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsEPS("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
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
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsEPS("Region") then %>
						<td align="center">
							<a href="p_harianaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsEPS("region")%>
							</a>
						</td>
						<% SprevRegion = rsEPS("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
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
					   tempCompany =rsEPS("Company")
					   tempfcba = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsEPS("Company") then %>
						<td align="center">
							<a href="p_harianaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsEPS("company")%>
							</a>
						</td>
						<% SprevCompany = rsEPS("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
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
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if sprevFcba <> rsEPS("fcba") then %>
						<td align="center">
							<a href="p_harianaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> "" then %>
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
					   if txtownership <> "--SUMMARY--"  and txtownership <> "" then
					      tempownership = rsEPS("ownership")
					   else
					      tempownership = ""
                       end if
					   if v_txtDivision <> "--SUMMARY--"  and v_txtDivision <> "" then
					      tempDivision = rsEPS("division")
					   else
					      tempDivision = ""
                       end if
					   tempBlok = "--ALL--"
					 %>
				
				
				    <% if sprevDivision <> rsEPS("division") then %>
					    <td align="center">
							<a href="p_harianaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsEPS("division")%>
							</a>
						</td>
						<% sprevDivision = rsEPS("division") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
				<td align="center"><%=rsEPS("fddate")%></td>

				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <>  "" then %>
				    <% if sprevFieldCode <> rsEPS("fieldcode") then %>				
					    <td align="center"><%=rsEPS("fieldcode")%></td>				
						<% sprevFieldCode = rsEPS("fieldcode") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>				
				
				<td align="right"><%=FormatNumber(rsEPS("ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("tonase"),0)%></td>				
			</tr>
<%
			NO = NO + 1
			gt_ha = gt_ha + Cdbl(to_nol(rsEPS("ha")))
			gt_tonase = gt_tonase + Cdbl(to_nol(rsEPS("tonase")))
			
			rsEPS.MoveNext			
		Wend
	end if
	rsEPS.Close
	set rsEPS = Nothing
	
	
%>
			<tr bgcolor="#F4F4A0">			   
				<td align="center" colspan="<%=1 +add_colspan%>"><b>Total</b></td>				
				<td align="right"><b><%=FormatNumber(gt_ha,2)%></b></td>
				<td align="right"><b><%=FormatNumber(gt_tonase,0)%></b></td>				
			</tr>
</tbody>
</table>
</form>
</div>
</body>
</html>

