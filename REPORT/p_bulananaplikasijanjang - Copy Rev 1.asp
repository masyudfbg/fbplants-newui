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
			txtYear		= Request("txtYear")
			
			if "" & txtYear = "" then txtYear = Year(textDate(date))
			
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
	
	strMain = " SELECT " & strFieldCons & "" & _
	          "        division" & _
	          "        , fieldcode" & _
			  "        , SUM(ha) ha" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 1, tonase, 0)) jan" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 2, tonase, 0)) feb" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 3, tonase, 0)) mar" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 4, tonase, 0)) apr" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 5, tonase, 0)) mei" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 6, tonase, 0)) jun" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 7, tonase, 0)) jul" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 8, tonase, 0)) agu" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 9, tonase, 0)) sep" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 10, tonase, 0)) okt" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 11, tonase, 0)) nov" & _
	          "        , SUM(DECODE(TO_CHAR(fddate, 'mm'), 12, tonase, 0)) des" & _
	          "        , SUM(tonase) total" & _
	          "        , SUM(rekom) rekom" & _
	          " FROM (" & _
	          " SELECT fcba" & _
			  "        , TRUNC(fddate, 'mm') fddate" & _
			  "        , field_api.getValid_Afdeling(fieldcode, fcba, fddate) division" & _
			  "        , field_api.getValid_Ha(fieldcode, fcba, fddate) ha" & _
			  "        , fieldcode" & _
			  "        , SUM(tonase) tonase" & _
			  "        , SUM(rekom) rekom" & _
              " FROM( SELECT fcba, fddate, locationcode fieldcode, SUM(ROUND(unit,0)) tonase, 0 rekom" & _
              "       FROM attendance_gad" & _
              "       WHERE jobcode IN ("&getjobcode(arr_param, txtparam)&")" & _
              "             AND TO_CHAR(fddate, 'yyyy') = '" & txtyear & "'" & _
              "       GROUP BY fcba, fddate, locationcode" & _
			  "       UNION ALL" & _
			  "       SELECT fcba, fddate, fieldcode, 0 tonase, SUM(rekom) rekom" & _
			  "       FROM (" & _
			  "           SELECT fcba" & _
			  "                  , TO_DATE('1/1/'||'" & txtyear & "', 'mm/dd/yyyy') fddate" & _
			  "                  , fieldcode" & _
			  "                  , qty_01 + qty_02 + qty_03 + qty_04 + qty_05 + qty_06" & _
			  "                    + qty_07 + qty_08 + qty_09 + qty_10 + qty_11 + qty_12 rekom" & _
			  "           FROM budget_transaction_estate" & _
			  "           WHERE budget_year = '" & txtyear & "'" & _
			  "           AND jobcode IN ("&getjobcode(arr_param, txtparam)&") )" & _
			  "       GROUP BY fcba, fddate, fieldcode  )" & _
              " WHERE USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', fcba) = 'TRUE'" & _
              " GROUP BY fcba" & _
			  "          , TRUNC(fddate, 'mm')" & _
			  "          , field_api.getValid_Afdeling(fieldcode, fcba, fddate)" & _
			  "          , field_api.getValid_Ha(fieldcode, fcba, fddate)" & _
			  "          , fieldcode )" & _
			  " GROUP BY fcba" & _
	          "          , division" & _
	          "          , fieldcode"
		
	 strEPS2  =  " , SUM(ha) ha" & _
	             " , SUM(rekom) rekom" & _
	             " , SUM(jan) jan" & _
				 " , SUM(feb) feb" & _
				 " , SUM(mar) mar" & _
				 " , SUM(apr) apr" & _
				 " , SUM(mei) mei" & _
				 " , SUM(jun) jun" & _
				 " , SUM(jul) jul" & _
				 " , SUM(agu) agu" & _
				 " , SUM(sep) sep" & _
				 " , SUM(okt) okt" & _
				 " , SUM(nov) nov" & _
				 " , SUM(des) des" & _
				 " , SUM(total) total"
				 
	strEPS =  " SELECT groupplantation " & selectCons & strEPS2 &"" & _
	          " FROM ( " & strMain & " )" & _
			  " " & whereCons & "" & _
			  "       AND USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E'" & _
			  " GROUP BY groupplantation" & selectCons & "" & _
			  " ORDER BY groupplantation" & orderCons & ""

'response.write strEPS

Set rsEPS = Server.CreateObject("ADODB.Recordset")
rsEPS.Open strEPS , connstr, adOpenStatic


function f_persen(pembilang__, pembagi__)
	res__ = 0
	if pembagi__ = 0 then
		res__ = 0
	else
		res__ = round(pembilang__/pembagi__*100,0)
	end if
	
	f_persen = res__
end function


%>

<html>

<head>
<title>Aplikasi <%=gettitle(arr_param, txtparam)%> - iPlaS Report</title>
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
		<th>TAHUN</th>
		<th>AFDELING</th>
		<th>BLOK</th>
		<th>LAPORAN</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
				<%
					Yr = Year(Session("v_Costbook"))
					MinYr = 2015
					MaxYr = Yr  + 2
					If len(yyyy)= 0 then yyyy = Yr
				%>
		
		<td align="center">
			<select name="txtYear">
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
			</select>
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
<th colspan="<%=16 +add_colspan%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
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
			<h2>MONITORING APLIKASI <%=UCASE(gettitle(arr_param, txtparam))%></h2>
				<h5>PERIODE : <%=txtyear%></h5>
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
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% 			
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% 			
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% 		
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% 		
	    end if %>

		
		<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% 		
	    end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" And v_txtPlantingDate <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">TT</th>
		<% 		
	    end if %>
		
		<% if v_txtBlok <> "--SUMMARY--" And v_txtBlok <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<% end if %>
		
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >Luas (Ha)</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >Rekomendasi (Kg)</th>
		<th colspan="12" style="background-color:#00A300;color:#fff;" >Aplikasi <%=PCASE(gettitle(arr_param, txtparam))%></th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >TOTAL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >%</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" >JAN</th>
		<th style="background-color:#00A300;color:#fff;" >FEB</th>
		<th style="background-color:#00A300;color:#fff;" >MAR</th>
		<th style="background-color:#00A300;color:#fff;" >APR</th>
		<th style="background-color:#00A300;color:#fff;" >MEI</th>
		<th style="background-color:#00A300;color:#fff;" >JUN</th>
		<th style="background-color:#00A300;color:#fff;" >JUL</th>
		<th style="background-color:#00A300;color:#fff;" >AGU</th>
		<th style="background-color:#00A300;color:#fff;" >SEP</th>
		<th style="background-color:#00A300;color:#fff;" >OKT</th>
		<th style="background-color:#00A300;color:#fff;" >NOV</th>
		<th style="background-color:#00A300;color:#fff;" >DES</th>
	</tr>
        </thead>
 <tbody>
<%
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
							<a href="p_bulananaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
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
							<a href="p_bulananaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
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
							<a href="p_bulananaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
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
							<a href="p_bulananaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
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
							<a href="p_bulananaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
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
							<a href="p_bulananaplikasijanjang.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>">
							<%=rsEPS("division")%>
							</a>
						</td>
						<% sprevDivision = rsEPS("division") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>

				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <>  "" then %>
				    <% if sprevFieldCode <> rsEPS("fieldcode") then %>				
					    <td align="center"><%=rsEPS("fieldcode")%></td>				
						<% sprevFieldCode = rsEPS("fieldcode") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>				
				
				<td align="right"><%=FormatNumber(rsEPS("ha"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("rekom"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("jan"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("feb"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("mar"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("apr"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("mei"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("jun"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("jul"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("agu"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("sep"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("okt"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("nov"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("des"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("total"),0)%></td>
				<td align="right"><%=f_persen(cdbl(to_nol(rsEPS("total"))), 0)%></td>
			</tr>
<%
			NO = NO + 1			
			rsEPS.MoveNext			
		Wend
	end if
	rsEPS.Close
	set rsEPS = Nothing
%>
</tbody>
</table>
</form>
</div>
</body>
</html>

