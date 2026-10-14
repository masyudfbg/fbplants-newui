<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%	  
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
	v_txtYear = Request("txtYear")
	v_txtMonth = Request("txtMonth")


	txtyear 		= right(Request("txttodaydate"),4)
	txtmonth 		= left(Request("txttodaydate"),2)
	txttodaydate 	= Request("txttodaydate")


	txtComp         = Request("txtCompany")
	nSpanTotal = 8
	v_txtComp           =   txtComp
	
	if request("txtYear")="" then
	   v_txtYear	=	Year(Session("v_Costbook"))
	end if
	if request("txtMonth")="" then
	   v_txtMonth	=	Month(Session("v_Costbook"))
	end if
		 
		'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
		
		Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
		strFieldCons = 	"   USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "', fcba) GroupPlantation" & _
						" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "', fcba) Plantation" & _
						" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "', fcba) groupregion" & _
						" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "', fcba) region" & _
						" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', fcba) company" & _
						" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba,"

		GroupCons = 	"   USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "', fcba)" & _
						" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "', fcba)" & _
						" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "', fcba)" & _
						" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "', fcba)" & _
						" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', fcba)" & _
						" , BUSINESSUNIT_API.get_OrderData(fcba),fcba,"
		
		wherecons = " where company is not null "
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
		       whereCons = whereCons & " And fcba = '" & v_txtfcba & "' "
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
			   orderCons  = orderCons & ",PLANTINGYEAR "
			else
		    nSpanTotal = nSpanTotal - 1    
			end if
	
		
			if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
			   if v_txtBlok <>  "--ALL--" then
				  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
			   end if
			   selectCons = selectCons & " , FIELDCODE "
			   orderCons  = orderCons & ",FIELDCODE "
			else
		    nSpanTotal = nSpanTotal - 1    
			end if
        
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	strMain = " SELECT " & strFieldCons & _
				" 		  MAINX.assetgroupcode," & _
				"         MAINX.assetgroupname," & _
				"         MAX (MAINX.acquisitionvalue_prev) acquisitionvalue_prev," & _
				"         MAX (MAINX.adjustment_amount_prev) adjustment_amount_prev," & _
				"         MAX (MAINX.fixassetamount_prev) fixassetamount_prev," & _
				"         MAX (MAINX.accumulated_depreciation_prev) accumulated_depreciation_prev," & _
				"         MAX (MAINX.yearlydepreciation_prev) yearlydepreciation_prev," & _
				"         MAX (MAINX.monthlydepreciation_prev) monthlydepreciation_prev," & _
				"         MAX (MAINX.acquisitionvalue_curr) acquisitionvalue_curr," & _
				"         MAX (MAINX.adjustment_amount_curr) adjustment_amount_curr," & _
				"         MAX (MAINX.fixassetamount_curr) fixassetamount_curr," & _
				"         MAX (MAINX.accumulated_depreciation_curr) accumulated_depreciation_curr," & _
				"         MAX (MAINX.yearlydepreciation_curr) yearlydepreciation_curr," & _
				"         MAX (MAINX.monthlydepreciation_curr) monthlydepreciation_curr" & _
				"    FROM (SELECT fcba, assetgroupcode," & _
				"                 assetgroupname," & _
				"                 assetcode," & _
				"                 assetname," & _
				"                 fccode," & _
				"                 fcname," & _
				"                 transdate," & _
				"                 acquisitionvalue_prev," & _
				"                 adjustment_amount_prev," & _
				"                 fixassetamount_prev," & _
				"                 accumulated_depreciation_prev," & _
				"                 yearlydepreciation_prev," & _
				"                 monthlydepreciation_prev," & _
				"                 0 AS acquisitionvalue_curr," & _
				"                 0 AS adjustment_amount_curr," & _
				"                 0 AS fixassetamount_curr," & _
				"                 0 AS accumulated_depreciation_curr," & _
				"                 0 AS yearlydepreciation_curr," & _
				"                 0 AS monthlydepreciation_curr" & _
				"            FROM (SELECT result1.fcba," & _
				"                         result1.assetgroupcode," & _
				"                         result1.assetgroupname," & _
				"                         result1.assetcode," & _
				"                         result1.assetname," & _
				"                         result1.fccode," & _
				"                         result1.fcname," & _
				"                         result1.transdate AS transdate," & _
				"                         result1.acquisitionvalue AS acquisitionvalue_prev," & _
				"                         result2.adjustment_amount AS adjustment_amount_prev," & _
				"                         result3.fixassetamount AS fixassetamount_prev," & _
				"                         result3.accumulated_depreciation AS accumulated_depreciation_prev," & _
				"                         result4.yearlydepreciation AS yearlydepreciation_prev," & _
				"                         result5.monthlydepreciation monthlydepreciation_prev" & _
				"                    FROM (  SELECT fa.fcba," & _
				"                                   fa.assetgroupcode," & _
				"                                   fa.assetgroupname," & _
				"                                   fa.assetcode," & _
				"                                   fa.assetname," & _
				"                                   fa.fccode," & _
				"                                   fa.fcname," & _
				"                                   fa.transdate," & _
				"                                   SUM (fa.acquisitionvalue) acquisitionvalue" & _
				"                              FROM fixedasset fa" & _
				"                             WHERE TO_NUMBER (TO_CHAR (transdate, 'yyyy')) < "& v_txtyear &"" & _
				"                          GROUP BY fa.fcba," & _
				"                                   fa.assetgroupcode," & _
				"                                   fa.assetgroupname," & _
				"                                   fa.assetcode," & _
				"                                   fa.assetname," & _
				"                                   fa.fccode," & _
				"                                   fa.fcname," & _
				"                                   fa.transdate) result1" & _
				"                         LEFT JOIN" & _
				"                         (  SELECT fad.fcba, fad.fixedassetcode, SUM (fad.adjustment_amount) adjustment_amount" & _
				"                              FROM fixedasset_adjustment fad" & _
				"                             WHERE TO_NUMBER (TO_CHAR (transdate, 'yyyy')) = "& v_txtyear &" - 1 AND fad.fctype = 'ADDITION'" & _
				"                          GROUP BY fad.fcba, fad.fixedassetcode) result2" & _
				"                            ON (    result1.fcba = result2.fcba" & _
				"                                AND result1.fccode = result2.fixedassetcode)" & _
				"                         LEFT JOIN" & _
				"                         (  SELECT fwd.fcba," & _
				"                                   fwd.fixassetcode," & _
				"                                   SUM (fwd.fixassetamount) AS fixassetamount," & _
				"                                   SUM (fwd.accumulated_depreciation)" & _
				"                                      AS accumulated_depreciation" & _
				"                              FROM fixedasset_writeoff_det fwd" & _
				"                                   INNER JOIN fixedasset_writeoff fw" & _
				"                                      ON (    fw.fccode = fwd.fccode" & _
				"                                          AND fw.fcba = fwd.fcba)" & _
				"                             WHERE TO_NUMBER (TO_CHAR (transdate, 'yyyy')) =" & _
				"                                      "& v_txtyear &" - 1" & _
				"                          GROUP BY fwd.fcba, fwd.fixassetcode) result3" & _
				"                            ON (    result1.fcba = result3.fcba" & _
				"                                AND result1.fccode = result3.fixassetcode)" & _
				"                         LEFT JOIN" & _
				"                         (  SELECT fdl.fcba," & _
				"                                   fdl.fccode," & _
				"                                   SUM (fdl.yearlydepreciation)" & _
				"                                      AS yearlydepreciation" & _
				"                              FROM fixedasset_depreciation_ledger fdl" & _
				"                             WHERE TO_NUMBER (TO_CHAR (date_transaction, 'yyyy')) =" & _
				"                                      "& v_txtyear &" -1 " & _
				"                          GROUP BY fdl.fcba, fdl.fccode) result4" & _
				"                            ON (    result1.fcba = result4.fcba" & _
				"                                AND result1.fccode = result4.fccode)" & _
				"                         INNER JOIN" & _
				"                         (  SELECT fdl.fcba," & _
				"                                   fdl.fccode," & _
				"                                   SUM (monthlydepreciation) monthlydepreciation" & _
				"                              FROM fixedasset_depreciation_ledger fdl" & _
				"                             WHERE     TO_NUMBER (" & _
				"                                          TO_CHAR (date_transaction, 'yyyy')) =" & _
				"                                          "& v_txtyear &"" & _
				"                                   AND (TO_CHAR (date_transaction, 'mm')) = "& v_txtmonth &"" & _
				"                          GROUP BY fdl.fcba, fdl.fccode) result5" & _
				"                            ON (    result1.fcba = result5.fcba" & _
				"                                AND result1.fccode = result5.fccode))" & _
				"          UNION ALL" & _
				"  SELECT 		  fcba, assetgroupcode," & _
				"                 assetgroupname," & _
				"                 assetcode," & _
				"                 assetname," & _
				"                 fccode," & _
				"                 fcname," & _
				"                 transdate," & _
				"                 0 AS acquisitionvalue_prev," & _
				"                 0 AS adjustment_amount_prev," & _
				"                 0 AS fixassetamount_prev," & _
				"                 0 AS accumulated_depreciation_prev," & _
				"                 0 AS yearlydepreciation_prev," & _
				"                 0 AS monthlydepreciation_prev," & _
				"                 acquisitionvalue_curr," & _
				"                 adjustment_amount_curr," & _
				"                 fixassetamount_curr," & _
				"                 accumulated_depreciation_curr," & _
				"                 yearlydepreciation_curr," & _
				"                 monthlydepreciation_curr" & _
				"            FROM (SELECT result1.fcba," & _
				"                         result1.assetgroupcode," & _
				"                         result1.assetgroupname," & _
				"                         result1.assetcode," & _
				"                         result1.assetname," & _
				"                         result1.fccode," & _
				"                         result1.fcname," & _
				"                         result1.transdate," & _
				"                         result1.acquisitionvalue AS acquisitionvalue_curr," & _
				"						  case TO_CHAR(result1.transdate,'yyyy')" & _
				"                         when '"& v_txtyear &"' then" & _
				"                            NVL((SELECT " & _
				"                                SUM (fad.adjustment_amount) adjustment_amount" & _
				"                                FROM fixedasset_adjustment fad" & _
                "                WHERE " & _
				"                                TO_NUMBER (TO_CHAR (transdate, 'yyyy')) = TO_CHAR(result1.transdate,'yyyy')" & _
				"                                AND fad.fctype = 'ADDITION'" & _
				"                                AND fad.fcba = result1.fcba" & _
				"                                AND fad.fixedassetcode = result1.fccode),0) + " & _
				"                            NVL((SELECT" & _
				"                                SUM (fa.acquisitionvalue) acquisitionvalue" & _
				"                                FROM fixedasset fa" & _
				"                                WHERE TO_NUMBER (TO_CHAR (transdate, 'yyyy'))  = TO_CHAR(result1.transdate,'yyyy')" & _
				"                                AND fa.fcba = result1.fcba" & _
				"                                AND fa.fccode = result1.fccode),0)" & _
				"                         else" & _
				"                            result2.adjustment_amount" & _
				"                         end AS adjustment_amount_curr, " & _
				"                         result3.fixassetamount AS fixassetamount_curr," & _				
				"                         result3.accumulated_depreciation" & _
				"                            AS accumulated_depreciation_curr," & _
				"                         result4.yearlydepreciation AS yearlydepreciation_curr," & _
				"                         result5.monthlydepreciation monthlydepreciation_curr" & _
				"                    FROM (  SELECT fa.fcba," & _
				"                                   fa.assetgroupcode," & _
				"                                   fa.assetgroupname," & _
				"                                   fa.assetcode," & _
				"                                   fa.assetname," & _
				"                                   fa.fccode," & _
				"                                   fa.fcname," & _
				"                                   fa.transdate," & _
				"                                   SUM (fa.acquisitionvalue) acquisitionvalue" & _
				"                              FROM fixedasset fa" & _
				"                             WHERE TO_NUMBER (TO_CHAR (transdate, 'yyyy')) <=" & _
				"                                      "& v_txtyear &"" & _
				"                          GROUP BY fa.fcba," & _
				"                                   fa.assetgroupcode," & _
				"                                   fa.assetgroupname," & _
				"                                   fa.assetcode," & _
				"                                   fa.assetname," & _
				"                                   fa.fccode," & _
				"                                   fa.fcname," & _
				"                                   fa.transdate) result1" & _
				"                         LEFT JOIN" & _
				"                         (  SELECT fad.fcba," & _
				"                                   fad.fixedassetcode," & _
				"                                   SUM (fad.adjustment_amount) adjustment_amount" & _
				"                              FROM fixedasset_adjustment fad" & _
				"                             WHERE     TO_NUMBER (TO_CHAR (transdate, 'yyyy')) <=" & _
				"                                          "& v_txtyear &"" & _
				"                                   AND fad.fctype = 'ADDITION'" & _
				"                          GROUP BY fad.fcba, fad.fixedassetcode) result2" & _
				"                            ON (    result1.fcba = result2.fcba" & _
				"                                AND result1.fccode = result2.fixedassetcode)" & _
				"                         LEFT JOIN" & _
				"                         (  SELECT fwd.fcba," & _
				"                                   fwd.fixassetcode," & _
				"                                   SUM (fwd.fixassetamount) AS fixassetamount," & _
				"                                   SUM (fwd.accumulated_depreciation)" & _
				"                                      AS accumulated_depreciation" & _
				"                              FROM fixedasset_writeoff_det fwd" & _
				"                                   INNER JOIN fixedasset_writeoff fw" & _
				"                                      ON (    fw.fccode = fwd.fccode" & _
				"                                          AND fw.fcba = fwd.fcba)" & _
				"                             WHERE TO_NUMBER (TO_CHAR (transdate, 'yyyy')) <=" & _
				"                                      "& v_txtyear &"" & _
				"                          GROUP BY fwd.fcba, fwd.fixassetcode) result3" & _
				"                            ON (    result1.fcba = result3.fcba" & _
				"                                AND result1.fccode = result3.fixassetcode)" & _
				"                         INNER JOIN" & _
				"                         (  SELECT fdl.fcba," & _
				"                                   fdl.fccode," & _
				"                                   SUM (fdl.monthlydepreciation)" & _
				"                                      AS yearlydepreciation" & _
				"                              FROM fixedasset_depreciation_ledger fdl" & _
				"                             WHERE TO_NUMBER (TO_CHAR (date_transaction, 'yyyy')) =" & _
				"                                      "& v_txtyear &"" & _
				"                          GROUP BY fdl.fcba, fdl.fccode) result4" & _
				"                            ON (    result1.fcba = result4.fcba" & _
				"                                AND result1.fccode = result4.fccode)" & _
				"                         INNER JOIN" & _
				"                         (  SELECT fdl.fcba," & _
				"                                   fdl.fccode," & _
				"                                   SUM (monthlydepreciation) monthlydepreciation" & _
				"                              FROM fixedasset_depreciation_ledger fdl" & _
				"                             WHERE     TO_NUMBER (" & _
				"                                          TO_CHAR (date_transaction, 'yyyy')) =" & _
				"                                          "& v_txtyear &"" & _
				"                                   AND (TO_CHAR (date_transaction, 'mm')) = "& v_txtmonth &"" & _
				"                          GROUP BY fdl.fcba, fdl.fccode) result5" & _
				"                            ON (    result1.fcba = result5.fcba" & _
				"                                AND result1.fccode = result5.fccode)))MAINX" & _
				" INNER JOIN ASSETGROUP ON MAINX.assetgroupcode = ASSETGROUP.FCCODE AND ASSETGROUP.ISMATURITY = 'FALSE'" & _
				" GROUP BY " & GroupCons &" MAINX.assetgroupcode," & _
				" MAINX.assetgroupname"

	strEPS2 =   "  		,assetgroupcode," & _
				"        assetgroupname," & _
				"        acquisitionvalue_prev," & _
				"        adjustment_amount_prev," & _
				"        fixassetamount_prev," & _
				"        accumulated_depreciation_prev," & _
				"        yearlydepreciation_prev," & _
				"        monthlydepreciation_prev," & _
				"        acquisitionvalue_curr," & _
				"        adjustment_amount_curr," & _
				"        fixassetamount_curr," & _
				"        accumulated_depreciation_curr," & _
				"        yearlydepreciation_curr," & _
				"        monthlydepreciation_curr," & _
				" 		 ( acquisitionvalue_prev - yearlydepreciation_prev ) bookvalue," & _
         		"		 ( (acquisitionvalue_prev - yearlydepreciation_prev) - yearlydepreciation_curr) netbookvalue"

	strEPS =  	" SELECT GroupPlantation " & selectCons & strEPS2 &"  From (" & strMain & ")" & whereCons & "" & _
				" Order by GroupPlantation" & orderCons &", assetgroupcode"
 
Set rsEPS = Server.CreateObject("ADODB.Recordset")
' response.write strEPS &" - "& v_txtmonth

rsEPS.Open strEPS , connstr, adOpenStatic  

%>

<html>

<head>
<title>ePlaS Reporting - Asset</title>
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
	<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
	<form name="frmEntry" method="post" action="<%=ScriptName%>">
	<!--#include file="FormSelectNewSum.html"-->
	<table border="0" id="form-table" class="centered-margin">		
		<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
		<tr>	
			<td><font face="verdana" size="1">YEAR</font></td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, i, MM
					mon = Month(Session("v_Costbook"))
					If len(v_txtMonth)=0 then v_txtMonth=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = 2015
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
			<td width="3">
					<select name="txtMonth">
				<%
					For i = 1 to 15
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
				<%
					Next
				%>
					</select>
			</td>
			<td align="center">
				<input type="submit" value="View" />
			</td> 
		</tr>
	</table>
</form>
</div>


<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
			<th colspan="18" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
				</div>
				<div id="subheader"class="centered">
					<h2>LAPORAN SUMMARY ASSET</h2>
					<br/>
					<h4>s/d&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
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
			<th rowspan="2" style="  background-color:#00A300;color:#fff; "></th>
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

			<th rowspan="2" style="background-color:#00A300;color:#fff;">ASSET GROUP CODE</th>
			<th style="background-color:#00A300;color:#fff;">ORIGINAL COST</th>
			<th colspan="4" style="background-color:#00A300;color:#fff;">PREVIOUS FINANCIAL YEAR</th>
			<th colspan="7" style="background-color:#00A300;color:#fff;">CURRENT FINANCIAL YEAR</th>		
		</tr>  
		<tr>
			<th style="background-color:#00A300;color:#fff;">CHANGE COST</th>
			<th style="background-color:#00A300;color:#fff;">ADD IN</th>
			<th style="background-color:#00A300;color:#fff;">SOLD WRITE OFF AMOUNT</th>
			<th style="background-color:#00A300;color:#fff;">SOLD AMOUNT DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL ACC. DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">BOOK VALUE</th>
			<th style="background-color:#00A300;color:#fff;">ADD IN</th>
			<th style="background-color:#00A300;color:#fff;">SOLD WRITE OFF AMOUNT</th>
			<th style="background-color:#00A300;color:#fff;">SOLD AMOUNT DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">ACC. DEPRE.</th>
			<th style="background-color:#00A300;color:#fff;">NET BOOK VALUE</th>
			<th style="background-color:#00A300;color:#fff;">MONTHLY DEPRE.</th>
		</tr>
	</thead>
<tbody>
<%  
	if not rsEPS.EOF then
		NO = 1
		nacquisitionvalue_prev = 0
		nadjustment_amount_prev = 0
		nfixassetamount_prev = 0
		naccumulated_depreciation_prev = 0
		nyearlydepreciation_prev = 0

		nbookvalue = 0
		nadjustment_amount = 0
		nfixassetamount = 0
		naccumulated_depreciation = 0
		nyearlydepreciation = 0
		nnetbookvalue = 0
		nmonthlydepreciation = 0
		
		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		rows = 1
		while not rsEPS.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>"> 
			<td align="center" Rowspan="<%=temprows %>"></td> 
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
					<% if SprevPlantation <> rsEPS("Plantation") then %>
						<td align="center" Rowspan="<%=temprows%>">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsEPS("plantation") %>	
					<% else	%>
						<td align="center" Rowspan="<%=temprows %>"></td>					
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
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsEPS("GroupRegion") then %>
						<td align="center" Rowspan="<%=temprows %>">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsEPS("GroupRegion") %>	
					<% else	%>
						<td align="center" Rowspan="<%=temprows %>"></td>					
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
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsEPS("Region") then %>
						<td align="center" Rowspan="<%=temprows %>">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("region")%>
							</a>
						</td>
						<% SprevRegion = rsEPS("region") %>	
					<% else	%>
						<td align="center" Rowspan="<%=temprows %>"></td>					
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
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsEPS("Company") then %>
						<td align="center" Rowspan="<%=temprows %>">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("company")%>
							</a>
						</td>
						<% SprevCompany = rsEPS("Company") %>	
					<% else	%>
						<td align="center" Rowspan="<%=temprows %>"></td>					
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
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsEPS("fcba") then %>
						<td align="center" Rowspan="<%=temprows %>">
							<a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") %>	
					<% else	%>
						<td align="center" Rowspan="<%=temprows %>"></td>					
					<% end if %>	 
				<% end if %>
	 
			
				<td align="left" Rowspan="<%=temprows %>"><%=rsEPS("assetgroupname")%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("acquisitionvalue_prev"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("adjustment_amount_prev"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("fixassetamount_prev"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("accumulated_depreciation_prev"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("yearlydepreciation_prev"))),2)%></td> 
				<% 
				'rumus 
				%>
				
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("bookvalue"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("adjustment_amount_curr"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("fixassetamount_curr"))),2)%></td> 
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("accumulated_depreciation_curr"))),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("yearlydepreciation_curr"))),2)%></td> 
				<% 
				'rumus 
				%>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("netbookvalue"))),2)%></td>
				<% 
				'rumus 
				%>
				<td align="right"><%=FormatNumber(cdbl(to_nol(rsEPS("monthlydepreciation_curr"))),2)%></td> 
		</tr>
				<%
			nacquisitionvalue_prev = nacquisitionvalue_prev + cdbl(to_nol(rsEPS("acquisitionvalue_prev")))
			nadjustment_amount_prev = nadjustment_amount_prev + cdbl(to_nol(rsEPS("adjustment_amount_prev")))
			nfixassetamount_prev = nfixassetamount_prev + cdbl(to_nol(rsEPS("fixassetamount_prev")))
			naccumulated_depreciation_prev = naccumulated_depreciation_prev + cdbl(to_nol(rsEPS("accumulated_depreciation_prev")))
			nyearlydepreciation_prev = nyearlydepreciation_prev + cdbl(to_nol(rsEPS("yearlydepreciation_prev")))

			nbookvalue = nbookvalue + cdbl(to_nol(rsEPS("bookvalue")))
			nadjustment_amount = nadjustment_amount + cdbl(to_nol(rsEPS("adjustment_amount_curr")))
			nfixassetamount = nfixassetamount + cdbl(to_nol(rsEPS("fixassetamount_curr")))
			naccumulated_depreciation = naccumulated_depreciation + cdbl(to_nol(rsEPS("accumulated_depreciation_curr")))
			nyearlydepreciation = nyearlydepreciation + cdbl(to_nol(rsEPS("yearlydepreciation_curr")))
			nnetbookvalue = nnetbookvalue + cdbl(to_nol(rsEPS("netbookvalue")))
			nmonthlydepreciation = nmonthlydepreciation + cdbl(to_nol(rsEPS("monthlydepreciation_curr")))
			rsEPS.MoveNext
		Wend
	end if
	rsEPS.Close
	set rsEPS = Nothing
	%>
	</tbody>
	<tr bgcolor="#F4F4A0">
		<td colspan="<%=nSpanTotal +2 %>" align="center"><strong>TOTAL </strong></td>
		<td align="right"><strong><%=FormatNumber(nacquisitionvalue_prev, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nadjustment_amount_prev, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nfixassetamount_prev, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(naccumulated_depreciation_prev, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nyearlydepreciation_prev, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nbookvalue, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nadjustment_amount, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nfixassetamount, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(naccumulated_depreciation, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nyearlydepreciation, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nnetbookvalue, 2 )%></strong></td>
		<td align="right"><strong><%=FormatNumber(nmonthlydepreciation, 2 )%></strong></td> 
	</tr>
</table>
</form>
</div>
</body>
</html>

