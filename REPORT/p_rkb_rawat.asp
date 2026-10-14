<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtMonth, v_txtYear
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName, v_txtsectionname, v_txtGangCode, v_txtStatus
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no

v_txtYear		= Request("txtYear")
v_txtMonth		= Request("txtMonth")
v_txtGroupPlantation  = Request("txtGroupPlantation")
v_txtPlantation  	= Request("txtPlantation")
v_txtGroupRegion  	= Request("txtGroupRegion")
v_txtRegion  		= Request("txtRegion")
v_txtcompany  		= Request("txtCompany")
v_txtfcba 			= Request("txtfcba")
v_txtDivision		= Request("txtDivision")
v_txtBlok			= Request("txtBlok")
v_txtPlantingDate	= Request("txtPlantingDate")
v_txtJobcode		= Request("txtJobcode")
txtownership 		= Request("txtownership")

If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
If "" & txttdate = "" Then txttdate = textDate(date)

txttdate = txtMonth & "" & txtYear

ScriptName		= Request.ServerVariables("SCRIPT_NAME")
nSpanTotal = 0
Dim v_paymenttype : v_paymenttype = Request("paymenttype")

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManagerName = rsBaInfo("FCMANAGER")
	varKTUName = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
	varManagerName = ""
	varKTUName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

txtComp         = Request("txtCompany")
nSpanTotal = 8

'response.write "Year = " & v_txtyear & " Month = " & v_txtmonth 


'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate

Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
strFieldCons = 	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',mwp.fcba) GroupPlantation" & _
				" ,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',mwp.fcba) Plantation " & _
			   	" ,USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',mwp.fcba) groupregion " & _
			   	" ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',mwp.fcba) region" & _
			   	" ,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',mwp.fcba) company, mwp.fcba, BUSINESSUNIT_API.get_OrderData(mwp.fcba) norut" 

			   
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
 

if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
	if v_txtDivision <>  "--ALL--"  then
		whereCons = whereCons & " AND unitkerja = '" & v_txtDivision & "'"
	end if
	selectCons = selectCons &  " , unitkerja "
	orderCons  = orderCons & ",unitkerja "
else
	nSpanTotal = nSpanTotal - 1 	  
end if   



if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
	if v_txtPlantingDate <>  "--ALL--" then
		whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
	end if
	selectCons = selectCons &  " "
	orderCons  = orderCons & " "
	nSpanTotal = nSpanTotal - 1    
else
	nSpanTotal = nSpanTotal - 1    
end if


if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
   if v_txtBlok <>  "--ALL--" then
	  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
   end if
   selectCons = selectCons & " , FIELDCODE "
   orderCons  = orderCons & ",FIELDCODE "
   nSpanTotal = nSpanTotal - 1    
else
nSpanTotal = nSpanTotal - 1    
end if

if v_txtJobcode <> "--SUMMARY--" and v_txtJobcode <> "" then
   if v_txtJobcode <>  "--ALL--" then
	  whereCons = whereCons & " AND activitycode = '" & v_txtJobcode & "'"
   end if
   selectCons = selectCons & " , activitycode "
   orderCons  = orderCons & ",activitycode "
   nSpanTotal = nSpanTotal - 1    
else
nSpanTotal = nSpanTotal - 1    
end if

		strMain	=	" SELECT " & strFieldCons & ", mwp.rkbno, mwp.activitycode," & _
					"        job_api.get_jobcode_name (mwp.activitycode, mwp.fcba) activityname," & _
					"        to_number(to_char(field_api.getvalid_plantingdate (mwp.fieldcode," & _
					"                                         mwp.fcba," & _
					"                                         TO_DATE ('"& v_txtMonth &"/01/"& v_txtYear &"', 'mm/dd/yyyy')),'yyyy'))" & _
					"           plantingyear," & _
					"        mwp.fieldcode," & _
					"        mwp.unitkerja," & _ 
					"        field_api.getvalid_hectarageplanted (" & _
					"           mwp.fieldcode," & _
					"           mwp.fcba," & _
					"           TO_DATE ('"& v_txtMonth &"/01/"& v_txtYear &"', 'mm/dd/yyyy'))" & _
					"           luas," & _
					"        field_api.getvalid_pokok (mwp.fieldcode," & _
					"                                  mwp.fcba," & _
					"                                  TO_DATE ('"& v_txtMonth &"/01/"& v_txtYear &"', 'mm/dd/yyyy'))" & _
					"           pokok," & _
					"        NVL (mwp.volume, 0) volume," & _
					"        NVL (0, 0) luas_dikerjakan," & _
					"        job_api.get_job_uom (mwp.activitycode, mwp.fcba) uom_job," & _					
					"        NVL(mwp.normamandays,0)   norma_hk," & _
					"        nvl(mwp.mandays,0) mandays," & _
					"         nvl(mwp.totalamount,0)totalamount," & _
					"        NVL (mwp.normapremi, 0) normapremi," & _
					"        NVL (mwp.boronganpremi, 0) boronganpremi," & _
					"        mwd.stockcode," & _
					"        /*stock_api.get_stockname (mwd.stockcode, mwd.fcba)*/ stock.fcname stockname," & _
					"        /*stock_api.get_uom (mwd.stockcode, mwd.fcba)*/ stock.uomactivematerial uom_stock," & _
					"        NVL(mwd.activematerialnorm,0) normabahan," & _
					"        NVL(mwd.qty,0) qty," & _
					"         NVL(mwd.ratestock,0) ratestock," & _
					"        NVL( mwd.totalamount,0) totalamountstock" & _
					"   FROM monthlyworkplan mwp" & _
					"        INNER JOIN monthlyworkplandet mwd" & _
					"           ON (mwp.fcba = mwd.fcba AND mwp.rkbno = mwd.rkbno)" & _
					"        LEFT JOIN (SELECT a.job,a.actmat,a.uomactivematerial,a.fcba,b.fcname from materialactivity_norm a inner join activematerial b on a.actmat = b.fccode ) stock  " & _ 
					"           ON (mwd.stockcode = stock.actmat and mwd.fcba=stock.fcba and mwp.activitycode=stock.job) " & _
					"  WHERE mwp.period_year = '"& v_txtYear &"' AND mwp.period_month = '"& v_txtMonth &"'" 

		strEPS =	"  ,rkbno,activitycode," & _
					"  activityname," & _
					"  plantingyear," & _
					"  fieldcode," & _
					"  nvl(luas,0)luas," & _
					"  pokok," & _
					"  volume," & _
					"  luas_dikerjakan," & _
					"  uom_job," & _
					"  norma_hk," & _
					"  mandays," & _
					"  totalamount," & _
					"  normapremi," & _
					"  boronganpremi," & _
					"  stockcode," & _
					"  stockname," & _
					"  uom_stock," & _
					"  normabahan," & _
					"  qty," & _
					"  ratestock," & _
					"  totalamountstock" 
		
	strPayroll = "select GroupPlantation " & selectCons & strEPS &"  From (" & strMain & " )" & whereCons & " "

	strPayroll = strPayroll & " order by plantingyear,rkbno " 
	' response.write strPayroll
	Set rsPayroll = Server.CreateObject("ADODB.Recordset")
	rsPayroll.Open strPayroll, connstr, adOpenStatic
			
%>
<html>
<head>
	<title>ePlaS Reporting - Rencana Kerja Bulanan Rawat </title>
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
	  }, false
	);
	</script>
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
	
<style type="text/css">
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
			<div class="printhide">
			<input type="button" value="&#xf02f;" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmAR" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include file="FormSelectNewsum.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<th>TAHUN</th>
						<th>BULAN</th> 
						<th>PEKERJAAN</th>
						<th>UNIT KERJA</th>
						<th>TAHUN TANAM</th>
						<th>BLOK</th>
					</tr>
					<tr>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM

							Yr = v_txtYear
							MinYr = 2015
							MaxYr = Yr  + 2
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
								<td width="2">
									<select name="txtMonth">
						<% 
							For i = 1 to 12
								
								if i <= 9 then 
									iBulan = 0&i
								else
									iBulan = i
								end if
								i	= iBulan
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
						<%
							Next
						%>
									</select>
								</td>
						<td align="center">
							<select name="txtJobcode">	
								<option value="--SUMMARY--">--SUMMARY--</option> 
								<% 
									Set rsDivision = Server.CreateObject("ADODB.Recordset")
									rsDivision.CursorLocation = adUseClient
									

							
									if len(trim(v_txtJobcode)) > 0 then 
										strsql =  	" SELECT '--ALL--' as fccode,'' fcName FROM DUAL UNION ALL SELECT fccode, fcname FROM job WHERE fcba = '"& v_txtFcba &"'" & _
													" AND job_forfieldstatus NOT IN ('BALANCE SHEET','BIBITAN') ORDER BY fccode"  
									else
										strsql =  	" SELECT '--ALL--' as fccode,'' fcName FROM DUAL UNION ALL SELECT fccode, fcname FROM job WHERE fcba = '"& Session("fcba") &"'" & _
													" AND job_forfieldstatus NOT IN ('BALANCE SHEET','BIBITAN') ORDER BY fccode"
									end if
									
									
									rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsDivision.eof then
										DO while not rsDivision.eof
										if v_txtJobcode = rsDivision("fccode") then
								%>			<option value="<%=rsDivision("fccode")%>" Selected><%=rsDivision("fccode")%> : <%=rsDivision("fcname")%></option>
										<%else%>
											<option value="<%=rsDivision("fccode")%>"><%=rsDivision("fccode")%> : <%=rsDivision("fcname")%></option>
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
							<select name="txtDivision">	
								<option value="--SUMMARY--">--SUMMARY--</option> 
								<%
									Dim rsDivision
									Set rsDivision = Server.CreateObject("ADODB.Recordset")
									rsDivision.CursorLocation = adUseClient
									

							
									if len(trim(v_txtDivision)) > 0 then 
										strsql =  	" SELECT '--ALL--' as sectionname, '' fcName from DUAL UNION ALL SELECT DISTINCT emp.sectionname, sc.fcname" & _
													"    FROM employee emp" & _
													"         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
													"   WHERE fcba = '"& v_txtFcba &"' order by sectionname" 
									else
										strsql =  	" SELECT '--ALL--' as sectionname, '' fcName FROM DUAL UNION ALL SELECT DISTINCT emp.sectionname, sc.fcname" & _
													"    FROM employee emp" & _
													"         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
													"   WHERE fcba = '"& Session("fcba") &"' order by sectionname"
									end if
									
									
									rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsDivision.eof then
										DO while not rsDivision.eof
										if v_txtDivision = rsDivision("sectionname") then
								%>			<option value="<%=rsDivision("sectionname")%>" Selected><%=rsDivision("sectionname")%> : <%=rsDivision("fcname")%></option>
										<%else%>
											<option value="<%=rsDivision("sectionname")%>"><%=rsDivision("sectionname")%> : <%=rsDivision("fcname")%></option>
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
							<select name="txtPlantingDate">	
									<option value="--SUMMARY--">--SUMMARY--</option>
									<%
										Dim rsPlantingDate
										Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
										rsPlantingDate.CursorLocation = adUseClient
								
										if len(trim(v_txtPlantingDate)) > 0 then 
											strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
										else
											strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
										end if
										
										
										rsPlantingDate.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
										if not rsPlantingDate.eof then
											DO while not rsPlantingDate.eof
											if v_txtPlantingDate = rsPlantingDate("PlantingDate") then
									%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
											<%else%>
												<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
											<%end if%>
									<%			rsPlantingDate.movenext
											loop
										End if
										rsPlantingDate.Close
										set rsPlantingDate = nothing
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
										strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
									else
										strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
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
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>
		
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="30" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<h4>PT. <%=varCompanyName%></h4>
						<h5><%=varFcbaName%></h5>
					</div>
					<div id="subheader" class="centered">
						<h2>LAPORAN RENCANA KERJA BULANAN RAWAT</h2>
						<br/>
						<h4>
							PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%>
						</h4> 

					</div>
					</div>
				</th>
			</tr>
		<tr> 
			<th rowspan="3" style="color:#fff;">NO</th>
			<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
				<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
			<% 			
		    end if %>
			
			<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
			<% 			
		    end if %>
			
			<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
			<%
		    end if %>
			
			<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
			<% 		
		    end if %>
			
			<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
			<% 		
		    end if %>
			
			<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
			    <th rowspan="3" style="background-color:#00A300;color:#fff;">Unit Kerja</th>
			<% 		
		    end if %>
			
			<th rowspan="3" style="color:#fff;">Tahun Tanam</th>
			<th rowspan="3" style="color:#fff;">Nomor RKB</th>
			<th rowspan="3" style="color:#fff;">Kode Pekerjaan</th>
			<th rowspan="3" style="color:#fff;">Nama Pekerjaan</th>
			<th colspan="5" style="color:#fff;">Lokasi</th>
			<th colspan="5" style="color:#fff;">Tenaga Kerja</th>
			<th colspan="6" style="color:#fff;">Pemakaian Bahan/ Alokasi Unit</th> 
			<th rowspan="3" style="color:#fff;">Total Biaya (HK + Bahan)</th>
			<th rowspan="3" style="color:#fff;">Cost/Satuan</th>
			<th rowspan="3" style="color:#fff;">Terakhir Dikerjakan</th>
		</tr>
		<tr>	
			<th rowspan="2" style="color:#fff;">Blok</th>
			<th rowspan="2" style="color:#fff;">Luas</th>
			<th rowspan="2" style="color:#fff;">Pokok</th>
			<th rowspan="2" style="color:#fff;">Luas Dikerjakan</th> 
			<th rowspan="2" style="color:#fff;">Satuan</th>
			<th colspan="3" style="color:#fff;">Harian</th>
			<th colspan="2" style="color:#fff;">Borongan / Premi</th>
			<th rowspan="2" style="color:#fff;">Nama Bahan / Unit</th>
			<th rowspan="2" style="color:#fff;">Satuan Bahan</th>
			<th rowspan="2" style="color:#fff;">Norma Bahan/HM</th>
			<th rowspan="2" style="color:#fff;">Jumlah Bahan</th>
			<th rowspan="2" style="color:#fff;">Harga Satuan(Rp)</th>
			<th rowspan="2" style="color:#fff;">Biaya Bahan(Rp)</th>
		</tr>
		<tr>
			<th style="color:#fff;">Norma HK/HA/KG</th>
			<th style="color:#fff;">Jumlah HK</th>
			<th style="color:#fff;">Jumlah Biaya (Rp)</th>
			<th style="color:#fff;">Harga Satuan</th>	
			<th style="color:#fff;">Jumlah Biaya Borongan/Premi (Rp)</th>	
		</tr>
	</thead>
	<tbody>
<%if rsPayroll.BOF and rsPayroll.EOF then%>
		<tr>
			<td colspan="27" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 0
		rows = 1
		while not rsPayroll.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
		
		totalbiaya = cdbl(to_nol(rsPayroll("totalamount"))) + cdbl(to_nol(rsPayroll("boronganpremi"))) + cdbl(to_nol(rsPayroll("totalamountstock"))) 
		If cdbl(to_nol(rsPayroll("volume"))) = 0 Then costsatuan = 0 Else costsatuan = cdbl(to_nol(totalbiaya)) / cdbl(to_nol(rsPayroll("volume")))
%>
		<tr class="<%=cClass%>">
			<% if rows = 1 then
					sqlrow = "  select count(plantingyear) rowsd  from ( " & strPayroll & " ) where plantingyear = '" & rsPayroll("plantingyear") & "'"
					' response.write sqlrow
					Set rsrow = Server.CreateObject("ADODB.Recordset")
					rsrow.Open sqlrow, connstr, adOpenStatic
					if not rsrow.EOF then
						temprows = CDbl(rsrow("rowsd"))
					else
						temprows = 0
					end if
				end if 
			no = NO+1
			%>
			<td align="center"><%=No%></td>
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% tempPlantation = rsPayroll("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsPayroll("Plantation") then %>
						<td align="center">
							<%=rsPayroll("Plantation")%>
							<!--<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							
							</a>-->
						</td>
						<% SprevPlantation = rsPayroll("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsPayroll("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsPayroll("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsPayroll("GroupRegion") then %>
						<td align="center">
							<%=rsPayroll("groupregion")%>
							<!--<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							
							</a>-->
						</td>
						<% SprevGroupRegion = rsPayroll("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsPayroll("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsPayroll("Region") then %>
						<td align="center">
							<%=rsPayroll("region")%>
							<!--<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							
							</a>-->
						</td>
						<% SprevRegion = rsPayroll("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPayroll("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsPayroll("Company")
					   tempfcba = "--ALL--"
					   tempownership = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsPayroll("Company") then %>
						<td align="center">
							<%=rsPayroll("company")%>					
					<!--<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">					
						</a>-->
						</td>
						<% SprevCompany = rsPayroll("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPayroll("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsPayroll("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsPayroll("fcba")
					   tempownership = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsPayroll("fcba") then %>
						<td align="center">
			    <%=rsPayroll("fcba")%>			
			<!--<a href="p_ProductionMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtownership=<%=tempownership%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
			</a>-->
						</td>
						<% sprevFcba = rsPayroll("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				 
				<% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> "" then %>
				   <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsPayroll("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsPayroll("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsPayroll("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsPayroll("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsPayroll("fcba")
					   else
					       tempfcba = ""
					   end if
					   if txtownership <> "--SUMMARY--"  and txtownership <> "" then
					      tempownership = rsPayroll("ownership")
					   else
					      tempownership = ""
                       end if
					   if v_txtDivision <> "--SUMMARY--"  and v_txtDivision <> "" then
					      tempDivision = rsPayroll("unitkerja")
					   else
					      tempDivision = ""
                       end if					   
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevDivision <> rsPayroll("unitkerja") then %>
					    <td align="center"> 
							<%=rsPayroll("unitkerja")%> 
						</td>
						<% sprevDivision = rsPayroll("unitkerja") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %> 
			<%if rows = 1 then%>
				<td align="center" Rowspan="<%=0+temprows%>"><%=rsPayroll("plantingyear")%></td> 
			<%end if%>
			<td align="center">
			<a href="..\MonthlyWorkPlan.asp?transno=<%=rsPayroll("rkbno")%>&fcba=<%=rsPayroll("fcba")%>&txtModule=<%="Estate"%>&txtTable=<%="MonthlyWorkPlan"%>%>">
					<%=rsPayroll("rkbno")%>				
			</a>
			</td>
			<td align="center"><%=rsPayroll("activitycode")%></td>
			<td align="left"><%=rsPayroll("activityname")%></td> 
			<td align="center"><%=rsPayroll("fieldcode")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("luas")), 0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("pokok")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("volume")),0)%></td> 
			<td align="center"><%=rsPayroll("uom_job")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("norma_hk")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("mandays")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("totalamount")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("normapremi")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("boronganpremi")),0)%></td>
			<td align="left"><%=rsPayroll("stockname")%></td>
			<td align="center"><%=rsPayroll("uom_stock")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("normabahan")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("qty")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("ratestock")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPayroll("totalamountstock")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(totalbiaya),0)%></td> 
			<td align="right"><%=formatnumber( cdbl(costsatuan),0)%></td>
			<td align="right"></td> 
		</tr>
<%	if rows = temprows then
			noline			= noline		+ 1
			nLuasDikerjakan = nLuasDikerjakan + cdbl(rsPayroll("luas_dikerjakan"))
			nMandays 		= nMandays 	+ cdbl(rsPayroll("mandays"))
			nTotalamount 	= nTotalamount	+ cdbl(rsPayroll("totalamount"))
			nBoronganpremi 	= nBoronganpremi + cdbl(rsPayroll("boronganpremi"))
			nQty	 		= nQty + cdbl(rsPayroll("qty"))
			nTotalamountstock = nTotalamountstock 	+ cdbl(rsPayroll("totalamountstock"))
			nTotalbiaya 	= nTotalbiaya + cdbl(totalbiaya)
			nCostsatuan 	= nCostsatuan + cdbl(costsatuan) 
			rows = 1
			
	%>
		<tr style="background-color:#F4F4A0;">	 
			<td align="center" colspan="<%= 10+ nSpanTotal %>" ><strong>Sub Total</strong></td> 
			<td align="right"><strong><%=formatnumber(nLuasDikerjakan,0)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(nMandays,0)%></strong></td> 
			<td align="right"><strong><%=formatnumber(nTotalamount,0)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(nBoronganpremi,0)%></strong></td>
			<td align="right" colspan="3"></td>
			<td align="right"><strong><%=formatnumber(nQty,0)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(nTotalamountstock,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nTotalbiaya,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nCostsatuan,0)%></strong></td>
			<td align="right"><strong></td>
		</tr>
		<tr>
			<td align="center" colspan="<%= 10+ nSpanTotal %>" ></td>
		</tr>
		<%	
			noline = 1
			nLuasDikerjakan = cdbl("0")
			nMandays = cdbl("0")
			nTotalamount = cdbl("0")
			nBoronganpremi = cdbl("0")
			nTotalamountstock = cdbl("0")
			nQty = cdbl("0")
			nTotalbiaya = cdbl("0")
			nCostsatuan = cdbl("0")
			 
		else
		' cari sub total
			noline			= noline		+ 1
			nLuasDikerjakan = nLuasDikerjakan + cdbl(rsPayroll("luas_dikerjakan"))
			nMandays 		= nMandays + cdbl(rsPayroll("mandays"))
			nTotalamount 	= nTotalamount + cdbl(rsPayroll("totalamount"))
			nBoronganpremi 	= nBoronganpremi	+ cdbl(rsPayroll("boronganpremi"))
			nQty	 		= nQty + cdbl(rsPayroll("qty"))
			nTotalamountstock = nTotalamountstock 	+ cdbl(rsPayroll("totalamountstock"))
			nTotalbiaya 	= nTotalbiaya + cdbl(totalbiaya)
			nCostsatuan 	= nCostsatuan + cdbl(costsatuan)
			rows = rows + 1
		End if
							  
			' cari grand total 
			gt_LuasDikerjakan = gt_LuasDikerjakan + cdbl(rsPayroll("luas_dikerjakan"))
			gt_Mandays 		= gt_Mandays + cdbl(rsPayroll("mandays"))
			gt_Totalamount 	= gt_Totalamount + cdbl(rsPayroll("totalamount"))
			gt_Boronganpremi 	= gt_Boronganpremi	+ cdbl(rsPayroll("boronganpremi"))
			gt_Qty	 		= gt_Qty + cdbl(rsPayroll("qty"))
			gt_Totalamountstock = gt_Totalamountstock 	+ cdbl(rsPayroll("totalamountstock"))
			gt_Totalbiaya 	= gt_Totalbiaya + cdbl(totalbiaya)
			gt_Costsatuan 	= gt_Costsatuan + cdbl(costsatuan) 				
		 
			rsPayroll.MoveNext			
			Wend
		rsPayroll.Close
		set rsPayroll = Nothing
	%> 
		<tr style="background-color:#27AE60;">	 
			<td align="center" colspan="<%= 10+ nSpanTotal %>" ><strong>Grand Total</strong></td>
			<td align="right"><strong><%=formatnumber(gt_LuasDikerjakan,0)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(gt_Mandays,0)%></strong></td> 
			<td align="right"><strong><%=formatnumber(gt_Totalamount,0)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(gt_Boronganpremi,0)%></strong></td>
			<td align="right" colspan="3"></td>
			<td align="right"><strong><%=formatnumber(gt_Qty,0)%></strong></td>
			<td align="right"></td>
			<td align="right"><strong><%=formatnumber(gt_Totalamountstock,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(gt_Totalbiaya,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(gt_Costsatuan,0)%></strong></td>
			<td align="right"><strong></td>
		</tr> 
	</tbody>
</table> 
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
   </body>
   </html>