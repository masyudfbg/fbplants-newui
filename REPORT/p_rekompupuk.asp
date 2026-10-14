<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="../css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="../css/jquery-ui.css" />
<script language="JavaScript" src='../js/x.js'></script>
<script language="JavaScript" src='../js/xtableheaderfixed.js'></script>
<script type="text/javascript" src="dtree.js"></script>
<script src="../js/jquery-1.9.1.js"></script>
<script src="../js/jquery-ui.js"></script>
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
<script language="JavaScript" src="../js/date-picker.js"></script> 
<title>ePlaS Reporting - Rekom Pupuk</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

' v_txtCUSTOMERCODE= trim(request("periodyear"))
division =request("division")
transdate_From =request("transdate_From")
transdate_To =request("transdate_To")
pemeriksa=request("pemeriksa")
line=request("line")
fieldcode=request("fieldcode")
empcode = Request("empcode")
txtlineno = request("txtlineno")
supervision =request("supervision")
gangcode =request("gangcode")
janjangpanen=request("janjangpanen")
buahtinggal=request("buahtinggal")
brondolantinggal=request("brondolantinggal")
overprunning = Request("overprunning")
pelepahsengkleh = request("pelepahsengkleh")
pelepahsusun=request("pelepahsusun")

v_txtfcba 			= Request("txtfcba")
v_txtDivision		= Request("txtDivision")
v_txtBlok			= Request("txtBlok")
v_txtPlantingDate	= Request("txtPlantingDate")
fcname 				= Request("fcname")
fcCode 				= Request("fcCode")

txtYear = Request("txtYear")
txtMonth = Request("txtMonth")

txttdate = Request("txttdate")
		If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
		If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))

		txttdate = textDate(txtMonth&"/01/"&txtYear)

		varDefaultBa = Session("fcBA")

		If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
			v_txtFcba = varDefaultBa
		End If	
			
	       if  v_txtfcba <> "" then
		       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
			end if
		    ' selectCons = selectCons & ",fcba"
		    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
			q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
		else
		    nSpanTotal = nSpanTotal - 1 
		end if
	    
	   if v_txtDivision <>  "--ALL--"  then
			whereCons = whereCons & " AND fp.DIVISION = '" & v_txtDivision & "'"
	   else
	   nSpanTotal = nSpanTotal - 1  
	   end if   


		if v_txtPlantingDate <>  "--ALL--" then
			  whereCons = whereCons & " AND to_char(field_api.getvalid_plantingyear (fp.fccode, fp.fcba, LAST_DAY (TO_DATE ('"& txtyear &"', 'YYYY')))) ='" & v_txtplantingdate & "'"
		else
		   nSpanTotal = nSpanTotal - 1 
		end if

	
		if v_txtBlok <>  "--ALL--" then
			  whereCons = whereCons & " AND fp.fccode ='" & v_txtBlok & "'"
		else
			nSpanTotal = nSpanTotal - 1 
		end if


txtSQL = 	"	SELECT fp.fccode," & _
			"       fp.fcname," & _
			"       fp.division," & _
			"       fp.plantingdate," & _
			"       fp.hectarageplanted," & _
			"       fp.totalstand," & _
			"       fp.soiltype," & _
			"       fp.plantingmaterial," & _
			"       rp.fertilizercode," & _
			"       rp.fertilizername," & _
			" 		to_char(sysdate,'yyyy') - to_char(plantingdate,'yyyy') as Umur," & _
			"       NVL (rp.dos_jan, 0) DOS_JAN," & _
			"       NVL (rp.dos_feb, 0) DOS_FEB," & _
			"       NVL (rp.dos_mar, 0) DOS_MAR," & _
			"       NVL (rp.dos_apr, 0) DOS_APR," & _
			"       NVL (rp.dos_mei, 0) DOS_MEI," & _
			"       NVL (rp.dos_jun, 0) DOS_JUN," & _
			"       NVL (rp.dos_jul, 0) DOS_JUL," & _
			"       NVL (rp.dos_aug, 0) DOS_AUG," & _
			"       NVL (rp.dos_sep, 0) DOS_SEP," & _
			"       NVL (rp.dos_okt, 0) DOS_OKT," & _
			"       NVL (rp.dos_nov, 0) DOS_NOV," & _
			"       NVL (rp.dos_des, 0) DOS_DES," & _
			"       NVL (rp.kg_jan, 0) KG_JAN," & _
			"       NVL (rp.kg_feb, 0) KG_FEB," & _
			"       NVL (rp.kg_mar, 0) KG_MAR," & _
			"       NVL (rp.kg_apr, 0) KG_APR," & _
			"       NVL (rp.kg_mei, 0) KG_MEI," & _
			"       NVL (rp.kg_jun, 0) KG_JUN," & _
			"       NVL (rp.kg_jul, 0) KG_JUL," & _
			"       NVL (rp.kg_agu, 0) KG_AUG," & _
			"       NVL (rp.kg_sep, 0) KG_SEP," & _
			"       NVL (rp.kg_okt, 0) KG_OKT," & _
			"       NVL (rp.kg_nov, 0) KG_NOV," & _
			"       NVL (rp.kg_des, 0) KG_DES" & _
			"  FROM field_pupuk fp" & _
			"       INNER JOIN rekom_pupuk rp" & _
			"          ON (fp.fccode = rp.fccode AND fp.fcba = rp.fcba) " & whereCons & " order by fp.division, fp.fccode, fp.plantingdate" 
		 ' response.write txtSQL
Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSQL, connstr

if  len(trim(v_txtFcba)) > 0 then
    txtSQL = txtSQL & " and fp.fcba = '"& v_txtfcba &"'"
end if
if  len(trim(v_txtDivision)) > 0 then
    txtSQL = txtSQL & " and fp.division = '"& v_txtDivision &"'"
end if
if  len(trim(v_txtBlok)) > 0 then
    txtSQL = txtSQL & " and fp.fcCode = '"& v_txtBlok &"'"
end if
if  len(trim(plantingdate)) > 0 then
    txtSQL = txtSQL & " and fp.plantingdate >=to_Date('" & v_txtplantingdate & "','YYYY')"
end if


txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = ""
end if
'& " group by WIDATE, WIDIVISION"

' response.Write(txtsql)
' response.Flush()
%>
 <body>
		<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<form name="frmAR" method="post" action="<%=ScriptName%>">
							<!-- Added Sendy 09/05/2018 -->
						<tr>
							<th>KEBUN</th>
							<th>DIVISION</th>
							<th>BLOK</th>
							<th>TAHUN TANAM</th>
							<th>&nbsp;</th>
						</tr>
							<td>
								<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
								<%
								Dim rsFcba
								Set rsFcba = Server.CreateObject("ADODB.Recordset")
								rsFcba.CursorLocation = adUseClient
								strsql =   " select fccode fcba from businessunit " & _
											" where " & _
											" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
								  
								
								rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
								if not rsFcba.eof then
									DO while not rsFcba.eof
										if v_txtFcba = rsFcba("fcba") then
										%>			
											<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
										<% Else %>	
											<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
										<% End If
										
										rsFcba.movenext
									loop
								End if
								rsFcba.Close
								set rsFcba = Nothing
								%>
								</select>	
							</td>
							<td align="center">
								<select name="txtDivision">	
									<%
										Dim rsDivision
										Set rsDivision = Server.CreateObject("ADODB.Recordset")
										rsDivision.CursorLocation = adUseClient
										
										if len(trim(v_txtDivision)) > 0 then 
											strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
										else
											strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
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
								<select name="txtPlantingDate">	
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
							<!-- End Added -->
						

						<td><input type="submit" value="View" /></td>
						</form>
					</tr>
				</table>
			</div>
 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="33" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>REKOM PUPUK</h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="33" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th rowspan="2">Afdeling</th>
					<th rowspan="2">Blok</th>
					<th rowspan="2">Tahun Tanam</th>
					<th rowspan="2">Luas (Ha)</th>
					<th rowspan="2">Jumlah Pokok</th>
					<th rowspan="2">Umur (BST)</th>
					<th rowspan="2">Jenis Pupuk</th>
					<th colspan="13">Dosis Pupuk</th>
					<th colspan="13">Jumlah Pupuk</th>
				</tr>
				<tr>
					<th>Jan</th>
					<th>Feb</th>
					<th>Mar</th>
					<th>Apr</th>
					<th>Mei</th>
					<th>Jun</th>
					<th>Jul</th>
					<th>Aug</th>
					<th>Sep</th>
					<th>Okt</th>
					<th>Nov</th>
					<th>Des</th>
					<th>Total</th>
			
					<th>Jan</th>
					<th>Feb</th>
					<th>Mar</th>
					<th>Apr</th>
					<th>Mei</th>
					<th>Jun</th>
					<th>Jul</th>
					<th>Aug</th>
					<th>Sep</th>
					<th>Okt</th>
					<th>Nov</th>
					<th>Des</th>
					<th>Total</th>
				</tr>
            </thead>
        <tbody>
	<%
	if Not rscat.eof then
		  v_dos_jan = 0
		  v_dos_feb = 0
		  v_dos_mar = 0
		  v_dos_apr = 0
		  v_dos_mei = 0
		  v_dos_jun = 0
		  v_dos_jul = 0
		  v_dos_aug = 0
		  v_dos_sep = 0
		  v_dos_okt = 0
		  v_dos_nov = 0
		  v_dos_des = 0
		  v_totaldosis = 0

		  v_kg_jan  = 0
		  v_kg_feb  = 0
		  v_kg_mar  = 0
		  v_kg_apr  = 0
		  v_kg_mei  = 0
		  v_kg_jun  = 0
		  v_kg_jul  = 0 
		  v_kg_aug  = 0
		  v_kg_sep  = 0
		  v_kg_okt  = 0
		  v_kg_nov  = 0
		  v_kg_des  = 0 
		  v_totalkg = 0

	While Not rscat.EOF
		
		TotalDosis = cdbl(to_nol(rscat("DOS_JAN"))) + cdbl(to_nol(rscat("DOS_FEB"))) + cdbl(to_nol(rscat("DOS_MAR"))) + cdbl(to_nol(rscat("DOS_APR"))) + cdbl(to_nol(rscat("DOS_MEI"))) + cdbl(to_nol(rscat("DOS_JUN"))) + cdbl(to_nol(rscat("DOS_JUL"))) + cdbl(to_nol(rscat("DOS_AUG"))) + cdbl(to_nol(rscat("DOS_SEP"))) + cdbl(to_nol(rscat("DOS_OKT"))) + cdbl(to_nol(rscat("DOS_NOV"))) + cdbl(to_nol(rscat("DOS_DES")))

		TotalKg = cdbl(to_nol(rscat("KG_JAN"))) + cdbl(to_nol(rscat("KG_FEB"))) + cdbl(to_nol(rscat("KG_MAR"))) + cdbl(to_nol(rscat("KG_APR"))) + cdbl(to_nol(rscat("KG_MEI"))) + cdbl(to_nol(rscat("KG_JUN"))) + cdbl(to_nol(rscat("KG_JUL"))) + cdbl(to_nol(rscat("KG_AUG"))) + cdbl(to_nol(rscat("KG_SEP"))) + cdbl(to_nol(rscat("KG_OKT"))) + cdbl(to_nol(rscat("KG_NOV"))) + cdbl(to_nol(rscat("KG_DES")))  
	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("division")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("plantingdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("hectarageplanted")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("totalstand")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("umur")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fertilizername")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_JAN")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_FEB")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_MAR")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_APR")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_MEI")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_JUN")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_JUL")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_AUG")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_SEP")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_OKT")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_NOV")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOS_DES")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TotalDosis%></td>

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_JAN")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_FEB")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_MAR")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_APR")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_MEI")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_JUN")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_JUL")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_AUG")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_SEP")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_OKT")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_NOV")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("KG_DES")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TotalKg%></td>
		</tr>
	  <%
	  v_dos_jan =v_dos_jan+trim(cdbl(rsCat("DOS_JAN")))
	  v_dos_feb =v_dos_jan+trim(cdbl(rsCat("DOS_FEB")))
	  v_dos_mar =v_dos_jan+trim(cdbl(rsCat("DOS_MAR")))
	  v_dos_apr =v_dos_jan+trim(cdbl(rsCat("DOS_APR")))
	  v_dos_mei =v_dos_jan+trim(cdbl(rsCat("DOS_MEI")))
	  v_dos_jun =v_dos_jan+trim(cdbl(rsCat("DOS_JUN")))
	  v_dos_jul =v_dos_jan+trim(cdbl(rsCat("DOS_JUL")))
	  v_dos_aug =v_dos_jan+trim(cdbl(rsCat("DOS_AUG")))
	  v_dos_sep =v_dos_jan+trim(cdbl(rsCat("DOS_SEP")))
	  v_dos_okt =v_dos_jan+trim(cdbl(rsCat("DOS_OKT")))
	  v_dos_nov =v_dos_jan+trim(cdbl(rsCat("DOS_NOV")))
	  v_dos_des =v_dos_jan+trim(cdbl(rsCat("DOS_DES")))
	  v_totaldosis = v_totaldosis + TotalDosis

	  v_kg_jan =v_kg_jan+trim(cdbl(rsCat("KG_JAN")))
	  v_kg_feb =v_kg_feb+trim(cdbl(rsCat("KG_FEB")))
	  v_kg_mar =v_kg_mar+trim(cdbl(rsCat("KG_MAR")))
	  v_kg_apr =v_kg_apr+trim(cdbl(rsCat("KG_APR")))
	  v_kg_mei =v_kg_mei+trim(cdbl(rsCat("KG_MEI")))
	  v_kg_jun =v_kg_jun+trim(cdbl(rsCat("KG_JUN")))
	  v_kg_jul =v_kg_jul+trim(cdbl(rsCat("KG_JUL")))
	  v_kg_aug =v_kg_aug+trim(cdbl(rsCat("KG_AUG")))
	  v_kg_sep =v_kg_sep+trim(cdbl(rsCat("KG_SEP")))
	  v_kg_okt =v_kg_okt+trim(cdbl(rsCat("KG_OKT")))
	  v_kg_nov =v_kg_nov+trim(cdbl(rsCat("KG_NOV")))
	  v_kg_des =v_kg_des+trim(cdbl(rsCat("KG_DES")))
	  v_totalkg = v_totalkg + TotalKg

	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='33'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="7"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_jan%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_feb%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_mar%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_apr%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_mei%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_jun%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_jul%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_aug%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_sep%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_okt%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_nov%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_dos_des%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_totaldosis%></td>

		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_jan%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_feb%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_mar%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_apr%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_mei%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_jun%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_jul%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_aug%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_sep%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_okt%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_nov%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_kg_des%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=v_totalkg%></td>
	</tr>
	</tbody>
</table>
 <table width="90%" class="data">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=MgrName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=KTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
       </tbody>
   </table>
   </body>
   </html>