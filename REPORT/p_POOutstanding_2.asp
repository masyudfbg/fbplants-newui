<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtSupplierCode, v_txtPOCode, v_txtItemCode,v_txtReceivedQty, NO,v_txtOSRQTY
Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType

dim Vmonth,Vday
	if len(month(date))=1 then
	Vmonth = "0" & month(date)
	else
	Vmonth = month(date)
	end if
	if len(day(date))=1 then
	Vday = "0" & day(date)
	else
	Vday = day(date)
	end if
	
v_txtGroupPlantation  = Request("txtGroupPlantation")
v_txtPlantation  	  = Request("txtPlantation")
v_txtGroupRegion  	  = Request("txtGroupRegion")
v_txtRegion  		  = Request("txtRegion")
v_txtcompany  		  = Request("txtCompany")
v_txtfcba 			  = Request("txtfcba")
			
v_txtDateFrom		=	trim(Request("txtDateFrom"))
v_txtDateTo			=	trim(Request("txtDateTo"))
v_txtSupplierCode	=	trim(Request("txtSupplierCode"))
v_txtPOCode			=	trim(Request("txtPOCode"))
v_txtItemCode		=	trim(Request("txtItemCode"))
v_txtOSRQTY			=	trim(Request("OUTSTANDINGRECEIVEDQTY"))
v_txtReceivedQty	=	trim(Request("txtReceivedQty"))
deliverydestination	=	trim(Request("deliverydestination"))
ScriptName			=	Request.ServerVariables("SCRIPT_NAME")

' If Len(Trim(v_txtDateFrom)) = 0 Then
	' v_txtDateFrom = textDate(Date)
' Else
	' v_txtDateFrom = Request("txtDateFrom")
' End If
' If Len(Trim(v_txtDateTo)) = 0 Then
	' v_txtDateTo = textDate(Date)
' Else
	' v_txtDateTo = Request("txtDateTo")
' End If

Dim nSpanTotal : nSpanTotal = 6
			Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
			
			strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
								" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
								" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
								" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
								" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
								" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
								
			wherecons	=	" where company is not null "

			orderCons = " Order by GroupPlantation"
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

				' Added By Sendy untuk validasi menu 11/19/2018
				Dim Menu : Menu = "N"
				Dim oRecordSetMenu, oSQLQueryMenu
				oSQLQueryMenu = " SELECT 'Y' istrue FROM users_tab " & _
							" WHERE fccode = '" & Session("CodeName") & "' AND MENU_PO = 'TRUE'" 
				' response.write oSQLQueryMenu
				set oRecordSetMenu = server.CreateObject("Adodb.recordset")
				oRecordSetMenu.Open oSQLQueryMenu, connstr
				If Not oRecordSetMenu.Eof then
					Menu = oRecordSetMenu("istrue")
				Else
					Menu = "N"
				End if
				' End
%>
<html>

<head>
<title>ePlaS Reporting - PO Outstanding</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<!-- Fuction Checked Lebih Kurang -->
<style>
   .icheckbox_square-aero,
	input[type="radio"].iradio_square-aero {
	padding: 0;
	width: 24px;
	height: 24px;
	background-image: url(images/aero2.png);
	background-repeat:no-repeat;
	border: none;
	border-bottom: none;
	cursor: pointer;
	}
		input[type="radio"].iradio_square-aero {
		display: block;
		margin: 0;
		opacity:0;
		}
		.icheckbox_square-aero{
		 display: block;
		 transition: none;
		 }
		.icheckbox_square-aero:hover{
		 border-bottom: none;
		 background-image: url(images/aero.png);
		 }
	
		.checkedkurang {
		 background-position: -24px 0;
		 cursor: default;
		}
		.checkedlebih {
		 background-position: -0px 0;
		 cursor: default;
		 }
</style>
<script>
function displayPhrase()
	{
		var LEBIH = document.getElementById("lebihcondition")
		var KURANG = document.getElementById("kurangcondition")
		if(KURANG.checked) {
			LEBIH.style.display = 'block';
			KURANG.style.display = 'none';
			document.getElementById("a-cond1").className  = 'icheckbox_square-aero checkedkurang';
		}else if (LEBIH.checked){
			LEBIH.style.display = 'none';
			KURANG.style.display = 'block';
			document.getElementById("a-cond1").className  = 'icheckbox_square-aero checkedlebih';
		}
	}
</script>
<!-- END Fuction Checked Lebih Kurang -->	
		
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);
	</script>
<link rel="stylesheet" href="/css/jquery-ui.css" />
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
     <script>
	$(function() {
	$( ".datepicker" ).datepicker();
	});
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
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
		<form name="frmEntry" method="post" action="<%=ScriptName%>">
		<div class="printhide">
		<!--#include file="FormSelectNewSum.html"-->
		<table id="form-table" class="centered-margin">
			<tr>
				
				<td><font face="verdana" size="1">Dari</font></td>
				<td align="center">
				<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
				<% else %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
				<% end if%>
					
				</td>
				<td><font face="verdana" size="1">Sampai</font></td>
				<td align="center">
				<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
				<% else %>
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
				<% end if%>
					
				</td>
				<td align="left"><font face="Verdana" size="1">Delivery Destination</font></td>
				<td align="left">
					<select name="deliverydestination" >
						<option value=""></option> 
						<% 	Dim rsComboTT
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboTT = Server.CreateObject("ADODB.Recordset")
							rsComboTT.CursorType = 3
							rsComboTT.CursorLocation = 3
							rsComboTT.Open q, connstr
							if not rsComboTT.eof then%>
							<%
							Do While Not rsComboTT.eof
							%>
								<option value="<%=rsComboTT("fccode")%>" 
									<%if cstr(""&deliverydestination)=cstr(""&rsComboTT("fccode")) then %> selected 
									<%End If%>><%=rsComboTT("fccode")%>
								</option> 
							<%rsComboTT.movenext
							loop
							End if
							rsComboTT.Close
							set rsComboTT = Nothing
							q=""
						%>
					</select>
				</td>
				<td><font face="verdana" size="1">NO PO</font></td><td><input type="text" name="txtPOCode" value="<%=v_txtPOCode%>" /></td>
				<td><font face="verdana" size="1">KODE BARANG</font></td><td><input type="text" name="txtItemCode" value="<%=v_txtItemCode%>" /></td>
				<td>
				<font face="verdana" size="1">Type Laporan</font>
				</td>
				<td>
					<select name="OUTSTANDINGRECEIVEDQTY">
					<% if v_txtOSRQTY = "1" then%>
						<option value=""></option>
						<option value="1" selected>Out Standing</option>
						<option value="0">PO Closing</option>
					<%elseif v_txtOSRQTY = "0" then%>
						<option value=""></option>
						<option value="1">Out Standing</option>
						<option value="0" selected>PO Closing</option>
					<%else%>
						<option value="" selected></option>
						<option value="1">Out Standing</option>
						<option value="0">PO Closing</option>
					<%end if%>
					</select>
				</td>
				<% if  v_txtOSRQTY = "1" then%>
				<td>
				
				<% 
				dim checked_cond1,display_checked,display_checked2
					if request("LebihKurang") = "LEBIH" then 
						checked_cond1 = "checkedlebih"
						display_checked = "display:none;"
						display_checked2 = "display:block;"
					elseif request("LebihKurang") = "KURANG" then
						checked_cond1 = "checkedkurang"
						display_checked = "display:block;"
						display_checked2 = "display:none;"
					else
						checked_cond1 = "checkedlebih"
						display_checked = "display:none;"
						display_checked2 = "display:block;"
					end if
				%>
				<a id="a-cond1" class="icheckbox_square-aero <%=checked_cond1%>" >
				<input id="lebihcondition" onclick="displayPhrase()" class="iradio_square-aero " type="radio" name="LebihKurang" value="LEBIH" checked style="<%=display_checked%>">
				<input id="kurangcondition" onclick="displayPhrase()" class="iradio_square-aero " type="radio" name="LebihKurang" value="KURANG" style="<%=display_checked2%>">
				</a>
				</td>
				<td>
				<div id="sisa"><font face="verdana" size="1">SISA</font></div>
				</td>
				<td>
				<input style="width:50px; text-align:center;" type="text" name="txtReceivedQty" value="<%=v_txtReceivedQty%>" /></td>
				<%end if%>
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		</div>
		</form>
	</div>
	
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="<%=nSpanTotal+14%>" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4>
				<%Else%>
					<h4>&nbsp;</h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					<h5><%=FcbaName%></h5>
				<%Else%>
					<h5>&nbsp;</h5>
				<% End if%>
					</div> 
	 
					<div id="subheader" class="centered">
						<h2>LAPORAN
						<%if v_txtOSRQTY = "1" then%>
						PO OUTSTANDING
						<%elseif v_txtOSRQTY = "0" then%>
						PO CLOSING
						<%else%>
						
						<%end if%>
						</h2>	
<br/>						
					<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
						<h5>TANGGAL: <%=v_txtdateto%></h5>
					<% else %>
						<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
					<% end if %>	
<br/>					
					</div>
				</div>
			</th>
		</tr>
     <tr>

		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		<th rowspan="2" style="color:#fff;">NO PO</th>		
		<th rowspan="2" style="color:#fff;">DELIVERY DESTINATION</th>
		<th rowspan="2" style="color:#fff;">LINENOKEY</th>
		<th rowspan="2" style="color:#fff;">TANGGAL</th>
		<th rowspan="2" style="color:#fff;">KODE BARANG</th>
		<th rowspan="2" style="color:#fff;">NAMA BARANG</th>
		<%If Menu = "Y" Then%>
		<th rowspan="2" style="color:#fff;">JUMLAH (Rp)</th>
		<%End If%>
		<th colspan="4" style="color:#fff;">QUANTITY</th>
		<th rowspan="2" style="color:#fff;">KETERANGAN</th>
		<th rowspan="2" style="color:#fff;">NO PR</th>
		<th rowspan="2" style="color:#fff;">LINE NO</th>
		<th rowspan="2" style="color:#fff;">NOTES</th>
	</tr>
	<tr>		
		<th style="color:#fff;">DI ORDER</th>
		<th style="color:#fff;">DI TERIMA</th>
		<th style="color:#fff;">DI MINTA</th>	
		<th style="color:#fff;">SISA</th>
	</tr>
	</thead>
	<tbody>
	<%


	Dim rsPOOut, strPOOut , noline
	
	strPOOut	=	" SELECT " & strFieldCons & "," & _
					"        pocode," & _
					"        deliverydestination," & _
					"        linenokey," & _
					"        podate," & _
					"        itemcode," & _
					"        fcname," & _
					"        amount," & _
					"        orderedqty," & _
					"        receivedqty," & _
					"        requestedqty," & _
					"        outstandingreceivedqty," & _
					"        otheritemdesc," & _
					"        prcode," & _
					"        prlineid," & _
					"        notes" & _
					"   FROM (SELECT data.fcba," & _
					"                data.pocode," & _
					"                data.deliverydestination," & _
					"                data.linenokey," & _
					"                data.podate," & _
					"                data.itemcode," & _
					"                stock.fcname," & _
					"                data.amount," & _
					"                data.orderedQty," & _
					"                data.receivedQty," & _
					"                data.requestedqty," & _
					"                data.approvedqty," & _
					"                data.OutstandingReceivedQty," & _
					"                data.otheritemdesc," & _
					"                data.prcode," & _
					"                data.prlineid," & _
					"                'PO without PR Information' NOTES" & _
					"           FROM (SELECT a.fcba," & _
					"                        a.pocode," & _
					"                        b.deliverydestination," & _
					"                        a.linenokey," & _
					"                        b.podate," & _
					"                        a.itemcode," & _
					"                        a.quantity orderedQty," & _
					"                        a.amount," & _
					"                        a.received receivedQty," & _
					"                        0 requestedqty," & _
					"                        0 approvedqty," & _
					"                        (a.quantity - a.received) OutstandingReceivedQty," & _
					"                        a.otheritemdesc," & _
					"                        a.prcode," & _
					"                        a.prlineid" & _
					"                   FROM purchaseorderdetails a, purchaseorder b" & _
					"                  WHERE     a.pocode = b.pocode" & _
					"                        AND a.fcba = b.fcba" & _
					"                        AND b.rowstateid = 'Ordered'" & _
					"                        AND a.prcode IS NULL) data," & _
					"                stock" & _
					"          WHERE data.itemcode = stock.fccode AND data.fcba = stock.fcba" & _
					"         UNION" & _
					"         SELECT data.fcba," & _
					"                data.pocode," & _
					"                data.deliverydestination," & _
					"                data.linenokey," & _
					"                data.podate," & _
					"                data.itemcode," & _
					"                stock.fcname," & _
					"                data.amount," & _
					"                data.orderedQty," & _
					"                data.receivedQty," & _
					"                data.requestedqty," & _
					"                data.approvedqty," & _
					"                data.OutstandingReceivedQty," & _
					"                data.otheritemdesc," & _
					"                data.prcode," & _
					"                data.prlineid," & _
					"                'PO with PR Information' NOTES" & _
					"           FROM (SELECT a.fcba," & _
					"                        a.pocode," & _
					"                        b.deliverydestination," & _
					"                        a.linenokey," & _
					"                        b.podate," & _
					"                        a.itemcode," & _
					"                        a.quantity orderedQty," & _
					"                        a.amount," & _
					"                        a.received receivedQty," & _
					"                        c.requestedqty," & _
					"                        c.approvedqty," & _
					"                        (a.quantity - a.received) OutstandingReceivedQty," & _
					"                        a.otheritemdesc," & _
					"                        a.prcode," & _
					"                        a.prlineid" & _
					"                   FROM purchaseorderdetails a," & _
					"                        purchaseorder b," & _
					"                        purchaserequestdetails c" & _
					"                  WHERE     a.pocode = b.pocode" & _
					"                        AND a.fcba = b.fcba" & _
					"                        AND A.PRCODE = c.prcode" & _
					"                        AND A.PRLINEID = C.LINEID" & _
					"                        AND b.rowstateid in('Ordered','Confirmed','AuthorizeOnProgress')" & _
					"                        AND a.prcode IS NOT NULL) data," & _
					"                stock" & _
					"          WHERE data.itemcode = stock.fccode AND data.fcba = stock.fcba) WHERE " 
if  len(trim(v_txtDateFrom)) > 0 then
	strPOOut = strPOOut & "  podate >=	to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
	else
	strPOOut = strPOOut & "  podate >=	to_Date('" & DateAdd("YYYY",-1,date) & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	strPOOut = strPOOut &  " AND podate <= to_Date('" & v_txtDateTo & "','MM/DD/YYYY')"
	else
	strPOOut = strPOOut &  " AND podate <= to_Date('" & date & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPOCode)) > 0 then
	strPOOut = strPOOut &  " AND pocode = '" & v_txtPOCode & "'"
end if	
if  len(trim(v_txtItemCode)) > 0 then
	strPOOut = strPOOut &  " AND itemcode = '" & v_txtItemCode & "'"
end if	
if  len(trim(deliverydestination)) > 0 then
	strPOOut = strPOOut &  " AND deliverydestination = '" & deliverydestination & "'"
end if	
if  len(trim(v_txtOSRQTY)) > 0 then
	if v_txtOSRQTY = "1" then
		if request("LebihKurang") = "LEBIH" then 
			if  len(trim(v_txtReceivedQty)) > 0 then
			strPOOut = strPOOut &  " AND OUTSTANDINGRECEIVEDQTY >= '" & v_txtReceivedQty & "'"
			else
			strPOOut = strPOOut &  " AND OUTSTANDINGRECEIVEDQTY >= '" & v_txtOSRQTY & "'"
			end if	
		elseif request("LebihKurang") = "KURANG" then
			if  len(trim(v_txtReceivedQty)) > 0 then
			strPOOut = strPOOut &  " AND OUTSTANDINGRECEIVEDQTY <= '" & v_txtReceivedQty & "' AND OUTSTANDINGRECEIVEDQTY > 0 " 
			else
			strPOOut = strPOOut &  " AND OUTSTANDINGRECEIVEDQTY <= '" & v_txtOSRQTY & "' AND OUTSTANDINGRECEIVEDQTY > 0 "
			end if	
		else
			if  len(trim(v_txtReceivedQty)) > 0 then
			strPOOut = strPOOut &  " AND OUTSTANDINGRECEIVEDQTY >= '" & v_txtReceivedQty & "'"
			else
			strPOOut = strPOOut &  " AND OUTSTANDINGRECEIVEDQTY >= '" & v_txtOSRQTY & "'"
			end if	
		end if
		
	else
	
	strPOOut = strPOOut &  " AND OUTSTANDINGRECEIVEDQTY = '" & v_txtOSRQTY & "'"
	end if
end if
	
	orderCons = orderCons & ", podate, pocode, itemcode, linenokey"
	Set rsPOOut = Server.CreateObject("ADODB.Recordset")
	strPOOut = " select GroupPlantation " & selectCons & "" & _
	           "        , pocode" & _
	           "        , deliverydestination" & _
	           "        , linenokey" & _
	           "        , podate" & _
	           "        , itemcode" & _
	           "        , fcname" & _
	           "        , sum(amount) amount" & _
	           "        , sum(orderedqty) orderedqty" & _
	           "        , sum(receivedqty) receivedqty" & _
	           "        , sum(requestedqty) requestedqty" & _
	           "        , sum(outstandingreceivedqty) outstandingreceivedqty" & _
	           "        , otheritemdesc" & _
	           "        , prcode" & _
	           "        , prlineid" & _
	           "        , notes" & _
	           " From (" & strPOOut & " )" & whereCons & "" & _
			   " group by GroupPlantation" & selectCons & "" & _
			   "          , pocode" & _
			   "          , deliverydestination" & _
	           "          , linenokey" & _
	           "          , podate" & _
	           "          , itemcode" & _
	           "          , fcname" & _
	           "          , otheritemdesc" & _
	           "          , prcode" & _
	           "          , prlineid" & _
	           "          , notes" & "" & _
			   orderCons
	
	' Response.Write strPOOut
	rsPOOut.Open strPOOut, connstr, adOpenStatic
	
	
	
	if (rsPOOut.bof or rsPOOut.eof) then
	%>
	<tr>
			<td align="left" colspan="<%=nSpanTotal+14%>"><strong>DATA TIDAK DITEMUKAN</strong></td>
	</tr>	
	<%
	elseif Not rsPOOut.eof then		
		while Not rsPOOut.eof  
		if cdbl(rsPOOut("outstandingReceivedQty")) = 0 then
		%>          
		  <tr style="background:#DEFFC4;">
		<%else%>
		  <tr style="background:#FFC4C4;">
		<%end if%>
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsPOOut("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPOOut("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPOOut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsPOOut("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPOOut("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPOOut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPOOut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsPOOut("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPOOut("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPOOut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPOOut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsPOOut("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsPOOut("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPOOut("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPOOut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPOOut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsPOOut("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsPOOut("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsPOOut("Company")
					   tempfcba = "--ALL--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPOOut("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPOOut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPOOut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsPOOut("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsPOOut("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsPOOut("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsPOOut("fcba")
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
						    <%=rsPOOut("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
			<td align="center"><%=rsPOOut("pocode")%></td>
			<td align="center"><%=rsPOOut("deliverydestination")%></td>
			<td align="right"><%=rsPOOut("linenokey")%></td>
			<td align="right"><%=FbpNum2(rsPOOut("podate"))%></td>
			<td align="left"><%=rsPOOut("itemcode")%></td>
			<td align="left"><%=rsPOOut("fcname")%></td>
			<%If Menu = "Y" Then%>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("amount")),2)%></td>
			<% end if %>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("orderedqty")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("receivedqty")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("requestedqty")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("outstandingReceivedQty")),2)%></td>
			<td align="left"><%=rsPOOut("otheritemdesc")%></td>
			<td align="left"><%=rsPOOut("prcode")%></td>
			<td align="right"><%=rsPOOut("prlineid")%></td>
			<td align="left"><%=rsPOOut("notes")%></td>
		  </tr>	
		      
<%

			rsPOOut.MoveNext
		wend
	end if
%>
	</tbody>
</table>
<!--#include file="bottom.Asp"-->