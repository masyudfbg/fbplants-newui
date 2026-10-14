<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Harvesting Receptions</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="javascript" type="text/javascript">
				function combo_(){
					var v = document.getElementById("Vehicle");
					var c = document.getElementById("CONTRACTNUMBER");
					v.value = "";
					c.value = "";
				}
                function showHint(str)
					{
				var checker;
					if (str=="")
					  {
					  document.getElementById("checker").innerHTML="*";
					  $("#Vehicle").removeClass("error_val");
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  checker=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  checker=new ActiveXObject("Microsoft.XMLHTTP");
					  }
						checker.onreadystatechange=function()
						  {
						  if (checker.readyState==4 && checker.status==200)
							{
								var own=document.getElementById("OWN").value; 
								var vehicle = document.getElementById("Vehicle");
								var chacker = document.getElementById("checker");
								if (own == "01" ){
									if (checker.responseText == "TRUE"){
										chacker.innerHTML="<img src='IMAGES/check.GIF'>";
										$("#Vehicle").removeClass("error_val");
									}else{
										chacker.innerHTML="<img src='IMAGES/error.GIF'>";
										vehicle.value="";
										$("#Vehicle").addClass("error_val");
										alert('" ERROR ! " Kode Vehicle ' + str + ', tidak ada di Vehicle Master');
										$("#Vehicle").focus();
									}
								}
							}
						  }
						checker.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=Check_Vehicle&FcCode="+str,true);
						checker.send();
					}

</script>
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
function validasi(form){

  if (form.CROPCODE.value == ""){

    alert("Anda belum mengisikan Crop Code.");
	
    form.CROPCODE.focus();

    return (false);

  }
if (form.CHITNO.value == ""){

    alert("Anda belum mengisikan Chit No.");

    form.CHITNO.focus();

    return (false);

  }
<%IF Request("txtJenis")<>"Edit" then%> 
	
if (form.FIELDCODE.value == ""){

    alert("Anda belum mengisikan Field Code.");

    form.FIELDCODE.focus();

    return (false);

  } 
<%End If%>  
if (form.ProductCODE.value == ""){

    alert("Anda belum mengisikan Product Code.");

    form.ProductCODE.focus();

    return (false);

  }  
if (form.OWN.value == ""){

    alert("Anda belum mengisikan Own.");

    form.OWN.focus();

    return (false);

  }  
if (form.Vehicle.value == ""){

    alert("Anda belum mengisikan Vehicle.");

    form.Vehicle.focus();

    return (false);
}
if (form.TRANSPORTTYPE.value == ""){

    alert("Anda belum mengisikan Transport Type.");

    form.TRANSPORTTYPE.focus();

    return (false);

  }    
 }
</script>

</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%
	Dim oSQL, oRS
 Dim FCCODE
 Dim FCENTRY, FCEDIT, FCIP, FCBA
 Dim CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, RECEPTIONDATE, VEHICLE, OWN, TRANSPORTTYPE, CONTRACTNUMBER, QUANTITY, ESTATE_WEIGHT, MILL_WEIGHT
 Dim cmdView
	
	Dim val_VehicleStatus : val_VehicleStatus = ""
	Dim val_AvailableVehicle : val_AvailableVehicle = 0
	Dim val_AvailableContract : val_AvailableContract = 0
	Dim val_ContractorCode : val_ContractorCode = ""
	Dim val_LineNo : val_LineNo = ""
	Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

RECEPTIONDATE =request("header_date")
CHITNO=request("CHITNO")
FIELDCODE=request("FIELDCODE") 
CROPCODE=request("CROPCODE")
PRODUCTCODE=request("ProductCODE")
RECEPTIONDATE=request("RECEPTIONDATE")
VEHICLE=request("VEHICLE")
OWN=request("OWN")
TRANSPORTTYPE = Request("TRANSPORTTYPE")
CONTRACTNUMBER = Request("CONTRACTNUMBER")
QUANTITY=request("QUANTITY")
ESTATE_WEIGHT=request("ESTATE_WEIGHT")
MILL_WEIGHT=request("MILL_WEIGHT")

Sub InsertDataHarvesting
	q	=	" Insert into Harvesting (" & _
			" CHITNO" & _
			" , FIELDCODE" & _
			" , CROPCODE" & _
			" , PRODUCTCODE" & _
			" , RECEPTIONDATE" & _
			" , VEHICLE" & _
			" , OWN" & _
			" , QUANTITY" & _
			" , ESTATE_WEIGHT" & _
			" , MILL_WEIGHT" & _
			" , FCENTRY" & _
			" , FCEDIT" & _
			" , FCIP" & _
			" , FCBA" & _
			" , LASTUPDATE" & _
			" , LASTTIME" & _
			" , HARVESTINGTRANPORTTYPECODE" & _
			" , AGREEMENTCODE" & _
			" , CONTRACTORCODE" & _
			" , LINENO" & _
			" ) values (" & _
			" '" & CHITNO & "'" & _
			" , '" & FIELDCODE & "'" & _
			" , '" & CROPCODE & "'" & _
			" , '" & PRODUCTCODE & "'" & _
			" , to_date('" & RECEPTIONDATE & "','mm/dd/yyyy')" & _
			" , '" & VEHICLE & "'" & _
			" , '" & OWN & "'" & _
			" , '" & To_Nol(QUANTITY) & "'" & _
			" , '" & To_Nol(ESTATE_WEIGHT) & "'" & _
			" , '" & To_Nol(MILL_WEIGHT) & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
			" , '" & Session("fcBA") & "'" & _
			" , sysdate " & _
			" , To_Char(sysdate, 'HH24:MI')" & _
			" , '" & TRANSPORTTYPE & "'" & _
			" , '" & CONTRACTNUMBER & "'" & _
			" , '" & val_ContractorCode & "'" & _
			" , '" & val_LineNo & "'" & _
			" )"
			
	'Response.Write q
	connStr.execute(q)
End Sub

Sub UpdateDataHarvesting
	q	=	" Update Harvesting set" & _
			" CROPCODE = '" & CROPCODE & "'" & _
			" , PRODUCTCODE = '" & PRODUCTCODE & "'" & _
			" , VEHICLE = '" & VEHICLE & "'" & _
			" , OWN = '" & OWN & "'" & _
			" , QUANTITY = '" & To_Nol(QUANTITY) & "'" & _
			" , ESTATE_WEIGHT = '" & To_Nol(ESTATE_WEIGHT) & "'" & _
			" , MILL_WEIGHT = '" & To_Nol(MILL_WEIGHT) & "'" & _
			" , LASTUPDATE = sysdate " & _
			" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			" , HARVESTINGTRANPORTTYPECODE = '" & TRANSPORTTYPE & "'" & _
			" , AGREEMENTCODE = '" & CONTRACTNUMBER & "'" & _
			" , CONTRACTORCODE = '" & val_ContractorCode & "'" & _
			" , LINENO = '" & val_LineNo & "'" & _
			" where" & _
			" CHITNO = '" & CHITNO & "'" & _
			" And FIELDCODE = '" & FIELDCODE & "'" & _
			" And RECEPTIONDATE = to_date('" & RECEPTIONDATE & "','mm/dd/yyyy')" & _
			" and fcBA = '" & Session("fcBA") & "'"
	
	'Response.Write q
	connStr.execute(q)
End Sub

Sub DeleteDataHarvesting
	q	=	" Delete from Harvesting where" & _
			" CHITNO = '" & CHITNO & "'" & _
			" And FIELDCODE = '" & FIELDCODE & "'" & _
			" And RECEPTIONDATE = to_date('" & RECEPTIONDATE & "','mm/dd/yyyy')" & _
			" and fcBA = '" & Session("fcBA") & "'"

	'Response.Write q
	connStr.execute(q)
End Sub
 
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), RECEPTIONDATE)
		If Cek_IsDailyClosing = "FALSE" then
			Call DeleteDataHarvesting
		else
			err.number = 1001
			err.description = "Tanggal """ & RECEPTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		oSQL = " SELECT UPPER(FCNAME) VEHICLESTATUSNAME FROM HARVESTINGVEHICLESTATUS WHERE FCCODE = '" & OWN & "'"
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open oSQL, connstr
		
		If Not oRS.EOF Then
			val_VehicleStatus = oRS("VEHICLESTATUSNAME")
		Else
			val_VehicleStatus = "OTHERS"
		End If
		
		oSQL = ""
		oRS.Close
		set oRS = Nothing
	
		If val_VehicleStatus = "OWN" Then
			val_AvailableVehicle = CheckAvailableVehicle(VEHICLE, Session("FCBA"))
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), RECEPTIONDATE)
			If Len(Trim(VEHICLE)) = 0 Then
				err.number = 1001
				err.description = "Kendaraan tidak boleh kosong, isikan dengan Kode Kendaraan"
			Else
				If val_AvailableVehicle = 1 Then
					If Cek_IsDailyClosing = "FALSE" then
						Call InsertDataHarvesting
					else
						err.number = 1001
						err.description = "Tanggal """ & RECEPTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
				ElseIf val_AvailableVehicle < 1 Then
					err.number = 1001
					err.description = "Kendaraan """ & VEHICLE & """, tidak ada didalam tabel master, koreksi kembali kendaraan yang di input"
				Else
					err.number = 1001
					err.description = "Kendaraan """ & VEHICLE & """, ada 2 atau lebih. Cek kembali Master Kendaraan anda."	
				End If
			End If
		ElseIf val_VehicleStatus = "CONTRACT" Then
			val_AvailableContract = CheckAvailableContract(CONTRACTNUMBER, RECEPTIONDATE, Session("FCBA"))
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), RECEPTIONDATE)
			If Len(Trim(CONTRACTNUMBER)) = 0 Then
				err.number = 1001
				err.description = "No.Kontrak tidak boleh kosong, isikan dengan No.Kontrak"
			Else
				If val_AvailableContract = 1 Then
					oSQL =	" SELECT" & _
							" AGREEMENTDETAIL_API.GET_CONTRACTORCODE('" & Session("FCBA") & "', '" & CONTRACTNUMBER & "') CONTRACTORCODE" & _
							" , AGREEMENTDETAIL_API.GET_LINENO('" & Session("FCBA") & "', '" & CONTRACTNUMBER & "', '" & FIELDCODE & "') LINENO" & _
							" FROM DUAL"
		
					Set oRS = Server.CreateObject("ADODB.Recordset")
					oRS.Open oSQL, connstr
					
					If Not oRS.EOF Then
						val_ContractorCode = oRS("CONTRACTORCODE")
						val_LineNo = Cdbl(oRS("LINENO"))
					Else
						val_ContractorCode = "N/A"
						val_LineNo = ""
					End If
					
					oSQL = ""
					oRS.Close
					set oRS = Nothing
					If Cek_IsDailyClosing = "FALSE" then
						Call InsertDataHarvesting
					else
						err.number = 1001
						err.description = "Tanggal """ & RECEPTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
				ElseIf val_AvailableContract < 1 Then
					err.number = 1001
					err.description = "No.Kontrak """ & CONTRACTNUMBER & """, tidak ada didalam Kontrak Pengangkutan atau sudah jatuh tempo, koreksi kembali No.Kontrak yang di input"
				Else
					err.number = 1001
					err.description = "No.Kontrak """ & CONTRACTNUMBER & """, da 2 atau lebih. Cek kembali Kontrak anda."
				End If
			End If
		Else
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), RECEPTIONDATE)
			If Cek_IsDailyClosing = "FALSE" then
				Call InsertDataHarvesting
			else
				err.number = 1001
				err.description = "Tanggal """ & RECEPTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		oSQL = " SELECT UPPER(FCNAME) VEHICLESTATUSNAME FROM HARVESTINGVEHICLESTATUS WHERE FCCODE = '" & OWN & "'"
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open oSQL, connstr
		
		If Not oRS.EOF Then
			val_VehicleStatus = oRS("VEHICLESTATUSNAME")
		Else
			val_VehicleStatus = "OTHERS"
		End If
		
		oSQL = ""
		oRS.Close
		set oRS = Nothing
	
		If val_VehicleStatus = "OWN" Then
			val_AvailableVehicle = CheckAvailableVehicle(VEHICLE, Session("FCBA"))
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), RECEPTIONDATE)
			If Len(Trim(VEHICLE)) = 0 Then
				err.number = 1001
				err.description = "Kendaraan tidak boleh kosong, isikan dengan Kode Kendaraan"
			Else
				If val_AvailableVehicle = 1 Then
					If Cek_IsDailyClosing = "FALSE" then
						Call UpdateDataHarvesting
					else
						err.number = 1001
						err.description = "Tanggal """ & RECEPTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
				ElseIf val_AvailableVehicle < 1 Then
					err.number = 1001
					err.description = "Kendaraan """ & VEHICLE & """, tidak ada didalam tabel master, koreksi kembali kendaraan yang di input"
				Else
					err.number = 1001
					err.description = "Kendaraan """ & VEHICLE & """, ada 2 atau lebih. Cek kembali Master Kendaraan anda."	
				End If
			End If
		ElseIf val_VehicleStatus = "CONTRACT" Then
			val_AvailableContract = CheckAvailableContract(CONTRACTNUMBER, RECEPTIONDATE, Session("FCBA"))
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), RECEPTIONDATE)
			If Len(Trim(CONTRACTNUMBER)) = 0 Then
				err.number = 1001
				err.description = "No.Kontrak tidak boleh kosong, isikan dengan No.Kontrak"
			Else
				If val_AvailableContract = 1 Then
					oSQL =	" SELECT" & _
							" AGREEMENTDETAIL_API.GET_CONTRACTORCODE('" & Session("FCBA") & "', '" & CONTRACTNUMBER & "') CONTRACTORCODE" & _
							" , AGREEMENTDETAIL_API.GET_LINENO('" & Session("FCBA") & "', '" & CONTRACTNUMBER & "', '" & FIELDCODE & "') LINENO" & _
							" FROM DUAL"
		
					Set oRS = Server.CreateObject("ADODB.Recordset")
					oRS.Open oSQL, connstr
					
					If Not oRS.EOF Then
						val_ContractorCode = oRS("CONTRACTORCODE")
						val_LineNo = Cdbl(oRS("LINENO"))
					Else
						val_ContractorCode = "N/A"
						val_LineNo = ""
					End If
					
					oSQL = ""
					oRS.Close
					set oRS = Nothing
					If Cek_IsDailyClosing = "FALSE" then
						Call UpdateDataHarvesting
					else
						err.number = 1001
						err.description = "Tanggal """ & RECEPTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
				ElseIf val_AvailableContract < 1 Then
					err.number = 1001
					err.description = "No.Kontrak """ & CONTRACTNUMBER & """, tidak ada didalam Kontrak Pengangkutan atau sudah jatuh tempo, koreksi kembali No.Kontrak yang di input"
				Else
					err.number = 1001
					err.description = "No.Kontrak """ & CONTRACTNUMBER & """, da 2 atau lebih. Cek kembali Kontrak anda."
				End If
			End If
		Else
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), RECEPTIONDATE)
			If Cek_IsDailyClosing = "FALSE" then
				Call UpdateDataHarvesting
			else
				err.number = 1001
				err.description = "Tanggal """ & RECEPTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
		End If
		cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if RECEPTIONDATE <> "" then
   RECEPTIONDATE   =Request("RECEPTIONDATE")
else
   RECEPTIONDATE   =textDate(date)
end if
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

If Err.Number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if
%>
 <!-------------ISI Payment Form!------------->
 <form name="frmEntry" method="post" action="Harvesting.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<tr>
			<td colspan="2">
				<div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE HARVESTING</strong></font></div>
			</td>
		</tr>

		<tr>
			<td align="left" width="30%">Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE" size="10" value="<%=RECEPTIONDATE%>">
				
			</td>
		</tr>
			
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;<a href="PrintHarvesting.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&RECEPTIONDATE=<%=RECEPTIONDATE%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Harvesting"
				dataSubmodule 	= "Harvesting Receptions"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" And len(trim(RECEPTIONDATE)) > 0 Then
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q	=	" select" & _
			" CHITNO, FIELDCODE, CROPCODE, ProductCODE, to_char(RECEPTIONDATE,'mm/dd/yyyy') as RECEPTIONDATE" & _
			" , VEHICLE, OWN, HARVESTINGTRANPORTTYPECODE, AGREEMENTCODE, CONTRACTORCODE" & _
			" , LINENO, QUANTITY, ESTATE_WEIGHT, MILL_WEIGHT" & _
			" , FCENTRY, FCEDIT, FCIP, FCBA from Harvesting where fcBA='" & session("fcBA") & "'" & _
			" and RECEPTIONDATE = to_date('" & RECEPTIONDATE & "', 'MM/DD/YYYY')" & _
			" ORDER BY CHITNO, FIELDCODE"
	
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	'response.Write(q)
	'response.Flush()
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
   
%>
   <form action="Harvesting.Asp" method="post" name="eForm" onSubmit="javascript:return validasi(this)">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		
		<tr>
			<td align="left" width="30%">Chit No</td>
			<td align="left">
				<input type="text" size="15" name="CHITNO" value="<%=CHITNO%>" <%IF txtJenis="Edit" Then%> readonly="True" style="background:#999999"<%End if%>><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Field Code</td>
			<td align="left">
				<select name="FIELDCODE" <%IF Request("txtJenis")="Edit" then%> disabled = "disabled" <%End If%>>
					<option value="" disabled selected></option> 
				<%
					Dim rsFieldDeatail, strFieldDeatail
					strFieldDeatail = " select fccode, field_api.getValid_Name(fccode, fcba, to_date('" & RECEPTIONDATE & "', 'mm/dd/yyyy')) fcname" & _
									  " from field" & _
									  " where fcba = '" & Session("fcBA") & "'" & _
									  "       and field_api.getvalid_status(fccode, fcba, to_date('" & RECEPTIONDATE & "', 'mm/dd/yyyy')) = 'TM'" & _
									  "       and nvl(field_api.getvalid_activation(fccode, fcba, to_date('" & RECEPTIONDATE & "', 'mm/dd/yyyy')), 'N') = 'Y'" & _
									  " order by fccode"

					
					Set rsFieldDeatail = Server.CreateObject("ADODB.Recordset")
					rsFieldDeatail.Open strFieldDeatail, connstr
					
					If Not rsFieldDeatail.EOF then
						Do While Not rsFieldDeatail.EOF
				%>
					<option value="<%=rsFieldDeatail("FCCODE")%>" <%If FIELDCODE = rsFieldDeatail("FCCODE") Then%> selected <%End If%>><%=rsFieldDeatail("FCCODE")%> : <%=rsFieldDeatail("FCNAME")%></option> 
				<%			rsFieldDeatail.movenext
						Loop
					End if
					rsFieldDeatail.Close
					set rsFieldDeatail = Nothing
				%>
				</select><font color="#FF0000"> *</font>
				<%IF Request("txtJenis")="Edit" then%> 
				<input type="hidden" name="FIELDCODE" value="<%=FIELDCODE%>" />
				<%End If%>
			</td>
		</tr>
		
		<tr>
			<td align="left">Crop Code</td>
			<td align="left">
				<select name="CROPCODE">
				<%
					Dim rsCrop, strCrop
					strCrop =	" SELECT P.FCCODE,P.FCNAME FROM  PARAMETER P  WHERE P.PARCODE= 'F0005' AND FCBA = '" &  Session("fcBA") & "'" & _
								" AND P.FCCODE = 'OP' ORDER BY P.FCNAME DESC"
					
					Set rsCrop = Server.CreateObject("ADODB.Recordset")
					rsCrop.Open strCrop, connstr
					
					If Not rsCrop.EOF then
						Do While Not rsCrop.EOF
				%>
					<option value="<%=rsCrop("FCCODE")%>" <%If CROPCODE = rsCrop("FCCODE") Then%> selected <%End If%>><%=rsCrop("FCNAME")%></option> 
				<%			rsCrop.movenext
						Loop
					End if
					rsCrop.Close
					set rsCrop = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
			<td align="left"><font face="Verdana" size="1">Product Code</font></td>
			<td align="left">
				<select name="ProductCODE">
				<%
					Dim rsProductCode, strProductCode
					strProductCode = " SELECT PRODUCTCODE, PRODUCTDESCRIPTION FROM PRODUCT WHERE FCBA = '" &  Session("fcBA") & "' AND PRODUCTCODE = 'FFB' ORDER BY PRODUCTCODE"
					
					Set rsProductCode = Server.CreateObject("ADODB.Recordset")
					rsProductCode.Open strProductCode, connstr
					
					If Not rsProductCode.EOF then
						Do While Not rsProductCode.EOF
				%>
					<option value="<%=rsProductCode("PRODUCTCODE")%>" <%If PRODUCTCODE = rsProductCode("PRODUCTCODE") Then%> selected <%End If%>><%=rsProductCode("PRODUCTCODE")%> : <%=rsProductCode("PRODUCTDESCRIPTION")%></option> 
				<%			rsProductCode.movenext
						Loop
					End if
					rsProductCode.Close
					set rsProductCode = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>

		<tr>
			<td align="left"><font face="Verdana" size="1">Vehicle Status</font></td>
			<td align="left">
				<select id="OWN" name="OWN" onchange="javascript:combo_(this)">
				<%
					Dim rsVehicleStatus, strVehicleStatus
					strVehicleStatus = " SELECT FCCODE, FCNAME FROM HARVESTINGVEHICLESTATUS ORDER BY FCCODE"
					
					Set rsVehicleStatus = Server.CreateObject("ADODB.Recordset")
					rsVehicleStatus.Open strVehicleStatus, connstr
					
					If Not rsVehicleStatus.EOF then
						Do While Not rsVehicleStatus.EOF
				%>
					<option value = "<%=rsVehicleStatus("FCCODE")%>" <%If OWN = rsVehicleStatus("FCCODE") Then%> selected <%End If%>>
						<%=rsVehicleStatus("FCNAME")%>
					</option> 
				<%			rsVehicleStatus.movenext
						Loop
					End if
					rsVehicleStatus.Close
					set rsVehicleStatus = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
            <td align="left">Vehicle</td>
			<td align="left">
				<input id="Vehicle" onblur="showHint(this.value)" type="text" size="15" maxlength="50" name="Vehicle" value="<%=Vehicle%>">
				<span id="checker" style="color:red;margin-top:10px;"> *</span>
			</td>
		</tr>
		
		<tr>
			<td align="left"><font face="Verdana" size="1">Transport Type</font></td>
			<td align="left">
				<select name="TRANSPORTTYPE">
				<%
					Dim rsTransportType, strTransportType
					strTransportType = " SELECT FCCODE, FCNAME FROM HARVESTINGTRANSPORTTYPE ORDER BY FCCODE"
					
					Set rsTransportType = Server.CreateObject("ADODB.Recordset")
					rsTransportType.Open strTransportType, connstr
					
					If Not rsTransportType.EOF then
						Do While Not rsTransportType.EOF
				%>
					<option value = "<%=rsTransportType("FCCODE")%>" <%If TRANSPORTTYPE = rsTransportType("FCCODE") Then%> selected <%End If%>>
						<%=rsTransportType("FCNAME")%>
					</option> 
				<%			rsTransportType.movenext
						Loop
					End if
					rsTransportType.Close
					set rsTransportType = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
		</tr>
		
		<tr>
            <td align="left">Contract Number</td>
			<td align="left">
			<input type="text" size="30" maxlength="30" name="CONTRACTNUMBER" value="<%=CONTRACTNUMBER%>">
				<span id="checker_contract"></span>
			</td>
		</tr>

		<tr><td align="left">Bunch</td><td align="left"><input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>"></td></tr>
		<tr><td align="left">Estate Weight (Kg)</td><td align="left"><input type="text" size="15" name="ESTATE_WEIGHT" value="<%=ESTATE_WEIGHT%>"></td></tr>
		<tr><td align="left">Mill Weight (Kg)</td><td align="left"><input type="text" size="15" name="MILL_WEIGHT" value="<%=MILL_WEIGHT%>"></td></tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="RECEPTIONDATE" value="<%=RECEPTIONDATE%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
 </form>
  <!--#include file="validation_show.Asp"--> 
   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CHIT NO</TD>	
	<TD align="left">FILED CODE</TD>	
	<TD align="left">CROP CODE</TD>	
	<TD align="left">PRODUCT CODE</TD>	
	<TD align="left">RECEPTION DATE</TD>	
	<TD align="left">VEHICLE</TD>	
	<TD align="left">OWN</TD>
	<td align="left">TRANSPORT TYPE</td>
	<td align="left">AGREEMENT CODE</td>
	<td align="left">CONTRACTOR CODE</td>
	<td align="left">LINE NO</td>
	<TD align="left">BUNCH</TD>	
	<TD align="left">ESTATE WEIGHT</TD>	
	<TD align="left">MILL WEIGHT</TD>	
	</tr>
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="Harvesting.Asp?CHITNO=<%=rsVH("CHITNO")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&CROPCODE=<%=rsVH("CROPCODE")%>&PRODUCTCODE=<%=rsVH("PRODUCTCODE")%>&RECEPTIONDATE=<%=rsVH("RECEPTIONDATE")%>&VEHICLE=<%=rsVH("VEHICLE")%>&OWN=<%=rsVH("OWN")%>&TRANSPORTTYPE=<%=rsVH("HARVESTINGTRANPORTTYPECODE")%>&CONTRACTNUMBER=<%=rsVH("AGREEMENTCODE")%>&QUANTITY=<%=rsVH("QUANTITY")%>&ESTATE_WEIGHT=<%=rsVH("ESTATE_WEIGHT")%>&MILL_WEIGHT=<%=rsVH("MILL_WEIGHT")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="30">
			<a href="Harvesting.Asp?CHITNO=<%=rsVH("CHITNO")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&RECEPTIONDATE=<%=rsVH("RECEPTIONDATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></td>
		    <%for i = 0 to nCountofField-4%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </TABLE>
  </div>

<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
