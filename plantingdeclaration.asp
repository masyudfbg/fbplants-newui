<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Planting Declaration</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script>
		function hanyaAngka(evt) {
		  var charCode = (evt.which) ? evt.which : event.keyCode
		   if (charCode > 31 && (charCode < 48 || charCode > 57))
 
		    return false;
		  return true;
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

Dim qSupplyingDeclaration, rsSupplyingDeclarationInfo, tempQuantity, tempDecType
Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, v_FCCODE, v_FCNAME
Dim qSupplyingInfo, rsSupplyingInfo, qDeclarationType, rsDeclarationType
Dim hectare, parcode, totalstand, fullstand
Dim FIELDCODE 
Dim QUANTITY
Dim REMARKS
Dim TRANSDATE
Dim cmdView
Dim txtDocumentNo ,txtlineno
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

hectare =request("hectare")
TRANSDATE =request("TRANSDATE")
quantity=request("quantity")
remarks=request("remarks")
fieldcode=request("fieldcode")
txtDocumentNo = Request("txtDocumentNo")
fullstand = Request("fullstand")
txtlineno = request("txtlineno")


Function CheckFieldCode(initialFieldCode, initialFcba, initialDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from dual " & _
					" where field_api.GetValid_Activation('" & initialFieldCode & "', '" & initialFcba & "', to_date('"& initialDate &"', 'mm/dd/yyyy')) = 'Y'"
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckFieldCode = initialAvailable
End Function

Function CheckPlantingDate(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select count (1) total" & _
	            " from dual " & _
	            " where to_char(field_api.GetValid_Plantingdate('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')),'yyyy')= to_char(to_date('"& initDate &"', 'mm/dd/yyyy'),'yyyy')"
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	CheckPlantingDate = initialAvail
End Function

Function F_CurrentTotalStand(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select field_api.getValid_POKOK('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')) total" & _
	            " from dual " 
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	F_CurrentTotalStand = initialAvail
End Function

Function F_CurrentHectaragePlanted(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select field_api.getvalid_Ha('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')) total" & _
	            " from dual " 
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	F_CurrentHectaragePlanted = initialAvail
End Function

Function PokokDataTrans(initfieldcode, initialfcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" SELECT sum(pdd.quantity) total " & _
		  " FROM plantingdeclarationdtl pdd" & _
		  " INNER JOIN plantingdeclarationhdr pdh"& _
		  " ON (pdd.documentno = pdh.documentno AND pdd.fcba = pdh.fcba)" & _
		  " WHERE pdd.fcba = '" & initialfcba & "'" & _
		  " AND pdd.fieldcode = '" & initfieldcode & "'" 
	 ' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(to_nol(rsAvailable("Total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	PokokDataTrans = initialAvailable
End Function

Function HectareDataTrans(initfieldcode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	  " SELECT sum(pdd.hectare) total " & _
					  " FROM plantingdeclarationdtl pdd" & _
					  " INNER JOIN plantingdeclarationhdr pdh"& _
					  " ON (pdd.documentno = pdh.documentno AND pdd.fcba = pdh.fcba)" & _
					  " WHERE pdd.fcba = '" & initialfcba & "'" & _
					  " AND pdd.fieldcode = '" & initfieldcode & "'" 		  
	 ' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(to_nol(rsAvailable("Total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	HectareDataTrans = initialAvailable
End Function

Function CheckFullStand(initfieldcode, initialFcba, initDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	 " select field_api.getvalid_fullstand('" & initFieldCode & "', '" & initialFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')) total" & _
	           		 " from dual " 

	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("Total")
	Else
		initialAvailable = "FALSE"
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckFullStand = initialAvailable
End Function

if ""&fullstand = "" then fullstand = "FALSE"

If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q = " delete from plantingdeclarationhdr " & _
				" 		where documentno = '" & txtDocumentNo & "'" & _
				"       and fcBA = '" & Session("fcBA") & "'"
							' Response.Write q
							connStr.execute(q)
			
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			Current_Fullstand = CheckFullStand(fieldcode, session("fcBA"), TRANSDATE)
			' Post_fullstand = Current_Fullstand
			' response.write Current_Fullstand
			If Current_Fullstand = "TRUE" Then
				Err.Number = 1001
				Err.Description = "Fieldcode """& fieldcode &""" sudah Full Stand tidak dapat di hapus."
			Else
			
			q = " delete from plantingdeclarationdtl " & _
						" 		where lineno = '" & txtlineno & "'" & _
						"       and fieldcode = '" & fieldcode & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
				' Response.Write q
				connStr.execute(q)
				cmdView	= "True"
				' Current_Hectare = F_CurrentHectaragePlanted(fieldcode, session("fcBA"), TRANSDATE)
				' Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
				Current_Hectare = F_CurrentHectaragePlanted(fieldcode, session("fcBA"), TRANSDATE)
				strHectare =	" select field_api.getValid_ha('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& transdate &"', 'mm/dd/yyyy')) total" & _
	 	                		" from dual "								
									' Response.Write strHectare

									Set rsQty = Server.CreateObject("ADODB.Recordset")
									rsQty.Open strHectare, connstr							
				if not rsQty.Eof then
				do while not rsQty.eof
						maxhectare = cdbl (to_nol(rsQty("total")))
						rsQty.MoveNext
					loop
				End If
				' response.write hectare&"<br>"
				' response.write maxhectare&"<br>"
				' response.write val_QtyHectare&"<br>"

				' Current_Hectare = F_CurrentHectaragePlanted(fieldcode, session("fcBA"), TRANSDATE)
				Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
				strFieldDetail =	" select field_api.getValid_pokok('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& transdate &"', 'mm/dd/yyyy')) total" & _
	 	                			" from dual "								
									' Response.Write strFieldDetail

									Set rsQty = Server.CreateObject("ADODB.Recordset")
									rsQty.Open strFieldDetail, connstr							
				if not rsQty.Eof then
				do while not rsQty.eof
						maxquantity = cdbl (to_nol(rsQty("total")))
						rsQty.MoveNext
					loop
				End If
				
				Post_hectarageplanted = maxhectare - hectare
				post_totalstand = maxquantity - quantity
					a =	" Update field" & _
						" set hectarageplanted = '" & Post_hectarageplanted  & "'" & _
						"      ,totalstand = '" & post_totalstand  & "'" & _
						"      ,fullstand = 'FALSE'" & _
						" 		where fccode = '" & fieldcode & "'" & _
						"		and fcba = '" & Session("fcba") & "'"
				 	response.write  maxhectare
				 	response.write maxquantity
					connStr.execute(a)


					strFieldDetail =	"select fccode" & _
										"        , field_api.getvalid_name(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FCNAME " & _
										"        , field_api.getvalid_afdeling(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) DIVISION " & _
										"        , field_api.getvalid_Ha(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HECTARAGEPLANTED" & _
										"        , field_api.getvalid_terraintype(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) TERRAINTYPE" & _
										"        , field_api.getvalid_soiltype(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) SOILTYPE" & _
										"        , field_api.getvalid_ownership(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) OWNERSHIP" & _
										"        , field_api.getvalid_activation(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) ACTIVATION" & _
										"        , field_api.getvalid_croptype(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) CROPTYPE" & _
										"        , field_api.getvalid_plantingdate(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PLANTINGDATE" & _
										"        , field_api.getvalid_plantingmaterial(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PLANTINGMATERIAL" & _
										"        , field_api.getvalid_pokok(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) TOTALSTAND" & _
										"        , field_api.getvalid_status(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) STATUS" & _
										"        , field_api.getvalid_prevsemester_abw(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PREVSEMESTER_ABW" & _
										"        , field_api.getvalid_pressemester_abw(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PRESSEMESTER_ABW" & _
										"        , field_api.getvalid_harvestingbased_abw(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HARVESTINGBASED_ABW" & _
										"        , field_api.getvalid_main_road(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) MAIN_ROAD" & _
										"        , field_api.getvalid_collectionroad(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) COLLECTIONROAD" & _
										"        , field_api.getvalid_maindrain(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))MAINDRAIN" & _
										"        , field_api.getvalid_collectiondrain(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) COLLECTIONDRAIN" & _
										"        , field_api.getvalid_subsidiarydrain(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) SUBSIDIARYDRAIN" & _
										"        , field_api.getvalid_collectionpoint(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) COLLECTIONPOINT" & _
										"        , field_api.getvalid_harvestingbridge(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) HARVESTINGBRIDGE" & _
										"        , field_api.getvalid_permanentbridge(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PERMANENTBRIDGE" & _
										"        , field_api.getvalid_woodbridge(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) WOODBRIDGE" & _
										"        , field_api.getvalid_buist(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) BUIST" & _
										"        , field_api.getvalid_owlnest(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) OWLNEST" & _
										" 		 , '" & Session("CodeName") & "' FCENTRY" & _
								       	" 	 	 , '" & Session("CodeName") & "'FCEDIT" & _
								       	" 		 , '" & Request.ServerVariables("REMOTE_ADDR") & "' FCIP" & _
								       	"		 , '" & Session("fcba") & "' FCBA " & _
										"        , sysdate LASTUPDATE" & _
										"        , to_char (sysdate, 'hh24:mi')LASTTIME" & _
										"        , field_api.getvalid_harvestingsection(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) HARVESTINGSECTION" & _
										"        , field_api.getvalid_field_info_01(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_01" & _
										"        , field_api.getvalid_field_info_02(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_02" & _
										"        , field_api.getvalid_field_info_03(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_03" & _
										"        , field_api.getvalid_field_info_04(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_04" & _
										"        , field_api.getvalid_field_info_05(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_05" & _
										"        , field_api.getvalid_isactive_harvested(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) ISACTIVE_HARVESTED" & _
										"        , trunc(to_date('"& transdate &"', 'mm/dd/yyyy'),'mm') VALID_FROM " & _
										"        , (select nvl(max(lineno), 0) from fielddetail fd where fd.fccode = ff.fccode and fd.fcba = ff.fcba) +1 LINENO , 'Updated' ACTION_STATE " & _
										"        , field_api.getvalid_AREAL_LC(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))AREAL_LC" & _
										"        , field_api.getvalid_CADANGAN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))CADANGAN" & _
										"        , field_api.getvalid_OKUPASI(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))OKUPASI" & _
										"        , field_api.getvalid_BIBITAN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))BIBITAN" & _
										"        , field_api.getvalid_PABRIK(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PABRIK" & _
										"        , field_api.getvalid_EMPLASMEN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))EMPLASMEN" & _
										"        , field_api.getvalid_PARIT(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PARIT" & _
										"        , field_api.getvalid_JALAN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))JALAN" & _
										"        , field_api.getvalid_PRASARANA_LAINNYA(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PRASARANA_LAINNYA" & _
										"        , field_api.getvalid_QUARRY(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))QUARRY" & _
										"        , field_api.getvalid_TANGGUL(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))TANGGUL" & _
										"        , field_api.getvalid_ENCLAVE(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))ENCLAVE" & _
										"        , field_api.getvalid_TAMBANG(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))TAMBANG" & _
										"        , field_api.getvalid_KONSERVASI(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))KONSERVASI" & _
										"        , field_api.getvalid_BUKIT_SUNGAI(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))BUKIT_SUNGAI" & _
										"        , field_api.getvalid_NO_IMAGERY(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))NO_IMAGERY" & _
										"        , field_api.getvalid_LUAS_GROSS(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))LUAS_GROSS" & _
										"        , field_api.getvalid_HECTARAGEHARVESTED(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HECTARAGEHARVESTED" & _
										"        , field_api.getvalid_PRODUCTIVESTAND(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PRODUCTIVESTAND" & _
										"        , field_api.getvalid_HARVESTINGSECTIONNUMB(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HARVESTINGSECTIONNUMBER" & _
										"        , field_api.getvalid_PLANTINGPOINT(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PLANTINGPOINT" & _
										"        , field_api.getvalid_FULLSTAND(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))FULLSTAND" & _
										" 		 from field ff" & _
										" 		 where fccode = '"& fieldcode &"'" & _
										"        and fcba = '" & Session("fcba") & "'"								
										'Response.Write q
										Set rsAvail = Server.CreateObject("ADODB.Recordset")
										rsAvail.Open strFieldDetail, connstr

							if not rsAvail.Eof then
								do while not rsAvail.eof
									q = "insert into fielddetail (FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE, BUIST, OWLNEST, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, HARVESTINGSECTION, FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05, ISACTIVE_HARVESTED, VALID_FROM, LINENO, ACTION_STATE, AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA, QUARRY, TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER, PLANTINGPOINT, FULLSTAND) Values (" & _
									"  '" & rsAvail("FCCODE") & "'" & _
									" , '" & rsAvail("FCNAME") & "'" & _
									" , '" & rsAvail("DIVISION") & "'" & _
									" , '" & Post_hectarageplanted  & "'" & _
									" , '" & rsAvail("TERRAINTYPE") & "'" & _
									" , '" & rsAvail("SOILTYPE") & "'" & _
									" , '" & rsAvail("OWNERSHIP") & "'" & _
									" , '" & rsAvail("ACTIVATION") & "'" & _
									" , '" & rsAvail("CROPTYPE") & "'" & _
									" , to_date('" & rsAvail("PLANTINGDATE") & "','mm/dd/yyyy')" & _
									" , '" & rsAvail("PLANTINGMATERIAL") & "'" & _
									" , '" & post_totalstand & "'" & _
									" , '" & rsAvail("STATUS") & "'" & _
									" , '" & rsAvail("PREVSEMESTER_ABW") & "'" & _
									" , '" & rsAvail("PRESSEMESTER_ABW") & "'" & _
									" , '" & rsAvail("HARVESTINGBASED_ABW") & "'" & _
									" , '" & rsAvail("MAIN_ROAD") & "'" & _
									" , '" & rsAvail("COLLECTIONROAD") & "'" & _
									" , '" & rsAvail("MAINDRAIN") & "'" & _
									" , '" & rsAvail("COLLECTIONDRAIN") & "'" & _
									" , '" & rsAvail("SUBSIDIARYDRAIN") & "'" & _
									" , '" & rsAvail("COLLECTIONPOINT") & "'" & _
									" , '" & rsAvail("HARVESTINGBRIDGE") & "'" & _
									" , '" & rsAvail("PERMANENTBRIDGE") & "'" & _
									" , '" & rsAvail("WOODBRIDGE") & "'" & _
									" , '" & rsAvail("BUIST") & "'" & _
									" , '" & rsAvail("OWLNEST") & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									" , '" & Session("fcba") & "'" & _
									" , sysdate" & _
									" , to_char (sysdate, 'hh24:mi')" & _
									" , '" & rsAvail("HARVESTINGSECTION") & "'" & _
									" , '" & rsAvail("FIELD_INFO_01") & "'" & _
									" , '" & rsAvail("FIELD_INFO_02") & "'" & _
									" , '" & rsAvail("FIELD_INFO_03") & "'" & _
									" , '" & rsAvail("FIELD_INFO_04") & "'" & _
									" , '" & rsAvail("FIELD_INFO_05") & "'" & _
									" , '" & rsAvail("ISACTIVE_HARVESTED") & "'" & _
									" , to_date('" & TRANSDATE & "','mm/dd/yyyy')" & _
									" , '" & rsAvail("LINENO") & "'" & _
									" , '" & rsAvail("ACTION_STATE") & "'" & _
									" , '" & rsAvail("AREAL_LC") & "'" & _
									" , '" & rsAvail("CADANGAN") & "'" & _
									" , '" & rsAvail("OKUPASI") & "'" & _
									" , '" & rsAvail("BIBITAN") & "'" & _
									" , '" & rsAvail("PABRIK") & "'" & _
									" , '" & rsAvail("EMPLASMEN") & "'" & _
									" , '" & rsAvail("PARIT") & "'" & _
									" , '" & rsAvail("JALAN") & "'" & _
									" , '" & rsAvail("PRASARANA_LAINNYA") & "'" & _
									" , '" & rsAvail("QUARRY") & "'" & _
									" , '" & rsAvail("TANGGUL") & "'" & _
									" , '" & rsAvail("ENCLAVE") & "'" & _
									" , '" & rsAvail("TAMBANG") & "'" & _
									" , '" & rsAvail("KONSERVASI") & "'" & _
									" , '" & rsAvail("BUKIT_SUNGAI") & "'" & _
									" , '" & rsAvail("NO_IMAGERY") & "'" & _
									" , '" & rsAvail("LUAS_GROSS") & "'" & _
									" , '" & rsAvail("HECTARAGEHARVESTED") & "'" & _
									" , '" & rsAvail("PRODUCTIVESTAND") & "'" & _
									" , '" & rsAvail("HARVESTINGSECTIONNUMBER") & "'" &_
									" , '" & rsAvail("PLANTINGPOINT") & "'" &_
									" , '" & rsAvail("FULLSTAND") & "'" &_
									")" 
									'Response.Write q
									connStr.execute(q)
									rsAvail.MoveNext
									
								loop
							End If
						End if
						cmdView	= "True"

	elseif Request("txtJenis")="New" then		
			Dim rsPlantingDeclarationHdr
			Set rsPlantingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsPlantingDeclarationHdr.Open " SELECT documentno FROM plantingdeclarationhdr WHERE documentno = '" & txtDocumentNo & "' and fcba ='" & Session("fcba") & "' ", connstr
			If rsPlantingDeclarationHdr.EOF Then
				q = " INSERT INTO plantingdeclarationhdr" & _
					" (documentno, transdate, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) VALUES (" & _
					" '" &txtDocumentNo & "'" & _
					" , to_date('" & TRANSDATE & "','mm/dd/yyyy')" & _ 
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcba") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"								
				'Response.Write q
				connStr.execute(q)
				
			End If
			rsPlantingDeclarationHdr.Close
			set rsPlantingDeclarationHdr = Nothing
			
			If Len(Trim(fieldcode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Fieldcode Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(hectare)) = 0 Then
				Err.Number = 1001
				Err.Description = "Hectare Harus diisi, tidak boleh kosong"
			Else
			If Len(Trim(quantity)) = 0 Then
				Err.Number = 1001
				Err.Description = "Quantity Total Stand Harus diisi, tidak boleh kosong"
			Else
			Dim val_PlantingCode : val_PlantingCode = 0
			val_PlantingCode = CheckFieldCode(fieldcode, session("fcBA"), TRANSDATE)
			If val_PlantingCode = 0 Then
				Err.Number = 1001
				Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
			Else
			
			strFieldDetail =	" select field_api.getvalid_plantingyear('"& fieldcode &"','"& session("fcba") &"',to_date('"& transdate &"', 'MM/DD/YYYY')) total from dual "
								' Response.Write strFieldDetail
								Set rsQty = Server.CreateObject("ADODB.Recordset")
								rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
			do while not rsQty.eof
					plantingyear = cdbl (to_nol(rsQty("total")))
					rsQty.MoveNext
				loop
			End If
			
			  oSQLQuery =" SELECT currentaccountingyear as periodyear FROM periodcontrol" & _
			 			 " WHERE fcba = '" & Session("fcBA") & "'" & _
			       		 " AND UPPER (SYSTEM) = UPPER ('costbook')"

			' response.write oSQLQuery
			set oRecordSet = server.CreateObject("Adodb.recordset")
			oRecordSet.Open oSQLQuery, connstr
			If Not oRecordSet.Eof then
				periodyear = oRecordSet("periodyear")
			End if

			If cstr(""&plantingyear) <> cstr(""&periodyear) Then
				err.number = 1001
				Err.Description = "Data yang anda masukan tidak sesuai dengan Planting Date"
			Else

			Current_Fullstand = CheckFullStand(fieldcode, session("fcBA"), TRANSDATE)
			' Post_fullstand = Current_Fullstand
			' response.write Current_Fullstand
			If Current_Fullstand = "TRUE" Then
				Err.Number = 1001
				Err.Description = "Planting Declaration sudah Full Stand tidak dapat di input kembali."
			Else

			Current_Hectare = F_CurrentHectaragePlanted(fieldcode, session("fcBA"), TRANSDATE)
			strHectare =	" select field_api.getValid_LUAS_GROSS('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& transdate &"', 'mm/dd/yyyy')) total" & _
 	                		" from dual "								
								' Response.Write strHectare

								Set rsQty = Server.CreateObject("ADODB.Recordset")
								rsQty.Open strHectare, connstr							
			if not rsQty.Eof then
			do while not rsQty.eof
					maxhectare = cdbl (to_nol(rsQty("total")))
					rsQty.MoveNext
				loop
			End If
			' response.write hectare&"<br>"
			' response.write maxhectare&"<br>"
			' response.write val_QtyHectare&"<br>"
			If (cdbl (to_nol(hectare)) + Current_Hectare) > cdbl (to_nol(maxhectare)) Then
				Err.Number = 1001
				Err.Description = "Data tidak berhasil disimpan, Total Luas Gross = """ & maxhectare & """, Total Ha Planted = """ & Current_Hectare & """,  mohon cek kembali data yang anda masukan"
			Else

			' Current_Hectare = F_CurrentHectaragePlanted(fieldcode, session("fcBA"), TRANSDATE)
			Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
			strFieldDetail =	" select field_api.getValid_plantingpoint('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& transdate &"', 'mm/dd/yyyy')) total" & _
 	                			" from dual "								
								' Response.Write strFieldDetail

								Set rsQty = Server.CreateObject("ADODB.Recordset")
								rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
			do while not rsQty.eof
					maxquantity = cdbl (to_nol(rsQty("total")))
					rsQty.MoveNext
				loop
			End If
			' response.write maxquantity&"<br>"
			' response.write quantity&"<br>"
			' response.write Current_TotalStand&"<br>"
			' response.write val_QtySupplying&"<br>"
			If (cdbl (to_nol(quantity)) +  Current_TotalStand) > cdbl (to_nol(maxquantity)) Then
				Err.Number = 1001
				Err.Description = "Data tidak berhasil disimpan, Total Planting Point = """ & maxquantity & """, Total Stand = """ & Current_TotalStand & """, mohon cek kembali data yang anda masukan"
			Else

					' Dim val_Plantingdate : val_Plantingdate = 0
					' val_Plantingdate = CheckPlantingDate(fieldcode, session("fcBA"), TRANSDATE)
					' Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
					
					' If cstr(""&hectare) <> cstr("1")  Then
			If QUANTITY <= 0 Then
				err.number = 1001
				Err.Description = "Anda harus Masukkan Angka di atas 0"
			
			Else
				qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
						   " from plantingdeclarationdtl" & _
						   " where documentno = '" & txtDocumentNo & "'" & _
						   " and fcba = '" & session("fcBA") & "'"			
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 

				tempMaxLine = tempMaxLine + 1
				txtLineNoHist = tempMaxLine

				q = " INSERT INTO plantingdeclarationdtl" & _
					" (documentno, fieldcode, quantity, hectare, remarks, fcentry, fcedit, fcip, fcba, lastupdate, lasttime, fullstand, lineno) VALUES (" & _
					" '" & txtdocumentno & "'" & _
					" , '" & fieldcode & "'" & _
					" , '" & quantity & "'" & _
					" , '" & hectare & "'" & _
					" , '" & remarks & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcba") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" , '" & fullstand & "'" & _
					" , '" & tempMaxLine & "'" & _
					" )"
					' Response.Write q
					connStr.execute(q)
				End If


					Post_hectarageplanted = Current_Hectare + hectare
					post_totalstand = Current_TotalStand + quantity
				a =	" Update field" & _
					" set hectarageplanted = '" & Post_hectarageplanted  & "'" & _
					"      ,totalstand = '" & post_totalstand  & "'" & _
					"      ,fullstand = '" & fullstand  & "'" & _
					" 		where fccode = '" & fieldcode & "'" & _
					"		and fcba = '" & Session("fcba") & "'"
			 	' response.write  a
				connStr.execute(a)
						
					strFieldDetail =	"select fccode" & _
										"        , field_api.getvalid_name(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FCNAME " & _
										"        , field_api.getvalid_afdeling(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) DIVISION " & _
										"        , field_api.getvalid_Ha(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HECTARAGEPLANTED" & _
										"        , field_api.getvalid_terraintype(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) TERRAINTYPE" & _
										"        , field_api.getvalid_soiltype(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) SOILTYPE" & _
										"        , field_api.getvalid_ownership(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) OWNERSHIP" & _
										"        , field_api.getvalid_activation(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) ACTIVATION" & _
										"        , field_api.getvalid_croptype(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) CROPTYPE" & _
										"        , field_api.getvalid_plantingdate(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PLANTINGDATE" & _
										"        , field_api.getvalid_plantingmaterial(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PLANTINGMATERIAL" & _
										"        , field_api.getvalid_pokok(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) TOTALSTAND" & _
										"        , field_api.getvalid_status(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) STATUS" & _
										"        , field_api.getvalid_prevsemester_abw(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PREVSEMESTER_ABW" & _
										"        , field_api.getvalid_pressemester_abw(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PRESSEMESTER_ABW" & _
										"        , field_api.getvalid_harvestingbased_abw(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HARVESTINGBASED_ABW" & _
										"        , field_api.getvalid_main_road(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) MAIN_ROAD" & _
										"        , field_api.getvalid_collectionroad(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) COLLECTIONROAD" & _
										"        , field_api.getvalid_maindrain(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))MAINDRAIN" & _
										"        , field_api.getvalid_collectiondrain(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) COLLECTIONDRAIN" & _
										"        , field_api.getvalid_subsidiarydrain(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) SUBSIDIARYDRAIN" & _
										"        , field_api.getvalid_collectionpoint(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) COLLECTIONPOINT" & _
										"        , field_api.getvalid_harvestingbridge(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) HARVESTINGBRIDGE" & _
										"        , field_api.getvalid_permanentbridge(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) PERMANENTBRIDGE" & _
										"        , field_api.getvalid_woodbridge(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) WOODBRIDGE" & _
										"        , field_api.getvalid_buist(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) BUIST" & _
										"        , field_api.getvalid_owlnest(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) OWLNEST" & _
										" 		 , '" & Session("CodeName") & "' FCENTRY" & _
								       	" 	 	 , '" & Session("CodeName") & "'FCEDIT" & _
								       	" 		 , '" & Request.ServerVariables("REMOTE_ADDR") & "' FCIP" & _
								       	"		 , '" & Session("fcba") & "' FCBA " & _
										"        , sysdate LASTUPDATE" & _
										"        , to_char (sysdate, 'hh24:mi')LASTTIME" & _
										"        , field_api.getvalid_harvestingsection(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) HARVESTINGSECTION" & _
										"        , field_api.getvalid_field_info_01(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_01" & _
										"        , field_api.getvalid_field_info_02(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_02" & _
										"        , field_api.getvalid_field_info_03(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_03" & _
										"        , field_api.getvalid_field_info_04(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_04" & _
										"        , field_api.getvalid_field_info_05(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) FIELD_INFO_05" & _
										"        , field_api.getvalid_isactive_harvested(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy')) ISACTIVE_HARVESTED" & _
										"        , trunc(to_date('"& transdate &"', 'mm/dd/yyyy'),'mm') VALID_FROM " & _
										"        , (select nvl(max(lineno), 0) from fielddetail fd where fd.fccode = ff.fccode and fd.fcba = ff.fcba) +1 LINENO , 'Updated' ACTION_STATE " & _
										"        , field_api.getvalid_AREAL_LC(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))AREAL_LC" & _
										"        , field_api.getvalid_CADANGAN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))CADANGAN" & _
										"        , field_api.getvalid_OKUPASI(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))OKUPASI" & _
										"        , field_api.getvalid_BIBITAN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))BIBITAN" & _
										"        , field_api.getvalid_PABRIK(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PABRIK" & _
										"        , field_api.getvalid_EMPLASMEN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))EMPLASMEN" & _
										"        , field_api.getvalid_PARIT(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PARIT" & _
										"        , field_api.getvalid_JALAN(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))JALAN" & _
										"        , field_api.getvalid_PRASARANA_LAINNYA(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PRASARANA_LAINNYA" & _
										"        , field_api.getvalid_QUARRY(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))QUARRY" & _
										"        , field_api.getvalid_TANGGUL(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))TANGGUL" & _
										"        , field_api.getvalid_ENCLAVE(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))ENCLAVE" & _
										"        , field_api.getvalid_TAMBANG(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))TAMBANG" & _
										"        , field_api.getvalid_KONSERVASI(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))KONSERVASI" & _
										"        , field_api.getvalid_BUKIT_SUNGAI(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))BUKIT_SUNGAI" & _
										"        , field_api.getvalid_NO_IMAGERY(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))NO_IMAGERY" & _
										"        , field_api.getvalid_LUAS_GROSS(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))LUAS_GROSS" & _
										"        , field_api.getvalid_HECTARAGEHARVESTED(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HECTARAGEHARVESTED" & _
										"        , field_api.getvalid_PRODUCTIVESTAND(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PRODUCTIVESTAND" & _
										"        , field_api.getvalid_HARVESTINGSECTIONNUMB(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))HARVESTINGSECTIONNUMBER" & _
										"        , field_api.getvalid_PLANTINGPOINT(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))PLANTINGPOINT" & _
										"        , field_api.getvalid_FULLSTAND(fccode, fcba, to_date('"& transdate &"', 'mm/dd/yyyy'))FULLSTAND" & _
										" 		 from field ff" & _
										" 		 where fccode = '"& fieldcode &"'" & _
										"        and fcba = '" & Session("fcba") & "'"								
										'Response.Write q
										Set rsAvail = Server.CreateObject("ADODB.Recordset")
										rsAvail.Open strFieldDetail, connstr

							if not rsAvail.Eof then
								do while not rsAvail.eof
									q = "insert into fielddetail (FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE, BUIST, OWLNEST, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, HARVESTINGSECTION, FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05, ISACTIVE_HARVESTED, VALID_FROM, LINENO, ACTION_STATE, AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA, QUARRY, TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER, PLANTINGPOINT, FULLSTAND) Values (" & _
									"  '" & rsAvail("FCCODE") & "'" & _
									" , '" & rsAvail("FCNAME") & "'" & _
									" , '" & rsAvail("DIVISION") & "'" & _
									" , '" & Post_hectarageplanted  & "'" & _
									" , '" & rsAvail("TERRAINTYPE") & "'" & _
									" , '" & rsAvail("SOILTYPE") & "'" & _
									" , '" & rsAvail("OWNERSHIP") & "'" & _
									" , '" & rsAvail("ACTIVATION") & "'" & _
									" , '" & rsAvail("CROPTYPE") & "'" & _
									" , to_date('" & rsAvail("PLANTINGDATE") & "','mm/dd/yyyy')" & _
									" , '" & rsAvail("PLANTINGMATERIAL") & "'" & _
									" , '" & post_totalstand & "'" & _
									" , '" & rsAvail("STATUS") & "'" & _
									" , '" & rsAvail("PREVSEMESTER_ABW") & "'" & _
									" , '" & rsAvail("PRESSEMESTER_ABW") & "'" & _
									" , '" & rsAvail("HARVESTINGBASED_ABW") & "'" & _
									" , '" & rsAvail("MAIN_ROAD") & "'" & _
									" , '" & rsAvail("COLLECTIONROAD") & "'" & _
									" , '" & rsAvail("MAINDRAIN") & "'" & _
									" , '" & rsAvail("COLLECTIONDRAIN") & "'" & _
									" , '" & rsAvail("SUBSIDIARYDRAIN") & "'" & _
									" , '" & rsAvail("COLLECTIONPOINT") & "'" & _
									" , '" & rsAvail("HARVESTINGBRIDGE") & "'" & _
									" , '" & rsAvail("PERMANENTBRIDGE") & "'" & _
									" , '" & rsAvail("WOODBRIDGE") & "'" & _
									" , '" & rsAvail("BUIST") & "'" & _
									" , '" & rsAvail("OWLNEST") & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									" , '" & Session("fcba") & "'" & _
									" , sysdate" & _
									" , to_char (sysdate, 'hh24:mi')" & _
									" , '" & rsAvail("HARVESTINGSECTION") & "'" & _
									" , '" & rsAvail("FIELD_INFO_01") & "'" & _
									" , '" & rsAvail("FIELD_INFO_02") & "'" & _
									" , '" & rsAvail("FIELD_INFO_03") & "'" & _
									" , '" & rsAvail("FIELD_INFO_04") & "'" & _
									" , '" & rsAvail("FIELD_INFO_05") & "'" & _
									" , '" & rsAvail("ISACTIVE_HARVESTED") & "'" & _
									" , to_date('" & TRANSDATE & "','mm/dd/yyyy')" & _
									" , '" & rsAvail("LINENO") & "'" & _
									" , '" & rsAvail("ACTION_STATE") & "'" & _
									" , '" & rsAvail("AREAL_LC") & "'" & _
									" , '" & rsAvail("CADANGAN") & "'" & _
									" , '" & rsAvail("OKUPASI") & "'" & _
									" , '" & rsAvail("BIBITAN") & "'" & _
									" , '" & rsAvail("PABRIK") & "'" & _
									" , '" & rsAvail("EMPLASMEN") & "'" & _
									" , '" & rsAvail("PARIT") & "'" & _
									" , '" & rsAvail("JALAN") & "'" & _
									" , '" & rsAvail("PRASARANA_LAINNYA") & "'" & _
									" , '" & rsAvail("QUARRY") & "'" & _
									" , '" & rsAvail("TANGGUL") & "'" & _
									" , '" & rsAvail("ENCLAVE") & "'" & _
									" , '" & rsAvail("TAMBANG") & "'" & _
									" , '" & rsAvail("KONSERVASI") & "'" & _
									" , '" & rsAvail("BUKIT_SUNGAI") & "'" & _
									" , '" & rsAvail("NO_IMAGERY") & "'" & _
									" , '" & rsAvail("LUAS_GROSS") & "'" & _
									" , '" & rsAvail("HECTARAGEHARVESTED") & "'" & _
									" , '" & rsAvail("PRODUCTIVESTAND") & "'" & _
									" , '" & rsAvail("HARVESTINGSECTIONNUMBER") & "'" &_
									" , '" & rsAvail("PLANTINGPOINT") & "'" &_
									" , '" & fullstand & "'" &_
									")" 
									'Response.Write q
									connStr.execute(q)
									rsAvail.MoveNext
									
								loop
							End If
						End if
						End If
					End If
				End If
			End If
		End If
		End If
		End If
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
				If Len(Trim(fieldcode)) = 0 Then
					Err.Number = 1001
					Err.Description = "Fieldcode Harus diisi, tidak boleh kosong"
				Else
					Dim val_SupplyingCodes : val_SupplyingCodes = 0
					val_SupplyingCodes = CheckFieldCode(fieldcode, session("fcBA"), TRANSDATE)
				If val_SupplyingCodes = 0 Then
					Err.Number = 1001
					Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
				Else
					
				' Dim val_Plantingdates : val_Plantingdates = 0
				' val_Plantingdates = CheckPlantingDate(fieldcode, session("fcBA"), TRANSDATE)
				' Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
				
				' If cstr(""&declarationtype) = cstr("1")  Then
					If QUANTITY <= 0 Then
						err.number = 1001
						Err.Description = "Anda harus Masukkan Angka di atas 0"
					
				Else
					q =	" Update plantingdeclarationdtl" & _
						" set quantity = '" & quantity & "'" & _
						"     , hectare = '" & hectare & "'" & _
						"     , remarks = '" & remarks & "'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where documentno = '" & txtDocumentNo & "'" & _
						"       and fieldcode = '" & fieldcode & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
		 			' response.write q
					connStr.execute(q)
					
					End If
				End If
			End If
		End If
	End If
	cmdView	= "True"	
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

If Len(Trim(TRANSDATE)) = 0 Then
	TRANSDATE = textDate(date)
Else
	TRANSDATE = Request("TRANSDATE")
End If

Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select documentno, to_char(TRANSDATE,'mm/dd/yyyy') TRANSDATE" & _
							" from plantingdeclarationhdr" & _
							" where documentno = '" & txtDocumentNo & "'" & _
							" and fcba = '" & session("fcBA") & "'"
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	txtDocumentNo	=	rsSupplyingDeclaration("documentno")
	TRANSDATE	=	rsSupplyingDeclaration("TRANSDATE")	
End if
'response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

' Dim val_IsTrue : val_IsTrue = "N"
'For Show Tabel Supplying Detail
Dim oRS, oSQL
oSQL = 	  " SELECT pdd.documentno, pdh.transdate, pdd.hectare, pdd.fieldcode, pdd.quantity, pdd.remarks " & _
		  " FROM plantingdeclarationdtl pdd" & _
		  " INNER JOIN plantingdeclarationhdr pdh"& _
		  " ON (pdd.documentno = pdh.documentno AND pdd.fcba = pdh.fcba)" & _
		  " WHERE pdd.fcba = '" & session("fcBA") & "'" & _
		  " AND pdd.documentno = '" & txtDocumentNo & "'"
' response.write oSQL	   
set oRS = server.CreateObject("Adodb.recordset")
oRS.Open oSQL, connstr

Dim val_IsTrue : val_IsTrue = "N"

' oSQLQuery = " select 'Y' istrue" & _
'             " from periodcontrol" & _
' 	        " where fcba = '" & Session("fcBA") & "'" & _
' 	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') >= currentdate" & _
' 	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
' 	        "		AND TO_CHAR(TO_DATE ('" & TRANSDATE & "', 'mm/dd/yyyy'),'yyyy') = currentaccountingyear" & _
' 	        "       and upper(system) = upper('costbook')"

  oSQLQuery =" SELECT 'Y' istrue " & _
  			 " FROM periodcontrol" & _
 			 " WHERE fcba = '" & Session("fcBA") & "'" & _
       		 " AND TO_DATE ('" & TRANSDATE & "', 'mm/dd/yyyy') between currentdate and TO_DATE ('12/31/' || currentaccountingyear , 'mm/dd/yyyy')" & _
       		 " AND UPPER (SYSTEM) = UPPER ('costbook')"

			' response.write oSQLQuery
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if

'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing

' Dim val_Fullstand : val_Fullstand = "N"

' oSQLQuery = " select 'Y' istrue" & _
'             " from fielddetail" & _
' 	        " where fcba = '" & Session("fcBA") & "'" & _
' 	        "       and fccode ='" & fieldcode & "'" & _
' 	        "       and fullstand ='" & fullstand & "'" 
' set oRecordSet = server.CreateObject("Adodb.recordset")
' oRecordSet.Open oSQLQuery, connstr
' If Not oRecordSet.Eof then
' 	val_Fullstand = oRecordSet("istrue")
' Else
' 	val_Fullstand = "N"
' End if

' 'response.Write strNurseryCulling
' oRecordSet.Close
' Set oRecordSet = Nothing
%>

 <!-------------ISI Supplying Declaration Form!------------->
 
	<form name="frmEntry" method="post" action="plantingdeclaration.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Document No</td>
			<td align="left">
				<input type="text" maxlength="15" name="txtDocumentNo" size="30" value="<%=txtDocumentNo%>" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Document Date</td>
			<td align="left"><input class="datepicker" type="text" size="15" name="TRANSDATE" value="<%=TRANSDATE%>"></td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				<%If val_IsTrue = "Y" Then%>
					<a href="plantingdeclaration.asp?txtDocumentNo=<%=txtDocumentNo%>&TRANSDATE=<%=TRANSDATE%>&txtJenis=HeaderDelete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				<%Else%>
					 <img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> 
				<%End If%>
				<!-- HADIR -->
				<!-- <%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Culling"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a> -->
				<!-- END:HADIR -->
			</td>
		</tr>
		</table>
	</form>
<%
if cmdView = "True" and len(trim(txtDocumentNo)) > 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " SELECT pdd.documentno, pdh.transdate, pdd.hectare, pdd.fieldcode, pdd.quantity, pdd.lineno, pdd.fullstand, pdd.remarks " & _
		  " FROM plantingdeclarationdtl pdd" & _
		  " INNER JOIN plantingdeclarationhdr pdh"& _
		  " ON (pdd.documentno = pdh.documentno AND pdd.fcba = pdh.fcba)" & _
		  " WHERE pdd.fcba = '" & session("fcBA") & "'" & _
		  " AND pdd.documentno = '" & txtDocumentNo & "'"

	'response.Write q

	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="plantingdeclaration.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td align="left" width="20%">Field Code</td>
				<td align="left">
					<input type="text" size="15" name="FIELDCODE" style="background:#FFFFCC;"  value="<%=fieldcode%>" maxlength="32" autocomplete="off">
					<font color="#FF0000">*</font>
				</td>
			</tr>
			<tr> 
	            <td align="left">Hectare</td>
				<td align="left"> 
					<input type="text" size="15" name="hectare" style="background:#FFFFCC;"  value="<%=hectare%>" maxlength="32" autocomplete="off" onkeypress="return hanyaAngka(event)" >	
					<font color="#FF0000">*</font>
				</td>
          	</tr>
			<tr>
				<td align="left" width="20%">Quantity</td>
				<td align="left">
					<input type="text" size="15" name="QUANTITY" style="background:#FFFFCC;" value="<%=quantity%>" autocomplete="off" maxlength="32" onkeypress="return hanyaAngka(event)">
					<font color="#FF0000">*</font>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Remarks</td>
				<td align="left">
					<textarea name="REMARKS" cols="50" rows="3" maxlength="400" autocomplete="off" ><%=remarks%></textarea>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Full Stand</td>
				<td>
					<input style="float: left;" type="checkbox" id="fullstand" name="fullstand" value="<%=fullstand%>" onclick="myFunction()">
					<p id="text" style="display:none; padding-right: 700px;">Full Stand</p>
				</td>

			</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="txtDocumentNo" value="<%=txtDocumentNo%>">
					<input type="hidden" name="TRANSDATE" value="<%=TRANSDATE%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If val_IsTrue = "Y" Then%>
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input type="hidden" name="txtlineno" value="<%=txtlineno%>">
						<input Type="Submit"  name="Add" value="Update">
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add">
					<%End if%>

					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				<%End if%>	
				</td>
			</tr>
		</table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="1" >&nbsp;</td>
		<td align="left">DOCUMENT NO</td>
		<td align="left">DOCUMENT DATE</td>
		<td align="left">FIELDCODE</td>
		<td align="left">QUANTITY</td>
		<td align="left">HECTARE</td>
		<td align="left">REMARKS</td>
		<td align="left">FULL STAND</td>
	</tr>
   
   <%
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
			
			
			<td valign="middle" >
				<%If val_IsTrue = "Y" Then%>
				<a href="plantingdeclaration.asp?txtlineno=<%=rsVH("lineno")%>&txtDocumentNo=<%=txtDocumentNo%>&fullstand=<%=rsVH("fullstand")%>&transdate=<%=rsVH("TRANSDATE")%>&hectare=<%=rsVH("hectare")%>&quantity=<%=rsVH("quantity")%>&fieldcode=<%=rsVH("fieldcode")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				<%Else%>
					 <img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> 
				<%End If%>
			</td>

			    <td align="left" title="<%=rsVH("documentno").Name%>" valign="top"><%=rsVH("documentno")%></td>
			    <td align="left" title="<%=rsVH("transdate").Name%>" valign="top"><%=rsVH("transdate")%></td>
			    <td align="left" title="<%=rsVH("fieldcode").Name%>" valign="top"><%=rsVH("fieldcode")%></td>
			    <td align="left" title="<%=rsVH("quantity").Name%>" valign="top"><%=rsVH("quantity")%></td>
			    <td align="left" title="<%=rsVH("hectare").Name%>" valign="top"><%=rsVH("hectare")%></td>
			    <td align="left" title="<%=rsVH("remarks").Name%>" valign="top"><%=rsVH("remarks")%></td>
			    <td align="left" title="<%=rsVH("fullstand").Name%>" valign="top"><%=rsVH("fullstand")%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

		<script>
		function myFunction() {
		    var checkBox = document.getElementById("fullstand");
		    var text = document.getElementById("text");
		    if (checkBox.checked === true){
		        text.style.display = "block";
		        checkBox.value="TRUE";
		    } else {
		       text.style.display = "none";
		       checkBox.value="FALSE";
		    }
		}
		</script>

</body>
</html>
