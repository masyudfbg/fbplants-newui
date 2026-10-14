<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - BBC Transaction</title>

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
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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
Dim txtmonth, txtyear 
Dim parcode, totalstand, ScriptName
Dim QUANTITY, censuscode, fieldcode
Dim HarvestingBased_ABW, HectarageHarvested, ProductiveStand, SampleStand, SampleBunch
Dim TRANSDATE
Dim cmdView
Dim distribution1, distribution2, distribution3, distribution4
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

txtyear =request("txtyear")
censuscode =request("censuscode")
HarvestingBased_ABW=request("HarvestingBased_ABW")
remarks=request("remarks")
fieldcode=request("fieldcode")
HectarageHarvested = Request("HectarageHarvested")
ProductiveStand = Request("ProductiveStand")
SampleStand = Request("SampleStand")
SampleBunch = Request("SampleBunch")

mon = Month(Session("v_Costbook"))
If len(txtmonth)=0 then txtmonth=mon

Yr = Year(Session("v_Costbook"))

If len(txtyear)= 0 then txtyear = Yr
transdate = txtmonth & "/1/" & txtyear

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   txtyear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   txtmonth	=	Month(Session("v_Costbook"))
end if


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
	'Select count(1) from dual where field_api.GetValid_Activation('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')) = 'Y'
	'select to_char (field_api.GetValid_Plantingdate('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')),'yyyy') from dual
End Function

Function CheckHaHarvested(initialFieldCode, InitialCensuscode, initialFcba, initDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" select hectarageharvested from bbctransaction total" & _
				    " WHERE periodyear = '" & initDate & "'" & _
				    " AND censuscode = '" & InitialCensuscode & "'" & _
				    " AND fcba = '" & initialFcba & "'"  & _
				    " AND fieldcode = '" & initialFieldCode & "'" 

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

	CheckHaHarvested = initialAvailable
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
	
	strAvail =	" select productivestand from bbctransaction total" & _
			    " WHERE periodyear = '" & initDate & "'" & _
			    " AND censuscode = '" & InitialCensuscode & "'" & _
			    " AND fcba = '" & initialFcba & "'"  & _
			    " AND fieldcode = '" & initialFieldCode & "'"  
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

If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			q = " delete from bbctransaction " & _
						" 		where periodyear = '" & txtyear & "'" & _
						"       and censuscode = '" & censuscode & "'" & _
						"       and fieldcode = '" & fieldcode & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
				' Response.Write q
				connStr.execute(q)

			Current_TotalStand = F_CurrentTotalStand(fieldcode, session("fcBA"), TRANSDATE)
			Current_Hectare = CheckHaHarvested(fieldcode, session("fcBA"), TRANSDATE, censuscode)
			Post_totalstand = cdbl(to_nol(PRODUCTIVESTAND)) - cdbl(to_nol(PRODUCTIVESTAND))
			post_Hectare = cdbl(to_nol(HECTARAGEHARVESTED)) - cdbl(to_nol(HECTARAGEHARVESTED))

			a =	" Update field" & _
				" set HectarageHarvested = '" & post_Hectare  & "'" & _
				"      ,ProductiveStand = '" & Post_totalstand  & "'" & _
				" 		where fcba = '" & Session("fcba") & "'" & _
				"		and fccode = '" & fieldcode & "'"
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
						" , '" & rsAvail("HECTARAGEPLANTED")  & "'" & _
						" , '" & rsAvail("TERRAINTYPE") & "'" & _
						" , '" & rsAvail("SOILTYPE") & "'" & _
						" , '" & rsAvail("OWNERSHIP") & "'" & _
						" , '" & rsAvail("ACTIVATION") & "'" & _
						" , '" & rsAvail("CROPTYPE") & "'" & _
						" , to_date('" & rsAvail("PLANTINGDATE") & "','mm/dd/yyyy')" & _
						" , '" & rsAvail("PLANTINGMATERIAL") & "'" & _
						" , '" & rsAvail("TOTALSTAND") & "'" & _
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
						" , '" & Post_totalstand & "'" & _
						" , '" & post_Hectare & "'" & _
						" , '" & rsAvail("HARVESTINGSECTIONNUMBER") & "'" &_
						" , '" & rsAvail("PLANTINGPOINT") & "'" &_
						" , '" & rsAvail("FULLSTAND") & "'" &_
						")" 
						'Response.Write q
						connStr.execute(q)
						rsAvail.MoveNext
							
						loop
					End If
				cmdView	= "True"


	elseif Request("txtJenis")="New" then		
			
			Dim val_SupplyingCodes : val_SupplyingCodes = 0
				val_SupplyingCodes = CheckFieldCode(fieldcode, session("fcBA"), "01/01" & txtyear)
			If val_SupplyingCodes = 0 Then
				Err.Number = 1001
				Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
			Else	
				If cdbl(to_nol(HectarageHarvested)) = 0 Then
					Err.Number = 1001
					Err.Description = "Hectarage Harvested Harus diisi, tidak boleh kosong"
				Else
					If cdbl(to_nol(ProductiveStand)) = 0 Then
						Err.Number = 1001
						Err.Description = "Productive Stand Harus diisi, tidak boleh kosong"
					Else
						If cdbl(to_nol(SampleStand)) = 0 Then
							Err.Number = 1001
							Err.Description = "Sample Stand Harus diisi, tidak boleh kosong"
						Else
						If cdbl(to_nol(SampleStand)) > cdbl(to_nol(ProductiveStand)) Then
							Err.Number = 1001
							Err.Description = "Sample Stand tidak boleh lebih dari Productive Stand"
						Else
							If cdbl(to_nol(SampleBunch)) = 0 Then
								Err.Number = 1001
								Err.Description = "Sample Bunch Harus diisi, tidak boleh kosong"
							Else
								If cdbl(to_nol(HarvestingBased_ABW)) = 0 Then
									Err.Number = 1001
									Err.Description = "Harvestingbased ABW Harus diisi, tidak boleh kosong"
								Else

							'Validasi untuk Hectarage Harvested tidak boleh lebih dari ha planted master field
								
							    strHectare =  " select field_api.getvalid_Ha('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& transdate &"', 'mm/dd/yyyy')) total" & _
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
								' response.write HectarageHarvested&"<br>"
								' response.write maxhectare&"<br>"
								' response.write val_QtyHectare&"<br>"
								If cdbl (to_nol(HectarageHarvested)) > cdbl (to_nol(maxhectare)) Then
									Err.Number = 1001
									Err.Description = "Data tidak berhasil disimpan, Total Hectarage Planted = """ & maxhectare & """,  mohon cek kembali data yang anda masukan"
								Else
								
								'Validasi untuk produktif stand tidak boleh lebih dari total stand master field

							    strPokok =  " select field_api.getValid_POKOK('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& transdate &"', 'mm/dd/yyyy')) total" & _
	            							  " from dual " 
												' Response.Write strPokok

												Set rsQty = Server.CreateObject("ADODB.Recordset")
												rsQty.Open strPokok, connstr							
								if not rsQty.Eof then
								do while not rsQty.eof
										maxstand = cdbl (to_nol(rsQty("total")))
										rsQty.MoveNext
									loop
								End If
								' response.write ProductiveStand&"<br>"
								' response.write maxstand&"<br>"
								' response.write val_QtyHectare&"<br>"
								If cdbl (to_nol(ProductiveStand)) > cdbl (to_nol(maxstand)) Then
									Err.Number = 1001
									Err.Description = "Data tidak berhasil disimpan, Total Stand = """ & maxstand & """,  mohon cek kembali data yang anda masukan"
								Else

									q = " INSERT INTO bbctransaction" & _
										" (periodyear, censuscode, fieldcode, HectarageHarvested, ProductiveStand, SampleStand, SampleBunch, HarvestingBased_ABW, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) VALUES (" & _
										" '" & txtyear & "'" & _
										" , '" & censuscode & "'" & _
										" , '" & fieldcode & "'" & _
										" , '" & HectarageHarvested & "'" & _
										" , '" & ProductiveStand & "'" & _
										" , '" & SampleStand & "'" & _
										" , '" & SampleBunch & "'" & _
										" , '" & HarvestingBased_ABW & "'" & _
										" , '" & Session("CodeName") & "'" & _
										" , '" & Session("CodeName") & "'" & _
										" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
										" , '" & Session("fcba") & "'" & _
										" , sysdate" & _
										" , to_char(sysdate, 'hh24:mi')" & _
										" )"
										' Response.Write q
										connStr.execute(q)
									
									a =	" Update field" & _
										" set HectarageHarvested = '" & HectarageHarvested  & "'" & _
										"      ,ProductiveStand = '" & ProductiveStand  & "'" & _
										" 		where fcba = '" & Session("fcba") & "'" & _
										"		and fccode = '" & fieldcode & "'"
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
											" , '" & rsAvail("HECTARAGEPLANTED")  & "'" & _
											" , '" & rsAvail("TERRAINTYPE") & "'" & _
											" , '" & rsAvail("SOILTYPE") & "'" & _
											" , '" & rsAvail("OWNERSHIP") & "'" & _
											" , '" & rsAvail("ACTIVATION") & "'" & _
											" , '" & rsAvail("CROPTYPE") & "'" & _
											" , to_date('" & rsAvail("PLANTINGDATE") & "','mm/dd/yyyy')" & _
											" , '" & rsAvail("PLANTINGMATERIAL") & "'" & _
											" , '" & rsAvail("TOTALSTAND") & "'" & _
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
											" , '" & HECTARAGEHARVESTED & "'" & _
											" , '" & PRODUCTIVESTAND & "'" & _
											" , '" & rsAvail("HARVESTINGSECTIONNUMBER") & "'" &_
											" , '" & rsAvail("PLANTINGPOINT") & "'" &_
											" , '" & rsAvail("FULLSTAND") & "'" &_
											")" 
											'Response.Write q
											connStr.execute(q)
											rsAvail.MoveNext
											
										loop
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		End If
		End If

	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			
			Dim val_SupplyingCode : val_SupplyingCode = 0
				val_SupplyingCode = CheckFieldCode(fieldcode, session("fcBA"), "01/01" & txtyear)
			If val_SupplyingCode = 0 Then
				Err.Number = 1001
				Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
			Else	
				If cdbl(to_nol(HectarageHarvested)) = 0 Then
					Err.Number = 1001
					Err.Description = "Hectarage Harvested Harus diisi, tidak boleh kosong"
				Else
					If cdbl(to_nol(ProductiveStand)) = 0 Then
						Err.Number = 1001
						Err.Description = "Productive Stand Harus diisi, tidak boleh kosong"
					Else
						If cdbl(to_nol(SampleStand)) = 0 Then
							Err.Number = 1001
							Err.Description = "Sample Stand Harus diisi, tidak boleh kosong"
						Else
						If cdbl(to_nol(SampleStand)) > cdbl(to_nol(ProductiveStand)) Then
							Err.Number = 1001
							Err.Description = "Sample Stand tidak boleh lebih dari Productive Stand"
						Else
							If cdbl(to_nol(SampleBunch)) = 0 Then
								Err.Number = 1001
								Err.Description = "Sample Bunch Harus diisi, tidak boleh kosong"
							Else
								If cdbl(to_nol(HarvestingBased_ABW)) = 0 Then
									Err.Number = 1001
									Err.Description = "Harvestingbased ABW Harus diisi, tidak boleh kosong"
								Else

							'Validasi untuk Hectarage Harvested tidak boleh lebih dari ha planted master field
								
							    strHectare =  " select field_api.getvalid_Ha('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& transdate &"', 'mm/dd/yyyy')) total" & _
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
								' response.write HectarageHarvested&"<br>"
								' response.write maxhectare&"<br>"
								' response.write val_QtyHectare&"<br>"
								If cdbl (to_nol(HectarageHarvested)) > cdbl (to_nol(maxhectare)) Then
									Err.Number = 1001
									Err.Description = "Data tidak berhasil disimpan, Total Hectarage Planted = """ & maxhectare & """,  mohon cek kembali data yang anda masukan"
								Else
								
								'Validasi untuk produktif stand tidak boleh lebih dari total stand master field

							    strHectare =  " select field_api.getValid_POKOK('" & fieldcode & "', '" & session("fcBA") & "', to_date('"& TRANSDATE &"', 'mm/dd/yyyy')) total" & _
	            							  " from dual " 
												' Response.Write strHectare

												Set rsQty = Server.CreateObject("ADODB.Recordset")
												rsQty.Open strHectare, connstr							
								if not rsQty.Eof then
								do while not rsQty.eof
										maxstand = cdbl (to_nol(rsQty("total")))
										rsQty.MoveNext
									loop
								End If
								' response.write ProductiveStand&"<br>"
								' response.write maxstand&"<br>"
								' response.write val_QtyHectare&"<br>"
								If cdbl (to_nol(ProductiveStand)) > cdbl (to_nol(maxstand)) Then
									Err.Number = 1001
									Err.Description = "Data tidak berhasil disimpan, Total Stand = """ & maxstand & """,  mohon cek kembali data yang anda masukan"
								Else
									q =	" Update bbctransaction" & _
										" set  fieldcode = '" & fieldcode & "'" & _
										"     , HectarageHarvested = '" & HectarageHarvested & "'" & _
										"     , ProductiveStand = '" & ProductiveStand & "'" & _
										"     , SampleStand = '" & SampleStand & "'" & _
										"     , SampleBunch = '" & SampleBunch & "'" & _
										"     , HarvestingBased_ABW = '" & HarvestingBased_ABW & "'" & _
										"     , fcentry = '" & Session("CodeName") & "'" & _
										"     , fcedit = '" & Session("CodeName") & "'" & _
										" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
										"     , lastupdate = sysdate " & _
										"     , lasttime = to_char(sysdate,'hh24:mi') " & _
										" 		where periodyear = '" & txtyear & "'" & _
										"       and censuscode = '" & censuscode & "'" & _
										"       and fcBA = '" & Session("fcBA") & "'" &_
										"       and fieldcode = '" & fieldcode & "'" 
						 			' response.write q
									connStr.execute(q)
									
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		End If
	cmdView	= "True"
	else
		cmdView	= "True"
 	End if
	Else
	 	cmdview = request("cmdview")
 	End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")


If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('"& txtmonth &"'||'/01/'||'" & txtyear &"', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_char(to_date('"& txtmonth &"'||'/01/'||'" & txtyear &"','mm/dd/yyyy'),'yyyy')= to_char(currentdate,'yyyy')" & _
	        "       and upper(system) = upper('costbook')"
			
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

'response.write cmdView
%>

 <!-------------ISI Supplying Declaration Form!------------->
 
	<form name="frmEntry" method="post" action="<%=ScriptName%>">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Year</td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, MM

					MinYr = Yr - 2
					MaxYr = Yr  + 2
				%>
			<td align="left">
				<select name="txtYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(txtyear) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
				</select>
			</td>
		</tr>
		<tr> 
	            <td align="left"><font face="Verdana" size="1">Census Code</font></td>
					<td align="left"> 
						<select name="censuscode" required="required"  title="CENSUS CODE Character a s/d z or 1 s/d 9">
							<option value=""></option> 
								<% 	Dim rsComboV
								q = "SELECT * FROM PARAMETER WHERE PARCODE = 'BBC' and fcba = '" & Session("fcba") & "'"
								Set rsComboV = Server.CreateObject("ADODB.Recordset")
								rsComboV.CursorType = 3
								rsComboV.CursorLocation = 3
								rsComboV.Open q, connstr
								if not rsComboV.eof then%>
								<%
								Do While Not rsCombov.eof
								%>
									<option value="<%=rsComboV("FCCODE")%>" 
										<%if cstr(""&censuscode)=cstr(""&rsComboV("FCCODE")) then %> selected 
										<%End If%>><%=rsComboV("FCNAME")%>
									</option> 
								<%rsComboV.movenext
								loop
								End if
								rsComboV.Close
								set rsComboV = Nothing
								q=""
								%>
						</select> 
				</td>
          </tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				<!-- <img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> -->
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
if cmdView = "True" and len(trim(txtyear)) > 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " SELECT * FROM bbctransaction " & _
		  " WHERE periodyear = '" & txtyear & "'" & _
		  " AND censuscode = '" & censuscode & "'" & _
		  " AND fcba = '" & session("fcBA") & "'"   
	'response.Write q

	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="bbctransaction.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			
			<tr>
				<td align="left" width="20%">Field Code</td>
				<td align="left">
					<input type="text" size="15" name="FIELDCODE" style="background:#FFFFCC;"  value="<%=fieldcode%>" maxlength="32" autocomplete="off">

					<font color="#FF0000">*</font>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Hectarage Harvested</td>
				<td align="left">
					<input type="text" size="15" name="HectarageHarvested" style="background:#FFFFCC;" value="<%=HectarageHarvested%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);">
					HA
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Productive Stand</td>
				<td align="left">
					<input type="text" size="15" name="ProductiveStand" style="background:#FFFFCC;" value="<%=ProductiveStand%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					PKK
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Sample Stand</td>
				<td align="left">
					<input type="text" size="15" name="SampleStand" style="background:#FFFFCC;" value="<%=SampleStand%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					PKK
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Harvesting Based ABW (BJR)</td>
				<td align="left">
					<input type="text" size="15" name="HarvestingBased_ABW" style="background:#FFFFCC;" value="<%=HarvestingBased_ABW%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					KG
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Sample Bunch</td>
				<td align="left">
					<input type="text" size="15" name="SampleBunch"  style="background:#FFFFCC;" value="<%=SampleBunch%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					JJG
				</td>
			</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="txtyear" value="<%=txtyear%>">
					<input type="hidden" name="censuscode" value="<%=censuscode%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If val_IsTrue = "Y" Then%>
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
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
		<td align="left">CENSUS CODE</td>
		<td align="left">FIELDCODE</td>
		<td align="left">HECTARAGE HARVESTED</td>
		<td align="left">PRODUCTIVE STAND</td>
		<td align="left">SAMPLE STAND</td>
		<td align="left">SAMPLE BUNCH</td>
		<td align="left">HARVESTINGBASED ABW</td>
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
						
			<td valign="middle" width="50" >
				<%If val_IsTrue = "Y" Then%>

					<a href="bbctransaction.asp?txtyear=<%=txtyear%>&censuscode=<%=rsVH("censuscode")%>&fieldcode=<%=rsVH("fieldcode")%>&HectarageHarvested=<%=rsVH("HectarageHarvested")%>&ProductiveStand=<%=rsVH("ProductiveStand")%>&SampleStand=<%=rsVH("SampleStand")%>&SampleBunch=<%=rsVH("SampleBunch")%>&HarvestingBased_ABW=<%=rsVH("HarvestingBased_ABW")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtyear")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> 
				<%end if%>
				
			</td>
			    <td align="left" title="<%=rsVH("censuscode").Name%>" valign="top"><%=rsVH("censuscode")%></td>
			    <td align="left" title="<%=rsVH("fieldcode").Name%>" valign="top"><%=rsVH("fieldcode")%></td>
			    <td align="left" title="<%=rsVH("HectarageHarvested").Name%>" valign="top"><%=rsVH("HectarageHarvested")%></td>
			    <td align="left" title="<%=rsVH("ProductiveStand").Name%>" valign="top"><%=rsVH("ProductiveStand")%></td>
			    <td align="left" title="<%=rsVH("SampleStand").Name%>" valign="top"><%=rsVH("SampleStand")%></td>
			    <td align="left" title="<%=rsVH("SampleBunch").Name%>" valign="top"><%=rsVH("SampleBunch")%></td>
			    <td align="left" title="<%=rsVH("HarvestingBased_ABW").Name%>" valign="top"><%=rsVH("HarvestingBased_ABW")%></td>
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

</body>
</html>
