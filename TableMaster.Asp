
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - Electronic Plantation Solutions Master</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />  
<link rel="StyleSheet" href="style-content-field-master.css" type="text/css" />  <!-- #51A351 = Hijau,  Additional Here Change background Color Input type=submit dan input type=button -->
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<!--
format Numeric Input ata ------------------------------------------>
<script>
	$(document).ready(function(){
    $('[class=numeric]').keypress(validateNumber);
	});
	function validateNumber(event) {
    var key = (event.which) ? event.which : event.keyCode ;
    
	
    if (event.ctrlKey || event.altKey || event.keyCode == 8 || event.keyCode == 46 
     || event.keyCode == 37 || event.keyCode == 39 || event.charCode == 188 || event.charCode == 190 || event.charCode == 110 )  {
        return true;
    }
    else if ( key < 48 || key > 57 ) {
        return false;
    }
    else return true;
	};
	
	
	function harusangka(evt) {
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
		}
}
	
	
	</script>

<!-- contoh HTML : <input class="numeric" type="text" size="20" title="CODE Number 1 s/d 9"> " -->

<!-- END format Numeric Input -->


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
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>


<script language="JavaScript" src="/js/validation_master_field.js" type="text/javascript"></script> <!--Additional Validation For Master Field -->
<script language="JavaScript" src="/js/gen_validatorv4.js" type="text/javascript"></script>

<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %>>
<!--#include file="top.jsp"-->

	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <%'-------------awal-----------------------
Dim cmdView,crashId,Pencarian,q_his,q_lineno,rsLineno,q_ipthis,q_input,qLineMax,rsqLineMax,tempMaxLine,txtLineNoHist,q1,q2
Dim ScriptName

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtJenis=Request("txtJenis")
q1 = ""
if txtTable = "EmployeeAllowdedDetail" then
crashId = "EMPLOYEECODE"
Pencarian = "Employee Code"
elseif Request("txttable") ="UntaxableIncome" then
crashId = "MARITALSTATUSCODE"
Pencarian = "Tax Code"
ELSE
crashId = "fcCode"
Pencarian = "Code"
end if

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("Add")= "Update" Then
	if Request("txtJenis")= "Delete" then
	    ' if txtModule <> "Administrator" then 
			if txtTable = "MaritalStatus" then
			  q="Delete from " & txtTable & " where  " & crashId & "='" & request(crashId) & "'"
			elseif txtTable = "EmpPosition" then
			  q="Delete from " & txtTable & " where POSITION ='" & request(crashId) & "' AND fcBA = '"& session("fcBA") &"'"
 			elseif txtTable = "Product" then
			  q="Delete from " & txtTable & " where PRODUCTCODE ='" & request(crashId) & "' AND fcBA = '"& session("fcBA") &"'"
			elseif txtTable = "SectionName" then
			  q="Delete from " & txtTable & " where FCCODE ='" & request(crashId) & "'"
			elseif txtTable = "MultiLevel_Harvesting_Rate" then
			  q = "Delete from " & txtTable & " where LEVELNO ='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "UnitOfMeasure" then
			  q="Delete from " & txtTable & " where UNITOFMEASURECODE ='" & request(crashId) & "' AND fcBA = '"& session("fcBA") &"'"
			elseif txtTable = "VehicleGroup" then
			  q="Delete from " & txtTable & " where VEHICLEGROUPCODE ='" & request(crashId) & "' AND fcBA = '"& session("fcBA") &"'"
			elseif txtTable = "CUSTOMER" then
			  q="UPDATE CUSTOMER SET ISACTIVE='N'  where CUSTOMERCODE ='" & request(crashId) & "' AND fcBA = '"& session("fcBA") &"'"
			elseif txtTable = "OpHarvRate" then
				qLineMax = " select nvl(max(LINENO),0) tempMaxLine" & _
						   " from OpHarvRate where fcba = '"& session("fcBA") &"'" 
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 
			  q="Delete from " & txtTable & " where CROPCODE ='" & request(crashId) & "' AND LINENO = '"& tempMaxLine &"' AND fcBA = '"& session("fcBA") &"'"
			elseif txtTable = "EmployeeAllowdedDetail" then
			q="Delete from " & txtTable & " where  " & crashId & "='" & request(crashId) & "' and ALLOWDEDCODE = '" & request("ALLOWDEDCODE") & "' AND FCBA ='" & Session("fcBA") & "'"
			elseif txtTable = "Field" then
			q_lineno = "select MAX(LINENO)+1 LINENO from FIELDDETAIL WHERE FCBA ='" & Session("fcBA") & "' AND FCCODE = '" & Request("fcCode") & "' "
				Set rsLineno = Server.CreateObject("ADODB.Recordset")
				rsLineno.Open q_lineno, connstr
			q_ipthis = "select * from FIELDDETAIL WHERE " & crashId & " = '" & request(crashId) & "' AND FCBA ='" & Session("fcBA") & "'"
				Set q_input = Server.CreateObject("ADODB.Recordset")
				q_input.Open q_ipthis, connstr
			q_his = "Insert into FIELDDETAIL (FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, " & _ 
				"PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, DX_BLOCK_STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, " & _ 
				"COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE, " & _ 
				"BUIST, OWLNEST, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, HARVESTINGSECTION, " & _
				"FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05, ISACTIVE_HARVESTED,VALID_FROM,LINENO,ACTION_STATE)" & _
				"values ('" &  q_input("FCCODE") & "','" & q_input("FCNAME") & "','"  &  q_input("DIVISION") & "','" & _
				q_input("HECTARAGEPLANTED") & "','" & q_input("TERRAINTYPE") & "','"  &  q_input("SOILTYPE") & "','" &  q_input("OWNERSHIP") & "','" &  q_input("ACTIVATION") & "','" &  q_input("CROPTYPE") & "'," & _
				" to_date('" & q_input("PLANTINGDATE") & "','mm/dd/yyyy') " & ",'" & q_input("PLANTINGMATERIAL") & "','" & _
				q_input("TOTALSTAND") & "','" & _
				q_input("DX_BLOCK_STATUS")  & "','" & q_input("PREVSEMESTER_ABW") & "','" & q_input("PRESSEMESTER_ABW") & "','" & _
				q_input("HARVESTINGBASED_ABW")  & "','" & q_input("MAIN_ROAD") & "','" & q_input("COLLECTIONROAD") & "','" & _
				q_input("MAINDRAIN")  & "','" & q_input("COLLECTIONDRAIN") & "','" & q_input("SUBSIDIARYDRAIN") & "','" & _
				q_input("COLLECTIONPOINT")  & "','" & q_input("HARVESTINGBRIDGE") & "','" & q_input("PERMANENTBRIDGE") & "','" & _
				q_input("WOODBRIDGE")  & "','" & q_input("BUIST") & "','" & q_input("OWLNEST") & "','" & _
				Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'), '" &_
				q_input("HARVESTINGSECTION") & "','" & q_input("FIELD_INFO_01")  & "','" & q_input("FIELD_INFO_02") & "','" &_ 
				q_input("FIELD_INFO_03") & "','" & q_input("FIELD_INFO_04")  & "','" & q_input("FIELD_INFO_05") & "','" &_
				q_input("ISACTIVE_HARVESTED") & "',to_date('" & q_input("VALID_FROM") & "','mm/dd/yyyy'),'" & rsLineno("LINENO") &"','Deleted')"
				'RESPONSE.WRITE(q_his)
				
			q2	=	" UPDATE FIELD_HIST_TAB SET" & _
					" ACTIVATION = 'N'" & _
					" ,LASTUPDATE = SYSDATE " & _
					" ,FCEDIT = '" & Session("CodeName") & "'" & _ 
					" WHERE fcCode = '" & Request("fcCode") & "'" & _
					"       AND fcBA = '" & Session("fcBA") & "'" & _ 
					"       AND year >= " & Request("YEAR") & "" & _ 
					"       AND month >= " & Request("MONTH") & "" 
		    'connStr.execute(q2)	
			ExecuteAndLog ScriptName, q2
				 
            q="Delete from " & txtTable & " where  " & crashId & "='" & request(crashId) & "' and fcBA='" & Session("fcBA") & "'"
			elseif txtTable = "FixedAsset" then
			 q="DELETE FROM fixedasset_hdr_tab WHERE  " & crashId & "='" & request(crashId) & "' AND fcBA='" & Session("fcBA") & "'"
			elseif txtTable = "BankStock" then
			 q="DELETE FROM BankStock WHERE  " & crashId & "='" & request(crashId) & "'"
			elseif  txtTable = "Roles" then             
		     q="Delete from Roles_Tab where  " & crashId & "='" & request(crashId) & "' "
			ElseIf Request("txttable") ="Employee_WorkingTime" then
			 q="Delete Employee_WorkingTime " & _
			   "WHERE  EMPCODE='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'" & _
			   "AND PERIOD_YEAR='" & Request("TXT_YEAR") & "' AND PERIOD_MONTH >='" & Request("TXT_MONTH")  & "'" 
			else
			 q="Delete from " & txtTable & " where  " & crashId & "='" & request(crashId) & "' and fcBA='" & Session("fcBA") & "'"
			end if
		' elseif  txtModule = "Administrator" and txtTable = "Roles" then             
		     ' q="Delete from Roles_Tab where  " & crashId & "='" & request(crashId) & "' "
        ' end if
	elseif Request("txtJenis")= "New" then  
	%>
	
	<!-------------------------------- jika insert baru ---------------------------------->	  
	
	  <%IF txtTable ="Stock" then
       q= "Insert into Stock (FCCODE, FCNAME, ITEMTYPE, MANUFACTUREPARTNUMBER, MANUFACTURECODE,ValidUntil, SUPPLIERPARTNUMBER, SUPPLIERCODE, UNITOFMEASUREMENT, STOCKGROUPCODE, FCENTRY, FCEDIT, FCIP, FCBA,ISLESSZERO, LASTUPDATE, LASTTIME, ACTIVEMATERIAL)" & _
		  "values ('" &  trim(Request("FCCODE")) & "','" & trim(Request("FCNAME")) & "','"  &  Request("ITEMTYPE") & "','" & _
		  Request("MANUFACTUREPARTNUMBER") & "','" & Request("MANUFACTURECODE") & "',to_date('" & Request("VALIDUNTIL") & "','mm/dd/yyyy'),'" & _
		  Request("SUPPLIERPARTNUMBER") & "','" &  Request("SUPPLIERCODE") & "','" &  Request("UNITOFMEASUREMENT") & "','" & _
		  Request("STOCKGROUPCODE")  & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', '" & Request("IsLessZero") & "', sysdate, To_Char(sysdate,'HH24:MI'), '" & Request("ACTIVEMATERIAL") & "' )" 		  
	  ElseIf txtTable ="BankStock" then
       q = " INSERT INTO bankstock ( fccode, fcname, itemtype, unitofmeasurement, stockgroupcode, activematerial, manufacturepartnumber, manufacturecode, supplierpartnumber, suppliercode, islesszero, validuntil, isactive, min_value, max_value, stockminimum, isbudget_control, fcentry, fcedit, fcip, lastupdate, lasttime, stocksubgroupcode)" & _
		   " VALUES ( '" & trim(Request("FCCODE")) & "'" & _
		   "          , '" & trim(Request("FCNAME")) & "'" & _
		   "          , '" &  Request("ITEMTYPE") & "'" & _
		   "          , '" &  Request("UNITOFMEASUREMENT") & "'" & _
		   "          , '" & Request("STOCKGROUPCODE")  & "'" & _
		   "          , '" & Request("ACTIVEMATERIAL") & "'" & _
		   "          , '" & Request("MANUFACTUREPARTNUMBER") & "'" & _
		   "          , '" & Request("MANUFACTURECODE") & "'" & _
		   "          , '" & Request("SUPPLIERPARTNUMBER") & "'" & _
		   "          , '" &  Request("SUPPLIERCODE") & "'" & _
		   "          , '" & Request("IsLessZero") & "'" & _
		   "          , TO_DATE('" & Request("VALIDUNTIL") & "','mm/dd/yyyy')" & _
		   "          , 'Y' " & _
		   "          , '" & Request("min_value") & "'" & _
		   "          , '" & Request("max_value") & "'" & _
		   "          , '" & Request("stockminimum") & "'" & _
		   "          , '" & Request("isbudget_control") & "'" & _
		   "          , '" & Session("CodeName") & "'" & _
		   "          , '" & Session("CodeName") & "'" & _
		   "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
		   "          , SYSDATE" & _
		   "          , TO_CHAR(SYSDATE,'hh24:mi')" & _
		   "          , '" & trim(Request("STOCKSUBGROUPCODE")) & "'" & _
		   "           )"
		   
	 q1 = "Insert into Stock (FCCODE, FCNAME, ITEMTYPE, MANUFACTUREPARTNUMBER, MANUFACTURECODE,ValidUntil, SUPPLIERPARTNUMBER, SUPPLIERCODE, UNITOFMEASUREMENT, STOCKGROUPCODE, FCENTRY, FCEDIT, FCIP, FCBA,ISLESSZERO, LASTUPDATE, LASTTIME, ACTIVEMATERIAL,STOCKSUBGROUPCODE,isactive)" & _
		  "SELECT '" &  trim(Request("FCCODE")) & "' FCCODE,'" & trim(Request("FCNAME")) & "' FCNAME,'"  &  Request("ITEMTYPE") & "' ITEMTYPE,'" & _
		  Request("MANUFACTUREPARTNUMBER") & "' MANUFACTUREPARTNUMBER,'" & Request("MANUFACTURECODE") & "' MANUFACTURECODE,to_date('" & Request("VALIDUNTIL") & "','mm/dd/yyyy') ValidUntil,'" & _
		  Request("SUPPLIERPARTNUMBER") & "' SUPPLIERPARTNUMBER,'" &  Request("SUPPLIERCODE") & "' SUPPLIERCODE,'" &  Request("UNITOFMEASUREMENT") & "' UNITOFMEASUREMENT,'" & _
		  Request("STOCKGROUPCODE")  & "' STOCKGROUPCODE,'" & Session("CodeName") & "' FCENTRY,'" & Session("CodeName") & "' FCEDIT,'" & _
		  Request.ServerVariables("REMOTE_ADDR") &"' FCIP ,fccode fcba, '" & Request("IsLessZero") & "' ISLESSZERO, sysdate LASTUPDATE, To_Char(sysdate,'HH24:MI') LASTTIME, '" & Request("ACTIVEMATERIAL") & "' ACTIVEMATERIAL , '" & Request("STOCKSUBGROUPCODE") & "' STOCKSUBGROUPCODE,'Y' isactive " & _
          " FROM Businessunit " & _
          " WHERE isactive='TRUE' " 		  
		  
	  ElseIf txtTable ="Bank" then
       q= "Insert into Bank (FCCODE, FCNAME, CONTROLJOB, ACCOUNTCODE, ADDRESS, CITY, STATE, REMARKS, DATECREATED, ACTIVATION, TRIGGER_SYSTEM, BANKTYPE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("CONTROLJOB") & "','" & _
		  Request("ACCOUNTCODE") & "','" & Request("ADDRESS") & "','" & _
		  Request("CITY") & "','" &  Request("STATE") & "','" &  Request("REMARKS") & "'," & _
		  " to_date('" & Request("datecreated") & "','mm/dd/yyyy') " & ",'" & Request("ACTIVATION") & "','" & Request("TRIGGER_SYSTEM") & "','" & Request("BANKTYPE") & "','" & Session("CodeName") & "','"_
		  & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		ElseIf txtTable ="StockGroup" then
       q= "Insert into StockGroup (FCCODE, FCNAME, CONTROLJOB, ACTIVATION, GROUPTYPE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &_
		  Request("CONTROLJOB") & "','" & request("ACTIVATION") & "','" & request("GROUPTYPE") & "','" & Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		ElseIf txtTable ="StockSubGroup" then
       q= "Insert into StockSubGroup (FCCODE, FCNAME, STOCKGROUPCODE, ACTIVATION, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME, FCBA ) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &_
		  Request("STOCKGROUPCODE") & "','" & request("ACTIVATION") & "','" & Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI'),'" & Session("fcBA") & "' ) "
		ElseIf txtTable ="EmpPosition" then
       q= "Insert into empposition (POSITION, DESCRIPTION, FCES_EM, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &_
		  Request("ACTIVATION") & "','" & Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		ElseIf txtTable ="Product" then
       q= "Insert into Product (PRODUCTCODE, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("PRODUCTDESCRIPTION") & "','"  &_
		  Request("CROPCODE") & "','" & Request("UNITOFMEASUREMENT") & "','','" & Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		ElseIf txtTable ="SectionName" then
				qLineMax = " select nvl(max(ORDERDATA),0) tempMaxLine" & _
						   " from sectionname" 
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 

				tempMaxLine = tempMaxLine + 1
				txtLineNoHist = tempMaxLine
       q= "Insert into sectionname (FCCODE, FCNAME, ORDERDATA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"& tempMaxLine &"', sysdate, To_Char(sysdate,'HH24:MI') )"
	   ElseIf txtTable ="UnitOfMeasure" then
       q= "Insert into UNITOFMEASURE (UNITOFMEASURECODE, UNITOFMEASUREDESC, ISACTIVE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" & Request("ISACTIVE") & "','" & Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','-', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="VehicleGroup" then
				qLineMax = " select nvl(max(ORDERREPORT),0) tempMaxLine" & _
						   " from VehicleGroup where fcba = '"& session("fcBA") &"'" 
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 

				tempMaxLine = tempMaxLine + 1
				txtLineNoHist = tempMaxLine
				
				Dim is_WS
				if Request("VEHICLETYPECODE")="WS" then 
					is_WS="Y"
				else
					is_WS="N"
				end if
       q= "Insert into VehicleGroup (VEHICLEGROUPCODE, DESCRIPTION, ORDERREPORT, VEHICLETYPECODE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME,workshopfactor) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"& tempMaxLine &"','" & Request("VEHICLETYPECODE") & "', '" & Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & is_WS & "' )"
	  ElseIf txtTable ="OpHarvRate" then
	  			qLineMax = " select nvl(max(LINENO),0) tempMaxLine" & _
						   " from OpHarvRate where fcba = '"& session("fcBA") &"'" 
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 

				tempMaxLine = tempMaxLine + 1
				txtLineNoHist = tempMaxLine
	   	q= "Insert into opharvrate (CROPCODE, ABW_CENSUS, TERRAIN_TYPE, ABW_MIN, ABW_MAX, BASIS_QTY, BASIS_RATE, EXTRA_BASIS_RATE, BASIS_HA, LINENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		   "values ('" &  Trim(Request("FCCODE")) & "','" & Request("bjr_sensus") & "','"  &_
		  Request("soil_type") & "','" & Request("bjr_min") & "','" & Request("bjr_max") & "','" & Request("bas_qty") & "','" & Request("bas_rate") & "','" & Request("ext_bas") & "','" & Request("bas_ha") & "','"& tempMaxLine &"','" & Session("CodeName") & "','" & _

		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	   ElseIf txtTable ="FixedAsset" then
		   Dim p_assetgroupcode, p_estimatelifeyear
		   p_assetgroupcode = ""
		   p_estimatelifeyear = 0
		   queryselisih = " SELECT ass.fccode, ass.fcname, ass.unitofmeasurement, ass.assetgroupcode, asg.fcname assetgroupname, asg.estimatelifeyear, asg.fcba" & _
						  " FROM asset ass" & _
						  " INNER JOIN assetgroup asg ON (asg.fcba = ass.fcba AND asg.fccode = ass.assetgroupcode)" & _
						  " WHERE ass.fcba = '" & Session("fcBA") & "' AND ass.fccode = '" & Request("assetcode") & "'"
			'response.write "queryselisih : " & queryselisih & "<br/>"
			set rsselisih = Server.CreateObject("ADODB.Recordset")
			rsselisih.Open queryselisih, connstr
			if not rsselisih.eof then
				p_assetgroupcode = rsselisih("assetgroupcode")
				p_estimatelifeyear = cdbl(to_nol(rsselisih("estimatelifeyear")))
			else
				p_assetgroupcode = ""
				p_estimatelifeyear = 0
			end if
			rsselisih.close() : set rsselisih = nothing
	   
       q= " INSERT INTO fixedasset_hdr_tab (fccode, fcname, assetcode, assetgroupcode, estimatelifeyear, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
		  " VALUES ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" &  Request("assetcode") & "','"&p_assetgroupcode&"','"&p_estimatelifeyear&"','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	   ElseIf txtTable ="Asset" then
       q= "Insert into Asset (FCCODE, FCNAME, UNITOFMEASUREMENT, ASSETGROUPCODE, FCENTRY, FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" &  Request("UNITOFMEASUREMENT") & "','" & Request("ASSETGROUPCODE")  & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	   ElseIf txtTable ="AssetGroup" then
       q= "Insert into AssetGroup (FCCODE, FCNAME, estimatelifeyear, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &_
		  Request("estimatelifeyear") & "','" & Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Supplier" then 
		<!-- Modify Aam Jan 21-->
		 q= "Insert into Supplier (FCCODE, FCNAME, CONTROLJOB, DATECREATED, ACTIVATION, ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT, BANK_OWNERSHIP, PHONE, FAX, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, CPERSON, MAIL, NPWP,FCTYPE) " & _
		  "SELECT '" &  trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("CONTROLJOB") & "'," & _
		  " to_date('" & Request("datecreated") & "','mm/dd/yyyy') " & ",'Y','" & _
		  Request("ADDRESS") & "','" &  Request("CITY") & "','" &  Request("STATE") & "','" & _
		  Request("REMARKS")  & "','" & Request("BANKNAME") & "','" & Request("BANKACCOUNT") & "','" & Request("BANK_OWNERSHIP") & "','" & Request("PHONE") & "','" & Request("FAX") & "','" & _
		  Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"',businessunit.fccode, sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("CPERSON") & "','" & Request("MAIL") & "','" & Request("NPWP") & "','" & Request("FCTYPE") & "' " & _ 
		  "FROM businessunit where isactive='TRUE'" 
        
		'q= "Insert into Supplier (FCCODE, FCNAME, CONTROLJOB, DATECREATED, ACTIVATION, ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT, BANK_OWNERSHIP, PHONE, FAX, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, CPERSON, MAIL, NPWP) " & _
		'  "values ('" &  Request("FCCODE") & "','" & Request("FCNAME") & "','"  &  Request("CONTROLJOB") & "'," & _
		'  " to_date('" & Request("datecreated") & "','mm/dd/yyyy') " & ",'" & Request("ACTIVATION") & "','" & _
		'  Request("ADDRESS") & "','" &  Request("CITY") & "','" &  Request("STATE") & "','" & _
		'  Request("REMARKS")  & "','" & Request("BANKNAME") & "','" & Request("BANKACCOUNT") & "','" & Request("BANK_OWNERSHIP") & "','" & Request("PHONE") & "','" & Request("FAX") & "','" & _
		'  Session("CodeName") & "','" & _
		'  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("CPERSON") & "','" & Request("MAIL") & "','" & Request("NPWP") & "')"
        
	   ElseIf txtTable ="Transporter" then
       q= "Insert into transporter (FCCODE, FCNAME, CONTROLJOB, DATECREATED, ACTIVATION, ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("CONTROLJOB") & "'," & _
		  " to_date('" & Request("datecreated") & "','mm/dd/yyyy') " & ",'" & Request("ACTIVATION") & "','" & _
		  Request("ADDRESS") & "','" &  Request("CITY") & "','" &  Request("STATE") & "','" & _
		  Request("REMARKS")  & "','" & Request("BANKNAME") & "','" & Request("BANKACCOUNT") & "','" & _
		  Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	   ElseIf txtTable ="CUSTOMER" then
       q= "Insert into CUSTOMER (CUSTOMERCODE,DESCRIPTION,SUPPLIER,FCENTRY,FCEDIT,FCIP,FCBA,OWN_ESTATE,LASTUPDATE,LASTTIME,CUSTOMER_FCBA,CUSTFFBTYPE,ISTRANSPORTER,TELEPHONE,ISACTIVE,CONTROLJOB) " & _
		  "values ('" &  Trim(Request("CUSTOMERCODE")) & "','" & Request("DESCRIPTION") & "','"  &  Request("SUPPLIER") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Session("fcBA") & "','" & Request("OWN_ESTATE") & "',sysdate,To_Char(sysdate,'HH24:MI'),'" & _
		  Session("fcBA") & "','" & Request("CUSTFFBTYPE") & "','" &  Request("ISTRANSPORTER") & "','"  &  Request("TELEPHONE") & "','Y','2110102')"
	   ElseIf txtTable ="MultiLevel_Harvesting_Rate" then
       q= "Insert into MultiLevel_Harvesting_Rate (LEVELNO, CROPCODE, DIVISION, PLANTING_YEAR, TERRAIN_TYPE, MIN_BUNCH, MAX_BUNCH, RATE, RATE_HOLIDAY, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("CROPCODE") & "','" & Request("DIVISION") & "','" & Request("PLANTINGYEAR") & "','" & Request("TERRAINTYPE") & "','" & _
		  Request("MIN_BUNCH") & "','" & Request("MAX_BUNCH") & "','" & Request("RATE") & "','" & _
		  Request("RATE_HOLIDAY") & "','" & _
		  Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	   ElseIf txtTable ="Store" then
       q= "Insert into Store (FCCODE, FCNAME, REMARKS, ACTIVATION, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("REMARKS") & "','" & _
		  Request("ACTIVATION") & "','" & _
		  Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Workshop" then
       q= "Insert into Workshop (FCCODE, FCNAME, REMARKS, ACTIVATION, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("REMARKS") & "','" & _
		  Request("ACTIVATION") & "','" & _
		  Session("CodeName") & "','" & _
		  Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Vehicle" then
       q= "Insert into Vehicle (FCCODE, FCNAME, VEHICLEGROUPCODE,VEHICLESUBGROUPCODE, DATECREATED, ACTIVATION,DATETERMINATE, REGISTRATIONNO, MAKE, MODEL, ENGINENO, CHASISNO, YEAROFMADE, YEAROFPURCHASE, OWNERSHIP, VEHICLESTATUS, CONTROLALLOCATION, CONTROLJOB, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, WBREGISTERED)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("VEHICLEGROUPCODE") & "','"  &  Request("VEHICLESUBGROUPCODE") & "'," & _
		  " to_date('" & Request("datecreated") & "','mm/dd/yyyy') " & ",'" & Request("ACTIVATION") & "',to_date('" & Request("DATETERMINATE") & "','mm/dd/yyyy'),'" & _
		  Request("REGISTRATIONNO") & "','" &  Request("MAKE") & "','" &  Request("MODEL") & "','" & _
		  Request("ENGINENO")  & "','" & Request("CHASISNO")  & "'," & _
		  " to_date('" & Request("yearofmade") & "','mm/dd/yyyy') " &"," & _
		  " to_date('" & Request("yearofpurchase") & "','mm/dd/yyyy') " & ",'" & Request("OWNERSHIP") & "'" & _
		  ",'" & Request("VEHICLESTATUS")  & "','" & _
		  Request("CONTROLALLOCATION")  & "','" & Request("CONTROLJOB") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("WBREGISTERED") & "' )"
	  ElseIf txtTable ="Contractor" then
       q= "Insert into Contractor (FCCODE, FCNAME, CONTROLJOB, DATECREATED, ACTIVATION, ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "SELECT '" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("CONTROLJOB") & "'," & _
		  " to_date('" & Request("DATECREATED") & "','mm/dd/yyyy') " & ",'" & Request("ACTIVATION") & "','" & _
		  Request("ADDRESS") & "','" &  Request("CITY") & "','" &  Request("STATE") & "','" & _
		  Request("REMARKS")  & "','" & Request("BANKNAME")  & _
		  "','" & Request("BANKACCOUNT")  & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"', '" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') " & _ 
		  "FROM DUAL " 
	  ElseIf txtTable ="Building" then
       q= "Insert into Building (FCCODE, FCNAME, BUILDGROUPCODE_DESC, ITEM, ACTIVATION, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("BUILDGROUPCODE_DESC") & "','" & _
		  Request("ITEM") & _
		  "','" & Request("ACTIVATION")  & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="MillQuality" then
       q= "Insert into MillQuality (FCCODE, FCNAME, DIRT, FFA, MOISTURE, DOBI, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("DIRT") & "','" & _
		  Request("FFA") & _
		  "','" & Request("MOISTURE")  & "','" & Request("DOBI")  & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="UserProfile" then
       q= "Insert into UserProfile (FCCODE, FCPASSWORD, FULLNAME, DEPARTMENT, ACCESSRIGHTS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("fcPassword") & "','" & Request("fullname") & "','" & Request("department") & "','" & Request("ACCESSRIGHTS") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Gang" then
       q= "Insert into Gang (FCCODE, FCNAME, REMARKS, ACTIVATION, GANGTYPECODE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("fcName") & "','" & Request("Remarks") & "','" & Request("Activation") & "','"  & _
		  Request("GangTypeCode") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Station" then
       q= "Insert into Station (FCCODE, FCNAME, JOBCODE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("fcName") & "','" & Request("JOBCODE") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Machine" then
       q= "Insert into Machine (FCCODE, DESCRIPTION, ACTIVATION, FIXEDASSETCODE, OWNERSHIP, DATEPURCHASE, PURCHASECOST, STATION_STATIONCODE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("DESCRIPTION") & "','" & Request("ACTIVATION") & "','"  &  Request("FIXEDASSETCODE") & "','" &  Request("OWNERSHIP") & "'," & _
		  " to_date('" & Request("DATEPURCHASE") & "','mm/dd/yyyy') " & ",'" & Request("PURCHASECOST") & "','" & _
		  Request("STATION_STATIONCODE") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
      ElseIf txtTable ="WorkingTime" then
		q= "Insert into WorkingTime (FCCODE, FCNAME, TIMEIN, TIMEOUT, ISRESTIME, ISNEXTDAY, ADDIN_MINUTES, ADDOUT_MINUTES, ACTIVATION, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" & Request("TIMEIN") & "','"  &  Request("TIMEOUT") & "','" &  Request("ISRESTIME") & "','" & _
		   Request("ISNEXTDAY") & "','" & Request("ADDIN") & "','" & Request("ADDOUT") & "','" & Request("Activation") & "','" & _ 
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		  
	 ElseIf txtTable ="WorkingTime_Template" then
		q= "Insert into WorkingTime_Template (FCCODE, FCNAME, GANGCODE,D_2,D_3,D_4,D_5,D_6,D_7,D_1,PERIOD,ACTIVATION, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" & Request("GANGCODE") & "','"  &  _ 
		   Request("D_2") & "','" &  Request("D_3") & "','" & Request("D_4") & "','" &  Request("D_5") & "','" & Request("D_6") & "','" &  Request("D_7") & "','"  &  Request("D_1") & "','" & _ 
		   Request("PERIOD") & "','" & Request("Activation") & "','" & _ 
		   Session("CodeName") & "','" & Session("CodeName") & "','" & _
		   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		
		'response.write (q)
		  
	  ElseIf txtTable ="Field" then
       q= "insert into Field (FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, " & _ 
	       "PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, DX_BLOCK_STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, " & _ 
	       "COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE, " & _ 
		   "BUIST, OWLNEST, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, HARVESTINGSECTION, " & _
		   "FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05, ISACTIVE_HARVESTED,VALID_FROM, " & _
		   "AREAL_LC,CADANGAN,OKUPASI,BIBITAN,PABRIK,EMPLASMEN,PARIT,JALAN,PRASARANA_LAINNYA,QUARRY,TANGGUL,ENCLAVE,TAMBANG,KONSERVASI,BUKIT_SUNGAI,NO_IMAGERY,LUAS_GROSS,HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER, PLANTINGPOINT, FULLSTAND)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("DIVISION") & "','" & _
		  Request("HECTARAGEPLANTED") & "','" & Request("TERRAINTYPE") & "','"  &  Request("SOILTYPE") & "','" &  Request("OWNERSHIP") & "','" &  Request("ACTIVATION") & "','" &  Request("CROPTYPE") & "'," & _
		  " to_date('" & Request("PLANTINGDATE") & "','mm/dd/yyyy'),'" & Request("PLANTINGMATERIAL") & "','" & _
		  Request("TOTALSTAND") & "','" & _
		  Request("STATUS")  & "','" & Request("PREVSEMESTER_ABW") & "','" & Request("PRESSEMESTER_ABW") & "','" & _
		  Request("HARVESTINGBASED_ABW")  & "','" & Request("MAIN_ROAD") & "','" & Request("COLLECTIONROAD") & "','" & _
		  Request("MAINDRAIN")  & "','" & Request("COLLECTIONDRAIN") & "','" & Request("SUBSIDIARYDRAIN") & "','" & _
		  Request("COLLECTIONPOINT")  & "','" & Request("HARVESTINGBRIDGE") & "','" & Request("PERMANENTBRIDGE") & "','" & _
		  Request("WOODBRIDGE")  & "','" & Request("BUIST") & "','" & Request("OWLNEST") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'), '" &_
		  Request("HARVESTINGSECTION") & "','" & Request("FIELD_INFO_01")  & "','" & Request("FIELD_INFO_02") & "','" &_ 
		  Request("FIELD_INFO_03") & "','" & Request("FIELD_INFO_04")  & "','" & Request("FIELD_INFO_05") & "','" &_
		  Request("ISACTIVE_HARVESTED") & "',to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy'),'"& _
		  Request("AREAL_LC")  & "','" & Request("CADANGAN") & "','" & Request("OKUPASI") &"','" & _
		  Request("BIBITAN")  & "','" & Request("PABRIK") &"','" & _
		  Request("EMPLASMEN")  & "','" & Request("PARIT") & "','" & Request("JALAN") & "','" & _
		  Request("PRASARANA_LAINNYA")  & "','" & Request("QUARRY") & "','" & Request("TANGGUL") & "','" & _
		  Request("ENCLAVE")  & "','" & Request("TAMBANG") & "','" & Request("KONSERVASI") & "','" & _
		  Request("BUKIT_SUNGAI")  & "','" & Request("NO_IMAGERY") & "','" & Request("LUAS_GROSS") & "','" & Request("hectarageharvested") & "','" & Request("productivestand") & "','" & Request("harvestingsectionnumber") & "','" & Request("plantingpoint") & "','" & Request("fullstand") & "')"
		
		<!-- Additional New Insert FieldDetail-->
		q_his = "insert into FIELDDETAIL (FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, " & _ 
	       "PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, DX_BLOCK_STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, " & _ 
	       "COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE, " & _ 
		   "BUIST, OWLNEST, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, HARVESTINGSECTION, " & _
		   "FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05, ISACTIVE_HARVESTED,VALID_FROM,LINENO,ACTION_STATE, " & _
		   "AREAL_LC,CADANGAN,OKUPASI,BIBITAN,PABRIK,EMPLASMEN,PARIT,JALAN,PRASARANA_LAINNYA,QUARRY,TANGGUL,ENCLAVE,TAMBANG,KONSERVASI,BUKIT_SUNGAI,NO_IMAGERY,LUAS_GROSS,HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTIONNUMBER,PLANTINGPOINT, FULLSTAND ) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("DIVISION") & "','" & _
		  Request("HECTARAGEPLANTED") & "','" & Request("TERRAINTYPE") & "','"  &  Request("SOILTYPE") & "','" &  Request("OWNERSHIP") & "','" &  Request("ACTIVATION") & "','" &  Request("CROPTYPE") & "'," & _
		  " to_date('" & Request("PLANTINGDATE") & "','mm/dd/yyyy'),'" & Request("PLANTINGMATERIAL") & "','" & _
		  Request("TOTALSTAND") & "','" & _
		  Request("STATUS")  & "','" & Request("PREVSEMESTER_ABW") & "','" & Request("PRESSEMESTER_ABW") & "','" & _
		  Request("HARVESTINGBASED_ABW")  & "','" & Request("MAIN_ROAD") & "','" & Request("COLLECTIONROAD") & "','" & _
		  Request("MAINDRAIN")  & "','" & Request("COLLECTIONDRAIN") & "','" & Request("SUBSIDIARYDRAIN") & "','" & _
		  Request("COLLECTIONPOINT")  & "','" & Request("HARVESTINGBRIDGE") & "','" & Request("PERMANENTBRIDGE") & "','" & _
		  Request("WOODBRIDGE")  & "','" & Request("BUIST") & "','" & Request("OWLNEST") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'), '" &_
		  Request("HARVESTINGSECTION") & "','" & Request("FIELD_INFO_01")  & "','" & Request("FIELD_INFO_02") & "','" &_ 
		  Request("FIELD_INFO_03") & "','" & Request("FIELD_INFO_04")  & "','" & Request("FIELD_INFO_05") & "','" &_
		  Request("ISACTIVE_HARVESTED") & "',to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy'),'1','Inserted','"& _
		  Request("AREAL_LC")  & "','" & Request("CADANGAN") & "','" & Request("OKUPASI") &"','" & _
		  Request("BIBITAN")  & "','" & Request("PABRIK") &"','" & _
		  Request("EMPLASMEN")  & "','" & Request("PARIT") & "','" & Request("JALAN") & "','" & _
		  Request("PRASARANA_LAINNYA")  & "','" & Request("QUARRY") & "','" & Request("TANGGUL") & "','" & _
		  Request("ENCLAVE")  & "','" & Request("TAMBANG") & "','" & Request("KONSERVASI") & "','" & _
		  Request("BUKIT_SUNGAI")  & "','" & Request("NO_IMAGERY") & "','" & Request("LUAS_GROSS") & "','" & Request("hectarageharvested") & "','" & Request("productivestand") & "','" & Request("harvestingsectionnumber") & "','" & Request("plantingpoint") & "','" & Request("fullstand") & "')"
		  
		  'RESPONSE.Write(q_his)
		  ' FIELD_HIST_TAB
		  for i= Cint(Trim(Request("MONTH")))  to 12 
          q2  = " MERGE INTO FIELD_HIST_TAB  a " & _ 
				" USING ( SELECT " & i &  " MONTH, " & Trim(Request("YEAR")) & " YEAR ,'" &  Trim(Request("FCCODE")) & "' FCCODE,'" & Request("FCNAME") & "' FCNAME,'"  &  Request("DIVISION") & "' DIVISION,'" & _
				  Request("HECTARAGEPLANTED") & "' HECTARAGEPLANTED,'" & Request("TERRAINTYPE") & "' TERRAINTYPE,'"  &  Request("SOILTYPE") & "' SOILTYPE,'" &  Request("OWNERSHIP") & "' OWNERSHIP,'" &  Request("ACTIVATION") & "' ACTIVATION,'" &  Request("CROPTYPE") & "' CROPTYPE," & _
				  " to_date('" & Request("PLANTINGDATE") & "','mm/dd/yyyy') PLANTINGDATE,'" & Request("PLANTINGMATERIAL") & "' PLANTINGMATERIAL,'" & _
				  Request("TOTALSTAND") & "' TOTALSTAND,'" & _
				  Request("STATUS")  & "' DX_BLOCK_STATUS,'" & Request("PREVSEMESTER_ABW") & "' PREVSEMESTER_ABW,'" & Request("PRESSEMESTER_ABW") & "' PRESSEMESTER_ABW,'" & _
				  Request("HARVESTINGBASED_ABW")  & "' HARVESTINGBASED_ABW,'" & Request("MAIN_ROAD") & "' MAIN_ROAD,'" & Request("COLLECTIONROAD") & "' COLLECTIONROAD,'" & _
				  Request("MAINDRAIN")  & "' MAINDRAIN,'" & Request("COLLECTIONDRAIN") & "' COLLECTIONDRAIN,'" & Request("SUBSIDIARYDRAIN") & "' SUBSIDIARYDRAIN,'" & _
				  Request("COLLECTIONPOINT")  & "' COLLECTIONPOINT,'" & Request("HARVESTINGBRIDGE") & "' HARVESTINGBRIDGE,'" & Request("PERMANENTBRIDGE") & "' PERMANENTBRIDGE,'" & _
				  Request("WOODBRIDGE")  & "' WOODBRIDGE,'" & Request("BUIST") & "' BUIST,'" & Request("OWLNEST") & "' OWLNEST,'" & _
				  Session("fcBA") & "' FCBA, '" &_
				  Request("HARVESTINGSECTION") & "' HARVESTINGSECTION,'" & Request("FIELD_INFO_01")  & "' FIELD_INFO_01,'" & Request("FIELD_INFO_02") & "' FIELD_INFO_02,'" &_ 
				  Request("FIELD_INFO_03") & "' FIELD_INFO_03,'" & Request("FIELD_INFO_04")  & "' FIELD_INFO_04,'" & Request("FIELD_INFO_05") & "' FIELD_INFO_05,'" &_
				  Request("ISACTIVE_HARVESTED") & "' ISACTIVE_HARVESTED ,to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy') VALID_FROM,'"& _
				  Request("AREAL_LC")  & "' AREAL_LC,'" & Request("CADANGAN") & "' CADANGAN,'" & Request("OKUPASI") &"' OKUPASI,'" & _
				  Request("BIBITAN")  & "' BIBITAN,'" & Request("PABRIK") &"' PABRIK,'" & _
				  Request("EMPLASMEN")  & "' EMPLASMEN,'" & Request("PARIT") & "' PARIT,'" & Request("JALAN") & "' JALAN,'" & _
				  Request("PRASARANA_LAINNYA")  & "' PRASARANA_LAINNYA,'" & Request("QUARRY") & "' QUARRY,'" & Request("TANGGUL") & "' TANGGUL,'" & _
				  Request("ENCLAVE")  & "' ENCLAVE,'" & Request("TAMBANG") & "' TAMBANG,'" & Request("KONSERVASI") & "' KONSERVASI,'" & _
				  Request("BUKIT_SUNGAI")  & "' BUKIT_SUNGAI,'" & Request("NO_IMAGERY") & "' NO_IMAGERY,'" & Request("LUAS_GROSS") & "' LUAS_GROSS,'"  & _ 
				  Request("hectarageharvested") & "' hectarageharvested,'" & Request("productivestand") & "' productivestand,'" & _  
				  Request("harvestingsectionnumber") & "' harvestingsectionnumber,'" & Request("plantingpoint") & "' plantingpoint,'" & _
				  Request("fullstand") & "' fullstand,sysdate,To_Char(sysdate,'HH24:MI') LASTTIME " & _
				" FROM DUAL ) b  ON (a.MONTH=b.MONTH AND a.YEAR=b.YEAR AND a.FCCODE=b.FCCODE AND a.FCBA=b.FCBA ) " & _ 
				" WHEN NOT MATCHED THEN INSERT ( " & _
				"  a.MONTH,a.YEAR,a.FCCODE,a.FCNAME,a.DIVISION,a.HECTARAGEPLANTED,a.TERRAINTYPE" & _
				" , a.SOILTYPE, a.OWNERSHIP, a.ACTIVATION, a.CROPTYPE, a.PLANTINGDATE" & _
				" , a.PLANTINGMATERIAL, a.TOTALSTAND, a.DX_BLOCK_STATUS, a.PREVSEMESTER_ABW, a.PRESSEMESTER_ABW" & _
				" , a.HARVESTINGBASED_ABW, a.MAIN_ROAD, a.COLLECTIONROAD, a.MAINDRAIN, a.COLLECTIONDRAIN" & _
				" , a.SUBSIDIARYDRAIN, a.COLLECTIONPOINT, a.HARVESTINGBRIDGE, a.PERMANENTBRIDGE, a.WOODBRIDGE" & _
				" , a.BUIST, a.OWLNEST" & _
				" , a.FIELD_INFO_01, a.FIELD_INFO_02, a.FIELD_INFO_03, a.FIELD_INFO_04, a.FIELD_INFO_05,a.ISACTIVE_HARVESTED, a.VALID_FROM,a.LINENO" & _
				" , a.AREAL_LC, a.CADANGAN, a.OKUPASI, a.BIBITAN, a.PABRIK, a.EMPLASMEN, a.PARIT, a.JALAN, a.PRASARANA_LAINNYA" & _
				" , a.QUARRY, a.TANGGUL, a.ENCLAVE, a.TAMBANG, a.KONSERVASI, a.BUKIT_SUNGAI, a.NO_IMAGERY, a.LUAS_GROSS, a.HECTARAGEHARVESTED, a.PRODUCTIVESTAND, a.HARVESTINGSECTION, a.HARVESTINGSECTIONNUMBER, a.PLANTINGPOINT, a.FULLSTAND " & _
				" , a.FCENTRY, a.FCEDIT, a.FCIP, a.FCBA, a.lastupdate, a.lasttime) " & _ 
				" VALUES( b.MONTH,b.YEAR,b.FCCODE,b.FCNAME,b.DIVISION,b.HECTARAGEPLANTED,b.TERRAINTYPE" & _
				" , b.SOILTYPE, b.OWNERSHIP, b.ACTIVATION, b.CROPTYPE, b.PLANTINGDATE" & _
				" , b.PLANTINGMATERIAL, b.TOTALSTAND, b.DX_BLOCK_STATUS, b.PREVSEMESTER_ABW, b.PRESSEMESTER_ABW" & _
				" , b.HARVESTINGBASED_ABW, b.MAIN_ROAD, b.COLLECTIONROAD, b.MAINDRAIN, b.COLLECTIONDRAIN" & _
				" , b.SUBSIDIARYDRAIN, b.COLLECTIONPOINT, b.HARVESTINGBRIDGE, b.PERMANENTBRIDGE, b.WOODBRIDGE" & _
				" , b.BUIST, b.OWLNEST" & _
				" , b.FIELD_INFO_01, b.FIELD_INFO_02, b.FIELD_INFO_03, b.FIELD_INFO_04, b.FIELD_INFO_05,b.ISACTIVE_HARVESTED, b.VALID_FROM,'" & i & "'" & _
				" , b.AREAL_LC, b.CADANGAN, b.OKUPASI, b.BIBITAN, b.PABRIK, b.EMPLASMEN, b.PARIT, b.JALAN, b.PRASARANA_LAINNYA" & _
				" , b.QUARRY, b.TANGGUL, b.ENCLAVE, b.TAMBANG, b.KONSERVASI, b.BUKIT_SUNGAI, b.NO_IMAGERY, b.LUAS_GROSS, b.HECTARAGEHARVESTED, b.PRODUCTIVESTAND, b.HARVESTINGSECTION, b.HARVESTINGSECTIONNUMBER, b.PLANTINGPOINT, b.FULLSTAND " & _
				" ,'" & Session("CodeName") & "' ,'" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', b.FCBA, sysdate,To_Char(sysdate,'HH24:MI')  )"
		 '   response.Write(q2)
				'connStr.execute(q2)
				ExecuteAndLog ScriptName, q2				
				
            next 				
		 
		  'response.Flush()
	  ElseIf txtTable ="Employee" then
	        'Cek id Exist
				dim rsEC, txtEC, stringEC
				stringEC =" SELECT 1 FROM " & txtTable & "" & _
				" WHERE FCCODE = '" & Request("EMPLOYEECODE") & "' AND FCBA = '" & Session("FCBA") & "'"
				Set rsEC = Server.CreateObject("ADODB.Recordset")
				  rsEC.Open stringEC , connstr
				  IF Not rsEC.eof then
					     txtEC = rsEC("FCCODE")						 
				  End if
				rsEC.Close
				set rsEC = Nothing  	
				
			if txtEC = Trim(Request("FCCODE"))  then
             response.write("<script language=""javascript"">alert('NIK have entered already exists!'); history.go(-1);</script>")
			 response.write("<script language=""javascript"">history.go(-1);</script>")
		
			else
			
				Dim queryselisih, rsselisih, CurrentAgeofJoin
				queryselisih = " select to_number(to_char(to_date('" & Request("DATEOFJOIN") & "','mm/dd/yyyy'), 'yyyy')) - to_number(to_char(to_date('" & Request("DATEOFBIRTH") & "','mm/dd/yyyy'), 'yyyy')) CurrentAgeofJoin" & _
							   " from dual"
				'response.write "queryselisih : " & queryselisih & "<br/>"
				set rsselisih = Server.CreateObject("ADODB.Recordset")
				rsselisih.Open queryselisih, connstr
				if not rsselisih.eof then
					CurrentAgeofJoin = cdbl(rsselisih("CurrentAgeofJoin"))
				else
					CurrentAgeofJoin = 0
				end if
				rsselisih.close()
				rsselisih=null
				'response.write "CurrentAgeofJoin : " & CurrentAgeofJoin & "<br/>"
			
			   q= "Insert into Employee (FCCODE, FCNAME, SECTIONNAME, GANGCODE, DATEOFBIRTH, RELIGION, RACE, IDENTITYCARD, HIGHESTEDUCATION, " & _
				   "YEAROFGRADUATION, DATEOFJOIN, DATEOFDESIGNATION, PRESENTGRADE, PRESENTGRADE_DETAIL, PRESENTGRADE2, PRESENTGRADE_DETAIL2, PREVIOUSGRADE2, PREVIOUSGRADE_DETAIL2," & _ 
				   "PRESENTRATE, PREVIOUSGRADE, " & _ 
				   "PREVIOUSGRADE_DETAIL, PREVIOUSRATE, EMP_TYPE, MARITALSTATUS, SPOUSE, DEPENDENTCHILD, NPWP, JAMSOSTEKNO, JAMSOSTEKCATEGORY, POSITION, " & _ 
				   "DATETERMINATE, FCENTRY, FCEDIT, FCIP, FCBA, NATIONALITY, INCLUDEDTAXCALCULATION, TAXSTATUS, NATURETYPE,GENDER,PAYMENTTYPE,BANKNAME,ACCOUNTNO,OWNERSHIPOFACCOUNT,  TAXPAIDBYCOMPANY, ADDRESS, BPJSKES_CATEGORY, BPJSKES_NO, LASTUPDATE, LASTTIME, company_nik, subgrade ) " & _
				  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("SECTIONNAME") & "','" & _
				  Request("GANGCODE") & "', to_date('" & Request("DATEOFBIRTH") & "','mm/dd/yyyy') " & ",'" & Request("RELIGION") & "','" & _
				  Request("RACE") & "','" & Request("IDENTITYCARD") & "','" & Request("HIGHESTEDUCATION") & "'," & _
				  " '" & Request("YEAROFGRADUATION") & "'," & _
				  " to_date('" & Request("DATEOFJOIN") & "','mm/dd/yyyy') " & "," & _
				  " to_date('" & Request("DATEOFDESIGNATION") & "','mm/dd/yyyy') " & ",'" & _
				  Request("PRESENTGRADE") & "','" & Request("PRESENTGRADE_DETAIL") & "','" & _
				  Request("PRESENTGRADE2") & "','" & Request("PRESENTGRADE_DETAIL2") & "','" & _
				  Request("PREVIOUSGRADE2") & "','" & Request("PREVIOUSGRADE_DETAIL2") & "','" & _
				  Request("PRESENTRATE")  & "','" & Request("PREVIOUSGRADE") & "','" & _
				  Request("PREVIOUSGRADE_DETAIL")  & "','" & Request("PREVIOUSRATE") & "','" & Request("EMP_TYPE") & "','" & Request("MARITALSTATUS") & "','" & _
				  Request("SPOUSE")  & "','" & Request("DEPENDENTCHILD") & "','" & Request("NPWP") & "','" & _
				  Request("JAMSOSTEKNO")  & "','" & Request("JAMSOSTEKCATEGORY") & "','" & Request("POSITION") & "'," & _
				  " to_date('" & Request("DATETERMINATE") & "','mm/dd/yyyy') " & ",'" & _
				  Session("CodeName") & "','" & Session("CodeName") & "','" & _
				  Request.ServerVariables("REMOTE_ADDR") & "','" & Session("fcBA") & "', '" & Request("NATIONALITY")  & "'" & _
				  " , '" & Request("INCLUDEDTAXCALCULATION") & "'" & _
				  " , '" & Request("TAXSTATUS") & "'" & _
				  " , '" & Request("NATURETYPE") & "','" & Request("GENDER") & "','" & Request("PAYMENTTYPE") & "','" & Request("BANKNAME") & "','" & Request("ACCOUNTNO") & "','" & Request("OWNERSHIPOFACCOUNT") & "', '" & Request("TAXPAIDBYCOMPANY") & "', '" & Request("txtaddress") & "', '" & Request("txtbpjs_category") & "', '" & Request("BPJSKES_NO") & "', sysdate, To_Char(sysdate,'HH24:MI'), '" & Request("company_nik") & "', '" & Request("subgrade") & "' )"
			end if  
	  
	  ElseIf txtTable ="RangeTaxGroup" then
		q= "Insert into rangetaxgroup (FCCODE, FCNAME, RANGEFROM, RANGEUNTIL, TARIFFPERCENTAGE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
		   "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','"  &  Request("RANGEFROM") & "','" & _
		   Request("RANGEUNTIL") & "','" & Request("TARIFFPERCENTAGE") & "','" & Request("REMARKS") & "','" & _
		   Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
		  
	  ElseIf txtTable ="MaritalStatus" then
		q= "Insert into maritalstatus (FCCODE, FCNAME, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME ) " & _
		   "values ('" &  Trim(Request("FCCODE")) & "','" & Request("FCNAME") & "','" & _
		   Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
		
	  ElseIf txtTable ="UntaxableIncome" then
		q= "Insert into untaxableincome (MARITALSTATUSCODE, FCNAME, UTI_EMPLOYEE, UTI_SPOUSE, UTI_CHILD, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
		   "values ('" &  Request("MARITALSTATUSCODE") & "','" & Request("FCNAME") & "','"  &  Request("UTI_EMPLOYEE") & "','" & _
		   Request("UTI_SPOUSE") & "','" & Request("UTI_CHILD") & "','" & Request("REMARKS") & "','" & _
		   Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"

	  ElseIf txtTable ="EmployeeAllowdedDetail" then 
	  	dim rsEAD, txtCODE, stringEAD, txtAllowCODE
	  stringEAD =" SELECT" & _
				" EMPLOYEECODE, ALLOWDEDCODE FROM " & txtTable & "" & _
				" WHERE EMPLOYEECODE = '" & Request("EMPLOYEECODE") & "' AND ALLOWDEDCODE = '" & Request("ALLOWDEDCODE") & "' AND FCBA = '" & Session("FCBA") & "'"
      Set rsEAD = Server.CreateObject("ADODB.Recordset")
				  rsEAD.Open stringEAD , connstr
				  IF Not rsEAD.eof then
					     txtCODE = rsEAD("EMPLOYEECODE")
						 txtAllowCODE = rsEAD("ALLOWDEDCODE")
					  End if
					  rsEAD.Close
					  set rsEAD = Nothing  	
		if txtCODE = Request("EMPLOYEECODE") and txtAllowCODE = Request("ALLOWDEDCODE") then
             response.write("<script language=""javascript"">alert('Id code have entered already exists!'); history.go(-1);</script>")
			 response.write("<script language=""javascript"">history.go(-1);</script>")
			else
			
			q= "Insert into employeeallowdeddetail (EMPLOYEECODE, ALLOWDEDCODE, RATE_1, RATE_2, RATE_3, RATE_4, RATE_5, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
		  "values ('" &  Request("EMPLOYEECODE") & "','" & Request("ALLOWDEDCODE") & "','"  &  Request("RATE_1") & "','" & _
		  Request("RATE_2") & "','" & Request("RATE_3") & "','" & Request("RATE_4") & "','" & Request("RATE_5") & "','" & Request("REMARKS") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
		end if
	  ElseIf txtTable ="CONTROLALLOWDED" then
       q= "Insert into CONTROLALLOWDED (FCCODE, CROPCODE, ALLOWDEDCODE, MAIN_ACTIVITY_CODE, USE_ALLOWDED_JOB, USE_ALL_GAD_JOB, " & _
	      "CHARGEALLGADLOCATION, CHARGEMAINLOCATION, CHARGINGMECHANISM, ALLOCATE, MANUAL_OR_AUTO, ACTIVESTATUS, PREMI_ALLOWANCE, " & _
		  "PENALTY_DEDUCTION, FOR_EMPLOYER_EMPLOYEE, RICE_OR_VALUE, UOM, BS_PW_OT, RPT_SEQ, MINAMT, MAXAMT, RATE1, RATE2, " & _
		  "RATE3, RATE4, RATE5, TYPE1, TYPE2, TYPE3, REFER_TABLE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("CROPCODE") & "','"  &  Request("ALLOWDEDCODE") & "','" & _
		  Request("MAIN_ACTIVITY_CODE") & "','" & Request("USE_ALLOWDED_JOB") & "','" & Request("USE_ALL_GAD_JOB") & "','" & _
		  Request("CHARGEALLGADLOCATION")  & "','" & Request("CHARGEMAINLOCATION") & "','" & Request("CHARGINGMECHANISM") & "','" & _
		  Request("ALLOCATE")  & "','" & Request("MANUAL_OR_AUTO") & "','" & Request("ACTIVESTATUS") & "','" & _
		  Request("PREMI_ALLOWANCE")  & "','" & Request("PENALTY_DEDUCTION") & "','" & Request("FOR_EMPLOYER_EMPLOYEE") & "','" & _
		  Request("RICE_OR_VALUE")  & "','" & Request("UOM") & "','" & Request("BS_PW_OT") & "','" & _
		  Request("RPT_SEQ")  & "','" & Request("MINAMT") & "','" & Request("MAXAMT") & "','" & _
		  Request("RATE1")  & "','" & Request("RATE2") & "','" & Request("RATE3") & "','" & _
		  Request("RATE4")  & "','" & Request("RATE5") & "','" & Request("TYPE1") & "','" & _
		  Request("TYPE2")  & "','" & Request("TYPE3") & "','" & Request("REFER_TABLE") & "','" & _
		  Request("REMARKS")  & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") & "','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Job" then
       q= "Insert into Job (FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, " & _
	      "UNITOFMEASUREMENT, ACCOUNTCODE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "SELECT '" &  Trim(Request("FCCODE")) & "','" & Request("fcName") & "','" & Request("JobGroupCode") & "','" & Request("JOB_FORCROPTYPE") & "','" & Request("JOB_FORFIELDSTATUS") & "','"  & _
		  Request("JOB_CATEGORY") & "','" & Request("UNITOFMEASUREMENT") & "','" & Request("ACCOUNTCODE") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"',fcCode, sysdate, To_Char(sysdate,'HH24:MI') " & _
		  "FROM businessunit where isactive='TRUE' ;" & _
		  "Insert into Jobmaster (FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, " & _
	      "UNITOFMEASUREMENT, ACCOUNTCODE, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME) " & _
		  "SELECT '" &  Trim(Request("FCCODE")) & "','" & Request("fcName") & "','" & Request("JobGroupCode") & "','" & Request("JOB_FORCROPTYPE") & "','" & Request("JOB_FORFIELDSTATUS") & "','"  & _
		  Request("JOB_CATEGORY") & "','" & Request("UNITOFMEASUREMENT") & "','" & Request("ACCOUNTCODE") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"',sysdate, To_Char(sysdate,'HH24:MI') " & _
		  "FROM DUAL ;" 
		  
	  ElseIf txtTable ="JobGroup" then
       q= "Insert into JobGroup (FCCODE, FCNAME, PARENTCODE, PARENTDESC, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Request("fcName") & "','" & Request("PARENTCODE") & "','" & Request("PARENTDESC") & "','"  & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	  ElseIf txtTable ="Allowded" then
       q= "Insert into Allowded (FCCODE, FCBA, FCNAME, ALLDED_GROUP, ALLDED, CROP, CONTROL_JOB, MAN_AUTO, TYPE_1, TYPE_2, " & _
	       "TYPE_3, TYPE_4, MIN_AMT, MAX_AMT, RATE_1, RATE_2, RATE_3, RATE_4, REMARKS, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Trim(Request("FCCODE")) & "','" & Session("fcBA") & "','" & Request("FCNAME") & "','" & Request("ALLDED_GROUP") & "','" & Request("ALLDED") & "','" & _
		  Request("CROP") & "','" &  Request("CONTROL_JOB") & "','" &  Request("MAN_AUTO") & "','" & _
		  Request("TYPE_1")  & "','" & Request("TYPE_2")  & "','" & Request("TYPE_3")  & "','" & Request("TYPE_4")  & "','" &  Request("MIN_AMT") & "','" & _
		  Request("MAX_AMT")  & "','" & Request("RATE_1")  & "','" & Request("RATE_2")  & "','" & Request("RATE_3")  & "','" &  Request("RATE_4") & "','" & _		  
		  Request("REMARKS")  & "','" & _		  
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
	  elseif txtTable = "Roles" then 
        q  = " Insert into roles_tab (fccode, fcname, fcentry, fcedit, fcip, LASTUPDATE, LASTTIME) values ('" & Request("FCCODE") & "','" & Request("FCNAME")& "'," & _
		        "'" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
	  elseif txtTable = "Koperasi" then
	    q = " INSERT INTO koperasi ( fccode" & _
				"                        , fcname" & _
				"                        , datecreated" & _
				"                        , registrationno" & _
				"                        , registrationdate" & _
				"                        , address" & _
				"                        , city" & _
				"                        , state" & _
				"                        , phone" & _
				"                        , fax" & _
				"                        , formco" & _
				"                        , typeco" & _
				"                        , groupco" & _
				"                        , sectorco" & _
				"                        , idno" & _
				"                        , iscertified" & _
				"                        , activation" & _
				"                        , fcentry" & _
				"                        , fcedit" & _
				"                        , fcip" & _
				"                        , fcba" & _
				"                        , lastupdate" & _
				"                        , lasttime)" & _
				" VALUES ( '" & Trim(Request("FCCODE")) & "'" & _
				"          , '" & Request("fcname") & "'" & _
				"          , TO_DATE('" & Request("datecreated") & "', 'mm/dd/yyyy')" & _
				"          , '" & Request("registrationno") & "'" & _
				"          , TO_DATE('" & Request("registrationdate") & "', 'mm/dd/yyyy')" & _
				"          , '" & Request("address") & "'" & _
				"          , '" & Request("city") & "'" & _
				"          , '" & Request("state") & "'" & _
				"          , '" & Request("phone") & "'" & _
				"          , '" & Request("fax") & "'" & _
				"          , '" & Request("formco") & "'" & _
				"          , '" & Request("typeco") & "'" & _
				"          , '" & Request("groupco") & "'" & _
				"          , '" & Request("sectorco") & "'" & _
				"          , '" & Request("idno") & "'" & _
				"          , '" & Request("iscertified") & "'" & _
				"          , '" & Request("activation") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , '" & Session("fcba") & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
	  end if%>
	 
	  
    <!-------------------------------- jika edit ---------------------------------->
		  
	<%elseif Request("Add") ="Update" or Request("txtJenis")= "Edit"  then
	
    	If Request("txttable") ="Bank" then
			q="Update bank set FCNAME ='" & Request("FCNAME") & "',CONTROLJOB='" & Request("CONTROLJOB") & "',ACCOUNTCODE='" & Request("ACCOUNTCODE") & "'" & _
				",ADDRESS='" & REQUEST("ADDRESS") & "',CITY='" & REQUEST("CITY") & "',STATE='" & REQUEST("STATE") & "',REMARKS='" & REQUEST("REMARKS") & "'" & _
				",ACTIVATION='" & REQUEST("ACTIVATION") & "'" & _
				",TRIGGER_SYSTEM='" & REQUEST("TRIGGER_SYSTEM") & "'" & _
				",BANKTYPE='" & REQUEST("BANKTYPE") & "'" & _
				",DATECREATED =to_date('" & Request("DATECREATED") & "','mm/dd/yyyy') " & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _		
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 				
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
    	ELSEIf Request("txttable") ="Stock" then
			q	=	" Update Stock set" & _
					" FCNAME = '" & Request("FCNAME") & "'" & _
					" , ITEMTYPE = '" & Request("ITEMTYPE") & "'" & _
					" , MANUFACTUREPARTNUMBER = '" & Request("MANUFACTUREPARTNUMBER") & "'" & _
					" , MANUFACTURECODE = '" & REQUEST("MANUFACTURECODE") & "'" & _
				    " , VALIDUNTIL=to_date('" & Request("VALIDUNTIL") & "','mm/dd/yyyy')" & _
					" , SUPPLIERPARTNUMBER = '" & REQUEST("SUPPLIERPARTNUMBER") & "'" & _
					" , SUPPLIERCODE = '" & REQUEST("SUPPLIERCODE") & "'" & _
					" , UNITOFMEASUREMENT = '" & REQUEST("UNITOFMEASUREMENT") & "'" & _
					" , STOCKGROUPCODE = '" & REQUEST("STOCKGROUPCODE") & "'" & _
					" , ISLESSZERO = '" & REQUEST("IsLessZero") & "'" & _
					" , LASTUPDATE = sysdate " & _
					" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
					" , ACTIVEMATERIAL = '" & REQUEST("ACTIVEMATERIAL") & "'" & _
					" , FCEDIT = '" & Session("CodeName") & "'" & _ 
					" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="BankStock" then
			q	=	" UPDATE bankstock SET" & _
					" FCNAME = '" & Request("FCNAME") & "'" & _
					" , ITEMTYPE = '" & Request("ITEMTYPE") & "'" & _
					" , UNITOFMEASUREMENT = '" & REQUEST("UNITOFMEASUREMENT") & "'" & _
					" , STOCKGROUPCODE = '" & REQUEST("STOCKGROUPCODE") & "'" & _
					" , ACTIVEMATERIAL = '" & REQUEST("ACTIVEMATERIAL") & "'" & _
					" , MANUFACTUREPARTNUMBER = '" & Request("MANUFACTUREPARTNUMBER") & "'" & _
					" , MANUFACTURECODE = '" & REQUEST("MANUFACTURECODE") & "'" & _
					" , SUPPLIERPARTNUMBER = '" & REQUEST("SUPPLIERPARTNUMBER") & "'" & _
					" , SUPPLIERCODE = '" & REQUEST("SUPPLIERCODE") & "'" & _
					" , ISLESSZERO = '" & REQUEST("IsLessZero") & "'" & _
					" , VALIDUNTIL=TO_DATE('" & Request("VALIDUNTIL") & "','mm/dd/yyyy')" & _
					" , isactive = 'Y' " & _
					" , min_value = '" & REQUEST("min_value") & "'" & _
					" , max_value = '" & REQUEST("max_value") & "'" & _
					" , stockminimum = '" & REQUEST("stockminimum") & "'" & _
					" , isbudget_control = '" & REQUEST("isbudget_control") & "'" & _
					" , LASTUPDATE = SYSDATE " & _
					" , LASTTIME = TO_CHAR(SYSDATE,'HH24:MI') " & _
					" , FCEDIT = '" & Session("CodeName") & "'" & _ 
					" where fcCode = '" & Request("fcCode") & "'" 
					
			q2  =   " MERGE INTO stock s1 " & _ 
					" USING (SELECT b.fccode fcba , a.fccode " & _
					"                     , a.fcname,a.itemtype" & _
					"                     , a.unitofmeasurement,a.MANUFACTUREPARTNUMBER,a.MANUFACTURECODE,a.stockminimum,a.activematerial,a.validuntil,a.isactive,a.stockgroupcode,a.ISLESSZERO,a.lastupdate,a.lasttime " & _
					"		FROM bankstock a LEFT JOIN businessunit b on 1=1 WHERE b.isactive='TRUE' and a.fccode = '" & Request("fcCode") & "') s2 "  & _
					" ON (s1.fccode=s2.fccode and s1.fcba=s2.fcba) " & _ 
					" WHEN MATCHED THEN UPDATE SET s1.fcname=s2.fcname,s1.itemtype=s2.itemtype,s1.unitofmeasurement=s2.unitofmeasurement,s1.MANUFACTUREPARTNUMBER=s2.MANUFACTUREPARTNUMBER,s1.MANUFACTURECODE=s2.MANUFACTURECODE,s1.stockminimum=s2.stockminimum,s1.activematerial=s1.activematerial,s1.fcedit='" & Session("CodeName") & "',lastupdate=SYSDATE " & _ 
                    " WHEN NOT MATCHED THEN INSERT (s1.fccode,s1.fcname,s1.itemtype,s1.manufacturepartnumber,s1.manufacturecode,s1.unitofmeasurement,s1.stockgroupcode,s1.fcba,s1.activematerial,s1.validuntil,s1.isactive,s1.stockminimum,s1.islesszero,s1.lastupdate,s1.lasttime) " & _ 
					" VALUES(s2.fccode,s2.fcname,s2.itemtype,s2.manufacturepartnumber,s2.manufacturecode,s2.unitofmeasurement,s2.stockgroupcode,s2.fcba,s2.activematerial,s2.validuntil,s2.isactive,s2.stockminimum,s2.islesszero,s2.lastupdate,s2.lasttime)"
				'connStr.execute(q2)
				ExecuteAndLog ScriptName, q2
				' response.write q1
					
    	ELSEIf Request("txttable") ="StockGroup" then
			q="Update StockGroup set FCNAME ='" & Request("FCNAME") & "', ACTIVATION ='" & Request("ACTIVATION") & "', GROUPTYPE ='" & Request("GROUPTYPE") & "',CONTROLJOB='" & Request("CONTROLJOB") & "'"& _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="StockSubGroup" then
			q="Update StockSubGroup set FCNAME ='" & Request("FCNAME") & "', ACTIVATION ='" & Request("ACTIVATION") & "', STOCKGROUPCODE ='" & Request("STOCKGROUPCODE") & "' "& _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcba='" & Session("fcBA") & "'"
		
		
		ELSEIf Request("txttable") ="EmpPosition" then
			q="Update empposition set DESCRIPTION ='" & Request("FCNAME") & "',FCES_EM='" & Request("ACTIVATION") & "'"& _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  POSITION='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="SectionName" then
			q="Update sectionname set FcName ='" & Request("FCNAME") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  FCCODE ='" & Request("fcCode") & "'"
		ELSEIf Request("txttable") ="Product" then
			q="Update Product set PRODUCTDESCRIPTION ='" & Request("PRODUCTDESCRIPTION") & "'" & _
				" , CROPCODE ='" & Request("CROPCODE") & "'" & _
				" , UNITOFMEASURE ='" & Request("UNITOFMEASUREMENT") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  PRODUCTCODE ='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="UnitOfMeasure" then
			q="Update UNITOFMEASURE set UNITOFMEASUREDESC ='" & Request("FCNAME") & "',ISACTIVE ='" & Request("ISACTIVE") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  UNITOFMEASURECODE='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="MillQuality" then
			q=	" Update MillQuality set fcname ='" & Request("FCNAME") & "',dirt='" & Request("dirt") & "', moisture='" & Request("moisture") & "',ffa='" & Request("ffa") & "'" & _
				" , dobi='" & Request("dobi") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  FCCODE='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="VehicleGroup" then
			q="Update VehicleGroup set DESCRIPTION ='" & Request("FCNAME") & "'" & _
				" ,	VEHICLETYPECODE ='" & Request("VEHICLETYPECODE") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  VEHICLEGROUPCODE='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="OpHarvRate" then
				qLineMax = " select nvl(max(LINENO),0) tempMaxLine" & _
						   " from OpHarvRate where fcba = '"& session("fcBA") &"'" 
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 

			q="Update OpHarvRate set ABW_CENSUS ='" & Request("bjr_sensus") & "',TERRAIN_TYPE='" & Request("soil_type") & "',ABW_MIN='" & Request("bjr_min") & "',ABW_MAX='" & Request("bjr_max") & "',BASIS_QTY='" & Request("bas_qty") & "',BASIS_RATE='" & Request("bas_rate") & "',EXTRA_BASIS_RATE='" & Request("ext_bas") & "',BASIS_HA='" & Request("bas_ha") & "'"& _

				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  CROPCODE='" & Request("fcCode") & "' and LINENO = '"& tempMaxLine &"' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="FixedAsset" then
		   p_assetgroupcode = ""
		   p_estimatelifeyear = 0
		   queryselisih = " SELECT ass.fccode, ass.fcname, ass.unitofmeasurement, ass.assetgroupcode, asg.fcname assetgroupname, asg.estimatelifeyear, asg.fcba" & _
						  " FROM asset ass" & _
						  " INNER JOIN assetgroup asg ON (asg.fcba = ass.fcba AND asg.fccode = ass.assetgroupcode)" & _
						  " WHERE ass.fcba = '" & Session("fcBA") & "' AND ass.fccode = '" & Request("assetcode") & "'"
			'response.write "queryselisih : " & queryselisih & "<br/>"
			set rsselisih = Server.CreateObject("ADODB.Recordset")
			rsselisih.Open queryselisih, connstr
			if not rsselisih.eof then
				p_assetgroupcode = rsselisih("assetgroupcode")
				p_estimatelifeyear = cdbl(to_nol(rsselisih("estimatelifeyear")))
			else
				p_assetgroupcode = ""
				p_estimatelifeyear = 0
			end if
			rsselisih.close() : set rsselisih = nothing
			q	=	" UPDATE fixedasset_hdr_tab SET" & _
					" fcname = '" & Request("fcname") & "'" & _
					" , assetcode = '" & REQUEST("assetcode") & "'" & _
					" , assetgroupcode = '" & p_assetgroupcode & "'" & _
					" , estimatelifeyear = '" & p_estimatelifeyear & "'" & _
					" , lastupdate = SYSDATE " & _
					" , lasttime = TO_CHAR(SYSDATE,'hh24:mi') " & _
					" WHERE  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="Asset" then
			q	=	" Update Asset set" & _
					" FCNAME = '" & Request("FCNAME") & "'" & _
					" , UNITOFMEASUREMENT = '" & REQUEST("UNITOFMEASUREMENT") & "'" & _
					" , ASSETGROUPCODE = '" & REQUEST("ASSETGROUPCODE") & "'" & _
					" , LASTUPDATE = sysdate " & _
					" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
					" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="AssetGroup" then
			q="Update AssetGroup set FCNAME ='" & Request("FCNAME") & "',estimatelifeyear='" & Request("estimatelifeyear") & "'"& _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
    	ELSEIf Request("txttable") ="Supplier" then
			q="Update Supplier set FCNAME ='" & Request("FCNAME") & "',CONTROLJOB='" & Request("CONTROLJOB") & "',ACTIVATION='" & Request("ACTIVATION") & "'" & _
				",ADDRESS='" & REQUEST("ADDRESS") & "',CITY='" & REQUEST("CITY") & "',STATE='" & REQUEST("STATE") & "',REMARKS='" & REQUEST("REMARKS") & "'" & _
				",BANKNAME='" & REQUEST("BANKNAME") & "',BANKACCOUNT='" & REQUEST("BANKACCOUNT") & "',BANK_OWNERSHIP='" & REQUEST("BANK_OWNERSHIP") & "'" & _
				",FAX='" & REQUEST("FAX") & "',PHONE='" & REQUEST("PHONE") & "'" & _
				",DATECREATED =to_date('" & Request("DATECREATED") & "','mm/dd/yyyy') " & _	
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _		
				" , CPERSON = '" & REQUEST("CPERSON") & "' " & _
				" , MAIL = '" & REQUEST("MAIL") & "' " & _
				" , NPWP = '" & REQUEST("NPWP") & "' " & _
				" , FCTYPE = '" & REQUEST("FCTYPE") & "'" & _
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA in (select fccode from businessunit where isactive='TRUE')"
		ELSEIf Request("txttable") ="Transporter" then
			q="Update transporter set FCNAME ='" & Request("FCNAME") & "',CONTROLJOB='" & Request("CONTROLJOB") & "',ACTIVATION='" & Request("ACTIVATION") & "'" & _
				",ADDRESS='" & REQUEST("ADDRESS") & "',CITY='" & REQUEST("CITY") & "',STATE='" & REQUEST("STATE") & "',REMARKS='" & REQUEST("REMARKS") & "'" & _
				",BANKNAME='" & REQUEST("BANKNAME") & "',BANKACCOUNT='" & REQUEST("BANKACCOUNT") & "'" & _
				",DATECREATED =to_date('" & Request("DATECREATED") & "','mm/dd/yyyy') " & _	
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _				
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 				
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="CUSTOMER" then
			q	=	" Update CUSTOMER set" & _
					" DESCRIPTION = '" & Request("DESCRIPTION") & "'" & _
					" , FCEDIT = '" & session("CodeName") & "'"  & _
					" , LASTUPDATE = sysdate " & _
					" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
					" , FCEDIT = '" & Session("CodeName") & "'" & _ 
					" where  CUSTOMERCODE='" & Request("CUSTOMERCODE") & "' and fcBA='" & Session("fcBA") & "'"
		ELSEIf Request("txttable") ="MultiLevel_Harvesting_Rate" then
			q="Update MultiLevel_Harvesting_Rate set LEVELNO ='" & Request("FCCODE") & "',CROPCODE='" & Request("CROPCODE") & "',DIVISION='" & Request("DIVISION") & "'" & _
				",PLANTING_YEAR='" & REQUEST("PLANTINGYEAR") & "',TERRAIN_TYPE='" & REQUEST("TERRAINTYPE") & "',MIN_BUNCH='" & REQUEST("MIN_BUNCH") & "',MAX_BUNCH='" & REQUEST("MAX_BUNCH") & "'" & _
				",RATE='" & REQUEST("RATE") & "',RATE_HOLIDAY='" & REQUEST("RATE_HOLIDAY") & "'" & _	
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _					
				" where  LEVELNO='" & Request("FCCODE") & "' and fcBA='" & Session("fcBA") & "'"
				' response.write q
		ELSEIf Request("txttable") ="Store" then
			q="Update Store set FCNAME ='" & Request("FCNAME") & "',REMARKS='" & Request("REMARKS") & "',ACTIVATION='" & Request("ACTIVATION") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
    	ELSEIf Request("txttable") ="Gang" then
			q="Update Gang set FCNAME ='" & Request("FCNAME") & "',REMARKS='" & Request("REMARKS") & "',ACTIVATION='" & Request("ACTIVATION") & "', GANGTYPECODE='" & _
				Request("GangTypeCode") & "' " & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
    	ELSEIf Request("txttable") ="Workshop" then
			q="Update Workshop set FCNAME ='" & Request("FCNAME") & "',REMARKS='" & Request("REMARKS") & "',ACTIVATION='" & Request("ACTIVATION") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
    	ELSEIf Request("txttable") ="Vehicle" then
			q="Update Vehicle set FCNAME ='" & Request("FCNAME") & "',VEHICLEGROUPCODE='" & Request("VEHICLEGROUPCODE") & "',VEHICLESUBGROUPCODE='" & Request("VEHICLESUBGROUPCODE") & "'" & _
				",ACTIVATION='" & Request("ACTIVATION") & "',DATETERMINATE=to_date('" & Request("DATETERMINATE") & "','mm/dd/yyyy')" & _
				",REGISTRATIONNO ='" & Request("REGISTRATIONNO") & "',MAKE='" & Request("MAKE") & "',MODEL='" & Request("MODEL") & "'" & _
				",ENGINENO ='" & Request("ENGINENO") & "',CHASISNO='" & Request("CHASISNO") & "',OWNERSHIP='" & Request.FORM("OWNERSHIP") & "'" & _
				",VEHICLESTATUS ='" & Request("VEHICLESTATUS") & "',CONTROLALLOCATION='" & Request("CONTROLALLOCATION") & "',CONTROLJOB='" & Request("CONTROLJOB") & "'" & _
				",DATECREATED =to_date('" & Request("DATECREATED") & "','mm/dd/yyyy') " & _
				",YEAROFMADE =to_date('" & Request("YEAROFMADE") & "','mm/dd/yyyy') " & _
				",YEAROFPURCHASE =to_date('" & Request("YEAROFPURCHASE") & "','mm/dd/yyyy') " & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _					
				" , WBREGISTERED = '"& request("WBREGISTERED") &"'" & _
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
    	ELSEIf Request("txttable") ="Contractor" then
			q="Update Contractor set FCNAME ='" & Request("FCNAME") & "',CONTROLJOB='" & Request("CONTROLJOB") & "',ACTIVATION='" & Request("ACTIVATION") & "'" & _
				",ADDRESS ='" & Request("ADDRESS") & "',CITY='" & Request("CITY") & "',STATE='" & Request("STATE") & "'" & _
				",REMARKS ='" & Request("REMARKS") & "',BANKNAME='" & Request("BANKNAME") & "',BANKACCOUNT='" & Request("BANKACCOUNT") & "'" & _
				",DATECREATED =to_date('" & Request("DATECREATED") & "','mm/dd/yyyy') " & _	
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _					
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
    	ELSEIf Request("txttable") ="Building" then
			q="Update Building set FCNAME ='" & Request("FCNAME") & "',BUILDGROUPCODE_DESC='" & Request("BUILDGROUPCODE_DESC") & "',ITEM='" & Request("ITEM") & "'" & _
				",ACTIVATION ='" & Request("ACTIVATION") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		 ElseIf Request("txttable") ="Field" then <!-- Additional Field Update -->
			q_lineno = "select NVL(MAX(LINENO),0)+1 LINENO from FIELDDETAIL WHERE FCBA ='" & Session("fcBA") & "' AND FCCODE = '" & Request("fcCode") & "' "
			Set rsLineno = Server.CreateObject("ADODB.Recordset")
			rsLineno.Open q_lineno, connstr
			q_his = "Insert into FIELDDETAIL (FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE, SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, " & _ 
	        "PLANTINGDATE, PLANTINGMATERIAL, TOTALSTAND, DX_BLOCK_STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW, HARVESTINGBASED_ABW, MAIN_ROAD, " & _ 
	        "COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN, SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE, " & _ 
		    "BUIST, OWLNEST, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, HARVESTINGSECTION, " & _
		    "FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05, ISACTIVE_HARVESTED,VALID_FROM,LINENO,ACTION_STATE, " & _
			"AREAL_LC,CADANGAN,OKUPASI,BIBITAN," & _
			"PABRIK,EMPLASMEN,PARIT,JALAN,PRASARANA_LAINNYA,QUARRY,TANGGUL,ENCLAVE,TAMBANG,KONSERVASI,BUKIT_SUNGAI,NO_IMAGERY,LUAS_GROSS,hectarageharvested,productivestand,harvestingsectionnumber, plantingpoint, fullstand)" & _
		    "values ('" &  Request("FCCODE") & "','" & Request("FCNAME") & "','"  &  Request("DIVISION") & "','" & _
			Request("HECTARAGEPLANTED") & "','" & Request("TERRAINTYPE") & "','"  &  Request("SOILTYPE") & "','" &  Request("OWNERSHIP") & "','" &  Request("ACTIVATION") & "','" &  Request("CROPTYPE") & "'," & _
			" to_date('" & Request("PLANTINGDATE") & "','mm/dd/yyyy') " & ",'" & Request("PLANTINGMATERIAL") & "','" & _
			Request("TOTALSTAND") & "','" & _
			Request("STATUS")  & "','" & Request("PREVSEMESTER_ABW") & "','" & Request("PRESSEMESTER_ABW") & "','" & _
			Request("HARVESTINGBASED_ABW")  & "','" & Request("MAIN_ROAD") & "','" & Request("COLLECTIONROAD") & "','" & _
			Request("MAINDRAIN")  & "','" & Request("COLLECTIONDRAIN") & "','" & Request("SUBSIDIARYDRAIN") & "','" & _
			Request("COLLECTIONPOINT")  & "','" & Request("HARVESTINGBRIDGE") & "','" & Request("PERMANENTBRIDGE") & "','" & _
			Request("WOODBRIDGE")  & "','" & Request("BUIST") & "','" & Request("OWLNEST") & "','" & _
			Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'), '" &_
			Request("HARVESTINGSECTION") & "','" & Request("FIELD_INFO_01")  & "','" & Request("FIELD_INFO_02") & "','" &_ 
			Request("FIELD_INFO_03") & "','" & Request("FIELD_INFO_04")  & "','" & Request("FIELD_INFO_05") & "','" &_
			Request("ISACTIVE_HARVESTED") & "',to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy'),'" & rsLineno("LINENO") &"','Updated','" & _
			Request("AREAL_LC")  & "','" & Request("CADANGAN") & "','" & Request("OKUPASI") &"','" & _
			Request("BIBITAN")  & "','" & Request("PABRIK") &"','" & _
			Request("EMPLASMEN")  & "','" & Request("PARIT") & "','" & Request("JALAN") & "','" & _
			Request("PRASARANA_LAINNYA")  & "','" & Request("QUARRY") & "','" & Request("TANGGUL") & "','" & _
			Request("ENCLAVE")  & "','" & Request("TAMBANG") & "','" & Request("KONSERVASI") & "','" & _
			Request("BUKIT_SUNGAI")  & "','" & Request("NO_IMAGERY") & "','" & Request("LUAS_GROSS") & "','" & Request("hectarageharvested") & "','" & Request("productivestand") & "','" & Request("harvestingsectionnumber") & "','" & Request("plantingpoint") & "','" & Request("fullstand") & "')"

			''response.write "Hasil  : " & q_his & "<br/>"
			
			q	=	" UPDATE field SET" & _
					" FCNAME = '" & Request("FCNAME") & "'" & _
					" , DIVISION = '" & Request("DIVISION") & "'" & _
					" , HECTARAGEPLANTED = '" & Request("HECTARAGEPLANTED") & "'" & _
					" , TERRAINTYPE = '" & Request("TERRAINTYPE") & "'" & _
					" , SOILTYPE = '" & Request("SOILTYPE") & "'" & _
					" , OWNERSHIP = '" & Request("OWNERSHIP") & "'" & _
					" , ACTIVATION = '" & Request("ACTIVATION") & "'" & _
					" , PLANTINGMATERIAL = '" & Request("PLANTINGMATERIAL") & "'" & _
					" , TOTALSTAND = '" & Request("TOTALSTAND") & "'" & _
					" , DX_BLOCK_STATUS ='" & Request("STATUS") & "'" & _
					" , PREVSEMESTER_ABW = '" & Request("PREVSEMESTER_ABW") & "'" & _
					" , PRESSEMESTER_ABW = '" & Request("PRESSEMESTER_ABW") & "'" & _
					" , HARVESTINGBASED_ABW = '" & Request("HARVESTINGBASED_ABW") & "'" & _
					" , MAIN_ROAD = '" & Request("MAIN_ROAD") & "'" & _
					" , COLLECTIONROAD = '" & Request("COLLECTIONROAD") & "'" & _
					" , MAINDRAIN = '" & Request("MAINDRAIN") & "'" & _
					" , COLLECTIONDRAIN = '" & Request("COLLECTIONDRAIN") & "'" & _
					" , SUBSIDIARYDRAIN = '" & Request("SUBSIDIARYDRAIN") & "'" & _
					" , COLLECTIONPOINT = '" & Request("COLLECTIONPOINT") & "'" & _
					" , HARVESTINGBRIDGE = '" & Request("HARVESTINGBRIDGE") & "'" & _
					" , PERMANENTBRIDGE = '" & Request("PERMANENTBRIDGE") & "'" & _
					" , WOODBRIDGE = '" & Request("WOODBRIDGE") & "'" & _
					" , BUIST = '" & Request("BUIST") & "'" & _
					" , OWLNEST = '" & Request("OWLNEST") & "'" & _
					" , PLANTINGDATE = to_date('" & Request("PLANTINGDATE") & "','mm/dd/yyyy') " & _
					" , LASTUPDATE = sysdate " & _
					" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
					" , HARVESTINGSECTION = '" & Request("HARVESTINGSECTION") & "'" & _
					" , FIELD_INFO_01 = '" & Request("FIELD_INFO_01") & "'" & _
					" , FIELD_INFO_02 = '" & Request("FIELD_INFO_02") & "'" & _
					" , FIELD_INFO_03 = '" & Request("FIELD_INFO_03") & "'" & _
					" , FIELD_INFO_04 = '" & Request("FIELD_INFO_04") & "'" & _
					" , FIELD_INFO_05 = '" & Request("FIELD_INFO_05") & "'" & _
					" , ISACTIVE_HARVESTED = '" & Request("ISACTIVE_HARVESTED") & "'" & _
					" , VALID_FROM = to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy')" & _
					" , AREAL_LC = '" & Request("AREAL_LC") & "'" & _
					" , CADANGAN = '" & Request("CADANGAN") & "'" & _
					" , OKUPASI = '" & Request("OKUPASI") & "'" & _
					" , BIBITAN = '" & Request("BIBITAN") & "'" & _
					" , PABRIK = '" & Request("PABRIK") & "'" & _
					" , EMPLASMEN = '" & Request("EMPLASMEN") & "'" & _
					" , PARIT = '" & Request("PARIT") & "'" & _
					" , JALAN = '" & Request("JALAN") & "'" & _
					" , PRASARANA_LAINNYA = '" & Request("PRASARANA_LAINNYA") & "'" & _
					" , QUARRY = '" & Request("QUARRY") & "'" & _
					" , TANGGUL = '" & Request("TANGGUL") & "'" & _
					" , ENCLAVE = '" & Request("ENCLAVE") & "'" & _
					" , TAMBANG = '" & Request("TAMBANG") & "'" & _
					" , KONSERVASI = '" & Request("KONSERVASI") & "'" & _
					" , BUKIT_SUNGAI = '" & Request("BUKIT_SUNGAI") & "'" & _
					" , NO_IMAGERY = '" & Request("NO_IMAGERY") & "'" & _
					" , LUAS_GROSS = '" & Request("LUAS_GROSS") & "'" & _		
					" , hectarageharvested = '" & Request("hectarageharvested") & "'" & _
					" , productivestand = '" & Request("productivestand") & "'" & _
					" , harvestingsectionnumber = '" & Request("harvestingsectionnumber") & "'" & _
					" , plantingpoint = '" & Request("plantingpoint") & "'" & _
					" , fullstand = '" & Request("fullstand") & "'" & _
					" , FCEDIT = '" & Session("CodeName") & "'" & _ 
					" WHERE fcCode = '" & Request("fcCode") & "'" & _
					"       AND fcBA = '" & Session("fcBA") & "'"
				''response.write "Hasil : " & q & "<br/>"
				
			q2	=	" UPDATE FIELD_HIST_TAB SET" & _
					" FCNAME = '" & Request("FCNAME") & "'" & _
					" , DIVISION = '" & Request("DIVISION") & "'" & _
					" , HECTARAGEPLANTED = '" & Request("HECTARAGEPLANTED") & "'" & _
					" , TERRAINTYPE = '" & Request("TERRAINTYPE") & "'" & _
					" , SOILTYPE = '" & Request("SOILTYPE") & "'" & _
					" , OWNERSHIP = '" & Request("OWNERSHIP") & "'" & _
					" , ACTIVATION = '" & Request("ACTIVATION") & "'" & _
					" , PLANTINGMATERIAL = '" & Request("PLANTINGMATERIAL") & "'" & _
					" , TOTALSTAND = '" & Request("TOTALSTAND") & "'" & _
					" , DX_BLOCK_STATUS ='" & Request("STATUS") & "'" & _
					" , PREVSEMESTER_ABW = '" & Request("PREVSEMESTER_ABW") & "'" & _
					" , PRESSEMESTER_ABW = '" & Request("PRESSEMESTER_ABW") & "'" & _
					" , HARVESTINGBASED_ABW = '" & Request("HARVESTINGBASED_ABW") & "'" & _
					" , MAIN_ROAD = '" & Request("MAIN_ROAD") & "'" & _
					" , COLLECTIONROAD = '" & Request("COLLECTIONROAD") & "'" & _
					" , MAINDRAIN = '" & Request("MAINDRAIN") & "'" & _
					" , COLLECTIONDRAIN = '" & Request("COLLECTIONDRAIN") & "'" & _
					" , SUBSIDIARYDRAIN = '" & Request("SUBSIDIARYDRAIN") & "'" & _
					" , COLLECTIONPOINT = '" & Request("COLLECTIONPOINT") & "'" & _
					" , HARVESTINGBRIDGE = '" & Request("HARVESTINGBRIDGE") & "'" & _
					" , PERMANENTBRIDGE = '" & Request("PERMANENTBRIDGE") & "'" & _
					" , WOODBRIDGE = '" & Request("WOODBRIDGE") & "'" & _
					" , BUIST = '" & Request("BUIST") & "'" & _
					" , OWLNEST = '" & Request("OWLNEST") & "'" & _
					" , PLANTINGDATE = to_date('" & Request("PLANTINGDATE") & "','mm/dd/yyyy') " & _
					" , LASTUPDATE = sysdate " & _
					" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
					" , HARVESTINGSECTION = '" & Request("HARVESTINGSECTION") & "'" & _
					" , FIELD_INFO_01 = '" & Request("FIELD_INFO_01") & "'" & _
					" , FIELD_INFO_02 = '" & Request("FIELD_INFO_02") & "'" & _
					" , FIELD_INFO_03 = '" & Request("FIELD_INFO_03") & "'" & _
					" , FIELD_INFO_04 = '" & Request("FIELD_INFO_04") & "'" & _
					" , FIELD_INFO_05 = '" & Request("FIELD_INFO_05") & "'" & _
					" , ISACTIVE_HARVESTED = '" & Request("ISACTIVE_HARVESTED") & "'" & _
					" , VALID_FROM = to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy')" & _
					" , AREAL_LC = '" & Request("AREAL_LC") & "'" & _
					" , CADANGAN = '" & Request("CADANGAN") & "'" & _
					" , OKUPASI = '" & Request("OKUPASI") & "'" & _
					" , BIBITAN = '" & Request("BIBITAN") & "'" & _
					" , PABRIK = '" & Request("PABRIK") & "'" & _
					" , EMPLASMEN = '" & Request("EMPLASMEN") & "'" & _
					" , PARIT = '" & Request("PARIT") & "'" & _
					" , JALAN = '" & Request("JALAN") & "'" & _
					" , PRASARANA_LAINNYA = '" & Request("PRASARANA_LAINNYA") & "'" & _
					" , QUARRY = '" & Request("QUARRY") & "'" & _
					" , TANGGUL = '" & Request("TANGGUL") & "'" & _
					" , ENCLAVE = '" & Request("ENCLAVE") & "'" & _
					" , TAMBANG = '" & Request("TAMBANG") & "'" & _
					" , KONSERVASI = '" & Request("KONSERVASI") & "'" & _
					" , BUKIT_SUNGAI = '" & Request("BUKIT_SUNGAI") & "'" & _
					" , NO_IMAGERY = '" & Request("NO_IMAGERY") & "'" & _
					" , LUAS_GROSS = '" & Request("LUAS_GROSS") & "'" & _		
					" , hectarageharvested = '" & Request("hectarageharvested") & "'" & _
					" , productivestand = '" & Request("productivestand") & "'" & _
					" , harvestingsectionnumber = '" & Request("harvestingsectionnumber") & "'" & _
					" , plantingpoint = '" & Request("plantingpoint") & "'" & _
					" , fullstand = '" & Request("fullstand") & "'" & _
					" , FCEDIT = '" & Session("CodeName") & "'" & _ 
					" WHERE fcCode = '" & Request("fcCode") & "'" & _
					"       AND fcBA = '" & Session("fcBA") & "'" & _ 
					"       AND year >= " & Request("YEAR") & "" & _ 
					"       AND month >= " & Request("MONTH") & "" 
				 'response.write q2
				 'connStr.execute(q2)
				 ExecuteAndLog ScriptName, q2				 
				
			
			
		 ElseIf Request("txttable") ="Employee" then
				'Dim queryselisih, rsselisih, CurrentAgeofJoin
				queryselisih = " select to_number(to_char(to_date('" & Request("DATEOFJOIN") & "','mm/dd/yyyy'), 'yyyy')) - to_number(to_char(to_date('" & Request("DATEOFBIRTH") & "','mm/dd/yyyy'), 'yyyy')) CurrentAgeofJoin" & _
							   " from dual"
				'response.write "queryselisih : " & queryselisih & "<br/>"
				set rsselisih = Server.CreateObject("ADODB.Recordset")
				rsselisih.Open queryselisih, connstr
				if not rsselisih.eof then
					CurrentAgeofJoin = cdbl(rsselisih("CurrentAgeofJoin"))
				else
					CurrentAgeofJoin = 0
				end if
				rsselisih.close()
				rsselisih=null
				'response.write "CurrentAgeofJoin : " & CurrentAgeofJoin & "<br/>"
			q="Update Employee set FCNAME ='" & Request("FCNAME") & "',SECTIONNAME='" & Request("SECTIONNAME") & "',GANGCODE='" & Request("GANGCODE") & "'" & _
				",DATEOFBIRTH =to_date('" & Request("DATEOFBIRTH") & "','mm/dd/yyyy') " & _
				",RELIGION ='" & Request("RELIGION") & "',RACE ='" & Request("RACE") & "'" & _
				",IDENTITYCARD ='" & Request("IDENTITYCARD") & "',HIGHESTEDUCATION ='" & Request("HIGHESTEDUCATION") & "'" & _
				",YEAROFGRADUATION ='" & Request("YEAROFGRADUATION") & "'" & _								
				",DATEOFJOIN =to_date('" & Request("DATEOFJOIN") & "','mm/dd/yyyy') " & _
				",DATEOFDESIGNATION =to_date('" & Request("DATEOFDESIGNATION") & "','mm/dd/yyyy') " & _								
				",PRESENTGRADE ='" & Request("PRESENTGRADE") & "',PRESENTGRADE_DETAIL ='" & Request("PRESENTGRADE_DETAIL") & "'" & _ 
				",PRESENTGRADE2 ='" & Request("PRESENTGRADE2") & "',PRESENTGRADE_DETAIL2 ='" & Request("PRESENTGRADE_DETAIL2") & "'" & _ 
				",PRESENTRATE ='" & Request("PRESENTRATE") & "'" & _
				",PREVIOUSGRADE ='" & Request("PREVIOUSGRADE") & "',PREVIOUSGRADE_DETAIL ='" & Request("PREVIOUSGRADE_DETAIL") & "'" & _
				",PREVIOUSGRADE2 ='" & Request("PREVIOUSGRADE2") & "',PREVIOUSGRADE_DETAIL2 ='" & Request("PREVIOUSGRADE_DETAIL2") & "'" & _
				",PREVIOUSRATE ='" & Request("PREVIOUSRATE") & "',MARITALSTATUS ='" & Request("MARITALSTATUS") & "',SPOUSE ='" & Request("SPOUSE") & "'" & _
				",DEPENDENTCHILD ='" & Request("DEPENDENTCHILD") & "',NPWP ='" & Request("NPWP") & "', JAMSOSTEKNO ='" & Request("JAMSOSTEKNO") & "'" & _
				",JAMSOSTEKCATEGORY ='" & Request("JAMSOSTEKCATEGORY") & "',POSITION ='" & Request("POSITION") & "'" & _									
				",DATETERMINATE =to_date('" & Request("DATETERMINATE") & "','mm/dd/yyyy') " & _
				",NATIONALITY ='" & Request("NATIONALITY") & "' ,EMP_TYPE ='" & Request("EMP_TYPE") & "'" & _	
				" , INCLUDEDTAXCALCULATION = '" & Request("INCLUDEDTAXCALCULATION") & "'" & _
				" , TAXSTATUS = '" & Request("TAXSTATUS") & "'" & _
                ",NATURETYPE ='" & Request("NATURETYPE") & "'" & _
				",GENDER ='" & Request("GENDER") & "'" & _
				",PAYMENTTYPE = '" & Request("PAYMENTTYPE") & "'" & _
				",BANKNAME = '" & Request("BANKNAME") & "'" & _
				",ACCOUNTNO = '" &Request("ACCOUNTNO") & "'" & _
				",OWNERSHIPOFACCOUNT = '" & Request("OWNERSHIPOFACCOUNT") & "'" & _
				" , TAXPAIDBYCOMPANY = '" & Request("TAXPAIDBYCOMPANY") & "'" & _
				" , ADDRESS = '" & Request("txtaddress") & "'" & _
				" , BPJSKES_CATEGORY = '" & Request("txtbpjs_category") & "'" & _
				" , BPJSKES_NO = '" & Request("BPJSKES_NO") & "'" & _
				" , company_nik = '" & Request("company_nik") & "'" & _
				" , subgrade = '" & Request("subgrade") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 				
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		 ElseIf txtTable ="RangeTaxGroup" then
			q= "Update rangetaxgroup set " & _
			   "FCNAME ='" & Request("FCNAME") & "', " & _
			   "RANGEFROM ='" & Request("RANGEFROM") & "', " & _
			   "RANGEUNTIL ='" & Request("RANGEUNTIL") & "', " & _
			   "TARIFFPERCENTAGE ='" & Request("TARIFFPERCENTAGE") & "', " & _
			   "REMARKS ='" & Request("REMARKS") & "', " & _
			   "FCENTRY ='" & Session("CodeName") & "', " & _
			   "FCEDIT ='" & Session("CodeName") & "', " & _
			   "FCIP ='" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			   "LASTUPDATE =sysdate, "& _
			   "LASTTIME =To_Char(sysdate,'HH24:MI') "& _
			   " WHERE FCCODE ='" & Request("FCCODE") & "' AND FCBA = '" & Session("fcBA") & "' "
			'response.write(Q)
		 ElseIf txtTable ="MaritalStatus" then
			q= "Update maritalstatus set " & _
			   "FCNAME ='" & Request("FCNAME") & "', " & _
			   "FCENTRY ='" & Session("CodeName") & "', " & _
			   "FCEDIT ='" & Session("CodeName") & "', " & _
			   "FCIP ='" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			   "LASTUPDATE =sysdate, "& _
			   "LASTTIME =To_Char(sysdate,'HH24:MI') "& _
			   " WHERE FCCODE ='" & Request("FCCODE") & "' "
			'response.write(Q)
		 ElseIf txtTable ="UntaxableIncome" then
			q= "Update untaxableincome set " & _
			   "FCNAME ='" & Request("FCNAME") & "', " & _
			   "UTI_EMPLOYEE ='" & Request("UTI_EMPLOYEE") & "', " & _
			   "UTI_SPOUSE ='" & Request("UTI_SPOUSE") & "', " & _
			   "UTI_CHILD ='" & Request("UTI_CHILD") & "', " & _
			   "REMARKS ='" & Request("REMARKS") & "', " & _
			   "FCENTRY ='" & Session("CodeName") & "', " & _
			   "FCEDIT ='" & Session("CodeName") & "', " & _
			   "FCIP ='" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			   "LASTUPDATE =sysdate, "& _
			   "LASTTIME =To_Char(sysdate,'HH24:MI') "& _
			   " WHERE MARITALSTATUSCODE ='" & Request("MARITALSTATUSCODE") & "' AND FCBA = '" & Session("fcBA") & "' "
			'response.write(Q)
		 ElseIf txtTable ="EmployeeAllowdedDetail" then
			q= "Update employeeallowdeddetail set " & _
			   "RATE_1 ='" & Request("RATE_1") & "', " & _
			   "RATE_2 ='" & Request("RATE_2") & "', " & _
			   "RATE_3 ='" & Request("RATE_3") & "', " & _
			   "RATE_4 ='" & Request("RATE_4") & "', " & _
			   "RATE_5 ='" & Request("RATE_5") & "', " & _
			   "REMARKS ='" & Request("REMARKS") & "', " & _
			   "FCENTRY ='" & Session("CodeName") & "', " & _
			   "FCEDIT ='" & Session("CodeName") & "', " & _
			   "FCIP ='" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
			   "LASTUPDATE =sysdate, "& _
			   "LASTTIME =To_Char(sysdate,'HH24:MI') "& _
			   " WHERE EMPLOYEECODE ='" & Request("EMPLOYEECODE") & "' AND ALLOWDEDCODE ='" & Request("ALLOWDEDCODE") & "' AND FCBA = '" & Session("fcBA") & "' "
			'response.write(Q)
		 ElseIf Request("txttable") ="CONTROLALLOWDED" then
			q="Update CONTROLALLOWDED set CROPCODE ='" & Request("CROPCODE") & "',ALLOWDEDCODE='" & Request("ALLOWDEDCODE") & "',MAIN_ACTIVITY_CODE='" & Request("MAIN_ACTIVITY_CODE") & "'" & _
				",USE_ALLOWDED_JOB ='" & Request("USE_ALLOWDED_JOB") & "',USE_ALL_GAD_JOB ='" & Request("USE_ALL_GAD_JOB") & "',CHARGEALLGADLOCATION ='" & Request("CHARGEALLGADLOCATION") & "'" & _
				",CHARGEMAINLOCATION ='" & Request("CHARGEMAINLOCATION") & "',CHARGINGMECHANISM ='" & Request("CHARGINGMECHANISM") & "',ALLOCATE ='" & Request("ALLOCATE") & "'" & _
				",MANUAL_OR_AUTO ='" & Request("MANUAL_OR_AUTO") & "',ACTIVESTATUS ='" & Request("ACTIVESTATUS") & "',PREMI_ALLOWANCE ='" & Request("PREMI_ALLOWANCE") & "'" & _
				",PENALTY_DEDUCTION ='" & Request("PENALTY_DEDUCTION") & "',FOR_EMPLOYER_EMPLOYEE ='" & Request("FOR_EMPLOYER_EMPLOYEE") & "',RICE_OR_VALUE ='" & Request("RICE_OR_VALUE") & "'" & _
				",UOM ='" & Request("UOM") & "',BS_PW_OT ='" & Request("BS_PW_OT") & "',RPT_SEQ ='" & Request("RPT_SEQ") & "'" & _
				",MINAMT ='" & Request("MINAMT") & "',MAXAMT ='" & Request("MAXAMT") & "',RATE1 ='" & Request("RATE1") & "'" & _
				",RATE2 ='" & Request("RATE2") & "',RATE3 ='" & Request("RATE3") & "',RATE4 ='" & Request("RATE4") & "'" & _
				",RATE5 ='" & Request("RATE5") & "',TYPE1 ='" & Request("TYPE1") & "',TYPE2 ='" & Request("TYPE2") & "'" & _
				",TYPE3 ='" & Request("TYPE3") & "',REFER_TABLE ='" & Request("REFER_TABLE") & "',REMARKS ='" & Request("REMARKS") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		 ElseIf Request("txttable") ="Station" then
			q="Update Station set FcName ='" & Request("FcName") & "',JOBCODE='" & Request("JOBCODE") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"	
		 ElseIf Request("txttable") ="Machine" then
			q="Update Machine set DESCRIPTION ='" & Request("DESCRIPTION") & "',FIXEDASSETCODE='" & Request("FIXEDASSETCODE") & "',ACTIVATION='" & Request("ACTIVATION") & "',OWNERSHIP='" & Request("OWNERSHIP") & "'" & _
				",PURCHASECOST ='" & Request("PURCHASECOST") & "',STATION_STATIONCODE ='" & Request("STATION_STATIONCODE") & "'" & _
				",DATEPURCHASE =to_date('" & Request("DATEPURCHASE") & "','mm/dd/yyyy') " & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" , FCEDIT = '" & Session("CodeName") & "'" & _ 
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		  ElseIf Request("txttable") ="WorkingTime" then
			q="Update WorkingTime set FCNAME ='" & Request("FCNAME") & "',TIMEIN='" & Request("TIMEIN") & "',ACTIVATION='" & Request("ACTIVATION") & "',TIMEOUT='" & Request("TIMEOUT") & "'" & _
				",ISRESTIME ='" & Request("ISRESTIME") & "',ISNEXTDAY ='" & Request("ISNEXTDAY") & "'" & _
				",ADDIN_MINUTES ='" & Request("ADDIN") & "',ADDOUT_MINUTES ='" & Request("ADDOUT") & "'" & _
				" , FCEDIT = '" & Session("CodeName") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
				
		 ElseIf Request("txttable") ="WorkingTime_Template" then
			q="Update WorkingTime_Template set FCNAME ='" & Request("FCNAME") & "',PERIOD='" & Request("PERIOD") & "',ACTIVATION='" & Request("ACTIVATION") & "'" & _
			    ",D_2 ='" & Request("D_2") & "',D_3 ='" & Request("D_3") & "'" & _
				",D_4 ='" & Request("D_4") & "',D_5 ='" & Request("D_5") & "'" & _
				",D_6 ='" & Request("D_6") & "',D_7 ='" & Request("D_7") & "'" & _
				",D_1 ='" & Request("D_1") & "'" & _
				" , FCEDIT = '" & Session("CodeName") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		
		 ElseIf Request("txttable") ="Employee_WorkingTime" then
			q="Update Employee_WorkingTime set ACTIVATION='" & Request("ACTIVATION") & "'" & _
			    ",D_1 ='" & Request("D_1") & "',D_2 ='" & Request("D_2") & "'" & _
				",D_3 ='" & Request("D_3") & "',D_4 ='" & Request("D_4") & "'" & _
				",D_5 ='" & Request("D_5") & "',D_6 ='" & Request("D_6") & "'" & _
				",D_7 ='" & Request("D_7") & "',D_8 ='" & Request("D_8") & "'" & _
				",D_9 ='" & Request("D_9") & "',D_10 ='" & Request("D_10") & "'" & _
				",D_11 ='" & Request("D_11") & "',D_12 ='" & Request("D_12") & "'" & _
				",D_13 ='" & Request("D_13") & "',D_14 ='" & Request("D_14") & "'" & _
				",D_15 ='" & Request("D_15") & "',D_16 ='" & Request("D_16") & "'" & _
				",D_17 ='" & Request("D_17") & "',D_18 ='" & Request("D_18") & "'" & _
				",D_19 ='" & Request("D_19") & "',D_20 ='" & Request("D_20") & "'" & _
				",D_21 ='" & Request("D_21") & "',D_22 ='" & Request("D_22") & "'" & _
				",D_23 ='" & Request("D_23") & "',D_24 ='" & Request("D_24") & "'" & _
				",D_25 ='" & Request("D_25") & "',D_26 ='" & Request("D_26") & "'" & _
				",D_27 ='" & Request("D_27") & "',D_28 ='" & Request("D_28") & "'" & _
				",D_29 ='" & Request("D_29") & "',D_30 ='" & Request("D_30") & "'" & _
				", D_31 = '" & Session("D_31") & "'" & _
				", FCEDIT = '" & Session("CodeName") & "'" & _
				", LASTUPDATE = sysdate " & _
				", LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				"WHERE  EMPCODE='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'" & _
				"AND PERIOD_YEAR='" & Request("PERIOD_YEAR") & "' AND PERIOD_MONTH='" & Request("PERIOD_MONTH")  & "'"
		
		 ElseIf Request("txttable") ="Allowded" then
			q="Update Allowded set MAN_AUTO ='" & Request("MAN_AUTO") & "',TYPE_1 ='" & Request("TYPE_1") & "',TYPE_2='" & Request("TYPE_2") & "',TYPE_3='" & Request("TYPE_3") & "'" & _
				",TYPE_4 ='" & Request("TYPE_4") & "',MIN_AMT ='" & Request("MIN_AMT") & "',MAX_AMT ='" & Request("MAX_AMT") & "'" & _
				",RATE_1 ='" & Request("RATE_1") & "',RATE_2 ='" & Request("RATE_2") & "'" & _
				",RATE_3 ='" & Request("RATE_3") & "',RATE_4 ='" & Request("RATE_4") & "',REMARKS ='" & Request("REMARKS") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		 ElseIf Request("txttable") ="Roles" then
			q="Update Roles_Tab set FCNAME ='" & Request("FCNAME") &  "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "'"			
         ElseIf Request("txttable") = "Koperasi" then
		   q = " UPDATE koperasi" & _
			   " SET fcname = '" & Request("fcname") & "'" & _
			   "     , datecreated = TO_DATE('" & Request("datecreated") & "', 'mm/dd/yyyy')" & _
			   "     , registrationno = '" & Request("registrationno") & "'" & _
			   "     , registrationdate = TO_DATE('" & Request("registrationdate") & "', 'mm/dd/yyyy')" & _
			   "     , address = '" & Request("address") & "'" & _
			   "     , city = '" & Request("city") & "'" & _
			   "     , state = '" & Request("state") & "'" & _
			   "     , phone = '" & Request("phone") & "'" & _
			   "     , fax = '" & Request("fax") & "'" & _
			   "     , formco = '" & Request("formco") & "'" & _
			   "     , typeco = '" & Request("typeco") & "'" & _
			   "     , groupco = '" & Request("groupco") & "'" & _
			   "     , sectorco = '" & Request("sectorco") & "'" & _
			   "     , idno = '" & Request("idno") & "'" & _
			   "     , iscertified = '" & Request("iscertified") & "'" & _
			   "     , activation = '" & Request("activation") & "'" & _
			   "     , fcedit = '" & Session("CodeName") & "'" & _
			   "     , lastupdate = SYSDATE " & _
			   "     , lasttime = TO_CHAR(SYSDATE,'HH24:MI')" & _	
			   " WHERE fcba = '" & Session("fcba") & "' AND fccode = '" & Request("fccode") & "'"
 		 end if

	end if
	
	If Request("txttable") ="Employee"  then
		if Request("txtJenis")= "New" or Request("Add") ="Update" or Request("txtJenis")= "Edit"  then
			if CurrentAgeofJoin >= 18 then
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
			else
				err.number = 1001
				err.description = "Umur Karyawan tersebut belum 18+, Jadi tidak memenuhi syarat untuk bekerja sehingga tidak bisa di input ke dalam master employee! Hubungi manajemen anda!"
			end if
		else
			'connStr.execute(q)
			ExecuteAndLog ScriptName, q
			
		end if
	else
		'Response.Write "XX -" & q
		response.flush
		'connStr.execute(q)
		ExecuteAndLog ScriptName, q
		
	end if
	
	If Request("txttable") ="Field"  then
		''Response.Write q_his
	   'connStr.execute(q_his)
	   ExecuteAndLog ScriptName, q_his
	end if
	
	If Request("txttable") ="BankStock" and Request("txtJenis")= "New" Then
		response.flush
		'connStr.execute(q1)
		ExecuteAndLog ScriptName, q1
	end IF
	'If Request("txttable") ="BankStock" and (Request("Add") ="Update" or Request("txtJenis")= "Edit")  then
		''Response.Write q_his
	'   connStr.execute(q1)
	'end if
	' IF Length(trim(q1)) > 0 THEN
	'      connStr.execute(q1)
	' End IF
    cmdView	= "True"
else
    cmdView	= request("cmdView")
End if

if err.number <> 0 then
   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
   cmdView	= "True"
End if

Dim bln, thn
bln = Month(Date)
thn = Year(Date)

qLineMax = " select ARESTA_CLOSING from EMPALLDEDSTATUS where fcba = '"& session("fcBA") &"' and YEAR = "&thn&" and MONTH = '"&bln&"'" 
set rsqLineMax = Server.CreateObject("ADODB.Recordset")
rsqLineMax.Open qLineMax, connstr

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,fddate,txtJenis,fcCode,TXT_FCCODE,EMPLOYEECODE,TXT_YEAR,TXT_MONTH,GANGCODE 
TXT_FCCODE= Ucase(Request("TXT_FCCODE"))
EMPLOYEECODE = Request("EMPLOYEECODE")
fddate   =Request("fddate")
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
fcCode =   Request("fcCode")
txtJenis =Request("txtJenis")
TXT_YEAR = REQUEST("TXT_YEAR")
TXT_MONTH = REQUEST("TXT_MONTH")
GANGCODE = REQUEST("GANGCODE")


%>
<%If txtTable <> "Field" then%>
	<style>
		/* Card filter selebar halaman, isi tetap di kiri */
		form.tm-filter > table.data { width: 100% !important; max-width: none !important; }
		form.tm-filter td.tm-fill { width: 100%; padding: 0 !important; }
	</style>
	<form name="eform" class="tm-filter" method="post" action="TableMaster.asp" style="display: <%if txtTable <> "CUSTOMER" And txtTable <> "Employee" And txtTable <> "Transporter" then%> <%if rsqLineMax("ARESTA_CLOSING") = "TRUE" then%>none<%else%>block <%end if%> <%end if%>">
		<table class="data"  cellpadding="0" cellspacing="0">
			<Tr> 
			<td width="100" align="left"><font face="Verdana" size="1"><%=Pencarian%></font></td>
			<td width="200" align="left"><input type="text" size="25" maxlength="30"  name="TXT_FCCODE" value="<%=Request("TXT_FCCODE")%>"   title="CODE Character a s/d z or 1 s/d 9"></td>
		
		<%IF txttable="Employee_WorkingTime" or txttable="Employee" or txttable="Field"  then%>
			
			<td width="10" align="left"><font face="Verdana" size="1">Period</font></td>
			<td >
			<% 
				Dim mon, Yr, MinYr, MaxYr, xx, MM, v_txtMonth, v_txtYear
				mon = Month(Date)
				v_txtMonth = TXT_MONTH 
				v_txtYear  = TXT_YEAR
				If len(v_txtMonth)= 0 then v_txtMonth = mon

				Yr = Year(Date)
				If len(v_txtYear) = 0 then v_txtYear = Yr
				MinYr = Yr - 2
				MaxYr = Yr + 2
								
			%>
				<select name="TXT_YEAR" id="TXTYEAR_">
				<%	For xx = MinYr to MaxYr %>
					<Option Value="<%=xx%>" <%If cint(xx) = cint(v_txtYear) then%> Selected <%End If%>><%=xx%></option>
				<%	Next %>
				</select>
			</td>
			<td>
				<select name="TXT_MONTH" id="TXTMONTH_"> 
				<%	For xx = 1 to 12	%>
						<Option Value="<%=xx%>" <%If cint(xx) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(xx)%></option>
				<%	Next %>
				</select>	
			</td>
			<%End IF%>	
			
			<%IF txttable="Employee_WorkingTime" then%>
			
			<td width="10" align="left"><font face="Verdana" size="1">GangCode</font></td>
			<td >
				<select name="GANGCODE" id="GANGCODE_">
						<%cmdCallCombo "Gang",GANGCODE,"",""%>
				</select>
			</td>
			<%End IF%>	
			
			
			
			<input type="hidden" name="cmdView" value="True"> 
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<td><input type="Submit"  name="Add2" value="View"></td>
			
			<%IF txttable="WorkingTime_Template" then%>
			<td width="50" align="left"><font face="Verdana" size="1">Generate Period</font></td>
			<td >
			<% 
				Dim  YrNow, MinYrNow, MaxYrNow, xy, wt_txtYear
				
				YrNow = Year(Date)
				If len(wt_txtYear) = 0 then wt_txtYear = YrNow
				MinYrNow = YrNow - 2
				MaxYrNow = YrNow + 2
								
			%>
				<select name="TXT_YEAR" id="TXTYEAR_">
				<%	For xy = MinYrNow to MaxYrNow %>
					<Option Value="<%=xy%>" <%If cint(xy) = cint(wt_txtYear) then%> Selected <%End If%>><%=xy%></option>
				<%	Next %>
				</select>
			</td>
			<td>
			<script language="text/javascript">
			function GenerateWKTemplate() {
			var txtyear_ = document.getElementById("txtYear_");
			var yr_ = txtyear_.options[txtyear_.selectedIndex].value;
			alert yr_
			//window.location.href = "upload_ba_fp_01.Asp?txtJenis=Preview&txtYear="+yr_+"&txtMonth="+mn_;
			}		
			</script>  
			<input type="Button" value="Generate" onClick="javascript:var txtyear_ = document.getElementById('TXTYEAR_');var yr_ = txtyear_.options[txtyear_.selectedIndex].value;window.open('p_generate_workingtime.Asp?txtYear='+yr_); ">&nbsp;</td>
			
			<%End IF%>	
			<td class="tm-fill"></td>
			</tr>
			
		</table>  
	</form>
<%End If%>
<%
'' ------------GET INIT EDIT For List Data---------------
Dim rsSetup 

set rsSetup = server.CreateObject("Adodb.recordset")
if Request("txtJenis")="Edit" then
     ' if txtModule <> "Administrator" then 
			if txtTable = "MaritalStatus" then
			q = "select * from " & txtTable & " where  " & crashId & "='" & request(crashId) & "'"
			elseif txtTable = "FixedAsset" then
				q = "select * from fixedasset_hdr_tab where  " & crashId & "='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "BankStock" then
				q = "select * from bankstock where isactive='Y' AND " & crashId & "='" & request(crashId) & "'"
			elseif txtTable = "StockSubGroup" then
				q = "select * from StockSubGroup where  " & crashId & " = '" &  request(crashId) & "%'  AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "MultiLevel_Harvesting_Rate" then
				q = "select * from " & txtTable & " where LEVELNO ='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "EmpPosition" then
				q = "select POSITION fcCode, DESCRIPTION fcname, FCES_EM from " & txtTable & " where  POSITION ='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "SectionName" then
				q = "select * from " & txtTable & " where  " & crashId & "='" & request(crashId) & "'"
			elseif txtTable = "Product" then
				q = "select PRODUCTCODE fcCode, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA from " & txtTable & " where  PRODUCTCODE ='" & request(crashId) & "'"
			elseif txtTable = "UnitOfMeasure" then
				q = "select UNITOFMEASURECODE fcCode, UNITOFMEASUREDESC fcname, ISACTIVE from " & txtTable & " where  UNITOFMEASURECODE ='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "VehicleGroup" then
				q = "select VEHICLEGROUPCODE fcCode, DESCRIPTION fcname, VEHICLETYPECODE from " & txtTable & " where  VEHICLEGROUPCODE ='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "OpHarvRate" then
				qLineMax = " select nvl(max(LINENO),0) tempMaxLine" & _
								   " from OpHarvRate where fcba = '"& session("fcBA") &"'" 
						set rsqLineMax = Server.CreateObject("ADODB.Recordset")
						rsqLineMax.Open qLineMax, connstr
						if not rsqLineMax.eof then
							tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
						else
							tempMaxLine = 0
						end if 
				q = "select * from " & txtTable & " where  CROPCODE ='" & request(crashId) & "' AND LINENO = '"& tempMaxLine &"' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "Roles" then 
				q = "select fccode, fcname from Roles_Tab where  " & crashId & "='" & request(crashId) & "'"
			elseif txtTable = "Employee_WorkingTime" then 
				q = "select * from Employee_WorkingTime where  EmpCode='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "' " & _ 
				    "AND (Period_Month='" & REQUEST("TXT_MONTH") & "' or Period_Month='" & REQUEST("PERIOD_MONTH") & "') " & _ 
					"AND (Period_Year='" & REQUEST("TXT_YEAR") & "'  or Period_Year='" & REQUEST("PERIOD_YEAR") & "')"
			   'q = "select * from Employee_WorkingTime where  EmpCode='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
			elseif txtTable = "Field" then 
			q	=	" SELECT" & _
				"  MONTH,YEAR,FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE" & _
				" , SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, PLANTINGDATE" & _
				" , PLANTINGMATERIAL, TOTALSTAND, DX_BLOCK_STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW" & _
				" , HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN" & _
				" , SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE" & _
				" , BUIST, OWLNEST" & _
				" , FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05,ISACTIVE_HARVESTED, VALID_FROM" & _
				" , AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA" & _
				" , QUARRY, TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTION, HARVESTINGSECTIONNUMBER,PLANTINGPOINT, FULLSTAND " & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM Field_Hist_Tab " & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (MONTH = '" & TXT_MONTH & "' OR MONTH='" & REQUEST("MONTH") & "')" & _
				" AND (YEAR = '" & TXT_YEAR & "' OR YEAR='" & REQUEST("YEAR") & "')" & _ 
				" AND (FCCODE = '" & request(crashId) & "')" 
			elseif txtTable = "CUSTOMER" then 		
			q = "select * from CUSTOMER WHERE CUSTOMERCODE = '"& request(crashId) &"' AND FCBA='" & SESSION("FCBA") & "'"
		
			else

				q = "select * from " & txtTable & " where  " & crashId & "='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"

		        
		end if
					' else
						' q = "select * from " & txtTable & " where  " & crashId & "='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
					' end if
		        
		' end if
	 ' else
	     ' if txtTable = "Roles" then 
              ' q = "select fccode, fcname from Roles_Tab where  " & crashId & "='" & request(crashId) & "'"
		 ' end if
	 ' end if
	 
ELSE
     if txtTable = "MaritalStatus" then
			q = "select * from " & txtTable & ""
		elseif txtTable = "FixedAsset" then
			q = "SELECT * FROM fixedasset_hdr_tab WHERE fcba = '" & SESSION("FCBA") &"'"
		elseif txtTable = "MultiLevel_Harvesting_Rate" then
			q = "select * from " & txtTable & " where LEVELNO ='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "'"
		elseif txtTable = "BankStock" then
			q = "select * from bankstock where isactive='Y' AND  " & crashId & " like '" & TXT_FCCODE & "%' or Upper(fcname) Like '%" & TXT_FCCODE & "%'"
		elseif txtTable = "StockSubGroup" then
			q = "select * from StockSubGroup where  " & crashId & " like '" & TXT_FCCODE & "%' or Upper(fcname) Like '%" & TXT_FCCODE & "%' AND FCBA='" & SESSION("FCBA") & "'"
		elseif txtTable = "SectionName" then
			q = "select * from " & txtTable & " WHERE FCCODE = '"& TXT_FCCODE &"'"
		elseif txtTable = "Product" then
			q = "select PRODUCTCODE fcCode, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA from " & txtTable & " where  PRODUCTCODE ='" & request(crashId) & "'"
		elseif txtTable = "Roles" then 
              q = "select fccode, fcname from Roles_Tab where  " & crashId & "='" & request(crashId) & "'"
	    elseif txtTable = "Employee_WorkingTime" then 
			q = "select * from Employee_WorkingTime where  EmpCode='" & request(crashId) & "' AND FCBA='" & SESSION("FCBA") & "' " & _ 
				"AND (Period_Month='" & REQUEST("TXT_MONTH") & "' or Period_Month='" & REQUEST("PERIOD_MONTH") & "') " & _ 
				"AND (Period_Year='" & REQUEST("TXT_YEAR") & "'  or Period_Year='" & REQUEST("PERIOD_YEAR") & "')"
		elseif txtTable = "Field" then 
			q	=	" SELECT" & _
				"  MONTH,YEAR,FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE" & _
				" , SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, PLANTINGDATE" & _
				" , PLANTINGMATERIAL, TOTALSTAND, DX_BLOCK_STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW" & _
				" , HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN" & _
				" , SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE" & _
				" , BUIST, OWLNEST" & _
				" , FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05,ISACTIVE_HARVESTED, VALID_FROM" & _
				" , AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA" & _
				" , QUARRY, TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTION, HARVESTINGSECTIONNUMBER, PLANTINGPOINT, FULLSTAND " & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM Field_Hist_Tab " & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (MONTH = '" & TXT_MONTH & "' OR MONTH='" & REQUEST("MONTH") & "')" & _
				" AND (YEAR = '" & TXT_YEAR & "' OR YEAR='" & REQUEST("YEAR") & "')" & _ 
				" AND FCCODE = '" & TXT_FCCODE & "'" 
		elseif txtTable = "CUSTOMER" then 		
			q = "select * from CUSTOMER WHERE CUSTOMERCODE = '"& TXT_FCCODE &"'"
		else
			q = "SELECT * FROM " & txtTable & " WHERE fcba = '" & SESSION("FCBA") &"'"

		end if





END IF
'response.write("Jebis:<" & Request("txtJenis") & "> Table:<" & txtTable & "> Yr:"& REQUEST("TXT_MONTH") & " q:" & REQUEST("PERIOD_MONTH"))
Set rsSetup = Server.CreateObject("ADODB.Recordset")
rsSetup.CursorType = 3
rsSetup.CursorLocation = 3
'response.write "xx1-" & q
rsSetup.Open q, connstr
%>
 <!-------------isi Form-->
 <form name="frmEntry" method="post" action="TableMaster.asp" id="frmEntry" >
		<%if txtTable <> "CUSTOMER" And txtTable <> "Employee" And txtTable <> "Transporter" And txtTable <> "Field" And txtTable <> "StockGroup" And txtTable <> "StockSubGroup" And txtTable <> "BankStock" And txtTable <> "Bank" And txtTable <> "Supplier" And txtTable <> "Store" And txtTable <> "SectionName" And txtTable <> "Gang" And txtTable <> "EmpPosition" And txtTable <> "EmpPosition" And txtTable <> "MaritalStatus" And txtTable <> "Workshop" And txtTable <> "VehicleGroup" And txtTable <> "Vehicle" And txtTable <> "Contractor" And txtTable <> "Building" And txtTable <> "Station" And txtTable <> "Machine" And txtTable <> "Koperasi" And txtTable <> "MillQuality" And txtTable <> "02-paymentvoucher" then%>
			<%if rsqLineMax("ARESTA_CLOSING") = "TRUE" then%>
				<center>
					<p style="color: red; font-size: 16px">Mohon maaf Anda tidak bisa input data ARESTA, Karena sudah CLOSING!</p>
				</center>
			<%end if%>
		<%end if%>
	    <table class="data" 
			style="display: <%if txtTable <> "CUSTOMER" And txtTable <> "Employee" And txtTable <> "Transporter" And txtTable <> "Field" And txtTable <> "StockGroup" And txtTable <> "StockSubGroup" And txtTable <> "BankStock" And txtTable <> "Bank" And txtTable <> "Supplier" And txtTable <> "Store" And txtTable <> "SectionName" And txtTable <> "Gang" And txtTable <> "EmpPosition" And txtTable <> "EmpPosition" And txtTable <> "MaritalStatus" And txtTable <> "Workshop" And txtTable <> "VehicleGroup" And txtTable <> "Vehicle" And txtTable <> "Contractor" And txtTable <> "Building" And txtTable <> "Station" And txtTable <> "Machine" And txtTable <> "Koperasi" And txtTable <> "MillQuality" And txtTable <> "02-paymentvoucher" then%> <%if rsqLineMax("ARESTA_CLOSING") = "TRUE" then%>none<%else%>block <%end if%> <%end if%>">
			<%'if not rsSetup.eof Then
			
			if txtTable <> "Administrator" then 
				nCountofField=rsSetup.fields.count-1
			else
				nCountofField=rsSetup.fields.count
			end if
			'End if
			Dim txtBoxColor,txtTitle
			%>
		  
			<%'----------------------Atur Tampilan Form 
			%>
			<!--#include file="Select_Form_Show_Master.asp"-->
	
			<tr> 
			<td colspan="4"  align="left"> <input Type="Hidden" Value="True" Name="cmdView"> 
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable"> 
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule"> 
				
			<%if Request("txttable") ="Employee" then%> 
					<%IF Request("txtJenis")="Edit" then%> 
					<input type="hidden" name="txtJenis" value="Edit"> 
					<input Type="Submit"  name="Add" value="Update">
					
					<!-- <a href="TableMaster.Asp?txtModule=Data Master&txtTable=EmployeeAllowdedDetail&TXT_FCCODE=<%=rsSetup("FCCODE")%>&txtEMPLOYEECODE=<%=rsSetup("FCCODE")%>" class="jqplot-image-button ">Add Allowded Detail</a>-->
					<%Else%> 
					<input type="hidden" name="txtJenis" value="New"> 
					<input Type="Submit"  name="Add" value="Add"> 
					<%End if%> 
					
			<!--Additional Validasi Master Field-->	  
			<% elseif Request("txttable") ="Field" then%> 

					<%IF Request("txtJenis")="Edit"   then%> 
				  		<input type="hidden" name="txtJenis" value="Edit">
					<% 
					Dim yr_ , mn_
					yr_ = trim(REQUEST("TXT_YEAR"))
					mn_ = trim(REQUEST("TXT_MONTH"))
					
					if yr_  = "" then 
						yr_ = Request("YEAR")
					end if
					
					if mn_  = "" then 
						mn_ = Request("MONTH")
					end if
					
					if StatusAccClosing(yr_,mn_,session("fcBA"))="FALSE" then 
					%>
				  <input Type="button" name="Add" value="Update" OnClick="submitData()"> <!-- OnClick ="periksa(this.form)"> -->
				  <%Else%> 
				  <span ><Font color=red>PERIODE SUDAH DI CLOSE !!!</font></span>
					<% End if%>
				  <%Else%> 
				  <input type="hidden" name="txtJenis" value="New"> 
				  <input Type="button" name="Add" value="Add" OnClick="submitData()"> 	<!-- OnClick ="periksa(this.form)"> -->
				  <%End if%> 

			<% else %>
					<%IF Request("txtJenis")="Edit" then%> 
					<input type="hidden" name="txtJenis" value="Edit">
					<input Type="Submit"  name="Add" value="Update">
					
					<%Else%> 
						<input type="hidden" name="txtJenis" value="New"> 
						<input Type="Submit"  name="Add" value="Add"> 
					<%End if%> 
			<% end if %>
				
					<input Type="Reset" onclick="window.history.go(-1); return false;"  name="Reset" value="Cancel">
					<%IF txtModule="Data Master" or txtModule="Data Setup" Then%>	
					<a href="PrintData.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&txtSearch=<%=request("fcCode")%>&txtField=<%="fcCode"%>" 
					title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a> * Mandatory Field 
					<%Else%>
					<a href="PrintData.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&fcCode=<%=fcCode%>" 
					title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a> * Mandatory Field 
					<%End if%>
				
			<%IF txtModule="Data Master" or txtModule="Data Setup" Then%>	
				<% If txtTable <> "Field" then%>
					<a href="PrintData.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&txtSearch=<%=request("fcCode")%>&txtField=<%="fcCode"%>" 
					title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a> * Mandatory Field 
				<%End If%>
			<%Else%>
				<a href="PrintData.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&fcCode=<%=fcCode%>" 
				title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a> * Mandatory Field 
			<%End if%>
			</td>
			</tr>
		</table> 
         
 </form>
<!--#include file="validation_form_master.asp"--> 

<%if cmdView = "True" Then
   Dim rsGAD
	If Request("txttable") = "Bank" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, CONTROLJOB, ACCOUNTCODE, ADDRESS" & _
				" , CITY, STATE, REMARKS, DATECREATED, ACTIVATION" & _
				" , TRIGGER_SYSTEM, BANKTYPE" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "BankStock" Then
		q	=	" SELECT fccode, fcname, itemtype, unitofmeasurement, stockgroupcode, stocksubgroupcode,activematerial" & _
				"        , manufacturepartnumber, manufacturecode, supplierpartnumber, suppliercode, islesszero, validuntil" & _
				"        , min_value, max_value, stockminimum, isbudget_control, fcentry, fcedit, fcip, lastupdate FROM " & txtTable & "" & _
				" WHERE isActive='Y' AND ((FCCODE like '" & TXT_FCCODE & "%') or (UPPER(FCNAME) like '%" & TXT_FCCODE & "%')) ORDER BY FCCODE "
	ElseIf Request("txttable") = "StockGroup" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, CONTROLJOB, ACTIVATION, GROUPTYPE" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "StockSubGroup" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, STOCKGROUPCODE, ACTIVATION,'X' as Temp " & _
				" , FCENTRY, FCEDIT, FCIP FROM " & txtTable & "" & _
				" WHERE " & _
				" (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "') AND FCBA='" & SESSION("FCBA") & "' ORDER BY FCCODE"

	ElseIf Request("txttable") = "FixedAsset" Then
		q	=	" SELECT fccode, fcname, assetcode, assetgroupcode, estimatelifeyear, fcentry, fcedit, fcip, fcba FROM fixedasset_hdr_tab" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Asset" Then
		q	=	" SELECT FCCODE, FCNAME" & _
				" , UNITOFMEASUREMENT,ASSETGROUPCODE" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"	
	ElseIf Request("txttable") = "AssetGroup" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, estimatelifeyear" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Supplier" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, FCTYPE||'-'||parameter_api.get_fcname(fctype,'SUPPLIER_TYPE',fcba) FCTYPENAME,FCTYPE, CONTROLJOB, DATECREATED, ACTIVATION" & _
				" , ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT, BANK_OWNERSHIP, PHONE, FAX , CPERSON, MAIL, NPWP " & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND ((FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')" & _ 
				" OR (UPPER(FCNAME) like '%" & Ucase(TXT_FCCODE) & "%' OR UPPER(FCNAME) ='" & REQUEST("FCCODE") & "')) " & _
                " ORDER BY FCCODE "				

	ElseIf Request("txttable") = "EmpPosition" Then
		q	=	" SELECT" & _
				" POSITION, DESCRIPTION, FCES_EM" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'"
	ElseIf Request("txttable") = "OpHarvRate" Then
		q	=	" SELECT" & _
				" CROPCODE, ABW_CENSUS, TERRAIN_TYPE, ABW_MIN, ABW_MAX, BASIS_QTY, BASIS_RATE, EXTRA_BASIS_RATE, BASIS_HA" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (CROPCODE='" & TXT_FCCODE & "' OR CROPCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "UnitOfMeasure" Then
		q	=	" SELECT" & _
				" UNITOFMEASURECODE, UNITOFMEASUREDESC, ISACTIVE" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (UNITOFMEASURECODE like '" & TXT_FCCODE & "%' OR UNITOFMEASURECODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "SectionName" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, '','','',''" & _
				" FROM " & txtTable & "" & _
				" WHERE (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Product" Then
		q	=	" SELECT" & _
				" PRODUCTCODE, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA, FCENTRY, FCEDIT, FCIP, FCBA" & _
				" FROM " & txtTable & "" & _
				" WHERE (PRODUCTCODE like '" & TXT_FCCODE & "%' OR PRODUCTCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "VehicleGroup" Then
		q	=	" SELECT" & _
				" VEHICLEGROUPCODE, DESCRIPTION, VEHICLETYPECODE" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (VEHICLEGROUPCODE like '" & TXT_FCCODE & "%' OR VEHICLEGROUPCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Transporter" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, CONTROLJOB, DATECREATED, ACTIVATION" & _
				" , ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "MultiLevel_Harvesting_Rate" Then
		q	=	" SELECT" & _
				" LEVELNO, CROPCODE, DIVISION, PLANTING_YEAR, TERRAIN_TYPE" & _
				" , MIN_BUNCH, MAX_BUNCH, RATE, RATE_HOLIDAY" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (LEVELNO='" & TXT_FCCODE & "' OR LEVELNO='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Store" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, REMARKS, ACTIVATION" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Gang" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, REMARKS, ACTIVATION, GANGTYPECODE" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')" & _
				" ORDER BY REMARKS,FCCODE"
	ElseIf Request("txttable") = "RangeTaxGroup" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, RANGEFROM, RANGEUNTIL, TARIFFPERCENTAGE," & _
				" REMARKS , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE='" & TXT_FCCODE & "' OR FCCODE='" & REQUEST("FCCODE") & "') "
				'RESPONSE.WRITE(q)
	ElseIf Request("txttable") = "MaritalStatus" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME," & _
				" FCENTRY, FCEDIT, FCIP, LASTUPDATE FROM " & txtTable & "" & _
				" WHERE " & _
				" (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "') "
				'RESPONSE.WRITE(q)
	ElseIf Request("txttable") = "UntaxableIncome" Then
		q	=	" SELECT" & _
				" MARITALSTATUSCODE, FCNAME, UTI_EMPLOYEE, UTI_SPOUSE, UTI_CHILD," & _
				" REMARKS , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (MARITALSTATUSCODE='" & TXT_FCCODE & "' OR MARITALSTATUSCODE='" & REQUEST("MARITALSTATUSCODE") & "') "
				'RESPONSE.WRITE(q)
	ElseIf Request("txttable") = "Employee" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, SECTIONNAME, GANGCODE, EMP_TYPE, DATEOFBIRTH" & _
				" , RELIGION, RACE, IDENTITYCARD, HIGHESTEDUCATION, YEAROFGRADUATION" & _
				" , DATEOFJOIN, DATEOFDESIGNATION, PRESENTGRADE, PRESENTGRADE_DETAIL, PRESENTRATE" & _
				" , PREVIOUSGRADE, PREVIOUSGRADE_DETAIL, PREVIOUSRATE, PRESENTGRADE2, PRESENTGRADE_DETAIL2, PREVIOUSGRADE2, PREVIOUSGRADE_DETAIL2" & _ 
				" ,	MARITALSTATUS, SPOUSE" & _
				" , DEPENDENTCHILD, JAMSOSTEKNO, JAMSOSTEKCATEGORY, POSITION, DATETERMINATE" & _
				" , NPWP, NATIONALITY, INCLUDEDTAXCALCULATION, TAXSTATUS,GENDER,PAYMENTTYPE,ACCOUNTNO,BANKNAME,OWNERSHIPOFACCOUNT, TAXPAIDBYCOMPANY " & _
				" , ADDRESS, BPJSKES_CATEGORY, BPJSKES_NO" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
		
	ElseIf Request("txttable") = "EmployeeAllowdedDetail" Then
		q	=	" SELECT" & _
				" EMPLOYEECODE, ALLOWDEDCODE, RATE_1, RATE_2, RATE_3" & _
				" , RATE_4, RATE_5, REMARKS , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (EMPLOYEECODE='" & TXT_FCCODE & "' OR EMPLOYEECODE='" & REQUEST("EMPLOYEECODE") & "') "
				'RESPONSE.WRITE(q)
	ElseIf Request("txttable") = "Workshop" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, REMARKS, ACTIVATION" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Vehicle" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, VEHICLEGROUPCODE,VEHICLESUBGROUPCODE, DATECREATED, ACTIVATION,DATETERMINATE" & _
				" , REGISTRATIONNO, MAKE, MODEL, ENGINENO, CHASISNO" & _
				" , YEAROFMADE, YEAROFPURCHASE, OWNERSHIP, VEHICLESTATUS, CONTROLALLOCATION, CONTROLJOB, WBREGISTERED" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like'" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Contractor" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, CONTROLJOB, DATECREATED, ACTIVATION" & _
				" , ADDRESS, CITY, STATE, REMARKS, BANKNAME, BANKACCOUNT" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "MillQuality" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, DIRT, MOISTURE, FFA, DOBI " & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE='" & TXT_FCCODE & "' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Building" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, BUILDGROUPCODE_DESC, ITEM, ACTIVATION" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Allowded" Then
		q	=	" SELECT" & _
				" FCCODE, FCBA, FCNAME, ALLDED_GROUP, ALLDED" & _
				" , CROP, CONTROL_JOB, MAN_AUTO, TYPE_1, TYPE_2" & _
				" , TYPE_3, TYPE_4, MIN_AMT, MAX_AMT, RATE_1" & _
				" , RATE_2, RATE_3, RATE_4, REMARKS" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Field" Then <!--Additional MENU FIELD Judul -->
		q	=	" SELECT" & _
				"  FCCODE, FCNAME, DIVISION, HECTARAGEPLANTED, TERRAINTYPE" & _
				" , SOILTYPE, OWNERSHIP, ACTIVATION, CROPTYPE, PLANTINGDATE" & _
				" , PLANTINGMATERIAL, TOTALSTAND, DX_BLOCK_STATUS, PREVSEMESTER_ABW, PRESSEMESTER_ABW" & _
				" , HARVESTINGBASED_ABW, MAIN_ROAD, COLLECTIONROAD, MAINDRAIN, COLLECTIONDRAIN" & _
				" , SUBSIDIARYDRAIN, COLLECTIONPOINT, HARVESTINGBRIDGE, PERMANENTBRIDGE, WOODBRIDGE" & _
				" , BUIST, OWLNEST" & _
				" , FIELD_INFO_01, FIELD_INFO_02, FIELD_INFO_03, FIELD_INFO_04, FIELD_INFO_05,ISACTIVE_HARVESTED, VALID_FROM" & _
				" , AREAL_LC, CADANGAN, OKUPASI, BIBITAN, PABRIK, EMPLASMEN, PARIT, JALAN, PRASARANA_LAINNYA" & _
				" , QUARRY, TANGGUL, ENCLAVE, TAMBANG, KONSERVASI, BUKIT_SUNGAI, NO_IMAGERY, LUAS_GROSS, HECTARAGEHARVESTED, PRODUCTIVESTAND, HARVESTINGSECTION, HARVESTINGSECTIONNUMBER, PLANTINGPOINT, FULLSTAND, Year,Month " & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM Field_Hist_Tab " & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (MONTH = '" & TXT_MONTH & "' OR MONTH='" & REQUEST("TXT_MONTH") & "')" & _
				" AND (YEAR = '" & TXT_YEAR & "' OR YEAR='" & REQUEST("TXT_YEAR") & "')" & _ 
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')" & _
				" ORDER BY FCCODE " 
	ElseIf Request("txttable") = "Station" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, JOBCODE" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Machine" Then
		q	=	" SELECT" & _
				" FCCODE, DESCRIPTION, FIXEDASSETCODE, OWNERSHIP, DATEPURCHASE" & _
				" , PURCHASECOST, STATION_STATIONCODE, ACTIVATION" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
    ElseIf Request("txttable") = "WorkingTime" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, TIMEIN, TIMEOUT, ISRESTIME, ISNEXTDAY" & _
				" , ADDIN_MINUTES, ADDOUT_MINUTES, ACTIVATION" & _
				" , FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
    ElseIf Request("txttable") = "WorkingTime_Template" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, GANGCODE, PERIOD," & _
				" D_2,WORKINGTIME_API.get_timeinout(D_2,FCBA) D2_TIME," & _
				" D_3,WORKINGTIME_API.get_timeinout(D_3,FCBA) D3_TIME," & _
				" D_4,WORKINGTIME_API.get_timeinout(D_4,FCBA) D4_TIME," & _
				" D_5,WORKINGTIME_API.get_timeinout(D_5,FCBA) D5_TIME," & _
				" D_6,WORKINGTIME_API.get_timeinout(D_6,FCBA) D6_TIME," & _
				" D_7,WORKINGTIME_API.get_timeinout(D_7,FCBA) D7_TIME," & _
				" D_1,WORKINGTIME_API.get_timeinout(D_1,FCBA) D1_TIME," & _
				" ACTIVATION,FCENTRY, FCEDIT, FCIP, FCBA FROM " & txtTable & "" & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Employee_WorkingTime" Then
		q	=	" SELECT" & _
				" EMPCODE , EMPLOYEE_API.get_EmployeeName(FCBA,EMPCODE) FCNAME, PERIOD_YEAR, PERIOD_MONTH," & _
				" D_1,WORKINGTIME_API.get_timeinout(D_1,FCBA) D1_TIME," & _
				" D_2,WORKINGTIME_API.get_timeinout(D_2,FCBA) D2_TIME," & _
				" D_3,WORKINGTIME_API.get_timeinout(D_3,FCBA) D3_TIME," & _
				" D_4,WORKINGTIME_API.get_timeinout(D_4,FCBA) D4_TIME," & _
				" D_5,WORKINGTIME_API.get_timeinout(D_5,FCBA) D5_TIME," & _
				" D_6,WORKINGTIME_API.get_timeinout(D_6,FCBA) D6_TIME," & _
				" D_7,WORKINGTIME_API.get_timeinout(D_7,FCBA) D7_TIME," & _
				" D_8,WORKINGTIME_API.get_timeinout(D_8,FCBA) D8_TIME," & _
				" D_9,WORKINGTIME_API.get_timeinout(D_9,FCBA) D9_TIME," & _
				" D_10,WORKINGTIME_API.get_timeinout(D_10,FCBA) D10_TIME," & _
				" D_11,WORKINGTIME_API.get_timeinout(D_11,FCBA) D11_TIME," & _
				" D_12,WORKINGTIME_API.get_timeinout(D_12,FCBA) D12_TIME," & _
				" D_13,WORKINGTIME_API.get_timeinout(D_13,FCBA) D13_TIME," & _
				" D_14,WORKINGTIME_API.get_timeinout(D_14,FCBA) D14_TIME," & _
				" D_15,WORKINGTIME_API.get_timeinout(D_15,FCBA) D15_TIME," & _
				" D_16,WORKINGTIME_API.get_timeinout(D_16,FCBA) D16_TIME," & _
				" D_17,WORKINGTIME_API.get_timeinout(D_17,FCBA) D17_TIME," & _
				" D_18,WORKINGTIME_API.get_timeinout(D_18,FCBA) D18_TIME," & _
				" D_19,WORKINGTIME_API.get_timeinout(D_19,FCBA) D19_TIME," & _
				" D_20,WORKINGTIME_API.get_timeinout(D_20,FCBA) D20_TIME," & _
				" D_21,WORKINGTIME_API.get_timeinout(D_21,FCBA) D21_TIME," & _
				" D_22,WORKINGTIME_API.get_timeinout(D_22,FCBA) D22_TIME," & _
				" D_23,WORKINGTIME_API.get_timeinout(D_23,FCBA) D23_TIME," & _
				" D_24,WORKINGTIME_API.get_timeinout(D_24,FCBA) D24_TIME," & _
				" D_25,WORKINGTIME_API.get_timeinout(D_25,FCBA) D25_TIME," & _
				" D_26,WORKINGTIME_API.get_timeinout(D_26,FCBA) D26_TIME," & _
				" D_27,WORKINGTIME_API.get_timeinout(D_27,FCBA) D27_TIME," & _
				" D_28,WORKINGTIME_API.get_timeinout(D_28,FCBA) D28_TIME," & _
				" D_29,WORKINGTIME_API.get_timeinout(D_29,FCBA) D29_TIME," & _
				" D_30,WORKINGTIME_API.get_timeinout(D_30,FCBA) D30_TIME," & _
				" D_31,WORKINGTIME_API.get_timeinout(D_31,FCBA) D31_TIME," & _				
				" ACTIVATION,FCENTRY, FCEDIT, FCIP, FCBA FROM Employee_WorkingTime " & _
				" WHERE FCBA = '" & Session("FCBA") & "'" & _
				" AND (EMPCODE LIKE '" & TXT_FCCODE & "%' OR EMPCODE LIKE '" & REQUEST("FCCODE") & "%')" & _
				" AND (EMPLOYEE_API.get_GangCode(FCBA,EMPCODE) = '" & GANGCODE & "' OR EMPLOYEE_API.get_GangCode(FCBA,EMPCODE)='" & REQUEST("GANGCODE") & "')" & _
				" AND (PERIOD_MONTH = '" & TXT_MONTH & "' OR PERIOD_MONTH='" & REQUEST("TXT_MONTH") & "')" & _
				" AND (PERIOD_YEAR = '" & TXT_YEAR & "' OR PERIOD_YEAR='" & REQUEST("TXT_YEAR") & "')"
				
				
	ElseIf Request("txttable") = "Roles" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, FCENTRY, FCEDIT, FCIP, 'XXX' DUMMY_1" & _
				" FROM ROLES_TAB" & _
				" WHERE (FCCODE like '" & TXT_FCCODE & "%' OR FCCODE='" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "Koperasi" Then
		q = " SELECT fccode" & _
		    "        , fcname" & _
		    "        , datecreated" & _
		    "        , registrationno" & _
		    "        , registrationdate" & _
		    "        , address" & _
		    "        , city" & _
		    "        , state" & _
		    "        , phone" & _
		    "        , fax" & _
		    "        , formco" & _
		    "        , typeco" & _
		    "        , groupco" & _
		    "        , sectorco" & _
		    "        , idno" & _
		    "        , iscertified" & _
		    "        , activation" & _
		    "        , fcentry" & _
		    "        , fcedit" & _
		    "        , fcip" & _
		    "        , fcba" & _
		    " FROM koperasi" & _
		    " WHERE fcba='" & session("fcBA") & "' AND (fccode = '" & TXT_FCCODE & "' OR fccode = '" & REQUEST("FCCODE") & "')"
	ElseIf Request("txttable") = "CUSTOMER" Then
		q = " SELECT " & _
		    " CUSTOMERCODE,  " & _
			" DESCRIPTION,   " & _
			" SUPPLIER,      " & _
			" ICNO,          " & _
			" BANK,          " & _
			" BANKACCOUNTNO, " & _
			" ADDRESS,       " & _
			" COMPANYNO,     " & _
			" FCENTRY,       " & _
			" FCEDIT,        " & _
			" FCIP,          " & _
			" FCBA,          " & _
			" OWN_ESTATE,    " & _
			" LASTUPDATE,    " & _
			" LASTTIME,      " & _
			" CUSTOMER_FCBA, " & _
			" CUSTFFBTYPE,   " & _
			" ISTRANSPORTER, " & _
			" STREET,        " & _
			" CITY,          " & _
			" PROVINCE,      " & _
			" POSTALCODE,    " & _
			" TELEPHONE,     " & _
			" FAX,           " & _
			" EMAIL,         " & _
			" ISACTIVE,      " & _
			" ISBLOCK,       " & _
			" CONTROLJOB,    " & _
			" BANK_OWNERSHIP " & _
		    " FROM CUSTOMER  " & _
		    " WHERE fcba='" & session("fcBA") & "' AND ((CUSTOMERCODE Like '" & TXT_FCCODE & "%' OR CUSTOMERCODE Like '" & REQUEST("CUSTOMERCODE") & "%') ) "
	Elseif Request("txttable")="BusinessUnit"  or Request("txttable")="Pattern" or Request("txttable")="Parameter" or Request("txttable")="Trigger_Control" or Request("txttable")="GCMaster" then
      q="select * from " & txtTable & " WHERE FCBA='" & SESSION("FCBA") & "'"   
   else
      q="select * from " & txtTable & " where (FCCODE='" & TXT_FCCODE & "' OR FCCODE='" & REQUEST("FCCODE") & "') or (UPPER(FCNAME) Like '%" & TXT_FCCODE & "%' OR UPPER(FCNAME) Like '%" & REQUEST("FCCODE") & "%') AND fcba='" & session("fcBA") & "'"
   end if
   
   'Response.Write Request("txttable") & "-" & q 
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr
   JumRec=1
   intRecordsPerPage = 15
	
	Dim intRecordsPerPage,intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 15
	
	rsGAD.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGAD.PageCount
	intRecords = rsGAD.RecordCount
	'Set our AbsolutePage property.
	if not rsGAD.EOF then
		If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  rsGAD.AbsolutePage = 1
		Else
		  rsGAD.AbsolutePage = Request("iPageNo")
		End If  
	end if
	if Request("iPageNo") = "" then
		nPage = 1
	Else
		nPage = Request("iPageNo")
	End if
	Dim cColor

   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <style>
   	/* Daftar TableMaster: sel ditulis ringkas (dulu <td align="left" valign="top">), perataan lewat CSS */
   	table.tm-list > tbody > tr > td { text-align: left; vertical-align: top; }
   </style>
   <table class="data tm-list" >
	<!--#Include file="Judul_Tabel_Show.Asp"-->
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-1
	  ' Percepatan: pembuka sel per kolom (berisi nama kolom untuk tooltip) disiapkan sekali,
	  ' bukan memanggil rsGAD(i).Name di setiap sel; isi sel per baris digabung lalu ditulis sekali.
	  ' Kolom teks digabung sebagai string; kolom tanggal/angka tetap ditulis lewat Response.Write
	  ' (sama seperti cara lama: tulis langsung isi rsGAD(i)) agar format tanggal/angka mengikuti locale Response seperti sebelumnya.
	  Dim tmOpen(), tmIsText(), tmFlds, tmK, tmBuf
	  If nCountofField - 4 >= 0 Then
		ReDim tmOpen(nCountofField - 4)
		ReDim tmIsText(nCountofField - 4)
		For tmK = 0 To nCountofField - 4
			tmOpen(tmK) = "<td title=""" & rsGAD(tmK).Name & """>"
			Select Case rsGAD(tmK).Type
				Case 8, 129, 130, 200, 201, 202, 203: tmIsText(tmK) = True
				Case Else: tmIsText(tmK) = False
			End Select
		Next
	  End If
	  Set tmFlds = rsGAD.Fields
	  no=1
	  While Not rsGAD.EOF 
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
		  <% if request("txtTable") = "EmployeeAllowdedDetail" then %>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&ALLOWDEDCODE=<%=rsGAD(1)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a></TD>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&ALLOWDEDCODE=<%=rsGAD(1)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></TD>
		  <% elseif request("txtTable") = "Employee_WorkingTime" then %>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&TXT_YEAR=<%=rsGAD(2)%>&TXT_MONTH=<%=rsGAD(3)%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a></TD>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&TXT_YEAR=<%=rsGAD(2)%>&TXT_MONTH=<%=rsGAD(3)%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></TD>
		  <% elseif request("txtTable") = "Field" then %>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&TXT_YEAR=<%=rsGAD(57)%>&TXT_MONTH=<%=rsGAD(58)%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a></TD>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&TXT_YEAR=<%=rsGAD(57)%>&TXT_MONTH=<%=rsGAD(58)%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></TD>
		  <% else %>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a></TD>
		  <td align="left" valign="top"><a href="TableMaster.asp?<%=crashId%>=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></TD>
		  <% end if%>
				<%
				if nCountofField - 4 >= 0 then
					tmBuf = ""
					for i = 0 to nCountofField - 4
						if tmIsText(i) then
							tmBuf = tmBuf & tmOpen(i) & tmFlds(i).Value & "</td>"
						else
							Response.Write tmBuf & tmOpen(i)
							Response.Write tmFlds(i)
							tmBuf = "</td>"
						end if
					next
					Response.Write tmBuf
				end if
				%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
	<tr>
	<td colspan="10">
	   <%'if intNumberOfPages<>1 then
		  'For i = 1 to intNumberOfPages
			'  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""tableMaster.asp?txtModule=" &_
			'  txtModule & "&iPageNo=" &  i & "&txtTable=" & request("txtTable")& "&txtField=" & request("txtField")  & _
			'  "&txtSearch=" & request("txtSearch") & _
			'  """>" & i & "  "
		  'Next
	   'End if%>
	   </a><br>
	   <%if Request("iPageNo") = "" then%>
			<%nPage = 1%>
	   <%Else%>
			<%nPage = Request("iPageNo")%>
	   <%End if%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page : <%=nPage%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%></td>
	</tr>
   </TABLE>

<%End if%>  
 
  <%'-------------akhir----------------------%>
  
 </div> </div>
 

  <!--#include file="bottom.jsp"-->


</body>
</html>
