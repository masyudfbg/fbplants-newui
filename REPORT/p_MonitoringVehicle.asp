<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay,v_txtDate, v_txtAfdeling, v_txtBlok, v_txtColor
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName
Dim oRS, oSQL

If Len(Trim(Request("txtDate"))) = 0 Then
	v_txtDate		=	date
else
	v_txtDate		=	Request("txtDate")
End If

v_txtMonth		=	month(v_txtDate)
v_txtDay		=	DAY(v_txtDate)
v_txtYear		=	year(v_txtDate)
	
'v_txtYear		=	Request("txtYear")
'v_txtMonth		=	Request("txtMonth")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")
v_txtFcba		=	Request("txtFcba")
v_txtColor		=	Request("txtColor")

Dim v_txtPersen : v_txtPersen = Request("txtPersen")
If Len(trim(v_txtPersen)) = 0 Then v_txtPersen = 0


strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

If v_txtYear = "" Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If v_txtMonth = "" Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strPusingan	=	" select vm.fcba,vm.vehiclecode,vh.fcname, get_vehiclegroup_name(vh.vehiclegroupcode) vehiclegroup,vh.make,vh.model,vh.engineno,vh.chasisno,vh.yearofmade,vh.yearofpurchase, " & _
				"		   vm.DAY_1,vm.Hours_1,vm.Unit_1,vm.Output_1,vm.DAY_2,vm.Hours_2,vm.Unit_2,vm.Output_2,vm.DAY_3,vm.Hours_3,vm.Unit_3,vm.Output_3,vm.DAY_4,vm.Hours_4,vm.Unit_4,vm.Output_4,vm.DAY_5,vm.Hours_5,vm.Unit_5,vm.Output_5, " & _
				"          vm.DAY_6,vm.Hours_6,vm.Unit_6,vm.Output_6,vm.DAY_7,vm.Hours_7,vm.Unit_7,vm.Output_7,vm.DAY_8,vm.Hours_8,vm.Unit_8,vm.Output_8,vm.DAY_9,vm.Hours_9,vm.Unit_9,vm.Output_9,vm.DAY_10,vm.Hours_10,vm.Unit_10,vm.Output_10, " & _
				"		   vm.DAY_11,vm.Hours_11,vm.Unit_11,vm.Output_11,vm.DAY_12,vm.Hours_12,vm.Unit_12,vm.Output_12,vm.DAY_13,vm.Hours_13,vm.Unit_13,vm.Output_13,vm.DAY_14,vm.Hours_14,vm.Unit_14,vm.Output_14,vm.DAY_15,vm.Hours_15,vm.Unit_15,vm.Output_15, " & _
				"          vm.DAY_16,vm.Hours_16,vm.Unit_16,vm.Output_16,vm.DAY_17,vm.Hours_17,vm.Unit_17,vm.Output_17,vm.DAY_18,vm.Hours_18,vm.Unit_18,vm.Output_18,vm.DAY_19,vm.Hours_19,vm.Unit_19,vm.Output_19,vm.DAY_20,vm.Hours_20,vm.Unit_20,vm.Output_20, " & _
				"		   vm.DAY_21,vm.Hours_21,vm.Unit_21,vm.Output_21,vm.DAY_22,vm.Hours_22,vm.Unit_22,vm.Output_22,vm.DAY_23,vm.Hours_23,vm.Unit_23,vm.Output_23,vm.DAY_24,vm.Hours_24,vm.Unit_24,vm.Output_24,vm.DAY_25,vm.Hours_25,vm.Unit_25,vm.Output_25, " &_
				"          vm.DAY_26,vm.Hours_26,vm.Unit_26,vm.Output_26,vm.DAY_27,vm.Hours_27,vm.Unit_27,vm.Output_27,vm.DAY_28,vm.Hours_28,vm.Unit_28,vm.Output_28,vm.DAY_29,vm.Hours_29,vm.Unit_29,vm.Output_29,vm.DAY_30,vm.Hours_30,vm.Unit_30,vm.Output_30,vm.DAY_31,vm.Hours_31,vm.Unit_31,vm.Output_31 " & _
				"	from (    select  fcba, " & _
				"			  vehiclecode, " & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),1,vehiclestatus,0)) DAY_1,   " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),1,Hours,0))         Hours_1, " & _ 
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),1,Unit,0))          Unit_1, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),1,Output,0))        Output_1," & _ 
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),2,vehiclestatus,0)) DAY_2,   " & _ 
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),2,Hours,0))         Hours_2, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),2,Unit,0))          Unit_2, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),2,Output,0))        Output_2," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),3,vehiclestatus,0)) DAY_3,   " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),3,Hours,0))         Hours_3, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),3,Unit,0))          Unit_3, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),3,Output,0))        Output_3," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),4,vehiclestatus,0)) DAY_4,   " & _ 
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),4,Hours,0))         Hours_4, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),4,Unit,0))          Unit_4, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),4,Output,0))        Output_4," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),5,vehiclestatus,0)) DAY_5,   " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),5,Hours,0))         Hours_5, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),5,Unit,0))          Unit_5, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),5,Output,0))        Output_5," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),6,vehiclestatus,0)) DAY_6,   " & _
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),6,Hours,0))         Hours_6, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),6,Unit,0))          Unit_6, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),6,Output,0))        Output_6," & _				
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),7,vehiclestatus,0)) DAY_7,   " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),7,Hours,0))         Hours_7, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),7,Unit,0))          Unit_7, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),7,Output,0))        Output_7," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),8,vehiclestatus,0)) DAY_8,   " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),8,Hours,0))         Hours_8, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),8,Unit,0))          Unit_8, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),8,Output,0))        Output_8," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),9,vehiclestatus,0)) DAY_9,   " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),9,Hours,0))         Hours_9, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),9,Unit,0))          Unit_9, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),9,Output,0))        Output_9," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),10,vehiclestatus,0)) DAY_10, " & _
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),10,Hours,0))         Hours_10, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),10,Unit,0))          Unit_10, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),10,Output,0))        Output_10," & _ 				
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),11,vehiclestatus,0)) DAY_11, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),11,Hours,0))         Hours_11, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),11,Unit,0))          Unit_11, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),11,Output,0))        Output_11," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),12,vehiclestatus,0)) DAY_12, " & _
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),12,Hours,0))         Hours_12, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),12,Unit,0))          Unit_12, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),12,Output,0))        Output_12," & _				
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),13,vehiclestatus,0)) DAY_13, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),13,Hours,0))         Hours_13, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),13,Unit,0))          Unit_13, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),13,Output,0))        Output_13," & _ 
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),14,vehiclestatus,0)) DAY_14, " & _
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),14,Hours,0))         Hours_14, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),14,Unit,0))          Unit_14, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),14,Output,0))        Output_14," & _				
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),15,vehiclestatus,0)) DAY_15, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),15,Hours,0))         Hours_15, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),15,Unit,0))          Unit_15, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),15,Output,0))        Output_15," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),16,vehiclestatus,0)) DAY_16, " & _
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),16,Hours,0))         Hours_16, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),16,Unit,0))          Unit_16, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),16,Output,0))        Output_16," & _				
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),17,vehiclestatus,0)) DAY_17, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),17,Hours,0))         Hours_17, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),17,Unit,0))          Unit_17, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),17,Output,0))        Output_17," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),18,vehiclestatus,0)) DAY_18, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),18,Hours,0))         Hours_18, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),18,Unit,0))          Unit_18, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),18,Output,0))        Output_18," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),19,vehiclestatus,0)) DAY_19, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),19,Hours,0))         Hours_19, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),19,Unit,0))          Unit_19, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),19,Output,0))        Output_19," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),20,vehiclestatus,0)) DAY_20, " & _ 
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),20,Hours,0))         Hours_20, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),20,Unit,0))          Unit_20, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),20,Output,0))        Output_20," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),21,vehiclestatus,0)) DAY_21, " & _ 
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),21,Hours,0))         Hours_21, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),21,Unit,0))          Unit_21, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),21,Output,0))        Output_21," & _				
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),22,vehiclestatus,0)) DAY_22, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),22,Hours,0))         Hours_22, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),22,Unit,0))          Unit_22, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),22,Output,0))        Output_22," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),23,vehiclestatus,0)) DAY_23, " & _  
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),23,Hours,0))         Hours_23, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),23,Unit,0))          Unit_23, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),23,Output,0))        Output_23," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),24,vehiclestatus,0)) DAY_24, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),24,Hours,0))         Hours_24, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),24,Unit,0))          Unit_24, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),24,Output,0))        Output_24," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),25,vehiclestatus,0)) DAY_25, " & _ 
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),25,Hours,0))         Hours_25, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),25,Unit,0))          Unit_25, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),25,Output,0))        Output_25," & _				
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),26,vehiclestatus,0)) DAY_26, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),26,Hours,0))         Hours_26, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),26,Unit,0))          Unit_26, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),26,Output,0))        Output_26," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),27,vehiclestatus,0)) DAY_27, " & _ 
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),27,Hours,0))         Hours_27, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),27,Unit,0))          Unit_27, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),27,Output,0))        Output_27," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),28,vehiclestatus,0)) DAY_28, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),28,Hours,0))         Hours_28, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),28,Unit,0))          Unit_28, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),28,Output,0))        Output_28," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),29,vehiclestatus,0)) DAY_29, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),29,Hours,0))         Hours_29, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),29,Unit,0))          Unit_29, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),29,Output,0))        Output_29," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),30,vehiclestatus,0)) DAY_30, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),30,Hours,0))         Hours_30, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),30,Unit,0))          Unit_30, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),30,Output,0))        Output_30," & _
				"			  MAX(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),31,vehiclestatus,0)) DAY_31,  " & _ 
                "			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),31,Hours,0))         Hours_31, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),31,Unit,0))          Unit_31, " & _
				"			  SUM(DECODE(TO_NUMBER(TO_CHAR(ONDATE, 'DD')),31,Output,0))        Output_31 " & _				
				"	from vehicle_monitoring_tab " & _  
				"	where to_number(to_char(ondate,'YYYY'))= " & v_txtYear & " " & _  
				"	and to_number(to_char(ondate,'MM'))=" & v_txtMonth & "" & _
				"	and fcba='" & v_txtFcba & "'" & _
				"	and to_number(TO_CHAR(ONDATE, 'DD')) <=  " &  v_txtDay  &"" & _  
				"	group by fcba, vehiclecode,to_char(ondate,'YYYY'),to_char(ondate,'MM') " & _  
				"	) vm , vehicle vh " & _  
				"	where  vm.fcba=vh.fcba " & _  
				"	and vh.fccode =vm.vehiclecode order by vh.vehiclegroupcode,vm.vehiclecode"
				

									
strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
Else
	varFcbaName = ""
	varCompanyName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

'response.write strPusingan
%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>ePlaS Reporting - Monitoring Vehicle</title>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
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
-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	 <!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">	
						<tr>
						<!--<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM
						mon = Month(Session("v_Costbook"))
						If len(v_txtMonth)=0 then v_txtMonth=mon

						Yr = Year(Session("v_Costbook"))
						MinYr = Yr - 2
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
					
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>-->
							<td>
								<font face="verdana" size="1">DATE</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDate" value="<%=v_txtDate%>" />
							</td>
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												" AND FCTYPE = 'E'" & _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value=""></option>
									<%
									End If
									
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							
							
						
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=110+GetLastDay(v_txtMonth, v_txtYear)%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>MONTHLY MONITORING VEHICLE</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				<h4>s/d Tgl&nbsp;:&nbsp;<%=v_txtDay%>&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;&nbsp;<%=v_txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
<% 
dim selec_day,totalcolumn
	if len(trim(DAY(v_txtDate))) = 0 then
		selec_day = GetLastDay(v_txtMonth, v_txtYear)
	else
		selec_day = DAY(v_txtDate)
	end if
	totalcolumn = cdbl(selec_day) * 4
%>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">Vehicle Group</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">Vehicle Code</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">Vehicle Name</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">Merk</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">Type</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">Year Of Made</th>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">Year Of Purchase</th>
		<th colspan="<%=totalcolumn%>" style="background-color:#00A300;color:#fff;">TANGGAL</th>
		<th colspan="4" style="background-color:#EDF4EB;color:#000000;">Total Hari</th>
		<th rowspan="3" style="background-color:#EDF4EB;color:#000000;">Down Time (%) </th>
	</tr>
	<tr>
	
	
<%
Dim rsGetColorTanggal, strGetColorTanggal
Dim val_FirstDay, val_LastDay
Dim val_Background

val_FirstDay = v_txtMonth & "/" & 1 & "/" & v_txtYear
val_LastDay = v_txtMonth & "/" & v_txtDay & "/" & v_txtYear

strGetColorTanggal =	" SELECT" & _
						"     FCBA" & _
						"     , TO_NUMBER(TO_CHAR(TANGGAL, 'DD')) LIST_DAY" & _
						"     , DECODE(" & _
						"         BUDGET_TRANSACTION_WORKDAY_API.GETQTYWORKINGDAY_HI(FCBA, TO_NUMBER(TO_CHAR(TANGGAL, 'YYYY')), 1, 1,TANGGAL)" & _
						"         , 0,'LIBUR', 'NORMAL') KETERANGAN" & _
						" FROM (" & _
						"     SELECT" & _
						"         FCCODE FCBA" & _
						"         , TANGGAL" & _
						"     FROM (" & _
						"         SELECT TO_DATE('" & val_FirstDay & "' , 'MM/DD/YYYY') + ROWNUM - 1 TANGGAL" & _
						"         FROM ALL_OBJECTS" & _
						"         WHERE ROWNUM <= TO_DATE('" & val_LastDay & "','MM/DD/YYYY') - TO_DATE('" & val_FirstDay & "' , 'MM/DD/YYYY') + 1" & _
						"         )" & _
						"     , BUSINESSUNIT" & _
						"     ) WHERE FCBA = '" & v_txtFcba & "' ORDER BY TO_NUMBER(TO_CHAR(TANGGAL, 'DD'))"
						
'response.write(strGetColorTanggal)
Set rsGetColorTanggal = Server.CreateObject("ADODB.Recordset")
rsGetColorTanggal.Open strGetColorTanggal, connstr, adOpenStatic

If Not rsGetColorTanggal.EOF Then

	While Not rsGetColorTanggal.EOF
		If rsGetColorTanggal("KETERANGAN") = "LIBUR" Then
			val_Background = "background-color:#e7dc14;color:#fff;"
		Else
			val_Background = "background-color:#00A300;color:#fff;"
		End If
%>
	<th colspan="4" style="<%=val_Background%>"><%=rsGetColorTanggal("LIST_DAY")%></th>
<%
	
		rsGetColorTanggal.MoveNext
	Wend 
		
End If
%>
      <th rowspan="2" style="background-color:#EDF4EB;color:#000000;">B</th>
      <th rowspan="2" style="background-color:#EDF4EB;color:#000000;">SB</th>
	  <th rowspan="2" style="background-color:#EDF4EB;color:#000000;">R</th>
	  <th rowspan="2" style="background-color:#EDF4EB;color:#000000;">Total</th>
	</tr>
	<tr>
	   <%
	   For xNum = 1 to selec_day 
	   %>
	   <th align="center"  style="background-color:#00A300;color:#fff;"><strong>Status</strong></th>
	   <th align="center"  style="background-color:#00A300;color:#fff;"><strong>Hours</strong></th>
	   <th align="center"  style="background-color:#00A300;color:#fff;"><strong>Unit</strong></th>
	   <th align="center"  style="background-color:#00A300;color:#fff;"><strong>Output</strong></th>
	   <%
	   Next
	   %>
	</tr>
    </thead>
	<tbody>
<%
Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic

Dim xNum

Dim Total_Ha : Total_Ha = 0
Dim Total_LastInterval : Total_LastInterval = 0
Dim DataIsPersen : DataIsPersen = 0
Dim CountData : CountData = 0


if rsPusingan.bof and rsPusingan.eof then

%>
	<tr>
		<td colspan="<%=48+v_txtDay%>">NO DATA AVAILABLE</td>
	</tr>
<%end if
      DIM B_,SB_,R_
	  DIM Rnday(31),nHours(31),nUnit(31),nOutput(31)
	  DIM TotB_, TotSB_, TotR_

		while not rsPusingan.eof
			CountData = CountData + 1
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>

			<tr class="<%=cClass%>">
				<td align="center"><%=rsPusingan("vehiclegroup")%></td>
				<td align="center"><%=rsPusingan("vehiclecode")%></td>
				<td align="center"><%=rsPusingan("fcname")%></td>
				<td align="center"><%=rsPusingan("make")%></td>
				<td align="center"><%=rsPusingan("model")%></td>
				<td align="center"><%=rsPusingan("yearofmade")%></td>
				<td align="center"><%=rsPusingan("yearofpurchase")%></td>
				
				<%
				
				B_  =0 
				SB_ =0
				R_  =0
				
				For xNum = 1 to selec_day 
				%>
				<td align="center"><strong><%=rsPusingan("day_"&xNum&"")%></strong></td>
				<%
				   If rsPusingan("day_"&xNum ) = "B" Then
				       B_ = B_ + 1
					   TotB_  =  cdbl(TotB_) + 1
				   ElseIf rsPusingan("day_"&xNum ) = "SB" Then
					   SB_ = SB_ + 1	
					   TotSB_  =  cdbl(TotSB_) + 1
                   ElseIf rsPusingan("day_"&xNum ) = "R" Then
					    R_ = R_ + 1
						TotR_  =  cdbl(TotR_) + 1
						Rnday(xNum) = cdbl(Rnday(xNum)) + 1
				   End If
                %>
                <td align="center"><strong><%=rsPusingan("hours_"&xNum&"")%></strong></td>  				
				<td align="center"><strong><%=rsPusingan("unit_"&xNum&"")%></strong></td> 
				<td align="center"><strong><%=rsPusingan("output_"&xNum&"")%></strong></td> 
                <% 	
                    nHours(xNum)  =  cdbl(nHours(xNum))  + cdbl(rsPusingan("hours_"&xNum&""))
					nUnit(xNum)   =  cdbl(nUnit(xNum))  + cdbl(rsPusingan("unit_"&xNum&""))
					nOutput(xNum) =   cdbl(nOutput(xNum))  + cdbl(rsPusingan("output_"&xNum&""))				
				Next
				%>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=B_%></strong></td>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=SB_%></strong></td>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=R_%></strong></td>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=Cint(B_)+Cint(SB_)+Cint(R_)%></strong></td>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=round(Cint(R_)/(Cint(B_)+Cint(SB_)+Cint(R_))*100,1)%></strong></td>
				
			</tr>
			
			
			

<%
			
			rsPusingan.MoveNext
		Wend
rsPusingan.Close



   IF cdbl(CountData) > 0 THEN
%>

            <tr class="<%=cClass%>">
				<td align="Center"  colspan=6 Rowspan=2 ">Total</td>	
				<td align="Right"  >H. Rusak</td>	
                <% For xNum = 1 to selec_day  %> 
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=cdbl(Rnday(xNum))%></strong></td>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=cdbl(nHours(xNum))%></strong></td>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=cdbl(nUnit(xNum))%></strong></td>
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=cdbl(nOutput(xNum))%></strong></td>
                <% Next %>
				
				<td align="Right" Rowspan=2  bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=TotB_%></strong></td>
				<td align="Right" Rowspan=2  bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=TotSB_%></strong></td>
				<td align="Right" Rowspan=2  bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=TotR_%></strong></td>
				<td align="Right" Rowspan=2  bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=Cint(TotB_)+Cint(TotSB_)+Cint(TotR_)%></strong></td>
				<td align="Right" Rowspan=2  bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=round(Cint(TotR_)/(Cint(TotB_)+Cint(TotSB_)+Cint(TotR_))*100,1)%></strong></td>
				
			</tr>
			<tr class="<%=cClass%>">
		
				<td align="Right"  >% D.Time</td>
                <% For xNum = 1 to selec_day  %> 
				<td align="Right" bgcolor="EDF4EB" style="color:"#ffffff";><strong><%=round(cdbl(Rnday(xNum))/cdbl(CountData) *100,1)%></strong></td>
				<td align="Right" colspan=3 bgcolor="EDF4EB" style="color:"#ffffff";><strong></strong></td>
                <% Next %>
				
				
				
			</tr>

   <% END IF %>
</tbody>


   <table width="90%" class="data">
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
	   <%
		DIM rsttd,ttd
		
		ttd = "SELECT FCMANAGER,FCKTU FROM BusinessUnit WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rsttd = Server.CreateObject("ADODB.Recordset")
		rsttd.Open ttd, connstr, adOpenStatic

		If Not rsttd.EOF Then
	   %>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCKTU")%></u></strong></td>
	   </tr>
	<%  End If
%> 
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
