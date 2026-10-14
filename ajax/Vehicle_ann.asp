<!--#include file="../connDb1.Inc"-->
<%
ajaxUoM = Request("ajaxUoM")	
FcCode	= Request("FcCode")

if ajaxUoM = "Check_Vehicle" then 
      
	  Dim txtUomCheck
	
	response.expires=-1
	

	
	 stringCheck = " select vehicle_api.get_IsVehicleTrue ('" & FcCode &"','" & Session("fcBA") & "') checker from dual"

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringCheck, connstr

					  IF Not rsUoM.eof then
						 txtUomCheck = rsUoM("checker")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
	 				  
	 response.write(txtUomCheck)

end if

if ajaxUoM = "hours" then 
      
	  Dim txtUomHours
	
	response.expires=-1
	

	
	 stringHours = " select PARAMETER_COMB_UOM_API.GetUom_Hours ('VEHICLEACTIVITY'," & _
						"vehicle_api.get_vehiclegroupcode('" & FcCode &"','" & Session("fcBA") & "'),'VEHICLEGROUPCODE','" & _ 
						"" & request.querystring("q") & "','JOBCODE',NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') qHours from dual"

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringHours, connstr

					  IF Not rsUoM.eof then
						 txtUomHours = rsUoM("qHours")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))
	 response.write(txtUomHours)

end if

if ajaxUoM = "unit" then 
      
	  Dim txtUomUnit
	
	response.expires=-1
	

	
	 stringUnit = " select PARAMETER_COMB_UOM_API.GetUom_Unit ('VEHICLEACTIVITY'," & _
						"vehicle_api.get_vehiclegroupcode('" & FcCode &"','" & Session("fcBA") & "'),'VEHICLEGROUPCODE','" & _ 
						"" & request.querystring("q") & "','JOBCODE',NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') qUnit from dual"

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringUnit, connstr

					  IF Not rsUoM.eof then
						 txtUomUnit = rsUoM("qUnit")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))
	 response.write(txtUomUnit)

end if

if ajaxUoM = "output" then 
      
	  Dim txtUomOutput
	
	response.expires=-1
	

	
	 stringOutput = " select PARAMETER_COMB_UOM_API.GetUom_Output ('VEHICLEACTIVITY'," & _
						"vehicle_api.get_vehiclegroupcode('" & FcCode &"','" & Session("fcBA") & "'),'VEHICLEGROUPCODE','" & _ 
						"" & request.querystring("q") & "','JOBCODE',NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') qOutput from dual"

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringOutput, connstr

					  IF Not rsUoM.eof then
						 txtUomOutput = rsUoM("qOutput")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))
	 response.write(txtUomOutput)

end if
%>

