<!--#include file="../connDb1.Inc"-->
<%

ajaxUoM = Request("ajaxUoM")

if ajaxUoM = "hours" then

	Dim txtUoMhours
	
	response.expires=-1
	
	 stringhours = " select PARAMETER_COMB_UOM_API.GetUom_Hours('WORKSHOPACTIVITY','" & request.querystring("q") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') tempUoM from dual "

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringhours, connstr

						   

					  IF Not rsUoM.eof then
						 txtUoMhours = rsUoM("tempUoM")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))


	 response.write(txtUoMhours)


End if
if ajaxUoM = "unit" then
Dim txtUomUnit
	
	response.expires=-1
	
	 stringUnit = " select PARAMETER_COMB_UOM_API.GetUom_Unit('WORKSHOPACTIVITY','" & request.querystring("q") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') Unit from dual "

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringUnit, connstr

					  IF Not rsUoM.eof then
						 txtUomUnit = rsUoM("Unit")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))


	 response.write(txtUomUnit)
End if

if ajaxUoM = "output" then

	 Dim txtUomOutput
	
	response.expires=-1
	
	 stringOutput = " select PARAMETER_COMB_UOM_API.GetUom_Output('WORKSHOPACTIVITY','" & request.querystring("q") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') qOutput from dual "

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
End if


%>

