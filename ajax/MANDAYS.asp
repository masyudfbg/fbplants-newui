<!--#include file="../connDb1.Inc"-->
<%

ajaxUoM = Request("ajaxUoM")

if ajaxUoM = "mandays" then

	Dim txtUoMmandays 
	
	response.expires=-1
	
	 stringmandays = " select PARAMETER_COMB_UOM_API.GetUom_Mandays('ATTENDANCE_GAD','" & request.querystring("q") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') tempUoM from dual "

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringmandays, connstr

						   

					  IF Not rsUoM.eof then
						 txtUoMmandays = rsUoM("tempUoM")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))


	 response.write(txtUoMmandays)


End if
if ajaxUoM = "unit" then
Dim txtUomUnit
	
	response.expires=-1
	
	 stringUnit = " select PARAMETER_COMB_UOM_API.GetUom_Unit('ATTENDANCE_GAD','" & request.querystring("q") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') Unit from dual "

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
	
	 stringOutput = " select PARAMETER_COMB_UOM_API.GetUom_Output('ATTENDANCE_GAD','" & request.querystring("q") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') qOutput from dual "

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
if ajaxUoM = "oth" then
 Dim txtUomOthrs
	
	response.expires=-1
	
	 stringOthrs = " select PARAMETER_COMB_UOM_API.GetUom_Othrs('ATTENDANCE_GAD','" & request.querystring("q") & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') Othrs from dual "

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open stringOthrs, connstr

					  IF Not rsUoM.eof then
						 txtUomOthrs = rsUoM("Othrs")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	m=ucase(request.querystring("q"))
	
	response.write(txtUomOthrs)
	

	
End if
if ajaxUoM = "jobname" then
 Dim txtUomJobname
	
	response.expires=-1
	
	 stringJobname = " select job_api.get_jobcode_name('" & request.querystring("q") & "','" & Session("fcBA") & "') qJobname  from dual  "

					  Set rsUoM = Server.CreateObject("ADODB.Recordset")

					  rsUoM.Open  stringJobname, connstr

					  IF Not rsUoM.eof then
						 txtUomJobname = rsUoM("qJobname")
					  End if

					  rsUoM.Close

					  set rsUoM = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))

	response.write("(" + txtUomJobname + ")")

	 
End if

%>

