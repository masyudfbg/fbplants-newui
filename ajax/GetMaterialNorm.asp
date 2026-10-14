<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
stock = Request("stock")
activitycode = Request("activitycode")
 
'q = "select NORM.actmatnorm "&_
'	" from "&_
'	" stock "&_
'	" inner join ACTIVEMATERIAL AM ON stock.ACTIVEMATERIAL = AM.FCCODE "&_
'	" inner join MATERIALACTIVITY_NORM NORM ON AM.FCCODE = NORM.ACTMAT AND NORM.FCBA = stock.FCBA "&_
'	" inner join job on NORM.JOB = job.fccode and NORM.fcba = job.fcba "&_
'	" Where  "&_
'	" stock.FCBA = '" & fcba & "' "&_
'	" and stock.fccode = '" & stock & "' "&_
'	" and job.fccode ='" & activitycode & "' "
 
q = "Select actmatnorm " & _
	" from "&_
	" MATERIALACTIVITY_NORM "&_
	" Where  "&_
	" FCBA = '" & fcba & "' "&_
	" and actmat = '" & stock & "' "&_
	" and job ='" & activitycode & "' "


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	materialnorm= rsrecurtrans("actmatnorm") 
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{" 
response.write """materialnorm"" : """ & materialnorm & """"
response.write "}"

%>