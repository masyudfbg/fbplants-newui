<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
result = ""

q = " SELECT * " & _
" FROM ASSETGROUP" & _
" WHERE fccode= '"& fccode &"' and fcba = '" & Session("fcba") & "'" 


'response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	result = rsrecurtrans("fcname")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing


response.write result
%>