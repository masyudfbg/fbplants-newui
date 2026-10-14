<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
productcode = Request("productcode")
fcba = Request("fcba")
result = ""

q = " SELECT * " & _
" FROM PRODUCT" & _
" WHERE productcode= '"& productcode &"' and fcba = '" & fcba & "'" 


'response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	result = rsrecurtrans("unitofmeasure")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing


response.write result
%>