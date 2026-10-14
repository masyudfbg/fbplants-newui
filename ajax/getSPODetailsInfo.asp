<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
businessunit = Request("businessunit")
spolineno = Request("spolineno")

arr_json_eof = true
q = " SELECT fcba, pocode, stockname, quantity, unitprice, amount, othercharge, amountcharge" & _
	" FROM servicepurchaseorderdetails" & _
	" WHERE pocode = '"&fccode&"' AND fcba = '"&businessunit&"' AND linenokey = '"&spolineno&"'"

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	arr_json = rsrecurtrans.getrows()
	arr_json_eof = false
end if
rsrecurtrans.Close : set rsrecurtrans = Nothing

if not arr_json_eof then
	response.write "{"
	for xyyy=0 to ubound(arr_json,2)
		response.write """descriptionx"" : """ & arr_json(2,xyyy) & ""","
		response.write """amount"" : """ & cdbl(to_nol(arr_json(5,xyyy))) & ""","
		response.write """othercharge"" : """ & cdbl(to_nol(arr_json(6,xyyy))) & ""","
		response.write """amountcharge"" : """ & cdbl(to_nol(arr_json(7,xyyy))) & """"
	next
	response.write "}"
else
	response.write "{"
	response.write """descriptionx"" : """","
	response.write """amount"" : ""0"","
	response.write """othercharge"" : ""0"","
	response.write """amountcharge"" : ""0"""
	response.write "}"
end if

%>