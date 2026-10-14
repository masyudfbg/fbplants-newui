<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
prcode = Request("prcode")
chitnoharvesting = Request("chitnoharvesting")
result = ""

arr_json_eof = true
q = " SELECT bjr_per_chitno, kg_per_bucket FROM wb_harvesting WHERE fcba_estate = '"& session("fcba") &"' AND spbno = '"& chitnoharvesting &"'"  
' response.Write q

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
		response.write """bjr_per_chitno"" : """ & formatnumber(arr_json(0,xyyy),2) & ""","
		response.write """kg_per_bucket"" : """ & formatnumber(arr_json(1,xyyy),2) & """"
	next
	response.write "}"
else 
response.write "{"
response.write """bjr_per_chitno"" : "","
response.write """kg_per_bucket"" : """
response.write "}"

end if

%>