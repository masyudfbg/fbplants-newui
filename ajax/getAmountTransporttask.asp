<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
orderdate = Request("orderdate")
storecode = Request("storecode")
itemcode = Request("itemcode")

arr_json_eof = true
' q = "SELECT STOCKLEDGER_API.GetCurrentPrice(to_date('" & orderdate & "','mm/dd/yyyy'), '" & storecode  & "', '" & itemcode  & "','" & Session("fcBA") & "') price FROM DUAL" 

q = " SELECT storecode," & _
	"        itemcode," & _
	"        fcba," & _
	"        case when NVL(amtperunit,0)=0 then " & _
	" 		 	case when NVL(openingbalanceqty,0) = 0 then 0 " & _
	"        	else (openingbalanceamt / openingbalanceqty) " & _
	"        	end " & _ 
    "		 else " & _ 
    "			amtperunit " & _ 
    "		 end " & _	
	"  price" & _
	"  FROM stockledger" & _
	"  WHERE     fcba = '"& session("fcba") &"'" & _
	"        AND storecode = '"& storecode &"'" & _
	"        AND itemcode = '"& itemcode &"'" & _
	"        AND year =" & _
	"               TO_NUMBER (" & _
	"                  TO_CHAR (TO_DATE ('"& orderdate &"', 'mm/dd/yyyy'), 'yyyy'))" & _
	"        AND month =" & _
	"               TO_NUMBER (" & _
	"                  TO_CHAR (TO_DATE ('"& orderdate &"', 'mm/dd/yyyy'), 'mm'))" 

' response.write q
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
		' response.write """podate"" : """ & arr_json(0,xyyy) & ""","
		' response.write """suppliercode"" : """ & arr_json(1,xyyy) & ""","
		' response.write """dp"" : """ & arr_json(2,xyyy) & ""","
		' response.write """ppn"" : """ & arr_json(3,xyyy) & ""","
		response.write """price"" : """ & arr_json(3,xyyy) & """"
	next
	response.write "}"
else
	response.write "{"
	' response.write """podate"" : """","
	' response.write """suppliercode"" : """","
	' response.write """dp"" : """","
	' response.write """ppn"" : ""0"","
	response.write """price"" : """""
	response.write "}"
end if

%>