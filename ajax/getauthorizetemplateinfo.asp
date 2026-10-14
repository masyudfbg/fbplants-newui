<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
authorizetemplatecode = Request("authorizetemplatecode")

arr_json_eof = true
q = " SELECT aut.fccode" & _
	"        , aut.fcname" & _
	"        , aut.document_type" & _
	"        , aut.document_classification" & _
	"        , aut.authorizebyamount" & _
	"        , aut.amountrangefrom" & _
	"        , aut.amountrangeuntil" & _
	"        , aut.description" & _
	"        , aut.isactive" & _
	"        , dt.fcname document_type_name" & _
	"        , dc.fcname document_classification_name" & _
	" FROM authorizetemplate aut" & _
	" LEFT JOIN (select distinct fccode, fcname from document_classification) dc ON aut.document_classification = dc.fccode" & _
	" LEFT JOIN document_type dt ON aut.document_type = dt.fccode" & _
	" WHERE aut.fccode = '"&authorizetemplatecode&"'" & _
	" ORDER BY aut.fcname "

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
		response.write """fccode"" : """ & arr_json(0,xyyy) & ""","
		response.write """fcname"" : """ & arr_json(1,xyyy) & ""","
		response.write """document_type"" : """ & arr_json(2,xyyy) & ""","
		response.write """document_classification"" : """ & arr_json(3,xyyy) & ""","
		response.write """authorizebyamount"" : """ & arr_json(4,xyyy) & ""","
		response.write """amountrangefrom"" : """ & arr_json(5,xyyy) & ""","
		response.write """amountrangeuntil"" : """ & arr_json(6,xyyy) & ""","
		response.write """description"" : """ & arr_json(7,xyyy) & ""","
		response.write """isactive"" : """ & arr_json(8,xyyy) & ""","
		response.write """document_type_name"" : """ & arr_json(9,xyyy) & ""","
		response.write """document_classification_name"" : """ & arr_json(10,xyyy) & """"
	next
	response.write "}"
end if

%>