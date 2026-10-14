<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
agreementcode = Request("agreementcode")
fcba = Request("fcba")
lineno = Request("lineno")

arr_json_eof = true
q = " SELECT dtl.fcba" & _
	"        , dtl.agreementcode" & _
	"        , hdr.agreementdate" & _
	"        , hdr.contractorcode" & _
	"        , hdr.description" & _
	"        , hdr.retention" & _
	"        , hdr.startdate" & _
	"        , hdr.finishdate" & _
	"        , hdr.tax" & _
	"        , dtl.lineno" & _
	"        , dtl.jobcode" & _
	"        , job.fcname jobname" & _
	"        , dtl.qty" & _
	"        , dtl.rate" & _
	"        , dtl.reference" & _
	"        , dtl.remarks" & _
	" FROM contractagreement hdr" & _
	" INNER JOIN contractagreementdetail dtl ON (dtl.fcba = hdr.fcba AND dtl.agreementcode = hdr.agreementcode)" & _
	" LEFT JOIN job ON (dtl.fcba = job.fcba AND dtl.jobcode = job.fccode)" & _
	" WHERE dtl.agreementcode = '"&agreementcode&"'" & _
	"       AND dtl.fcba = '"&fcba&"'" & _
	"       AND dtl.lineno = '"&lineno&"'"


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
		response.write """agreementcode"" : """ & arr_json(1,xyyy) & ""","
		response.write """agreementdate"" : """ & arr_json(2,xyyy) & ""","
		response.write """contractorcode"" : """ & arr_json(3,xyyy) & ""","
		response.write """description"" : """ & arr_json(4,xyyy) & ""","
		response.write """retention"" : """ & arr_json(5,xyyy) & ""","
		response.write """startdate"" : """ & arr_json(6,xyyy) & ""","
		response.write """finishdate"" : """ & arr_json(7,xyyy) & ""","
		response.write """tax"" : """ & arr_json(8,xyyy) & ""","
		response.write """lineno"" : """ & arr_json(9,xyyy) & ""","
		response.write """jobcode"" : """ & arr_json(10,xyyy) & ""","
		response.write """jobname"" : """ & arr_json(11,xyyy) & ""","
		response.write """qty"" : """ & arr_json(12,xyyy) & ""","
		response.write """rate"" : """ & arr_json(13,xyyy) & ""","
		response.write """reference"" : """ & arr_json(14,xyyy) & ""","
		response.write """remarks"" : """ & arr_json(15,xyyy) & """"
	next
	response.write "}"
else
	response.write "{"
	response.write """agreementcode"" : """","
	response.write """agreementdate"" : """","
	response.write """contractorcode"" : """","
	response.write """description"" : """","
	response.write """retention"" : """","
	response.write """startdate"" : """","
	response.write """finishdate"" : """","
	response.write """tax"" : """","
	response.write """lineno"" : """","
	response.write """jobcode"" : """","
	response.write """jobname"" : """","
	response.write """qty"" : """","
	response.write """rate"" : """","
	response.write """reference"" : """","
	response.write """remarks"" : """""
	response.write "}"
end if

%>