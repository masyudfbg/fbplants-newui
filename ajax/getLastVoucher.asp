<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transtype = Request("transtype")
bankcode = Request("bankcode")
result = ""

q =   	 " select max(vouchercode) vouchercode from paymentvoucher where fcba = '"& session("fcba") &"'" & _
		 " and transtype = upper('"& transtype &"') and bankcode = '"& bankcode &"'order by lastupdate desc"

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	lastvouchercode= rsrecurtrans("vouchercode")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """lastvouchercode"" : """ & lastvouchercode & """"
response.write "}"
%>