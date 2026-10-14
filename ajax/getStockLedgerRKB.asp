<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
stock = Request("stock")
txtMonth = Request("txtMonth")
txtYear = Request("txtYear")
 
q = "select round(amtperunit,2) amtperunit from stockledger where fcba = '"& fcba &"' and itemcode = '"& stock &"' and month = '"& txtMonth &"' and year = '"& txtYear &"'" 
 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	amtperunit= rsrecurtrans("amtperunit") 
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{" 
response.write """amtperunit"" : """ & amtperunit & """"
response.write "}"

%>