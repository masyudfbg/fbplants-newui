<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
chitno = Request("chitno")
' transporter = Request("transporter")
result = ""

q = "   select wb.datein as datein, wb.dateout as dateout , wb.weightin as weightin, wb.weightout as weightout,nvl(gr.weight,0) as weight, (wb.weightin-wb.weightout-nvl(gr.weight,0)) as Netto " & _
	"	,wb.vehicleno as vehicleno" & _
	"	from weighbridgeticket wb" & _
	"	left join gradingrecord gr ON (wb.chitno = gr.wbt_chitno and wb.fcba = gr.fcba) " & _
 	"   WHERE wb.chitno= '"& chitno &"' and wb.fcba = '" & Session("fcba") & "'" 

' q = " SELECT dn.transportco, dn.postingdate, dn.quantity, wb.vehicleno " & _
' 	" FROM DELIVERYNOTE dn" & _
' 	" LEFT JOIN weighbridgeticket wb " & _
' 	" ON (dn.fcba = wb.fcba and dn.wbt_chitno = wb.chitno)" & _
' 	" WHERE dn.wbt_chitno= '"& wbt_chitno &"' and dn.fcba = '" & Session("fcba") & "'" 


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	' datein = rsrecurtrans("datein")
	' dateout = rsrecurtrans("dateout")
	' weightin = rsrecurtrans("weightin")
	' weightout = rsrecurtrans("weightout")
	' weight = rsrecurtrans("weight")
	' transporter = rsrecurtrans("transporter")
	Netto = rsrecurtrans("Netto")
	vehicleno = rsrecurtrans("vehicleno")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
' response.write """postingdate"" : """ & postingdate & ""","
' response.write """transporter"" : """ & transporter & ""","
response.write """vehicleno"" : """ & vehicleno & ""","
response.write """Netto"" : """ & Netto & """"
response.write "}"
%>