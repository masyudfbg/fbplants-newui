<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
txtYear = Request("txtYear")
txtMonth = Request("txtMonth")
activitycode = Request("activitycode")
divisi = Request("divisi")
fieldcode = Request("fieldcode")

transno = fcba & txtYear & txtMonth & activitycode & divisi & fieldcode

response.write "{"
response.write """transno"" : """ & transno &""""
' response.write """rate"" : """ & rate & ""","
' response.write """Description"" : """ & Description & """"
response.write "}"

%>