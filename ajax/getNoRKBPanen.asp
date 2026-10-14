<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
txtYear = Request("txtYear")
txtMonth = Request("txtMonth") 
activitycode = Request("activitycode") 
unitkerja = Request("unitkerja") 
fieldcode = Request("fieldcode") 

transno = fcba & txtYear & txtMonth & activitycode & unitkerja & fieldcode

response.write "{"
response.write """transno"" : """ & transno &""""
' response.write """rate"" : """ & rate & ""","
' response.write """Description"" : """ & Description & """"
response.write "}"

%>