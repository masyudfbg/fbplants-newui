<!--#include file="txtFunction.inc"-->
<%
dim firstdate_, lastdate_

firstdate_ = textDate( Month(Session("v_Employee")) & "/01/" & Year(Session("v_Employee")) )
lastdate_ = textDate( Month(Session("v_Employee")) & "/" & GetLastDay(Month(Session("v_Employee")), Year(Session("v_Employee"))) & "/" & Year(Session("v_Employee")) )

response.write "<br/> Hari Pertama: " & firstdate_
response.write "<br/> Hari Terakhir: " & lastdate_

%>