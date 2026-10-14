<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim RF_username, RF_ipaddress, RF_fcba, RF_date_act, RF_object_name
	
	RF_username = Request("RF_username")
	RF_ipaddress = Request("RF_ipaddress")
	RF_fcba = Request("RF_fcba")
	RF_date_act = Request("RF_date_act")
	RF_object_name = Request("RF_object_name")
	
	Dim sql_inserthistorytable
	sql_inserthistorytable = " INSERT INTO HISTORICAL_USER_TAB ( USERNAME" & _
							 "                                   , IP_ADDRESS" & _
							 "                                   , BU" & _
							 "                                   , DATE_ACT" & _
							 "                                   , TIME_ACT" & _
							 "                                   , LASTUPDATE" & _
							 "                                   , LASTTIME" & _
							 "                                   , DATE_ACT_FORM" & _
							 "                                   , OBJECT_NAME)" & _
							 " VALUES ( '" & RF_username & "'" & _
							 "          , '" & RF_ipaddress & "'" & _
							 "          , '" & RF_fcba & "'" & _
							 "          , '" & RF_date_act & "'" & _
							 "          , TO_CHAR(SYSDATE,'HH24:MI')" & _
							 "          , SYSDATE" & _
							 "          , TO_CHAR(SYSDATE,'HH24:MI')" & _
							 "          , SYSDATE" & _
							 "          , '" & RF_object_name & "')"
	Response.Write sql_inserthistorytable								
	connstr.execute(sql_inserthistorytable)
%>