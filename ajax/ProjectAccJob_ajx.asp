<!--#include file="../connDb1.Inc"-->
<%
	Dim PROJECTGROUPCODE,initialFcba
	PROJECTGROUPCODE = request("PROJECTGROUPCODE")
	initialFcba = request("initialFcba")
	
	strAvailable =	" select CONTROLJOBCODE " & _
	                " from projectaccgroup" & _
					" where FCCODE = '" & PROJECTGROUPCODE & "'" & _
					"       and fcba = '" & initialFcba & "'"
	'response.write(strAvailable)			
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("CONTROLJOBCODE"))
	Else
		initialAvailable = ""
	End If
	
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CONTROLJOBCODE = initialAvailable


	response.write(CONTROLJOBCODE)
	
%>	