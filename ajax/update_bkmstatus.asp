<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<%
dim txt_INDAY,txt_FCBA,gangcode

gangcode = Request("gangcode")
fddate = Request("fddate")
txt_FCBA = Request("FCBA")
txt_exec = Request("exec")
Dim txtajx
	
	' if txt_exec = "doit" then
	'   connstr.BeginTrans
	' 	' CREATE EXEC CLOSING
	' 	oQuery	= 	" UPDATE attendance_gad SET rowstate ='Planned' WHERE fcba = '"& txt_FCBA &"' AND gangcode='"& gangcode &"' AND fddate = TO_DATE('" & txt_INDAY & "','MM/DD/YYYY') "
					
	' 	'Response.Write oQuery & "<br/>"
	' 	connStr.execute(oQuery)
	' 	If err.number = 0 Then
	' 		connStr.CommitTrans
	' 	End If
	' 	If err.number <> 0 Then
	' 		connStr.RollbackTrans
	' 	End If
		
	' end if
	
	if txt_exec = "doupdate" then
	  connstr.BeginTrans
		' CREATE EXEC CLOSING
		oQuery	= 	" DELETE ag_authorization_comments WHERE fcba = '"& txt_FCBA &"' AND gangcode='"& gangcode &"' AND fddate = TO_DATE('" & fddate & "','MM/DD/YYYY') "
					
		'Response.Write oQuery & "<br/>"
		connStr.execute(oQuery)
		
		oQuery	= 	" DELETE  attendance_gad_hist WHERE fcba = '"& txt_FCBA &"' AND gangcode='"& gangcode &"' AND fddate = TO_DATE('" & fddate & "','MM/DD/YYYY') "
					
		'Response.Write oQuery & "<br/>"
		connStr.execute(oQuery)

		oQuery	= 	" DELETE  attendance_gad WHERE fcba = '"& txt_FCBA &"' AND gangcode='"& gangcode &"' AND fddate = TO_DATE('" & fddate & "','MM/DD/YYYY') "
					
		'Response.Write oQuery & "<br/>"
		connStr.execute(oQuery)

		oQuery	= 	" UPDATE attendance_gad_temp SET rowstate ='Planned' WHERE fcba = '"& txt_FCBA &"' AND gangcode='"& gangcode &"' AND fddate = TO_DATE('" & fddate & "','MM/DD/YYYY') "
					
		'Response.Write oQuery & "<br/>"
		connStr.execute(oQuery)

		If err.number = 0 Then
			connStr.CommitTrans
		End If
		If err.number <> 0 Then
			connStr.RollbackTrans
		End If
		
	end if
txt_exec = ""
	
	 SQLata = "SELECT distinct case when rowstate <> 'Planned' then 'TRUE' else 'FALSE' end cek FROM attendance_gad WHERE fcba = '"& txt_FCBA &"' AND gangcode='"& gangcode &"' AND fddate = TO_DATE('" & fddate & "','MM/DD/YYYY') "
	 ' response.write(SQLata)
		Set rsQ = Server.CreateObject("ADODB.Recordset")
		rsQ.Open SQLata, connstr			   

		IF Not rsQ.eof then
			txtajx = rsQ("cek")
		else
			txtajx = "Not Found."
		End if
		
		rsQ.Close
		set rsQ = Nothing  
					  
	response.expires=-1
	
	
	'Output "no suggestion" if no hint were found
	'or output the correct values

	if UCase(txtajx) = "TRUE" then
		response.write("Approved")
	else 
		response.write("Not Yet Approved")
	End if
	
%>