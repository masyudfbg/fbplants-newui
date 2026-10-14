<!--#include file="../connDb1.Inc"-->
<%

dataModule 		= request("dataModule")
dataSubmodule 	= request("dataSubmodule")
dataDetail 		= request("dataDetail")

if dataModule = "" then
dataModule = "NULL"
end if
if dataSubmodule = "" then
dataSubmodule = "NULL"
end if
if dataDetail = "" then
dataDetail = "NULL"
end if

check = "SELECT COUNT (*) jml " & _
        " FROM DailyInputModul " & _
        "WHERE FCBA = '" & Session("fcBA") & "' " & _
        "  AND TO_CHAR(FDDATE, 'YYYYMMDD') = to_char(sysdate,'YYYYMMDD') " & _
        "  AND upper(MODULE) = upper('" & dataModule & "') " & _
        "  and upper(submodule) = upper('" & dataSubmodule & "') " & _
        "  and upper(detail) = upper('" & dataDetail & "') "

	Set rscheck = Server.CreateObject("ADODB.Recordset")
			rscheck.CursorType = 3
			rscheck.CursorLocation = 3 

			rscheck.Open check, connstr

			if Not rscheck.Eof then
			   jml = cdbl(rscheck("jml"))
			end if	   
if jml = 0 then
    q = " Insert Into DailyInputModul" & _
                               " (fddate,fcba,module,submodule,detail," & _
        " fcentry,fcedit,fcip,lastupdate,lasttime)" & _
                                " values ( sysdate " & _
                                " ,'" & Session("fcBA") & "'" & _
                                " ,'" & dataModule & "'" & _
                                " ,'" & dataSubmodule & "'" & _
                                " ,'" & dataDetail & "'" & _
                                " ,'" & Session("CodeName") & "'" & _
                                " ,'" & Session("CodeName") & "'" & _
                                " ,'" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
                                " ,sysdate, " & _
                                " To_Char(sysdate,'HH24:MI')" & _
                                " )"
	'RESPONSE.WRITE(Q)							
    connStr.execute(q)
	response.write("<script>alert('Terima Kasih Atas Kehadiran Anda ');window.history.go(-1);</script>")
else
	response.write("<script>alert('" & Session("CodeName") & " Sudah Hadir hari ini');window.history.go(-1);</script>")
end if	
%>
