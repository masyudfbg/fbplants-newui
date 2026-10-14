<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%		
	fcba = Request("fcba")
	activitycode = Request("activity")
	
	html = html & "<select id='activitycode' name='activitycode' onchange='javascript:getselectedactivity(); javascript:gettargetcode(); javascript:getNormaHK(); javascript:getNoRKB();' class='chosen-select' style='width:80%;'>"
	html = html &"<option value=''></option>"

					strsectionname	=	" select b.fccode,b.fcname from activity_norm a inner join jobmaster b on (a.job=b.fccode) " & _ 
					                    " where job_category in ('PN','MN','FF') and a.fcba = '" & fcba & "'" & _ 
										" order by fccode asc"
					
					Set rssectionname = Server.CreateObject("ADODB.Recordset")
					rssectionname.Open strsectionname, connstr, adOpenStatic 
					
					if not rssectionname.eof then
						Do While Not rssectionname.eof
							if cstr(""&activitycode) = cstr(""&rssectionname("fccode")) then
								html = html & "<option value ='"&rssectionname("fccode")&"' selected>"&rssectionname("fccode")&": "&rssectionname("fcname")&"</option>"
							else
								html = html & "<option value ='"&rssectionname("fccode")&"'>"&rssectionname("fccode")&": "&rssectionname("fcname")&"</option>"
							end if
							rssectionname.movenext
						loop
					end if
				
	html = html &"</select><font color='#FF0000'>&nbsp;*</font>"
	response.write(html) 
%>