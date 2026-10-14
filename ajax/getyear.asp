<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fieldcode = Request("fieldcode")
txtYear = Request("year")
fcba = Request("fcba")

if txtYear = "" then
	txtYear = 0
end if

result = ""	
	Yr = Year(Session("v_Costbook"))
 
	html = html & "<select id='txtYear' name='txtYear' onchange='javascript:getselectedyear(); javascript:getFieldInfo(); javascript:getNoRKB(); javascript:getratematerial();' class='chosen-select' style='width:60%;'>"
	html = html &"<option value=''></option>"

					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(txtYear) = cint(xYr) then
						html = html & "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						html = html & "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next

	html = html &"</select><font color='#FF0000'>&nbsp;*</font>"
	response.write(html) 
%>