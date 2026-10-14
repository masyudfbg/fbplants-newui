<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%		
	txtMonth = Request("month")
	
	if txtMonth = "" then
		txtMonth = 0
	end if

	Yr = Year(Session("v_Costbook"))
	
	html = html & "<select id='txtMonth' name='txtMonth' onchange='javascript:getselectedmonth(); javascript:getFieldInfo(); javascript:getNoRKB(); javascript:getratematerial();' class='chosen-select' style='width:60%;'>"
	html = html &"<option value=''></option>"
					for xMn=1 to 12
					if cint(txtMonth) = cint(xMn) then
						html = html & "<option value ='"&xMn&"' selected>"&Bulan(xMn)&"</option>"
					else
						html = html & "<option value ='"&xMn&"'>"&Bulan(xMn)&"</option>"
					end if
					next
				
	html = html &"</select><font color='#FF0000'>&nbsp;*</font>"
	response.write(html) 
%>