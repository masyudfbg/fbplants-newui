<%Option Explicit%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<%
	Dim txtTable,txtModule,q,nCountofField,NO,txtSearch,txtCode
	Dim i
	txtTable =Request("txtTable")
	txtModule=Request("txtModule")
	txtCode  =Request("txtCode")
	txtSearch=Request("txtSearch")
	if len(trim(txtCode)) > 0 then
	   if txtTable="Attendance_GAD" then
		   connStr.Execute("Delete from "&txtTable&" where Gangcode='" & txtCode & "' and fcBA='" & session("fcBA") & "'")
		   Connstr.commit()
	   Elseif txtTable="PurchaseOrderDetails" then
		   connStr.Execute("Delete from PurchaseOrderDetails where PoCode='" & txtCode & "' and fcBA='" & session("fcBA") & "'")
		   Connstr.commit()
		   connStr.Execute("Delete from PurchaseOrder where PoCode='" & txtCode & "' and fcBA='" & session("fcBA") & "'")
		   Connstr.commit()
	   else
		   connStr.Execute("Delete from "&txtTable&" where fcCode='" & txtCode & "' and fcBA='" & session("fcBA") & "'")
		   Connstr.commit()
	   end if
	end if
	response.Redirect("show.asp?txtModule=" &_
				  txtModule & "&txtTable=" & request("txtTable")& "&txtField=" & request("txtField")  & _
				  "&txtSearch=" & request("txtSearch") & "&iPageNo=" & request("iPageNo"))
%>
