<%Option Explicit%>
<%On Error Resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Harvesting</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
<script language="JavaScript" src="date-picker.js"></script> 
<script language="JavaScript"> 
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script> 
</head>
<body>
<!--#include file="top.jsp"-->
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%'-------------awal-----------------------%>
 <%
 Dim FCCODE
 Dim PeriodDate,Txt_Fieldcode
 Dim FCENTRY, FCEDIT, FCIP, FCBA
 Dim CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, RECEPTIONDATE, VEHICLE, OWN, QUANTITY, ESTATE_WEIGHT, MILL_WEIGHT
 Dim cmdView

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	PeriodDate =request("PeriodDate")
	Txt_Fieldcode =request("Txt_Fieldcode")
    CHITNO=request("CHITNO")
	FIELDCODE=request("FIELDCODE") 
	CROPCODE=request("CROPCODE")
	PRODUCTCODE=request("ProductCODE")
	RECEPTIONDATE=request("RECEPTIONDATE")
	VEHICLE=request("VEHICLE")
	OWN=request("OWN")
	QUANTITY=request("QUANTITY")
	ESTATE_WEIGHT=request("ESTATE_WEIGHT")
	MILL_WEIGHT=request("MILL_WEIGHT")
	if Request("txtJenis")= "Delete" then
       q="Delete from Harvesting where  CHITNO='" & CHITNO & "' And FIELDCODE='" & FIELDCODE & "' and fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	  q="Insert into Harvesting (CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, RECEPTIONDATE, VEHICLE, OWN, QUANTITY, ESTATE_WEIGHT, MILL_WEIGHT, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & CHITNO & "','" & FIELDCODE & "','" & CROPCODE & "','" & PRODUCTCODE & "'," & _
		" to_date('" & RECEPTIONDATE & "','mm/dd/yyyy') " & ",'"  &  VEHICLE & "','" & OWN & "','" & QUANTITY & "'," & _
		"'" & ESTATE_WEIGHT & "','" & MILL_WEIGHT & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
	    connStr.execute(q)
		'RESPONSE.Write(q)
		'RESPONSE.Flush()
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update Harvesting set CROPCODE ='" & CROPCODE & "',PRODUCTCODE='" & PRODUCTCODE & "'," & _
		"RECEPTIONDATE=" & " to_date('" & RECEPTIONDATE & "','mm/dd/yyyy') ," & _
		"VEHICLE ='" & VEHICLE & "',OWN='" & OWN & "'," & _
		"QUANTITY ='" & QUANTITY & "',ESTATE_WEIGHT ='" & ESTATE_WEIGHT & "',MILL_WEIGHT ='" & MILL_WEIGHT & "'" & _
        " where  CHITNO='" & CHITNO & "' And FIELDCODE='" & FIELDCODE & "' and fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		'RESPONSE.Write(q)
		'RESPONSE.Flush()
		connStr.execute("commit")
		'RESPONSE.Write "SQLnsddsdya : " & Q
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("PeriodDate")<>"" then
   PeriodDate   =Request("PeriodDate")
else
   PeriodDate   =textDate(date)
end if
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

 '-------------ISI Payment Form%>
 <form name="frmEntry" method="post" action="Harvesting.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
	  <TR> 
		<td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
			HARVESTING</strong></font></div></td>
	  </TR>

		<tr><td align="left">Periode</td>
		<td align="left"><input type="text" maxlength="10" name="PeriodDate" size="10" value="<%=PeriodDate%>">
		<a href="javascript:show_calendar('frmEntry.PeriodDate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<TR>
		<td align="left"><font face="Verdana" size="1">Field Code</font></td>
		<td align="left"> <select name="Txt_Fieldcode">
		<%cmdCallCombo "Field",request("Txt_Fieldcode"),"",""%>
		</select></td>
		</tr>		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintHarvesting.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&PeriodDate=<%=PeriodDate%>&Txt_Fieldcode=<%=Txt_Fieldcode%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select CHITNO, FIELDCODE, CROPCODE, ProductCODE, to_char(RECEPTIONDATE,'mm/dd/yyyy') as RECEPTIONDATE, VEHICLE, OWN, QUANTITY, " & _
     "ESTATE_WEIGHT, MILL_WEIGHT, FCENTRY, FCEDIT, FCIP, FCBA from Harvesting where fcBA='" & session("fcBA") & "'"
	if  len(trim(request("PeriodDate"))) > 0 then
		q = q & " and to_char(RECEPTIONDATE,'mm/yyyy')=to_char(to_date('" & request("PeriodDate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("Txt_Fieldcode"))) > 0 then
		q = q & " and FIELDCODE='" & request("Txt_Fieldcode") & "'"
	end if
	q = q
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   'response.Write(q)
   'response.Flush()
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="Harvesting.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">

		<tr><td align="left">Reception Date</td><td align="left">
		<%IF txtJenis="Edit" Then%>
		<input type="text" size="15" name="RECEPTIONDATE" value="<%=RECEPTIONDATE%>">
		<%Else%>
		<input type="text" size="15" name="RECEPTIONDATE" value="<%=DATE()%>">
		<%End if%>
		<a href="javascript:show_calendar('eForm.RECEPTIONDATE');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>

		<%IF txtJenis="Edit" Then%>
			 <tr><td align="left">Chit No</td><td align="left"><input type="text" size="15" name="CHITNO" value="<%=CHITNO%>"  readonly="True" style="background:#999999">  <font color="#FF0000">*</font></td></tr>
			</TR>
			<TR>
			<td align="left"><font face="Verdana" size="1">Field Code</font></td>
			<td align="left"> <select name="FIELDCODE">
			<%cmdCallCombo "Field",FIELDCODE,"",""%>
			</select> <font color="#FF0000">*</font> </td>
			</tr>
			 <tr>
				<td align="left"><font face="Verdana" size="1">Crop Code</font></td>
					<% 	Dim rsCropA
					q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" & Session("fcBA") & "' order by P.FCNAME desc "
					Set rsCropA = Server.CreateObject("ADODB.Recordset")
					rsCropA.CursorType = 3
					rsCropA.CursorLocation = 3
					rsCropA.Open q, connstr
					if not rsCropA.eof then%>
					<td align="left"> <select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
					<option value=""></option> 
					<%
					Do While Not rsCropA.eoF%>
					<%IF Request("CROPCODE")=rsCropA("FCCODE") Then%>
						<option value="<%=rsCropA("FCCODE")%>" selected><%=rsCropA("FCNAME")%></option> 
					<%Else%>
						<option value="<%=rsCropA("FCCODE")%>"><%=rsCropA("FCNAME")%></option> 
					<%End if%>
					<%rsCropA.movenext
					loop
					End if
					rsCropA.Close
					set rsCropA = Nothing
					q=""
					%>
					</select> <font color="#FF0000">*</font></td>

			 </tr>			
			 <tr>
				<td align="left"><font face="Verdana" size="1">Product Code</font></td>
				<td align="left"> <select name="ProductCODE">
				<%cmdCallComboSetup "PRODUCT",PRODUCTCODE,"",""%>
				</select> <font color="#FF0000">*</font> </td>
			 </tr>			
		<%else%>
 		     <tr>
			 <td align="left">Chit No</td><td align="left"><input type="text" size="15" name="CHITNO" value="<%=CHITNO%>"> <font color="#FF0000">*</font></td>
			 </tr>
			 <tr>
				<td align="left"><font face="Verdana" size="1">Field Code</font></td>
				<td align="left"> <select name="FIELDCODE">
				<%cmdCallCombo "Field",FIELDCODE,"",""%>
				</select> <font color="#FF0000">*</font> </td>
			 </tr>
			 <tr>
				<td align="left"><font face="Verdana" size="1">Crop Code</font></td>
					<% 	Dim rsCrop
					q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
					Set rsCrop = Server.CreateObject("ADODB.Recordset")
					rsCrop.CursorType = 3
					rsCrop.CursorLocation = 3
					rsCrop.Open q, connstr
					if not rsCrop.eof then%>
					<td align="left"> <select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
					<option value=""></option> 
					<%
					Do While Not rsCrop.eoF%>
					<%IF Request("CROPCODE")=rsCrop("FCCODE") Then%>
						<option value="<%=rsCrop("FCCODE")%>" selected><%=rsCrop("FCNAME")%></option> 
					<%Else%>
						<option value="<%=rsCrop("FCCODE")%>"><%=rsCrop("FCNAME")%></option> 
					<%End if%>
					<%rsCrop.movenext
					loop
					End if
					rsCrop.Close
					set rsCrop = Nothing
					q=""
					%>
					</select> <font color="#FF0000">*</font></td>
			 </tr>
			 <tr>
				<td align="left"><font face="Verdana" size="1">Product Code</font></td>
				<td align="left"> <select name="ProductCODE">
				<%cmdCallComboSetup "PRODUCT",PRODUCTCODE,"",""%>
				</select> <font color="#FF0000">*</font> </td>
			 </tr>			
			 
		<%End if%>
			<tr>
            <td align="left"><font face="Verdana" size="1">Vehicle</font></td><td align="left"><input type="text" size="15" maxlength="50" name="Vehicle" value="<%=Vehicle%>"> <font color="#FF0000">*</font></td>
			</tr>
		<%IF txtJenis="Edit" Then%>
		</tr>
            <td align="left"><font face="Verdana" size="1">Own</font></td>
            <td align="left"><select name="OWN" title="Own Character a s/d z or 1 s/d 9">
				<%if request("OWN")="01" Then%>
                <Option value="01" selected>OWN</Option>
                <Option value="02">CONTRACT</Option>
                <Option value="03">OTHERS</Option>
				<%Elseif request("OWN")="02" Then%>
                <Option value="01">OWN</Option>
                <Option value="02" selected>CONTRACT</Option>
                <Option value="03">OTHERS</Option>
				<%else%>
                <Option value="01">OWN</Option>
                <Option value="02">CONTRACT</Option>
                <Option value="03" selected>OTHERS</Option>
				<%End If%>
              </select></td>
          </Tr>
		  <%Else%>
		</tr>
            <td align="left"><font face="Verdana" size="1">Own</font></td>
            <td align="left"><select name="OWN" title="Own Character a s/d z or 1 s/d 9">
                <Option value="01">OWN</Option>
                <Option value="02">CONTRACT</Option>
                <Option value="03">OTHERS</Option>
              </select></td>
          </Tr>
		 <%End IF%>
		<tr><td align="left">Bunch</td><td align="left"><input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>"></td></tr>
		<tr><td align="left">Estate Weight (Kg)</td><td align="left"><input type="text" size="15" name="ESTATE_WEIGHT" value="<%=ESTATE_WEIGHT%>"></td></tr>
		<tr><td align="left">Mill Weight (Kg)</td><td align="left"><input type="text" size="15" name="MILL_WEIGHT" value="<%=MILL_WEIGHT%>"></td></tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="PeriodDate" value="<%=PeriodDate%>">
		  <input type="hidden" name="Txt_Fieldcode" value="<%=Txt_Fieldcode%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" style="background-color: #298CF7; color: #F7C610; font-weight:
		  bold; font-size: 8pt; font-family:Verdana" name="Reset" value="Reset">
		  <input Type="Button" style="background-color: #298CF7; color: #F7C610; font-weight: 
		  bold; font-size: 8pt; font-family:Verdana" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
 </form>
  <!--#include file="validation_show.Asp"--> 
   <table class="data">
	<BR>
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CHIT NO</TD>	
	<TD align="left">FILED CODE</TD>	
	<TD align="left">CROP CODE</TD>	
	<TD align="left">PRODUCT CODE</TD>	
	<TD align="left">RECEPTION DATE</TD>	
	<TD align="left">VEHICLE</TD>	
	<TD align="left">OWN</TD>	
	<TD align="left">BUNCH</TD>	
	<TD align="left">ESTATE WEIGHT</TD>	
	<TD align="left">MILL WEIGHT</TD>	
	</tr>
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="Harvesting.Asp?CHITNO=<%=rsVH("CHITNO")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&CROPCODE=<%=rsVH("CROPCODE")%>&PRODUCTCODE=<%=rsVH("PRODUCTCODE")%>&RECEPTIONDATE=<%=rsVH("RECEPTIONDATE")%>&VEHICLE=<%=rsVH("VEHICLE")%>&OWN=<%=rsVH("OWN")%>&QUANTITY=<%=rsVH("QUANTITY")%>&ESTATE_WEIGHT=<%=rsVH("ESTATE_WEIGHT")%>&MILL_WEIGHT=<%=rsVH("MILL_WEIGHT")%>&TXT_FIELDCODE=<%=rsVH("FIELDCODE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="30">
			<a href="Harvesting.Asp?CHITNO=<%=rsVH("CHITNO")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></td>
		    <%for i = 0 to nCountofField-4%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </TABLE>
  

<%End if%>  
 
  <%'-------------akhir----------------------%>
 </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>
