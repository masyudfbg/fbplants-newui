
<script>
function clickNew(val)
    {
    var docnumb_clr = document.getElementById("DOCNUMBER");
	var docdate_clr = document.getElementById("DOCDATE");
	var strge_clr = document.getElementById("PSTORAGE_STORAGECODE");
	var Add_btn = document.getElementById("Add_btn")
	var Add_OPEN_btn = document.getElementById("Add_OPEN_btn")
		docnumb_clr.value = "(NEW)"
		Add_btn.style = "display:none;"
		Add_OPEN_btn.style = "display:none;"
		docdate_clr.value = ""

}
</script>

<% 
dim form_direct,FFA_QUALITY,MOISTURE_QUALITY,DIRT_QUALITY
dim PSTORAGE_STORAGECODE,q,q2,docdate,SOUNDINGTYPE,DOCNUMBER,ROWSTATEID,SOUNDINGSTEP,HIGHLEVELTANK,TEMPERATURE,QUANTITY,txtAdd,Add_OPEN


NORMALSOUNDINGNUMBER = Request("NORMALSOUNDINGNUMBER")
FFA_QUALITY = Request("FFA_QUALITY")
MOISTURE_QUALITY = Request("MOISTURE_QUALITY")
DIRT_QUALITY = Request("DIRT_QUALITY")

TRANSFERSOUNDINGNUMBER = Request("TRANSFERSOUNDINGNUMBER")
SALESSOUNDINGNUMBER = Request("SALESSOUNDINGNUMBER")
PSTORAGE_STORAGECODE = Request("PSTORAGE_STORAGECODE")
DOCDATE = Request("DOCDATE")
SOUNDINGTYPE = Request("SOUNDINGTYPE")
DOCNUMBER = Request("DOCNUMBER")
ROWSTATEID = Request("ROWSTATEID")
SOUNDINGSTEP = Request("SOUNDINGSTEP")
HIGHLEVELTANK = Request("HIGHLEVELTANK")
TEMPERATURE = Request("TEMPERATURE")
QUANTITY = Request("QUANTITY")
txtAdd = Request("Add")
txtJenis =  Request("txtJenis")
cmdView = Request("cmdView")
Add_OPEN = Request("Add_OPEN")

if txtModule = "Mill_Production" and txtTable = "Sounding_Normal" then 
	if Len(NORMALSOUNDINGNUMBER) > 0 and NORMALSOUNDINGNUMBER <> "(NEW)" and txtJenis = "Edit" then 
		form_direct = "sounding_normal.Asp?txtModule=Mill_Production&txttable=Sounding_Normal&txtGetSounding=GetSounding1st&NORMALSOUNDINGNUMBER="&NORMALSOUNDINGNUMBER&"&DOCDATE="&DOCDATE&"&PSTORAGE_STORAGECODE="&PSTORAGE_STORAGECODE&"&FFA_QUALITY="&FFA_QUALITY&"&MOISTURE_QUALITY="&MOISTURE_QUALITY&"&DIRT_QUALITY="&DIRT_QUALITY&"&formwindow=open"
	else
		form_direct = "sounding_normal.Asp?txtModule=Mill_Production&txttable=Sounding_Normal&txtGetSounding=GetSounding1st&formwindow=open"
	end if
else
form_direct = "sounding.asp"
end if


if NORMALSOUNDINGNUMBER = "" then 
	NORMALSOUNDINGNUMBER = "(NEW)"
else
	NORMALSOUNDINGNUMBER = Request("NORMALSOUNDINGNUMBER")
end if	
if DOCNUMBER = "" then 
	DOCNUMBER = "(NEW)"
	else
	DOCNUMBER = Request("DOCNUMBER")
end if	
if DOCNUMBER = "(NEW)" and len(trim(PSTORAGE_STORAGECODE)) > 0 and len(trim(DOCDATE)) > 0 then
txtJenis = "New"
end if
if txtJenis = "" and Add_OPEN = "Add_OPEN" then
txtJenis = "New"
end if


'response.write(Add_OPEN)

if txtJenis = "New" or txtJenis ="New" or txtJenis = "Delete" or txtJenis = "Edit" or txtJenis = "Header" or Add_OPEN = "Add_OPEN" Then 

	if txtJenis = "Header" then
	
					connstr.BeginTrans
					
					dim txtDOCNUMBER
					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.Open "Select DOCNUMBER,FCBA from DOCSOUNDINGDETAIL where DOCNUMBER='" & DOCNUMBER & "' and FCBA = '" & Session("fcBA") & "'" & _
					" and fcBA='" & Session("fcBA") & "'", connstr
					IF Not rs.eof then %>
					<script>
					alert("Please Check Your Data");
					</script>
				<%	else
					q = " delete from DOCSOUNDING" & _
						" where DOCNUMBER = '" & DOCNUMBER & "'" & _
						"       and pstorage_storagecode='" & PSTORAGE_STORAGECODE & "'" & _
						"       and fcba = '" & Session("fcBA") & "'"
				   'response.Write(q)
				   connStr.execute(q)
				   
					if err.number = 0 then
					   connStr.CommitTrans
					end if
					
					cmdView	= "True"
					
					End if
					rs.Close
					set rs = Nothing
					
					

	elseif txtJenis = "Delete" then
		connstr.BeginTrans
		q = " delete from DOCSOUNDINGDETAIL" & _
		    " where DOCNUMBER = '" & DOCNUMBER & "'" & _
			"       and DOCDATE= to_date('" & DOCDATE & "','MM/DD/YYYY')" & _
			"       and SOUNDINGSTEP='" & SOUNDINGSTEP & "'" & _
			"       and fcba = '" & Session("fcBA") & "'"
   	   'response.Write(q)
	   connStr.execute(q)
	   
		if err.number = 0 then
		   connStr.CommitTrans
		end if
		
		cmdView	= "True"
		
	
	   elseiF txtJenis = "New" and len(trim(TEMPERATURE)) > 0 and len(trim(HIGHLEVELTANK)) > 0 and Add_OPEN ="" then
		connstr.BeginTrans
				if DOCNUMBER = "(NEW)" then
				
					'auto number (NEW)

					 Set rsAUTONUM = Server.CreateObject("ADODB.Recordset")
					 numq = "select nvl(controlnumber_api.get_nextnumber(upper('DOCSOUNDING'), " & year(date) & ", " & month(date) & ", '" & session("fcBA") & "'), 'FALSE') prcode from dual "
					 'response.write(numq)
					 'response.write (numq)
					 rsAUTONUM.Open  numq, connstr	   
					if Not rsAUTONUM.Eof then
						AutoNum = rsAUTONUM("prcode")
					end if
								
					DOCNUMBER = AutoNum		
					rsAUTONUM.Close
					set rsAUTONUM = Nothing
				else
					DOCNUMBER = Request("DOCNUMBER")	
				end if
				'response.write(DOCNUMBER)
				if SOUNDINGTYPE = "NORMAL" then
					ROWSTATEID = "CLOSED"
				end if
				
				q = " insert into DOCSOUNDING (" & _
					" DOCNUMBER, DOCDATE,PSTORAGE_STORAGECODE,SOUNDINGTYPE, ROWSTATEID, fcentry, fcedit" & _
					" , fcip, fcba, lastupdate, lasttime " & _
					" ) values " & _
					" ('" & DOCNUMBER & "'" & _
					" , to_date('" & DOCDATE & "','mm/dd/yyyy')" & _
					" , '" & PSTORAGE_STORAGECODE & "'" & _
					" , '" & SOUNDINGTYPE & "'" & _
					" , '" & ROWSTATEID & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcBA") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate,'hh24:mi')" & _
					" )"
				'response.Write(q)
				'response.Flush()
				connStr.execute(q)	
				
				q2 = " insert into DOCSOUNDINGDETAIL (" & _
					" DOCNUMBER, DOCDATE,SOUNDINGSTEP,HIGHLEVELTANK, TEMPERATURE,QUANTITY, fcentry, fcedit" & _
					" , fcip, fcba, lastupdate, lasttime " & _
					" ) values " & _
					" ('" & DOCNUMBER & "'" & _
					" , to_date('" & DOCDATE & "','MM/DD/YYYY')" & _
					" , '" & SOUNDINGSTEP & "'" & _
					" , '" & HIGHLEVELTANK & "'" & _
					" , '" & TEMPERATURE & "'" & _
					" , '" & QUANTITY & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcBA") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate,'hh24:mi')" & _
					" )"
					
				'response.Write(q2)
				connStr.execute(q2)
		
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
		
		cmdView	= "True"
		txtJenis = "select_head"
		
	if txtModule = "Mill_Production" and txtTable = "Sounding_Normal" then
		Response.Redirect "sounding_normal.asp?txtModule=Mill_Production&txttable=Sounding_Normal&NORMALSOUNDINGNUMBER=" & NORMALSOUNDINGNUMBER & "&HIGHLEVELTANK=" & HIGHLEVELTANK & "&TEMPERATURE=" & TEMPERATURE & "&QUANTITY=" & QUANTITY & "&DOCNUMBER=" & DOCNUMBER & "&PSTORAGE_STORAGECODE=" & PSTORAGE_STORAGECODE & "&formwindow=open&ViewSounding=True"
		end if
	elseif txtJenis = "New" and Add_OPEN ="Add_OPEN"  then
		connstr.BeginTrans
		q2 = " insert into DOCSOUNDINGDETAIL (" & _
					" DOCNUMBER, DOCDATE,SOUNDINGSTEP,HIGHLEVELTANK, TEMPERATURE,QUANTITY, fcentry, fcedit" & _
					" , fcip, fcba, lastupdate, lasttime " & _
					" ) values " & _
					" ('" & DOCNUMBER & "'" & _
					" , to_date('" & DOCDATE & "','MM/DD/YYYY')" & _
					" , '" & SOUNDINGSTEP & "'" & _
					" , '" & HIGHLEVELTANK & "'" & _
					" , '" & TEMPERATURE & "'" & _
					" , '" & QUANTITY & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcBA") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate,'hh24:mi')" & _
					" )"
					
				'response.Write(q2)
				connStr.execute(q2)
				
		q = " update DOCSOUNDING" & _
					" set ROWSTATEID = 'CLOSED'" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _	
					"     , fcba = '" & Session("fcBA") & "'" & _	
					"     , lastupdate = sysdate" & _	
					"     , lasttime = to_char(sysdate,'hh24:mi')" & _						
					" where DOCNUMBER = '" & DOCNUMBER & "'" & _ 
					"       and DOCDATE = TO_DATE('" & DOCDATE & "','MM/DD/YYYY')" & _
					"       and fcba = '" & Session("fcBA") & "'"
					
				'response.Write(q)
				connStr.execute(q)

		if err.number = 0 then
		   connStr.CommitTrans
		end if
		
		cmdView	= "True"
		txtJenis = "select_head"
		Add_OPEN = ""
		if txtModule = "Mill_Production" and txtTable = "Sounding_Normal" then
		Response.Redirect "sounding_normal.asp?txtModule=Mill_Production&txttable=Sounding_Normal&NORMALSOUNDINGNUMBER=" & NORMALSOUNDINGNUMBER & "&HIGHLEVELTANK=" & HIGHLEVELTANK & "&TEMPERATURE=" & TEMPERATURE & "&QUANTITY=" & QUANTITY & "&DOCNUMBER=" & DOCNUMBER & "&PSTORAGE_STORAGECODE=" & PSTORAGE_STORAGECODE & "&formwindow=open&ViewSounding=True"
		END IF
	elseif txtJenis ="Edit" and txtAdd ="Update" then
	
		connstr.BeginTrans
		
				q = " update DOCSOUNDINGDETAIL" & _
					" set SOUNDINGSTEP = '" & SOUNDINGSTEP & "'" & _
					"     , HIGHLEVELTANK = '" & HIGHLEVELTANK & "'" & _
					"     , TEMPERATURE = '" & TEMPERATURE & "'" & _
					"     , QUANTITY = '" & QUANTITY & "'" & _					
					" where DOCNUMBER = '" & DOCNUMBER & "'" & _ 
					"       and DOCDATE = TO_DATE('" & DOCDATE & "','MM/DD/YYYY')" & _
					"       and fcba = '" & Session("fcBA") & "'"
					
				'response.Write(q)
				connStr.execute(q)

				'response.Flush()
				connStr.execute("commit")
				'RESPONSE.Write "SQLnsddsdya : " & Q
		
		if err.number = 0 then
		   connStr.CommitTrans
		   if txtModule = "Mill_Production" and txtTable = "Sounding_Normal" then
				Response.Redirect "sounding_normal.asp?txtModule=Mill_Production&txttable=Sounding_Normal&NORMALSOUNDINGNUMBER=" & NORMALSOUNDINGNUMBER & "&HIGHLEVELTANK=" & HIGHLEVELTANK & "&TEMPERATURE=" & TEMPERATURE & "&QUANTITY=" & QUANTITY & "&DOCNUMBER=" & DOCNUMBER & "&DOCDATE="&DOCDATE&"&PSTORAGE_STORAGECODE="&PSTORAGE_STORAGECODE&"&FFA_QUALITY="&FFA_QUALITY&"&MOISTURE_QUALITY="&MOISTURE_QUALITY&"&DIRT_QUALITY="&DIRT_QUALITY&"&formwindow=open&ViewSounding=True&txtJenis=Edit"
			END IF
		end if
		
		cmdView	= "True"
		txtJenis = "select_head"
	
    else
        cmdView	= "True"

	end if								 
	'Response.Write "sql " & q
	
else
    cmdView	= request("cmdView")
	
End if

if DOCNUMBER <> "(NEW)" and DOCNUMBER <> "" then

		dim where_val,where_strg,rsQ
		if DOCDATE = "" then
		txtJenis = "select_head"
		where_val = ""
		else
		where_val = " AND DOCDATE = " & DOCDATE & ""
		end if
		if PSTORAGE_STORAGECODE = "" then
		txtJenis = "select_head"
		where_strg = ""
		else
		where_strg = " AND PSTORAGE_STORAGECODE = " & PSTORAGE_STORAGECODE & ""
		end if

		if txtJenis <> "Edit" and PSTORAGE_STORAGECODE = "" and DOCDATE = "" then
		txtJenis = "select_head"
		elseif txtJenis = "Edit" then
		txtJenis = "Edit"
		end if
		
	if txtJenis = "select_head" then	
	Set rsQ = Server.CreateObject("ADODB.Recordset")
	rsQ.CursorType = 3
	rsQ.CursorLocation = 3	
	rsQ.Open "select DOCDATE,PSTORAGE_STORAGECODE,SOUNDINGTYPE,ROWSTATEID,NORMALSOUNDINGNUMBER,TRANSFERSOUNDINGNUMBER,SALESSOUNDINGNUMBER FROM DOCSOUNDING" & _
			" WHERE DOCNUMBER = '" & DOCNUMBER & "' and fcBA='" & Session("fcBA") & "'", connstr

			IF Not rsQ.eof then
			DOCDATE = rsQ("DOCDATE")
			PSTORAGE_STORAGECODE = rsQ("PSTORAGE_STORAGECODE")
			SOUNDINGTYPE = rsQ("SOUNDINGTYPE")
			ROWSTATEID = rsQ("ROWSTATEID")
			NORMALSOUNDINGNUMBER = rsQ("NORMALSOUNDINGNUMBER")
			TRANSFERSOUNDINGNUMBER = rsQ("TRANSFERSOUNDINGNUMBER")
			SALESSOUNDINGNUMBER = rsQ("SALESSOUNDINGNUMBER")
			
			ELSE%>
			<script>
					alert("No Data");
			</script>
		<%	END IF
	rsQ.Close
	set rsQ = Nothing

	end if	
	else
	 cmdView	= "True"
end if	
%>
<!-- END NEW WINDOW SCRIPT -->
<script language="javascript">
			function level_tank(str)
				{
				var tempr = document.getElementById("TEMPERATURE").value;
				if 	(tempr == ""){
					document.getElementById("QUANTITY").value="Temperature ?";
					}else{
						if (str.length==0)
						  {
						  document.getElementById("QUANTITY").value="0";
						  return;
						  }
						  
						if (window.XMLHttpRequest)
						  {// code for IE7+, Firefox, Chrome, Opera, Safari
						  xmlhttp=new XMLHttpRequest();
						  }
						else
						  {// code for IE6, IE5
						  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
						  }
						  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("QUANTITY").value=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/sounding_ajx.asp?initialStorageCode=<%=PSTORAGE_STORAGECODE%>&initialReadingNumber="+str+"&initialFcba=<%=session("fcBA")%>&initialTemperature=" + tempr,true);
						xmlhttp.send();
					}
				}
				function level_tempr(str)
				{
				var hlt = document.getElementById("HIGHLEVELTANK").value;
				if 	(hlt == ""){
					document.getElementById("QUANTITY").value="Reading Lvl Tank ?";
					}else{
						if (str.length==0)
						  {
						  document.getElementById("QUANTITY").value="0";
						  return;
						  }
						  
						if (window.XMLHttpRequest)
						  {// code for IE7+, Firefox, Chrome, Opera, Safari
						  xmlhttp=new XMLHttpRequest();
						  }
						else
						  {// code for IE6, IE5
						  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
						  }
						  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("QUANTITY").value=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/sounding_ajx.asp?initialStorageCode=<%=PSTORAGE_STORAGECODE%>&initialReadingNumber="+hlt+"&initialFcba=<%=session("fcBA")%>&initialTemperature=" + str,true);
						xmlhttp.send();
					}
				}
</script>
<script language="javascript">


	function validasi2(form){
		if (form.DOCNUMBER.value == 0){

		alert("Please enter Document Number");

		form.DOCNUMBER.focus();

		return (false);

		}

		
		if (form.SOUNDINGSTEP.value == 0){

		alert("Please enter SOUNDING STEP");

		form.SOUNDINGSTEP.focus();

		return (false);

		}

		if (form.HIGHLEVELTANK.value == 0){

		alert("Please enter HIGH LEVEL TANK");

		form.HIGHLEVELTANK.focus();

		return (false);

		}
		
		
		if (form.TEMPERATURE.value == 0){

		alert("Please enter TEMPERATURE");

		form.TEMPERATURE.focus();

		return (false);

		}
		
		if (form.QUANTITY.value == 0){

		alert("Please enter QUANTITY");

		form.QUANTITY.focus();

		return (false);

		}
		
	 
	  return (true);

	}
	
	
</script> 

<% 
if cmdView = "True" and Len(Trim(DOCNUMBER)) > 0 and Len(Trim(DOCDATE)) > 0 and Len(Trim(PSTORAGE_STORAGECODE)) > 0 Then 
  
	Dim rsGAD
	set rsGAD = server.CreateObject("Adodb.recordset")

	q = " select DOCNUMBER" & _
	    "        , DOCDATE" & _
	    "        , SOUNDINGSTEP" & _
		"        , HIGHLEVELTANK" & _
		"        , TEMPERATURE" & _
		"        , QUANTITY" & _
	    " from DOCSOUNDINGDETAIL" & _
	    " where DOCNUMBER = '" & DOCNUMBER & "'" & _
		"		and DOCDATE = to_date('" & DOCDATE & "','mm/dd/yyyy')" & _
	    "       and fcBA = '" & session("fcBA") & "'" 
		
	'response.Write(q)
	'response.Flush()
	rsGAD.CursorType = 3
	rsGAD.CursorLocation = 3
	rsGAD.Open q, connstr
	if rsGAD.EOF then
	txtJenis = "New_Edit"
	else
	txtJenis = "Edit"
	SOUNDINGSTEP = rsGAD("SOUNDINGSTEP")
	HIGHLEVELTANK = rsGAD("HIGHLEVELTANK")
	TEMPERATURE = rsGAD("TEMPERATURE")
	QUANTITY = rsGAD("QUANTITY")
	Add_OPEN = "Add_OPEN"
	end if	

End if	
			dim read_only,read_only_edit,disabled_edit,Input_PSTORAGE_STORAGECODE
			dim submit_disabled,style_btn
			
			if ROWSTATEID = "OPEN" then 
				read_only = "readonly='True' style='background:#FFFFCC'"
				else
				read_only = "readonly='True' style='background:#FFFFCC'"
			end if
			if txtJenis = "Edit" then
				read_only_edit = "readonly='True' style='background:#FFFFCC'"
				disabled_edit = "disabled"
				Input_PSTORAGE_STORAGECODE = "<input Type='Hidden' Value='" & PSTORAGE_STORAGECODE & "' Name='PSTORAGE_STORAGECODE'>"
				else
				read_only_edit = ""
				disabled_edit = ""
				Input_PSTORAGE_STORAGECODE = ""
			end if
			
			if txtJenis = "select_head" and Add_OPEN <> "Add_OPEN"then 
				style_btn = "Submit"
			elseif txtJenis = "New" or txtJenis = "Edit" then 
				submit_disabled = "disabled"
				style_btn = "button"
			elseif Add_OPEN = "Add_OPEN" then
				submit_disabled = "disabled"
				style_btn = "button"
			else
				style_btn = "Submit"
			end if
			'response.write(cmdView)
			'response.write(txtJenis)
			'esponse.write("add =" & txtAdd)
			'response.write(Add_OPEN)
			%>
			
<form action="<%=form_direct%>" method="post" name="eForm" onSubmit="javascript:return validasi2(this)" style="background:#ffffff;border: 0px;padding: 0px;">
<div style="border: 1px solid #CCCCCC;padding:20px;">
<!-- <form name="frmSend" method="post" action="sounding.asp" onSubmit="javascript:return validasi(this)">-->
	<table class="data" width="100%">
		<tr>
		<td align="left">Document Number</td>
		<td align="left"><input id="DOCNUMBER" type="text" maxlength="50" name="DOCNUMBER" size="30" value="<%=DOCNUMBER%>" <%=read_only_edit%>><font color="#FF0000">&nbsp;&nbsp;*</font>
		</td>
		</tr>
		
		<tr>
		<td align="left">Doc Date</td>
		<td align="left"><input id="DOCDATE" class="datepicker" type="text" maxlength="10" name="DOCDATE" size="30" value="<%=DOCDATE%>" <%=read_only_edit%>><font color="#FF0000">&nbsp;&nbsp;*</font>
		</td>
		</tr>
		<tr>
			<td align="left">Storage Code</td>
			<td align="left">
			<select name="PSTORAGE_STORAGECODE" style="width:208px" <%=disabled_edit%>>
				<% 
				DIM rsCallCombo
				q = "select STORAGECODE from PRODUCTSTORAGE where PRODUCT_PRODUCTCODE = 'CPO' and fcba = '" & session("fcba") & "'"
				Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
				rsCallCombo.CursorType = 3
				rsCallCombo.CursorLocation = 3
				rsCallCombo.Open q, connstr
				if not rsCallCombo.eof then
				 Do While Not rsCallCombo.eof

				 if rsCallCombo("STORAGECODE") = PSTORAGE_STORAGECODE then%>
					<option value="<%=rsCallCombo("STORAGECODE")%>" Selected><%=rsCallCombo("STORAGECODE")%></option> 
				 <%Else%>
					<option value="<%=rsCallCombo("STORAGECODE")%>"><%=rsCallCombo("STORAGECODE")%></option>  
				<%End if

					rsCallCombo.movenext
				 loop
				End if
				rsCallCombo.Close
				set rsCallCombo = Nothing 
				rsCallCombo = ""
				%>
			</select><font color="#FF0000">&nbsp;&nbsp;*</font>
				<%=Input_PSTORAGE_STORAGECODE%>
			</td>
		</tr>
		<tr>
		<td align="left">Sounding Type</td>
		<td align="left">
			<select name="SOUNDINGTYPE" style="width:208px" <%=disabled_edit%>>
				<% 
				q = "select * from SOUNDINGTYPE"
				Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
				rsCallCombo.CursorType = 3
				rsCallCombo.CursorLocation = 3
				rsCallCombo.Open q, connstr
				if not rsCallCombo.eof then
				 Do While Not rsCallCombo.eof

				 if rsCallCombo("FCCODE") = SOUNDINGTYPE then%>
					<option value="<%=rsCallCombo("FCCODE")%>" Selected><%=rsCallCombo("FCNAME")%></option> 
				 <%Else%>
					<option value="<%=rsCallCombo("FCCODE")%>"><%=rsCallCombo("FCNAME")%></option>  
				<%End if

					rsCallCombo.movenext
				 loop
				End if
				rsCallCombo.Close
				set rsCallCombo = Nothing %>
			</select>
			<font color="#FF0000">&nbsp;&nbsp;*</font>
		</td>
		</tr>
		<tr>
		<td align="left">Status</td>
		<td align="left">
			<% if ROWSTATEID = "OPEN" and Add_OPEN <> "Add_OPEN" then%>
				<input id="Add_OPEN_btn" type="hidden" name="Add_OPEN" value="Add_OPEN"  style="display:block;">
				<input id="Add_btn" class="jqplot-image-button" Type="Submit" name="Add" value="<%=ROWSTATEID%>" style="display:block;">
			<% elseif ROWSTATEID = "CLOSED" then%>
				<input class="jqplot-image-button" Type="button" name="Add" value="<%=ROWSTATEID%>" disabled>
			<% else %>	
				<input class="jqplot-image-button" Type="button" name="Add" value="No Status" disabled>
			<% end if%>  
			
		</td>
		<tr><td colspan="2"  align="left">
					<input Type="Hidden" Value="True" Name="cmdView">
					<input Type="Hidden" Value="570px" Name="Height_form">
					<% if txtJenis <> "Edit" and txtAdd <> "Update" then 
						if Add_OPEN = "Add_OPEN" then
						SOUNDINGSTEP = 2
						else
						SOUNDINGSTEP = 1
						end if
					end if%>	
					
					<% if Add_OPEN <> "Add_OPEN" then%>
						<input Type="Hidden" Value="<%=txtJenis%>" Name="txtJenis">
					<% end if %>
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">		
					<input Type="<%=style_btn%>"  name="Add" value="Select" <%=submit_disabled%>>
					<input Type="button" name="new_release" value="New" onClick="clickNew('Button');">&nbsp;&nbsp;
		</td>
		</tr>
    </table>
</div>	
 <!--</form>-->

 <%
 
 if cmdView = "True" and Len(Trim(DOCNUMBER)) > 0 and Len(Trim(DOCDATE)) > 0 and Len(Trim(PSTORAGE_STORAGECODE)) > 0 Then 
  

 %>
<script language="javascript">

	function validasi2(form){
		
		if (form.SOUNDINGSTEP.value == 0){

		alert("Please enter SOUNDING STEP");

		form.SOUNDINGSTEP.focus();

		return (false);

		}

		if (form.HIGHLEVELTANK.value == 0){

		alert("Please enter HIGH LEVEL TANK");

		form.HIGHLEVELTANK.focus();

		return (false);

		}
		
		
		if (form.TEMPERATURE.value == 0){

		alert("Please enter TEMPERATURE");

		form.TEMPERATURE.focus();

		return (false);

		}
		
		if (form.QUANTITY.value == 0){

		alert("Please enter QUANTITY");

		form.QUANTITY.focus();

		return (false);

		}
	 
	  return (true);

	}

</script>
  <!--<form action="sounding.asp" method="post" name="eForm" onSubmit="javascript:return validasi2(this)">-->
  
  <% 
  
  if txtJenis <> "select_head" or Add_OPEN = "Add_OPEN" then 
  
		if Add_OPEN = "Add_OPEN" then%>
		<input type="hidden" name="Add_OPEN" value="Add_OPEN">
		<% end if %>
  <div style="border: 1px solid #CCCCCC;padding:20;margin-top:20px;">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">

		<tr>
			<td align="left">SOUNDING STEP</td>
			<td align="left">
			
				<input type="text" size="15" name="SOUNDINGSTEP" value="<%=SOUNDINGSTEP%>" <%=read_only%>><font color="#FF0000">&nbsp;&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">READING LEVEL TANK (cm)</td>
			<td align="left">
				<input id="HIGHLEVELTANK" type="text" size="15" name="HIGHLEVELTANK" value="<%=HIGHLEVELTANK%>" onkeyup="level_tank(this.value)"> <font color="#FF0000">&nbsp;&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">TEMPERATURE (&#176;C)</td>
			<td align="left">
				<input id="TEMPERATURE" type="text" size="15" name="TEMPERATURE" value="<%=TEMPERATURE%>" onkeyup="level_tempr(this.value)"><font color="#FF0000">&nbsp;&nbsp;*</font>
			</td>
		</tr>
		<tr>
		
			<td align="left">QUANTITY</td>
			<td align="left">
			<%
			DIM TinggiMaxTerukur,ProductCode,VolumeMaxTerukur,IsReadytoDeliver,Round,Specific,Konversi,Bottom
			
			'TinggiMaxTerukur = Get_TinggiMaxTerukur(PSTORAGE_STORAGECODE, session("fcBA"))
			'ProductCode = Get_ProductCode(PSTORAGE_STORAGECODE, session("fcBA"))
			'VolumeMaxTerukur = Get_VolumeMaxTerukur(PSTORAGE_STORAGECODE, session("fcBA"))
			
			'IsReadytoDeliver = Get_IsReadytoDeliver(PSTORAGE_STORAGECODE,session("fcBA"))
			'Round = Get_Round(initialNumber, initialLenght)
			'Specific = Get_SpecificGravity(TEMPERATURE)
			'Konversi =  Get_Konversi(PSTORAGE_STORAGECODE, session("fcBA"))
			'Bottom = Get_Bottom(PSTORAGE_STORAGECODE, session("fcBA"))
			

			
			%>
				<input id="QUANTITY" type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>" readonly="True" style="background:#FFFFCC">
			</td>
		</tr>
	    <TR>
			<td colspan="2" align="left">
			<!--<input type="hidden" name="DOCNUMBER" value="<%=DOCNUMBER%>">
				<input type="hidden" name="DOCDATE" value="<%=DOCDATE%>">
				<input type="hidden" name="PSTORAGE_STORAGECODE" value="<%=PSTORAGE_STORAGECODE%>">
				<input type="hidden" name="SOUNDINGTYPE" value="<%=SOUNDINGTYPE%>">-->
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%IF Request("txtJenis")="Edit" or txtJenis = "Edit" then%>
					<input type="hidden" name="txtJenis" value="Edit">
					<input Type="Submit"  name="Add" value="Update">

				<%Else%>
					<input type="hidden" name="txtJenis" value="New">
					<input Type="Submit"  name="Add" value="Add">
				<%End if%>
			<!--<input Type="Reset" name="Reset" value="Reset">-->
				<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
			</td>
	  </tr>
   </table>
   </div>
  
	<%end if%>
	

<%end if%>
  </form>
 </div></div>
