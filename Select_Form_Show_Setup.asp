<%if Request("txttable") ="BusinessUnit" then%>
    <%if Request("txtJenis") <> "Edit" then%>

          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                BUSINESS UNIT</strong></font></div></td>
          </TR>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Company Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" name="FCCOMPANYCODE" title="Company Code Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Company Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="50"  name="FCCOMPANYNAME"  title="Company Name Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Business Unit Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Business Unit Name</font></td>
            <td align="left"><input type="text" size="70" maxlength="70"  name="FCNAME" title="NAME Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr> 
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Area</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCAREA" title="Area Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Ownership</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCOWNERSHIP"  title="Ownership Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Type</font></td>
            <td align="left"><input type="text" size="1" maxlength="50"  name="fcType"  title="[E] Estate [M] Mill [W] Workshop [C] Composting [O] Others">
			<font color="#FF0000">* [E] Estate [M] Mill [W] Workshop [C] Composting [O] Others</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Crop</font></td>
            <td align="left">
			<% 	Dim rsCropB
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsCropB = Server.CreateObject("ADODB.Recordset")
			rsCropB.CursorType = 3
			rsCropB.CursorLocation = 3
			rsCropB.Open q, connstr
			if not rsCropB.eof then%>
     	    <td align="left"> <select name="fcCrop" id="fcCrop" title="CROPTYPE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsCropB.eof%>
				<option value="<%=rsCropB("FCCODE")%>"><%=rsCropB("FCNAME")%></option> 
			<%rsCropB.movenext
			loop
			End if
			rsCropB.Close
			set rsCropB = Nothing
			q=""
			%>
			</td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Manager</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCMANAGER"  title="Manager Character a s/d z or 1 s/d 9"></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Telephone</font></td>
            <td align="left"><input type="text" size="10" maxlength="10"  name="FCTELEPHONE" title="Telephone Character a s/d z or 1 s/d 9"></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Fax</font></td>
            <td align="left"><input type="text" size="10" maxlength="10"  name="FCFAX" title="Fax Character a s/d z or 1 s/d 9"></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">E-Mail</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCEMAIL" title="E-Mail Character a s/d z or 1 s/d 9"></td>
          </Tr>
       <Tr> 
            <td align="left"><font face="Verdana" size="1">Address</font></td>
            <td align="left"><textarea cols="50" name="ADDRESS" title="Address Character a s/d z or 1 s/d 9"></textarea></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">KTU</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCKTU"  title="KTU Character a s/d z or 1 s/d 9"></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">ISACTIVE</font></td>
            <td align="left">
			<input type="radio" size="30" maxlength="50" value="TRUE" name="ISACTIVE">
			&nbsp;&nbsp;<span>TRUE</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="radio" value="FALSE" name="ISACTIVE" checked>
			&nbsp;&nbsp;<span>FALSE</span>
			</td>
          </Tr>
		   <Tr> 
            <td align="left"></td>
			<td align="left"></td>
          </Tr>
   <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                BUSINESS UNIT</strong></font></div></td>
          </TR>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Company Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" name="FCCOMPANYCODE" value="<%=rsSetup("FCCOMPANYCODE")%>" title="Company Code Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Company Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="50"  name="FCCOMPANYNAME"  value="<%=rsSetup("FCCOMPANYNAME")%>" title="Company Name Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Business Unit Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value="<%=rsSetup("FCCODE")%>" readonly="True" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Business Unit Name</font></td>
            <td align="left"><input type="text" size="70" maxlength="70"  name="FCNAME" value="<%=rsSetup("FCNAME")%>" title="NAME Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
<!--		  
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value="<%=rsSetup("FCCODE")%>" readonly="True" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font> </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="70" maxlength="70"  name="FCNAME" value="<%=rsSetup("FCNAME")%>"  title="NAME Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Company Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" name="FCCOMPANYCODE" value="<%=rsSetup("FCCOMPANYCODE")%>"  title="Company Code Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Company Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="50"  name="FCCOMPANYNAME" value="<%=rsSetup("FCCOMPANYNAME")%>"   title="Company Name Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
-->
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Area</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCAREA" value="<%=rsSetup("FCAREA")%>"  title="Area Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Ownership</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCOWNERSHIP" value="<%=rsSetup("FCOWNERSHIP")%>"   title="Ownership Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Type</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="fcType"  value="<%=rsSetup("fcType")%>" title="[E] Estate [M] Mill [W] Workshop [C] Composting [O] Others">
			<font color="#FF0000">* [E] Estate [M] Mill [W] Workshop [C] Composting [O] Others</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Crop</font></td>
            <td align="left">
			<input type="text" size="30" maxlength="50"  name="fcCrop"  value="<%=rsSetup("fcCrop")%>" title="[OP] Oil Palm [RB] Rubber [S] Sungkai [O] Others">
			<font color="#FF0000">* [OP] Oil Palm [RB] Rubber [S] Sungkai [O] Others</font></td>
          </Tr>
		  
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Manager</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCMANAGER" value="<%=rsSetup("FCMANAGER")%>"   title="Manager Character a s/d z or 1 s/d 9"></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Telephone</font></td>
            <td align="left"><input type="text" size="20" maxlength="20"  name="FCTELEPHONE" value="<%=rsSetup("FCTELEPHONE")%>"   title="Telephone Character a s/d z or 1 s/d 9"></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Fax</font></td>
            <td align="left"><input type="text" size="20" maxlength="20"  name="FCFAX" value="<%=rsSetup("FCFAX")%>"   title="Fax Character a s/d z or 1 s/d 9"></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">E-Mail</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCEMAIL" value="<%=rsSetup("FCEMAIL")%>"   title="E-Mail Character a s/d z or 1 s/d 9"></td>
          </Tr>
       <Tr> 
            <td align="left"><font face="Verdana" size="1">Address</font></td>
            <td align="left"><textarea cols = "50" name="ADDRESS" title="Address Character a s/d z or 1 s/d 9"><%=rsSetup("address")%></textarea></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">KTU</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCKTU" value="<%=rsSetup("FCKTU")%>"  title="KTU Character a s/d z or 1 s/d 9"></td>
          </Tr>
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">ISACTIVE</font></td>
            <td align="left">
			<% 
			DIM checked, unchecked
			if rsSetup("ISACTIVE") = "TRUE" THEN 
			checked = "checked"
			unchecked = ""
			ELSE
			unchecked = "checked"
			checked = ""
			END IF
			
			%>
			<input type="radio" size="30" maxlength="50" value="TRUE" name="ISACTIVE" <%=checked%>>
			&nbsp;&nbsp;<span>TRUE</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="radio" value="FALSE" name="ISACTIVE" <%=unchecked%>>
			&nbsp;&nbsp;<span>FALSE</span>
			</td>
          </Tr>
		   <Tr> 
            <td align="left"></td>
			<td align="left"></td>
          </Tr>
   
   <%End if%>

<%Elseif Request("txttable") ="UserProfile" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                USER AND PROFILE</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Password</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCPASSWORD" title="PASSWORD Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Full Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25" name="FULLNAME" title="FULLNAME Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Department</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="DEPARTMENT"  title="DEPARTMENT Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Access Rights</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ACCESSRIGHTS" title="ACCESSRIGHTS Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
   <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                USER AND PROFILE</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value="<%=rsSetup("FCCODE")%>" readonly="True" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Password</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCPASSWORD" title="PASSWORD Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Full Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25" name="FULLNAME" value="<%=rsSetup("FULLNAME")%>" title="FULLNAME Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Department</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="DEPARTMENT" value="<%=rsSetup("DEPARTMENT")%>" title="DEPARTMENT Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Access Rights</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ACCESSRIGHTS" value="<%=rsSetup("ACCESSRIGHTS")%>" title="ACCESSRIGHTS Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
   
   <%End IF%>
   
<%Elseif Request("txttable") ="Users" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                USER AND PROFILE</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Password</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCPASSWORD" title="PASSWORD Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Full Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25" name="FULLNAME" title="FULLNAME Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Title</font></td>
            <td align="left">
			<select name="TITLE">
			
			<%
			dim rstitle
			q = "select fcCode,description from usertitle Order by FCCODE"
			Set rstitle = Server.CreateObject("ADODB.Recordset")
			rstitle.CursorType = 3
			rstitle.CursorLocation = 3
			rstitle.Open q, connstr
			if not rstitle.eof then
			%>
			<OPTION value="">-- Title --</option>
			<%
			 Do While Not rstitle.eof
			%>
			<option value="<%=rstitle("fcCode")%>"><%=rstitle("fcCode")%> : <%=rstitle("description")%></option>
			<%rstitle.movenext
			 loop
			End if
			rstitle.Close
			set rstitle = Nothing
			q=""
			%>
			</SELECT><font color="#FF0000">*</font>
			
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Department</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="DEPARTMENT"  title="DEPARTMENT Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Access Rights</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ACCESSRIGHTS" title="ACCESSRIGHTS Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Active</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ISACTIVE" title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">First Address</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="FIRSTADDRESS" title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid From</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_FROM" title="VALID FROM Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid Until</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_UNTIL" title="VALID UNTIL Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>		  
		  
   <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                USER AND PROFILE</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value="<%=rsSetup("FCCODE")%>" readonly="True" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Password</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCPASSWORD" title="PASSWORD Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Full Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25" name="FULLNAME" value="<%=rsSetup("FULLNAME")%>" title="FULLNAME Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Title</font></td>
            <td align="left">
			<select name="TITLE">
			
			<%
			
			q = "select fcCode,description from usertitle Order by FCCODE"
			Set rstitle = Server.CreateObject("ADODB.Recordset")
			rstitle.CursorType = 3
			rstitle.CursorLocation = 3
			rstitle.Open q, connstr
			if not rstitle.eof then
			%>
			<OPTION value="">-- Title --</option>
			<%
			 Do While Not rstitle.eof
			%>
			<option value="<%=rstitle("fcCode")%>" <%if rsSetup("TITLE") = rstitle("fcCode") then%> Selected <%End if%>><%=rstitle("fcCode")%> : <%=rstitle("description")%></option>
			<%rstitle.movenext
			 loop
			End if
			rstitle.Close
			set rstitle = Nothing
			q=""
			%>
			</SELECT><font color="#FF0000">*</font>
			
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Department</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="DEPARTMENT" value="<%=rsSetup("DEPARTMENT")%>" title="DEPARTMENT Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Access Rights</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ACCESSRIGHTS" value="<%=rsSetup("ACCESSRIGHTS")%>" title="ACCESSRIGHTS Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Active</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ISACTIVE" value="<%=rsSetup("ISACTIVE")%>"  title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">First Address</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="FIRSTADDRESS"  value="<%=rsSetup("FIRSTADDRESS")%>" title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid From</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_FROM" value="<%=rsSetup("VALID_FROM")%>"  title="VALID FROM Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid Until</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_UNTIL" value="<%=rsSetup("VALID_UNTIL")%>"  title="VALID UNTIL Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>			  
		  
   
   <%End IF%>
   
 
<%Elseif Request("txttable") ="Pattern" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                PATTERN</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Table Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25"  name="TABLENAME" title="TABLENAME Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Display Pattern</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="DISPLAYPATTERN" title="DISPLAYPATTERN Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
         <td align="left"><font face="Verdana" size="1">Edit Pattern</font></td>
            <td align="left"><input type="text" size="30" maxlength="50" name="EDITPATTERN" title="EDITPATTERN Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		 <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                PATTERN</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Table Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25"  name="TABLENAME" value="<%=request("fccode")%>" title="TABLENAME Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Display Pattern</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="DISPLAYPATTERN" value="<%=rsSetup("DISPLAYPATTERN")%>" title="DISPLAYPATTERN Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
         <td align="left"><font face="Verdana" size="1">Edit Pattern</font></td>
            <td align="left"><input type="text" size="30" maxlength="50" name="EDITPATTERN" value="<%=rsSetup("EDITPATTERN")%>" title="EDITPATTERN Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
	<%End if%>
<%Elseif Request("txttable") ="Parameter" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                PARAMETER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Parameter Code</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parameter Name</font></td>
            <td align="left"><input type="text" size="20" maxlength="20"  name="PARAMETERNAME" title="PARAMETERNAME Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
		 <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                PARAMETER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Parameter Code</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCCODE" value="<%=rsSetup("FCCODE")%>" readonly="True" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parameter Name</font></td>
            <td align="left"><input type="text" size="20" maxlength="20"  name="PARAMETERNAME" value="<%=rsSetup("PARAMETERNAME")%>" title="PARAMETERNAME Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
		 
	<%End if%>


<%Elseif Request("txttable") ="Trigger_Control" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                TRIGGER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Activity</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="ACTIVITY" title="ACTIVITY Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Source</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="SOURCE" title="SOURCE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Target</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="TARGET" title="TARGET Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
		 <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                TRIGGER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Activity</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="ACTIVITY" value="<%=rsSetup("ACTIVITY")%>" title="ACTIVITY Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Source</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="SOURCE" value="<%=rsSetup("SOURCE")%>" title="SOURCE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Target</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="TARGET"  value="<%=rsSetup("TARGET")%>" title="TARGET Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
		 
	<%End if%>
<%Elseif Request("txttable") ="Job" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
             </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Group Code</font></td>
            <td align="left"><select name="JOBGROUPCODE" title="JOBGROUPCODE Character a s/d z or 1 s/d 9"  >
			<option value=""></option> 
			<% 	Dim rsComboV
			q = "select * from  jobgroupmaster " & " Order by FCCODE"
			Set rsComboV = Server.CreateObject("ADODB.Recordset")
			rsComboV.CursorType = 3
			rsComboV.CursorLocation = 3
			rsComboV.Open q, connstr
			if not rsComboV.eof then%>
           
			<%
			Do While Not rsCombov.eof
			%>
			<option value="<%=rsComboV("fccode")%>"><%=rsComboV("fccode") & "-" & rsComboV("fcname")%></option> 
			<%rsComboV.movenext
			loop
			End if
			rsComboV.Close
			set rsComboV = Nothing
			q=""
			%>
			</select> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Crop Type</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORCROPTYPE" title="JOB_FORCROPTYPE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Field Status</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORFIELDSTATUS" title="JOB_FORFIELDSTATUS Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Category</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_CATEGORY" title="JOB_CATEGORY Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Unit Of Measurement</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="UNITOFMEASUREMENT" title="UNITOFMEASUREMENT Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Account Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="ACCOUNTCODE" title="ACCOUNTCODE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Code</font></td>
            <td align="left">
			<select name="JOB_FORBALANCESHEETCODE" id="JOB_FORBALANCESHEETCODE" title="JOB_FORBALANCESHEETCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupBSName();">
			<option value=""></option> 
			<% 	Dim rsComboBS
			q = "select * from  Jobbsmaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
	        </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORBALANCESHEETNAME" id="JOB_FORBALANCESHEETNAME" title="JOB_FORBALANCESHEETNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORPROFITANDLOSSCODE" title="JOB_FORPROFITANDLOSSCODE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORPROFITANDLOSSNAME" title="JOB_FORPROFITANDLOSSNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
     <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value="<%=rsSetup("FCCODE")%>"  readonly="True" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
             </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Group Code</font></td>
         	<% 	Dim rsComboVG
			q = "select * from  jobgroupmaster " & " Order by fccode"
			Set rsComboVG = Server.CreateObject("ADODB.Recordset")
			rsComboVG.CursorType = 3
			rsComboVG.CursorLocation = 3
			rsComboVG.Open q, connstr
			if not rsComboVG.eof then%>
            <td align="left"> <select name="JOBGROUPCODE" title="JOBGROUPCODE Character a s/d z or 1 s/d 9">
			<%
			Do While Not rsComboVG.eof
			%>
						
			<%if rsComboVG("fccode") = rssetup("JOBGROUPCODE") then%>
			<option value="<%=rsComboVG("fccode")%>" Selected><%=rsComboVG("fccode") & "-" & rsComboVG("FCNAME")%></option> 
			<%Else%>
			<option value="<%=rsComboVG("fccode")%>"><%=rsComboVG("fccode") & "-" & rsComboVG("fcname")%></option> 
			<%End if%>
			<%rsComboVG.movenext
			loop
			End if
			rsComboVG.Close
			set rsComboVG = Nothing
			q=""
			%>
			</select> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Crop Type</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORCROPTYPE" value="<%=rsSetup("JOB_FORCROPTYPE")%>"  title="JOB_FORCROPTYPE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Field Status</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORFIELDSTATUS" value="<%=rsSetup("JOB_FORFIELDSTATUS")%>" title="JOB_FORFIELDSTATUS Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Category</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_CATEGORY" value="<%=rsSetup("JOB_CATEGORY")%>" title="JOB_CATEGORY Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Unit Of Measurement</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="UNITOFMEASUREMENT" value="<%=rsSetup("UNITOFMEASUREMENT")%>" title="UNITOFMEASUREMENT Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Account Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="ACCOUNTCODE" value="<%=rsSetup("ACCOUNTCODE")%>" title="ACCOUNTCODE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Code</font></td>
         	<% 	Dim rsComboBSE
			q = "select * from  Jobbsmaster " & " Order by fccode"
			Set rsComboBSE = Server.CreateObject("ADODB.Recordset")
			rsComboBSE.CursorType = 3
			rsComboBSE.CursorLocation = 3
			rsComboBSE.Open q, connstr
			if not rsComboBSE.eof then%>
            <td align="left"> <select name="JOB_FORBALANCESHEETCODE" id="JOB_FORBALANCESHEETCODE" title="JOB_FORBALANCESHEETCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupBSName();">
			<%
			Do While Not rsComboBSE.eof
			%>
						
			<%if rsComboBSE("fccode") = rssetup("JOB_FORBALANCESHEETCODE") then%>
			<option value="<%=rsComboBSE("fccode")%>" Selected><%=rsComboBSE("fccode") & "-" & rsComboBSE("FCNAME")%></option> 
			<%Else%>
			<option value="<%=rsComboBSE("fccode")%>"><%=rsComboBSE("fccode") & "-" & rsComboBSE("fcname")%></option> 
			<%End if%>
			<%rsComboBSE.movenext
			loop
			End if
			rsComboBSE.Close
			set rsComboBSE = Nothing
			q=""
			%>
			</select> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORBALANCESHEETNAME" id="JOB_FORBALANCESHEETNAME" value="<%=rsSetup("JOB_FORBALANCESHEETNAME")%>" title="JOB_FORBALANCESHEETNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORPROFITANDLOSSCODE" value="<%=rsSetup("JOB_FORPROFITANDLOSSCODE")%>" title="JOB_FORPROFITANDLOSSCODE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORPROFITANDLOSSNAME" value="<%=rsSetup("JOB_FORPROFITANDLOSSNAME")%>" title="JOB_FORPROFITANDLOSSNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
  <%end if%>
<%Elseif Request("txttable") ="Jobmaster" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
             </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Group Code</font></td>
            <td align="left"><select name="JOBGROUPCODE" title="JOBGROUPCODE Character a s/d z or 1 s/d 9"  >
			<option value=""></option> 
			<% 	
			q = "select * from  jobgroupmaster " & " Order by FCCODE"
			Set rsComboV = Server.CreateObject("ADODB.Recordset")
			rsComboV.CursorType = 3
			rsComboV.CursorLocation = 3
			rsComboV.Open q, connstr
			if not rsComboV.eof then%>
           
			<%
			Do While Not rsCombov.eof
			%>
			<option value="<%=rsComboV("fccode")%>"><%=rsComboV("fccode") & "-" & rsComboV("fcname")%></option> 
			<%rsComboV.movenext
			loop
			End if
			rsComboV.Close
			set rsComboV = Nothing
			q=""
			%>
			</select><font color="#FF0000">*</font>
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Crop Type</font></td>
            <td align="left">
			<% 	
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsCropB = Server.CreateObject("ADODB.Recordset")
			rsCropB.CursorType = 3
			rsCropB.CursorLocation = 3
			rsCropB.Open q, connstr
			if not rsCropB.eof then%>
     	    <select name="JOB_FORCROPTYPE" id="JOB_FORCROPTYPE" title="CROPTYPE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsCropB.eof%>
				<option value="<%=rsCropB("FCCODE")%>"><%=rsCropB("FCNAME")%></option> 
			<%rsCropB.movenext
			loop
			End if
			rsCropB.Close
			set rsCropB = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font>
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Field Status</font></td>
            <td align="left">
			<% 	Dim rsStatus
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0006' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsStatus = Server.CreateObject("ADODB.Recordset")
			rsStatus.CursorType = 3
			rsStatus.CursorLocation = 3
			rsStatus.Open q, connstr
			if not rsStatus.eof then%>
     	    <select name="JOB_FORFIELDSTATUS" id="JOB_FORFIELDSTATUS" title="STATUS Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsStatus.eof%>
				<option value="<%=rsStatus("FCCODE")%>"><%=rsStatus("FCNAME")%></option> 
			<%rsStatus.movenext
			loop
			End if
			rsStatus.Close
			set rsStatus = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font>
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Category</font></td>
            <td align="left">
			<% 	
			q = "select locationtypecode fccode, locationtypename fcname from locationtype order by locationtypecode"
			Set rsStatus = Server.CreateObject("ADODB.Recordset")
			rsStatus.CursorType = 3
			rsStatus.CursorLocation = 3
			rsStatus.Open q, connstr
			if not rsStatus.eof then%>
     	    <select name="JOB_CATEGORY" id="JOB_CATEGORY" title="STATUS Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsStatus.eof%>
				<option value="<%=rsStatus("FCCODE")%>"><%=rsStatus("fccode") & "-" & rsStatus("fcname")%></option> 
			<%rsStatus.movenext
			loop
			End if
			rsStatus.Close
			set rsStatus = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font>
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Unit Of Measurement</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="UNITOFMEASUREMENT" title="UNITOFMEASUREMENT Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Account Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="ACCOUNTCODE" title="ACCOUNTCODE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Code</font></td>
            <td align="left">
			<select name="JOB_FORBALANCESHEETCODE" id="JOB_FORBALANCESHEETCODE" title="JOB_FORBALANCESHEETCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupBSName();">
			<option value=""></option> 
			<% 	
			q = "select * from  Jobbsmaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
	        </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORBALANCESHEETNAME" id="JOB_FORBALANCESHEETNAME" title="JOB_FORBALANCESHEETNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Code</font></td>
            <td align="left">
			<select name="JOB_FORPROFITANDLOSSCODE" id="JOB_FORPROFITANDLOSSCODE" title="JOB_FORPROFITANDLOSSCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupISName();">
			<option value=""></option> 
			<% 	
			q = "select * from  jobismaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
			  </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" id="JOB_FORPROFITANDLOSSNAME" name="JOB_FORPROFITANDLOSSNAME" title="JOB_FORPROFITANDLOSSNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For CashFlow Code</font></td>
            <td align="left">
			<select name="JOB_FORCASHFLOWCODE" id="JOB_FORCASHFLOWCODE" title="JOB_FORCASHFLOWCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupCFName();">
			<option value=""></option> 
			<% 	
			q = "select * from  jobcfmaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
		    </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For CashFlow Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" id="JOB_FORCASHFLOWNAME" name="JOB_FORCASHFLOWNAME" title="JOB_FORCASHFLOWNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Budget Code</font></td>
            <td align="left">
			<select name="JOB_FORBUDGETCODE" id="JOB_FORBUDGETCODE" title="JOB_FORBUDGETCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupBTName();">
			<option value=""></option> 
			<% 	
			q = "select * from  jobbtmaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
	        </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Budget Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" id="JOB_FORBUDGETNAME"  name="JOB_FORBUDGETNAME" title="JOB_FORBUDGETNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
           <Tr> 
            <td align="left"><font face="Verdana" size="1">Budget Qty</font></td>
            <td align="left">
             <select name="budgetqty" id="budgetqty">
                <option value="Y">YA</option>
                <option value="N">Tidak</option>
              </select>
            </td>
          </Tr> 
     <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value="<%=rsSetup("FCCODE")%>"  readonly="True" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
             </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Group Code</font></td>
         	<% 	
			q = "select * from  jobgroupmaster " & " Order by fccode"
			Set rsComboVG = Server.CreateObject("ADODB.Recordset")
			rsComboVG.CursorType = 3
			rsComboVG.CursorLocation = 3
			rsComboVG.Open q, connstr
			if not rsComboVG.eof then%>
            <td align="left"> <select name="JOBGROUPCODE" title="JOBGROUPCODE Character a s/d z or 1 s/d 9">
			<%
			Do While Not rsComboVG.eof
			%>
						
			<%if rsComboVG("fccode") = rssetup("JOBGROUPCODE") then%>
			<option value="<%=rsComboVG("fccode")%>" Selected><%=rsComboVG("fccode") & "-" & rsComboVG("FCNAME")%></option> 
			<%Else%>
			<option value="<%=rsComboVG("fccode")%>"><%=rsComboVG("fccode") & "-" & rsComboVG("fcname")%></option> 
			<%End if%>
			<%rsComboVG.movenext
			loop
			End if
			rsComboVG.Close
			set rsComboVG = Nothing
			q=""
			%>
			</select> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Crop Type</font></td>
            <td align="left">
			<% 	
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsCropB = Server.CreateObject("ADODB.Recordset")
			rsCropB.CursorType = 3
			rsCropB.CursorLocation = 3
			rsCropB.Open q, connstr
			if not rsCropB.eof then%>
     	    <select name="JOB_FORCROPTYPE" id="JOB_FORCROPTYPE" title="CROPTYPE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsCropB.eof%>
				<%if rsCropB("fccode") = rssetup("JOB_FORCROPTYPE") then%>
				<option value="<%=rsCropB("fccode")%>" Selected><%=rsCropB("fccode") & "-" & rsCropB("FCNAME")%></option> 
				<%Else%>
				<option value="<%=rsCropB("fccode")%>"><%=rsCropB("fccode") & "-" & rsCropB("fcname")%></option> 
				<%End if%>
			<%rsCropB.movenext
			loop
			End if
			rsCropB.Close
			set rsCropB = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font>
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Field Status</font></td>
            <td align="left">
			<% 	
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0006' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsStatus = Server.CreateObject("ADODB.Recordset")
			rsStatus.CursorType = 3
			rsStatus.CursorLocation = 3
			rsStatus.Open q, connstr
			if not rsStatus.eof then%>
     	    <select name="JOB_FORFIELDSTATUS" id="JOB_FORFIELDSTATUS" title="STATUS Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsStatus.eof%>
				<%if rsStatus("fccode") = rssetup("JOB_FORFIELDSTATUS") then%>
				<option value="<%=rsStatus("fccode")%>" Selected><%=rsStatus("fccode") & "-" & rsStatus("FCNAME")%></option> 
				<%Else%>
				<option value="<%=rsStatus("fccode")%>"><%=rsStatus("fccode") & "-" & rsStatus("fcname")%></option> 
				<%End if%>
				
			<%rsStatus.movenext
			loop
			End if
			rsStatus.Close
			set rsStatus = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font>
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job Category</font></td>
            <td align="left">
		    <% 	
			q = "select locationtypecode fccode, locationtypename fcname from locationtype order by locationtypecode"
			Set rsStatus = Server.CreateObject("ADODB.Recordset")
			rsStatus.CursorType = 3
			rsStatus.CursorLocation = 3
			rsStatus.Open q, connstr
			if not rsStatus.eof then%>
     	    <select name="JOB_CATEGORY" id="JOB_CATEGORY" title="STATUS Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsStatus.eof%>
				<%if rsStatus("fccode") = rssetup("JOB_CATEGORY") then%>
				<option value="<%=rsStatus("fccode")%>" Selected><%=rsStatus("fccode") & "-" & rsStatus("FCNAME")%></option> 
				<%Else%>
				<option value="<%=rsStatus("fccode")%>"><%=rsStatus("fccode") & "-" & rsStatus("fcname")%></option> 
				<%End if%>
			<%rsStatus.movenext
			loop
			End if
			rsStatus.Close
			set rsStatus = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font>

			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Unit Of Measurement</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="UNITOFMEASUREMENT" value="<%=rsSetup("UNITOFMEASUREMENT")%>" title="UNITOFMEASUREMENT Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Account Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="ACCOUNTCODE" value="<%=rsSetup("ACCOUNTCODE")%>" title="ACCOUNTCODE Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Code</font></td>
         	<% 	
			q = "select * from  Jobbsmaster " & " Order by fccode"
			Set rsComboBSE = Server.CreateObject("ADODB.Recordset")
			rsComboBSE.CursorType = 3
			rsComboBSE.CursorLocation = 3
			rsComboBSE.Open q, connstr
			if not rsComboBSE.eof then%>
            <td align="left"> <select name="JOB_FORBALANCESHEETCODE" id="JOB_FORBALANCESHEETCODE" title="JOB_FORBALANCESHEETCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupBSName();">
			<%
			Do While Not rsComboBSE.eof
			%>
						
			<%if rsComboBSE("fccode") = rssetup("JOB_FORBALANCESHEETCODE") then%>
			<option value="<%=rsComboBSE("fccode")%>" Selected><%=rsComboBSE("fccode") & "-" & rsComboBSE("FCNAME")%></option> 
			<%Else%>
			<option value="<%=rsComboBSE("fccode")%>"><%=rsComboBSE("fccode") & "-" & rsComboBSE("fcname")%></option> 
			<%End if%>
			<%rsComboBSE.movenext
			loop
			End if
			rsComboBSE.Close
			set rsComboBSE = Nothing
			q=""
			%>
			</select> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Balance Sheet Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="JOB_FORBALANCESHEETNAME" id="JOB_FORBALANCESHEETNAME" value="<%=rsSetup("JOB_FORBALANCESHEETNAME")%>" title="JOB_FORBALANCESHEETNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Code</font></td>
            <td align="left">
			<select name="JOB_FORPROFITANDLOSSCODE" id="JOB_FORPROFITANDLOSSCODE" title="JOB_FORPROFITANDLOSSCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupISName();">
			<option value=""></option> 
			<% 	
			q = "select * from  jobismaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
		    </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Profit And Loss Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  id="JOB_FORPROFITANDLOSSNAME" name="JOB_FORPROFITANDLOSSNAME" value="<%=rsSetup("JOB_FORPROFITANDLOSSNAME")%>" title="JOB_FORPROFITANDLOSSNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For CashFlow Code</font></td>
            <td align="left">
			<select name="JOB_FORCASHFLOWCODE" id="JOB_FORCASHFLOWCODE" title="JOB_FORCASHFLOWCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupCFName();">
			<option value=""></option> 
			<% 	
			q = "select * from  jobcfmaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For CashFlow Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" id="JOB_FORCASHFLOWNAME"  name="JOB_FORCASHFLOWNAME" value="<%=rsSetup("JOB_FORCASHFLOWNAME")%>" title="JOB_FORCASHFLOWNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
		   <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Budget Code</font></td>
            <td align="left">
			<select name="JOB_FORBUDGETCODE" id="JOB_FORBUDGETCODE" title="JOB_FORBUDGETCODE Character a s/d z or 1 s/d 9" onchange="javascript:GetTxtGroupBTName();">
			<option value=""></option> 
			<% 	
			q = "select * from  jobbtmaster " & " Order by FCCODE"
			Set rsComboBS = Server.CreateObject("ADODB.Recordset")
			rsComboBS.CursorType = 3
			rsComboBS.CursorLocation = 3
			rsComboBS.Open q, connstr
			if not rsComboBS.eof then%>
           
			<%
			Do While Not rsComboBS.eof
			%>
			<option value="<%=rsComboBS("fccode")%>"><%=rsComboBS("fccode") & "-" &rsComboBS("fcname")%></option> 
			<%rsComboBS.movenext
			loop
			End if
			rsComboBS.Close
			set rsComboBS = Nothing
			q=""
			%>
			</select> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Job For Budget Name</font></td>
            <td align="left"><input type="text" size="30" maxlength="30" if="JOB_FORBUDGETNAME" name="JOB_FORBUDGETNAME" value="<%=rsSetup("JOB_FORBUDGETNAME")%>" title="JOB_FORBUDGETNAME Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr> 
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Budget Qty</font></td>
            <td align="left">
              <select name="budgetqty" id="budgetqty">
                <%
                    IF rsSetup("BUDGETQTY") = "Y" THEN
                %>
                <option value="Y" selected>YA</option>
                <option value="N">Tidak</option>
                <%else%>
                <option value="N" selected>Tidak</option>
                <option value="Y">YA</option>
                <%END IF%>
              </select>
            </td>
          </Tr> 
  <%end if%>  
<%Elseif Request("txttable") ="JobGroup" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB GROUP</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="PARENTCODE" title="PARENTCODE Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC_ENG" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
		<%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB GROUP</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value ="<%=rsSetup("FCCODE")%>" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value ="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="PARENTCODE" value ="<%=rsSetup("PARENTCODE")%>" title="PARENTCODE Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC" value ="<%=rsSetup("PARENTDESC")%>" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC_ENG" value ="<%=rsSetup("PARENTDESC_END")%>" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
	<%End if%>
<%Elseif Request("txttable") ="JobGroupMaster" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB GROUP MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="PARENTCODE" title="PARENTCODE Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC_ENG" value ="<%=rsSetup("PARENTDESC_ENG")%>" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
		<%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB GROUP MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value ="<%=rsSetup("FCCODE")%>" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value ="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="PARENTCODE" value ="<%=rsSetup("PARENTCODE")%>" title="PARENTCODE Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC" value ="<%=rsSetup("PARENTDESC")%>" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Parent Description(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="PARENTDESC_ENG" value ="<%=rsSetup("PARENTDESC")%>" title="PARENTDESC Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
	<%End if%>

<%Elseif Request("txttable") ="Jobbsmaster" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR BALANCESHEET MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name (English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
		<%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR BALANCESHEET MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value ="<%=rsSetup("FCCODE")%>" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value ="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" value ="<%=rsSetup("FCNAME_ENG")%>" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
	<%End if%>
<%Elseif Request("txttable") ="Jobismaster" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR PROFIT AND LOSS MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name (English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
		<%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR PROFIT AND LOSS MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value ="<%=rsSetup("FCCODE")%>" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value ="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" value ="<%=rsSetup("FCNAME_ENG")%>" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
	<%End if%>
	
<%Elseif Request("txttable") ="Jobcfmaster" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR CASHFLOW MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name (English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
		<%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR CASHFLOW MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value ="<%=rsSetup("FCCODE")%>" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value ="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" value ="<%=rsSetup("FCNAME_ENG")%>" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
	<%End if%>
<%Elseif Request("txttable") ="Jobbtmaster" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR BUDGET MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name (English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
		<%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                JOB FOR BUDGET MASTER</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value ="<%=rsSetup("FCCODE")%>" title="FCCODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME" value ="<%=rsSetup("FCNAME")%>" title="FCNAME Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Name(English)</font></td>
            <td align="left"><input type="text" size="50" maxlength="70"  name="FCNAME_ENG" value ="<%=rsSetup("FCNAME_ENG")%>" title="FCNAME(ENG) Character a s/d z or 1 s/d 9"> 
            </td>
          </Tr>
	<%End if%>
<%Elseif Request("txttable") ="GCMaster" then%>
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                GENERAL CHARGE SETUP</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCCODE" title="CODE Character a s/d z or 1 s/d 9">  <font color="#FF0000">*</font>
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Description</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCNAME" title="Description Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Item</font></td>
            <td align="left"><input type="text" size="10" maxlength="3" name="ITEM" title="ITEM Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
	<%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                GENERAL CHARGE SETUP</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCCODE" value="<%=rsSetup("FCCODE")%>" title="CODE Character a s/d z or 1 s/d 9">  <font color="#FF0000">*</font>
            </td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Description</font></td>
            <td align="left"><input type="text" size="30" maxlength="50"  name="FCNAME" value="<%=rsSetup("FCNAME")%>" title="Description Character a s/d z or 1 s/d 9"> 
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Item</font></td>
            <td align="left"><input type="text" size="10" maxlength="3" name="ITEM" value="<%=rsSetup("ITEM")%>" title="Item Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
	
	<%End if%>

<%else
            Dim k
			for k = 0 to  nCountofField-4%>
          <%
				if rsSetup(k).type = 135  then
					txtBoxColor="FFFF99"
					txtTitle   =rsSetup(k).name & " Type Date as mm/dd/yyyy"
				elseif rsSetup(k).type = 131  then
					txtBoxColor="FF0000"
					txtTitle   =rsSetup(i).name & " Numeric 1 s/d 9"
				else
					txtBoxColor="99FF99"
					txtTitle   =rsSetup(k).name & " Character a s/d z or 1 s/d 9"
				end if	 	 
				%>
			  <%if Request("txtJenis")="Edit" then%>
			  <tr> 
				<td align="left"><%=(rsSetup(k).name)%></td>
				<td><input type="text" value="<%=rsSetup(k)%>" name="<%=rsSetup(k).name%>"  align="left" 
						title="<%=txtTitle%>" style="background:#<%=txtBoxColor%>"></td>
			  </tr>
			  <%else%>
			  <tr> 
				<td align="left"><%=(rsSetup(k).name)%></td>
				<td><input type="text" value="<%=request(rsSetup(k).name)%>" name="<%=rsSetup(k).name%>" align="left" 
						title= "<%=txtTitle%>" style= "background:#<%=txtBoxColor%>"></td>
			  </tr>
			  <%end if%>
          <%next%>
<%end if%>

<script type="text/javascript">
function GetTxtGroupBSName()
	{
		 var pbulan = $("#JOB_FORBALANCESHEETCODE option:selected").text();
		 var result = pbulan.slice(pbulan.indexOf("-")+1);     
		 
		 document.getElementById('JOB_FORBALANCESHEETNAME').value = result;
	}
function GetTxtGroupISName()
	{
		 var pbulan = $("#JOB_FORPROFITANDLOSSCODE option:selected").text();
		 var result = pbulan.slice(pbulan.indexOf("-")+1);     
		 
		 document.getElementById('JOB_FORPROFITANDLOSSNAME').value = result;
	}
function GetTxtGroupCFName()
	{
		 var pbulan = $("#JOB_FORCASHFLOWCODE option:selected").text();
		 var result = pbulan.slice(pbulan.indexOf("-")+1);     
		 
		 document.getElementById('JOB_FORCASHFLOWNAME').value = result;
	}
function GetTxtGroupBTName()
	{
		 var pbulan = $("#JOB_FORBUDGETCODE option:selected").text();
		 var result = pbulan.slice(pbulan.indexOf("-")+1);     
		 
		 document.getElementById('JOB_FORBUDGETNAME').value = result;
	}
</script>