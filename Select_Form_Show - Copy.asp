
<% 
businessunit = session("fcba") 
%>

<%if Request("txttable") = "02-paymentvoucher" or Request("txttable") = "01-receivevoucher"  then%>
	<%
		transtype = ""
		banktype = ""
		
		transtype = Replace(txtTable, "voucher", "")
		
	%>
<a class="button_menu_top" href="paymentvoucherdetail.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>&transtype=<%=transtype%>&banktype=<%=banktype%>"
	title="New : <%=txtTable%>"><img src="ICON/Bank.png" border="0" >
    <span>Maintenance Voucher</span></a>
    <br><br>
    <span><font face="Verdana" size="1">Select Data Here :</font></span>
    <br>
    <table class="table_body">
		<tr>
			<td>
				<font face="Verdana" size="1">From</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			 </td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">To</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" onchange="javascript:getlistbank();">
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td> 
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Voucher Code</font>
			</td>
			<td>
				<input class="code" type="text" size="15" maxlength="30" name="VOUCHERCODE" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Bank</font>
			</td>
			<td>
				<div id="banklist"></div>
			</td>
		</tr>
		<tr>
			<!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
			<td>
				<font face="Verdana" size="1">Job</font>
			</td>
			<td>
				<select name="JOBCODE" class="chosen-select"><%cmdCallCombo "Job","","activation","Y"%></select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE" >
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="ITEMCODE" ></td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Status</font>
			</td>
			<td>
				<select name="rowstateid" id="rowstateid"> 
						<% 	q = " SELECT DISTINCT rowstateid " & _
							 	" FROM  paymentvoucher order by rowstateid desc"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("rowstateid")%>" 
									<%if cstr(""&rowstateid)=cstr(""&rsComboBST("rowstateid")) then %> selected 
									<%End If%>><%=rsComboBST("rowstateid")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td> 
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Paid For</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="remarks" >
			</td>
		</tr>
    </table>
    <br/><br/>
	<input type="submit" value="Preview" >
<%elseif Request("txttable") ="journalvoucher" then%>
<a class="button_menu_top" href="JournalVoucherDetail.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Bank.png" border="0" >
    <span>Maintenance Journal Voucher</span></a>
    <br><br>
    <span><font face="Verdana" size="1">Select Data Here :</font></span>
    <br>
   	<table class="table_body">
	    <tr>
	        <td><font face="Verdana" size="1">From</font></td>
	        <td>
	            <input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">
	            <!--<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	        </td>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">To</font></td>
	        <td>
	            <input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">
	            <!--<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	        </td>
	    </tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" onchange="javascript:getlistjournal();">
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td> 
		</tr>
	    <tr>
	        <td><font face="Verdana" size="1">Voucher Code</font></td>
	        <td>
	            <input class="code" type="text" size="30" maxlength="30" name="VOUCHERCODE">
	        </td>
	    </tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Bank</font>
			</td>
			<td>
				<div id="banklist"></div>
			</td>
		</tr>
	    
	    <tr>
	        <!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	        <td><font face="Verdana" size="1">Job</font></td>
	        <td>
	            <select name="JOBCODE" class="chosen-select">
	                <%cmdCallCombo "Job","","activation","Y"%>
	            </select>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">Target Type</font></td>
	        <td>
	            <select name="LOCATIONTYPE">
	                <%cmdCallComboSetup "LOCATIONTYPE", request("LOCATIONTYPE"),"",""%>
	            </select>
	        </td>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">Target Code</font></td>
	        <td>
	            <input type="text" size="30" maxlength="30" name="ITEMCODE">
	        </td>
	    </tr>
	</table>
    <br/><br/>
	<input type="submit" value="Preview" >
<%elseif Request("txttable") ="JournalReclass" then%>
<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Bank.png" border="0" >
    <span>Maintenance Journal Kredit</span></a>
    <br><br>
    <span><font face="Verdana" size="1">Select Data Here :</font></span>
    <br>
   	<table class="table_body">
	    <tr>
	        <td><font face="Verdana" size="1">From</font></td>
	        <td>
	            <input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">
	            <!--<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	        </td>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">To</font></td>
	        <td>
	            <input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">
	            <!--<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	        </td>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">Voucher Code</font></td>
	        <td>
	            <input class="code" type="text" size="30" maxlength="30" name="VOUCHERCODE">
	        </td>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">Bank</font></td>
	        <td>
	            <select name="BANKCODE">
	                <option value="88" <%if request( "BANKCODE")="88" then%> selected
	                    <%end if%>>JURNAL RECLASS</option>
	            </select>
	        </td>
	    </tr>
	    <tr>
	        <!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	        <td><font face="Verdana" size="1">Job</font></td>
	        <td>
	            <select name="JOBCODE" class="chosen-select">
	                <%cmdCallCombo "Job","","activation","Y"%>
	            </select>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">Target Type</font></td>
	        <td>
	            <select name="LOCATIONTYPE">
	                <%cmdCallComboSetup "LOCATIONTYPE", request("LOCATIONTYPE"),"",""%>
	            </select>
	        </td>
	    </tr>
	    <tr>
	        <td><font face="Verdana" size="1">Target Code</font></td>
	        <td>
	            <input type="text" size="30" maxlength="30" name="ITEMCODE">
	        </td>
	    </tr>
	</table>
    <br/><br/>
	<input type="submit" value="Preview" >
<%Elseif Request("txttable") ="ReclassVoucherDetail" then%>
    <a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Bank.png" border="0" >
    <span>Maintenance Reclass Voucher Activity</span></a>
	<br><br>
    
  
<%Elseif Request("txttable") ="Bank_Closing" then%>
<div class="close_box">
    <b>Bank Close</b></a><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Bank")) &"/"& Year(Session("v_Bank"))%> - <%=Session("CurP_Bank")%>
</div>

<%elseif Request("txttable") ="RolesMenu" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Roles Menu</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td>
				<font face="Verdana" size="1">Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="FCCODE">
			</td>
		</tr> 
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
 <%elseif Request("txttable") ="authorizetemplate" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Authorize Template</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td>
				<font face="Verdana" size="1">Authorize Templated Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="FCCODE">
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Module</font></td>
			<td align="left">
				<select name="document_type">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM document_type ORDER BY orderdata"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if document_type = rslov("fccode") then%>selected<%end if%>><%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Classification</font></td>
			<td align="left">
				<select name="document_classification">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM document_classification ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if cstr(""&document_classification) = cstr(""&rslov("fccode")) then%>selected<%end if%>><%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Authorize By Amount</font></td>
			<td align="left">
				<select name="authorizebyamount">
					<option value=""></option>
					<option value="TRUE" <%if isactive = "TRUE" then%>selected<%end if%>>TRUE</option>
					<option value="FALSE" <%if isactive = "FALSE" then%>selected<%end if%>>FALSE</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Is Active</font></td>
			<td align="left">
				<select name="isactive">
					<option value=""></option>
					<option value="TRUE" <%if isactive = "TRUE" then%>selected<%end if%>>TRUE</option>
					<option value="FALSE" <%if isactive = "FALSE" then%>selected<%end if%>>FALSE</option>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
  <%elseif Request("txttable") ="authorizelevel" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Authorize Level</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td>
				<font face="Verdana" size="1">Authorize Level Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="FCCODE">
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Start Level</font></td>
			<td align="left">
				<select id="startlevel" name="startlevel">
					<option value="FALSE" <%if startlevel = "FALSE" then%> selected <%end if%>>FALSE</option>
					<option value="TRUE" <%if startlevel = "TRUE" then%> selected <%end if%>>TRUE</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Final Level</font></td>
			<td align="left">
				<select id="finallevel" name="finallevel">
					<option value="FALSE" <%if finallevel = "FALSE" then%> selected <%end if%>>FALSE</option>
					<option value="TRUE" <%if finallevel = "TRUE" then%> selected <%end if%>>TRUE</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Is Active</font></td>
			<td align="left">
				<select id="isactive" name="isactive">
					<option value="FALSE" <%if isactive = "FALSE" then%> selected <%end if%>>FALSE</option>
					<option value="TRUE" <%if isactive = "TRUE" then%> selected <%end if%>>TRUE</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>

<%elseif Request("txttable") ="authorizestructure" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Authorize Structure</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td align="left"><font face="Verdana" size="1">Business Unit</font></td>
				<td align="left">
					<select name="businessunit" >
						<option value=""></option> 
						<% 	Dim rsComboASB
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboASB = Server.CreateObject("ADODB.Recordset")
							rsComboASB.CursorType = 3
							rsComboASB.CursorLocation = 3
							rsComboASB.Open q, connstr
							if not rsComboASB.eof then%>
							<%
							Do While Not rsComboASB.eof
							%>
								<option value="<%=rsComboASB("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboASB("fccode")) then %> selected 
									<%End If%>><%=rsComboASB("fccode")%>
								</option> 
							<%rsComboASB.movenext
							loop
							End if
							rsComboASB.Close
							set rsComboASB = Nothing
							q=""
						%>
					</select>
				</td>
			</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Authorize Template</font></td>
				<td align="left">
					<select name="authorizetemplatecode" >
						<option value=""></option> 
						<% 	Dim rsComboAS
							q = " SELECT Distinct aut.fccode" & _
								"        , aut.fcname" & _
								" FROM authorizetemplate aut" & _
								" LEFT JOIN document_classification dc ON (aut.document_classification = dc.fccode and aut.isactive=dc.isactive) " & _
								" LEFT JOIN document_type dt ON aut.document_type = dt.fccode" & _ 
                                " WHERE aut.isactive='TRUE' " & _ 
								" ORDER BY aut.fcname "

							' response.write q
							Set rsComboAS = Server.CreateObject("ADODB.Recordset")
							rsComboAS.CursorType = 3
							rsComboAS.CursorLocation = 3
							rsComboAS.Open q, connstr
							if not rsComboAS.eof then%>
							<%
							Do While Not rsComboAS.eof
							%>
								<option value="<%=rsComboAS("fccode")%>" 
									<%if cstr(""&authorizetemplatecode)=cstr(""&rsComboAS("fccode")) then %> selected 
									<%End If%>><%=rsComboAS("fccode")%>
								</option> 
							<%rsComboAS.movenext
							loop
							End if
							rsComboAS.Close
							set rsComboAS = Nothing
							q=""
						%>
					</select>
				</td>
			</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
<%elseif Request("txttable") ="PurchaseRequestDetails" then%>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1">PR Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	 
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PR Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="PRCODE"  style="background:#FFFFCC;">
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Item Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Requestor</font>
			</td>
			<td>
				<select name="REQUESTORID">
					<%cmdCallCombo "REQUESTOR","","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Status</font>
			</td>
			<td>				
				<select name="ROWSTATEID">
					<%cmdCallCombo "rowstatePurchaseRequest", "" ,"",""%>
				</select>
			</td>
		</tr>
		
		<tr>
			<td>
				 <input type="submit" value="Preview" >
			</td>
			<td>				
				<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Order.png" border="0">
	<span>Maintenance Purchase Request</span></a>
			</td>
		</tr>
	</table>	
   
 
	
<%elseif Request("txttable") ="PurchaseRequestAuthorization" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1">PR Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboPRA
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboPRA = Server.CreateObject("ADODB.Recordset")
						rsComboPRA.CursorType = 3
						rsComboPRA.CursorLocation = 3
						rsComboPRA.Open q, connstr
						if not rsComboPRA.eof then%>
						<%
						Do While Not rsComboPRA.eof
						%>
						<option value="<%=rsComboPRA("fccode")%>" 
							<%if cstr(""&businessunit)=cstr(""&rsComboPRA("fccode")) then %> selected 
							<%End If%>><%=rsComboPRA("fccode")%>
						</option> 
						<%rsComboPRA.movenext
						loop
						End if
						rsComboPRA.Close
						set rsComboPRA = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PR Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="PRCODE"  style="background:#FFFFCC">
			</td>
		</tr>		
	</table>
	
    <tr>
			<td>
				 <input type="submit" value="Preview" >
			</td>
			<td>				
				<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Order.png" border="0">
	<span>Maintenance Purchase Request Approval</span></a>
			</td>
		</tr>
	
<%elseif Request("txttable") ="PurchaseRequestApproval" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1">PR Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PR Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="PRCODE"  style="background:#FFFFCC">
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Item Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC">
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Requestor</font>
			</td>
			<td>
				<select name="REQUESTORID">
					<%cmdCallCombo "REQUESTOR","","",""%>
				</select>
			</td>
		</tr>

	
		<tr>
			<td>
				<font face="Verdana" size="1">Status</font>
			</td>
			<td>				
				<select name="ROWSTATEID">
					<%cmdCallCombo "rowstatePurchaseRequest", "" ,"",""%>
				</select>
			</td>
		</tr>			
	</table>
	
   <tr>
			<td>
				 <input type="submit" value="Preview" >
			</td>
			<td>				
				<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Order.png" border="0">
	<span>Maintenance Purchase Request Approval</span></a>
			</td>
		</tr>
	
<%elseif Request("txttable") ="PurchaseOrderDetails" then%>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body"> 
		<tr>
		    <td>
				<font face="Verdana" size="1">PO Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	 
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PO Number</font>
			</td>
			<td>		
				<input type="text" size="15" maxlength="30"  name="POCODE"  style="background:#FFFFCC">
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Item Code</font>
			</td>
			<td>				
				<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC">
			</td>	
		</tr>
		<tr>
			<td>	
				<font face="Verdana" size="1">Supplier</font>
			</td>
			<td>				
				<select name="SUPPLIERCODE" class="chosen-select">
				   <%cmdCallCombo "SUPPLIER","","",""%>
				</select>
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Status</font>
			</td>
			<td>				
				<select name="ROWSTATEID">
					<%cmdCallCombo "rowstatePurchaseOrder", "" ,"",""%>
				</select>
			<td>
		</tr>
		<tr>
			<td>
				 <input type="submit" value="Preview" >
			</td>
			<td>				
				<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Order.png" border="0">
	<span>Maintenance Purchase Order</span></a>
			</td>
		</tr>
	</table>
    
  <%elseif Request("txttable") ="PurchaseOrderAuthorization" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1">PO Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboPOA
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboPOA = Server.CreateObject("ADODB.Recordset")
						rsComboPOA.CursorType = 3
						rsComboPOA.CursorLocation = 3
						rsComboPOA.Open q, connstr
						if not rsComboPOA.eof then%>
						<%
						Do While Not rsComboPOA.eof
						%>
							<option value="<%=rsComboPOA("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboPOA("fccode")) then %> selected 
								<%End If%>><%=rsComboPOA("fccode")%>
							</option> 
						<%rsComboPOA.movenext
						loop
						End if
						rsComboPOA.Close
						set rsComboPOA = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PO Number</font>
			</td>
			<td>		
				<input type="text" size="15" maxlength="30"  name="POCODE"  style="background:#FFFFCC">
			</td>
		</tr>
		<tr>
			<td>
				 <input type="submit" value="Preview" >
			</td>
			<td>				
				<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Order.png" border="0" >
	<span>Maintenance Purchase Order Authorization</span></a>
			</td>
		</tr>
	</table>
    
<%elseif Request("txttable") ="PurchaseOrderApproval" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1">PO Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">PO Number</font>
			</td>
			<td>		
				<input type="text" size="15" maxlength="30"  name="POCODE"  style="background:#FFFFCC">
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Item Code</font>
			</td>
			<td>				
				<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC">
			</td>	
		</tr>
		<tr>
			<td>	
				<font face="Verdana" size="1">Supplier</font>
			</td>
			<td>				
				<select name="SUPPLIERCODE" class="chosen-select">
					<%cmdCallCombo "SUPPLIER","","",""%>
				</select>
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Status</font>
			</td>
			<td>				
				<select name="ROWSTATEID">
					<%cmdCallCombo "rowstatePurchaseOrder", "" ,"",""%>
				</select>
			<td>
		</tr>
		<tr>
			<td>
				 <input type="submit" value="Preview" >
			</td>
			<td>				
				<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Order.png" border="0" >
	<span>Maintenance Purchase Order Approval</span></a>
			</td>
		</tr>
	</table>
 
    
<%elseif Request("txttable") ="StockTransferOrder" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Stock Transfer Order</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="STODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="STODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr> 
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	Dim rsComboBST
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">STO No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="stono" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PO No.</font>
			</td>
				<td align="left">
					<input type="text" maxlength="50" name="ponumber" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Transporter Code</font>
			</td>
				<td align="left">
					<input type="text" maxlength="50" name="transportercode" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Rent By</font>
			</td>
			<td>
				<select name="rentby">
					<option value=""></option> 
					<% 	Dim rsComboRB
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboRB = Server.CreateObject("ADODB.Recordset")
						rsComboRB.CursorType = 3
						rsComboRB.CursorLocation = 3
						rsComboRB.Open q, connstr
						if not rsComboRB.eof then%>
						<%
						Do While Not rsComboRB.eof
						%>
							<option value="<%=rsComboRB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboRB("fccode")) then %> selected 
								<%End If%>><%=rsComboRB("fccode")%>
							</option> 
						<%rsComboRB.movenext
						loop
						End if
						rsComboRB.Close
						set rsComboRB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Term Of Payment</font>
			</td>
				<td align="left">
					<select name="termofpayment" value="<%=termofpayment%>">
						<option value="<%=termofpayment%>"></option>
						<option value="TOP7" <% If cstr(""&termofpayment) = "TOP7" Then %> Selected <%End if%>>TOP7</option> 
						<option value="TOP14" <% If cstr(""&termofpayment) = "TOP14" Then %> Selected <%End if%>>TOP14</option> 
						<option value="TOP30" <% If cstr(""&termofpayment) = "TOP30" Then %> Selected <%End if%>>TOP30</option> 
					</select>
				</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="servicepurchaseorder" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Service Purchase Order</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr> 
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	Dim rsComboBU
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBU = Server.CreateObject("ADODB.Recordset")
							rsComboBU.CursorType = 3
							rsComboBU.CursorLocation = 3
							rsComboBU.Open q, connstr
							if not rsComboBU.eof then%>
							<%
							Do While Not rsComboBU.eof
							%>
								<option value="<%=rsComboBU("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBU("fccode")) then %> selected 
									<%End If%>><%=rsComboBU("fccode")%>
								</option> 
							<%rsComboBU.movenext
							loop
							End if
							rsComboBU.Close
							set rsComboBU = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">SPO No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="spono" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Vendor</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="suppliercode" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Term Of Payment</font>
			</td>
				<td align="left">
					<select name="termofpayment" value="<%=termofpayment%>">
						<option value="<%=termofpayment%>"></option>
						<option value="TOP7" <% If cstr(""&termofpayment) = "TOP7" Then %> Selected <%End if%>>TOP7</option> 
						<option value="TOP14" <% If cstr(""&termofpayment) = "TOP14" Then %> Selected <%End if%>>TOP14</option> 
						<option value="TOP30" <% If cstr(""&termofpayment) = "TOP30" Then %> Selected <%End if%>>TOP30</option> 
					</select>
				</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Purchase By</font>
			</td>
			<td>
				<select name="purchaseby">
						<option value=""></option> 
						<% 	Dim rsComboPB
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboPB = Server.CreateObject("ADODB.Recordset")
							rsComboPB.CursorType = 3
							rsComboPB.CursorLocation = 3
							rsComboPB.Open q, connstr
							if not rsComboPB.eof then%>
							<%
							Do While Not rsComboPB.eof
							%>
								<option value="<%=rsComboPB("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboPB("fccode")) then %> selected 
									<%End If%>><%=rsComboPB("fccode")%>
								</option> 
							<%rsComboPB.movenext
							loop
							End if
							rsComboPB.Close
							set rsComboPB = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="deliveryorder_purchase" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Delivery Purchase Order</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">DO Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DO_DATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DO_DATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr> 
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	Dim rsComboB
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboB = Server.CreateObject("ADODB.Recordset")
							rsComboB.CursorType = 3
							rsComboB.CursorLocation = 3
							rsComboB.Open q, connstr
							if not rsComboB.eof then%>
							<%
							Do While Not rsComboB.eof
							%>
								<option value="<%=rsComboB("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboB("fccode")) then %> selected 
									<%End If%>><%=rsComboB("fccode")%>
								</option> 
							<%rsComboB.movenext
							loop
							End if
							rsComboB.Close
							set rsComboB = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">DO No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="do_no" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PO No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="ponumber" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="AuthorizeTransaction" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
			<td align="left">
				<font face="Verdana" size="1">From</font>
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">To</font>
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboAT
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboAT = Server.CreateObject("ADODB.Recordset")
						rsComboAT.CursorType = 3
						rsComboAT.CursorLocation = 3
						rsComboAT.Open q, connstr
						if not rsComboAT.eof then%>
						<%
						Do While Not rsComboAT.eof
						%>
						<option value="<%=rsComboAT("fccode")%>" 
							<%if cstr(""&businessunit)=cstr(""&rsComboAT("fccode")) then %> selected 
							<%End If%>><%=rsComboAT("fccode")%>
						</option> 
						<%rsComboAT.movenext
						loop
						End if
						rsComboAT.Close
						set rsComboAT = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PR Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="PRCODE"  style="background:#FFFFCC">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Purchase By</font>
			</td>
			<td>
				<select name="ManagedId" >
				    <%cmdCallCombo "PURCHASEMANAGER",ManagedId,"",""%>
			    </select>
			</td>
		</tr>		
	</table>
	
    <input type="submit" value="Preview" >

    <%elseif Request("txttable") ="AuthorizeTransactionPV" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
			<td align="left">
				<font face="Verdana" size="1">From</font>
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="VoucherDate__From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">To</font>
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="VoucherDate__To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	 
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboAT = Server.CreateObject("ADODB.Recordset")
						rsComboAT.CursorType = 3
						rsComboAT.CursorLocation = 3
						rsComboAT.Open q, connstr
						if not rsComboAT.eof then%>
						<%
						Do While Not rsComboAT.eof
						%>
						<option value="<%=rsComboAT("fccode")%>" 
							<%if cstr(""&businessunit)=cstr(""&rsComboAT("fccode")) then %> selected 
							<%End If%>><%=rsComboAT("fccode")%>
						</option> 
						<%rsComboAT.movenext
						loop
						End if
						rsComboAT.Close
						set rsComboAT = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>

		<tr>
			<td>
				<font face="Verdana" size="1">Voucher Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="vouchercode"  style="background:#FFFFCC">
			</td>
		</tr>
		<tr>
	        <td>
                <font face="Verdana" size="1">Bank Code</font>
            </td>
            <td align="left">
				<input type="hidden" id="transtype" name="transtype" value="<%=transtype%>">
				<input type="hidden" name="banktype" value="<%=banktype%>">
                <select name="xCode" id ="bankcode" onchange="javascript:reloadCode('<%=request("txtTable")%>','<%=transtype%>');">
				<option></option>
				<%
				q = " SELECT a.fccode, a.fcname" & _
					"  FROM bank a" & _
					" WHERE a.fcba = '"& session ("fcBA") &"'" & _
					"       AND a.activation = 'Y'" & _
				    "       AND UPPER(a.banktype) <> ('JOURNALRECLASS')" 

				if UCase(cstr(""&banktype)) = "JOURNAL" then
					q = q & "       AND UPPER(banktype) = UPPER('JOURNAL')"
				else
					q = q & "       AND UPPER(banktype) IN ('CASH', 'BANK')"
				end if
				' response.write q
				Set rsComboA = Server.CreateObject("ADODB.Recordset")
				
				rsComboA.Open q, connstr
				if not rsComboA.eof then
					Do While Not rsComboA.eof
				%>
					<option value="<%=rsComboA("fccode")%>" <% If cstr(""&xCode) = cstr(""&rsComboA("fccode")) Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
				<% 
						rsComboA.movenext
					loop
				End if
				rsComboA.Close : set rsComboA = Nothing : q = ""
				%>
                </select>
				<font color="#FF0000">&nbsp;*</font>
            </td>
	    </tr>
	    <tr>
			<td>
				<font face="Verdana" size="1">Status</font>
			</td>
			<td>
				<select name="rowstateid" id="rowstateid"> 
						<% 	q = " SELECT DISTINCT rowstateid " & _
							 	" FROM  paymentvoucher order by rowstateid desc"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("rowstateid")%>" 
									<%if cstr(""&rowstateid)=cstr(""&rsComboBST("rowstateid")) then %> selected 
									<%End If%>><%=rsComboBST("rowstateid")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td> 
		</tr>		
	</table>
	
    <input type="submit" value="Preview" >

<%elseif Request("txttable") ="AuthorizeTransactionPO" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
			<td align="left">
				<font face="Verdana" size="1">From</font>
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="PODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">To</font>
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="PODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboATPO
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboATPO = Server.CreateObject("ADODB.Recordset")
						rsComboATPO.CursorType = 3
						rsComboATPO.CursorLocation = 3
						rsComboATPO.Open q, connstr
						if not rsComboATPO.eof then%>
						<%
						Do While Not rsComboATPO.eof
						%>
						<option value="<%=rsComboATPO("fccode")%>" 
							<%if cstr(""&businessunit)=cstr(""&rsComboATPO("fccode")) then %> selected 
							<%End If%>><%=rsComboATPO("fccode")%>
						</option> 
						<%rsComboATPO.movenext
						loop
						End if
						rsComboATPO.Close
						set rsComboATPO = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PO Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="POCODE"  style="background:#FFFFCC">
			</td>
		</tr>		
	</table>
	
    <input type="submit" value="Preview" >
<%elseif Request("txttable") ="transporterinvoice" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Transporter Invoice</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="CREATIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="CREATIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr> 
		<tr>
			<td>
				<font face="Verdana" size="1">Invoice By</font>
			</td>
			<td>
				<select name="invoice_by">
						<option value=""></option> 
						<%  q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBTI = Server.CreateObject("ADODB.Recordset")
							rsComboBTI.CursorType = 3
							rsComboBTI.CursorLocation = 3
							rsComboBTI.Open q, connstr
							if not rsComboBTI.eof then%>
							<%
							Do While Not rsComboBTI.eof
							%>
								<option value="<%=rsComboBTI("fccode")%>" 
									<%if cstr(""&invoice_by)=cstr(""&rsComboBTI("fccode")) then %> selected 
									<%End If%>><%=rsComboBTI("fccode")%>
								</option> 
							<%rsComboBTI.movenext
							loop
							End if
							rsComboBTI.Close
							set rsComboBTI = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	Dim rsComboBTI
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBTI = Server.CreateObject("ADODB.Recordset")
							rsComboBTI.CursorType = 3
							rsComboBTI.CursorLocation = 3
							rsComboBTI.Open q, connstr
							if not rsComboBTI.eof then%>
							<%
							Do While Not rsComboBTI.eof
							%>
								<option value="<%=rsComboBTI("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBTI("fccode")) then %> selected 
									<%End If%>><%=rsComboBTI("fccode")%>
								</option> 
							<%rsComboBTI.movenext
							loop
							End if
							rsComboBTI.Close
							set rsComboBTI = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Inovice No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="invoicecode" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contract No.</font>
			</td>
				<td align="left">
					<input type="text" maxlength="50" name="contractno" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Transporter</font>
			</td>
			<td>
				<select name="transportercode">
						<option value=""></option> 
						<% 	 
							q = " SELECT DISTINCT fccode, fcname " & _
							 	" FROM transporter WHERE fcba = 'PSR'" 
							' response.write q
							Set rsComboBTI = Server.CreateObject("ADODB.Recordset")
							rsComboBTI.CursorType = 3
							rsComboBTI.CursorLocation = 3
							rsComboBTI.Open q, connstr
							if not rsComboBTI.eof then%>
							<%
							Do While Not rsComboBTI.eof
							%>
								<option value="<%=rsComboBTI("fccode")%>" 
									<%if cstr(""&transportercode)=cstr(""&rsComboBTI("fccode")) then %> selected 
									<%End If%>><%=rsComboBTI("fccode")%> : <%=rsComboBTI("fcname")%>
								</option> 
							<%rsComboBTI.movenext
							loop
							End if
							rsComboBTI.Close
							set rsComboBTI = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="contractorinvoice" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Contractor Invoice</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="CREATIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="CREATIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr>  
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	 
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBTI = Server.CreateObject("ADODB.Recordset")
							rsComboBTI.CursorType = 3
							rsComboBTI.CursorLocation = 3
							rsComboBTI.Open q, connstr
							if not rsComboBTI.eof then%>
							<%
							Do While Not rsComboBTI.eof
							%>
								<option value="<%=rsComboBTI("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBTI("fccode")) then %> selected 
									<%End If%>><%=rsComboBTI("fccode")%>
								</option> 
							<%rsComboBTI.movenext
							loop
							End if
							rsComboBTI.Close
							set rsComboBTI = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Inovice No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="invoicecode" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contract No.</font>
			</td>
				<td align="left">
					<input type="text" maxlength="50" name="contractno" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Transporter</font>
			</td>
			<td>
				<select name="transportercode">
						<option value=""></option> 
						<% 	 
							q = " SELECT DISTINCT fccode, fcname " & _
							 	" FROM transporter WHERE fcba = 'PSR'" 
							' response.write q
							Set rsComboBTI = Server.CreateObject("ADODB.Recordset")
							rsComboBTI.CursorType = 3
							rsComboBTI.CursorLocation = 3
							rsComboBTI.Open q, connstr
							if not rsComboBTI.eof then%>
							<%
							Do While Not rsComboBTI.eof
							%>
								<option value="<%=rsComboBTI("fccode")%>" 
									<%if cstr(""&transportercode)=cstr(""&rsComboBTI("fccode")) then %> selected 
									<%End If%>><%=rsComboBTI("fccode")%> : <%=rsComboBTI("fcname")%>
								</option> 
							<%rsComboBTI.movenext
							loop
							End if
							rsComboBTI.Close
							set rsComboBTI = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="contractorinvoicedp" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Contractor Invoice DP</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="CREATIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="CREATIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr>  
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	 
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBTI = Server.CreateObject("ADODB.Recordset")
							rsComboBTI.CursorType = 3
							rsComboBTI.CursorLocation = 3
							rsComboBTI.Open q, connstr
							if not rsComboBTI.eof then%>
							<%
							Do While Not rsComboBTI.eof
							%>
								<option value="<%=rsComboBTI("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBTI("fccode")) then %> selected 
									<%End If%>><%=rsComboBTI("fccode")%>
								</option> 
							<%rsComboBTI.movenext
							loop
							End if
							rsComboBTI.Close
							set rsComboBTI = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="Invoicesto" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Invoice STO</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr> 
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	Dim rsComboINVSTO
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboINVSTO = Server.CreateObject("ADODB.Recordset")
							rsComboINVSTO.CursorType = 3
							rsComboINVSTO.CursorLocation = 3
							rsComboINVSTO.Open q, connstr
							if not rsComboINVSTO.eof then%>
							<%
							Do While Not rsComboINVSTO.eof
							%>
								<option value="<%=rsComboINVSTO("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboINVSTO("fccode")) then %> selected 
									<%End If%>><%=rsComboINVSTO("fccode")%>
								</option> 
							<%rsComboINVSTO.movenext
							loop
							End if
							rsComboINVSTO.Close
							set rsComboINVSTO = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Inovice No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="invoicecode" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Transporter Code</font>
			</td>
			<td align="left">
				<select name="transportercode">
						<option value=""></option> 
						<% 	Dim rsComboTc
							q = " select distinct fccode, fcname from transporter where fcba in ('HOA','KNS','KKU','PSR')"
							' response.write q
							Set rsComboTc = Server.CreateObject("ADODB.Recordset")
							rsComboTc.CursorType = 3
							rsComboTc.CursorLocation = 3
							rsComboTc.Open q, connstr
							if not rsComboTc.eof then%>
							<%
							Do While Not rsComboTc.eof
							%>
								<option value="<%=rsComboTc("fccode")%>" 
									<%if cstr(""&transportercode)=cstr(""&rsComboTc("fccode")) then %> selected 
									<%End If%>><%=rsComboTc("fccode")%> : <%=rsComboTc("fcname")%> 
								</option> 
							<%rsComboTc.movenext
							loop
							End if
							rsComboTc.Close
							set rsComboTc = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="ReceiveNote" then%>

    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    
    <tr>
            <td>
				<font face="Verdana" size="1">Date</font>
			</td>
		    <td>
	<font face="Verdana" size="1"> 
	<input class="datepicker" type="text" maxlength="10" name="RNDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" ><font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="RNDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
            </td>
		</tr>
		
		<tr>
			<td>
	<font face="Verdana" size="1">RN Number</font> </td><td><input type="text" size="15" maxlength="30"  name="RECEIVENOTECODE"  style="background:#FFFFCC">
                </td>
		</tr>
		
		<tr>
			<td>
	<font face="Verdana" size="1">Item Code</font></td><td>	<input type="text" size="15" maxlength="30" name="PURCHASEITEMCODE"  style="background:#FFFFCC">
                </td>
		</tr>
		
		<tr>
			<td>
    <font face="Verdana" size="1">Supplier</font>
    </td><td>
	<select name="SUPPLIERCODE" class="chosen-select">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
                </td>
			<tr>

		<tr>
			<td>
				<font face="Verdana" size="1">Jobcode</font>
			</td>
			<td>
				<select name="jobcode" class="chosen-select">
					<%cmdCallCombo "JOB",Request("jobcode"),"activation","Y"%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				 <input type="submit" value="Preview" >
			</td>
			<td>				
				<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/note.png" border="0" >
	<span>Maintenance Receive Note</span></a>
			</td>
		</tr>
		
    </table>
    
<%elseif Request("txttable") ="Invoice" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/invoice.png" border="0" >
	<span>Maintenance Invoice</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	 
    <table>
		<tr>
            <td>
				<font face="Verdana" size="1">Date</font>
			</td>
		    <td>
				<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >				
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Supplier</font>
			</td>
			<td>
				<select name="SUPPLIERCODE" class="chosen-select">
					<%cmdCallCombo "SUPPLIER","","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Invoice Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="INVOICECODE"  style="background:#FFFFCC">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Item Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC">
			</td>
		</tr>
		<tr>
			<td rowspan="2">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
	
<%elseif Request("txttable") ="paymentrequest" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/invoice.png" border="0" >
	<span>Maintenance Invoice</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	 
    <table>
		<tr>
            <td>
				<font face="Verdana" size="1">Date</font>
			</td>
		    <td>
				<input class="datepicker" type="text" maxlength="10" name="transdate_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >				
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="transdate_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		
		<tr>
			<td rowspan="2">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
	
<%elseif Request("txttable") ="InvoiceService" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Special Invoice</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
		    <td>
				<font face="Verdana" size="1">Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="TRANSDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="TRANSDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr> 
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	Dim rsComboSI
							q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboSI = Server.CreateObject("ADODB.Recordset")
							rsComboSI.CursorType = 3
							rsComboSI.CursorLocation = 3
							rsComboSI.Open q, connstr
							if not rsComboSI.eof then%>
							<%
							Do While Not rsComboSI.eof
							%>
								<option value="<%=rsComboSI("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboSI("fccode")) then %> selected 
									<%End If%>><%=rsComboSI("fccode")%>
								</option> 
							<%rsComboSI.movenext
							loop
							End if
							rsComboSI.Close
							set rsComboSI = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Invoice</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="invoiceno" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PO No.</font>
			</td>
			<td>
				<input type="text" maxlength="50" name="ponumber" size="30" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="invoicetbs" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/invoice.png" border="0" >
	<span>Maintenance Invoice TBS</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	 
    <table>
		<tr>
            <td>
				<font face="Verdana" size="1">Date</font>
			</td>
		    <td>
				<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >				
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
					</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Supplier</font>
			</td>
			<td>
				<select name="suppliercode">
						<option value=""></option> 
						<% 	q = "select customercode, description from customer where fcba = 'ARM' and supplier = 'Y'order by customercode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("customercode")%>" 
									<%if cstr(""&suppliercode)=cstr(""&rsComboBST("customercode")) then %> selected 
									<%End If%>><%=rsComboBST("customercode")%> : <%=rsComboBST("description")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Invoice Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="INVOICECODE"  style="background:#FFFFCC">
			</td>
		</tr> 
		<tr>
			<td rowspan="2">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
<%Elseif Request("txttable") ="Purchasing_Closing" then%>
<div class="close_box">
    <b>Purchasing Close</b></a><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Purchase")) &"/"& Year(Session("v_Purchase"))%>
    </div>
<%elseif Request("txttable") ="IssueDetails" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/stock.png" border="0" >
	<span>Maintenance Stock Issue</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
     <table class="table_body">
    <tr>
        <td><font face="Verdana" size="1">From </font>
	</td>
    <td>
    <input class="datepicker" type="text" maxlength="10" name="SIVDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<font face="Verdana" size="1">To</font>
    <input class="datepicker" type="text" maxlength="10" name="SIVDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
    </td>
    <td>
	</td>
    </tr>
    <tr>
		<td>
			<font face="Verdana" size="1">Business Unit</font>
		</td>
		<td>
			<select name="businessunit" id="businessunit" >
					<option value=""></option> 
					<% 	q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboBST = Server.CreateObject("ADODB.Recordset")
						rsComboBST.CursorType = 3
						rsComboBST.CursorLocation = 3
						rsComboBST.Open q, connstr
						if not rsComboBST.eof then%>
						<%
						Do While Not rsComboBST.eof
						%>
							<option value="<%=rsComboBST("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
								<%End If%>><%=rsComboBST("fccode")%>
							</option> 
						<%rsComboBST.movenext
						loop
						End if
						rsComboBST.Close
						set rsComboBST = Nothing
						q=""
					%>
			</select>
		</td>
	</tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Issue No</font>
    </td>
    <td><input type="text" size="15" maxlength="30"  name="SIVCODE"  style="background:#FFFFCC">
	</tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Item Code</font>
    </td>
    <td><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#FFFFCC">
	</td>
    </tr>
    <!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	<tr>
    <td>
    <font face="Verdana" size="1">Job</font>
    </td>
    <td><select name="JOBCODE" class="chosen-select"><%cmdCallCombo "Job","","activation","Y"%></select>
    </td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Target Type</font>
    </td>
    <td>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<tr>
    <td>
    <font face="Verdana" size="1">Target Code</font>
    </td>
    <td>
    <input type="text" size="5" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
	</td>
    </tr>
    <tr>
    <td>
    <input type="submit" value="Preview" >
    </td>
    </tr>
    </table>
    
    
<%elseif Request("txttable") ="WriteOff" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/WriteOff.jpg" border="0" >
	<span>Maintenance Write Off</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
      <table class="table_body">
    <tr>
        <td><font face="Verdana" size="1">From </font>
	</td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="WODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="WODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
		<td>
			<font face="Verdana" size="1">Business Unit</font>
		</td>
		<td>
			<select name="businessunit" id="businessunit" >
					<option value=""></option> 
					<% 	q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboBST = Server.CreateObject("ADODB.Recordset")
						rsComboBST.CursorType = 3
						rsComboBST.CursorLocation = 3
						rsComboBST.Open q, connstr
						if not rsComboBST.eof then%>
						<%
						Do While Not rsComboBST.eof
						%>
							<option value="<%=rsComboBST("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
								<%End If%>><%=rsComboBST("fccode")%>
							</option> 
						<%rsComboBST.movenext
						loop
						End if
						rsComboBST.Close
						set rsComboBST = Nothing
						q=""
					%>
			</select>
		</td>
	</tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Writeoff Code</font>
    </td>
    <td>
    <input type="text" size="15" maxlength="30"  name="WriteOffCode"  style="background:#FFFFCC">
    </td>
    </tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Item Code</font>
    </td>
    <td>
    <input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#FFFFCC">
    </td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Target Type</font>
    </td>
    <td>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
    </td>
    </tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Target Code</font>
    </td>
    <td>
    <input type="text" size="5" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
    </td></tr></table>
	<input type="submit" value="Preview" >

<%elseif Request("txttable") ="transporttask" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/stock.png" border="0" >
	<span>Maintenance Transport Task</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
     <table class="table_body">
    <tr>
        <td><font face="Verdana" size="1">From </font>
	</td>
    <td>
    <input class="datepicker" type="text" maxlength="10" name="ORDERDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<font face="Verdana" size="1">To</font>
    <input class="datepicker" type="text" maxlength="10" name="ORDERDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
    </td>
    <td>
	</td>
    </tr>
	    <tr>
		<td align="left"><font face="Verdana" size="1">Business Unit</font></td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboTT
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboTT = Server.CreateObject("ADODB.Recordset")
						rsComboTT.CursorType = 3
						rsComboTT.CursorLocation = 3
						rsComboTT.Open q, connstr
						if not rsComboTT.eof then%>
						<%
						Do While Not rsComboTT.eof
						%>
							<option value="<%=rsComboTT("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboTT("fccode")) then %> selected 
								<%End If%>><%=rsComboTT("fccode")%>
							</option> 
						<%rsComboTT.movenext
						loop
						End if
						rsComboTT.Close
						set rsComboTT = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	    <tr>
		<td align="left"><font face="Verdana" size="1">Business Unit Destination</font></td>
			<td align="left">
				<select name="fcbadistination" >
					<option value=""></option> 
					<% 	Dim rsComboTD
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboTD = Server.CreateObject("ADODB.Recordset")
						rsComboTD.CursorType = 3
						rsComboTD.CursorLocation = 3
						rsComboTD.Open q, connstr
						if not rsComboTD.eof then%>
						<%
						Do While Not rsComboTD.eof
						%>
							<option value="<%=rsComboTD("fccode")%>" 
								<%if cstr(""&fcbadistination)=cstr(""&rsComboTD("fccode")) then %> selected 
								<%End If%>><%=rsComboTD("fccode")%>
							</option> 
						<%rsComboTD.movenext
						loop
						End if
						rsComboTD.Close
						set rsComboTD = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>

    <tr>
    <td>
	<font face="Verdana" size="1">Order No</font>
    </td>
    <td><input type="text" size="25" maxlength="30"  name="ORDERNO"  style="background:#FFFFCC">
	</tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Item Code</font>
    </td>
    <td><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#FFFFCC">
	</td>
    </tr>
    <tr>
	  <td>
		<font face="Verdana" size="1">Job Code</font>
	  </td>
			
			<% 	
				    q = "select fccode jobcode, fcname Jobname,fcba  from INTRACO_TAB where companycode= businessunit_api.get_CompanyCode('" &  Session("fcba") & "')  and status='TRUE'  order by fccode"
					set rsCombo = Server.CreateObject("ADODB.Recordset")
					'response.write q
					rsCombo.CursorType = 3
					rsCombo.CursorLocation = 3
					rsCombo.Open q, connstr
					if not rsCombo.eof then
				%>
				
				<td align="left"> 
					<select name="JobCode" >    
					<option value="" selected></option>
								<% Do While Not rsCombo.eof %>
									  <option data-fcba="<%=rsCombo("fcba")%>" value="<%=rsCombo("jobcode")%>" <% If Request("JobCode") = rsCombo("jobcode") Then %> Selected <%End if%>><%=rsCombo("jobcode")%> : <%=rsCombo("jobname")%></option> 
								<%    rsCombo.movenext
								   loop

								
					End if
					
					rsCombo.Close
					set rsCombo = Nothing
					q=""
				%>
					</select> <font color="#FF0000">*</font>
			</td>
		</tr>
    <tr>
    <td>
    <input type="submit" value="Preview" >
    </td>
    </tr>
    </table>

<%elseif Request("txttable") ="transporttaskdiv" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/stock.png" border="0" >
	<span>Maintenance Transport Task Divisi</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
     <table class="table_body">
    <tr>
        <td><font face="Verdana" size="1">From </font>
	</td>
    <td>
    <input class="datepicker" type="text" maxlength="10" name="ORDERDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<font face="Verdana" size="1">To</font>
    <input class="datepicker" type="text" maxlength="10" name="ORDERDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
    </td>
    <td>
	</td>
    </tr>
	    <tr>
		<td align="left"><font face="Verdana" size="1">Business Unit</font></td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	'Dim rsComboTT
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboTT = Server.CreateObject("ADODB.Recordset")
						rsComboTT.CursorType = 3
						rsComboTT.CursorLocation = 3
						rsComboTT.Open q, connstr
						if not rsComboTT.eof then%>
						<%
						Do While Not rsComboTT.eof
						%>
							<option value="<%=rsComboTT("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboTT("fccode")) then %> selected 
								<%End If%>><%=rsComboTT("fccode")%>
							</option> 
						<%rsComboTT.movenext
						loop
						End if
						rsComboTT.Close
						set rsComboTT = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	    

    <tr>
    <td>
	<font face="Verdana" size="1">Order No</font>
    </td>
    <td><input type="text" size="25" maxlength="30"  name="ORDERNO"  style="background:#FFFFCC">
	</tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Item Code</font>
    </td>
    <td><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#FFFFCC">
	</td>
    </tr>
    
    <td>
    <input type="submit" value="Preview" >
    </td>
    </tr>
    </table>


<%elseif Request("txttable") ="ReceivedTask" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/stock.png" border="0" >
	<span>Maintenance Received Task</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
    <table class="table_body">
	    <tr>
	        <td>
	        	<font face="Verdana" size="1">From </font>
			</td>
	    	<td>
	    		<input class="datepicker" type="text" maxlength="10" name="ORDERDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
	    		<input class="datepicker" type="text" maxlength="10" name="ORDERDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	    	</td>
    		<td>
			</td>
    	</tr>
    	<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
    	<tr>
    		<td>
				<font face="Verdana" size="1">Order No</font>
			</td>
    		<td>
    			<input type="text" size="25" maxlength="30"  name="ORDERNO"  style="background:#FFFFCC">
    		</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Fcba Source</font>
			</td>
			<td align="left" readonly="RUE" > 
		        <%if SudahClose="Y" then%>
				    <select name="FCBASOURCE" ><option disabled="disabled" value="<%=FCBASOURCE%>" Selected><%=FCBASOURCE%></option></select>
				<%else%>   		
		    
		    <% 
			dim rsCallCombo
			q = " select fccode,fcname " & _
							 "	from businessunit " & _
							 "	where fccompanycode = businessunit_api.get_CompanyCode('" & session("fcba") & "') " & _
							 "	and businessunit_api.Isactive(fccode)='TRUE'  and fccode <> '" & session("fcba") & "' order by orderdata "
							
			'response.write q
			Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
			rsCallCombo.CursorType = 3
			rsCallCombo.CursorLocation = 3
			rsCallCombo.Open q	,connstr			 
			if not rsCallCombo.eof then
			%>
			<select name="FCBASOURCE" >    
			<option value="" selected></option>
			<%
			 Do While Not rsCallCombo.eof
			    %> 
				   
				   <option value="<%=rsCallCombo("Fccode")%>" <% If Request("xCode") = rsCallCombo("Fccode") Then %> Selected <%End if%>><%=rsCallCombo("Fccode")%> : <%=rsCallCombo("fcname")%></option> 
				<%
				rsCallCombo.movenext
			 loop
			End if
			rsCallCombo.Close
			set rsCallCombo = Nothing
			%>
            </select>
	       <%End if%>
		<font color="#FF0000">&nbsp;*</font>
        </td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Item Code</font>
    </td>
    <td><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#FFFFCC">
	</td>
    </tr>
    
    <tr>
    <td>
    <input type="submit" value="Preview" >
    </td>
    </tr>
    </table>    


<%Elseif Request("txttable") ="StockLedger" then%>
<div class="close_box">
    <b>Stock Ledger Report</b><br/><br/>
    <img src="ICON/StockLedger.jpg" border="0" ></a>
</div>    
	<br/><br/>
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<form action="p_stockledger.asp" method="post" target="_blank">
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">Store Code</font></td><td><input type=text name=txtstorecode></td>
		<td><font face="verdana" size="1">STOCKGROUP</font></td><td><input type=text name=txtstockgroup></td>
		<td><font face="verdana" size="1">STOCK</font></td><td><input type=text name=txtstock></td>
		<td><input type=submit value="Preview" ></td>
	</tr>
	</table>
	</form>
    
<%Elseif Request("txttable") ="Store_Closing_New" then%>
    <div class="close_box">
    <b>Store Close</b></a><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Stores")) &"/"& Year(Session("v_Stores"))%>
    </div>
    
<%elseif Request("txttable") ="Attendance_GAD" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Employee.jpg" border="0" >
	<span>Maintenance GangActivity & Attendance</span></a>
    <br><br>
<table>
<script>
	function checkAll() {
		var checkbox_options = document.getElementById("checkbox_options");
		var checkbox_all = document.getElementById("checkbox_all");

		var checkboxes = checkbox_options.getElementsByTagName("input");        
		for (i = 0; i < checkboxes.length; i++) {
			checkboxes[i].checked = checkbox_all.checked;
		}
	}
	
	function checkAllAttendance() {
		var checkbox_options_attendance = document.getElementById("checkbox_options_attendance");
		var checkbox_all_attendance = document.getElementById("checkbox_all_attendance");

		var checkboxes_attendance = checkbox_options_attendance.getElementsByTagName("input");        
		for (i = 0; i < checkboxes_attendance.length; i++) {
			checkboxes_attendance[i].checked = checkbox_all_attendance.checked;
		}
	}
</script>
	<tr> 
		<td colspan="3">
		<font face="Verdana" size="1">Select Field :</Font><br/>		
		<div class="ckbox_list">
			<table>
				<tr>
					<td><input id="checkbox_all" type="checkbox" onclick="checkAll()" /><span><font face="Verdana" size="1">Select All</Font></span></td>
				</tr>
			</table>
			<div id = "checkbox_options" >
				<table>
					<tr>
						<td><input name="cField" type="checkbox" value="fddate" checked /><span><font face="Verdana" size="1">Date</font></span></td>
						<td><input name="cField" type="checkbox" value="gangcode" checked /><span><font face="Verdana" size="1">Gang Code</font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_1" /><span><font face="Verdana" size="1">Supervision 1</Font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_2" /><span><font face="Verdana" size="1">Supervision 2</Font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_3" /><span><font face="Verdana" size="1">Supervision 3</Font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_4" /><span><font face="Verdana" size="1">Supervision 4</Font></span></td>
					</tr>
					<tr>
						<td><input name="cField" type="checkbox" value="sectionname" checked /><span><font face="Verdana" size="1">Section Name</font></span></td>
						<td><input name="cField" type="checkbox" value="employeecode" checked /><span><font face="Verdana" size="1">Employee Code</Font></span></td>
						<td><input name="cField" type="checkbox" value="employeename" checked /><span><font face="Verdana" size="1">Employee Name</Font></span></td>
						<td><input name="cField" type="checkbox" value="attendance" checked /><span><font face="Verdana" size="1">Attendance</Font></span></td>
						<td><input name="cField" type="checkbox" value="jobcode" checked /><span><font face="Verdana" size="1">Job Code</Font></span></td>
						<td><input name="cField" type="checkbox" value="jobname" checked /><span><font face="Verdana" size="1">Job Description</Font></span></td>
					</tr>
					<tr>
						<td><input name="cField" type="checkbox" value="locationtype"checked /><span><font face="Verdana" size="1">Target Type</Font></span></td>
						<td><input name="cField" type="checkbox" value="locationcode"checked /><span><font face="Verdana" size="1">Target Code</Font></span></td>
						<td><input name="cField" type="checkbox" value="reference" /><span><font face="Verdana" size="1">Reference</Font></span></td>
						<td><input name="cField" type="checkbox" value="remarks"/><span><font face="Verdana" size="1">Remarks</Font></span></td>
						<td><input name="cField" type="checkbox" value="bucket"/><span><font face="Verdana" size="1">Bucket</Font></span></td>
						<td><input name="cField" type="checkbox" value="spbno"/><span><font face="Verdana" size="1">SPB No</Font></span></td>
					</tr>
					<tr>
						<td><input name="cField" type="checkbox" value="chargejob"/><span><font face="Verdana" size="1">Charge Job</Font></span></td>
						<td><input name="cField" type="checkbox" value="chargecode"/><span><font face="Verdana" size="1">Charge Code</Font></span></td>
						<td><input name="cField" type="checkbox" value="chargetype"/><span><font face="Verdana" size="1">Charge Type</Font></span></td>
						<td><input name="cField" type="checkbox" value="remarksallowded"/><span><font face="Verdana" size="1">Remarks Allowded</Font></span></td>
						<td><input name="cField" type="checkbox" value="linenokey"/><span><font face="Verdana" size="1">Line No Key</Font></span></td>
					</tr>
				</table>
			</div>
		</div>
		
		</td>
	</tr>
		<tr>
		    <td colspan="3">
		        <br><font face="Verdana" size="1">Select Data Here :</font></td>
		</tr>
		<tr>
		    <td>
		        <font face="Verdana" size="1">From</font>
		    </td>
		    <td>
		        <input class="datepicker" type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">

		        <font face="Verdana" size="1">To</font>
		        <input class="datepicker" type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">
		    </td>
		</tr>
    	<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	 	<tr>
		<td align="left">
			<font face="Verdana"   size="1">Attendance</font>
		</td>
			<%
				Dim rsAttCodeA, qAttCodeA
				qAttCodeA = "select FCCODE, FCNAME from  PARAMETER  where PARCODE= 'E0102'  and fcba='" & Session("fcBA") & "'  order by FCNAME ASC"
				
				Set rsAttCodeA = Server.CreateObject("ADODB.Recordset")
				rsAttCodeA.CursorType = 3
				rsAttCodeA.CursorLocation = 3
				rsAttCodeA.Open qAttCodeA, connstr
				if not rsAttCodeA.eof then
			%>
				<td align="left">
					<select name="ATTENDANCE">
						<option value=""></option>
					<%Do While Not rsAttCodeA.eof%>
						<%if rsAttCodeA("FCCODE")=request("ATTENDANCE") Then%>			
						<option value="<%=rsAttCodeA("FCCODE")%>" selected><%=rsAttCodeA("FCCODE")%> : <%=rsAttCodeA("FCNAME")%></option> 
						<%Else%>
						<option value="<%=rsAttCodeA("FCCODE")%>"><%=rsAttCodeA("FCCODE")%> : <%=rsAttCodeA("FCNAME")%></option> 
						<%End if%>
			<%			rsAttCodeA.movenext
					loop
				End if
				rsAttCodeA.Close
				set rsAttCodeA = Nothing
				qAttCodeA = ""
			%>
					</select>
					<font color="#FF0000">*</font>
				</td>
			</tr>
    <tr>
	    <td>

	        <font face="Verdana" size="1">Gang</font>
	    </td>
	    <td>
	        <select name="GANGCODE">
	            <%cmdCallCombo "Gang","","",""%>
	        </select>
	    </td>
  	</tr>
	<tr>
		<td><font face="Verdana" size="1">Section Name</font></td>
		<td>
			<select name="txtsectionname">
				<option value=""></option>
				<%
				Dim rssectionname, strsectionname
				strsectionname	=	"select distinct sectionname as fccode, sectionname_api.get_fcname(sectionname) fcname from employee where fcba= '"& session("fcba") &"' order by sectionname"
				
				Set rssectionname = Server.CreateObject("ADODB.Recordset")
				rssectionname.Open strsectionname, connstr, adOpenStatic 
				
				if not rssectionname.eof then
				Do While Not rssectionname.eof
				if rssectionname("fccode") = v_txtsectionname then%>
						<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
					 <%Else%>
						<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
					<%End if
					rssectionname.movenext
				loop
				end if
				rssectionname.Close
				set rssectionname = Nothing
				%>
			</select>
		</td>
	</tr>
  <tr> 
    <td >
    <font face="Verdana" size="1">Job</font>
     </td>
    <td>
    <select name="JOBCODE" class="chosen-select" >
        <%cmdCallCombo "Job","","activation","Y"%>
      </select>
         </td>
      </tr>
      <tr> 
        <td>
	  <font face="Verdana" size="1">Employee</font>
      </td>
      <td>
      <select name="EMPLOYEECODE" class="chosen-select">
        <%cmdCallCombo "EMPLOYEE","","",""%>
      </select>
      </td>
      </tr>
      <tr> 
        <td>
	  <font face="Verdana" size="1">Target Type</font>
      </td>
      <td>
      <select name="LOCATIONTYPE">
        <%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
      </select>
      </td>
  </tr>
  <tr> 
    <td><font face="Verdana" size="1">Target Code</font>
     </td>
    <td><input type="text" size="15" maxlength="30" name="LOCATIONCODE"  style="background:#FFFFCC;">
     </td>
     </tr>
     <tr>
    <td>
    <font face="Verdana" size="1">Report</font>
     </td>
    <td>
    <select name="Report" >
        <option value="ACTIVITY">ACTIVITY</option>
        <option value="ATTENDANCE">ATTENDANCE</option>
      </select>
     
      </td>
  </tr>
  <tr>
		<td>
			<font face="Verdana" size="1">Status</font>
		</td>
		<td>
			<select name="rowstate" id="rowstate"> 
				<option></option>
					<% 	q = " SELECT DISTINCT rowstate " & _
							" FROM  attendance_gad_temp order by rowstate desc"
						' response.write q
						Set rsComboBST = Server.CreateObject("ADODB.Recordset")
						rsComboBST.CursorType = 3
						rsComboBST.CursorLocation = 3
						rsComboBST.Open q, connstr
						if not rsComboBST.eof then%>
						<%
						Do While Not rsComboBST.eof
						%>
							<option value="<%=rsComboBST("rowstate")%>" 
								<%if cstr(""&rowstate)=cstr(""&rsComboBST("rowstate")) then %> selected 
								<%End If%>><%=rsComboBST("rowstate")%>
							</option> 
						<%rsComboBST.movenext
						loop
						End if
						rsComboBST.Close
						set rsComboBST = Nothing
						q=""
					%>
			</select>
		</td> 
	</tr>
  
</table>
<br />
   <input type="submit" value="Preview" >
   
   
<%elseif Request("txttable") ="EmpAllded" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Payslip.jpg" border="0" >
	<span>Maintenance Allowance & Deduction</span></a>
    <br><br>
	<script>
		function checkAll() {
			var checkbox_options = document.getElementById("checkbox_options");
			var checkbox_all = document.getElementById("checkbox_all");

			var checkboxes = checkbox_options.getElementsByTagName("input");        
			for (i = 0; i < checkboxes.length; i++) {
				checkboxes[i].checked = checkbox_all.checked;
			}
		}
	</script>
	<Table>
		<tr>
			<Td colspan="3">
				<font face="Verdana" size="1">Select Field :</Font><br>
				<div class="ckbox_list">
					<table>
						<tr>
							<td><input id="checkbox_all" type="checkbox" onclick="checkAll()" /><span><font face="Verdana" size="1">Select All</Font></span></td>
						</tr>
					</table>
					<div id = "checkbox_options" >
						<table>
							<tr>
								<td><input name="cField" type="checkbox" value="adhocdate" /><span><font face="Verdana" size="1">Date</font></span></td>
								<td><input name="cField" type="checkbox" value="employeecode" /><span><font face="Verdana" size="1">Employee Code</Font></span></td>
								<td><input name="cField" type="checkbox" value="employeename"/><span><font face="Verdana" size="1">Employee Name</Font> </span></td>
								<td><input name="cField" type="checkbox" value="gangcode" /><span><font face="Verdana" size="1">Gang Code</Font></span></td>
							</tr>
						</table>
					</div>
				</div>
			</td>
		</tr>
	<tr>
	<td><br><font face="Verdana" size="1">Select Data Here :</font><br> 
    </Td>
    </tr>
    <tr>
	<td>
    <font face="Verdana" size="1">From</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
	<td>
	<font face="Verdana" size="1">To</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</Td>
    </tr>
    <tr>
	<td>
    <font face="Verdana" size="1">Gang</font>
    </td>
    <td>
	<select name="GANGCODE" class="chosen-select">
		<%cmdCallCombo "Gang","","",""%>
	</select>
    </td>
	</tr>
	<tr>
	<td>
    <font face="Verdana" size="1">Allowance</font>
    </td>
    <td>
	<select name="Allowded" class="chosen-select">
		<%cmdCallCombo "Allowded","","",""%>
	</select>
    </Td>
    </tr>
    <tr>
	<td>
    <font face="Verdana" size="1">Employee</font>
    </td>
    <td>
	<select name="EMPLOYEECODE" class="chosen-select">
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select>
	
	</td>
	</tr>
	</Table>
    <br />
    <input type="submit" value="Preview" >
    
    
<%Elseif Request("txttable") ="Calculate" then%>
<div class="close_box">
    <b>Calculate Payroll</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Employee")) &"/"& Year(Session("v_Employee"))%>
	
</div>  
<%Elseif Request("txttable") ="closing_daily_payroll" then%>
<%
							'	Dim INDAY : INDAY = Request("INDAY")
								Dim FCBA : FCBA = Request("FCBA")
							'	DIM MINDATE,MAXDATE
								
							'	If Len(Trim(INDAY)) = 0 Then
							'		INDAY = textDate(Date)
							'	End if

							'	q = " SELECT MAX(FDDATE ) AS MaxDate, MIN(FDDATE ) AS MinDate FROM DAILYCLOSINGACTIVITY_PAYROLL where fcba='GWE'"
								' response.write q
							'	Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							'	rsComboBST.CursorType = 3
							'	rsComboBST.CursorLocation = 3
							'	rsComboBST.Open q, connstr
							'	IF  not rsComboBST.EOF THEN 
									'MINDATE = Date
									'MAXDATE = Date
							'		response.write(MINDATE)
							'	ELSE 
									'MINDATE = rsComboBST("MinDate")
									'MAXDATE = rsComboBST("MaxDate")
									
							'	END IF 
							
								%>

<script language="JavaScript">
$(document).ready(function(){	
	$("#inday").datepicker();
	//$("#inday").datepicker({minDate: new Date("<%=MINDATE%>"), maxDate: new Date("<%=MAXDATE%>"),mandatory: true});
	//alert(<%=MINDATE%>);
});
</script>
<div class="close_box">
 
					<table >
						<thead>
						</thead>
						<tbody >
							
							
							<tr>
								<td align="left" width="100">FCBA</td>
								<td align="left"> :&nbsp;&nbsp;&nbsp;&nbsp;<select name="businessunit" id="fcba">
										<option value=""></option> 
										<% 	 
											q = " SELECT DISTINCT fccode " & _
												" FROM (SELECT businessunit.fccode " & _
												" FROM user_fcba_tab " & _
												" INNER JOIN businessunit " & _
												" ON user_fcba_tab.fcba = " & _
												" businessunit.fccode " & _
												" WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
											' response.write q
											Set rsComboBST = Server.CreateObject("ADODB.Recordset")
											rsComboBST.CursorType = 3
											rsComboBST.CursorLocation = 3
											rsComboBST.Open q, connstr
											if not rsComboBST.eof then%>
											<%
											Do While Not rsComboBST.eof
											%>
												<option value="<%=rsComboBST("fccode")%>" 
													<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
													<%End If%>><%=rsComboBST("fccode")%>
												</option> 
											<%rsComboBST.movenext
											loop
											End if
											rsComboBST.Close
											set rsComboBST = Nothing
											q=""
										%>
									</select>
								</td>
							</tr>
							<tr>
								<td align="left">DATE</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;<input id="inday"  name="INDAY" size="10" value="" style="width:100px;">
								</td>
							</tr>
							
							<tr>
								<td colspan="2" align="left" style="padding-top:30px;"><input type="submit" value="Process" id="button_exe"></td>
								
							</tr>
						</tbody>
					</table>
				
	
</div>    
<%Elseif Request("txttable") ="CalculatePPH21" then%>
<div class="close_box">
    <b>Calculate Payroll-PPH21</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Employee")) &"/"& Year(Session("v_Employee"))%>
</div>   

<%Elseif Request("txttable") ="Employee_Closing_New" then%>
<div class="close_box">
    <b>Employee Close</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process"><br>
	Periode : <%=Month(Session("v_Employee")) &"/"& Year(Session("v_Employee"))%>
	
</div>

    
<%elseif Request("txttable") ="HarvestingQuality" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/HarvestingQuality.jpg" border="0" >
	<span>Maintenance Harvesting Quality</span></a>
    <br><br><br>
    
    <table>
    <tr>
    <td>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    </td>
    <td>
    <div class="ckbox_list">
    
	<input name="cField" type="checkbox" value="FDDATE" /><span><font face="Verdana" size="1">Tanggal</font></span>
	<input name="cField" type="checkbox" value="EMPCODE" /><span><font face="Verdana" size="1">NIK</Font></span>
	<input name="cField" type="checkbox" value="FIELDCODE"/><span><font face="Verdana" size="1">Blok</Font></span>
    
	</div>
    
    </td>
    </tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Form</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Employee</font>
    </td>
    <td>
	<select name="EMPLOYEECODE" class="chosen-select">
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select>
    </td>
    </tr>
    </table>
	<input type="submit" value="Preview">
    
<%elseif Request("txttable") ="WorkshopActivity" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/workshop.jpg" border="0">
	<span>Maintenance Workshop Activity</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
	    <tr>
		    <td>
		    	<font face="Verdana" size="1">From</font>
		    </td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >			
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
	    </tr>
	    <tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
					<option value=""></option> 
					<% 	 
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboBST = Server.CreateObject("ADODB.Recordset")
						rsComboBST.CursorType = 3
						rsComboBST.CursorLocation = 3
						rsComboBST.Open q, connstr
						if not rsComboBST.eof then%>
						<%
						Do While Not rsComboBST.eof
						%>
							<option value="<%=rsComboBST("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
								<%End If%>><%=rsComboBST("fccode")%>
							</option> 
						<%rsComboBST.movenext
						loop
						End if
						rsComboBST.Close
						set rsComboBST = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	    <tr>
	    	<td>
	    		<font face="Verdana" size="1">Activity</font>
	    	</td>
	    	<td>
				<select name="JOB" class="chosen-select">
					<%cmdCallCombo "job","","activation","Y"%>
				</select>
	    	</td>
	    </tr>
	    <tr>
	    	<td>
	    		<font face="Verdana" size="1">Target Type</font>
	    	</td>
		    <td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
				</select>
		    </td>
	    </tr>
	    <tr>
	    	<td>
	    		<font face="Verdana" size="1">Target Code</font>
	    	</td>
	    	<td>
	    		<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
	    	</td>
	    </tr>
    </table><br />

	<input type="submit" value="Preview" >
<%Elseif Request("txttable") ="Workshop_Closing" then%>
<div class="close_box">
    <b>Workshop Close</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Workshop")) &"/"& Year(Session("v_Workshop"))%>
</div>
<%elseif Request("txttable") ="VehicleActivity" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Vehicle.png" border="0" >
	<span>Maintenance Vehicle Activity</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
    <table>
	    <tr>
		    <td>
		    	<font face="Verdana" size="1">From</font>
		    </td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="VEHDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="VEHDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
	    </tr>
	    <tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
					<option value=""></option> 
					<% 	 
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboBST = Server.CreateObject("ADODB.Recordset")
						rsComboBST.CursorType = 3
						rsComboBST.CursorLocation = 3
						rsComboBST.Open q, connstr
						if not rsComboBST.eof then%>
						<%
						Do While Not rsComboBST.eof
						%>
							<option value="<%=rsComboBST("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
								<%End If%>><%=rsComboBST("fccode")%>
							</option> 
						<%rsComboBST.movenext
						loop
						End if
						rsComboBST.Close
						set rsComboBST = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	    <tr>
	    	<td>
	    		<font face="Verdana" size="1">Group</font>
	    	</td>
			<td>
				<select name="VEHICLEGROUPCODE">
					<%cmdCallComboSetup "VEHICLEGROUP","VEHICLEGROUPCODE","fcba",session("fcba")%>
				</select>
		    </td>
	    </tr>
	    <tr>
		    <td>
		    	<font face="Verdana" size="1">Vehicle</font>
		    </td>
			<td>
				<select name="VEHICLECODE" class="chosen-select">
					<%cmdCallCombo "Vehicle","","",""%>
				</select>
		    </td>
	    </tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Operator/Driver</font>
			</td>
			<td align="left">
				<input type="text" size="15" maxlength="30" name="employeecode" value="<%=employeecode%>"> 
			</td>
		</tr>
	    <tr>
	    	<td>
		   		<font face="Verdana" size="1">Activity</font>
		    </td>
			<td>
				<select name="JOBCODE" class="chosen-select">
					<%cmdCallCombo "job","","activation","Y"%>
				</select>
		    </td>
	    </tr>
	    <tr>
		    <td>
		    	<font face="Verdana" size="1">Target Type</font>
		    </td>
			<td>
				<select name="LOCATIONTYPE" >
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
				</select>
		    </td>
	    </tr>
	    <tr>
		    <td>
				<font face="Verdana" size="1">Target Code</font>
		    </td>
			<td>
		    	<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
		    </td>
		</tr>
		<tr>
			<td>
		    	<font face="Verdana" size="1">Category Transport</font>
		    </td>
			<td>
				<select name="categorytransport">
					<option value=""> --All-- </option>
				<% 
					q = " SELECT fccode, fcname FROM vehicleactivitycategory ORDER BY LPAD(fccode, 2, 0)"
					set rsCombo = Server.CreateObject("ADODB.Recordset")
					'response.write q
					rsCombo.CursorType = 3
					rsCombo.CursorLocation = 3
					rsCombo.Open q, connstr
					if not rsCombo.eof then
						Do While Not rsCombo.eof
				%>
					<option value="<%=rsCombo("fccode")%>" <% If Request("categorytransport") = rsCombo("fccode") Then %> Selected <%End if%>><%=rsCombo("fccode")%> : <%=rsCombo("fcname")%></option> 
				<%    		rsCombo.movenext
						loop						
					End if
					rsCombo.Close
					set rsCombo = Nothing
					q=""
				%>
				</select>
		    </td>
	    </tr>
	    <tr>
	    	<td colspan="2" align="left"> 
	    	<input type="submit" value="Preview" >
	    	</td>
		</tr>
    </table>
	
<%elseif Request("txttable") ="VehicleOperation" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Vehicle.png" border="0" >
	<span>Maintenance Vehicle Operation</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
     <table>
    <tr>
    <td><font face="Verdana" size="1">From</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="VEHDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="VEHDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Group</font>
    </td>
	<td>
	<select name="VEHICLEGROUPCODE">
		<%cmdCallComboSetup "VEHICLEGROUP","VEHICLEGROUPCODE","fcba",session("fcba")%>
	</select>
    </td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Vehicle</font>
    </td>
	<td>
	<select name="VEHICLECODE" class="chosen-select">
		<%cmdCallCombo "Vehicle","","",""%>
	</select>
    </td>
    </tr>
    <tr>
	<tr>
	<td>
    <font face="Verdana" size="1">Vehicle Status</font>
    </td>
	<td>
		<select name="vehiclestatus">
			<option value="" <% If Request("vehiclestatus") = "" Then %> Selected <%End if%>> --All-- </option>
			<%
			q = " SELECT fccode, fcname FROM parameter where fcba = '" &  Session("fcBA") & "' AND parcode = 'VO001'"
			Set rsComboA = Server.CreateObject("ADODB.Recordset")
			
			rsComboA.Open q, connstr
			if not rsComboA.eof then
				Do While Not rsComboA.eof
			%>
				<option value="<%=rsComboA("fccode")%>" <% If cstr("" & Request("vehiclestatus")) = cstr(rsComboA("fccode")) Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
			<% 
					rsComboA.movenext
				loop
			End if
			rsComboA.Close : set rsComboA = Nothing : q = ""
			%>
			
		</select>
    </td>
    </tr>
    <tr>
    <td colspan="2" align="left"> 
    <input type="submit" value="Preview" >
    </td>
	</tr>
    </table>
    
<%Elseif Request("txttable") ="Vehicle_Closing" then%>
    <b>Vehicle Close</b></a><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Vehicle")) &"/"& Year(Session("v_Vehicle"))%>
	
    
    
<%elseif Request("txttable") ="AGREEMENTDETAIL" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Contract.jpg" border="0" >
	<span>Maintenance Contract Agreement</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					Dim objectRs, objectSQL
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Due Date From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATE_FROM" size="10" value="<%=DATE_FROM%>"title="Type Date mm/dd/yyyy" >
		
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATE_TO" size="10" value="<%=DATE_TO%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>


<%elseif Request("txttable") ="WIPPLAN" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Contract.jpg" border="0" >
	<span>Maintenance Contract Work in Progress Plan</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATE_FROM" size="10" value="<%=DATE_FROM%>"title="Type Date mm/dd/yyyy" >
		
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATE_TO" size="10" value="<%=DATE_TO%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>

<%elseif Request("txttable") ="wipfinal" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Contract.jpg" border="0" >
	<span>Maintenance Contract Work in Progress Final</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
		<tr>
			<td><font face="Verdana" size="1">Year</font></td>
			<td>
				<select name="txtyear">
				<%
					Yr = Year(Session("v_Contract"))
					mon = Month(Session("v_Contract"))
					MinYr = Yr - 2
					MaxYr = Yr + 2
					
					if len(v_txtyear) = 0 then v_txtyear = Yr
					if len(v_txtmonth) = 0 then v_txtmonth = mon
				%>
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtyear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Month</font></td>
			<td>
				<select name="txtmonth">
				<%For i = 1 to 12%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtmonth) then%> Selected <%End If%>><%=textMonth(i)%></option>
				<%Next%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
	
<%elseif Request("txttable") ="WIP" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Contract.jpg" border="0" >
	<span>Maintenance Contract Work in Progress</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATE_FROM" size="10" value="<%=DATE_FROM%>"title="Type Date mm/dd/yyyy" >
		
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATE_TO" size="10" value="<%=DATE_TO%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
    
      
<%Elseif Request("txttable") ="Contract_Closing" then%>
<div class="close_box">
    <b>Contract Close</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br></a>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Contract")) &"/"& Year(Session("v_Contract"))%>
	
</div>

<%elseif Request("txttable") ="PRODUCTQUALITY" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Product Quality </span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		 <tr>
			<td>
			<font face="Verdana" size="1">Storage Code</font>
			</td><td>
			<select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","fcba",session("fcba")%></select>	
			</td>
		 </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>    
<%elseif Request("txttable") ="lhp" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance LHP</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font>
    </td>
    <td>
		<%
				'CCC GET DATE PRODUCTION'
			dim sqltgl, rsdata,tglprod
			sqltgl = "SELECT PRODUCTIONMILLDATE FROM configmst " & _
			" WHERE fcba='" & session("fcba") & "'"
			Set rsdata = Server.CreateObject("ADODB.Recordset")
			
			rsdata.CursorType = 3
			rsdata.CursorLocation = 3
			rsdata.Open sqltgl, connstr
			if not rsdata.eof then
				tglprod = rsdata("PRODUCTIONMILLDATE")
			end if
			rsdata.close 
			set rsdata = nothing
			'' response.write tglprod
		%>
		<input class="datepicker" type="text" maxlength="10" name="lhpdate_From" size="10" value="" title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>

<%elseif Request("txttable") ="bkm_operasional_input" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
    <font face="Verdana" size="1">Select Data Here : <br><br>
	

    <table>
    <tr>
    <td>
		<font face="Verdana" size="1">From</font>
    </td>
    <td>
		
		<input class="datepicker" type="text" maxlength="10" name="bkmfrom" id="bkmfrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
	</td>
	<td>
		<font face="Verdana" size="1">To</font>
    </td>
    <td>
		
		<input class="datepicker" type="text" maxlength="10" name="bkmto" id="bkmto" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
	</td>
	
    </tr>
	<tr>
		<td>
			<font face="Verdana" size="1">BKM Type</font>
		</td>
		<td colspan=2>
			<select name="pilihbkm" id="pilihbkm" >
				<option value="">--PILIH--</option>
				<option value="501">BKM-OPERATION PERAWATAN</option>
				<option value="501">BKM-OPERATION MILL</option>
				<option value="503">BKM-OPERATION PANEN</option>
				<option value="502">BKM-TRAKSI (VH/WS)</option>
				<option value="504">BKM-WIP</option>
			</select>
		</td>
	</tr>
	

	<tr>    
    <!--<td colspan="2" align="left">
		<input type="button" name="create" id="create" class="button_menu_top" value="Create" onclick="getViewBKM()"></input>
	
		<input type="button" id="tampilBkm" onclick="tampilData();" value="View" >
		<input type="button" id="tampilBkm" onclick="tampilListingBKM();" value="View Listing BKM" >
		
    </td>
	<td>
		<img  src="IMAGES/xls-icon.png" width=55 height=30 alt="CETAK XLS" onclick="cetakXLS()"> 
	</td>-->
	<td colspan="5" align="left">
	    <a class="button_menu_top" onclick="getViewBKM()" ><img src="Images/add2.gif" border="0" ><span> Create Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="tampilData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="tampilListingBKM()" ><img src="Images/View.png" border="0" ><span> View Listing BKM</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakXLS()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
	</td>
	
	
	</tr>	
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	
	
	
	
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	function tampilListingBKM()
	{
		var pilihbkm = $("#pilihbkm").val();
		var bkmfrom = $("#bkmfrom").val();
		var bkmto = $("#bkmto").val();
		if(pilihbkm=="")
		{
			alert("JENIS BKM BELUM DIPILIH");
		}else{
			if(pilihbkm=="501"){
				window.open('p_listing_bkm.asp?bkmfrom=' + bkmfrom + '&bkmto=' + bkmto + '&doc=' + 501, '_blank');
			}else if(pilihbkm=="502")
			{
				window.open('p_listing_bkm.asp?bkmfrom=' + bkmfrom + '&bkmto=' + bkmto + '&doc=' + 502, '_blank');
			}else if(pilihbkm=="503")
			{
				window.open('p_listing_bkm.asp?bkmfrom=' + bkmfrom + '&bkmto=' + bkmto + '&doc=' + 503, '_blank');
			}else if(pilihbkm=="504")
			{
				window.open('p_listing_bkm.asp?bkmfrom=' + bkmfrom + '&bkmto=' + bkmto + '&doc=' + 504, '_blank');
			}
		}
	}

	function getViewBKM()
	{
		var pilihbkm = $("#pilihbkm").val();
		if(pilihbkm=="")
		{
			alert("JENIS BKM BELUM DIPILIH");
		}else{
			if(pilihbkm=="501"){
				window.open('bkm_operasional_input.asp', '_blank');
			}else if(pilihbkm=="502")
			{
				window.open('bkm_traksi_input.asp', '_blank');
			}else if(pilihbkm=="503")
			{
				window.open('bkm_panen_input.asp', '_blank');
			}else if(pilihbkm=="504")
			{
				window.open('bkm_wip_input.asp', '_blank');
			}
		}
		
	}

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	function viewEdit(i)
	{
		alert(i);
	}

	function tampilData()
	{
		var bkmfrom = $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var jumlinv = parseInt($('#jumldata').val());
		var pilihbkm = $("#pilihbkm").val();
		var totupah		= 0;
		var totbahan	= 0;
		var totvehicle	= 0;
		var totwip		= 0;
		var tot  		= 0;
		
		const date1 = new Date(bkmfrom);
		const date2 = new Date(bkmto);
		const diffTime = Math.abs(date2 - date1);
		const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24))+1; 

		
		if (diffDays <= 31) {
		
		if (bkmfrom != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?bkmfrom=" + bkmfrom + "&bkmto=" + bkmto + "&pilihbkm=" + pilihbkm + "&idview=" + 1,function(data) {
				$('#bodytable').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					//var x = json.results[i].fccode
					var doc = json.results[i].document_classification;
					//var vehc= json.results[i].vehiclecode;
					//if(vehc=="" || vehc==null){
					//	vehc ="-";
					//}else{
					//	vehc = json.results[i].vehiclecode;
					//}

					//var posting = json.results[i].postingflg;
					//if(posting=="1"){
					//	posting = "YA";
					//}else{
					//	posting = "TIDAK";
					//}

					var docketlist = '';
						docketlist += '<tr id="tru'+i+'">';
						docketlist += '<td align="left">' + json.results[i].bkmno + '</td>';
						docketlist += '<td align="left">' + json.results[i].createdate + '</td>';
						docketlist += '<td align="left">' + json.results[i].dcname + '</td>';
						docketlist += '<td align="left">' + json.results[i].hujan + '</td>';
						
						//if (json.results[i].hujan == "0"){
						//docketlist += '<td align="left">' + '<input disabled name="cbhujan'+i+'" id="cbhujan+'+i+'" type="checkbox" value="'+json.results[i].hujan+'" /></td>';}
						//else{
						//docketlist += '<td align="left">' + '<input disabled name="cbhujan'+i+'" id="cbhujan+'+i+'" type="checkbox" value="'+json.results[i].hujan+'" checked/></td>';
						//}
						docketlist += '<td align="left">' + json.results[i].sectioncode + '</td>';
						docketlist += '<td align="left">' + json.results[i].gangcode + '</td>';
						docketlist += '<td align="center">' + json.results[i].vehiclecode + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].totalbkmupah) + '<input type="hidden" name="upahtext'+i+'" id="upahtext'+i+'" value='+json.results[i].totalbkmupah+' /></td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].totalbkmmaterial) + '<input type="hidden" name="bahantext'+i+'" id="bahantext'+i+'" value='+json.results[i].totalbkmmaterial+' /></td>';				
						docketlist += '<td align="right">' + numformat2(json.results[i].totalbkmvehicle) + '<input type="hidden" name="vehicletext'+i+'" id="vehicletext'+i+'" value='+json.results[i].totalbkmvehicle+' /></td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].totalbkmwip) + '<input type="hidden" name="wiptext'+i+'" id="wiptext'+i+'" value='+json.results[i].totalbkmwip+' /></td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].total) + '<input type="hidden" name="tottext'+i+'" id="tottext'+i+'" value='+json.results[i].total+' /></td>';
						docketlist += '<td align="left" class="">' + json.results[i].postingflg + '</td>';
						if(json.results[i].postingflg=="TIDAK"){

						
						if(doc=="501"){
							docketlist += '<td align="center"><a target=_blank href="bkm_operasional_edit.Asp?bkmno='+ json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data PERWATAN="0"></a></td>';
						}else if(doc=="502"){
							docketlist += '<td align="center"><a target=_blank href="bkm_traksi_edit.Asp?bkmno='+ json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data Traksi border="0"></a></td>';
						}
						else if(doc=="503"){

							docketlist += '<td align="center"><a target=_blank href="bkm_panen_edit.Asp?bkmno='+ json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data Panen border="0"></a></td>';
						}
						else if(doc=="504"){

							docketlist += '<td align="center"><a target=_blank href="bkm_wip_edit.Asp?bkmno='+ json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data WIP border="0"></a></td>';
						}
						}
						docketlist += '<td align="center"><a target=_blank href="bkm_jurnal_input.Asp?tglbkm='+ json.results[i].createdate + '"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
						docketlist += '</tr>';
						$('#bodytable').append(docketlist);
					no++;
					$("#jumldata").val(jumlview);
					totupah		+= json.results[i].totalbkmupah;
					totbahan	+= json.results[i].totalbkmmaterial;
					totvehicle	+= json.results[i].totalbkmvehicle;
					totwip		+= json.results[i].totalbkmwip;
					tot  		+= json.results[i].total;
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				getPuttotal(totupah,totbahan,totvehicle,totwip,tot);
					
			});
		 
		} else {
			$('#bodytable').empty();
		}
		} else {
		alert('Max 31 Hari !!');
		}
	}


	function getPuttotal(totupah,totbahan,totvehicle,totwip,tot) {
		//var totupah		= 0;
		//var totbahan	= 0;
		//var totvehicle	= 0;
		//var totwip		= 0;
		//var tot  		= 0;
		
		//var jumldata	= parseInt($('#jumldata').val());
		//for(var i=0; i<jumldata; i++) {
		//	totupah		+= parseInt($('#upahtext'+i).val());
		//	totbahan	+= parseInt($('#bahantext'+i).val());
		//	totvehicle	+= parseInt($('#vehicletext'+i).val());
		//	totwip		+= parseFloat($('#wiptext'+i).val());
		//	tot  		+= parseFloat($('#tottext'+i).val());
		//}
		//	totupah 	= numformat(totupah);
		//	totbahan 	= numformat(totbahan);
		//	totvehicle 	= numformat(totvehicle);
		//	totwip 		= numformat(totwip);
		//	tot         = numformat(tot);
		$('#totalupahview').html(numformat2(totupah));
		$('#totalbahanview').html(numformat2(totbahan));
		$('#totalvehicleview').html(numformat2(totvehicle));
		$('#totalwipview').html(numformat2(totwip));
		$('#totalview').html(numformat2(tot));
		
		
	}

	function cetakXLS()
	{
		var bkmfrom 	= $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var fcba 	= $("#fcba").val();
		
		window.open('p_biayaxls.Asp?tglawal=' + bkmfrom + '&tglakhir=' + bkmto, '_blank');
		//window.location.href = "p_biayaxls.Asp?tahun=" + tahun2 + "&bulan=" + bulan2, '_blank';
	}


	</script>


	<br>
	<table width="917" class=data>
		<tr>
			<th>BKM No</th>
			<th>BKM Date</th>
			<th>BKM Type</th>
			<th>HUJAN</th>			
			<th>SECTION</th>
			<th>GANG</th>
			<th>VEHCODE</th>
			<th>BKM Upah</th>
			<th>BKM Material</th>
			<th>BKM Vehicle</th>
			<th>BKM WIP</th>
			<th>TOTAL</th>
			<th>POSTING</th>
			<th colspan="2">Action</th>
		</tr>

		<tr id='tru'>
			<tbody id="bodytable">
			</tbody>
		</tr>
		<tr id="bodytable">
		<td colspan="7" align="right" class=" ratakanan huruftebal">Total</td>
		<td align='right'><span id="totalupahview" class=" ratakanan huruftebal">0</span></td>
		<td align='right'><span id="totalbahanview" class="ratakanan huruftebal">0</span></td>
		<td align='right'><span id="totalvehicleview" class="ratakanan huruftebal">0</span></td>
		<td align='right'><span id="totalwipview" class="ratakanan huruftebal">0</span></td>
		<td align='right'><span id="totalview" class="ratakanan huruftebal">0</span></td>
		
	</tr>
	</table>


<!-- view listing bkm-->
<%elseif Request("txttable") ="listing_bkm" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>
	

    <table>
    

	<tr>
		    <td>
		        <font face="Verdana" size="1">From</font>
		    </td>
		    <td>
		        <input class="datepicker" type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">

		        <font face="Verdana" size="1">To</font>
		        <input class="datepicker" type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy">
		    </td>
		</tr>
    	<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	 	<tr>
		<td align="left">
			<font face="Verdana"   size="1">Attendance</font>
		</td>
			<%
				'Dim rsAttCodeA, qAttCodeA
				qAttCodeA = "select FCCODE, FCNAME from  PARAMETER  where PARCODE= 'E0102'  and fcba='" & Session("fcBA") & "'  order by FCNAME ASC"
				
				Set rsAttCodeA = Server.CreateObject("ADODB.Recordset")
				rsAttCodeA.CursorType = 3
				rsAttCodeA.CursorLocation = 3
				rsAttCodeA.Open qAttCodeA, connstr
				if not rsAttCodeA.eof then
			%>
				<td align="left">
					<select name="ATTENDANCE">
						<option value=""></option>
					<%Do While Not rsAttCodeA.eof%>
						<%if rsAttCodeA("FCCODE")=request("ATTENDANCE") Then%>			
						<option value="<%=rsAttCodeA("FCCODE")%>" selected><%=rsAttCodeA("FCCODE")%> : <%=rsAttCodeA("FCNAME")%></option> 
						<%Else%>
						<option value="<%=rsAttCodeA("FCCODE")%>"><%=rsAttCodeA("FCCODE")%> : <%=rsAttCodeA("FCNAME")%></option> 
						<%End if%>
			<%			rsAttCodeA.movenext
					loop
				End if
				rsAttCodeA.Close
				set rsAttCodeA = Nothing
				qAttCodeA = ""
			%>
					</select>
					<font color="#FF0000">*</font>
				</td>
			</tr>
    <tr>
	    <td>

	        <font face="Verdana" size="1">Gang</font>
	    </td>
	    <td>
	        <select name="GANGCODE">
	            <%cmdCallCombo "Gang","","",""%>
	        </select>
	    </td>
  	</tr>
	<tr>
		<td><font face="Verdana" size="1">Section Name</font></td>
		<td>
			<select name="txtsectionname">
				<option value=""></option>
				<%
			''	Dim rssectionname, strsectionname
				strsectionname	=	"select distinct sectionname as fccode, sectionname_api.get_fcname(sectionname) fcname from employee where fcba= '"& session("fcba") &"' order by sectionname"
				
				Set rssectionname = Server.CreateObject("ADODB.Recordset")
				rssectionname.Open strsectionname, connstr, adOpenStatic 
				
				if not rssectionname.eof then
				Do While Not rssectionname.eof
				if rssectionname("fccode") = v_txtsectionname then%>
						<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
					 <%Else%>
						<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
					<%End if
					rssectionname.movenext
				loop
				end if
				rssectionname.Close
				set rssectionname = Nothing
				%>
			</select>
		</td>
	</tr>
  <tr> 
    <td >
    <font face="Verdana" size="1">Job</font>
     </td>
    <td>
    	<select name="JOBCODE" class="chosen-select" >
        	<%cmdCallCombo "Job","","activation","Y"%>
      </select>
    </td>
      </tr>
      <tr> 
        <td>
	  <font face="Verdana" size="1">Employee</font>
      </td>
      <td>
      <select name="EMPLOYEECODE" class="chosen-select">
        <%cmdCallCombo "EMPLOYEE","","",""%>
      </select>
      </td>
      </tr>
      
      
  </tr>
  
  <tr>
		<td>
			<font face="Verdana" size="1">Status</font>
		</td>
		<td>
			<select name="rowstate" id="rowstate"> 
				<option></option>
					<% 	q = " SELECT DISTINCT rowstate " & _
							" FROM  attendance_gad_temp order by rowstate desc"
						' response.write q
						Set rsComboBST = Server.CreateObject("ADODB.Recordset")
						rsComboBST.CursorType = 3
						rsComboBST.CursorLocation = 3
						rsComboBST.Open q, connstr
						if not rsComboBST.eof then%>
						<%
						Do While Not rsComboBST.eof
						%>
							<option value="<%=rsComboBST("rowstate")%>" 
								<%if cstr(""&rowstate)=cstr(""&rsComboBST("rowstate")) then %> selected 
								<%End If%>><%=rsComboBST("rowstate")%>
							</option> 
						<%rsComboBST.movenext
						loop
						End if
						rsComboBST.Close
						set rsComboBST = Nothing
						q=""
					%>
			</select>
		</td> 
	</tr>
  
	

	<tr>    
    <td colspan="2" align="left">
		<input type="submit" name="preview" id="preview" class="button_menu_top" value="Preview"></input>
		
    </td>
	
	
	</tr>	
	
	</table>
	
	
	
	
	<script type="text/javascript">

	

	function cetakXLS()
	{
		var bkmfrom 	= $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var fcba 	= $("#fcba").val();
		
		window.open('p_biayaxls.Asp?tglawal=' + bkmfrom + '&tglakhir=' + bkmto, '_blank');
		//window.location.href = "p_biayaxls.Asp?tahun=" + tahun2 + "&bulan=" + bulan2, '_blank';
	}


	</script>

<%elseif Request("txttable") ="lap_tanaman" then%>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahuntnmn,rstahuntnmn
				sqltahuntnmn = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahuntnmn = Server.CreateObject("ADODB.Recordset")
				rstahuntnmn.CursorType = 3
				rstahuntnmn.CursorLocation = 3
				rstahuntnmn.Open sqltahuntnmn, connstr
				While Not rstahuntnmn.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahuntnmn("tahun")) then %>
					<option value="<%=rstahuntnmn("tahun")%>" Selected ><%=rstahuntnmn("tahun")%></option>
					<%else%>
					<option value="<%=rstahuntnmn("tahun")%>"><%=rstahuntnmn("tahun")%></option>
					<%end if%>					
				<%rstahuntnmn.movenext
				wend 
				rstahuntnmn.close 
				

			%>	
			</select>
		</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulantnmn,rsbulantnmn
				sqlbulantnmn = "select distinct to_number(accountingperiod) bulan,description from accountingperiod WHERE accountingperiod not in('13','14','15') order by bulan " 
				Set rsbulantnmn = Server.CreateObject("ADODB.Recordset")
				rsbulantnmn.CursorType = 3
				rsbulantnmn.CursorLocation = 3
				rsbulantnmn.Open sqlbulantnmn, connstr
				While Not rsbulantnmn.EOF 
                    if cint(MONTH(Session("v_Costbook")))=cint(rsbulantnmn("bulan")) then %>
					<option value="<%=rsbulantnmn("bulan")%>" Selected><%=rsbulantnmn("description")%></option>
					<%else%>
					<option value="<%=rsbulantnmn("bulan")%>"><%=rsbulantnmn("description")%></option>
					<%end if%>
				<%rsbulantnmn.movenext
				wend 
				rsbulantnmn.close 
				

			%>	
				
			</select>
		</td>
		
    </tr>

	<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value="all">--ALL--</option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	

	<tr>    
   
	<td colspan="4" align="left">
		<a class="button_menu_top" onclick="tampilData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
	</td>
	
	</tr>	
	</table>

	<script>
	function tampilData()
	{
		var tahun = $("#tahun").val();
		var bulan 	= $("#bulan").val();
		window.open('p_lhp_rekap.Asp?tahun=' + tahun + '&bulan=' + bulan, '_blank');		
	}
	</script>


<%elseif Request("txttable") ="lhp_view" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
	</style>
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
    <font face="Verdana" size="1">Select Data Here : <br><br>
	

    <table>
    <tr>
    <td>
		<font face="Verdana" size="1">From</font>
    </td>
    <td>
		
		<input class="datepicker" type="text" maxlength="10" name="bkmfrom" id="bkmfrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
	</td>
	<td>
		<font face="Verdana" size="1">To</font>
    </td>
    <td>
		
		<input class="datepicker" type="text" maxlength="10" name="bkmto" id="bkmto" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
	

	<tr>    
    <!--<td colspan="2" align="left">
		
		<input type="button" id="tampilBkm" onclick="tampilData();" value="View" >
		
    </td>-->
	<td colspan="4" align="left">
		<a class="button_menu_top" onclick="tampilData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakXLS()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
	</td>
	
	</tr>	
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	
	
	
	
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
}


	function tampilData()
	{
		var bkmfrom = $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var jumlinv = parseInt($('#jumldata').val());

	
	 	var dt = new Date(bkmfrom);
     	dt.setDate( dt.getDate() - 1 );
		var today = dt.getDate();
		var tomnth = dt.getMonth() + 1;
		var tglawal = tomnth + '/' + today + '/' + dt.getFullYear();

 
		if (bkmfrom != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?bkmfrom=" + tglawal + "&bkmto=" + bkmto + "&idview=" + 4,function(data) {
				$('#bodytablelhp').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var tbsterimahi 	= json.results[i].tbsterima;
					var tbssisahi 		= json.results[i].tbssisa;
					var potongantbshi 	= json.results[i].potongantbs;

					var cpokirim  		= json.results[i].cpokirim;
					var adjcpo  		= json.results[i].adjcpo;
					var cpocothi 		= json.results[i].cpocot;
					var cpocsthi 		= json.results[i].cpocst;
					var cpoothi 		= json.results[i].cpoot;
					var cpotankisihi 	= json.results[i].cpotankisi;
					var totalcpohi		= parseInt(cpokirim)+(parseInt(cpocothi)+parseInt(cpocsthi)+parseInt(cpoothi)+parseInt(cpotankisihi)+parseInt(adjcpo));

					var kernelkirimhi	= json.results[i].kernelkirim;
					var kernelsilohi	= json.results[i].kernelsilo;
					var pktankisihi		= json.results[i].pktankisi;
					var totalpkhi		= parseInt(kernelkirimhi)+(parseInt(kernelsilohi)+parseInt(pktankisihi))


					//GET Sisa stock CPO
					
					var cposisa 		= parseInt(cpotankisihi)+parseInt(cpocothi)+parseInt(cpocsthi)+parseInt(cpoothi);

					//GET Sisa stock PK
				 	var pksisa 			= parseInt(pktankisihi)+parseInt(kernelsilohi);
						
						
							
					
					
					if(i > 0){   
						var sisakemarin	= json.results[i-1].tbssisa;
						var tersedia   	= parseInt(tbsterimahi) + parseInt(sisakemarin);
                		var tbsolah    	= tersedia - tbssisahi-potongantbshi;

						var cpocotkemarin 		= json.results[i-1].cpocot;
						var cpocstkemarin 		= json.results[i-1].cpocst;
						var cpootkemarin 		= json.results[i-1].cpoot;
						var cpotankisikemarin 	= json.results[i-1].cpotankisi;
						var totalcpokemarin		= parseInt(cpocotkemarin)+parseInt(cpocstkemarin)+parseInt(cpootkemarin)+parseInt(cpotankisikemarin);
						var cpohi				= parseInt(totalcpohi)-parseInt(totalcpokemarin);

						//var kernelkirimkemarin	= json.results[i-1].kernelkirim;
						var kernelsilokemarin	= json.results[i-1].kernelsilo;
						var pktankisikemarin	= json.results[i-1].pktankisi;
						var totalpkkemarin		= parseInt(kernelsilokemarin)+parseInt(pktankisikemarin);
						var produksipkhi		= parseInt(totalpkhi)-parseInt(totalpkkemarin);



						

						
					

						var docketlist = '';
							docketlist += '<tr id="tru'+i+'">';
							docketlist += '<td align="left">' + json.results[i].lhpno + '</td>';
							docketlist += '<td align="left">' + json.results[i].proddate + '</td>';
							docketlist += '<td align="right">' + numformat2(json.results[i].tbsterima) + '<input type="hidden" name="tbsterimatext'+i+'" id="tbsterimatext'+i+'" value='+json.results[i].tbsterima+' /></td>';
							docketlist += '<td align="right">' + numformat2(tbsolah) + '<input type="hidden" name="tbsolahtext'+i+'" id="tbsolahtext'+i+'" value='+ tbsolah +' /></td>';
							docketlist += '<td align="right">' + numformat2(json.results[i].tbssisa) + '</td>';				
							docketlist += '<td align="right">' + numformat2(cpohi) + '<input type="hidden" name="cpoproduksitext'+i+'" id="cpoproduksitext'+i+'" value='+ cpohi +' /></td>';
							docketlist += '<td align="right">' + numformat2(cpokirim) + '<input type="hidden" name="cpokirimtext'+i+'" id="cpokirimtext'+i+'" value='+ cpokirim +' /></td>';
							docketlist += '<td align="right">' + numformat2(cposisa) + '</td>';
							docketlist += '<td align="right">' + numformat2(produksipkhi) + '<input type="hidden" name="pkproduksitext'+i+'" id="pkproduksitext'+i+'" value='+ produksipkhi +' /></td>';
							docketlist += '<td align="right">' + numformat2(kernelkirimhi) + '<input type="hidden" name="pkkirimtext'+i+'" id="pkkirimtext'+i+'" value='+ kernelkirimhi +' /></td>';
							docketlist += '<td align="right">' + numformat2(pksisa) + '</td>';
							docketlist += '<td align="center"><a target=_blank href="lhp_detail_edit.Asp?lhpno='+ json.results[i].lhpno + '"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
							//docketlist += '<td align="center"><a target=_blank href="bkm_jurnal_input.Asp?tglbkm='+ json.results[i].lhpno + '"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
							docketlist += '</tr>';
							$('#bodytablelhp').append(docketlist);
						no++;
						$("#jumldata").val(jumlview);
						getPuttotallhp();

					}
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				
			});
		 
		} else {
			$('#bodytablelhp').empty();
		}
	}


	function getPuttotallhp() {
		var tottbsterima	= 0;
		var tottbsolah		= 0;
		var totcpoproduksi	= 0;
		var totcpokirim		= 0;
		var totpkproduksi	= 0;
		var totpkkirim		= 0;
		var jumldata	= parseInt($('#jumldata').val());
		for(var x=1; x<jumldata; x++) {
			tottbsterima		+= parseInt($('#tbsterimatext' + x).val());
			tottbsolah			+= parseInt($('#tbsolahtext' + x).val());
			totcpoproduksi		+= parseInt($('#cpoproduksitext'+x).val());
			totcpokirim			+= parseInt($('#cpokirimtext'+x).val());
			totpkproduksi		+= parseInt($('#pkproduksitext'+x).val());
			totpkkirim			+= parseInt($('#pkkirimtext'+x).val());
			
		}
			var tottbsterimaview 	= numformat(tottbsterima);
			var tottbsolahview 		= numformat(tottbsolah);
			totcpoproduksi 		= numformat2(totcpoproduksi);
			totcpokirim 		= numformat2(totcpokirim);
			totpkproduksi 		= numformat2(totpkproduksi);
			totpkkirim 			= numformat2(totpkkirim);
		$('#ttltbsterima').html(tottbsterimaview);
		$('#ttltbsolah').html(tottbsolahview);
		$('#ttlcpoproduksi').html(totcpoproduksi);
		$('#ttlcpokirim').html(totcpokirim);
		$('#ttlpkproduksi').html(totpkproduksi);
		$('#ttlpkkirim').html(totpkkirim);
		
	}

	</script>


	<br>
	<table width="917" class=data>
		<tr>
			<th rowspan='2'>LHP No</th>
			<th rowspan='2'>Date Created</th>
			<th colspan='3'>TBS</th>
			<th colspan='3'>CPO</th>
			<th colspan='3'>Kernel</th>
			<th rowspan="2" colspan='2'>Action</th>
		</tr>
		<tr>
			<th>Terima</th>
			<th>Olah</th>
			<th>Sisa</th>
			<th>Produksi</th>
			<th>Kirim</th>
			<th>Sisa</th>
			<th>Produksi</th>
			<th>Kirim</th>
			<th>Sisa</th>
		</tr>

		<tr id='tru'>
			<tbody id="bodytablelhp">
			</tbody>
		</tr>

		<tr>
			<td colspan=2 class=ratakanan><strong>Total</strong></td>
			<td colspan= class=ratakanan><strong><span id=ttltbsterima>0</span></strong></td>
			<td colspan= class=ratakanan><strong><span id=ttltbsolah>0</span></strong></td>
			<td colspan= class=ratakanan></td>
			<td colspan= class=ratakanan><strong><span id=ttlcpoproduksi>0</span></strong></td>
			<td colspan= class=ratakanan><strong><span id=ttlcpokirim>0</span></strong></td>
			<td colspan= class=ratakanan></td>
			<td colspan= class=ratakanan><strong><span id=ttlpkproduksi>0</span></strong></td>
			<td colspan= class=ratakanan><strong><span id=ttlpkkirim>0</span></strong></td>
		</tr>
	</table>

<%elseif Request("txttable") ="lhp_rekap" then%>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahunlhp,rstahunlhp
				sqltahunlhp = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahunlhp = Server.CreateObject("ADODB.Recordset")
				rstahunlhp.CursorType = 3
				rstahunlhp.CursorLocation = 3
				rstahunlhp.Open sqltahunlhp, connstr
				While Not rstahunlhp.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahunlhp("tahun")) then %>
					<option value="<%=rstahunlhp("tahun")%>" Selected ><%=rstahunlhp("tahun")%></option>
					<%else%>
					<option value="<%=rstahunlhp("tahun")%>"><%=rstahunlhp("tahun")%></option>
					<%end if%>					
				<%rstahunlhp.movenext
				wend 
				rstahunlhp.close 
				

			%>	
			</select>
		</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulanlhp,rsbulanlhp
				sqlbulanlhp = "select distinct to_number(accountingperiod) bulan,description from accountingperiod WHERE accountingperiod not in('13','14','15') order by bulan " 
				Set rsbulanlhp = Server.CreateObject("ADODB.Recordset")
				rsbulanlhp.CursorType = 3
				rsbulanlhp.CursorLocation = 3
				rsbulanlhp.Open sqlbulanlhp, connstr
				While Not rsbulanlhp.EOF 
                    if cint(MONTH(Session("v_Costbook")))=cint(rsbulanlhp("bulan")) then %>
					<option value="<%=rsbulanlhp("bulan")%>" Selected><%=rsbulanlhp("description")%></option>
					<%else%>
					<option value="<%=rsbulanlhp("bulan")%>"><%=rsbulanlhp("description")%></option>
					<%end if%>
				<%rsbulanlhp.movenext
				wend 
				rsbulanlhp.close 
				

			%>	
				
			</select>
		</td>
		
    </tr>
	

	<tr>    
   
	<td colspan="4" align="left">
		<a class="button_menu_top" onclick="tampilData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
	</td>
	
	</tr>	
	</table>

	<script>
	function tampilData()
	{
		var tahun = $("#tahun").val();
		var bulan 	= $("#bulan").val();
		window.open('p_lhp_rekap.Asp?tahun=' + tahun + '&bulan=' + bulan, '_blank');		
	}
	</script>




<%elseif Request("txttable") ="rkb_operasional_input" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
    <font face="Verdana" size="1">Select Data Here : <br><br>
	

    <table>
    	<tr>
		<td>
			<font face="Verdana" size="1">Company</font>
		</td>

		<td>
		
			<select name="companyrkb" id="companyrkb" onchange="getFCBArkb()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
						</select>
		</td>
		
		   
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Business unit</font>
		</td>

		<td>
		
			<select name="businessunitrkb" id="businessunitrkb" style=width:100px >
			</select>
		</td>
		
		   
	</tr>
	
    <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahun7,rstahun7
				sqltahun7 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun7 = Server.CreateObject("ADODB.Recordset")
				rstahun7.CursorType = 3
				rstahun7.CursorLocation = 3
				rstahun7.Open sqltahun7, connstr
				While Not rstahun7.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahun7("tahun")) then %>
					<option value="<%=rstahun7("tahun")%>" Selected ><%=rstahun7("tahun")%></option>
					<%else%>
					<option value="<%=rstahun7("tahun")%>"><%=rstahun7("tahun")%></option>
					<%end if%>					
				<%rstahun7.movenext
				wend 
				rstahun7.close 
				

			%>	
			</select>
		</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulan7,rsbulan7
				sqlbulan7 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
				Set rsbulan7 = Server.CreateObject("ADODB.Recordset")
				rsbulan7.CursorType = 3
				rsbulan7.CursorLocation = 3
				rsbulan7.Open sqlbulan7, connstr
				While Not rsbulan7.EOF 
                    if cint(MONTH(Session("v_Costbook")))=cint(rsbulan7("bulan")) then %>
					<option value="<%=rsbulan7("bulan")%>" Selected><%=rsbulan7("description")%></option>
					<%else%>
					<option value="<%=rsbulan7("bulan")%>"><%=rsbulan7("description")%></option>
					<%end if%>
				<%rsbulan7.movenext
				wend 
				rsbulan7.close 
				

			%>	
				
			</select>
		</td>
		
    </tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Type</font>
		</td>
		<td colspan=>
			
			<select name="typerkb" id="typerkb" style="width:200px">
			    <option value="">--ALL--</option>
				<option value="501">Perawatan</option>
				<option value="502">Traksi</option>
				<option value="503">Panen</option>
				<option value="504">WIP</option>
			</select>
		</td>
		
		
	</tr>
	

	<tr>    
    <!--<td colspan="2" align="left">
		<a class="button_menu_top"  target="_blank" href="<%=txtLink%>.Asp "><span>Create</span></a>&nbsp;
		<input type="button" id="tampilrkb" onclick="tampilData();" value="View" >
    </td>
	<td>
		<img  src="IMAGES/xls-icon.png" width=55 height=30 alt="CETAK XLS" onclick="cetakXLS()"> 
	</td>-->
	<td colspan="4" align="left">
		<a class="button_menu_top" target="_blank" href="<%=txtLink%>".Asp ><img src="Images/add2.gif" border="0" ><span> Create Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="tampilData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakXLS()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
	</td>
	
	</tr>	
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	
	
	
	
	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	function viewEdit(i)
	{
		alert(i);
	}

	function tampilData()
	{
	    var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var pcompany = $("#companyrkb").val();
		var pfcba 	= $("#businessunitrkb").val();
		var ptype 	= $("#typerkb").val();
		
		var totupah		= 0;
		var totbahan	= 0;
		
		
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?tahun=" + tahun + "&bulan=" + bulan + "&idview=" + 8 + "&company=" + pcompany+"&fcba=" + pfcba+"&type=" + ptype,function(data) {
				$('#bodytable').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var x = json.results[i].fccode
					//alert(x);
					var docketlist = '';
						docketlist += '<tr id="tru'+i+'">';
						docketlist += '<td align="left">' + json.results[i].fcba + '</td>';
						docketlist += '<td align="left">' + json.results[i].type + '</td>';
						docketlist += '<td align="left">' + json.results[i].rkbno + '</td>';
						docketlist += '<td align="left">' + json.results[i].period + '</td>';
						docketlist += '<td align="left">' + json.results[i].sectioncode + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].totupah) + '<input type="hidden" name="upahtext'+i+'" id="upahtext'+i+'" value='+json.results[i].totupah+' /></td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].totbahan) + '<input type="hidden" name="bahantext'+i+'" id="bahantext'+i+'" value='+json.results[i].totbahan+' /></td>';				
						docketlist += '<td align="center"><a target=_blank href="rkb_operasional_edit.Asp?rkbno='+json.results[i].rkbno+ '&fcba='+json.results[i].fcba+'"><img src="IMAGES/edit.gif" alt="Edit Data border="0"></a></td>';
											
						docketlist += '</tr>';
						$('#bodytable').append(docketlist);
						totupah		+= json.results[i].totupah;
						totbahan	+= json.results[i].totbahan;
					no++;
					$("#jumldata").val(jumlview);
					
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				getPuttotal(totupah,totbahan);
			});
		 
	} 


	function getPuttotal(p_upah,p_bahan) {
		
		totupah 	= numformat(p_upah);
		totbahan 	= numformat(p_bahan);
		$('#totalupahrkb').html(totupah);
		$('#totalbahanrkb').html(totbahan);
		
	}

	function cetakXLS()
	{
		var rkbfrom 	= $("#rkbfrom").val();
		var rkbto 	= $("#rkbto").val();
		var fcba 	= $("#fcba").val();
		
		window.open('p_biayaxls.Asp?tglawal=' + rkbfrom + '&tglakhir=' + rkbto, '_blank');
		//window.location.href = "p_biayaxls.Asp?tahun=" + tahun2 + "&bulan=" + bulan2, '_blank';
	}


	</script>


	<br>
	<table width="917" class=data>
		<tr>
		    <th>FCBA</th>
			<th>TYPE</th>
			<th>RKB NO</th>
			<th>RKB PERIOD</th>
			<th>SECTION</th>
			<th>RKB Upah</th>
			<th>RKB Material</th>
			<th></th>
		</tr>

		<tr id='tru'>
			<tbody id="bodytable">
			</tbody>
		</tr>
		<tr id="bodytable">
		<td colspan="5" align="right" class=" ratakanan huruftebal">Total</td>
		<td align='right'><span id="totalupahrkb" class=" ratakanan huruftebal">0</span></td>
		<td align='right'><span id="totalbahanrkb" class="ratakanan huruftebal">0</span></td>
		
	</tr>
	</table>
<%elseif Request("txttable") ="listing_mam" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
    <font face="Verdana" size="1">Select Data Here : <br><br>
	

    <table>
    	<tr>
		<td>
			<font face="Verdana" size="1">Company</font>
		</td>

		<td>
		
			<select name="companymam" id="companymam" onchange="getFCBAmam()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
						</select>
		</td>
		
		   
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Business unit</font>
		</td>

		<td>
		
			<select name="businessunitmam" id="businessunitmam" style=width:100px >
			</select>
		</td>
		
		   
	</tr>
	
	
    <tr>
		<td>
			<font face="Verdana" size="1">Tanggal</font>
		</td>

		<td>
			
			<input class="datepicker" type="text" maxlength="10" name="datemam" id="datemam" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		</td>
	
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Group</font>
		</td>
		<td colspan=>
			<select name="groupmam" id="groupmam" style="width:200px">	
							<option value="">--All--</option> 
						<% 	q = " SELECT DISTINCT fccode,Fcname FROM StockGroup WHERE grouptype = 'STOCK'"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("fccode")%>"><%=rsCompany("Fcname")%></option>
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>			
			</select>
		</td>
		
		
	</tr>
	

	<tr>    
    <!--<td colspan="2" align="left">
		<input type="button" id="tampilmam" onclick="tampilData();" value="View" >
    </td>
	<td>
		<img  src="IMAGES/xls-icon.png" width=55 height=30 alt="CETAK XLS" onclick="cetakXLS()"> 
	</td>-->
	
 		<td colspan="4" align="left">
		<a class="button_menu_top" onclick="tampilData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakXLS()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
		</td>
	</tr>	
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	
	
	
	
	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	function viewEdit(i)
	{
		alert(i);
	}

	function tampilData()
	{
	    var dt 		 = $("#datemam").val();
		var pcompany = $("#companymam").val();
		var pfcba 	 = $("#businessunitmam").val();
		var ptype 	 = $("#groupmam").val();
		
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?dt=" + dt + "&idview=" + 9 + "&company=" + pcompany+"&fcba=" + pfcba+"&type=" + ptype,function(data) {
				$('#bodytable').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var x = json.results[i].stockcode
					//alert(x);
					var docketlist = '';
						if (json.results[i].get_stockminimum != 0 && json.results[i].closing_balance_day <= json.results[i].get_stockminimum) {
						docketlist += '<tr bgcolor= "red" id="tru'+i+'">';}
						else {
						docketlist += '<tr id="tru'+i+'">';}
						
						docketlist += '<td align="left">' + json.results[i].fcba + '</td>';
						docketlist += '<td align="left">' + json.results[i].stockcode + '</td>';
						docketlist += '<td align="left">' + json.results[i].fcname + '</td>';
						docketlist += '<td align="left">' + json.results[i].stockgroupname + '</td>';
						docketlist += '<td align="left">' + json.results[i].unitofmeasurement + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].get_stockminimum) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].opening_balance_day) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].received) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].issued) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].adjustment) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].closing_balance_day) + '</td>';
						docketlist += '<td align="center"><a target=_blank href="mam_data_view.Asp?stockcode='+json.results[i].stockcode+ '&fcba='+json.results[i].fcba+'&dt='+dt+'"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
						docketlist += '</tr>';
						$('#bodytable').append(docketlist);
						
					no++;
					$("#jumldata").val(jumlview);
					
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				
			});
		 
	} 


	

	function cetakXLS()
	{
		var rkbfrom 	= $("#rkbfrom").val();
		var rkbto 	= $("#rkbto").val();
		var fcba 	= $("#fcba").val();
		
		window.open('p_biayaxls.Asp?tglawal=' + rkbfrom + '&tglakhir=' + rkbto, '_blank');
		//window.location.href = "p_biayaxls.Asp?tahun=" + tahun2 + "&bulan=" + bulan2, '_blank';
	}


	</script>


	<br>
	<table width="917" class=data>
		<tr>
		    <th>FCBA</th>
			<th>STOCK</th>
			<th>NAMA STOCK</th>
			<th>GROUP</th>
			<th>SATUAN</th>
			<th>MINIMAL STOCK</th>
			<th>SALDO AWAL</th>
			<th>TERIMA</th>
			<th>KELUAR</th>
			<th>ADJUSMENT</th>
			<th>SALDO AKHIR</th>			
			<th></th>
		</tr>

		<tr id='tru'>
			<tbody id="bodytable">
			</tbody>
		</tr>
		<tr id="bodytable">
		</tr>
	</table>	
<%elseif Request("txttable") ="p_jrnreport" then%>
<input type="hidden" name="jmljrn" id="jmljrn" value="0" />
<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}


			.btn {
					background-color: DodgerBlue;
					border: none;
					color: white;
					padding: 12px 30px;
					cursor: pointer;
					font-size: 20px;
					}

					/* Darker background on mouse-over */
					.btn:hover {
					background-color: RoyalBlue;
				}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>

		
    <table>
	<tr>
		<td>
			<font face="Verdana" size="1">Company</font>
		</td>

		<td>
		
			<select name="company" id="company" onchange="getFCBA()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
						</select>
		</td>
		
		   
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Business unit</font>
		</td>

		<td>
		
			<select name="businessunit" id="businessunit" style=width:100px >
			</select>
		</td>
		
		   
	</tr>
	
    <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahun5,rstahun5
				sqltahun5 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun5 = Server.CreateObject("ADODB.Recordset")
				rstahun5.CursorType = 3
				rstahun5.CursorLocation = 3
				rstahun5.Open sqltahun5, connstr
				While Not rstahun5.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahun5("tahun")) then %>
					<option value="<%=rstahun5("tahun")%>" Selected ><%=rstahun5("tahun")%></option>
					<%else%>
					<option value="<%=rstahun5("tahun")%>"><%=rstahun5("tahun")%></option>
					<%end if%>					
				<%rstahun5.movenext
				wend 
				rstahun5.close 
				

			%>	
			</select>
		</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulan5,rsbulan5
				sqlbulan5 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
				Set rsbulan5 = Server.CreateObject("ADODB.Recordset")
				rsbulan5.CursorType = 3
				rsbulan5.CursorLocation = 3
				rsbulan5.Open sqlbulan5, connstr
				While Not rsbulan5.EOF 
                    if cint(MONTH(Session("v_Costbook")))=cint(rsbulan5("bulan")) then %>
					<option value="<%=rsbulan5("bulan")%>" Selected><%=rsbulan5("description")%></option>
					<%else%>
					<option value="<%=rsbulan5("bulan")%>"><%=rsbulan5("description")%></option>
					<%end if%>
				<%rsbulan5.movenext
				wend 
				rsbulan5.close 
				

			%>	
				
			</select>
		</td>
		
    </tr>
	<tr>
		<td>
			<font face="Verdana" size="1">SOURCE</font>
		</td>
		<td colspan=>
			
			<select name="sourcecode" id="sourcecode" style="width:200px">
			    <option value="">--ALL--</option>
			    <%
				dim sqlsource5,rssource5
				sqlsource5 = "select distinct sourcecode from bkmjournalhdr order by sourcecode" 
				Set rssource5 = Server.CreateObject("ADODB.Recordset")
				rssource5.CursorType = 3
				rssource5.CursorLocation = 3
				rssource5.Open sqlsource5, connstr
				While Not rssource5.EOF %>

					<option value="<%=rssource5("sourcecode")%>"><%=rssource5("sourcecode")%></option>

				<%rssource5.movenext
				wend 
				rssource5.close 
				

			%>	
				
			</select>
		</td>
		
		
	</tr>
	
	<tr> 
	
		<!--
    	<td colspan="" align="left">
			<img  alt="VIEW" src="IMAGES/view-file.png"  width=40 height=40 onclick="viewData()">
			
		</td>
		<td>
			<img  src="IMAGES/pdf-icon.png" width=40 height=40 alt="CETAK PDF" onclick="cetakPDF()"> &nbsp;
			<img  src="IMAGES/xls-icon.png" width=40 height=40 alt="CETAK XLS" onclick="cetakXLS()"> 
		</td>-->
		<td colspan="4" align="left">
		<a class="button_menu_top" onclick="viewData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top" onclick="tampilexceldetail()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
		</td>
		
	</tr>
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	<br>
	
	<table class=data width="700px">
	<tr>
		<th >FCBA</th>
		<th >TANGGAL</th>
		<th >NO.JURNAL</th>
		<th >NO.REFERENCE</th>
		<th >SOURCECODE</th>
		<th >AMOUNT</th>
		<th >POSTING</th>
		<th ></th>
	</tr>

	<tr id='trjrn'>
			<tbody id="bodytablejrn">
				
			</tbody>
	</tr>
	<tr id="bodytablejrn">
					<td colspan="5" align="right" class="huruftebal">Total (Rp)</td>
					<td align='right'><span id="totalamountview" class="huruftebal">0</span></td>
					
	</tr>
	</table>

	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	//get show
	//ambil data detail
	function viewData()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var sourcecode 	= $("#sourcecode").val();
		var pcompany 	= $("#company").val();
		var pfcba 	= $("#businessunit").val();
		
		var jumgl = parseInt($('#jmljrn').val());
		
		
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?tahun=" + tahun + "&bulan=" + bulan + "&idview=" + 5 + "&sourcecode=" + sourcecode+"&company=" + pcompany+"&fcba=" + pfcba,function(data) {
				$('#bodytablejrn').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {

					var postingflg = json.results[i].postingflg;
					if(postingflg==1){
						postingflg="YA"
					}else{
						postingflg="TIDAK"
					}
					
					var docketlist = '';
						docketlist += '<tr id="trjrn'+i+'">';
						docketlist += '<td align="left">' + json.results[i].fcba + '</td>';
						docketlist += '<td align="left">' + json.results[i].docdate + '</td>';
						docketlist += '<td align="left">' + json.results[i].docno + '</td>';
						docketlist += '<td align="left">' + json.results[i].remarks + '</td>';
						docketlist += '<td align="left">' + json.results[i].sourcecode + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].totamount) + '<input class="ratakanan lebarupah" type="hidden" name="totamount' + i + '" id="totamount' + i + '" value="' + json.results[i].totamount + '"/></td>';
						docketlist += '<td align="CENTER">' + postingflg + '</td>';
						docketlist += '<td align="center"><a target=_blank href="bkm_jurnal_view.Asp?bkmno='+json.results[i].docno+'&tglbkm='+ json.results[i].docdate + '"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
						
						docketlist += '</tr>';
						$('#bodytablejrn').append(docketlist);
					no++;
					$("#jmljrn").val(jumlview);
					
				}

				$('#loadingjudul').hide();
				$('#loading').hide();
				getPuttotal();
			
				
			});
		 
		
	}


	function getPuttotal() {
		var totamount	= 0;
		
		var jmljrn	= parseInt($('#jmljrn').val());
		for(var i=0; i<jmljrn; i++) {
			totamount	+= parseFloat($('#totamount'+i).val());
		}	
		//alert(totdebet);
		totamount    = numformat2(totamount);
		
		$('#totalamountview').html((totamount));	
	}

	function cetakPDF()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var Sourcecode 	= $("#sourcecode").val();
		var fcba 	= $("#fcba").val();
		window.open('http://localhost/syncwb/report_gl/cetakpdf/' + bulan + '/' + tahun + '/' + sourcecode + '/' + fcba, '_blank');
		
			
	}

	function cetakXLS()
	{
		//var tglskrg = $("#txttgl").val();
		//window.location.href = "p_glpdf.Asp";
		//alert('tes');
	}
	
	function tampilexceldetail()
	{
		var bulan 	    = $("#bulan").val();
		var tahun 	    = $("#tahun").val();
		var sourcecode 	= $("#sourcecode").val();
		var pcompany 	= $("#company").val();
		var pfcba 	    = $("#businessunit").val();
		
		window.open('p_listing_bkmjournal.asp?bulan=' + bulan + '&tahun=' + tahun + '&pcompany=' + pcompany+ '&pfcba=' + pfcba+ '&sourcecode=' + sourcecode, '_blank');
		
	}

	</script>

<%elseif Request("txttable") ="p_glreport" then%>
<input type="hidden" name="jmlgl" id="jmlgl" value="0" />
<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}


			.btn {
					background-color: DodgerBlue;
					border: none;
					color: white;
					padding: 12px 30px;
					cursor: pointer;
					font-size: 20px;
					}

					/* Darker background on mouse-over */
					.btn:hover {
					background-color: RoyalBlue;
				}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>

		
    <table>
	<tr>
		<td>
			<font face="Verdana" size="1">Company</font>
		</td>

		<td>
		
			<select name="companygl" id="companygl" onchange="getFCBAgl()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
						</select>
		</td>
		
		   
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Business unit</font>
		</td>

		<td>
		
			<select name="businessunitgl" id="businessunitgl" style=width:100px >
			</select>
		</td>
		
		   
	</tr>
	
    <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahun1,rstahun1
				sqltahun1 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "
				Set rstahun1 = Server.CreateObject("ADODB.Recordset")
				rstahun1.CursorType = 3
				rstahun1.CursorLocation = 3
				rstahun1.Open sqltahun1, connstr
				While Not rstahun1.EOF 
				 if cint(YEAR(Session("v_Costbook")))=cint(rstahun1("tahun")) then %>
					<option value="<%=rstahun1("tahun")%>" Selected ><%=rstahun1("tahun")%></option>
					<%else%>
					<option value="<%=rstahun1("tahun")%>"><%=rstahun1("tahun")%></option>
				<%end if%>	
				<%rstahun1.movenext
				wend 
				rstahun1.close 
				

			%>	
			</select>
		</td>
		
			<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
			<font face="Verdana" size="1" class="huruftebal">Opening  </font>
			</td>
			
			<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
				<font face="Verdana" size="1" class="huruftebal">Debet  </font>
			</td>
			
			<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
				<font face="Verdana" size="1" class="huruftebal">Credit  </font>
			</td>
			
			<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
				<font face="Verdana" size="1" class="huruftebal">Ending  </font>
			</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulan1,rsbulan1
				sqlbulan1 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
				Set rsbulan1 = Server.CreateObject("ADODB.Recordset")
				rsbulan1.CursorType = 3
				rsbulan1.CursorLocation = 3
				rsbulan1.Open sqlbulan1, connstr
				While Not rsbulan1.EOF 
					if cint(MONTH(Session("v_Costbook")))=cint(rsbulan1("bulan")) then %>
					<option value="<%=rsbulan1("bulan")%>" Selected><%=rsbulan1("description")%></option>
					<%else%>
					<option value="<%=rsbulan1("bulan")%>"><%=rsbulan1("description")%></option>
					<%end if%>
				<%rsbulan1.movenext
				wend 
				rsbulan1.close 
				

			%>	
				
			</select>
		</td>
		<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
			<font face="Verdana" size="1" class="huruftebal ratakanan"><span id="opening" class="huruftebal ">0</span></font>
		</td>
			
			<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
				<font face="Verdana" size="1" class="huruftebal"><span id="debet" class="huruftebal ">0</span></font>
			</td>
			
			<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
				<font face="Verdana" size="1" class="huruftebal"><span id="credit" class="huruftebal ">0</span></font>
			</td>
			
			<td bgcolor="#9CB770" style="color:#FFFFFF" class=ratakanan>
				<font face="Verdana" size="1" class="huruftebal"><span id="ending" class="huruftebal ">0</span></font>
			</td>
		
    </tr>
	<tr>
		<td>
			<font face="Verdana" size="1">COA</font>
		</td>
		<td colspan=>
			
			<select name="coa" id="coa" style="width:200px">
				<option value="all">--SEMUA--</option>
			</select>
		</td>
		
		
	</tr>

	<tr> 
	
		<!--
    	<td colspan="" align="left">
			<img  alt="VIEW" src="IMAGES/view-file.png"  width=40 height=40 onclick="viewData()">
			
		</td>
		<td>
			<img  src="IMAGES/pdf-icon.png" width=40 height=40 alt="CETAK PDF" onclick="cetakPDF()"> &nbsp;
			<img  src="IMAGES/xls-icon.png" width=40 height=40 alt="CETAK XLS" onclick="cetakXLS()"> 
		</td>-->
		<td colspan="4" align="left">
		<a class="button_menu_top" onclick="viewData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakXLS()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
		</td>
		
	</tr>
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	<br>
	
	<table class=data width="700px">
	<tr>
		<th >FCBA</th>
		<th >TANGGAL</th>
		<th >NO. JURNAL</th>
		<th >Debet</th>
		<th >Kredit</th>
	</tr>

	<tr id='trgl'>
			<tbody id="bodytablegl">
				
			</tbody>
	</tr>
	<tr id="bodytablegl">
					<td colspan="3" align="right" class="huruftebal">Total (Rp)</td>
					<td align='right'><span id="totaldebetview" class="huruftebal">0</span></td>
					<td align='right'><span id="totalcreditview" class="huruftebal">0</span></td>
	</tr>
	</table>

	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 showCoa();
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	//get show
	function showCoa() {
		
		//$('#coa').empty();
		$('#coa').select2({
			ajax: {
				delay: 1000,
				url : 'bkm_get_coa.asp',
				dataType: 'json',
				type: "POST",
				allowClear: true,
				data: function (params) {
						var query = {
										keyword: params.term
									}
						return query;
				}
			}
		});
	}

	//ambil data detail
	function viewData()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var coa 	= $("#coa").val();
		var pcompanygl = $("#companygl").val();
		var pfcbagl    = $("#businessunitgl").val();
		var jumgl = parseInt($('#jmlgl').val());
		
		if (coa != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?tahun=" + tahun + "&bulan=" + bulan + "&idview=" + 2 + "&coa=" + coa+"&company=" + pcompanygl+"&fcba=" + pfcbagl,function(data) {
				$('#bodytablegl').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					
					var docketlist = '';
						docketlist += '<tr id="trgl'+i+'">';
						docketlist += '<td align="left">' + json.results[i].fcba + '</td>';
						docketlist += '<td align="left">' + json.results[i].journaldate + '</td>';
						docketlist += '<td align="left">' + json.results[i].journalno + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].debet) + '<input class="ratakanan lebarupah" type="hidden" name="debet' + i + '" id="debet' + i + '" value="' + json.results[i].debet + '"/></td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].credit) + '<input class="ratakanan lebarupah" type="hidden" name="credit' + i + '" id="credit' + i + '" value="' + json.results[i].credit + '"/></td>';
						
						docketlist += '</tr>';
						$('#bodytablegl').append(docketlist);
					no++;
					$("#jmlgl").val(jumlview);
					
				}

				$('#loadingjudul').hide();
				$('#loading').hide();
				getPuttotal();
				viewdataHDR();
				
			});
		 
		} else {
			$('#bodytablegl').empty();
		}
	}


	function viewdataHDR()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var coa 	= $("#coa").val();
		var pcompanygl = $("#companygl").val();
		var pfcbagl    = $("#businessunitgl").val();
		
		if (coa != "") { 
				$.post("bkm_viewdata.asp?tahun=" + tahun + "&bulan=" + bulan + "&idview=" + 3 + "&coa=" + coa+"&company=" + pcompanygl+"&fcba=" + pfcbagl,function(data) {

				var json  = JSON.parse(data);
				var opening  = json.results[0].opening_balance;
				var debet  = json.results[0].debet;
				var credit  = json.results[0].credit;
				var ending  = json.results[0].ending_balance;

				opening    = numformat2(opening);
				debet    = numformat2(debet);
				credit    = numformat2(credit);
				ending    = numformat2(ending);
				$('#opening').html((opening));
				$('#debet').html((debet));
				$('#credit').html((credit));
				$('#ending').html((ending));
			});
		 
		} 
	}

	function getPuttotal() {
		var totdebet	= 0;
		var totkredit	= 0;
		
		var jmlgl	= parseInt($('#jmlgl').val());
		for(var i=0; i<jmlgl; i++) {
			totdebet	+= parseFloat($('#debet'+i).val());
			totkredit	+= parseFloat($('#credit'+i).val());
			
		}	
		//alert(totdebet);
		totdebet    = numformat2(totdebet);
		totkredit    = numformat2(totkredit);
		
		$('#totaldebetview').html((totdebet));
		$('#totalcreditview').html((totkredit));
	}

	function cetakPDF()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var coa 	= $("#coa").val();
		var fcba 	= $("#fcba").val();
		window.open('http://localhost/syncwb/report_gl/cetakpdf/' + bulan + '/' + tahun + '/' + coa + '/' + fcba, '_blank');
		
			
	}

	function cetakXLS()
	{
		//var tglskrg = $("#txttgl").val();
		//window.location.href = "p_glpdf.Asp";
		//alert('tes');
	}

	</script>
<%elseif Request("txttable") ="p_hppreport_daily" then%>
<input type="hidden" name="jmlhpp_dly" id="jmlhpp_dly" value="0" />
<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}


			.btn {
					background-color: DodgerBlue;
					border: none;
					color: white;
					padding: 12px 30px;
					cursor: pointer;
					font-size: 20px;
					}

					/* Darker background on mouse-over */
					.btn:hover {
					background-color: RoyalBlue;
				}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>

		
    <table>
	<!--
	<tr>
		<td>
			<font face="Verdana" size="1">Company</font>
		</td>

		<td>
		
			<select name="company" id="company" onchange="getFCBA()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
						</select>
		</td>
		
		   
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Business unit</font>
		</td>

		<td>
		
			<select name="businessunit" id="businessunit" style=width:100px >
			</select>
		</td>
		
		   
	</tr>
	-->
	<tr>
	<td>
		<font face="Verdana" size="1">From</font>
    </td>
    <td>
		<input class="datepicker" type="text" maxlength="10" name="hppfrom" id="hppfrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<font face="Verdana" size="1">To</font>
        <input class="datepicker" type="text" maxlength="10" name="hppto" id="hppto" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >		
		
	</td>
	
	<tr> 
	
		<td colspan="" align="left">
		<a class="button_menu_top" onclick="viewData()" ><img src="Images/View.png" border="0" ><span>View Data</span></a>
		<td>
		
	</tr>
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	<br>
	
	<table class=data width="700px">
		<tr>
			<th rowspan='2'>Date</th>
			<th colspan='4'>TBS</th>
			<th colspan='3'>CPO</th>
			<th colspan='3'>PK</th>
			<th colspan='5'>Biaya</th>
			<th colspan='4'>HPP</th>
			
			
		</tr>
		<tr>
			<th>Terima-INT</th>
			<th>Terima-EXT</th>
			<th>Olah</th>
			<th>Sisa</th>
			
			<th>Produksi</th>
			<th>Kirim</th>
			<th>Sisa</th>
			
			<th>Produksi</th>
			<th>Kirim</th>
			<th>Sisa</th>
			
			<th>TBS-INT</th>
			<th>TBS-EXT</th>
			<th>Olah</th>
			<th>Total</th>
			<th>Avg</th>
			
			<th>CPO</th>
			<th>Rp/Kg</th>
			<th>PK</th>
			<th>Rp/Kg</th>
		</tr>

	<tr id='trhpp_dly'>
			<tbody id="bodytablehpp_dly">
				
			</tbody>
	</tr>
	<tr id="bodytablehpp_dly">
					<td align="right" class="huruftebal">Total</td>
					<td align='right'><span id="tbsterima_int" class="huruftebal">0</span></td>
					<td align='right'><span id="tbsterima_ext" class="huruftebal">0</span></td>
					<td align='right'><span id="tbsolah" class="huruftebal">0</span></td>
					<td align='right'><span id="tbssisa" class="huruftebal">0</span></td>
					
					<td align='right'><span id="cpo_prod_hi" class="huruftebal">0</span></td>
					<td align='right'><span id="cpo_kirim_hi" class="huruftebal">0</span></td>
					<td align='right'><span id="cpo_sisa_hi" class="huruftebal">0</span></td>
					
					<td align='right'><span id="pk_prod_hi" class="huruftebal">0</span></td>
					<td align='right'><span id="pk_kirim_hi" class="huruftebal">0</span></td>
					<td align='right'><span id="pk_sisa_hi" class="huruftebal">0</span></td>
					
					<td colspan='2' align='right'><span id="tbsterima_int_amnt" class="huruftebal">0</span></td>
					<td align='right'><span id="prod_olah_amnt" class="huruftebal">0</span></td>
					<td align='right'><span id="prod_total_amnt" class="huruftebal">0</span></td>
					<td align='right'><span id="prod_avr_amnt" class="huruftebal">0</span></td>
					
					<td align='right'><span id="hpp_cpo_amnt" class="huruftebal">0</span></td>
					<td align='right'><span id="hpp_cpo_rpkg" class="huruftebal">0</span></td>
					<td align='right'><span id="hpp_pk_amnt" class="huruftebal">0</span></td>
					<td align='right'><span id="hpp_pk_rpkg" class="huruftebal">0</span></td>
					
					
					
	</tr>
	
	</table>

	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	//get show
	//ambil data detail
	function viewData()
	{
		var hppfrom = $("#hppfrom").val();
		var hppto 	= $("#hppto").val();
		var jumhpp = parseInt($('#jmlhpp_dly').val());
		
		var tbsterima_int	= 0;
		var tbsterima_ext	= 0;
		var tbsolah = 0;
		var tbssisa = 0;
		var cpo_prod_hi = 0;
		var cpo_kirim_hi = 0;
		var cpo_sisa_hi = 0;
		var pk_prod_hi = 0;
		var pk_kirim_hi = 0;
		var pk_sisa_hi = 0;
		var tbsterima_int_amnt = 0;
		var tbsterima_ext_amnt = 0;
		var prod_olah_amnt = 0;
		var prod_total_amnt = 0;
		var prod_avr_amnt = 0;
		var hpp_cpo_amnt = 0;
		var hpp_cpo_rpkg = 0;
		var hpp_pk_amnt = 0;
		var hpp_pk_rpkg = 0;
		
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?bkmfrom=" + hppfrom + "&bkmto=" + hppto + "&idview=" + 10,function(data) {
				$('#bodytablehpp_dly').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
		
 		
					var docketlist = '';
						docketlist += '<tr id="trhpp_dly'+i+'">';
						docketlist += '<td align="left">' + json.results[i].proddate + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].tbsterima_int) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].tbsterima_ext) + '</td>';						
						docketlist += '<td align="right">' + numformat2(json.results[i].tbsolah) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].tbssisa) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_prod_hi) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_kirim_hi) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_sisa_hi) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_prod_hi) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_kirim_hi) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_sisa_hi) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].tbsterima_int_amnt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].tbsterima_ext_amnt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].prod_olah_amnt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].prod_total_amnt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].prod_avr_amnt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].hpp_cpo_amnt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].hpp_cpo_rpkg) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].hpp_pk_amnt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].hpp_pk_rpkg) + '</td>';
						
						docketlist += '</tr>';
						$('#bodytablehpp_dly').append(docketlist);
						tbsterima_int	+= json.results[i].tbsterima_int;
						tbsterima_ext   += json.results[i].tbsterima_ext;
						tbsolah         += json.results[i].tbsolah ;
						tbssisa 		+= json.results[i].tbssisa;
						cpo_prod_hi 	+= json.results[i].cpo_prod_hi;
						cpo_kirim_hi 	+= json.results[i].cpo_kirim_hi;
						cpo_sisa_hi 	+= json.results[i].cpo_sisa_hi;
						pk_prod_hi 		+= json.results[i].pk_prod_hi;
						pk_kirim_hi 	+= json.results[i].pk_kirim_hi;
						pk_sisa_hi 		+= json.results[i].pk_sisa_hi;
						tbsterima_int_amnt += json.results[i].tbsterima_int_amnt;
						tbsterima_ext_amnt += json.results[i].tbsterima_ext_amnt;
						prod_olah_amnt 	+= json.results[i].prod_olah_amnt;
						prod_total_amnt += json.results[i].prod_total_amnt;
						prod_avr_amnt 	+= json.results[i].prod_avr_amnt;
					no++;
					
					
				}
				$("#jmlhpp_dly").val(jumlview);
				$('#loadingjudul').hide();
				$('#loading').hide();
				getPuttotal(tbsterima_int,tbsterima_ext,tbsolah,tbssisa,cpo_prod_hi,cpo_kirim_hi,cpo_sisa_hi,pk_prod_hi,pk_kirim_hi,pk_sisa_hi,tbsterima_int_amnt,tbsterima_ext_amnt,prod_olah_amnt,prod_total_amnt,prod_avr_amnt);
				
			});
		 
		
	}

    function getPuttotal(tbsterima_int,tbsterima_ext,tbsolah,tbssisa,cpo_prod_hi,cpo_kirim_hi,cpo_sisa_hi,pk_prod_hi,pk_kirim_hi,pk_sisa_hi,tbsterima_int_amnt,tbsterima_ext_amnt,prod_olah_amnt,prod_total_amnt,prod_avr_amnt) {
		
		tbsterima       = (tbsterima_int + tbsterima_ext);
		tbsolah_amnt    = ((tbsterima_int_amnt + tbsterima_ext_amnt)/tbsterima) * tbsolah;
        prod_tbs_amnt   = tbsolah_amnt;
		prod_total_amnt = prod_tbs_amnt + prod_olah_amnt;
		prod_avr_amnt   = prod_total_amnt / (cpo_prod_hi + pk_prod_hi);
		
		Prop_cpoqty   	= cpo_prod_hi / (cpo_prod_hi + pk_prod_hi);
		Prop_pkqty    	= pk_prod_hi / (cpo_prod_hi + pk_prod_hi);
		Prop_oer      	= cpo_prod_hi/tbsolah;
		Prop_ker      	= pk_prod_hi/tbsolah;
		
		Prop_cpo        = prod_avr_amnt * Prop_cpoqty * Prop_oer;
		Prop_pk         = prod_avr_amnt * Prop_pkqty * Prop_ker;

		hpp_cpo_qty    = cpo_prod_hi ,
		hpp_cpo_amnt   = (Prop_cpo/(Prop_cpo+Prop_pk))* prod_total_amnt;

		hpp_pk_qty     = pk_prod_hi,
		hpp_pk_amnt    = (Prop_pk/(Prop_cpo+Prop_pk))* prod_total_amnt;
		hpp_cpo_rpkg   = Math.round(hpp_cpo_amnt/hpp_cpo_qty);
		hpp_pk_rpkg    = Math.round(hpp_pk_amnt/hpp_pk_qty);

		$('#tbsterima_int').html(numformat2(tbsterima_int));
		$('#tbsterima_ext').html(numformat2(tbsterima_ext));
		$('#tbsolah').html(numformat2(tbsolah));
		$('#tbssisa').html(numformat2(tbssisa));
					
		$('#cpo_prod_hi').html(numformat2(cpo_prod_hi));
		$('#cpo_kirim_hi').html(numformat2(cpo_kirim_hi));
		$('#cpo_sisa_hi').html(numformat2(cpo_sisa_hi));
					
		$('#pk_prod_hi').html(numformat2(pk_prod_hi));
		$('#pk_kirim_hi').html(numformat2(pk_kirim_hi));
		$('#pk_sisa_hi').html(numformat2(pk_sisa_hi));
					
		$('#tbsterima_int_amnt').html(numformat2(prod_tbs_amnt));
		$('#prod_olah_amnt').html(numformat2(prod_olah_amnt));
		$('#prod_total_amnt').html(numformat2(prod_total_amnt));
		$('#prod_avr_amnt').html(numformat2(prod_avr_amnt));
					
		$('#hpp_cpo_amnt').html(numformat2(hpp_cpo_amnt));
		$('#hpp_cpo_rpkg').html(numformat2(hpp_cpo_rpkg));
		$('#hpp_pk_amnt').html(numformat2(hpp_pk_amnt));
		$('#hpp_pk_rpkg').html(numformat2(hpp_pk_rpkg));
		
		
	}
	
	function cetakPDF()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var Sourcecode 	= $("#sourcecode").val();
		var fcba 	= $("#fcba").val();
		window.open('http://localhost/syncwb/report_gl/cetakpdf/' + bulan + '/' + tahun + '/' + sourcecode + '/' + fcba, '_blank');
		
			
	}

	function cetakXLS()
	{
		//var tglskrg = $("#txttgl").val();
		//window.location.href = "p_glpdf.Asp";
		//alert('tes');
	}

	</script>
	
<%elseif Request("txttable") ="p_hppreport_monthly" then%>
<input type="hidden" name="jmlhpp_bln" id="jmlhpp_bln" value="0" />
<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}


			.btn {
					background-color: DodgerBlue;
					border: none;
					color: white;
					padding: 12px 30px;
					cursor: pointer;
					font-size: 20px;
					}

					/* Darker background on mouse-over */
					.btn:hover {
					background-color: RoyalBlue;
				}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>

		
    <table>
	<tr>
	<td>
    	<font face="Verdana" size="1">Tahun </font>
    </td>
	<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahun,rstahun
				sqltahun = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun = Server.CreateObject("ADODB.Recordset")
				rstahun.CursorType = 3
				rstahun.CursorLocation = 3
				rstahun.Open sqltahun, connstr
				While Not rstahun.EOF %>
					<%if cint(YEAR(Session("v_Costbook")))=cint(rstahun("tahun")) then %>
					<option value="<%=rstahun("tahun")%>" Selected ><%=rstahun("tahun")%></option>
					<%else%>
					<option value="<%=rstahun("tahun")%>"><%=rstahun("tahun")%></option>
					<%end if%>						
				<%rstahun.movenext
				wend 
				rstahun.close 
				

			%>	
			</select>
	</td>
	
	<tr> 
	
		<td colspan="" align="left">
		<a class="button_menu_top" onclick="viewData()" ><img src="Images/View.png" border="0" ><span>View Data</span></a>
		<td>
		
	</tr>
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	<br>
	
	<table class=data width="700px">
		<tr>
			<th rowspan='2'>Periode</th>
			<th colspan='2'>TBS Olah(kg)</th>
			<th colspan='2'>CPO (kg)</th>
			<th colspan='2'>PK (kg)</th>
			<th colspan='2'>Persediaan Awal (Rp)</th>			
			<th colspan='2'>Biaya</th>
			<th colspan='2'>Persediaan Akhir (Rp)</th>			
			<th colspan='2'>HPP</th>
			<th rowspan='2' colspan='2'>Action</th>
			
			
		</tr>
		<tr>
			<th>Masuk</th>
			<th>Olah</th>
			
			<th>Beli</th>
			<th>Produksi</th>
			
			<th>Beli</th>
			<th>Produksi</th>
			
			<th>CPO</th>
			<th>PK</th>
			
			<th>Produksi</th>
			<th>Depresiasi</th>
			
			
			
			<th>CPO</th>
			<th>PK</th>
			
			<th>CPO</th>
			<th>PK</th>
			
		</tr>

	<tr id='trhpp_bln'>
			<tbody id="bodytablehpp_bln">
				
			</tbody>
	</tr>
	
	</table>

	<script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	
	<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	//get show
	//ambil data detail
	function viewData()
	{
		var hppperiod= $("#tahun").val();
		var jumhpp = parseInt($('#jmlhpp_bln').val());
		
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?hppperiod=" + hppperiod + "&idview=" + 12,function(data) {
				$('#bodytablehpp_bln').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
		
 					var docketlist = '';
						docketlist += '<tr id="trhpp_bln'+i+'">';
						docketlist += '<td align="left">' + json.results[i].period + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].tbs_terima_qty) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].tbs_olah_qty) + '</td>';						
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_beli_qty) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_proses_qty) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_beli_qty) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_proses_qty) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_awal_amt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_awal_amt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cost_prod_amt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cost_asset_amt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_akhir_amt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_akhir_amt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].cpo_hpp_amt) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].pk_hpp_amt) + '</td>';
						docketlist += '<td align="center"><a target=_blank href="hpp_bulanan_calculate.Asp?period_month='+ json.results[i].accountingperiod + '&period_year='+json.results[i].accountingyear+'"><img src="IMAGES/process.png" alt="Generate"></a></td>';
						docketlist += '<td align="center"><a target=_blank href="hpp_bulanan_jurnal.Asp?period_month='+ json.results[i].accountingperiod + '&period_year='+json.results[i].accountingyear+'"><img src="IMAGES/view.png" alt="Posting"></a></td>';
					
						docketlist += '</tr>';
						$('#bodytablehpp_bln').append(docketlist);
						
					no++;
					
					
				}
				$("#jmlhpp_bln").val(jumlview);
				$('#loadingjudul').hide();
				$('#loading').hide();
			});
		 
		
	}

    function getPuttotal(tbsterima_int,tbsterima_ext,tbsolah,tbssisa,cpo_prod_hi,cpo_kirim_hi,cpo_sisa_hi,pk_prod_hi,pk_kirim_hi,pk_sisa_hi,tbsterima_int_amnt,tbsterima_ext_amnt,prod_olah_amnt,prod_total_amnt,prod_avr_amnt) {
		
		tbsterima       = (tbsterima_int + tbsterima_ext);
		tbsolah_amnt    = ((tbsterima_int_amnt + tbsterima_ext_amnt)/tbsterima) * tbsolah;
        prod_tbs_amnt   = tbsolah_amnt;
		prod_total_amnt = prod_tbs_amnt + prod_olah_amnt;
		prod_avr_amnt   = prod_total_amnt / (cpo_prod_hi + pk_prod_hi);
		
		Prop_cpoqty   	= cpo_prod_hi / (cpo_prod_hi + pk_prod_hi);
		Prop_pkqty    	= pk_prod_hi / (cpo_prod_hi + pk_prod_hi);
		Prop_oer      	= cpo_prod_hi/tbsolah;
		Prop_ker      	= pk_prod_hi/tbsolah;
		
		Prop_cpo        = prod_avr_amnt * Prop_cpoqty * Prop_oer;
		Prop_pk         = prod_avr_amnt * Prop_pkqty * Prop_ker;

		hpp_cpo_qty    = cpo_prod_hi ,
		hpp_cpo_amnt   = (Prop_cpo/(Prop_cpo+Prop_pk))* prod_total_amnt;

		hpp_pk_qty     = pk_prod_hi,
		hpp_pk_amnt    = (Prop_pk/(Prop_cpo+Prop_pk))* prod_total_amnt;
		hpp_cpo_rpkg   = Math.round(hpp_cpo_amnt/hpp_cpo_qty);
		hpp_pk_rpkg    = Math.round(hpp_pk_amnt/hpp_pk_qty);

		$('#tbsterima_int').html(numformat2(tbsterima_int));
		$('#tbsterima_ext').html(numformat2(tbsterima_ext));
		$('#tbsolah').html(numformat2(tbsolah));
		$('#tbssisa').html(numformat2(tbssisa));
					
		$('#cpo_prod_hi').html(numformat2(cpo_prod_hi));
		$('#cpo_kirim_hi').html(numformat2(cpo_kirim_hi));
		$('#cpo_sisa_hi').html(numformat2(cpo_sisa_hi));
					
		$('#pk_prod_hi').html(numformat2(pk_prod_hi));
		$('#pk_kirim_hi').html(numformat2(pk_kirim_hi));
		$('#pk_sisa_hi').html(numformat2(pk_sisa_hi));
					
		$('#tbsterima_int_amnt').html(numformat2(prod_tbs_amnt));
		$('#prod_olah_amnt').html(numformat2(prod_olah_amnt));
		$('#prod_total_amnt').html(numformat2(prod_total_amnt));
		$('#prod_avr_amnt').html(numformat2(prod_avr_amnt));
					
		$('#hpp_cpo_amnt').html(numformat2(hpp_cpo_amnt));
		$('#hpp_cpo_rpkg').html(numformat2(hpp_cpo_rpkg));
		$('#hpp_pk_amnt').html(numformat2(hpp_pk_amnt));
		$('#hpp_pk_rpkg').html(numformat2(hpp_pk_rpkg));
		
		
	}
	
	function cetakPDF()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var Sourcecode 	= $("#sourcecode").val();
		var fcba 	= $("#fcba").val();
		window.open('http://localhost/syncwb/report_gl/cetakpdf/' + bulan + '/' + tahun + '/' + sourcecode + '/' + fcba, '_blank');
		
			
	}

	function cetakXLS()
	{
		//var tglskrg = $("#txttgl").val();
		//window.location.href = "p_glpdf.Asp";
		//alert('tes');
	}

	</script>	
	

<%elseif Request("txttable") ="p_neraca_sementara_report" then%>
<input type="hidden" name="jmlbs" id="jmlbs" value="0" />

	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
	</style>
	
    <font face="Verdana" size="1">Select Data Here : <br><br>
	

    <table>
	<td>
			<font face="Verdana" size="1">Company</font>
		</td>

		<td>
		
			<select name="companybs" id="companybs" onchange="getFCBAbs()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company " & _
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
						</select>
		</td>
		
		   
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Business unit</font>
		</td>

		<td>
		
			<select name="businessunitbs" id="businessunitbs" style=width:100px >
			</select>
		</td>
		
		   
	</tr>
	
		<tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				sqltahun = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun = Server.CreateObject("ADODB.Recordset")
				rstahun.CursorType = 3
				rstahun.CursorLocation = 3
				rstahun.Open sqltahun, connstr
				While Not rstahun.EOF %>
					<%if cint(YEAR(Session("v_Costbook")))=cint(rstahun("tahun")) then %>
					<option value="<%=rstahun("tahun")%>" Selected ><%=rstahun("tahun")%></option>
					<%else%>
					<option value="<%=rstahun("tahun")%>"><%=rstahun("tahun")%></option>
					<%end if%>						
				<%rstahun.movenext
				wend 
				rstahun.close 
				

			%>	
			</select>
		</td>
   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulan,rsbulan
				sqlbulan = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
				Set rsbulan = Server.CreateObject("ADODB.Recordset")
				rsbulan.CursorType = 3
				rsbulan.CursorLocation = 3
				rsbulan.Open sqlbulan, connstr
				While Not rsbulan.EOF %>
                    <%if cint(MONTH(Session("v_Costbook")))=cint(rsbulan("bulan")) then %>
					<option value="<%=rsbulan("bulan")%>" Selected><%=rsbulan("description")%></option>
					<%else%>
					<option value="<%=rsbulan("bulan")%>"><%=rsbulan("description")%></option>
					<%end if%>
					
				<%rsbulan.movenext
				wend 
				rsbulan.close 
				

			%>	
				
			</select>
		</td>
    </tr>
	

	<tr> 

		<td colspan="4" align="left">
		<a class="button_menu_top" onclick="viewData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top"  onclick="getExport()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
		</td>
		
    </tr>
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	<br>
	
	<table class=data id="product-table" width="800px">
	<tr>
		<th >NO</th>
		<th >AKUN</th>
		<th >PERKIRAAN</th>
		<th >SALDO AWAL</th>
		<th >DEBET</th>
		<th >KREDIT</th>
		<th >SALDO AKHIR</th>
	</tr>

	<tr id='trbs'>
			<tbody id="bodytablebs">
				
			</tbody>
			
	</tr>
	<tr><td colspan=7><hr size="1" noshade></td></tr>
	<tr id="bodytablebs">
					<td colspan="3" align="right" class="huruftebal">Total (Rp)</td>
					<td align='right'><span id="totalawalbs"   class="huruftebal">0</span></td>					
					<td align='right'><span id="totaldebetbs"  class="huruftebal">0</span></td>
					<td align='right'><span id="totalcreditbs" class="huruftebal">0</span></td>
					<td align='right'><span id="totalakhirbs"  class="huruftebal">0</span></td>
	</tr>
	</table>	
	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script src="/distexcel/tableToExcel.js"></script>
	<script src="/distexcel/jquery.tableToExcel.js"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
		 //sss
	});

	function getExport()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var pcompanygl = $("#companybs").val();
		var pfcbagl    = $("#businessunitbs").val();
		var nm = 'Neraca_' + pcompanygl + '_' + pfcbagl + '_'+tahun+bulan+".xlsx";  
	
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "Neraca"
            }
          });
		//$('table').tblToExcel();
	}
		
	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}

    //ambil data detail
	function viewData()
	{
		var bulan 	= $("#bulan").val();
		var tahun 	= $("#tahun").val();
		var coa 	= $("#coa").val();
		var pcompanygl = $("#companybs").val();
		var pfcbagl    = $("#businessunitbs").val();
		var jumgl = parseInt($('#jmlbs').val());
		
		if (coa != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?tahun=" + tahun + "&bulan=" + bulan + "&idview=" + 6 + "&coa=" + coa+"&company=" + pcompanygl+"&fcba=" + pfcbagl,function(data) {
				$('#bodytablebs').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					
					var docketlist = '';
						docketlist += '<tr id="trbs'+i+'">';
						docketlist += '<td align="left">' + no + '</td>';
						docketlist += '<td align="left">' + json.results[i].fccode + '</td>';
						docketlist += '<td align="left">' + json.results[i].fcname + '</td>';
						docketlist += '<td style="display:none;" data-t="n" >' + json.results[i].opening_balance + '</td>';
						docketlist += '<td data-exclude="true" align="right">' + numformat2(json.results[i].opening_balance) + '<input class="ratakanan lebarupah" type="hidden" name="opening_balance' + i + '" id="opening_balance' + i + '" value="' + json.results[i].opening_balance + '"/></td>';
						
						docketlist += '<td style="display:none;" data-t="n" >' + json.results[i].debet + '</td>';
						docketlist += '<td data-exclude="true" align="right">' + numformat2(json.results[i].debet) + '<input class="ratakanan lebarupah" type="hidden" name="debet' + i + '" id="debet' + i + '" value="' + json.results[i].debet + '"/></td>';
						
						docketlist += '<td style="display:none;" data-t="n" >' + json.results[i].credit + '</td>';
						docketlist += '<td data-exclude="true" align="right">' + numformat2(json.results[i].credit) + '<input class="ratakanan lebarupah" type="hidden" name="credit' + i + '" id="credit' + i + '" value="' + json.results[i].credit + '"/></td>';
						
						docketlist += '<td style="display:none;" data-t="n" >' + json.results[i].ending_balance + '</td>';
						docketlist += '<td data-exclude="true" align="right">' + numformat2(json.results[i].ending_balance) + '<input class="ratakanan lebarupah" type="hidden" name="ending_balance' + i + '" id="ending_balance' + i + '" value="' + json.results[i].ending_balance + '"/></td>';
						
						docketlist += '</tr>';
						$('#bodytablebs').append(docketlist);
					no++;
					$("#jmlbs").val(jumlview);
					
				}

				$('#loadingjudul').hide();
				$('#loading').hide();
				getPuttotal();
				
				
			});
		 
		} else {
			$('#bodytablebs').empty();
		}
	}
	
	function getPuttotal() {
		var totawal	= 0;
		var totakhir	= 0;
		var totdebet	= 0;
		var totkredit	= 0;
		
		var jmlgl	= parseInt($('#jmlbs').val());
		for(var i=0; i<jmlgl; i++) {
			totawal	    += parseFloat($('#opening_balance'+i).val());
		    totakhir	+= parseFloat($('#ending_balance'+i).val());
			totkredit	+= parseFloat($('#credit'+i).val());
			totdebet	+= parseFloat($('#debet'+i).val());
			
		}	
		totawal     = numformat2(totawal);
		totakhir     = numformat2(totakhir);
		
		totdebet     = numformat2(totdebet);
		totkredit    = numformat2(totkredit);
		
		$('#totalawalbs').html((totawal));
		$('#totalcreditbs').html((totkredit));
		$('#totaldebetbs').html((totdebet));
		$('#totalakhirbs').html((totakhir));
		
	}

	//get show
	function showCoa() {
		
		//$('#coa').empty();
		$('#coa').select2({
			ajax: {
				delay: 1000,
				url : 'bkm_get_coa.asp',
				dataType: 'json',
				type: "POST",
				allowClear: true,
				data: function (params) {
						var query = {
										keyword: params.term
									}
						return query;
				}
			}
		});
	}
	

	</script>
<%elseif Request("txttable") ="listing_bnk" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
			.huruftebal {
					font-weight: bold;
			}
			.ratakanan {
					text-align: right;
			}
	</style>
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="fcba" id="fcba" value="<%=session("fcba")%>" />
    <font face="Verdana" size="1">Select Data Here : <br><br>
	

    <table>
    	<tr>
		<td>
			<font face="Verdana" size="1">Company</font>
		</td>

		<td>
		
			<select name="companybnk" id="companybnk" onchange="getFCBAbnk()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company a INNER JOIN user_company_tab b ON a.companycode=b.fccompanycode " & _
								" WHERE b.fcusercode='" & Session("CodeName") & "'" & _
								
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
						</select>
		</td>
		
		   
	</tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Business unit</font>
		</td>

		<td>
		
			<select name="businessunitbnk" id="businessunitbnk" onchange="getBANKbnk()" style=width:100px >
			</select>
		</td>
		
		   
	</tr>
	
	
    <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahunbnk" id="tahunbnk" style="width:120px">
			<%
				dim sqltahun_bnk,rstahun_bnk
				sqltahun_bnk = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun_bnk = Server.CreateObject("ADODB.Recordset")
				rstahun_bnk.CursorType = 3
				rstahun_bnk.CursorLocation = 3
				rstahun_bnk.Open sqltahun_bnk, connstr
				While Not rstahun_bnk.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahun_bnk("tahun")) then %>
					<option value="<%=rstahun_bnk("tahun")%>" Selected ><%=rstahun_bnk("tahun")%></option>
					<%else%>
					<option value="<%=rstahun_bnk("tahun")%>"><%=rstahun_bnk("tahun")%></option>
					<%end if%>					
				<%rstahun_bnk.movenext
				wend 
				rstahun_bnk.close 
				

			%>	
			</select>
		</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulanbnk" id="bulanbnk" style="width:120px">
			<%
				dim sqlbulan_bnk,rsbulan_bnk
				sqlbulan_bnk = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
				Set rsbulan_bnk = Server.CreateObject("ADODB.Recordset")
				rsbulan_bnk.CursorType = 3
				rsbulan_bnk.CursorLocation = 3
				rsbulan_bnk.Open sqlbulan_bnk, connstr
				While Not rsbulan_bnk.EOF 
                    if cint(MONTH(Session("v_Costbook")))=cint(rsbulan_bnk("bulan")) then %>
					<option value="<%=rsbulan_bnk("bulan")%>" Selected><%=rsbulan_bnk("description")%></option>
					<%else%>
					<option value="<%=rsbulan_bnk("bulan")%>"><%=rsbulan_bnk("description")%></option>
					<%end if%>
				<%rsbulan_bnk.movenext
				wend 
				rsbulan_bnk.close 
				

			%>	
				
			</select>
		</td>
		
    </tr>
	
	<tr>
		<td>
			<font face="Verdana" size="1">Bank</font>
		</td>
		<td colspan=>
		    <select name="bankbnk" id="bankbnk" style=width:300px >
			</select>
		</td>
		
		
	</tr>
	

	<tr>    
    <!--<td colspan="2" align="left">
		<input type="button" id="tampilmam" onclick="tampilData();" value="View" >
    </td>
	<td>
		<img  src="IMAGES/xls-icon.png" width=55 height=30 alt="CETAK XLS" onclick="cetakXLS()"> 
	</td>-->
	
 		<td colspan="4" align="left">
		<a class="button_menu_top" onclick="tampilData()" ><img src="Images/View.png" border="0" ><span> View Data</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakPDF()" ><img src="Images/butt_export_pdf.png" border="0" ><span> Export PDF</span></a>&nbsp;
		<a class="button_menu_top" onclick="cetakXLS()" ><img src="Images/butt_export_xls.png" border="0" ><span> Export Excel</span></a>&nbsp;
		</td>
	</tr>	
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses menampilkan data...</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	</table>
	
	
	
	
	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
	}

	function numformat(val) {
		var nStr = Math.round(val).toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}


	function viewEdit(i)
	{
		alert(i);
	}

	function tampilData()
	{
	    var pbulan 		 = $("#bulanbnk").val();
		var ptahun 		 = $("#tahunbnk").val();
		var pcompany     = $("#companybnk").val();
		var pfcba 	     = $("#businessunitbnk").val();
		var ptype 	     = $("#bankbnk").val();
		
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?tahun=" + ptahun + "&bulan=" + pbulan + "&idview=" + 11 + "&company=" + pcompany+"&fcba=" + pfcba+"&type=" + ptype,function(data) {
				$('#bodytablebnk').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var x = json.results[i].stockcode
					//alert(x);
					var docketlist = '';
						if (json.results[i].nending < 0 ) {
						docketlist += '<tr bgcolor= "red" id="tru'+i+'">';}
						else {
						docketlist += '<tr id="tru'+i+'">';}
						
						docketlist += '<td align="left">' + json.results[i].fcba + '</td>';
						docketlist += '<td align="left">' + json.results[i].fccode + '</td>';
						docketlist += '<td align="left">' + json.results[i].fcname + '</td>';
						docketlist += '<td align="left">' + json.results[i].accountcode + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].nopening) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].ndebet) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].ncredit) + '</td>';
						docketlist += '<td align="right">' + numformat2(json.results[i].nending) + '</td>';
						docketlist += '<td align="center"><a target=_blank href="bnk_data_view.Asp?bankcode='+json.results[i].fccode+ '&fcba='+json.results[i].fcba+'&jobcode='+json.results[i].accountcode+'&month='+pbulan+'&tahun='+ptahun+'"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
						docketlist += '</tr>';
						$('#bodytablebnk').append(docketlist);
						
					no++;
					$("#jumldata").val(jumlview);
					
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				
			});
		 
	} 


	

	function cetakXLS()
	{
		var rkbfrom 	= $("#rkbfrom").val();
		var rkbto 	= $("#rkbto").val();
		var fcba 	= $("#fcba").val();
		
		window.open('p_biayaxls.Asp?tglawal=' + rkbfrom + '&tglakhir=' + rkbto, '_blank');
		//window.location.href = "p_biayaxls.Asp?tahun=" + tahun2 + "&bulan=" + bulan2, '_blank';
	}


	</script>


	<br>
	<table width="917" class=data>
		<tr>
		    <th>FCBA</th>
			<th>KODE BANK</th>
			<th>NAMA BANK</th>
			<th>KODE AKUN</th>
			<th>SALDO AWAL</th>
			<th>DEBET</th>
			<th>KREDIT</th>
			<th>SALDO AKHIR</th>			
			<th></th>
		</tr>

		<tr id='trbnk'>
			<tbody id="bodytablebnk">
				
			</tbody>
			
		</tr>
		
	</table>	
	
<%elseif Request("txttable") ="Process_closing_Opr" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
	</style>
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
    <font face="Verdana" size="1">Select Data Period Payroll : <br><br>
	

    <table>
     <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahun_cls_opr,rstahun_cls_opr
				sqltahun_cls_opr = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun_cls_opr = Server.CreateObject("ADODB.Recordset")
				rstahun_cls_opr.CursorType = 3
				rstahun_cls_opr.CursorLocation = 3
				rstahun_cls_opr.Open sqltahun_cls_opr, connstr
				While Not rstahun_cls_opr.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahun_cls_opr("tahun")) then %>
					<option value="<%=rstahun_cls_opr("tahun")%>" Selected ><%=rstahun_cls_opr("tahun")%></option>
					<%else%>
					<option value="<%=rstahun_cls_opr("tahun")%>"><%=rstahun_cls_opr("tahun")%></option>
					<%end if%>					
				<%rstahun_cls_opr.movenext
				wend 
				rstahun_cls_opr.close 
				

			%>	
			</select>
		</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulan_cls_opr,rsbulan_cls_opr
				sqlbulan_cls_opr = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
				Set rsbulan_cls_opr = Server.CreateObject("ADODB.Recordset")
				rsbulan_cls_opr.CursorType = 3
				rsbulan_cls_opr.CursorLocation = 3
				rsbulan_cls_opr.Open sqlbulan_cls_opr, connstr
				While Not rsbulan_cls_opr.EOF 
                    if cint(MONTH(Session("v_Costbook")))=cint(rsbulan_cls_opr("bulan")) then %>
					<option value="<%=rsbulan_cls_opr("bulan")%>" Selected><%=rsbulan_cls_opr("description")%></option>
					<%else%>
					<option value="<%=rsbulan_cls_opr("bulan")%>"><%=rsbulan_cls_opr("description")%></option>
					<%end if%>
				<%rsbulan_cls_opr.movenext
				wend 
				rsbulan_cls_opr.close 
				

			%>	
				
			</select>
		</td>
		
    </tr>
	

	<tr>    
    <td colspan="4" align="left">
		<a class="button_menu_top" onclick="ProsesData()" ><img src="Images/process.png" border="0" ><span> Process</span></a>&nbsp;
	</td>
	
	</tr>	
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses ......</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	
	</table>
	<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
	
		<link rel="stylesheet" href="css/jquery-ui.css" />
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
		 $('#finished').hide();
		 $('#failed').hide();
	});

	// function ProsesData()
	// {
	// 	var tahun   = $("#tahun").val();
	// 	var bulan 	= $("#bulan").val();
			
	//  			$('#loadingjudul').show();
	// 			$('#loading').show();
	// 			$.post("bkm_calculate_payroll.asp?tahun=" + tahun + "&bulan=" + bulan ,function(data) {
	// 			$('#bodytablebnk').empty();
	// 			callback(data);
	// 			var json  = JSON.parse(data);
	// 			var jumlview  = json.results.length;
	// 			var no = 1;
	// 			for (var i = 0; i < jumlview; i++) {
	// 				var x = json.results[i].stockcode
	// 				var docketlist = '';
	// 					docketlist += '<tr id="tru'+i+'">';
	// 					docketlist += '<td>' + json.results[i] + '</td>';
	// 					docketlist += '</tr>';
	// 					$('#bodytablebnk').append(docketlist);
						
	// 				no++;
	// 				$("#jumldata").val(jumlview);
					
					
	// 			}
	// 			$('#loadingjudul').hide();
	// 			$('#loading').hide();
				
	// 		});
	// }

	function ProsesData()
	{
		var tahun   = $("#tahun").val();
		var bulan 	= $("#bulan").val();
			
	 			$('#loadingjudul').show();
				$('#loading').show();
				$('#vfinish').empty();
				$.post("bkm_calculate_payroll.asp?tahun=" + tahun + "&bulan=" + bulan ,function(data) {
				var json  = JSON.parse(data);
				var msg = json.results[0].msg

				
				if(data){

					$('#loading').hide();
					$('#loadingjudul').hide();
					$('#finished').show(msg);
					$('#vfinish').html(msg);
				}
				
				
			});
	}
	</script>
	<div id="finished">
		<span id="vfinish"></span>
			
	</div>
	<div id="failed" class="redcolor">
		Ada Error Pada Proses Closing Payroll
		<div id="errormessege"></div>
	</div>	

<%elseif Request("txttable") ="Process_Closing_fac" then%>
	<style>
				body {
				/*background: #ececec;*/
			}
			/*Hidden class for adding and removing*/
			.lds-dual-ring.hidden {
				display: none;
			}

			/*Add an overlay to the entire page blocking any further presses to buttons or other elements.*/
			.overlay {
				position: fixed;
				top: 0;
				left: 0;
				width: 100%;
				height: 100vh;
				background: rgba(0,0,0,.8);
				z-index: 999;
				opacity: 1;
				transition: all 0.5s;
			}
			
			/*Spinner Styles*/
			.lds-dual-ring {
				display: inline-block;
				width: 80px;
				height: 80px;
			}
			.lds-dual-ring:after {
				content: " ";
				display: block;
				width: 64px;
				height: 64px;
				margin: 5% auto;
				border-radius: 50%;
				border: 6px solid #fff;
				border-color: #fff transparent #fff transparent;
				animation: lds-dual-ring 1.2s linear infinite;
			}
			@keyframes lds-dual-ring {
				0% {
					transform: rotate(0deg);
				}
				100% {
					transform: rotate(360deg);
				}
			}
	</style>
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
    <font face="Verdana" size="1">Select Data Closing : <br><br>
	

    <table>
     <tr>
		<td>
			<font face="Verdana" size="1">Tahun</font>
		</td>

		<td>
		
			<select name="tahun" id="tahun" style="width:120px">
			<%
				dim sqltahun_cls_fac,rstahun_cls_fac
				sqltahun_cls_fac = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun_cls_fac = Server.CreateObject("ADODB.Recordset")
				rstahun_cls_fac.CursorType = 3
				rstahun_cls_fac.CursorLocation = 3
				rstahun_cls_fac.Open sqltahun_cls_fac, connstr
				While Not rstahun_cls_fac.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahun_cls_fac("tahun")) then %>
					<option value="<%=rstahun_cls_fac("tahun")%>" Selected ><%=rstahun_cls_fac("tahun")%></option>
					<%else%>
					<option value="<%=rstahun_cls_fac("tahun")%>"><%=rstahun_cls_fac("tahun")%></option>
					<%end if%>					
				<%rstahun_cls_fac.movenext
				wend 
				rstahun_cls_fac.close 
				

			%>	
			</select>
		</td>
		
		   
	</tr>
	<tr>
		<td>
			<font face="Verdana" size="1">Bulan</font>
		</td>
		<td>
			
			<select name="bulan" id="bulan" style="width:120px">
			<%
				dim sqlbulan_cls_fac,rsbulan_cls_fac
				sqlbulan_cls_fac = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
				Set rsbulan_cls_fac = Server.CreateObject("ADODB.Recordset")
				rsbulan_cls_fac.CursorType = 3
				rsbulan_cls_fac.CursorLocation = 3
				rsbulan_cls_fac.Open sqlbulan_cls_fac, connstr
				While Not rsbulan_cls_fac.EOF 
                    if cint(MONTH(Session("v_Costbook")))=cint(rsbulan_cls_fac("bulan")) then %>
					<option value="<%=rsbulan_cls_fac("bulan")%>" Selected><%=rsbulan_cls_fac("description")%></option>
					<%else%>
					<option value="<%=rsbulan_cls_fac("bulan")%>"><%=rsbulan_cls_fac("description")%></option>
					<%end if%>
				<%rsbulan_cls_fac.movenext
				wend 
				rsbulan_cls_fac.close 
				

			%>	
				
			</select>
		</td>
		
    </tr>
	

	<tr>    
    <td colspan="4" align="left">
		<a class="button_menu_top" onclick="ProsesData()" ><img src="Images/process.png" border="0" ><span> Process</span></a>&nbsp;
	</td>
	
	</tr>	
	<tr>
		<td colspan="4">
			<span id="loadingjudul">Sedang di proses ......</span>
		</td>
		<td><img id="loading" src="IMAGES/bartransparent.gif" /></td>
	</tr>
	
	</table>
	
	
	
	
	<script type="text/javascript">

	$(function() {	
		
		 $('#loadingjudul').hide();
		 $('#loading').hide();
	});

	function ProsesData()
	{
		var tahun   = $("#tahun").val();
		var bulan 	= $("#bulan").val();
			
	 			$('#loadingjudul').show();
				$('#loading').show();
				//$.post("bkm_viewdata.asp?bkmfrom=" + tglawal + "&bkmto=" + bkmto + "&idview=" + 4,function(data) {
		
				//var json  = JSON.parse(data);
				//var jumlview  = json.results.length;
				$('#loadingjudul').hide();
				$('#loading').hide();
				
			});
	}
	</script>

<%elseif Request("txttable") ="Nursery" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Nursery.jpg" border="0" >
	<span>Maintenance Nursery Setup</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
    <tr>
    <td><font face="Verdana" size="1">From</font></td>
	<td>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Batch</font></td>
	<td><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
    </td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Crop</font></td>
	<td><input type="text" size="15" maxlength="30" name="Crop"  style="background:#FFFFCC;">
    </td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Source</font></td>
	<td><input type="text" size="15" maxlength="30" name="Source"  style="background:#FFFFCC;">
    </td>
    </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >
    </td>
    </tr>
    </table>
    
    
<%elseif Request("txttable") ="NurseryOrder" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Nursery.jpg" border="0" >
	<span>Maintenance Nursery Order</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
    <tr>
    <td><font face="Verdana" size="1">From</font></td>
	<td>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Batch</font></td>
	<td><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >
    </td></tr>
    </table>
    
    
<%elseif Request("txttable") ="NurseryCulling" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
		<img src="ICON/Nursery.jpg" border="0" >
		<span>Maintenance Nursery Culling</span>
	</a>
    <br/><br/><br/>
    <font face="Verdana" size="1">Select Data Here : </font>
	<br/><br/>
	<table>
		<tr>
			<td>
				<font face="Verdana" size="1">From</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Document Number</font>
			</td>
			<td>
				<input type="text" size="50" maxlength="50" name="txtNurseryNo"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Batch</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
	
<%elseif Request("txttable") ="NurseryAdjustment" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
		<img src="ICON/Nursery.jpg" border="0" >
		<span>Maintenance Nursery Adjustment</span>
	</a>
    <br/><br/><br/>
    <font face="Verdana" size="1">Select Data Here : </font>
	<br/><br/>
	<table>
		<tr>
			<td>
				<font face="Verdana" size="1">From</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Document Number</font>
			</td>
			<td>
				<input type="text" size="50" maxlength="50" name="txtNurseryNo"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Batch</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
    
<%elseif Request("txttable") ="NurseryIssue" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
		<img src="ICON/Nursery.jpg" border="0" >
		<span>Maintenance Nursery Issue</span>
	</a>
    <br/><br/><br/>
    <font face="Verdana" size="1">Select Data Here :</font>
	<br/><br/>
	<table>
		<tr>
			<td>
				<font face="Verdana" size="1">From</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Document Number</font>
			</td>
			<td>
				<input type="text" size="50" maxlength="50" name="txtNurseryNo"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Batch</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Job</font>
			</td>
			<td>
				<select name="txtJobCode" class="chosen-select">
					<%cmdCallCombo "Job",Request("txtJobCode"),"activation","Y"%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
	</table>
    
<%elseif Request("txttable") ="NurseryLedger" then%>
    <br><img src="ICON/NurseryLedger.jpg" border="0">
    <font face="Verdana" size="1"><br>Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
		<font face="verdana" size="1">YEAR</font><input type=text name=txtyear value="<%=txtyear%>">
		<font face="verdana" size="1">MONTH</font><input type=text name=txtmonth value="<%=txtmonth%>">
	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="txtnurserycode"  style="background:#FFFFCC;">
    <input type="submit" value="Preview" >
    
<%Elseif Request("txttable") ="Nursery_Closing" then%>
    <b>Nursery Close</b><br /><br />
    <img src="ICON/database_process.png" border="0" width="80" height="80"><br /><br />
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Nursery")) &"/"& Year(Session("v_Nursery"))%>
	
<%Elseif Request("txttable") ="Asset_Closing" then%>
    <b>Asset Close</b><br /><br />
    <img src="ICON/database_process.png" border="0" width="80" height="80"><br /><br />
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Asset")) &"/"& Year(Session("v_Asset"))%>
    
<%elseif Request("txttable") ="Project" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Setup</b></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
 	<font face="Verdana" size="1">Crop</font>
		<% 	Dim rsCropproject
		q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
		Set rsCropproject = Server.CreateObject("ADODB.Recordset")
		rsCropproject.CursorType = 3
		rsCropproject.CursorLocation = 3
		rsCropproject.Open q, connstr
		if not rsCropproject.eof then%>
		<select name="Crop">         
		<option value=""></option> 
		<%
		Do While Not rsCropproject.eof
		%>
		<option value="<%=rsCropproject("FCCODE")%>"><%=rsCropproject("FCNAME")%></option> 
		<%rsCropproject.movenext
		loop
		End if
		rsCropproject.Close
		set rsCropproject = Nothing
		q=""
		%>
		</select>				
 	<font face="Verdana" size="1">Source</font><input type="text" size="15" maxlength="30" name="Source"  style="background:#FFFFCC;">
	<input type="submit" value="Preview" >
    
<%elseif Request("txttable") ="Project_Production" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Production</b></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
	<input type="submit" value="Preview" >
    
<%elseif Request("txttable") ="Project_Adjustment" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Adjusment</b></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
	<input type="submit" value="Preview" >
    
<%elseif Request("txttable") ="Project_Issue" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Issue</b></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
	<font face="Verdana" size="1">LC</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
	<input type="submit" value="Preview" >		
<%elseif Request("txttable") ="Project_Ledger" then%>
    <br><img src="ICON/NurseryLedger.jpg" border="0">
    <font face="Verdana" size="1"><br>Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
		<font face="verdana" size="1">Year</font><input type=text name=txtyear value="<%=txtyear%>">
		<font face="verdana" size="1">Month</font><input type=text name=txtmonth value="<%=txtmonth%>">
	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="txtprojectcode"  style="background:#FFFFCC;">
    <input type="submit" value="Preview" >
<%Elseif Request("txttable") ="Project_Close" then%>
    <b>Project Close</b><br /><br />
    <img src="ICON/database_process.png" border="0" width="80" height="80"><br /><br />
	<input type="submit" value="Process" >
<%Elseif Request("txttable") ="Building_Closing" then%>
    <b>Building Close</b><br /><br />
    <img src="ICON/database_process.png" border="0" width="80" height="80"><br /><br />
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Building")) &"/"& Year(Session("v_Building"))%>
<%Elseif Request("txttable") ="General_Charges_Closing" then%>
    <b>General Charge Close</b><br /><br />
    <img src="ICON/database_process.png" border="0" width="80" height="80"><br /><br />
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_General_Charges")) &"/"& Year(Session("v_General_Charges"))%>
<%Elseif Request("txttable") ="Harvesting_Closing" then%>
    <b>Close Harvesting</b></a><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Harvesting")) &"/"& Year(Session("v_Harvesting"))%>
<%Elseif Request("txttable") ="Financial_Monthly_Closing" then%>
    <b>Closing This Month</b></a>
    <br/><br/>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br/><br/>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Costbook")) &"/"& Year(Session("v_Costbook"))%> - <%=Session("CurP_Costbook")%>
	
<%Elseif Request("txttable") ="Open_Period_CostBook" then%>
    <b>Open Period Cost Book</b></a>
    <br/><br/>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br/><br/>
	<input type="submit" value="Process" >	
<%Elseif Request("txttable") ="Generate_Account_Report" then%>
    <b>Generate Account Report</b>
	<br/><br/>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br/><br/>
	<input type="submit" value="Process" >
<%Elseif Request("txttable") ="Generate_ERP_Report" then%>
    <b>Generate ERP Report</b>
	<br/><br/>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br/><br/>
	<input type="submit" value="Process" >	

	
<%Elseif Request("txttable") ="Journal_Reversed" then%>
    <b>Voucher Reversed</b>
	<br/><br/>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br/><br/>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Reversed")) &"/"& Year(Session("v_Reversed"))%>
	
<%Elseif Request("txttable") ="Previous_Month_UnClosing" then%>
    <b>un-Closing This Month</b>
    <br/><br/>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br/><br/>
	<input type="submit" value="Process" >
<%Elseif Request("txttable") ="Financial_Year_Closing" then%>
    <b>Close Financial Year</b>
   <br/><br/>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br/><br/>
	<input type="submit" value="Process" >
<%Elseif Request("txttable") ="GLMasterRegisterMonthly" then%>
    <b><br><img src="ICON/NurseryLedger.jpg" border="0"><br>Trial Balance</b></a>
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<form action=p_glmasterregistermonthly.asp method=post>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear value="<%=txtyear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth value="<%=txtmonth%>"></td>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td><input type=text name=txtjobgroup></td>
		<td><font face="verdana" size="1">JOB</font></td><td><input type=text name=txtjob></td>
		<td><input type=submit value="Preview" ></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txttable") ="Cost_Activity" then
    if request("txtYear")="" then
	   txtYear=Year(Now)
	end if
    if request("txtMonth")="" then
	   txtMonth=Month(Now)
	end if
	%>
    <b><br><img src="ICON/NurseryLedger.jpg" border="0"><br>Activity Base Costing</b></a>
	<form action=p_glmasterregistermonthly.asp method=post>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type=text name=txtyear width="4" value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type=text name=txtMonth width="3" value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td width="4"><input type=text name=txtjobgroup width="6"></td>
		<td><font face="verdana" size="1">JOB</font></td><td width="4"><input type=text name=txtjob width="10"></td></TR><TR>
		<td><font face="verdana" size="1">ST</font></td><td width="4"><input type=text name=txtST width="5"></td>
		<td><font face="verdana" size="1">SC </font></td><td width="4"><input type=text name=txtSC width="10"></td>
		<td><font face="verdana" size="1">TT</font></td><td width="4"><input type=text name=txtTT width="5"></td>
		<td><font face="verdana" size="1">TC</font></td><td width="4"><input type=text name=txtTC width="10"></td>
	</tr>

	<tr>
		<td colspan="3" align="right"><font face="verdana" size="1">Reference</font></td><td width="4"><input type=text name=txtreference width="4" value="<%=txtReference%>"></td>
		<td colspan="5"><input type=submit value="Preview" ></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txttable") ="Neraca" then%>
    <b><br><img src="ICON/Neraca.jpg" border="0"><br>Neraca</b></a>
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<form action=p_neraca.asp method=post>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth></td>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td><input type=text name=txtjobgroup></td>
		<td><font face="verdana" size="1">JOB</font></td><td><input type=text name=txtjob></td>
		<td><input type=submit value="Preview" ></td>
	</tr>
	</table>
	</form>
<%elseif Request("txttable") ="Harvesting" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Harvesting</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
	<tr>
	<td><font face="Verdana" size="1"> From</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
 	<td><font face="Verdana" size="1">Chit No</font></td>
	<td><input type="text" size="15" maxlength="30" name="CHITNO"></td>
	</tr>
	<tr>
	<td><font face="Verdana" size="1">Field Code</font></td>
	<td><select name="FIELDCODE">
	<%cmdCallCombo "Field","FIELDCODE","",""%>
	</select></td>
    </tr>
    <tr>
	<td><font face="Verdana" size="1">Cropcode Code</font></td>					
	<% 	Dim rsCrop
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCrop = Server.CreateObject("ADODB.Recordset")
	rsCrop.CursorType = 3
	rsCrop.CursorLocation = 3
	rsCrop.Open q, connstr
	if not rsCrop.eof then%>
	<td><select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCrop.eof
	%>
	<option value="<%=rsCrop("FCCODE")%>"><%=rsCrop("FCNAME")%></option> 
	<%rsCrop.movenext
	loop
	End if
	rsCrop.Close
	set rsCrop = Nothing
	q=""
	%></select>
    </td>
    </tr>
    <tr>
	<td><font face="Verdana" size="1">Product Code</font>
    </td>
	<td>
    <select name="ProductCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%>
		</select>
	</td>
    </tr>

		<tr>
			<td align="left"><font face="Verdana" size="1">Vehicle Status</font></td>
			<td align="left">
				<select id="OWN" name="OWN">
					<option value=""></option>
				<%
					Dim rsVehicleStatus, strVehicleStatus
					strVehicleStatus = " SELECT FCCODE, FCNAME FROM HARVESTINGVEHICLESTATUS ORDER BY FCCODE"
					
					Set rsVehicleStatus = Server.CreateObject("ADODB.Recordset")
					rsVehicleStatus.Open strVehicleStatus, connstr
					
					If Not rsVehicleStatus.EOF then
						Do While Not rsVehicleStatus.EOF
				%>
					<option value = "<%=rsVehicleStatus("FCCODE")%>" <%If request("OWN") = rsVehicleStatus("FCCODE") Then%> selected <%End If%>>
						<%=rsVehicleStatus("FCNAME")%>
					</option> 
				<%			rsVehicleStatus.movenext
						Loop
					End if
					rsVehicleStatus.Close
					set rsVehicleStatus = Nothing
				%>
				</select>
			</td>
		</tr>
		
		<tr>
            <td align="left"><font face="Verdana" size="1">Vehicle</font></td>
			<td align="left">
				<input id="Vehicle" type="text" size="15" maxlength="50" name="Vehicle">
			</td>
		</tr>
		
		<tr>
			<td align="left"><font face="Verdana" size="1">Transport Type</font></td>
			<td align="left">
				<select name="TRANSPORTTYPE">
					<option value=""></option>
				<%
					Dim rsTransportType, strTransportType
					strTransportType = " SELECT FCCODE, FCNAME FROM HARVESTINGTRANSPORTTYPE ORDER BY FCCODE"
					
					Set rsTransportType = Server.CreateObject("ADODB.Recordset")
					rsTransportType.Open strTransportType, connstr
					
					If Not rsTransportType.EOF then
						Do While Not rsTransportType.EOF
				%>
					<option value = "<%=rsTransportType("FCCODE")%>" <%If request("TRANSPORTTYPE") = rsTransportType("FCCODE") Then%> selected <%End If%>>
						<%=rsTransportType("FCNAME")%>
					</option> 
				<%			rsTransportType.movenext
						Loop
					End if
					rsTransportType.Close
					set rsTransportType = Nothing
				%>
				</select>
			</td>
		</tr>
		
		<tr>
            <td align="left"><font face="Verdana" size="1">Contract Number</font></td>
			<td align="left">
				<input type="text" size="30" maxlength="30" name="CONTRACTNUMBER">
			</td>
		</tr>
	
    <tr>
	<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
	</tr>
	</table>
<%elseif Request("txttable") ="spbharvesting" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Harvesting</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
		<tr>
			<td>
				<font face="Verdana" size="1"> From</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
	    </tr>
	    <tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit">
					<option value=""></option> 
					<% 	 
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboBST = Server.CreateObject("ADODB.Recordset")
						rsComboBST.CursorType = 3
						rsComboBST.CursorLocation = 3
						rsComboBST.Open q, connstr
						if not rsComboBST.eof then%>
						<%
						Do While Not rsComboBST.eof
						%>
							<option value="<%=rsComboBST("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
								<%End If%>><%=rsComboBST("fccode")%>
							</option> 
						<%rsComboBST.movenext
						loop
						End if
						rsComboBST.Close
						set rsComboBST = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
	    <tr>
	 		<td>
	 			<font face="Verdana" size="1">Chit No</font>
	 		</td>
			<td>
				<input type="text" size="15" maxlength="30" name="CHITNO">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Field Code</font>
			</td>
			<td>
				<select name="FIELDCODE">
					<%cmdCallCombo "Field","FIELDCODE","",""%>
				</select>
			</td>
	    </tr>
	    <tr>
			<td>
				<font face="Verdana" size="1">Cropcode Code</font>
			</td>					
				<% 	Dim rsCropSPB
				q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
				Set rsCropSPB = Server.CreateObject("ADODB.Recordset")
				rsCropSPB.CursorType = 3
				rsCropSPB.CursorLocation = 3
				rsCropSPB.Open q, connstr
				if not rsCropSPB.eof then%>
			<td>
				<select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
					<option value=""></option> 
					<%
						Do While Not rsCropSPB.eof
					%>
					<option value="<%=rsCropSPB("FCCODE")%>"><%=rsCropSPB("FCNAME")%></option> 
					<% rsCropSPB.movenext
						loop
						End if
						rsCropSPB.Close
						set rsCropSPB = Nothing
						q="" 
					%>
				</select>
	    	</td>
	    </tr>
	    <tr>
			<td>
				<font face="Verdana" size="1">Product Code</font>
	    	</td>
			<td>
	    		<select name="ProductCODE">
					<%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%>
				</select>
			</td>
	    </tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Vehicle Status</font>
			</td>
			<td align="left">
				<select id="OWN" name="OWN">
					<option value=""></option>
				<%
					Dim rsVehicleStatusSpb, strVehicleStatusSPb
					strVehicleStatusSPb = " SELECT FCCODE, FCNAME FROM HARVESTINGVEHICLESTATUS ORDER BY FCCODE"
					
					Set rsVehicleStatusSpb = Server.CreateObject("ADODB.Recordset")
					rsVehicleStatusSpb.Open strVehicleStatusSPb, connstr
					
					If Not rsVehicleStatusSpb.EOF then
						Do While Not rsVehicleStatusSpb.EOF
				%>
					<option value = "<%=rsVehicleStatusSpb("FCNAME")%>" <%If request("OWN") = rsVehicleStatusSpb("FCNAME") Then%> selected <%End If%>>
						<%=rsVehicleStatusSpb("FCNAME")%>
					</option> 
				<%			rsVehicleStatusSpb.movenext
						Loop
					End if
					rsVehicleStatusSpb.Close
					set rsVehicleStatusSpb = Nothing
				%>
				</select>
			</td>
		</tr>
		<tr>
	        <td align="left">
	        	<font face="Verdana" size="1">Vehicle</font>
	        </td>
			<td align="left">
				<input id="Vehicle" type="text" size="15" maxlength="50" name="Vehicle">
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Transport Type</font>
			</td>
			<td align="left">
				<select name="TRANSPORTTYPE">
					<option value=""></option>
				<%
					Dim rsTransportTypeSpb, strTransportTypeSpb
					strTransportTypeSpb = " SELECT FCCODE, FCNAME FROM HARVESTINGTRANSPORTTYPE ORDER BY FCCODE"
					
					Set rsTransportTypeSpb = Server.CreateObject("ADODB.Recordset")
					rsTransportTypeSpb.Open strTransportTypeSpb, connstr
					
					If Not rsTransportTypeSpb.EOF then
						Do While Not rsTransportTypeSpb.EOF
				%>
					<option value = "<%=rsTransportTypeSpb("FCCODE")%>" <%If request("TRANSPORTTYPE") = rsTransportTypeSpb("FCCODE") Then%> selected <%End If%>>
						<%=rsTransportTypeSpb("FCNAME")%>
					</option> 
				<%			rsTransportTypeSpb.movenext
						Loop
					End if
					rsTransportTypeSpb.Close
					set rsTransportTypeSpb = Nothing
				%>
				</select>
			</td>
		</tr>	
			<tr>
	            <td align="left"><font face="Verdana" size="1">Contract Number</font></td>
				<td align="left">
					<input type="text" size="30" maxlength="30" name="CONTRACTNUMBER">
				</td>
			</tr>
	    <tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
	</table>
<%elseif Request("txttable") ="HarvestingRound" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Harvesting Round</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
	<tr>
	<td><font face="verdana" size="1">Year</font></td><td><input type=text size="10" maxlength="4" name=HRNDYEAR></td>
    </tr>
    <tr>
	<td><font face="verdana" size="1">Month</font></td><td><input type=text size="10" maxlength="2" name=HRNDMONTH></td>
    </tr>
    <tr>
	<td><font face="Verdana" size="1">Field Code</font></td>
	<td><select name="FIELDCODE">
	<%cmdCallCombo "Field","FIELDCODE","",""%>
	</select></td>
    </tr>
    <tr>
	<td><font face="Verdana" size="1">Cropcode Code</font></td>					
	<% 	Dim rsCropA
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCropA = Server.CreateObject("ADODB.Recordset")
	rsCropA.CursorType = 3
	rsCropA.CursorLocation = 3
	rsCropA.Open q, connstr
	if not rsCropA.eof then%>
	<td><select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCropA.eof
	%>
	<option value="<%=rsCropA("FCCODE")%>"><%=rsCropA("FCNAME")%></option> 
	<%rsCropA.movenext
	loop
	End if
	rsCropA.Close
	set rsCropA = Nothing
	q=""
	%></select>
	</td>
    </tr>
    <tr>
	<td rowspan="2"><input type="submit" value="Preview" ></td>
	</tr>
	</table>
    <%elseif Request("txttable") ="TBSPriceList" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance TBS Price</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
	<tr>
	<td><font face="Verdana" size="1"> From</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="DATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
	<td><font face="Verdana" size="1">Product Code</font>
    </td>
	<td>
    <select name="ProductCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%>
		</select>
	</td>
    </tr>
    <tr>
	<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
	</tr>
	</table>
	<%elseif Request("txttable") ="HarvestingWriteOff" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="IMAGES/add2.gif" border="0" >
	<span>Maintenance Harvesting Write Off</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
      <table class="table_body">
    <tr>
        <td><font face="Verdana" size="1">From </font>
	</td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="WRITEOFFDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="WRITEOFFDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Writeoff Code</font>
    </td>
    <td>
    <input type="text" size="15" maxlength="30"  name="WriteOffCode"  style="background:#FFFFCC">
    </td>
    </tr>
    
   </table>
	<input type="submit" value="Preview" >
    
    <%elseif Request("txttable") ="HarvestingValue" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="IMAGES/add2.gif" border="0" >
	<span>Maintenance Harvesting Value</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1"> 
      <table class="table_body">
    <tr>
        <td><font face="Verdana" size="1">From </font>
	</td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="HARVESINGVALUEDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="HARVESINGVALUEDATE_tO" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Harvesting Value Code</font>
    </td>
    <td>
    <input type="text" size="15" maxlength="30"  name="HARVESTINGVALUECODE"  style="background:#FFFFCC">
    </td>
    </tr>
    
   </table>
	<input type="submit" value="Preview" >
    
<%elseif Request("txttable") ="Weather" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Weather</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="WIDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
    
	<input class="datepicker" type="text" maxlength="10" name="WIDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Division</font></td>
	<td><input type="text" size="15" maxlength="30" name="WIDIVISION">
    </td>
	</tr>
	</table>
	<input type="submit" value="Preview" >	

<%elseif Request("txttable") ="PRODUCT" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Product</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
 	<font face="Verdana" size="1">Product Code</font>
    </td>
    <td>
    <input type="text" size="15" maxlength="30" name="PRODUCTCODE">
    </td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Crop Code</font>
    </td>
    <td>	
	<% 	Dim rsCropP
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCropP = Server.CreateObject("ADODB.Recordset")
	rsCropP.CursorType = 3
	rsCropP.CursorLocation = 3
	rsCropP.Open q, connstr
	if not rsCropP.eof then%>
	<select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCropP.eof
	%>
	<option value="<%=rsCropP("FCCODE")%>"><%=rsCropP("FCNAME")%></option> 
	<%rsCropP.movenext
	loop
	End if
	rsCropP.Close
	set rsCropP = Nothing
	q=""
	%></select>
    </td>
    </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >
    </td></tr></table>
    
<%elseif Request("txttable") ="WEIGHBRIDGE" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Weightbridge</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
 	<font face="Verdana" size="1">Weighbridge Code</font>
    </td><td>
    <input type="text" size="15" maxlength="30" name="WEIGHBRIDGECODE">
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>

<%elseif Request("txttable") ="CUSTOMER" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Customer</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
 	<font face="Verdana" size="1">Customer Code</font></td><td><input type="text" size="15" maxlength="30" name="CUSTOMERCODE">
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>
    
    
<%elseif Request("txttable") ="GRADINGITEM" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Grading Item</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
 	<font face="Verdana" size="1">Item Code</font></td><td><input type="text" size="15" maxlength="30" name="ITEMCODE">
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>
    
<%elseif Request("txttable") ="OPHarvRate" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance OP HARVESTING RATE</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
 	<table>
	<tr>
	<td><font face="Verdana" size="1">Crop Code</font></td>	
	<% 	Dim rsCropo
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCropo = Server.CreateObject("ADODB.Recordset")
	rsCropo.CursorType = 3
	rsCropo.CursorLocation = 3
	rsCropo.Open q, connstr
	if not rsCropo.eof then%>
	<td><select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCropo.eof
	%>
	<option value="<%=rsCropo("FCCODE")%>"><%=rsCropo("FCNAME")%></option> 
	<%rsCropo.movenext
	loop
	End if
	rsCropo.Close
	set rsCropo = Nothing
	q=""
	%></select></td>
 	<td><font face="Verdana" size="1">Terrain Type</font></td>
	<td><input type="text" size="15" maxlength="30" name="TERRAIN_TYPE"></td>
	<td><input type="submit" value="Preview" ></td>
	</tr>
	</table>

<%elseif Request("txttable") ="WEIGHBRIDGETICKET" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Weighbridge Ticket</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<!--ADDED BY UKIE 14-NOV-2012-->
    <table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="Dateout_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="Dateout_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
    </td>
	</tr>
	 <tr>
    <td>
	<font face="Verdana" size="1">Posting Date</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="Posting_Date" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
    </td>
	</tr>
    <tr>
    <td>
	<!--END ADDING-->
	<font face="Verdana" size="1">Product Code</font>
    </td>
    <td>
		<select name="PRODUCT_PRODUCTCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","fcba",session("fcba")%>
		</select>
    </td>
	</tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Chit No</font>
    </td>
    <td><input type="text" size="15" maxlength="30" name="CHITNO">
    </td>
	</tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Customer Code</font>
    </td>
    <td>
		<select name="CUST_CUSTOMERCODE" class="chosen-select">
		<%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","fcba",session("fcba")%>
		</select>
        </td>
	</tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr>
    </table>
	
	
<%elseif Request("txttable") ="GRADINGRECORD" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Grading Record</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="DATETIME_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATETIME_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Grading No</font></td><td><input type="text" size="15" maxlength="30" name="GRADINGNO">
    </td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Vehicle No</font></td><td><input type="text" size="15" maxlength="30" name="VEHICLENO">
    </td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Chitno</font> </td><td><input type="text" size="15" maxlength="30" name="WBT_CHITNO">
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr>
    </table>
<%elseif Request("txttable") ="SHIFT" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Shift</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
 	<font face="Verdana" size="1">Shift Code</font></td><td><input type="text" size="15" maxlength="30" name="SHIFTCODE">
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>
<%elseif Request("txttable") ="SHIFTREPORT" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Shift Report</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="PRODUCTIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="PRODUCTIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Shift Code</font></td>
    <td><select name="SHIFT_SHIFTCODE"><%cmdCallComboSetup "SHIFT","SHIFT_SHIFTCODE","fcba",session("fcba")%></select>
    </td>
    </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>
<%elseif Request("txttable") ="PRODUCTSTORAGE" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Product Storage</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
 	<font face="Verdana" size="1">Storage Code</font>
    </td>
    <td><input type="text" size="15" maxlength="30" name="STORAGECODE">
    </td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Product Code</font>
    </td>
    <td>
		<select name="PRODUCT_PRODUCTCODE">
		<%cmdCallComboSetup "PRODUCT","PRODUCT_PRODUCTCODE","fcba",session("fcba")%>
		</select>
        </td>
    </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>
    	

<%elseif Request("txttable") ="PRODUCTSTORAGEREADING" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Product Storage Reading</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font>
    </td><td>
	<input class="datepicker" type="text" maxlength="10" name="READINGDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="READINGDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Storage Code</font>
    </td><td>
    <select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","fcba",session("fcba")%></select>	
    </td>
    </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>
    
    
<%elseif Request("txttable") ="PRODUCTADJUSTMENT" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Product Adjusment</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font>
    </td><td>
	<input class="datepicker" type="text" maxlength="10" name="ADJUSTMENTDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="ADJUSTMENTDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Adjusment No</font></td><td><input type="text" size="15" maxlength="30" name="ADJUSTMENTNO">
    </td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Storage Code</font></td><td><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","fcba",session("fcba")%></select>
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>

<%elseif Request("txttable") ="PRODUCTSTORAGETRANSFER" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Storage Transfer</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
    <tr>
    <td>
	<font face="Verdana" size="1">From</font>
    </td><td>
	<input class="datepicker" type="text" maxlength="10" name="TRANSFERDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="TRANSFERDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Transfer No</font></td><td><input type="text" size="15" maxlength="30" name="TRANSFERNO">
    </td></tr>
    <tr><td>
 	<font face="Verdana" size="1">Storage Code</font></td><td><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","fcba",session("fcba")%></select>
    </td></tr>
    <tr><td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>

<%elseif Request("txttable") ="DELIVERYCONTRACT" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Delivery Contract</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<TABLE>
	<TR>
	<TD><font face="Verdana" size="1">From</FONT></TD><TD><input class="datepicker" type="text" maxlength="10" name="STARTDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<font face="Verdana" size="1">To</font><input class="datepicker" type="text" maxlength="10" name="STARTDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</TD>
    </TR>
    <TR>
 	<TD><font face="Verdana" size="1">Contract Code</font></TD><TD><input type="text" size="15" maxlength="30" name="CONTRACTCODE">
    </TD>
    </TR>
	<TR>
 	<TD><font face="Verdana" size="1">Customer Code</font></TD><TD><select name="CUST_CUSTOMERCODE" class="chosen-select"><%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","fcba",session("fcba")%></select>
    </TD>
    </TR>
	<TR>
 	<TD>
    <font face="Verdana" size="1">Product Code</font></TD><TD><select name="PRODUCT_PRODUCTCODE"><%cmdCallComboSetup "PRODUCT","ProductCODE","fcba",session("fcba")%></select></TD>
    </TR>
	<TR>
 	<td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td></tr></table>
    
<%elseif Request("txttable") ="DELIVERYNOTE" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Delivery Note</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<TABLE>
	<TR>
	<TD><font face="Verdana" size="1">From</FONT></TD><TD><input class="datepicker" type="text" maxlength="10" name="DNOTE_DATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<font face="Verdana" size="1">To</font><input class="datepicker" type="text" maxlength="10" name="DNOTE_DATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</TD></TR>
	<TR>
 	<TD>
	<font face="Verdana" size="1">Posting Date</FONT></TD><TD><input class="datepicker" type="text" maxlength="10" name="Postingdate" size="10" value=""title="Type Date mm/dd/yyyy" >
	</TD>
	</TR>
    <TR>
 	<TD><font face="Verdana" size="1">Delivery NOte Code</font></TD><TD><input type="text" size="15" maxlength="30" name="DELIVERYNOTECODE"></TD></TR>
	<TR>
 	<TD><font face="Verdana" size="1">Contract Code</font></TD>
	    	<% 	Dim rsContract
			q = "select D.CONTRACTCODE from  DELIVERYCONTRACT D"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.CursorType = 3
			rsContract.CursorLocation = 3
			rsContract.Open q, connstr
			if not rsContract.eof then%>
			<td align="left"> <select name="CONTRACTCODE" title="CONTRACTCODE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%
			Do While Not rsContract.eoF%>
				<option value="<%=rsContract("CONTRACTCODE")%>"><%=rsContract("CONTRACTCODE")%></option> 
			<%rsContract.movenext
			loop
			End if
			rsContract.Close
			set rsContract = Nothing
			q=""
			%>
			</select></td>
	</TR><TR><TD colspan="2" align="left"><input type="submit" value="Preview" ></TD>	
	</TR>
	</TABLE>
    
    
<%elseif Request("txttable") ="DELIVERYRETURNNOTE" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Delivery Return Note</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<table>
	<tr>
	<td><font face="Verdana" size="1">From</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="RETURNDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
    <font face="Verdana" size="1">To</font><input class="datepicker" type="text" maxlength="10" name="RETURNDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
 	<td><font face="Verdana" size="1">Delivery Note Code</font></td><td><input type="text" size="15" maxlength="30" name="DNOTE_DELIVERYNOTECODE"></td>
	</tr>
	<tr> 
 	<td><font face="Verdana" size="1">WB Chit No</font></td><td><input type="text" size="15" maxlength="30" name="WBT_CHITNO"></td>
    </tr>
	<tr> 
 	<td><font face="Verdana" size="1">Storage Code</font></td><td><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","fcba",session("fcba")%></select></td>
    </tr>
	<tr> 
	<td colspan="2" align="left"><input type="submit" value="Preview" ></td>	
	</tr>
	</table>
    
    
<%elseif Request("txttable") ="DELIVERYNOTESEAL" then%>	
<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Delivery Note Seal</span></a>
	<br><br>
	<table>
	<tr>
	<td><font face="Verdana" size="1">Delivery Note Code</font></td><td><input type="text" size="15" maxlength="30" name="DELIVERYNOTECODE"></td></tr>
	<tr>
	<td><font face="Verdana" size="1">Seal No</font></td><td><input type="text" size="15" maxlength="30" name="Seal No"></td>	</tr>
    <tr>
	<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
    </tr>
    </table>
	
<%elseif Request("txttable") ="DeliveryQuality" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Delivery Product Quality</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	<font face="Verdana" size="1">From
	<input class="datepicker" type="text" maxlength="10" name="QUALITYDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="QUALITYDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
 	<font face="Verdana" size="1">Delivery Note Code</font><input type="text" size="15" maxlength="30" name="DELIVERYNOTE">
	<input type="submit" value="Preview" >	
<%elseif Request("txttable") ="EmployeeAllowdedDetail" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Data</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>	
	<font face="Verdana" size="1">Employee Code</font></td><td><input type="text" size="15" maxlength="30" name="FCCODE">
	</td>
    </tr>
	<tr>
	<td><font face="Verdana" size="1">Allowance</font></td>
	<td>
	<select name="ALLOWDEDCODE" class="chosen-select">
		<%cmdCallCombo "allowded","","",""%>
	</select>
	</td>
	</tr>
	<tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td>
    </tr>
	</table>
<%elseif Request("txttable") ="UntaxableIncome" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Data</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>
	<font face="Verdana" size="1">Code</font></td><td><input type="text" size="15" maxlength="30" name="FCCODE">
	</td>
    </tr>
	<tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td>
    </tr>
	</table>
<%elseif Request("txttable") ="MaritalStatus" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Data</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>	
	<font face="Verdana" size="1">Code</font></td><td><input type="text" size="15" maxlength="30" name="FCCODE">
	</td>
    </tr>
	<tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td>
    </tr>
	</table>
<%elseif Request("txttable") ="RangeTaxGroup" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Data</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>	
	<font face="Verdana" size="1">Code</font></td><td><input type="text" size="15" maxlength="30" name="FCCODE">
	</td>
    </tr>
	<tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td>
    </tr>
	</table>
<%elseif Request("txttable") ="Employee" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<span>Maintenance Employee</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
    <tr>
    <td>	
	<font face="Verdana" size="1">Code</font></td><td><input type="text" size="15" maxlength="30" name="FCCODE"></td>
    </tr>
    <tr>
    <td>	
 	<font face="Verdana" size="1">Employee Name</font></td><td><input type="text" size="15" maxlength="30" name="FCNAME"></td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Gang</font>
    </td><td>
	<select name="GANGCODE">
		<%cmdCallCombo "Gang","","",""%>
	</select>
    </td>
    </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >	
    </td>
    </tr>
    </table>
<%elseif Request("txttable") ="assetgroup" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Asset Group</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td align="left"><font face="Verdana" size="1">Asset Type Code</font></td>
			<td align="left">
				<select name="assettypecode" style="width: 100%">
					<option></option>
					<option value="CAPEX" <%if assettypecode = "CAPEX" then%>selected<%end if%>>CAPEX</option>
					<option value="LVA" <%if assettypecode = "LVA" then%>selected<%end if%>>LVA</option>
				</select> 
			</td>
		</tr> 
		<tr>
			<td align="left"><font face="Verdana" size="1">Commercial Depreciation Method</font></td>
			<td align="left">
				<select id="depreciationmethod" name="depreciationmethod" onchange="loaddepreciationrange();"  style="width: 100%">
					<option></option>
					<option value="NOT DEPRECIABLE" <%if depreciationmethod = "NOT DEPRECIABLE" then%>selected<%end if%>>NOT DEPRECIABLE</option>
					<option value="STRAIGHT LINE" <%if depreciationmethod = "STRAIGHT LINE" then%>selected<%end if%>>STRAIGHT LINE</option>
				</select> 
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="fixedasset_taxgroup" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Asset Tax Group</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td>
				<font face="Verdana" size="1">Asset Tax Group Code</font>
			</td>
			<td>		
				<input type="text" size="15" maxlength="30"  name="assettaxgroupcode"  >
			</td>
		</tr> 
		<tr>
			<td align="left"><font face="Verdana" size="1">Commercial Depreciation Method</font></td>
			<td align="left">
				<select id="depreciationmethod" name="depreciationmethod" onchange="loaddepreciationrange();">
					<option></option>
					<option value="NOT DEPRECIABLE" <%if depreciationmethod = "NOT DEPRECIABLE" then%>selected<%end if%>>NOT DEPRECIABLE</option>
					<option value="STRAIGHT LINE" <%if depreciationmethod = "STRAIGHT LINE" then%>selected<%end if%>>STRAIGHT LINE</option>
				</select> 
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="assetmaster" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Asset Master</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
			<td>
				<font face="Verdana" size="1">Asset Master Code</font>
			</td>
			<td>		
				<input type="text" size="15" maxlength="30"  name="fccode"  >
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Asset Group</font>
			</td>
			<td align="left">
				<select name="assetgroupcode">
					<option></option>
					<%
					qlov = "SELECT fccode, fcname FROM assetgroup ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if assetgroupcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Is Active</font>
			</td>
			<td align="left">
				<select name="isactive">
					<option value="1" <%if isactive = "1" then%>selected<%end if%>>TRUE</option>
					<option value="0" <%if isactive = "0" then%>selected<%end if%>>FALSE</option>
				</select>
			</td>
		</tr>
	</table>
    <input type="submit" value="Preview" >	
<%elseif Request("txttable") ="fixedasset_trans" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Fixed Asset</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1"> Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboFAB
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Asset Code</font>
			</td>
			<td align="left">
				<select id="assetcode" name="assetcode" onchange="javascript:getAssetInfo();">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM asset ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if assetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Is Active</font>
			</td>
			<td align="left">
				<select name="isactive">
					<option value="TRUE" <%if cstr(""&isactive) = "TRUE" then%>selected<%end if%>>TRUE</option>
					<option value="FALSE" <%if cstr(""&isactive) = "FALSE" then%>selected<%end if%>>FALSE</option>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
	</table>
    <input type="submit" value="Preview" >
<%elseif Request("txttable") ="activity_norm" then%>
<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Activity Norm</span></a>
    <br><br><br>
	 <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">		
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Job Code</font>
			</td>
			<td align="left">
				<select id="jobcode" name="jobcode" class="chosen-select">
					<option value=""></option>
					<%
					qlov = " SELECT fccode, fcname FROM job where fcba = '"& Session("fcBA") &"' order by fccode" 
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if jobcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
	</table>
    <input type="submit" value="Preview" >
<%elseif Request("txttable") ="activity_norm_harv" then%>
<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Activity Norm Harvesting</span></a>
    <br><br><br>
	 <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">		
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Planting Year</font>
			</td>
			<td align="left">
				<select id="plantingyear" name="plantingyear">
					<option value=""></option>
					<%
					qlov = " SELECT distinct to_char(plantingyear)plantingyear FROM ACTIVITY_NORM_HARV where fcba = '"& Session("fcBA") &"' order by plantingyear" 
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("plantingyear")%>" <%if plantingyear = rslov("plantingyear") then%>selected<%end if%>><%=rslov("plantingyear")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
	</table>
    <input type="submit" value="Preview" >
<%elseif Request("txttable") ="materialactivity_norm" then%>
<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Material Activity Norm</span></a>
    <br><br><br>
	 <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">		
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Jobcode</font>
			</td>
			<td align="left">
				<select id="jobcode" name="jobcode" class="chosen-select">
					<option value=""></option>
					<%
					qlov = " SELECT fccode, fcname FROM job where fcba = '"& Session("fcBA") &"' order by fccode" 
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if jobcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Active Material</font>
			</td>
			<td align="left">
				<select id="actmat" name="actmat" class="chosen-select">
					<option value=""></option>
					<%
					qlov =  " SELECT fccode, fcname FROM activematerial order by fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if actmat = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
	</table>
    <input type="submit" value="Preview" >	
<%elseif Request("txttable") ="fixedasset_auc" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Fixed Asset AUC</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1"> Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboFAUC
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAUC = Server.CreateObject("ADODB.Recordset")
						rsComboFAUC.CursorType = 3
						rsComboFAUC.CursorLocation = 3
						rsComboFAUC.Open q, connstr
						if not rsComboFAUC.eof then%>
						<%
						Do While Not rsComboFAUC.eof
						%>
							<option value="<%=rsComboFAUC("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAUC("fccode")) then %> selected 
								<%End If%>><%=rsComboFAUC("fccode")%>
							</option> 
						<%rsComboFAUC.movenext
						loop
						End if
						rsComboFAUC.Close
						set rsComboFAUC = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Document No</font>
			</td>
			<td align="left">
				<input type="text" maxlength="50" name="documentno" size="30" >
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Fixed Asset Code</font>
			</td>
			<td align="left">
				<input type="text" maxlength="50" name="fixedassetcode" size="30" >
			</td>
		</tr>
	</table>
    <input type="submit" value="Preview" >	
 <%elseif Request("txttable") ="fixedasset_mature" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Fixed Asset Mature</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1"> Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboFAM
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAM = Server.CreateObject("ADODB.Recordset")
						rsComboFAM.CursorType = 3
						rsComboFAM.CursorLocation = 3
						rsComboFAM.Open q, connstr
						if not rsComboFAM.eof then%>
						<%
						Do While Not rsComboFAM.eof
						%>
							<option value="<%=rsComboFAM("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAM("fccode")) then %> selected 
								<%End If%>><%=rsComboFAM("fccode")%>
							</option> 
						<%rsComboFAM.movenext
						loop
						End if
						rsComboFAM.Close
						set rsComboFAM = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Document No</font>
			</td>
			<td align="left">
				<input type="text" maxlength="50" name="documentno" size="30" >
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Field Code</font></td>
			<td align="left">
				<input type="text" maxlength="50" id="fieldcode" name="fieldcode" size="30" >
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Fixed Asset Code</font></td>
			<td align="left">
				<input type="text" maxlength="50" name="fixedassetcode" size="30" >
			</td>
		</tr>
	</table>
    <input type="submit" value="Preview" >	
<%elseif Request("txttable") ="fixedasset_adjustment" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Fixed Asset Adjustment</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1"> Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Fixed Asset Code Adjusment</font></td>
			<td align="left">
				<input type="text" maxlength="50" name="assetadjustcode" size="30" >
			</td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Fixed Asset Code</font></td>
			<td align="left">
				<input type="text" maxlength="50" name="assetcode" size="30" >
			</td>
		</tr>		
	</table>
    <input type="submit" value="Preview" >
<%elseif Request("txttable") ="fixedasset_transporttask" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Fixed Asset Transport Task</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1"> Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit Destination</font>
			</td>
			<td align="left">
				<select name="businessunitDestin" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunitDestin)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Fixed Asset Transport Task No</font>
			</td>
			<td align="left">
				<select id="TransportTaskNo" name="TransportTaskNo" onchange="">
					<option value=""></option>
					<%
					qlov = "SELECT fccode FROM FIXEDASSET_TRANSPORTASK Where fcba = '" & Session("fcBA") & "' ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if TransportTaskNo = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") %></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Fixed Asset Code</font>
			</td>
			<td align="left">
				<select id="assetcode" name="assetcode" onchange="">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM asset ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if assetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		
	</table>
    <input type="submit" value="Preview" >	
<%elseif Request("txttable") ="fixedasset_receivedtask" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Fixed Asset Received Task</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1"> Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit Source</font>
			</td>
			<td align="left">
				<select name="businessunitSource" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunitDestin)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Fixed Asset Transport Task No</font>
			</td>
			<td align="left">
				<select id="TransportTaskNo" name="TransportTaskNo" onchange="">
					<option value=""></option>
					<%
					qlov = "SELECT fccode FROM FIXEDASSET_TRANSPORTASK Where fcba = '" & Session("fcBA") & "' ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if TransportTaskNo = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") %></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Fixed Asset Code</font>
			</td>
			<td align="left">
				<select id="assetcode" name="assetcode" onchange="">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM asset ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if assetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		
	</table>
    <input type="submit" value="Preview" >	 
<%elseif Request("txttable") ="fixedasset_writeoff" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<span>Maintenance Fixed Asset Write Off</span></a>
    <br><br><br>
    <font face="Verdana" size="1">Select Data Here : <br><br>
	
	<font face="Verdana" size="1"> 
	<table class="table_body">
		<tr>
		    <td>
				<font face="Verdana" size="1"> Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboFAB = Server.CreateObject("ADODB.Recordset")
						rsComboFAB.CursorType = 3
						rsComboFAB.CursorLocation = 3
						rsComboFAB.Open q, connstr
						if not rsComboFAB.eof then%>
						<%
						Do While Not rsComboFAB.eof
						%>
							<option value="<%=rsComboFAB("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboFAB("fccode")) then %> selected 
								<%End If%>><%=rsComboFAB("fccode")%>
							</option> 
						<%rsComboFAB.movenext
						loop
						End if
						rsComboFAB.Close
						set rsComboFAB = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Fixed Asset Write Off</font>
			</td>
			<td align="left">
				<select id="fcCode" name="fcCode" >
					<option value=""></option>
					<%
					qlov = "SELECT fccode FROM fixedasset_writeoff where fcba = '"& session("fcba") &"' ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if assetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Fixed Asset Write Off Type</font>
			</td>
			<td align="left">
				<select id="fctype" name="fctype" onchange="GetDepreciationRange();" class="chosen-select" style="width: 40%;">
					<option></option>
					<option value="SALES" <%if fctype = "SALES" then%>selected<%end if%>>SALES</option>
					<option value="DISPOSAL" <%if fctype = "DISPOSAL" then%>selected<%end if%>>DISPOSAL</option>
				</select>
			</td>
		</tr>
		
	</table>
    <input type="submit" value="Preview" >	 
<%elseif Request("txttable") ="recurring" then%>
<%
Yr = Year(Session("v_Bank"))
txtyear = request("txtyear")
if "" & txtyear = "" then txtyear = Yr
MinYr = Yr - 2
MaxYr = Yr + 2
%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
		<img src="IMAGES/add2.gif" border="0">
		<span>Maintenance Recurring</span>
	</a>
    <br/><br/><br/>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    <table>
		<tr>
			<td><font face="Verdana" size="1">Accounting Year</font></td>
			<td>
				<select name="txtyear">
				<%For xi = MinYr to MaxYr%>
					<Option Value="<%=xi%>" <%If cint(xi) = cint(txtyear) then%> Selected <%End If%>><%=xi%></option>
				<%Next%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Document No</font></td>
			<td><input type="text" size="50" maxlength="35" name="documentno"></td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Use Settlement</font></td>
			<td>
				<select name="txtusesettlement">
					<Option Value="" <%If "" & request("txtusesettlement") = "" then%> Selected <%End If%>></option>
					<Option Value="TRUE" <%If request("txtusesettlement") = "TRUE" then%> Selected <%End If%>>TRUE</option>
					<Option Value="FALSE" <%If request("txtusesettlement") = "FALSE" then%> Selected <%End If%>>FALSE</option>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Status</font></td>
			<td>
				<select name="txtstatus">
					<Option Value="" <%If "" & request("txtstatus") = "" then%> Selected <%End If%>></option>
					<Option Value="Planned" <%If request("txtstatus") = "Planned" then%> Selected <%End If%>>Planned</option>
					<Option Value="Confirmed" <%If request("txtstatus") = "Confirmed" then%> Selected <%End If%>>Confirmed</option>
					<Option Value="Executed" <%If request("txtstatus") = "Executed" then%> Selected <%End If%>>Executed</option>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
    </table>
<%Elseif Request("txttable") ="Stock_Product" then%>
    <b>Stock Product Ledger</b><img src="ICON/StockLedger.jpg" border="0" ></a>
	
	<table>
	<tr>
		<td><font face="verdana" size="1">Year</font></td><td><input type=text name=txtyear></td>
		<td><font face="verdana" size="1">Month</font></td><td><input type=text name=txtmonth></td>
		<td><font face="verdana" size="1">Storage Code</font></td><td><input type=text name=txtSTORAGE_CODE></td>
		<td><font face="verdana" size="1">Product Code</font></td><td><input type=text name=txtPRODUCT_CODE></td>
		<td><input type=submit value="Preview" ></td>
	</tr>
	</table>
	
<%Elseif Request("txttable") ="Close_Delivery" then%>
    <b>Close Delivery</b></a><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" >
    
<%Elseif Request("txttable") ="Calculate_Mill_Product" then%>
    <b>Calculate Mill Product & Delivery</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" >
    
<%Elseif Request("txttable") ="WeighBridge_Closing" then%>
    <b>WeighBridge Closing </b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_WeighBridge")) &"/"& Year(Session("v_WeighBridge"))%>
  
<%Elseif Request("txttable") ="Production_Closing" then%>
    <b>Production Closing </b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Mill_Production")) &"/"& Year(Session("v_Mill_Production"))%>
    
<%Elseif Request("txttable") ="Delivery_Closing" then%>
    <b>Delivery Closing</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_Delivery")) &"/"& Year(Session("v_Delivery"))%>	
<%Elseif Request("txttable") ="Plasma_Reconciliation" then%>
    <b>Plasma Reconciliation Closing</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
	Periode : <%=Month(Session("v_PlasmaReconsil")) &"/"& Year(Session("v_PlasmaReconsil"))%>	
<%Elseif Request("txttable") ="Sales_Closing_New" then%>
    <b>Sales Closing</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="Process" ><br>
<%Elseif Request("txttable") ="Clone" then%>
    <b>Clone Data Master to All Business Unit</b><br><br>
	<img src="ICON/database_process.png" border="0" width="80" height="80"><br><br>
	<input type="submit" value="::Clone::" >
<%elseif Request("txttable") ="UpkeepRotation" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Upkeep Rotation</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Year</font></td>
			<td>
				<select name="txtYear">
				<%
					Yr = Year(Session("v_Costbook"))
					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(Request("txtYear")) = cint(xYr) then
						response.write "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						response.write "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Month</font></td>
			<td>
				<select name="txtMonth">
					<option value="">--All--</option>
				<%
					
					for xMn=1 to 12
					if cint(Request("txtMonth")) = cint(xMn) then
						response.write "<option value ='"&xMn&"' selected>"&MonthName(xMn)&"</option>"
					else
						response.write "<option value ='"&xMn&"'>"&MonthName(xMn)&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Jobcode</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="jobcode">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Fieldcode</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="fieldcode">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="MonthlyWorkPlan" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Monthly Work Plan</span>
		</a>
		<br><br><br>
		<!--<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Year</font></td>
			<td>
				<select name="txtYear">
				<%
					Yr = Year(Session("v_Costbook"))
					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(Request("txtYear")) = cint(xYr) then
						response.write "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						response.write "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Month</font></td>
			<td>
				<select name="txtMonth">
					<option value="">--All--</option>
				<%
					
					for xMn=1 to 12
					if cint(Request("txtMonth")) = cint(xMn) then
						response.write "<option value ='"&xMn&"' selected>"&MonthName(xMn)&"</option>"
					else
						response.write "<option value ='"&xMn&"'>"&MonthName(xMn)&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Jobcode</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="jobcode">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Fieldcode</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="fieldcode">
			</td>
		</tr>-->
		<tr>
			<!-- <td colspan="2" align="left"><input type="submit" value="Preview" ></td> -->
			<td colspan="2" align="left"><a href="report/p_rkb_rawat.asp" target="_blank"><input type="button" value="Preview" id="mh"></a></td>
		</tr>
		</table>
<%elseif Request("txttable") ="MonthlyWorkPlanHarv" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Monthly Work Plan Harvesting</span>
		</a>
		<br><br><br>
		<!-- <font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Year</font></td>
			<td>
				<select name="txtYear">
				<%
					Yr = Year(Session("v_Costbook"))
					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(Request("txtYear")) = cint(xYr) then
						response.write "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						response.write "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Month</font></td>
			<td>
				<select name="txtMonth">
					<option value="">--All--</option>
				<%
					
					for xMn=1 to 12
					if cint(Request("txtMonth")) = cint(xMn) then
						response.write "<option value ='"&xMn&"' selected>"&MonthName(xMn)&"</option>"
					else
						response.write "<option value ='"&xMn&"'>"&MonthName(xMn)&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Jobcode</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="jobcode">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Fieldcode</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="fieldcode">
			</td>
		</tr> -->
		<tr>
			<!-- <td colspan="2" align="left"><input type="submit" value="Preview" ></td> -->
			<td colspan="2" align="left"><a href="report/p_rkb_panen.asp" target="_blank"><input type="button" value="Preview" id="mh"></a></td>
		</tr>
		</table>
<%elseif Request("txttable") ="bbctransaction" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance BBC Transaction</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Year</font></td>
			<td>
				<select name="txtYear">
				<%
					Yr = Year(Session("v_Costbook"))
					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(Request("txtYear")) = cint(xYr) then
						response.write "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						response.write "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Census Code</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="censuscode">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="bbcdistribution" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance BBC Distribution</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Year</font></td>
			<td>
				<select name="txtYear">
				<%
					Yr = Year(Session("v_Costbook"))
					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(Request("txtYear")) = cint(xYr) then
						response.write "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						response.write "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="supplyingdeclaration" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Supplying Declaration</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Date</font></td>

			<td>
				<input class="datepicker" type="text" maxlength="10" name="date_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="date_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >

			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Document No</font></td>
			<td><input class="code" type="text" size="15" name="documentno" ></td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Type</font></td>
			<td>
				<select name="declarationtype">
					<Option Value=""></option>
					<%

					strsql = " SELECT fccode, fcname FROM parameter WHERE parcode = 'AGROTECH_SUPPLYINGTYPE' AND fcba = '" & session("fcBA") & "' ORDER BY fccode"
					'response.Write strsql
					Set rsparam = Server.CreateObject("ADODB.Recordset")
					rsparam.Open strsql, connstr
					
					if not rsparam.eof then
						while not rsparam.eof
					%>
						<Option Value="<%=rsparam("fccode")%>" <%If cstr(""&request("declarationtype")) = cstr(""&rsparam("fccode")) then%> Selected <%End If%>><%=rsparam("fcname")%></option>
					<%
							rsparam.movenext
						wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Division</font></td>
			<td>
				<select name="division">
					<Option Value=""></option>
					<%

					strsql = " SELECT fccode, fcname FROM parameter WHERE parcode = 'F0001' AND fcba = '" & session("fcBA") & "' ORDER BY fccode"
					'response.Write strsql
					Set rsparam = Server.CreateObject("ADODB.Recordset")
					rsparam.Open strsql, connstr
					
					if not rsparam.eof then
						while not rsparam.eof
					%>
						<Option Value="<%=rsparam("fccode")%>" <%If cstr(""&request("division")) = cstr(""&rsparam("fccode")) then%> Selected <%End If%>><%=rsparam("fcname")%></option>
					<%
							rsparam.movenext
						wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Fieldcode</font></td>
			<td><input class="code" type="text" size="15" name="fieldcode" ></td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="plantingdeclaration" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Planting Declaration</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Date</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="TRANSDATE_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="TRANSDATE_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Document No</font></td>
			<td><input class="code" type="text" size="15" name="documentno" ></td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Fieldcode</font></td>
			<td><input class="code" type="text" size="15" name="fieldcode" ></td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Fullstand</font></td>
			<td>
				<select name="fullstand">
					<Option Value=""> -- Choose Select -- </option>
					<Option Value="TRUE"> TRUE </option>
					<Option Value="FALSE"> FALSE </option>
				</select>
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="pemeriksaan_mutubuah" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Pemeriksaan Mutu Buah</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Date</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="transdate_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="transdate_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Division</font></td>
			<td>
				<select name="division">
					<%
					strsql = " SELECT fccode, fcname FROM parameter WHERE parcode = 'F0001' AND fcba = '" & session("fcBA") & "' ORDER BY fccode"
					'response.Write strsql
					Set rsparam = Server.CreateObject("ADODB.Recordset")
					rsparam.Open strsql, connstr
					
					if not rsparam.eof then
						while not rsparam.eof
					%>
						<Option Value="<%=rsparam("fccode")%>" <%If cstr(""&division) = cstr(""&rsparam("fccode")) then%> Selected <%End If%>><%=rsparam("fcname")%></option>
					<%
							rsparam.movenext
						wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="pemeriksaan_mutuancak" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Pemeriksaan Mutu Ancak</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Date</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="transdate_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="transdate_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Division</font></td>
			<td>
				<select name="division">
					<%
					strsql = " SELECT fccode, fcname FROM parameter WHERE parcode = 'F0001' AND fcba = '" & session("fcBA") & "' ORDER BY fccode"
					'response.Write strsql
					Set rsparam = Server.CreateObject("ADODB.Recordset")
					rsparam.Open strsql, connstr
					
					if not rsparam.eof then
						while not rsparam.eof
					%>
						<Option Value="<%=rsparam("fccode")%>" <%If cstr(""&division) = cstr(""&rsparam("fccode")) then%> Selected <%End If%>><%=rsparam("fcname")%></option>
					<%
							rsparam.movenext
						wend
					end if
					%>
				</select>
			</td>
		</tr>

		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="sn_finger" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Serial Number</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">Serial  Number</font></td>
			<td>
				<input type="text" size="15" maxlength="30" name="serial_number">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="Fingerprint_BA" then%>
	<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/Payslip.jpg" border="0" >
	<span>Maintenance BA Fingerprint</span></a>
    <br><br>
	<script>
		function checkAll() {
			var checkbox_options = document.getElementById("checkbox_options");
			var checkbox_all = document.getElementById("checkbox_all");

			var checkboxes = checkbox_options.getElementsByTagName("input");        
			for (i = 0; i < checkboxes.length; i++) {
				checkboxes[i].checked = checkbox_all.checked;
			}
		}
	</script>
	<Table>
		<tr>
			<Td colspan="3">
				<font face="Verdana" size="1">Select Field :</Font><br>
				<div class="ckbox_list">
					<table>
						<tr>
							<td><input id="checkbox_all" type="checkbox" onclick="checkAll()" /><span><font face="Verdana" size="1">Select All</Font></span></td>
						</tr>
					</table>
					<div id = "checkbox_options" >
						<table>
							<tr>
								<td><input name="cField" type="checkbox" value="adhocdate" /><span><font face="Verdana" size="1">Date</font></span></td>
								<td><input name="cField" type="checkbox" value="employeecode" /><span><font face="Verdana" size="1">Employee Code</Font></span></td>
								<td><input name="cField" type="checkbox" value="employeename"/><span><font face="Verdana" size="1">Employee Name</Font> </span></td>
								<td><input name="cField" type="checkbox" value="gangcode" /><span><font face="Verdana" size="1">Gang Code</Font></span></td>
							</tr>
						</table>
					</div>
				</div>
			</td>
		</tr>
	<tr>
	<td><br><font face="Verdana" size="1">Select Data Here :</font><br> 
    </Td>
    </tr>
    <tr>
	<td>
    <font face="Verdana" size="1">From</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
	<td>
	<font face="Verdana" size="1">To</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</Td>
    </tr>
    <tr>
	<td>
    <font face="Verdana" size="1">Gang</font>
    </td>
    <td>
	<select name="GANGCODE" class="chosen-select">
		<%cmdCallCombo "Gang","","",""%>
	</select>
    </td>
	</tr>
	<tr>
	<td>
    <font face="Verdana" size="1">Employee</font>
    </td>
    <td>
	<select name="EMPLOYEECODE" class="chosen-select">
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select>
	
	</td>
	</tr>
	</Table>
    <br />
    <input type="submit" value="Preview" >		
<%elseif Request("txttable") ="salesagreement" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Sales Contract</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsComboPV
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboPV = Server.CreateObject("ADODB.Recordset")
						rsComboPV.CursorType = 3
						rsComboPV.CursorLocation = 3
						rsComboPV.Open q, connstr
						if not rsComboPV.eof then%>
						<%
						Do While Not rsComboPV.eof
						%>
							<option value="<%=rsComboPV("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboPV("fccode")) then %> selected 
								<%End If%>><%=rsComboPV("fccode")%>
							</option> 
						<%rsComboPV.movenext
						loop
						End if
						rsComboPV.Close
						set rsComboPV = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Sales Contract No</font></td>
			<td>
				<input type="text" size="30" name="salescontractno">
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Customer</font></TD>
			<td>
				<select name="customercode">
					<option value=""></option>
	    	<% 
			q = " SELECT customercode fccode, description fcname FROM customer WHERE supplier = 'N' AND fcba = '" & Session("fcBA") & "' ORDER BY customercode"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.Open q, connstr
			if not rsContract.eof then
				Do While Not rsContract.eoF%>
					<option value="<%=rsContract("fccode")%>"><%=rsContract("fccode") & ": " & rsContract("fcname") %></option> 
			<%rsContract.movenext
				loop
			End if
			rsContract.Close : set rsContract = Nothing : q = ""
			%>
				</select>
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Product</font></TD>
			<td>
				<select name="productcode">
					<option value=""></option>
	    	<% 
			q = " SELECT productcode fccode, productdescription fcname FROM product WHERE fcba = '" & Session("fcBA") & "' ORDER BY productcode"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.Open q, connstr
			if not rsContract.eof then
				Do While Not rsContract.eoF%>
					<option value="<%=rsContract("fccode")%>"><%=rsContract("fccode") & ": " & rsContract("fcname") %></option> 
			<%rsContract.movenext
				loop
			End if
			rsContract.Close : set rsContract = Nothing : q = ""
			%>
				</select>
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Delivery Type</font></TD>
			<td>
				<select name="deliverytype">
					<option value=""></option>
	    	<% 
			q = " SELECT fccode, fcname FROM parameter WHERE fcba = '" & Session("fcBA") & "' AND parcode = 'DELIVERY_TYPE' ORDER BY fccode"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.Open q, connstr
			if not rsContract.eof then
				Do While Not rsContract.eoF%>
					<option value="<%=rsContract("fccode")%>"><%=rsContract("fccode") & ": " & rsContract("fcname") %></option> 
			<%rsContract.movenext
				loop
			End if
			rsContract.Close : set rsContract = Nothing : q = ""
			%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="salesinvoicedp" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Sales Invoice DP</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboPV = Server.CreateObject("ADODB.Recordset")
						rsComboPV.CursorType = 3
						rsComboPV.CursorLocation = 3
						rsComboPV.Open q, connstr
						if not rsComboPV.eof then%>
						<%
						Do While Not rsComboPV.eof
						%>
							<option value="<%=rsComboPV("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboPV("fccode")) then %> selected 
								<%End If%>><%=rsComboPV("fccode")%>
							</option> 
						<%rsComboPV.movenext
						loop
						End if
						rsComboPV.Close
						set rsComboPV = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Sales Invoice DP No</font></td>
			<td>
				<input type="text" size="30" name="invoicedpno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Sales Contract No</font></td>
			<td>
				<input type="text" size="30" name="salescontractno">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="transportercontract" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Transporter Contract</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Contract No</font></td>
			<td>
				<input type="text" size="30" name="transportercontractno">
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Transporter</font></TD>
			<td>
				<select name="transporter">
					<option value=""></option>
	    	<% 
			q = " SELECT fccode, fcname FROM transporter WHERE fcba = '" & Session("fcBA") & "' ORDER BY fccode"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.Open q, connstr
			if not rsContract.eof then
				Do While Not rsContract.eoF%>
					<option value="<%=rsContract("fccode")%>"><%=rsContract("fccode") & ": " & rsContract("fcname") %></option> 
			<%rsContract.movenext
				loop
			End if
			rsContract.Close : set rsContract = Nothing : q = ""
			%>
				</select>
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="deliveryorder" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Delivery Order</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1"> From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="deliveryorderdate_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1"> To</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="deliveryorderdate_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Delivery Order</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="deliveryorderno">
			</td>
		</tr>	
		<tr>
			<td><font face="Verdana" size="1">Sales Contract</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salescontractno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Transporter Contract</font></td>
			<td>
				<select name="transportercontractno" id="transportercontractno"  title="TRANSPORTER CODE Character a s/d z or 1 s/d 9" onblur="javascript:getTransporter();">
					<option value=""></option> 
						<% 	Dim rsComboP
							q = "SELECT * FROM TRANSPORTERCONTRACT WHERE fcba = '" & Session("fcba") & "'"
							Set rsComboP = Server.CreateObject("ADODB.Recordset")
							rsComboP.CursorType = 3
							rsComboP.CursorLocation = 3
							rsComboP.Open q, connstr
							if not rsComboP.eof then%>
							<%
							Do While Not rsComboP.eof
							%>
								<option value="<%=rsComboP("TRANSPORTERCONTRACTNO")%>" 
									<%if cstr(""&transportercontractno)=cstr(""&rsComboP("TRANSPORTERCONTRACTNO")) then %> selected 
									<%End If%>><%=rsComboP("TRANSPORTERCONTRACTNO")%>
								</option> 
							<%rsComboP.movenext
							loop
							End if
							rsComboP.Close
							set rsComboP = Nothing
							q=""
						%>
				</select> 
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="salesshipment" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Sales Shipment</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Sales Shipment</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salesshipmentno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Delivery Order</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="deliveryorderno">
			</td>
		</tr>	
		<tr>
			<td><font face="Verdana" size="1">Sales Contract</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salescontractno">
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="salesinvoice" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Sales Invoice</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Sales Invoice</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="invoicecode">
			</td>
		</tr>	
		<tr>
			<td><font face="Verdana" size="1">Sales Contract</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salescontractno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Shipment</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salesshipmentno">
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Customer</font></TD>
			<td>
				<select name="customercode">
					<option value=""></option>
	    	<% 
			q = " SELECT customercode fccode, description fcname FROM customer WHERE supplier = 'N' AND fcba = '" & Session("fcBA") & "' ORDER BY customercode"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.Open q, connstr
			if not rsContract.eof then
				Do While Not rsContract.eoF%>
					<option value="<%=rsContract("fccode")%>"><%=rsContract("fccode") & ": " & rsContract("fcname") %></option> 
			<%rsContract.movenext
				loop
			End if
			rsContract.Close : set rsContract = Nothing : q = ""
			%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Product Code</font>
		    </td>
			<td>
		    	<select name="ProductCODE">
					<%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%>
				</select>
			</td>
	    </tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="purchaseagreement" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Purchase Contract</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	Dim rsBUPA
						q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsBUPA = Server.CreateObject("ADODB.Recordset")
						rsBUPA.CursorType = 3
						rsBUPA.CursorLocation = 3
						rsBUPA.Open q, connstr
						if not rsBUPA.eof then%>
						<%
						Do While Not rsBUPA.eof
						%>
							<option value="<%=rsBUPA("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsBUPA("fccode")) then %> selected 
								<%End If%>><%=rsBUPA("fccode")%>
							</option> 
						<%rsBUPA.movenext
						loop
						End if
						rsBUPA.Close
						set rsBUPA = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Purchase Contract No</font></td>
			<td>
				<input type="text" size="30" name="salescontractno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Supplier</font></td>
			<td>
		 	<select name="SUPPLIERCODE" class="chosen-select">
					<%cmdCallCombo "SUPPLIER","","",""%>
			</select>
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Product</font></TD>
			<td>
				<select name="productcode">
					<option value=""></option>
	    	<% 
			q = " SELECT productcode fccode, productdescription fcname FROM product WHERE fcba = '" & Session("fcBA") & "' ORDER BY productcode"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.Open q, connstr
			if not rsContract.eof then
				Do While Not rsContract.eoF%>
					<option value="<%=rsContract("fccode")%>"><%=rsContract("fccode") & ": " & rsContract("fcname") %></option> 
			<%rsContract.movenext
				loop
			End if
			rsContract.Close : set rsContract = Nothing : q = ""
			%>
				</select>
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Delivery Type</font></TD>
			<td>
				<select name="deliverytype">
					<option value=""></option>
	    	<% 
			q = " SELECT fccode, fcname FROM parameter WHERE fcba = '" & Session("fcBA") & "' AND parcode = 'DELIVERY_TYPE' ORDER BY fccode"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.Open q, connstr
			if not rsContract.eof then
				Do While Not rsContract.eoF%>
					<option value="<%=rsContract("fccode")%>"><%=rsContract("fccode") & ": " & rsContract("fcname") %></option> 
			<%rsContract.movenext
				loop
			End if
			rsContract.Close : set rsContract = Nothing : q = ""
			%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="purchasedeliveryorder" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Purchase Delivery Order</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1"> From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="deliveryorderdate_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1"> To</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="deliveryorderdate_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Purchase Delivery Order</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="deliveryorderno">
			</td>
		</tr>	
		<tr>
			<td><font face="Verdana" size="1">Purchase Contract</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salescontractno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Transporter Contract</font></td>
			<td>
				<select name="transportercontractno" id="transportercontractno"  title="TRANSPORTER CODE Character a s/d z or 1 s/d 9" onblur="javascript:getTransporter();">
					<option value=""></option> 
						<% 	'Dim rsComboP
							q = "SELECT * FROM TRANSPORTERCONTRACT WHERE fcba = '" & Session("fcba") & "'"
							Set rsComboP = Server.CreateObject("ADODB.Recordset")
							rsComboP.CursorType = 3
							rsComboP.CursorLocation = 3
							rsComboP.Open q, connstr
							if not rsComboP.eof then%>
							<%
							Do While Not rsComboP.eof
							%>
								<option value="<%=rsComboP("TRANSPORTERCONTRACTNO")%>" 
									<%if cstr(""&transportercontractno)=cstr(""&rsComboP("TRANSPORTERCONTRACTNO")) then %> selected 
									<%End If%>><%=rsComboP("TRANSPORTERCONTRACTNO")%>
								</option> 
							<%rsComboP.movenext
							loop
							End if
							rsComboP.Close
							set rsComboP = Nothing
							q=""
						%>
				</select> 
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>	
<%elseif Request("txttable") ="purchaseshipment" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Purchase Shipment</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="creationdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Purchase Shipment</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salesshipmentno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Delivery Order</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="deliveryorderno">
			</td>
		</tr>	
		<tr>
			<td><font face="Verdana" size="1">Purchase Contract</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salescontractno">
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>
<%elseif Request("txttable") ="purchaseinvoicedp" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Purchase Invoice DP</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td align="left">
				<select name="businessunit" >
					<option value=""></option> 
					<% 	q = " SELECT DISTINCT fccode " & _
						 	" FROM (SELECT businessunit.fccode " & _
						    " FROM user_fcba_tab " & _
						    " INNER JOIN businessunit " & _
						    " ON user_fcba_tab.fcba = " & _
						    " businessunit.fccode " & _
						    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
						' response.write q
						Set rsComboPV = Server.CreateObject("ADODB.Recordset")
						rsComboPV.CursorType = 3
						rsComboPV.CursorLocation = 3
						rsComboPV.Open q, connstr
						if not rsComboPV.eof then%>
						<%
						Do While Not rsComboPV.eof
						%>
							<option value="<%=rsComboPV("fccode")%>" 
								<%if cstr(""&businessunit)=cstr(""&rsComboPV("fccode")) then %> selected 
								<%End If%>><%=rsComboPV("fccode")%>
							</option> 
						<%rsComboPV.movenext
						loop
						End if
						rsComboPV.Close
						set rsComboPV = Nothing
						q=""
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Purchase Invoice DP No</font></td>
			<td>
				<input type="text" size="30" name="invoicedpno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Purchase Contract No</font></td>
			<td>
				<input type="text" size="30" name="salescontractno">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
		</table>
<%elseif Request("txttable") ="purchaseinvoice" then%>
		<a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>" title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			<span>Maintenance Purchase Invoice</span>
		</a>
		<br><br><br>
		<font face="Verdana" size="1">Select Data Here : <br><br>
		<table>
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_from" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td><font face="Verdana" size="1">To</font></td>
			<td><input class="datepicker" type="text" maxlength="10" name="transdate_to" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Business Unit</font>
			</td>
			<td>
				<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	q = " SELECT DISTINCT fccode " & _
							 	" FROM (SELECT businessunit.fccode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
							' response.write q
							Set rsComboBST = Server.CreateObject("ADODB.Recordset")
							rsComboBST.CursorType = 3
							rsComboBST.CursorLocation = 3
							rsComboBST.Open q, connstr
							if not rsComboBST.eof then%>
							<%
							Do While Not rsComboBST.eof
							%>
								<option value="<%=rsComboBST("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboBST("fccode")) then %> selected 
									<%End If%>><%=rsComboBST("fccode")%>
								</option> 
							<%rsComboBST.movenext
							loop
							End if
							rsComboBST.Close
							set rsComboBST = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Purchase Invoice</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="invoicecode">
			</td>
		</tr>	
		<tr>
			<td><font face="Verdana" size="1">Purchase Contract</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salescontractno">
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Shipment</font></td>
			<td>
				<input type="text" size="25" maxlength="30" name="salesshipmentno">
			</td>
		</tr>
		<tr>
		 	<td><font face="Verdana" size="1">Supplier</font></TD>
			<td>
			<select name="customercode" class="chosen-select">
					<%cmdCallCombo "SUPPLIER","","",""%>
			</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Product Code</font>
		    </td>
			<td>
		    	<select name="ProductCODE">
					<%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%>
				</select>
			</td>
	    </tr>
		<tr>
			<td colspan="2" align="left"><input type="submit" value="Preview" ></td>
		</tr>
	</table>		
<%else%>
    <a class="button_menu_top" href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
    <span>Maintenance Data</span>
    </a><br><br>
    <table>
    <!-- <tr>
    <td>
	<font face="verdana" size="1">Search : </font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5">
    </td></tr>
    <tr>
    <td>
    <font face="verdana" size="1">&nbsp;from Column</font>
	</td> -->
	<td>
	<%if Request("txttable") ="Bank" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control Job</Option>
		<Option value="ACCOUNTCODE">Account Code</Option>
		<Option value="ADDRESS">Address</Option>
		<Option value="CITY">City</Option>
		<Option value="STATE">State</Option>
		</select>
	<%Elseif Request("txttable") ="Stock" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ITEMTYPE">Item Type</Option>
		<Option value="MANUFACTUREPARTNUMBER">Manufacture Part Number</Option>
		<Option value="MANUFACTURECODE">Manufacture Code</Option>
		<Option value="SUPPLIERPARTNUMBER">Supplier Part Number</Option>
		<Option value="SUPPLIERCODE">Supplier Code</Option>
		<Option value="UNITOFMEASUREMENT">UOM</Option>
		<Option value="STOCKGROUPCODE">Stock<br>Group</Option>
		</select>
	<%Elseif Request("txttable") ="StockGroup" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>
	<%Elseif Request("txttable") ="Supplier" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="ADDRESS">Address</Option>
		<Option value="CITY">City</Option>
		<Option value="STATE">State</Option>
		</select>
	<%Elseif Request("txttable") ="Store" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
	<%Elseif Request("txttable") ="Gang" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
	<%Elseif Request("txttable") ="Workshop" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
	<%Elseif Request("txttable") ="Vehicle" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="VEHICLEGROUPCODE">Vehicle<br>Group Code</Option>
		<Option value="REGISTRATIONNO">Registration<br>No</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="OWNERSHIP">Ownership</Option>
		<Option value="VEHICLESTATUS">Vehicle<br>Status</Option>
		<Option value="CONTROLALLOCATION">Control<br>Allocation</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>
	<%Elseif Request("txttable") ="Contractor" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="ADDRESS">Address</Option>
		<Option value="CITY">City</Option>
		<Option value="STATE">State</Option>
		</select>
	<%Elseif Request("txttable") ="Building" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="BUILDGROUPCODE_DESC">Building<br>Group Code</Option>
		<Option value="ITEM">Item</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
	<%Elseif Request("txttable") ="Allowded" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ALLDED_GROUP">Allowded<br>Group</Option>
		<Option value="ALLDED">Allowded</Option>
		<Option value="CROP">Crop</Option>
		<Option value="CONTROL_JOB">Control<br>Account</Option>
		<Option value="MAN_AUTO">MANUAL<br>Auto</Option>
		<Option value="TYPE_1">TYPE<br>1</Option>
		<Option value="TYPE_2">TYPE<br>2</Option>
		<Option value="TYPE_3">TYPE<br>3</Option>
		<Option value="TYPE_4">TYPE<br>4</Option>
		<Option value="MIN_AMT">MINIMAL<br>AMT</Option>
		<Option value="MAX_AMT">MAKSIMAL<br>AMT</Option>
		<Option value="RATE_1">RATE<br>1</Option>
		<Option value="RATE_2">RATE<br>2</Option>
		<Option value="RATE_3">RATE<br>3</Option>
		<Option value="RATE_4">RATE<br>4</Option>
		
		</select>
	<%Elseif Request("txttable") ="Field" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="DIVISION">Division</Option>
		<Option value="TERRAINTYPE">Terrain<br>Type</Option>
		<Option value="SOILTYPE">Soil Type</Option>		
		<Option value="OWNERSHIP">Ownership</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="CROPTYPE">Crop<br>Type</Option>
		</select>
	<%Elseif Request("txttable") ="Station" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="JOBCODE">Control<br>Account</Option>
		</select>
	<%Elseif Request("txttable") ="Machine" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="DESCRIPTION">Description</Option>
		<Option value="FIXEDASSETCODE">Fixex<br>Asset</Option>
		<Option value="OWNERSHIP">Ownership</Option>
		<Option value="STATION_STATIONCODE">Station<br>Code</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>		
	<%Elseif Request("txttable") ="BusinessUnit" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="FCCOMPANYCODE">Company Code</Option>
		<Option value="FCCOMPANYNAME">Company Description</Option>
		<Option value="FCAREA">Area</Option>
		<Option value="FCOWNERSHIP">Ownership</Option>
		<Option value="FCMANAGER">Manager</Option>
		<Option value="FCTELEPHONE">Phone</Option>
		<Option value="FCFAX">Fax</Option>
		<Option value="FCEMAIL">e-Mail</Option>
		</select>		
	<%Elseif Request("txttable") ="UserProfile" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FULLNAME">Description</Option>
		<Option value="DEPARTMENT">Department</Option>
		<Option value="ACCESSRIGHTS">Access Right</Option>
		</select>	
	<%Elseif Request("txttable") ="Users" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">User Id</Option>
		<Option value="FCPASSWORD">Password</Option>
		<Option value="FULLNAME">Full Name</Option>
		<Option value="DEPARTMENT">Department</Option>
		<Option value="ACCESSRIGHTS">Access Rights</Option>
		<Option value="ISACTIVE">Active</Option>
		<Option value="VALID_FROM">Valid From</Option>
		<Option value="VALID_UNTIL">Valid Until</Option>
		</select>						
	<%Elseif Request("txttable") ="AccessRights" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">User</Option>
		<Option value="MODULECODE">Module</Option>
		<Option value="SUBMODULECODE">Sub Module</Option>
		<Option value="AUTHORIZED">Authorized</Option>
		</select>		
	<%Elseif Request("txttable") ="Trigger_Control" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="ACTIVITY">Activity</Option>
		<Option value="SOURCE">Source Type</Option>
		<Option value="TARGET">Target Type</Option>
		</select>		
	<%Elseif Request("txttable") ="Job" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="JOBGROUPCODE">Job Group<br>Code</Option>
		<Option value="JOB_FORCROPTYPE">Job For Type</Option>
		<Option value="JOB_FORFIELDSTATUS">Job For Status</Option>
		<Option value="JOB_CATEGORY">Job Category</Option>
		<Option value="UNITOFMEASUREMENT">Unit of Measurement</Option>
		<Option value="ACCOUNTCODE">Account<br>Code</Option>
		</select>		
	</select>		
	<%Elseif Request("txttable") ="JobGroup" then%>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="PARENTCODE">Parent<br>Code</Option>
		<Option value="PARENTDESC">Parent<br>Description</Option>
		</select>
	<%Else%>
		<!-- <input type="text" title="Case is not Sensitive" name="txtField" value="<%=txtField%>" width="5">	 -->
	<%End if%>
	<!-- &nbsp;&nbsp;
    </td>
    </tr>
    <tr>
    <td rowspan="2">
	<input type=submit value="Preview" >
    </td>
    </tr> -->
    </table>
<%end if%>
<script>
	

	function getFCBA()
	{
		var company = $("#company").val();
		showFCBA(company);

	}

	function showFCBA(company) {
				$('#businessunit').empty();
				$('#businessunit').select2({
					ajax: {
						delay: 1000,
						url : 'getListfcbaallajax.asp?fccompanycode=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}

    function getFCBAgl()
	{
		var company = $("#companygl").val();
		showFCBAgl(company);

	}

	function showFCBAgl(company) {
				$('#businessunitgl').empty();
				$('#businessunitgl').select2({
					ajax: {
						delay: 1000,
						url : 'getListfcbaallajax.asp?fccompanycode=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}
	
	function getFCBAbs()
	{
		var company = $("#companybs").val();
		showFCBAbs(company);

	}

	function showFCBAbs(company) {
				$('#businessunitbs').empty();
				$('#businessunitbs').select2({
					ajax: {
						delay: 1000,
						url : 'getListfcbaallajax.asp?fccompanycode=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}
	
	function getFCBArkb()
	{
		var company = $("#companyrkb").val();
		showFCBArkb(company);

	}

	function showFCBArkb(company) {
				$('#businessunitrkb').empty();
				$('#businessunitrkb').select2({
					ajax: {
						delay: 1000,
						url : 'getListfcbaallajax.asp?fccompanycode=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}
	
	function getFCBAmam()
	{
		var company = $("#companymam").val();
		showFCBAmam(company);

	}
	
	function showFCBAmam(company) {
				$('#businessunitmam').empty();
				$('#businessunitmam').select2({
					ajax: {
						delay: 1000,
						url : 'getListfcbaallajax.asp?fccompanycode=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}

    function getFCBAbnk()
	{
		var company = $("#companybnk").val();
		showFCBAbnk(company);

	}
	
	function showFCBAbnk(company) {
				$('#businessunitbnk').empty();
				$('#businessunitbnk').select2({
					ajax: {
						delay: 1000,
						url : 'getListfcbaajax.asp?fccompanycode=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}
	
	function getBANKbnk()
	{
		var bnk = $("#businessunitbnk").val();
		showBANKbnk(bnk);

	}
	
	function showBANKbnk(company) {
				$('#bankbnk').empty();
				$('#bankbnk').select2({
					ajax: {
						delay: 1000,
						url : 'getListbankallajax.asp?businessunit=' + company,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}
	
</script>