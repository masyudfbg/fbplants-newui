<%
Function LastDayOfMonth (ByVal CheckDate, ReturnType)
		' return type 1 = Day, 2 = Date
		Dim LastDay
		Dim LastDayDate
		If  ReturnType = 1 Then
			LastDay = Day(Dateadd("m",1,Dateadd("d", -day(CheckDate),CheckDate)))
			LastDayOfMonth = LastDay
		Else
			LastDayDate = Dateadd("m",1,Dateadd("d", -day(CheckDate),CheckDate))
			LastDayOfMonth = LastDayDate 
		End If
End Function

Function set_format_tanggal(inittanggal,initformat) 
                Dim initialformat_tanggal
                Dim rsformat_tanggal, strformat_tanggal
                
                strformat_tanggal =        " Select to_char(to_date('" & inittanggal & "','MM/DD/YYYY'),'" & initformat & "') formattgl from dual"
                                                                                                                
                Set rsformat_tanggal = Server.CreateObject("ADODB.Recordset")
                rsformat_tanggal.Open strformat_tanggal, connstr, 3
                
                If Not rsformat_tanggal.EOF Then
                                initialformat_tanggal = cstr(rsformat_tanggal("formattgl"))
                Else
                                initialformat_tanggal = inittanggal
                End If
                
                rsformat_tanggal.Close
                set rsformat_tanggal = Nothing
                set_format_tanggal = initialformat_tanggal
end function



Sub cmdCallCombo(txtRequestTable,txtRequestFields,txtFieldWhere,txtValueWhere)
 	Dim rsCallCombo
    if txtFieldWhere = "" then
   	   q = "select * from " & txtRequestTable & " where fcBA='" &  Session("fcBA") & "' Order by FCCODE"
	else
   	   q = "select * from " & txtRequestTable & " Where " & txtFieldWhere & " = '" & txtValueWhere & "' and fcBA='" & session("fcBA") & "' Order by fcName"
	End if
	'response.write 
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.CursorType = 3
	rsCallCombo.CursorLocation = 3
	rsCallCombo.Open q, connstr
	if not rsCallCombo.eof then
	%>
	<option value=""></option> 
	<%
	 Do While Not rsCallCombo.eof
	%>
	<option value="<%=rsCallCombo("fcCode")%>" <%if txtRequestFields = rsCallCombo("fcCode") then%> Selected <%End if%>><%=rsCallCombo("fcCode")%> : <%=rsCallCombo("fcName")%></option>
		<%rsCallCombo.movenext
	 loop
	End if
	rsCallCombo.Close
	set rsCallCombo = Nothing
	q=""
End Sub


Sub cmdCallComboWhereStatement(txtRequestTable,txtWhereStatement)
 	Dim rsCallCombo
    q = "select * from " & txtRequestTable & " Where " & txtWhereStatement & "' Order by fcName"

	'response.write 
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.CursorType = 3
	rsCallCombo.CursorLocation = 3
	rsCallCombo.Open q, connstr
	if not rsCallCombo.eof then
	%>
	<option value=""></option> 
	<%
	 Do While Not rsCallCombo.eof
	%>
	<option value="<%=rsCallCombo("fcCode")%>" <%if txtRequestFields = rsCallCombo("fcCode") then%> Selected <%End if%>><%=rsCallCombo("fcCode")%> : <%=rsCallCombo("fcName")%></option>
		<%rsCallCombo.movenext
	 loop
	End if
	rsCallCombo.Close
	set rsCallCombo = Nothing
	q=""
End Sub




Sub cmdCallComboSetup(txtRequestTable,txtRequestFields,txtFieldWhere,txtValueWhere)
 	Dim rsCallCombo
    if txtFieldWhere = "" then
   	   q = "select * from " & txtRequestTable
	else
   	   q = "select * from " & txtRequestTable & " Where " & txtFieldWhere & " = '" & txtValueWhere & "'"
	End if
	response.write q
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.CursorType = 3
	rsCallCombo.CursorLocation = 3
	rsCallCombo.Open q, connstr
	if not rsCallCombo.eof then%>
	 <option value=""></option><%
	 Do While Not rsCallCombo.eof
        if txtRequestTable="LOCATIONTYPE" then
		     if rsCallCombo("lOCATIONTYPECODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("lOCATIONTYPECODE")%>" Selected><%=rsCallCombo("lOCATIONTYPECODE")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("lOCATIONTYPECODE")%>"><%=rsCallCombo("lOCATIONTYPECODE")%></option>  
		    <%End if
        elseif txtRequestTable="VEHICLEGROUP" then
		     if rsCallCombo("VEHICLEGROUPCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("VEHICLEGROUPCODE")%>" Selected><%=rsCallCombo("VEHICLEGROUPCODE")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("VEHICLEGROUPCODE")%>"><%=rsCallCombo("VEHICLEGROUPCODE")%></option>  
		    <%End if
        elseif txtRequestTable="Crop" then
		     if rsCallCombo("CROPCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("CROPCODE")%>" Selected><%=rsCallCombo("DESCRIPTION")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("CROPCODE")%>"><%=rsCallCombo("DESCRIPTION")%></option>  
		    <%End if
        elseif txtRequestTable="PRODUCTSTORAGE" then
		     if rsCallCombo("STORAGECODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("STORAGECODE")%>" Selected><%=rsCallCombo("STORAGECODE")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("STORAGECODE")%>"><%=rsCallCombo("STORAGECODE")%></option>  
		    <%End if
        elseif txtRequestTable="DELIVERYNOTE" then
		     if rsCallCombo("DELIVERYNOTECODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("DELIVERYNOTECODE")%>" Selected><%=rsCallCombo("DELIVERYNOTECODE")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("DELIVERYNOTECODE")%>"><%=rsCallCombo("DELIVERYNOTECODE")%></option>  
		    <%End if
        elseif txtRequestTable="DELIVERYCONTRACT" then
		     if rsCallCombo("CONTRACTCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("CONTRACTCODE")%>" Selected><%=rsCallCombo("CONTRACTCODE")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("CONTRACTCODE")%>"><%=rsCallCombo("CONTRACTCODE")%></option>  
		    <%End if
        elseif txtRequestTable="PRODUCT" then
		     if rsCallCombo("PRODUCTCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("PRODUCTCODE")%>" Selected><%=rsCallCombo("PRODUCTDESCRIPTION")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("PRODUCTCODE")%>"><%=rsCallCombo("PRODUCTDESCRIPTION")%></option>  
		    <%End if
        elseif txtRequestTable="CUSTOMER" then
		     if rsCallCombo("CUSTOMERCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("CUSTOMERCODE")%>" Selected><%=rsCallCombo("DESCRIPTION")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("CUSTOMERCODE")%>"><%=rsCallCombo("DESCRIPTION")%></option>  
		    <%End if
        elseif txtRequestTable="SHIFT" then
		     if rsCallCombo("SHIFTCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("SHIFTCODE")%>" Selected><%=rsCallCombo("DESCRIPTION")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("SHIFTCODE")%>"><%=rsCallCombo("DESCRIPTION")%></option>  
		    <%End if
        elseif txtRequestTable="WEIGHBRIDGE" then
		     if rsCallCombo("WEIGHBRIDGECODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("WEIGHBRIDGECODE")%>" Selected><%=rsCallCombo("DESCRIPTION")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("WEIGHBRIDGECODE")%>"><%=rsCallCombo("DESCRIPTION")%></option>  
		    <%End if
        elseif txtRequestTable="AGREEMENTDETAIL" then
		     if rsCallCombo("AGREEMENTCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("AGREEMENTCODE")%>" Selected><%=rsCallCombo("AGREEMENTCODE")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("AGREEMENTCODE")%>"><%=rsCallCombo("AGREEMENTCODE")%></option>  
		    <%End if%>
		<%End if%>
		<%rsCallCombo.movenext
	 loop
	End if
	rsCallCombo.Close
	set rsCallCombo = Nothing
	q=""
End Sub


Sub cmdCallComboLast(txtRequestTable,txtRequestFields,txtFieldWhere,txtValueWhere)
    if txtFieldWhere = "" then
   	   sqlUser = "select * from " & txtRequestTable & " Order by fcTitle"
	else
   	   sqlUser = "select * from " & txtRequestTable & " Where " & txtFieldWhere & " = '" & txtValueWhere & "'" & " Order by fcTitle"
	End if
	Set rsCallCombo = con.Execute(sqlUser)
	if not rsCallCombo.eof then
	 Do While Not rsCallCombo.eof
		if rsCallCombo("fcValue") = txtRequestFields then%>
		   <option value="<%=rsCallCombo("fcValue")%>" Selected><%=rsCallCombo("fcTitle")%></option> 
		<%Else%>
		   <option value="<%=rsCallCombo("fcValue")%>"><%=rsCallCombo("fcTitle")%></option>  
		<%End if%>
		<%rsCallCombo.movenext
	 loop
	End if
	rsCallCombo.Close
	set rsCallCombo = Nothing
	sqlUser=""
End Sub

Function cmdFindTitle(txtRequestTable,txtFieldWhere,txtValueWhere)
    Dim sqlUser,rsFindTitle
	sqlUser = "select * from " & txtRequestTable & " Where " & txtFieldWhere & " = '" & trim(txtValueWhere) & "'"
	Set rsFindTitle = connStr.Execute(sqlUser)
	if not rsFindTitle.eof then
       cmdFindTitle = rsFindTitle("fcName")
	else
	   cmdFindTitle = ""   
	End if
	rsFindTitle.Close
	set rsFindTitle = Nothing
End Function

Function cmdFindTitle_BA(txtRequestTable,txtFieldWhere,txtValueWhere)
    Dim sqlUser,rsFindTitle
	sqlUser = "select * from " & txtRequestTable & " Where " & txtFieldWhere & " = '" & trim(txtValueWhere) & "' and fcBA='" & session("fcBA") & "'"
	Set rsFindTitle = connStr.Execute(sqlUser)
	if not rsFindTitle.eof then
       cmdFindTitle_BA = rsFindTitle("fcName")
	else
	   cmdFindTitle_BA = ""   
	End if
	rsFindTitle.Close
	set rsFindTitle = Nothing
End Function

Function cmdFindPeriod(txtValueWhere)
    dim sqluser,con,rsFindTitle
	rsFindTitle=Server.CreateObject("ADODB.Recordset")
	sqlUser = "select to_char(currentdate,'mm/dd/yyyy') as txtPeriod from PeriodControl Where upper(system) = '" & uCase(trim(txtValueWhere)) & "'" & _
              " and fcba='" & Session("fcBa") & "'"
	Set rsFindTitle = connStr.Execute(sqlUser)
	if not rsFindTitle.eof then
       cmdFindPeriod = rsFindTitle("txtPeriod")
	else
	   cmdFindPeriod = ""   
	End if
	rsFindTitle.Close
	set rsFindTitle = Nothing
End Function

Function cmdPeriodClose(txtValueWhere)
    dim sqluser,con,rsFindClose
	rsFindClose=Server.CreateObject("ADODB.Recordset")
	if Trim(txtValueWhere)="Financial" Then
	   sqlUser = "select to_char(add_months(LASTPERIODPOSTED,2),'mm/dd/yyyy') as txtPeriod from PeriodControl Where upper(system) = '" & "COSTBOOK" & "'" & _
	   " and fcba='" & Session("fcBa") & "'"
	   'response.write sqlUser
    else
	   sqlUser = "select to_char(LASTPERIODPOSTED,'mm/dd/yyyy') as txtPeriod from PeriodControl Where upper(system) = '" & uCase(trim(txtValueWhere)) & "'" & _
	   " and fcba='" & Session("fcBa") & "'"
	end if
	
    'response.write sqlUser
	Set rsFindClose = connStr.Execute(sqlUser)
	if not rsFindClose.eof then
       cmdPeriodClose = rsFindClose("txtPeriod")
	else
	   cmdPeriodClose = ""   
	End if
	rsFindClose.Close
	set rsFindClose = Nothing
End Function



Function CheckValue(nVarValue)
    if isnull(nVarValue) then
	   CheckValue=0
	else
	   CheckValue=FormatNumber(nVarValue,2)
    end if
End Function

Function To_Nol(nVarValue)
    if len(trim(nVarValue)) > 0 then
	   To_Nol=nVarValue
	else
	   To_Nol=0
    end if
End Function

Function FindEndOfMonth(nMonth,nYear) 
	dMonth = nMonth
	dYear  = nYear
	if dMonth=12 then
	   dDummyDate = "01" & "-" & "01" & "-" & dYear+1
	else
	   dDummyDate = dMonth+1 & "-" & "01" & "-" & dYear
	end if
	dDummyDate = cdate(dDummyDate) - 1
	FindEndOfMonth = dDummyDate 
End Function

Function textMonth(nMonth)
    nMonth = cInt(nMonth)
    Select Case nMonth
	   case 1:textMonth="Januari"
	   case 2:textMonth="Februari"
	   case 3:textMonth="Maret"
	   case 4:textMonth="April"
	   case 5:textMonth="Mei"
	   case 6:textMonth="Juni"
	   case 7:textMonth="Juli"
	   case 8:textMonth="Agustus"
	   case 9:textMonth="September"
	   case 10:textMonth="Oktober"
	   case 11:textMonth="November"
	   case 12:textMonth="Desember"
	   case 13:textMonth="Bulan 13"
	   case 14:textMonth="Bulan 14"
	   case 15:textMonth="Bulan 15"
	end Select
End Function


Function textMonthProg(nMonth)
    nMonth = cInt(nMonth)
	if nMonth=8 or nMonth=9 then
	   nMonth=10
	elseif nMonth=11 or nMonth=12 then   
	   nMonth=1
	elseif nMonth=2 or nMonth=3 then   
	   nMonth=4
	elseif nMonth=5 or nMonth=6 then   
	   nMonth=7	   
	end if
    Select Case nMonth
	   case 1:textMonth="Januari"
	   case 2:textMonth="Februari"
	   case 3:textMonth="Maret"
	   case 4:textMonth="April"
	   case 5:textMonth="Mei"
	   case 6:textMonth="Juni"
	   case 7:textMonth="Juli"
	   case 8:textMonth="Agustus"
	   case 9:textMonth="September"
	   case 10:textMonth="Oktober"
	   case 11:textMonth="November"
	   case 12:textMonth="Desember"
	   case 13:textMonth="Bulan 13"
	   case 14:textMonth="Bulan 14"
	   case 15:textMonth="Bulan 15"
	end Select
End Function

Function encrypt(x1, x2)
    Dim s,t,y,i
    s = ""
    t = 0
    For i = 1 to len(x1)
        t = t + asc(mid(x1,i,1))
    Next
    For i = 1 to len(x2)
        y = (t + asc(mid(x2,i,1)) * asc(mid(x2,((i+1) mod len(x2)+1),1))) mod 255
        s = s & chr(y)
    Next
    For i = (len(x2) + 1) to 10
        If t>598.8 Then t = 598.8
        y = t^3*i mod 255
        s = s & chr(y)
    Next
    encrypt = s
End Function

Function textDate(xDate)
    Dim cDay,cMonth
    cDay = Day(xDate)
	if len(cDay) = 1 then
	   cday="0"&cDay
	end if
	cMonth = Month(xDate)
	if len(cMonth) = 1 then
	   cMonth="0"&cMonth
	end if
	textDate=cMonth&"/"&cDay&"/"&year(xDate)
End Function

Function Rpad(sValue, sPadchar, iLength)
	Rpad = sValue & string(iLength - Len(sValue), sPadchar)
End Function

Function Lpad(sValue, sPadchar, iLength)
	Lpad = string(iLength - Len(sValue),sPadchar) & sValue
End Function

' Function to get last day of the month
' Requires to have the month number and year passed to the function
Function GetLastDay(initialMonth,initialYear)

	Dim initialDay

	Select Case initialMonth
		Case 1, 3, 5, 7, 8, 10, 12
			initialDay = 31
		Case 4, 6, 9, 11
			initialDay = 30
		Case 2
			If initialYear mod 4 = 0 Then
				If initialYear mod 100 = 0 AND initialYear mod 400 <> 0 Then
					initialDay = 28
				Else
					initialDay = 29
				End If
			Else
				initialDay = 28
			End If
	End Select

	GetLastDay = initialDay

End Function


Function CheckIsTM(initialField, initialFcba)
	Dim initialAvailableTM
	Dim rsAvailableEmployee, strAvailableEmployee
	
	strAvailableEmployee =	" Select Count(1) Total From Field Where" & _
							" Fccode = '" & initialField & "'" & _
							" And Activation = 'Y' " & _
							" And Status = 'TM' " & _
							" And Fcba = '" & initialFcba & "'"
							
	Set rsAvailableEmployee = Server.CreateObject("ADODB.Recordset")
	rsAvailableEmployee.Open strAvailableEmployee, connstr, 3
	
	If Not rsAvailableEmployee.EOF Then
		initialAvailableTM = Cdbl(rsAvailableEmployee("Total"))
	Else
		initialAvailableTM = 0
	End If
	
	rsAvailableEmployee.Close
	set rsAvailableEmployee = Nothing

	CheckIsTM = initialAvailableTM
End Function

Function PCase(strInput)
    Dim iPosition ' Our current position in the string (First character = 1)
    Dim iSpace ' The position of the next space after our iPosition
    Dim strOutput ' Our temporary string used to build the function's output

    iPosition = 1

    Do While InStr(iPosition, strInput, " ", 1) <> 0
            iSpace = InStr(iPosition, strInput, " ", 1)
            strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
            strOutput = strOutput & LCase(Mid(strInput, iPosition + 1, iSpace - iPosition))
            iPosition = iSpace + 1
    Loop

    strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
    strOutput = strOutput & LCase(Mid(strInput, iPosition + 1))

    PCase = strOutput
End Function

Function CheckAvailableVehicle(initialVehicle, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" Select Count(1) Total From Vehicle Where" & _
					" Fccode = '" & Trim(initialVehicle) & "'" & _
					" And Upper(Activation) = 'Y' " & _
					" And Fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckAvailableVehicle = initialAvailable
End Function

Function CheckAvailableContract(initialContract, initialTanggal, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" Select Count(1) Total From (" & _
					"     Select distinct AgreementCode, DueDate, Fcba From AgreementDetail Where" & _
					"         fcba||job in (" & _
					"             select fcba||param_1 from parameterdetail" & _
					"                 where parcode = 'HARVESTINGVEHICLESTATUS' and parheadcode = '02'" & _
					"         )" & _
					"     ) Where" & _					
					" AgreementCode = '" & Trim(initialContract) & "'" & _
					" And To_Date('" & initialTanggal & "', 'MM/DD/YYYY') <= DueDate" & _
					" And Fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckAvailableContract = initialAvailable
End Function
Function dateCallCombo(txtMont,txtTgl,txtYear)
	if LEN(trim(txtMont)) > 0 then %>
		<select name="<%=txtMont%>">
		<option style="color:#ccc;" value="0" disabled selected>Month</option>
		<% for CombMonth=1 to 12 %>
			<option value="<%=CombMonth%>"><%=MonthName(CombMonth,True)%></option>

		<% Next %>
		</select> 
	<%
    end if
	if LEN(trim(txtTgl)) > 0 then %>
		<select name="<%=txtTgl%>">
		<option style="color:#ccc;" value="0" disabled selected>Day</option>
		<% for CombDay=1 to 31 %>
				<option value="<%=CombDay%>"><%=CombDay%></option>
		<% Next %>
		</select> 
	<%

    end if 	
	if LEN(trim(txtYear)) > 0 then %>	
		<select name="<%=txtYear%>">
		<option style="color:#ccc;" value="0" disabled selected>Year</option>
		<% for CombYear=Year(date()) to Year(date())-10 step -1 %>
			<option value="<%=CombYear%>"><%=CombYear%></option>
		<% Next %>
		</select> 
	<%
    end if
End Function
sub dateCallComboUpdate(txtMont,txtTgl,txtYear,txtDate)
	if LEN(trim(txtMont)) > 0 then %>
		<select name="<%=txtMont%>">
		<option style="color:#ccc;" value="" disabled selected>Month</option>
		<% for CombMonth=1 to 12 
			if LEN(trim(txtDate)) > 0 and Month(txtDate) = CombMonth then %>
			<option value="<%=CombMonth%>" selected><%=MonthName(CombMonth,True)%></option>
			<%ELSE%>
			<option value="<%=CombMonth%>"><%=MonthName(CombMonth,True)%></option>
			<%END IF%>
		<% Next %>
		</select> 
	<%
    end if
	if LEN(trim(txtTgl)) > 0 then %>
		<select name="<%=txtTgl%>">
		<option style="color:#ccc;" value="" disabled selected>Day</option>
		<% for CombDay=1 to 31 
			if LEN(trim(txtDate)) > 0 and Day(txtDate) = CombDay then %>
				<option value="<%=CombDay%>" selected><%=CombDay%></option>
				<%ELSE%>
				<option value="<%=CombDay%>"><%=CombDay%></option>
				<%END IF%>
		<% Next %>
		</select> 
	<%

    end if 	
	if LEN(trim(txtYear)) > 0 then %>	
	<input type="text" name="<%=txtYear%>" value="<%=Year(txtDate)%>" style="width:60px;" placeholder="Year">
	<%
    end if
End sub

'fungsion null
Function Nz(val, defaultVal)
    If IsNull(val) Or val = "" Then
        Nz = defaultVal
    Else
        Nz = val
    End If
End Function

%>