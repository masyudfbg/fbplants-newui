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

Sub cmdCallComboNonFcba(txtRequestTable,txtRequestFields,txtFieldWhere,txtValueWhere)
 	Dim rsCallCombo
    if txtFieldWhere = "" then
   	   q = "select * from " & txtRequestTable & " Order by FCCODE"
	else
   	   q = "select * from " & txtRequestTable & " Where " & txtFieldWhere & " = '" & txtValueWhere & "'  Order by FCCODE"
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

Sub cmdCallCombo(txtRequestTable,txtRequestFields,txtFieldWhere,txtValueWhere)
 	Dim rsCallCombo

    if txtFieldWhere = "" then
		if txtRequestTable = "STOCKSUBGROUP" THEN 
   	   		q = "select * from " & txtRequestTable & " Order by FCCODE"
		ELSE 
   	   		q = "select * from " & txtRequestTable & " where fcBA='" &  Session("fcBA") & "' Order by FCCODE"
		END IF
	else
   	   q = "select * from " & txtRequestTable & " Where " & txtFieldWhere & " = '" & txtValueWhere & "' and fcBA='" & session("fcBA") & "' Order by FCCODE"
	End if
	response.write q
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
	'response.write q
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
		elseif txtRequestTable="COMPANY" then
		     if rsCallCombo("COMPANYCODE") = txtRequestFields then%>
		        <option value="<%=rsCallCombo("COMPANYCODE")%>" Selected><%=rsCallCombo("COMPANYCODE")%></option> 
		     <%Else%>
		        <option value="<%=rsCallCombo("COMPANYCODE")%>"><%=rsCallCombo("COMPANYCODE")%></option>  
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


Function cmdFindAccountPeriod(txtValueWhere)
    dim sqluser,con,rsFindTitle
	rsFindTitle=Server.CreateObject("ADODB.Recordset")
	sqlUser = "select currentaccountingyear||lpad(currentaccountingperiod,2,0) as txtPeriod from PeriodControl Where upper(system) = '" & uCase(trim(txtValueWhere)) & "'" & _
              " and fcba='" & Session("fcBa") & "'"
	Set rsFindTitle = connStr.Execute(sqlUser)
	if not rsFindTitle.eof then
       cmdFindAccountPeriod = rsFindTitle("txtPeriod")
	else
	   cmdFindAccountPeriod = ""   
	End if
	rsFindTitle.Close
	set rsFindTitle = Nothing
End Function

Function cmdAccountPeriodClose(txtValueWhere)
    dim sqluser,con,rsFindClose
	rsFindClose=Server.CreateObject("ADODB.Recordset")
	if Trim(txtValueWhere)="Financial" Then
	   sqlUser = "select closeaccountingyear||lpad(closeaccountingperiod+2,2,0) as txtPeriod from PeriodControl Where upper(system) = '" & "COSTBOOK" & "'" & _
	   " and fcba='" & Session("fcBa") & "'"
	   'response.write sqlUser
    else
	   sqlUser = "select closeaccountingyear||lpad(closeaccountingperiod,2,0) as txtPeriod from PeriodControl Where upper(system) = '" & uCase(trim(txtValueWhere)) & "'" & _
	   " and fcba='" & Session("fcBa") & "'"
	end if
	
    'response.write sqlUser
	Set rsFindClose = connStr.Execute(sqlUser)
	if not rsFindClose.eof then
       cmdAccountPeriodClose = rsFindClose("txtPeriod")
	else
	   cmdAccountPeriodClose = ""   
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

Function cDay(xDate)
    Dim sDay
    sDay = Day(xDate)
	if len(sDay) = 1 then
	   sDay="0"&sDay
	end if
	cDay=sDay
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
					" AgreementCode = '" & initialContract & "'" & _
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
function Terbilang(satuan)
			dim huruf
				huruf = array ("", "SATU", "DUA", "TIGA", "EMPAT", "LIMA", "ENAM","TUJUH", "DELAPAN", "SEMBILAN", "SEPULUH","SEBELAS")
				
				if satuan < 12 then  
				  response.write (" " & huruf(satuan))
				elseif (satuan < 20) then   
				  response.write (Terbilang(satuan - 10) & " BELAS")
				elseif (satuan < 100) then 
				  response.write (Terbilang((satuan - (satuan mod 10)) / 10) & " PULUH")
				  response.write (Terbilang(satuan mod 10))
				elseif (satuan < 200)  then
				  response.write ("SERATUS")
				  response.write (Terbilang(satuan - 100))
				elseif (satuan < 1000) then 
				  response.write (Terbilang((satuan - (satuan mod 100)) / 100) & " RATUS")
				  response.write (Terbilang(satuan mod 100))
				elseif (satuan < 2000) then 
				  response.write ("SERIBU")
				  response.write (Terbilang(satuan - 1000))
				elseif (satuan < 1000000)  then
				  response.write (Terbilang((satuan - (satuan mod 1000)) / 1000)&" RIBU")
				  response.write (Terbilang(satuan mod 1000))
				elseif (satuan < 1000000000)  then
				  response.write (Terbilang((satuan - (satuan mod 1000000)) / 1000000) & " JUTA")
				  response.write (Terbilang(satuan mod 1000000))
				elseif (satuan < 1000000000000)  then
				  response.write (Terbilang((satuan - (satuan mod 1000000000)) / 1000000000) & " MILYAR")
				  response.write (Terbilang(satuan mod 1000000000)) 
				elseif (satuan >= 1000000000000)  then
				  response.write ("")
				end if 
End Function  


Function PROFIL(txtName)
    if len(trim(txtName)) > 0 then 
		dim QCOPY,COPY,COMPANY_IMAGE,APPLICATION_IMAGE,REPORTDASHBOARD_IMAGE,REPORT_IMAGE
			
			QCOPY="select COPYRIGHT,COMPANY_IMAGE,APPLICATION_IMAGE,REPORTDASHBOARD_IMAGE,REPORT_IMAGE from APPLICATIONIMAGE where fcBA='" & session("fcBA") & "'"
			 
			 Set COPY = Server.CreateObject("ADODB.Recordset")
			   COPY.CursorType = 3
			   COPY.CursorLocation = 3
			   COPY.Open QCOPY, connstr 
			if Not COPY.eof then   			  
				   PROFIL = COPY(txtName)
			end if 
				   COPY.Close
				   set COPY = Nothing
	else
		response.write ("COMPANY_IMAGE-APPLICATION_IMAGE-REPORTDASHBOARD_IMAGE-REPORT_IMAGE")
	end if
End Function


Function RoundUp(dblNumber, intDecimals)
	Dim dblFactor, dblTemp
	'round up dblNumber to next highest number with intDecimals accuracy
    dblFactor = 10 ^ intDecimals
    dblTemp = dblNumber * dblFactor + 0.99999
    RoundUp = Int(dblTemp) / dblFactor
End Function

Function GetQtyHariKerja(StartDate,EndDate) 
                Dim QtyHariKerja
                Dim rsQtyHariKerja, strQtyHariKerja
                
                strQtyHariKerja =  " WITH dates AS ( " & _
									"	SELECT TO_DATE('" & StartDate & "', 'MM/DD/YYYY') AS start_date," & _
									"	       TO_DATE('" & EndDate   & "', 'MM/DD/YYYY') AS end_date " & _
									"	FROM dual " & _
									" )," & _
									"sundays AS ( " & _
									"SELECT start_date," & _
									"	   end_date," & _
									"	   NEXT_DAY(TRUNC(start_date) - 1, 'SUNDAY') AS first_sunday," & _
									"	   NEXT_DAY(TRUNC(end_date) - 1, 'SUNDAY') AS last_sunday " & _
									"FROM dates " & _
									") " & _
									" " & _
									"SELECT Sum(num_days) num_days FROM " & _
									"( " & _
									"SELECT 'X' id,(end_date - start_date) num_days FROM sundays " & _
									"UNION ALL " & _
									"SELECT 'Y' id," & _
									"	CASE " & _
									"		WHEN first_sunday > end_date THEN 0 " & _
									"		ELSE FLOOR((end_date - first_sunday) / 7) + 1 " & _
									"	END *-1 AS num_days " & _
									"FROM sundays " & _
									"UNION ALL " & _
									"SELECT 'Z' id,Count(DISTINCT holidaydate)*-1 num_days " & _
									"FROM budget_transaction_holiday a INNER JOIN sundays b ON (a.holidaydate BETWEEN b.start_date AND b.end_date) " & _
									"WHERE a.fcba ='ARJ' AND TO_CHAR(a.holidaydate, 'D') <> '1' " & _
									") "
									
									'response.write strQtyHariKerja
                                                                             
                Set rsQtyHariKerja = Server.CreateObject("ADODB.Recordset")
                rsQtyHariKerja.Open strQtyHariKerja, connstr, 3
                
                If Not rsQtyHariKerja.EOF Then
                                QtyHariKerja = cint(rsQtyHariKerja("num_days"))
                Else
                                QtyHariKerja = 0
                End If
                
                rsQtyHariKerja.Close
                set rsQtyHariKerja = Nothing
                GetQtyHariKerja = QtyHariKerja
end function

Function FmtDateMMDDYYYY(dateStr)
    ' Memecah string tanggal menjadi bagian bulan, hari, dan tahun
    Dim month, day, year
    month = CInt(Mid(dateStr, 1, 2))
    day = CInt(Mid(dateStr, 4, 2))
    year = CInt(Mid(dateStr, 7, 4))
    
    ' Menghasilkan objek Date dengan menggunakan DateSerial
    FmtDateMMDDYYYY = DateSerial(year, month, day)
End Function


Function GetHariLibur(p_data)
Dim q_libur ,rsLibur
q_libur =   "SELECT Count(DISTINCT holidaydate) num_days " & _
			"FROM budget_transaction_holiday " & _
			"WHERE fcba ='ARJ' AND holidaydate = TO_DATE('" & p_data & "', 'DD/MM/YYYY') AND TO_CHAR(holidaydate, 'D') <> '1' " 

	 Set rsLibur = Server.CreateObject("ADODB.Recordset")
     rsLibur.Open q_libur, connstr, 3
                
     If Not rsLibur.EOF Then
        QtyHariKerja = cint(rsLibur("num_days"))
     Else
        QtyHariKerja = 0
     End If
                
    rsLibur.Close
    set rsLibur = Nothing
    GetHariLibur = QtyHariKerja			

end Function

Function GetQtyHariKerjaBaru(tglstart, tglend)
    ' Mengubah string tanggal menjadi objek Date
	   
	Dim startDate, endDate
    startDate = FmtDateMMDDYYYY(tglstart)
    endDate   = FmtDateMMDDYYYY(tglend)
    
    ' Membuat variabel untuk menghitung jumlah hari Minggu
    Dim JmlHari
    JmlHari = 0
    
    ' Loop melalui setiap hari di antara rentang tanggal
    Dim currentDate
    currentDate = startDate
    Do While currentDate < endDate
	    JmlHari = JmlHari + 1
        
		' Memeriksa apakah hari ini adalah hari Minggu
        If Weekday(currentDate) = vbSunday Then
		    JmlHari = JmlHari - 1
		End If
		
		' Memeriksa apakah hari ini adalah hari Libur
		if GetHariLibur(currentDate) > 0 Then
			JmlHari = JmlHari - 1
		End if
        	
        ' Menambahkan satu hari ke tanggal saat ini
        currentDate = DateAdd("d", 1, currentDate)
    Loop
    
	' Mengembalikan jumlah hari 
    GetQtyHariKerjaBaru = JmlHari
End Function


Function StatusAccClosingTgl(pTgl,pFcba) 
                Dim StatClosing
                Dim rsClosing, strClosing
                
                strClosing =     " SELECT  " & _ 
								 " empalldedstatus_api.getstatusclosing_acc(EXTRACT(YEAR FROM to_date('"&pTgl&"', 'mm/dd/yyyy')),EXTRACT(MONTH FROM to_date('"&pTgl&"', 'mm/dd/yyyy')),'" & pFcba & "') closingAcc " & _
								 " FROM dual"
								'response.write strClosing
                                                                             
                Set rsClosing = Server.CreateObject("ADODB.Recordset")
                rsClosing.Open strClosing, connstr, 3
                
                If Not rsClosing.EOF Then
                                StatClosing = rsClosing("closingAcc")
                Else
                                StatClosing = "FALSE"
                End If
                
                rsClosing.Close
                set rsClosing = Nothing
                StatusAccClosingTgl = StatClosing
end function

Function StatusAccClosing(pYear,pMonth,pFcba) 
                Dim StatClosing
                Dim rsClosing, strClosing
                
                strClosing =     " SELECT  " & _ 
								 " empalldedstatus_api.getstatusclosing_acc(" & pYear & "," & pMonth & ",'" & pFcba & "') closingAcc " & _
								 " FROM dual"
								'response.write strClosing
                                                                             
                Set rsClosing = Server.CreateObject("ADODB.Recordset")
                rsClosing.Open strClosing, connstr, 3
                
                If Not rsClosing.EOF Then
                                StatClosing = rsClosing("closingAcc")
                Else
                                StatClosing = "FALSE"
                End If
                
                rsClosing.Close
                set rsClosing = Nothing
                StatusAccClosing = StatClosing
end function

Function KonversiTanggalKeDbl(dateStr)
    ' Pecah tanggal berdasarkan pemisah "/"
    Dim dateParts
	Dim dateInt
    dateParts = Split(dateStr, "/")
    
    ' Jika format tanggal tidak valid, kembalikan nilai kosong
    If UBound(dateParts) <> 2 Then
        KonversiTanggalKeDbl = 0
        Exit Function
    End If

    ' Ambil bagian bulan, hari, dan tahun
    Dim month, day, year
    month = dateParts(0)
    day = dateParts(1)
    year = dateParts(2)

    ' Pastikan bulan dan hari selalu dua digit
    If Len(month) = 1 Then month = "0" & month
    If Len(day) = 1 Then day = "0" & day

    ' Gabungkan kembali dalam format yyyymmdd
	dateInt = Cdbl(year & month & day)
    KonversiTanggalKeDbl = dateInt
End Function

'fungsion null
Function Nz(val, defaultVal)
    If IsNull(val) Or val = "" Then
        Nz = defaultVal
    Else
        Nz = val
    End If
End Function

%>