    <%
	
		    finalRowSpan = 0
			
			if not rsEndingBalance.EOF then 
				while not rsEndingBalance.EOF
					strCompare =   " select factorreport from parameterreport " & _
									  " where " & _
									  " reportcode = 'LHK' and " & _
									  " itemcode = '" & v_txtGroupItem & "' "
					
					Set rsCompare = Server.CreateObject("ADODB.Recordset")	
					rsCompare.Open strCompare, connstr, adOpenStatic 		

					if not rsCompare.EOF then
						tempStrCompare = rsCompare ("factorreport")
					else
						tempStrCompare = ""					
					end if	

				
				
				    if tempStrCompare  = rsEndingBalance("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsEndingBalance.MoveNext
				Wend
				rsEndingBalance.MoveFirst
			end if
			
	
	%>
