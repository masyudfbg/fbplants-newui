    <%
	
		    finalRowSpan = 0
			
			if not rsOutput.EOF then 
				while not rsOutput.EOF
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

				
				
				    if tempStrCompare  = rsOutput("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsOutput.MoveNext
				Wend
				rsOutput.MoveFirst
			end if
			
	
	%>
