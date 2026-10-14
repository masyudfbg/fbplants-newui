    <%
	
		    finalRowSpan = 0
			
			if not rsRendemenFinalAdd.EOF then 
				while not rsRendemenFinalAdd.EOF
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

				
				
				    if tempStrCompare  = "O E R" or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsRendemenFinalAdd.MoveNext
				Wend
			    rsRendemenFinalAdd.MoveFirst
			end if
			
	

			
	%>
