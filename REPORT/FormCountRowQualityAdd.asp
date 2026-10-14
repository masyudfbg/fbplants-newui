    <%
	
		    finalRowSpan = 0
			
			if not rsCPO_FFAQuality.EOF then 
				while not rsCPO_FFAQuality.EOF
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

				
				
				    if tempStrCompare  = "F F A" or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsCPO_FFAQuality.MoveNext
				Wend
			    rsCPO_FFAQuality.MoveFirst
			end if
			
		
		

			
	%>
