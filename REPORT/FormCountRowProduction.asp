    <%
	
		    finalRowSpan = 0
			
			if not rsProduction.EOF then 
				while not rsProduction.EOF
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

				
				
				    if tempStrCompare  = rsProduction("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsProduction.MoveNext
				Wend
			    rsProduction.MoveFirst
			end if
			
			if not rsPusingan.EOF then 
				while not rsPusingan.EOF
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

				
				
				    if tempStrCompare  = rsPusingan("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsPusingan.MoveNext
				Wend
				rsPusingan.MoveFirst	
			end if

			if not rsOutputPanenFinal.EOF then 
				while not rsOutputPanenFinal.EOF
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

				
				
				    if tempStrCompare  = rsOutputPanenFinal("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsOutputPanenFinal.MoveNext
				Wend
			    rsOutputPanenFinal.MoveFirst	
			end if
			
			' if v_txtDetailInfo = "Yes" then
				' if not rsOutputPanenFinal.EOF then 			
					' while not rsOutputPanenFinal.EOF
						' finalRowSpan = finalRowSpan + 1
						' rsOutputPanenFinal.MoveNext
					' Wend
					' rsOutputPanenFinal.MoveFirst	
				' end if
			' end if
			
			if not rsBiayaMengumpulFinal.EOF then 
				while not rsBiayaMengumpulFinal.EOF
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

				
				
				    if tempStrCompare  = rsBiayaMengumpulFinal("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsBiayaMengumpulFinal.MoveNext
				Wend
				rsBiayaMengumpulFinal.MoveFirst		
			end if
			
			' if v_txtDetailInfo = "Yes" then
				' if not rsBiayaMengumpulFinal.EOF then 
					' while not rsBiayaMengumpulFinal.EOF
						' finalRowSpan = finalRowSpan + 1
						' rsBiayaMengumpulFinal.MoveNext
					' Wend
					' rsBiayaMengumpulFinal.MoveFirst	
					
					' while not rsBiayaMengumpulFinal.EOF
						' finalRowSpan = finalRowSpan + 1
						' rsBiayaMengumpulFinal.MoveNext
					' Wend
					' rsBiayaMengumpulFinal.MoveFirst					
				' end if
			' end if

			
			if not rsTBSPengangkutanFinal.EOF then 
				while not rsTBSPengangkutanFinal.EOF
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

				
				
				    if tempStrCompare  = rsTBSPengangkutanFinal("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsTBSPengangkutanFinal.MoveNext
				Wend
				rsTBSPengangkutanFinal.MoveFirst	
			end if
			
			' if v_txtDetailInfo = "Yes" then
				' if not rsTBSPengangkutanFinal.EOF then 			
					' while not rsTBSPengangkutanFinal.EOF
						' finalRowSpan = finalRowSpan + 1
						' rsTBSPengangkutanFinal.MoveNext
					' Wend
					' rsTBSPengangkutanFinal.MoveFirst	
					
					' while not rsTBSPengangkutanFinal.EOF
						' finalRowSpan = finalRowSpan + 1
						' rsTBSPengangkutanFinal.MoveNext
					' Wend
					' rsTBSPengangkutanFinal.MoveFirst
				' end if				
			' end if			

			
			if not rsCurahHujan.EOF then 
				while not rsCurahHujan.EOF
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

				
				
				    if tempStrCompare  = rsCurahHujan("description") or v_txtGroupItem  = "--NONE--"  then 						
						finalRowSpan = finalRowSpan + 1
					end if
					rsCurahHujan.MoveNext
				Wend
				rsCurahHujan.MoveFirst		
			end if
	
	%>
