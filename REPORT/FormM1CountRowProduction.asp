    <%
	
		    finalRowSpan = 0
			
			if not rsProduction.EOF then 
				while not rsProduction.EOF
					finalRowSpan = finalRowSpan + 1
					rsProduction.MoveNext
				Wend
			    rsProduction.MoveFirst
			end if
			
			if not rsPusingan.EOF then 
				while not rsPusingan.EOF
					finalRowSpan = finalRowSpan + 1
					rsPusingan.MoveNext
				Wend
				rsPusingan.MoveFirst	
			end if

			if not rsOutputPanenFinal.EOF then 
				while not rsOutputPanenFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsOutputPanenFinal.MoveNext
				Wend
			    rsOutputPanenFinal.MoveFirst	
			end if
			
			if v_txtDetailInfo = "Yes" then
				while not rsOutputPanenFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsOutputPanenFinal.MoveNext
				Wend
			    rsOutputPanenFinal.MoveFirst	
			end if
			
			if not rsBiayaMengumpulFinal.EOF then 
				while not rsBiayaMengumpulFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsBiayaMengumpulFinal.MoveNext
				Wend
				rsBiayaMengumpulFinal.MoveFirst		
			end if
			
			if v_txtDetailInfo = "Yes" then
				while not rsBiayaMengumpulFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsBiayaMengumpulFinal.MoveNext
				Wend
			    rsBiayaMengumpulFinal.MoveFirst	
				
				while not rsBiayaMengumpulFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsBiayaMengumpulFinal.MoveNext
				Wend
			    rsBiayaMengumpulFinal.MoveFirst					
			end if

			
			if not rsTBSPengangkutanFinal.EOF then 
				while not rsTBSPengangkutanFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsTBSPengangkutanFinal.MoveNext
				Wend
				rsTBSPengangkutanFinal.MoveFirst	
			end if
			
			if v_txtDetailInfo = "Yes" then
				while not rsTBSPengangkutanFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsTBSPengangkutanFinal.MoveNext
				Wend
			    rsTBSPengangkutanFinal.MoveFirst	
				
				while not rsTBSPengangkutanFinal.EOF
					finalRowSpan = finalRowSpan + 1
					rsTBSPengangkutanFinal.MoveNext
				Wend
			    rsTBSPengangkutanFinal.MoveFirst
				
			end if			

			
			if not rsCurahHujan.EOF then 
				while not rsCurahHujan.EOF
					finalRowSpan = finalRowSpan + 1
					rsCurahHujan.MoveNext
				Wend
				rsCurahHujan.MoveFirst		
			end if
	
	%>
