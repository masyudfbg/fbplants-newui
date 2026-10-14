    <%
	
		    finalRowSpan = 0
			
			if not rsSortasi.EOF then 
				while not rsSortasi.EOF
					finalRowSpan = finalRowSpan + 1
					rsSortasi.MoveNext
				Wend
			    rsSortasi.MoveFirst
			end if
			
	
	%>
