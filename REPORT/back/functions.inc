<%
Function Bulan(xMonth)
    xMonth = cInt(xMonth)
    Select Case xMonth
	   case 1:Bulan="Januari"
	   case 2:Bulan="Februari"
	   case 3:Bulan="Maret"
	   case 4:Bulan="April"
	   case 5:Bulan="Mei"
	   case 6:Bulan="Juni"
	   case 7:Bulan="Juli"
	   case 8:Bulan="Agustus"
	   case 9:Bulan="September"
	   case 10:Bulan="Oktober"
	   case 11:Bulan="November"
	   case 12:Bulan="Desember"
	end Select
End Function
%>