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

Function BulanAP(xMonth)
    xMonth = cInt(xMonth)
    Select Case xMonth
	   case 1:BulanAP="Januari"
	   case 2:BulanAP="Februari"
	   case 3:BulanAP="Maret"
	   case 4:BulanAP="April"
	   case 5:BulanAP="Mei"
	   case 6:BulanAP="Juni"
	   case 7:BulanAP="Juli"
	   case 8:BulanAP="Agustus"
	   case 9:BulanAP="September"
	   case 10:BulanAP="Oktober"
	   case 11:BulanAP="November"
	   case 12:BulanAP="Desember"
	   case 13:BulanAP="Period 13"
	   case 14:BulanAP="Period 14"
	   case 15:BulanAP="Period 15"
	end Select
End Function


Function ToMonth(xMonth)
    xMonth = cInt(xMonth)
    Select Case xMonth
	   case 1:Bulan="January"
	   case 2:Bulan="February"
	   case 3:Bulan="March"
	   case 4:Bulan="April"
	   case 5:Bulan="Mey"
	   case 6:Bulan="June"
	   case 7:Bulan="July"
	   case 8:Bulan="August"
	   case 9:Bulan="September"
	   case 10:Bulan="October"
	   case 11:Bulan="November"
	   case 12:Bulan="December"
	end Select
End Function

%>