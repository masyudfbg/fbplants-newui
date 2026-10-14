<%
'=====================================================================
' qrcode.inc - pembuat QR Code (VBScript murni, tanpa komponen luar)
' Mode byte, koreksi error level M, versi 1-6 (maks. 106 karakter).
' Dipakai untuk QR nomor PO di PDF lampiran email (email/po_mail_sender.inc).
'
' Pemakaian:
'   Dim qrSize, qrMods
'   qrMods = PoQr_Encode("001/ARJE1/PO/2026/10/0007", qrSize)
'   qrMods(y * qrSize + x) = True  -> modul gelap
' Mengembalikan Empty bila teks terlalu panjang.
'=====================================================================

Dim PoQr_Exp(511), PoQr_Log(255), PoQr_Ready
PoQr_Ready = False

Sub PoQr_InitGf()
	Dim i, x
	If PoQr_Ready Then Exit Sub
	x = 1
	For i = 0 To 254
		PoQr_Exp(i) = x
		PoQr_Log(x) = i
		x = x * 2
		If x >= 256 Then x = x Xor &H11D
	Next
	For i = 255 To 511
		PoQr_Exp(i) = PoQr_Exp(i - 255)
	Next
	PoQr_Ready = True
End Sub

Function PoQr_Mul(a, b)
	If a = 0 Or b = 0 Then
		PoQr_Mul = 0
	Else
		PoQr_Mul = PoQr_Exp(PoQr_Log(a) + PoQr_Log(b))
	End If
End Function

' Polinom pembagi Reed-Solomon berderajat n (koefisien tanpa suku tertinggi)
Function PoQr_Divisor(n)
	Dim res(), root, i, j
	ReDim res(n - 1)
	For i = 0 To n - 1 : res(i) = 0 : Next
	res(n - 1) = 1
	root = 1
	For i = 0 To n - 1
		For j = 0 To n - 1
			res(j) = PoQr_Mul(res(j), root)
			If j + 1 < n Then res(j) = res(j) Xor res(j + 1)
		Next
		root = PoQr_Mul(root, 2)
	Next
	PoQr_Divisor = res
End Function

Function PoQr_Remainder(data, dStart, dLen, divisor)
	Dim n, res(), i, j, factor
	n = UBound(divisor) + 1
	ReDim res(n - 1)
	For i = 0 To n - 1 : res(i) = 0 : Next
	For i = dStart To dStart + dLen - 1
		factor = data(i) Xor res(0)
		For j = 0 To n - 2 : res(j) = res(j + 1) : Next
		res(n - 1) = 0
		For j = 0 To n - 1
			res(j) = res(j) Xor PoQr_Mul(divisor(j), factor)
		Next
	Next
	PoQr_Remainder = res
End Function

' Tabel versi 1-6 level M: total codeword, jumlah blok, codeword ECC per blok, posisi alignment
Function PoQr_Spec(ver, which)
	Dim total, blocks, ecc, align
	Select Case ver
		Case 1 : total = 26  : blocks = 1 : ecc = 10 : align = 0
		Case 2 : total = 44  : blocks = 1 : ecc = 16 : align = 18
		Case 3 : total = 70  : blocks = 1 : ecc = 26 : align = 22
		Case 4 : total = 100 : blocks = 2 : ecc = 18 : align = 26
		Case 5 : total = 134 : blocks = 2 : ecc = 24 : align = 30
		Case 6 : total = 172 : blocks = 4 : ecc = 16 : align = 34
	End Select
	Select Case which
		Case "total" : PoQr_Spec = total
		Case "blocks" : PoQr_Spec = blocks
		Case "ecc" : PoQr_Spec = ecc
		Case "align" : PoQr_Spec = align
		Case "data" : PoQr_Spec = total - blocks * ecc
	End Select
End Function

Function PoQr_Encode(text, ByRef size)
	Dim bytes(), nBytes, i, c, ver, cap, bits(), nb, dataCw(), ndata
	Dim total, blocks, ecc, shortLen, nShort, k, b, blk(), blkLen(), divisor, j, outCw(), no
	Dim mods(), isFn(), x, y, best, bestPen, m, pen

	PoQr_InitGf
	PoQr_Encode = Empty
	' --- teks ke byte (non-ASCII diganti "?") ---
	nBytes = Len(text)
	If nBytes = 0 Then Exit Function
	ReDim bytes(nBytes - 1)
	For i = 1 To nBytes
		c = AscW(Mid(text, i, 1))
		If c < 0 Or c > 126 Then c = 63
		bytes(i - 1) = c
	Next
	' --- pilih versi terkecil ---
	ver = 0
	For i = 1 To 6
		If 4 + 8 + 8 * nBytes <= PoQr_Spec(i, "data") * 8 Then ver = i : Exit For
	Next
	If ver = 0 Then Exit Function
	ndata = PoQr_Spec(ver, "data")
	cap = ndata * 8

	' --- bit data: mode byte 0100, panjang 8 bit, isi, terminator, padding ---
	ReDim bits(cap - 1)
	nb = 0
	Call PoQr_PushBits(bits, nb, 4, 4)
	Call PoQr_PushBits(bits, nb, nBytes, 8)
	For i = 0 To nBytes - 1 : Call PoQr_PushBits(bits, nb, bytes(i), 8) : Next
	For i = 1 To 4
		If nb < cap Then bits(nb) = 0 : nb = nb + 1
	Next
	Do While nb Mod 8 <> 0 : bits(nb) = 0 : nb = nb + 1 : Loop
	ReDim dataCw(ndata - 1)
	For i = 0 To nb \ 8 - 1
		b = 0
		For j = 0 To 7 : b = b * 2 + bits(i * 8 + j) : Next
		dataCw(i) = b
	Next
	k = 0
	For i = nb \ 8 To ndata - 1
		If k Mod 2 = 0 Then dataCw(i) = &HEC Else dataCw(i) = &H11
		k = k + 1
	Next

	' --- ECC per blok + interleave ---
	total = PoQr_Spec(ver, "total")
	blocks = PoQr_Spec(ver, "blocks")
	ecc = PoQr_Spec(ver, "ecc")
	nShort = blocks - (total Mod blocks)
	shortLen = total \ blocks
	divisor = PoQr_Divisor(ecc)
	ReDim blk(blocks - 1), blkLen(blocks - 1)
	k = 0
	For i = 0 To blocks - 1
		blk(i) = PoQr_MakeBlock(dataCw, k, i, nShort, shortLen, ecc, divisor)
		blkLen(i) = shortLen - ecc
		If i >= nShort Then blkLen(i) = blkLen(i) + 1
		k = k + blkLen(i)
	Next
	ReDim outCw(total - 1)
	no = 0
	For i = 0 To shortLen
		For j = 0 To blocks - 1
			If i <> shortLen - ecc Or j >= nShort Then
				outCw(no) = blk(j)(i) : no = no + 1
			End If
		Next
	Next

	' --- matriks ---
	size = 17 + 4 * ver
	ReDim mods(size * size - 1), isFn(size * size - 1)
	For i = 0 To size * size - 1 : mods(i) = False : isFn(i) = False : Next
	Call PoQr_DrawFunctions(mods, isFn, size, ver)
	Call PoQr_DrawCodewords(mods, isFn, size, outCw, total)

	' --- pilih mask dengan penalti terkecil ---
	best = 0 : bestPen = -1
	For m = 0 To 7
		Call PoQr_ApplyMask(mods, isFn, size, m)
		Call PoQr_DrawFormat(mods, isFn, size, m)
		pen = PoQr_Penalty(mods, size)
		If bestPen < 0 Or pen < bestPen Then best = m : bestPen = pen
		Call PoQr_ApplyMask(mods, isFn, size, m)   ' kembalikan (XOR)
	Next
	Call PoQr_ApplyMask(mods, isFn, size, best)
	Call PoQr_DrawFormat(mods, isFn, size, best)
	PoQr_Encode = mods
End Function

' Satu blok: data + ECC; blok pendek diberi 1 placeholder agar panjang semua blok sama
Function PoQr_MakeBlock(dataCw, k, i, nShort, shortLen, ecc, divisor)
	Dim dl, arr(), eccRem, j
	dl = shortLen - ecc
	If i >= nShort Then dl = dl + 1
	ReDim arr(shortLen)
	For j = 0 To dl - 1 : arr(j) = dataCw(k + j) : Next
	eccRem = PoQr_Remainder(dataCw, k, dl, divisor)
	If i < nShort Then
		arr(dl) = 0
		For j = 0 To ecc - 1 : arr(dl + 1 + j) = eccRem(j) : Next
	Else
		For j = 0 To ecc - 1 : arr(dl + j) = eccRem(j) : Next
	End If
	PoQr_MakeBlock = arr
End Function

Sub PoQr_PushBits(bits, ByRef nb, value, length)
	Dim i
	For i = length - 1 To 0 Step -1
		bits(nb) = (value \ (2 ^ i)) Mod 2
		nb = nb + 1
	Next
End Sub

Sub PoQr_Set(mods, isFn, size, x, y, dark)
	mods(y * size + x) = dark
	isFn(y * size + x) = True
End Sub

Sub PoQr_DrawFunctions(mods, isFn, size, ver)
	Dim i, a, dx, dy, d, xx, yy, cx, cy, f
	For i = 0 To size - 1
		Call PoQr_Set(mods, isFn, size, 6, i, (i Mod 2 = 0))
		Call PoQr_Set(mods, isFn, size, i, 6, (i Mod 2 = 0))
	Next
	For f = 0 To 2
		If f = 0 Then cx = 3 : cy = 3
		If f = 1 Then cx = size - 4 : cy = 3
		If f = 2 Then cx = 3 : cy = size - 4
		For dy = -4 To 4
			For dx = -4 To 4
				d = Abs(dx) : If Abs(dy) > d Then d = Abs(dy)
				xx = cx + dx : yy = cy + dy
				If xx >= 0 And xx < size And yy >= 0 And yy < size Then
					Call PoQr_Set(mods, isFn, size, xx, yy, (d <> 2 And d <> 4))
				End If
			Next
		Next
	Next
	a = PoQr_Spec(ver, "align")
	If a > 0 Then
		For dy = -2 To 2
			For dx = -2 To 2
				d = Abs(dx) : If Abs(dy) > d Then d = Abs(dy)
				Call PoQr_Set(mods, isFn, size, a + dx, a + dy, (d <> 1))
			Next
		Next
	End If
	Call PoQr_DrawFormat(mods, isFn, size, 0)   ' cadangkan area format
End Sub

Sub PoQr_DrawFormat(mods, isFn, size, mask)
	Dim data, r, i, bits
	data = 0 * 8 + mask            ' level M = 00
	r = data
	For i = 0 To 9
		r = r * 2
		If r >= 1024 Then r = r Xor &H537
	Next
	bits = (data * 1024 + r) Xor &H5412
	For i = 0 To 5 : Call PoQr_Set(mods, isFn, size, 8, i, PoQr_Bit(bits, i)) : Next
	Call PoQr_Set(mods, isFn, size, 8, 7, PoQr_Bit(bits, 6))
	Call PoQr_Set(mods, isFn, size, 8, 8, PoQr_Bit(bits, 7))
	Call PoQr_Set(mods, isFn, size, 7, 8, PoQr_Bit(bits, 8))
	For i = 9 To 14 : Call PoQr_Set(mods, isFn, size, 14 - i, 8, PoQr_Bit(bits, i)) : Next
	For i = 0 To 7 : Call PoQr_Set(mods, isFn, size, size - 1 - i, 8, PoQr_Bit(bits, i)) : Next
	For i = 8 To 14 : Call PoQr_Set(mods, isFn, size, 8, size - 15 + i, PoQr_Bit(bits, i)) : Next
	Call PoQr_Set(mods, isFn, size, 8, size - 8, True)    ' modul gelap tetap
End Sub

Function PoQr_Bit(v, i)
	PoQr_Bit = (((v \ (2 ^ i)) Mod 2) = 1)
End Function

Sub PoQr_DrawCodewords(mods, isFn, size, cw, total)
	Dim i, rightCol, vert, j, x, y, upward, nbits
	nbits = total * 8
	i = 0
	rightCol = size - 1
	Do While rightCol >= 1
		If rightCol = 6 Then rightCol = 5
		upward = (((rightCol + 1) \ 2) Mod 2 = 0)
		For vert = 0 To size - 1
			For j = 0 To 1
				x = rightCol - j
				If upward Then y = size - 1 - vert Else y = vert
				If Not isFn(y * size + x) And i < nbits Then
					mods(y * size + x) = (((cw(i \ 8) \ (2 ^ (7 - (i Mod 8)))) Mod 2) = 1)
					i = i + 1
				End If
			Next
		Next
		rightCol = rightCol - 2
	Loop
End Sub

Sub PoQr_ApplyMask(mods, isFn, size, m)
	Dim x, y, inv, idx
	For y = 0 To size - 1
		For x = 0 To size - 1
			idx = y * size + x
			If Not isFn(idx) Then
				Select Case m
					Case 0 : inv = ((x + y) Mod 2 = 0)
					Case 1 : inv = (y Mod 2 = 0)
					Case 2 : inv = (x Mod 3 = 0)
					Case 3 : inv = ((x + y) Mod 3 = 0)
					Case 4 : inv = (((x \ 3) + (y \ 2)) Mod 2 = 0)
					Case 5 : inv = ((x * y) Mod 2 + (x * y) Mod 3 = 0)
					Case 6 : inv = ((((x * y) Mod 2) + ((x * y) Mod 3)) Mod 2 = 0)
					Case 7 : inv = ((((x + y) Mod 2) + ((x * y) Mod 3)) Mod 2 = 0)
				End Select
				If inv Then mods(idx) = Not mods(idx)
			End If
		Next
	Next
End Sub

' Penalti (aturan 1, 2, 3, 4 standar QR) untuk memilih mask
Function PoQr_Penalty(mods, size)
	Dim p, x, y, run, prev, cur, dark, total, k, a, i
	p = 0
	For y = 0 To size - 1                       ' aturan 1 - baris
		run = 0
		For x = 0 To size - 1
			cur = mods(y * size + x)
			If x > 0 And cur = prev Then
				run = run + 1
				If run = 5 Then
					p = p + 3
				ElseIf run > 5 Then
					p = p + 1
				End If
			Else
				run = 1
			End If
			prev = cur
		Next
	Next
	For x = 0 To size - 1                       ' aturan 1 - kolom
		run = 0
		For y = 0 To size - 1
			cur = mods(y * size + x)
			If y > 0 And cur = prev Then
				run = run + 1
				If run = 5 Then
					p = p + 3
				ElseIf run > 5 Then
					p = p + 1
				End If
			Else
				run = 1
			End If
			prev = cur
		Next
	Next
	For y = 0 To size - 2                       ' aturan 2 - blok 2x2
		For x = 0 To size - 2
			cur = mods(y * size + x)
			If mods(y * size + x + 1) = cur And mods((y + 1) * size + x) = cur And mods((y + 1) * size + x + 1) = cur Then p = p + 3
		Next
	Next
	For y = 0 To size - 1                       ' aturan 3 - pola mirip finder (1011101 + 4 terang)
		For x = 0 To size - 11
			If PoQr_Finderish(mods, size, x, y, 1, 0) Then p = p + 40
		Next
	Next
	For x = 0 To size - 1
		For y = 0 To size - 11
			If PoQr_Finderish(mods, size, x, y, 0, 1) Then p = p + 40
		Next
	Next
	dark = 0 : total = size * size                ' aturan 4 - keseimbangan
	For i = 0 To total - 1
		If mods(i) Then dark = dark + 1
	Next
	k = Int(Abs(dark * 20 - total * 10) / total)
	p = p + k * 10
	PoQr_Penalty = p
End Function

Function PoQr_Finderish(mods, size, x, y, dx, dy)
	Dim pat1, pat2, i, v, ok1, ok2
	pat1 = Array(1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0)
	pat2 = Array(0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1)
	ok1 = True : ok2 = True
	For i = 0 To 10
		v = 0
		If mods((y + dy * i) * size + (x + dx * i)) Then v = 1
		If v <> pat1(i) Then ok1 = False
		If v <> pat2(i) Then ok2 = False
	Next
	PoQr_Finderish = ok1 Or ok2
End Function
%>
