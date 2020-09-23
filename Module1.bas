Attribute VB_Name = "RSGF"
' The RS and Galois field code is ported from C source code not written by me.
' I wrote the high speed assembler code based on the C source.
' The C sourcecode is freeware

Option Base 0
Option Explicit

Private Const prim_poly_32 As Long = &O20000007
Private Const prim_poly_16 As Long = &O210013
Private Const prim_poly_8  As Long = &O435
Private Const prim_poly_4  As Long = &O23
Private Const prim_poly_2  As Long = &O7

Private init_done As Boolean

Private modar_w    As Long
Private modar_nw   As Long
Private modar_nwm1 As Long
Private modar_poly As Long

Private modar_m    As Long
Private modar_n    As Long
Private modar_iam  As Long

Private B_TO_J()   As Long
Private J_TO_B()   As Long

Public Type struc_Condensed_Matrix
   condensed_mat() As Long
   row_identities() As Long
End Type

Public Declare Function xorBuffer Lib "ffixlib.dll" Alias "fix_xorbuffer" (ByVal ptrorgBuffer As Long, ByVal ptrxorBuffer As Long, ByVal Size As Long) As Long
Public Declare Function mulBuffer Lib "ffixlib.dll" Alias "fix_mulbuffer" (ByVal ptrorgBuffer As Long, ByVal ptrmulBuffer As Long, ByVal ptrB2J As Long, ByVal ptrJ2B As Long, ByVal Size As Long, ByVal flog As Long) As Long
Public Declare Function empBuffer Lib "ffixlib.dll" Alias "fix_emptybuffer" (ByVal ptrbuffer As Long, ByVal Size As Long) As Long
Public Declare Function crcBuffer Lib "ffixlib.dll" Alias "fix_crc32buffer" (ByVal ptrCRC As Long, ByVal ptrbuffer As Long, ByVal ptrmtable As Long, ByVal Size As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal ptrDest As Long, ByVal ptrSrc As Long, ByVal Length As Long)

Public dataParity() As Byte
Public dataMul() As Byte
Public dataCRCt(255) As Long

' w = degree 2,4,8,16 or 32
Public Function initialize(w As Long) As Boolean
  modar_w = w
  modar_nw = 2 ^ modar_w
  modar_nwm1 = modar_nw - 1
  Select Case w
     'Case 2:  modar_poly = prim_poly_2
     'Case 4:  modar_poly = prim_poly_4
     Case 8:  modar_poly = prim_poly_8       ' supported only
     'Case 16: modar_poly = prim_poly_16
     'Case 32: modar_poly = prim_poly_32
  End Select
   
  'CRC32 Lookup table
  dataCRCt(0) = &H0
  dataCRCt(1) = &H77073096
  dataCRCt(2) = &HEE0E612C
  dataCRCt(3) = &H990951BA
  dataCRCt(4) = &H76DC419
  dataCRCt(5) = &H706AF48F
  dataCRCt(6) = &HE963A535
  dataCRCt(7) = &H9E6495A3
  dataCRCt(8) = &HEDB8832
  dataCRCt(9) = &H79DCB8A4
  dataCRCt(10) = &HE0D5E91E
  dataCRCt(11) = &H97D2D988
  dataCRCt(12) = &H9B64C2B
  dataCRCt(13) = &H7EB17CBD
  dataCRCt(14) = &HE7B82D07
  dataCRCt(15) = &H90BF1D91
  dataCRCt(16) = &H1DB71064
  dataCRCt(17) = &H6AB020F2
  dataCRCt(18) = &HF3B97148
  dataCRCt(19) = &H84BE41DE
  dataCRCt(20) = &H1ADAD47D
  dataCRCt(21) = &H6DDDE4EB
  dataCRCt(22) = &HF4D4B551
  dataCRCt(23) = &H83D385C7
  dataCRCt(24) = &H136C9856
  dataCRCt(25) = &H646BA8C0
  dataCRCt(26) = &HFD62F97A
  dataCRCt(27) = &H8A65C9EC
  dataCRCt(28) = &H14015C4F
  dataCRCt(29) = &H63066CD9
  dataCRCt(30) = &HFA0F3D63
  dataCRCt(31) = &H8D080DF5
  dataCRCt(32) = &H3B6E20C8
  dataCRCt(33) = &H4C69105E
  dataCRCt(34) = &HD56041E4
  dataCRCt(35) = &HA2677172
  dataCRCt(36) = &H3C03E4D1
  dataCRCt(37) = &H4B04D447
  dataCRCt(38) = &HD20D85FD
  dataCRCt(39) = &HA50AB56B
  dataCRCt(40) = &H35B5A8FA
  dataCRCt(41) = &H42B2986C
  dataCRCt(42) = &HDBBBC9D6
  dataCRCt(43) = &HACBCF940
  dataCRCt(44) = &H32D86CE3
  dataCRCt(45) = &H45DF5C75
  dataCRCt(46) = &HDCD60DCF
  dataCRCt(47) = &HABD13D59
  dataCRCt(48) = &H26D930AC
  dataCRCt(49) = &H51DE003A
  dataCRCt(50) = &HC8D75180
  dataCRCt(51) = &HBFD06116
  dataCRCt(52) = &H21B4F4B5
  dataCRCt(53) = &H56B3C423
  dataCRCt(54) = &HCFBA9599
  dataCRCt(55) = &HB8BDA50F
  dataCRCt(56) = &H2802B89E
  dataCRCt(57) = &H5F058808
  dataCRCt(58) = &HC60CD9B2
  dataCRCt(59) = &HB10BE924
  dataCRCt(60) = &H2F6F7C87
  dataCRCt(61) = &H58684C11
  dataCRCt(62) = &HC1611DAB
  dataCRCt(63) = &HB6662D3D
  dataCRCt(64) = &H76DC4190
  dataCRCt(65) = &H1DB7106
  dataCRCt(66) = &H98D220BC
  dataCRCt(67) = &HEFD5102A
  dataCRCt(68) = &H71B18589
  dataCRCt(69) = &H6B6B51F
  dataCRCt(70) = &H9FBFE4A5
  dataCRCt(71) = &HE8B8D433
  dataCRCt(72) = &H7807C9A2
  dataCRCt(73) = &HF00F934
  dataCRCt(74) = &H9609A88E
  dataCRCt(75) = &HE10E9818
  dataCRCt(76) = &H7F6A0DBB
  dataCRCt(77) = &H86D3D2D
  dataCRCt(78) = &H91646C97
  dataCRCt(79) = &HE6635C01
  dataCRCt(80) = &H6B6B51F4
  dataCRCt(81) = &H1C6C6162
  dataCRCt(82) = &H856530D8
  dataCRCt(83) = &HF262004E
  dataCRCt(84) = &H6C0695ED
  dataCRCt(85) = &H1B01A57B
  dataCRCt(86) = &H8208F4C1
  dataCRCt(87) = &HF50FC457
  dataCRCt(88) = &H65B0D9C6
  dataCRCt(89) = &H12B7E950
  dataCRCt(90) = &H8BBEB8EA
  dataCRCt(91) = &HFCB9887C
  dataCRCt(92) = &H62DD1DDF
  dataCRCt(93) = &H15DA2D49
  dataCRCt(94) = &H8CD37CF3
  dataCRCt(95) = &HFBD44C65
  dataCRCt(96) = &H4DB26158
  dataCRCt(97) = &H3AB551CE
  dataCRCt(98) = &HA3BC0074
  dataCRCt(99) = &HD4BB30E2
  dataCRCt(100) = &H4ADFA541
  dataCRCt(101) = &H3DD895D7
  dataCRCt(102) = &HA4D1C46D
  dataCRCt(103) = &HD3D6F4FB
  dataCRCt(104) = &H4369E96A
  dataCRCt(105) = &H346ED9FC
  dataCRCt(106) = &HAD678846
  dataCRCt(107) = &HDA60B8D0
  dataCRCt(108) = &H44042D73
  dataCRCt(109) = &H33031DE5
  dataCRCt(110) = &HAA0A4C5F
  dataCRCt(111) = &HDD0D7CC9
  dataCRCt(112) = &H5005713C
  dataCRCt(113) = &H270241AA
  dataCRCt(114) = &HBE0B1010
  dataCRCt(115) = &HC90C2086
  dataCRCt(116) = &H5768B525
  dataCRCt(117) = &H206F85B3
  dataCRCt(118) = &HB966D409
  dataCRCt(119) = &HCE61E49F
  dataCRCt(120) = &H5EDEF90E
  dataCRCt(121) = &H29D9C998
  dataCRCt(122) = &HB0D09822
  dataCRCt(123) = &HC7D7A8B4
  dataCRCt(124) = &H59B33D17
  dataCRCt(125) = &H2EB40D81
  dataCRCt(126) = &HB7BD5C3B
  dataCRCt(127) = &HC0BA6CAD
  dataCRCt(128) = &HEDB88320
  dataCRCt(129) = &H9ABFB3B6
  dataCRCt(130) = &H3B6E20C
  dataCRCt(131) = &H74B1D29A
  dataCRCt(132) = &HEAD54739
  dataCRCt(133) = &H9DD277AF
  dataCRCt(134) = &H4DB2615
  dataCRCt(135) = &H73DC1683
  dataCRCt(136) = &HE3630B12
  dataCRCt(137) = &H94643B84
  dataCRCt(138) = &HD6D6A3E
  dataCRCt(139) = &H7A6A5AA8
  dataCRCt(140) = &HE40ECF0B
  dataCRCt(141) = &H9309FF9D
  dataCRCt(142) = &HA00AE27
  dataCRCt(143) = &H7D079EB1
  dataCRCt(144) = &HF00F9344
  dataCRCt(145) = &H8708A3D2
  dataCRCt(146) = &H1E01F268
  dataCRCt(147) = &H6906C2FE
  dataCRCt(148) = &HF762575D
  dataCRCt(149) = &H806567CB
  dataCRCt(150) = &H196C3671
  dataCRCt(151) = &H6E6B06E7
  dataCRCt(152) = &HFED41B76
  dataCRCt(153) = &H89D32BE0
  dataCRCt(154) = &H10DA7A5A
  dataCRCt(155) = &H67DD4ACC
  dataCRCt(156) = &HF9B9DF6F
  dataCRCt(157) = &H8EBEEFF9
  dataCRCt(158) = &H17B7BE43
  dataCRCt(159) = &H60B08ED5
  dataCRCt(160) = &HD6D6A3E8
  dataCRCt(161) = &HA1D1937E
  dataCRCt(162) = &H38D8C2C4
  dataCRCt(163) = &H4FDFF252
  dataCRCt(164) = &HD1BB67F1
  dataCRCt(165) = &HA6BC5767
  dataCRCt(166) = &H3FB506DD
  dataCRCt(167) = &H48B2364B
  dataCRCt(168) = &HD80D2BDA
  dataCRCt(169) = &HAF0A1B4C
  dataCRCt(170) = &H36034AF6
  dataCRCt(171) = &H41047A60
  dataCRCt(172) = &HDF60EFC3
  dataCRCt(173) = &HA867DF55
  dataCRCt(174) = &H316E8EEF
  dataCRCt(175) = &H4669BE79
  dataCRCt(176) = &HCB61B38C
  dataCRCt(177) = &HBC66831A
  dataCRCt(178) = &H256FD2A0
  dataCRCt(179) = &H5268E236
  dataCRCt(180) = &HCC0C7795
  dataCRCt(181) = &HBB0B4703
  dataCRCt(182) = &H220216B9
  dataCRCt(183) = &H5505262F
  dataCRCt(184) = &HC5BA3BBE
  dataCRCt(185) = &HB2BD0B28
  dataCRCt(186) = &H2BB45A92
  dataCRCt(187) = &H5CB36A04
  dataCRCt(188) = &HC2D7FFA7
  dataCRCt(189) = &HB5D0CF31
  dataCRCt(190) = &H2CD99E8B
  dataCRCt(191) = &H5BDEAE1D
  dataCRCt(192) = &H9B64C2B0
  dataCRCt(193) = &HEC63F226
  dataCRCt(194) = &H756AA39C
  dataCRCt(195) = &H26D930A
  dataCRCt(196) = &H9C0906A9
  dataCRCt(197) = &HEB0E363F
  dataCRCt(198) = &H72076785
  dataCRCt(199) = &H5005713
  dataCRCt(200) = &H95BF4A82
  dataCRCt(201) = &HE2B87A14
  dataCRCt(202) = &H7BB12BAE
  dataCRCt(203) = &HCB61B38
  dataCRCt(204) = &H92D28E9B
  dataCRCt(205) = &HE5D5BE0D
  dataCRCt(206) = &H7CDCEFB7
  dataCRCt(207) = &HBDBDF21
  dataCRCt(208) = &H86D3D2D4
  dataCRCt(209) = &HF1D4E242
  dataCRCt(210) = &H68DDB3F8
  dataCRCt(211) = &H1FDA836E
  dataCRCt(212) = &H81BE16CD
  dataCRCt(213) = &HF6B9265B
  dataCRCt(214) = &H6FB077E1
  dataCRCt(215) = &H18B74777
  dataCRCt(216) = &H88085AE6
  dataCRCt(217) = &HFF0F6A70
  dataCRCt(218) = &H66063BCA
  dataCRCt(219) = &H11010B5C
  dataCRCt(220) = &H8F659EFF
  dataCRCt(221) = &HF862AE69
  dataCRCt(222) = &H616BFFD3
  dataCRCt(223) = &H166CCF45
  dataCRCt(224) = &HA00AE278
  dataCRCt(225) = &HD70DD2EE
  dataCRCt(226) = &H4E048354
  dataCRCt(227) = &H3903B3C2
  dataCRCt(228) = &HA7672661
  dataCRCt(229) = &HD06016F7
  dataCRCt(230) = &H4969474D
  dataCRCt(231) = &H3E6E77DB
  dataCRCt(232) = &HAED16A4A
  dataCRCt(233) = &HD9D65ADC
  dataCRCt(234) = &H40DF0B66
  dataCRCt(235) = &H37D83BF0
  dataCRCt(236) = &HA9BCAE53
  dataCRCt(237) = &HDEBB9EC5
  dataCRCt(238) = &H47B2CF7F
  dataCRCt(239) = &H30B5FFE9
  dataCRCt(240) = &HBDBDF21C
  dataCRCt(241) = &HCABAC28A
  dataCRCt(242) = &H53B39330
  dataCRCt(243) = &H24B4A3A6
  dataCRCt(244) = &HBAD03605
  dataCRCt(245) = &HCDD70693
  dataCRCt(246) = &H54DE5729
  dataCRCt(247) = &H23D967BF
  dataCRCt(248) = &HB3667A2E
  dataCRCt(249) = &HC4614AB8
  dataCRCt(250) = &H5D681B02
  dataCRCt(251) = &H2A6F2B94
  dataCRCt(252) = &HB40BBE37
  dataCRCt(253) = &HC30C8EA1
  dataCRCt(254) = &H5A05DF1B
  dataCRCt(255) = &H2D02EF8D
  
  initialize = gf_modar_setup()
End Function

Private Function gf_modar_setup() As Boolean
   Dim j As Long
   Dim b As Long
   ReDim B_TO_J(modar_nwm1)
   ReDim J_TO_B(modar_nwm1 * 2 + 1)
   For j = 0 To modar_nwm1
      B_TO_J(j) = -1
      J_TO_B(j) = 0
      J_TO_B(j + modar_nwm1 + 1) = 0
   Next
   b = 1
   For j = 0 To modar_nwm1 - 1
      If B_TO_J(b) <> -1 Then Exit Function
      B_TO_J(b) = j
      J_TO_B(j) = b
      J_TO_B(j + modar_nwm1) = b
      b = b + b
      If (b And modar_nw) Then b = (b Xor modar_poly) And modar_nwm1
   Next
   init_done = True
   gf_modar_setup = True
End Function

Public Function gf_single_multiply(a As Long, b As Long) As Long
   Dim j As Long
   If a = 0 Or b = 0 Then Exit Function
   j = B_TO_J(a) + B_TO_J(b)
   If j >= modar_nwm1 Then j = j - modar_nwm1
   gf_single_multiply = J_TO_B(j)
End Function

Public Function gf_single_divide(a As Long, b As Long) As Long
   Dim j As Long
   If b = 0 Then gf_single_divide = -1: Exit Function
   If a = 0 Then gf_single_divide = 0:  Exit Function
   j = B_TO_J(a) - B_TO_J(b)
   If j < 0 Then j = j + modar_nwm1
   gf_single_divide = J_TO_B(j)
End Function

Public Function gf_mult_region(ByVal Size As Long, ByVal factor As Long, Optional ByVal Offset As Long = 0, Optional ByVal Cumulative As Boolean = False)
   'handle mul by 0 or 1
   If factor = 1 Then
      If Not Cumulative Then
         CopyMemory VarPtr(dataMul(LBound(dataMul))), _
                    VarPtr(dataBuffer(LBound(dataBuffer) + Offset)), _
                    Size
      End If
      Exit Function
   End If
   If factor = 0 Then
      empBuffer VarPtr(dataMul(LBound(dataMul))), Size
      Exit Function
   End If
   mulBuffer IIf(Cumulative, VarPtr(dataMul(LBound(dataMul))), VarPtr(dataBuffer(LBound(dataBuffer) + Offset))), _
             VarPtr(dataMul(LBound(dataMul))), _
             VarPtr(B_TO_J(LBound(B_TO_J))), _
             VarPtr(J_TO_B(LBound(J_TO_B))), _
             ByVal Size, _
             ByVal B_TO_J(factor)
End Function

Public Function gf_add_parity(ByVal Size As Long)
   xorBuffer VarPtr(dataMul(LBound(dataMul))), _
             VarPtr(dataParity(LBound(dataParity))), _
             ByVal Size
End Function

'/* This returns the rows*cols vandermonde matrix.  N+M must be
'   < 2^w -1.  Row 0 is in elements 0 to cols-1.  Row one is
'   in elements cols to 2cols-1.  Etc.*/

Public Function gf_make_vandermonde(r As Long, C As Long) As Long()
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim vdm() As Long
  If r >= modar_nwm1 Or C >= modar_nwm1 Then Exit Function
  ReDim vdm(r - 1, C - 1) As Long
  For i = 0 To r - 1
     k = 1
     For j = 0 To C - 1
       vdm(i, j) = k
       k = gf_single_multiply(k, i)
     Next
  Next
  gf_make_vandermonde = vdm()
End Function

Public Function find_swap_row(m() As Long, r As Long, C As Long, n As Long)
  Dim j As Long
  For j = n To r - 1
     If m(j, n) <> 0 Then
        find_swap_row = j
        Exit Function
     End If
  Next
  find_swap_row = -1
End Function

Public Function gf_make_dispersal_matrix(r As Long, C As Long) As Long()
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  Dim inv As Long
  Dim tmp As Long
  Dim vdm() As Long
  Dim coli As Long
  vdm() = gf_make_vandermonde(r, C)
  i = 0
  Do While i < C And i < r
     j = find_swap_row(vdm(), r, C, i)
     If j = -1 Then
'        Debug.Print "Error: make_dispersal_matrix.  Can't find swap row %d\n"
        Exit Function
     End If
     If j <> i Then
        For k = 0 To C - 1
           tmp = vdm(j, k)
           vdm(j, k) = vdm(i, k)
           vdm(i, k) = tmp
        Next
     End If
     If vdm(i, i) = 0 Then
'        Debug.Print "Internal error -- this shouldn't happen\n"
        Exit Function
     End If
     If vdm(i, i) <> 1 Then
        inv = gf_single_divide(1, vdm(i, i))
        For j = 0 To r - 1
          vdm(j, i) = gf_single_multiply(inv, vdm(j, i))
        Next
     End If
     If vdm(i, i) <> 1 Then
'         Debug.Print "Internal error -- this shouldn't happen #2\n"
        Exit Function
     End If
     For j = 0 To C - 1
        coli = vdm(i, j)
        If j <> i And coli <> 0 Then
           For l = 0 To r - 1
              vdm(l, j) = vdm(l, j) Xor gf_single_multiply(coli, vdm(l, i))
           Next
        End If
     Next
     i = i + 1
  Loop
  gf_make_dispersal_matrix = vdm()
End Function

Public Function gf_condense_dispersal_matrix(disp() As Long, existing_rows() As Long, r As Long, C As Long) As struc_Condensed_Matrix
  Dim i   As Long
  Dim j   As Long
  Dim k   As Long
  Dim tmp As Long
  Dim cm  As struc_Condensed_Matrix
  ReDim cm.condensed_mat(C - 1, C - 1)
  ReDim cm.row_identities(C - 1)
  For i = 0 To C - 1
    cm.row_identities(i) = -1
    If existing_rows(i) <> 0 Then
       cm.row_identities(i) = i
       For j = 0 To C - 1
          cm.condensed_mat(i, j) = disp(i, j)
       Next
    End If
  Next
  k = 0
  For i = C To r - 1
    If existing_rows(i) <> 0 Then
       Do While k < C And cm.row_identities(k) <> -1
         k = k + 1
         If k >= C Then Exit Do
       Loop
       If k = C Then
         gf_condense_dispersal_matrix = cm
         Exit Function
       End If
       cm.row_identities(k) = i
       For j = 0 To C - 1
         cm.condensed_mat(k, j) = disp(i, j)
       Next
    End If
  Next
  Do While k < C And cm.row_identities(k) <> -1
    k = k + 1
    If k >= C Then Exit Do
  Loop
  If k = C Then
    gf_condense_dispersal_matrix = cm
    Exit Function
  End If
End Function


Public Function gf_invert_matrix(m() As Long, r As Long) As Long()
  Dim inv() As Long
  Dim copy() As Long
  Dim C As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim X As Long
  Dim rs2 As Long
  Dim row_start As Long
  Dim tmp As Long
  Dim inverse As Long
  C = r
  ReDim inv(r - 1, C - 1)
  ReDim copy(r - 1, C - 1)
  For i = 0 To r - 1
     For j = 0 To C - 1
        inv(i, j) = IIf(i = j, 1, 0)
        copy(i, j) = m(i, j)
     Next
  Next
  For i = 0 To C - 1
     
     If copy(row_start, i) = 0 Then
        For j = i + 1 To r - 1
          If copy(j, i) <> 0 Then Exit For
        Next
        If j = r Then Exit Function
        rs2 = j
        For k = 0 To C - 1
          tmp = copy(row_start, k)
          copy(row_start, k) = copy(rs2, k)
          copy(rs2, k) = tmp
          tmp = inv(row_start, k)
          inv(row_start, k) = inv(rs2, k)
          inv(rs2, k) = tmp
        Next
     End If
     tmp = copy(row_start, i)
     If tmp <> 1 Then
        inverse = gf_single_divide(1, tmp)
        For j = 0 To C - 1
           copy(row_start, j) = gf_single_multiply(copy(row_start, j), inverse)
           inv(row_start, j) = gf_single_multiply(inv(row_start, j), inverse)
        Next
     End If
     k = row_start + 1
     For j = i + 1 To C - 1
        If copy(k, i) <> 0 Then
           If copy(k, i) = 1 Then
              rs2 = j
              For X = 0 To C - 1
                 copy(rs2, X) = copy(rs2, X) Xor copy(row_start, X)
                 inv(rs2, X) = inv(rs2, X) Xor inv(row_start, X)
              Next
           Else
              tmp = copy(k, i)
              rs2 = j
              For X = 0 To C - 1
                 copy(rs2, X) = copy(rs2, X) Xor gf_single_multiply(tmp, copy(row_start, X))
                 inv(rs2, X) = inv(rs2, X) Xor gf_single_multiply(tmp, inv(row_start, X))
              Next
           End If
        End If
        k = k + 1
     Next
     row_start = row_start + 1
   Next
   For i = r - 1 To 0 Step -1
     row_start = i
     For j = 0 To i - 1
        rs2 = j
        If copy(rs2, i) <> 0 Then
           tmp = copy(rs2, i)
           copy(rs2, i) = 0
           For k = 0 To C - 1
              inv(rs2, k) = inv(rs2, k) Xor gf_single_multiply(tmp, inv(row_start, k))
           Next
        End If
     Next
   Next
   ReDim copy(0)
   gf_invert_matrix = inv
End Function
