Attribute VB_Name = "Deflate"
Option Explicit

Private Type HashSlotType
    Char As Byte
    ToHash As Long
    ToSlot As Long
End Type
Private Type HashTableType
    FirstSlot As Long
    LastSlot As Long
    Pointer As Long
    NumSlots As Integer
End Type
Private Type SlotType
    Next As Long
    Prev As Long
    Winpos As Long
End Type
Private Type DataType
    data() As Byte
    pos As Long
    Bit As Long
End Type
Private Type CodesType
    Code As Long                            'The huffman codes
    Lenght As Long                          'Lenght of the huffman codes
End Type


Private Const MinMatch As Integer = 3
Private MaxMatch As Long
Private MaxHistore As Long
Private NxtFreeSlot As Long
Private NxtFreeHashSlot1 As Long
Private NxtFreeHashSlot2 As Long
Private NxtFreeHash As Long
Private Hash1() As Long
Private HashSlot1() As HashSlotType
Private hashSlot2() As HashSlotType
Private Hash() As HashTableType
Private Slot() As SlotType
Private Udata As DataType               'Output data
Private IData As DataType               'Input data
Private GoodMatch As Integer
Private LazyMatch As Integer
Private NiceMatch As Integer
Private MaxChain As Integer
Private MatchL As Long
Private MatchP As Long
Private FileLenght As Long
'Initiations for the huffman trees
Private LL(287) As Long                     'counter for literal code
Private DC(31) As Long                      'Distance code counter
Private Lens(31) As Long                    'Lenght codes
Private Lext(31) As Long                    'Lenght extra bits
Private Dist(31) As Long                    'Distance codes
Private Dext(31) As Long                    'Distance extra bits
Private Border(18) As Long                  'Store order for headerlenghts
Private LitLen(287) As CodesType            'Literal and lenght codes holder
Private Distance(31) As CodesType           'Distance codes holder
Private Bl_Tree() As CodesType              'Huffman codes for lenght/distance table
Private Max_BL As Long                      'Maximum block lenght
Private MaxLitLen As Long                   'Maximum Literal lenght code
Private MaxDist As Long                     'Maximum distance code
Private CBP As Long                         'Current bitposition
Private CurVal As Long                      'Current value
Private NumDist As Long                     'Number of distance codes to use

Public Sub Deflate(ByteArray() As Byte, CompLevel As Integer, Optional Deflate64 As Boolean = False)
    Dim HoldPos As Long
    Dim LitCount As Long
    Dim PrevMatch As Integer
    Dim PrevPos As Long
    Dim Char As Integer
    Dim BestStore As Long           'Lenght if stored
    Dim BestStatic As Long          'Lenght is static tree used
    Dim BestDynamic As Long         'Lenght as dynamic tree used
    Dim Bestcase As Long            'Wich case is best
    Dim Header As Byte              'header byte for already compressed data
    Dim tb As Byte                  'Bytevalue to be stored as huffman code
    Dim StLen As Long
    Dim X As Long
    Dim FileSize As Long
    Call INIT(CompLevel, Deflate64)
    Call init_Trees
    FileLenght = 0
    On Error Resume Next            'in case bytearray has no data
    FileSize = UBound(ByteArray) + 1
    Err.Clear
    On Error GoTo 0
    FileLenght = FileSize - 1
    IData.data = ByteArray
    Erase ByteArray                 'to clear memory
    IData.pos = 0
    If CompLevel > 0 And FileSize > 0 Then
        Do While IData.pos + MinMatch <= FileLenght
'search back until a good match is found
            Call FindLongestMatch(IData.pos, GoodMatch, True)
'if match is long enough but could maybe be longer, search for longer match
            If MatchL < LazyMatch And MatchL >= MinMatch Then
                PrevMatch = MatchL
                PrevPos = MatchP
                Call FindLongestMatch(IData.pos, LazyMatch, False)
                If MatchL <= PrevMatch Then
                    MatchL = PrevMatch
                    MatchP = PrevPos
                End If
            End If
'if match is long enough, search for better match at next position
            If CompLevel > 3 And MatchL >= MinMatch And MatchL < NiceMatch Then
                PrevMatch = MatchL
                PrevPos = MatchP
                Call FindLongestMatch(IData.pos + 1, NiceMatch, True)
                If MatchL <= PrevMatch Then
                    MatchL = PrevMatch
                    MatchP = PrevPos
                Else
'if better match found, store last byte as literal and store match at next position
                    Char = IData.data(IData.pos)
                    If LitCount = 127 Then Udata.data(HoldPos) = LitCount: LitCount = 0
                    If LitCount = 0 Then HoldPos = PutByte(0)
                    Call InsertSlot(IData.pos)
                    Call PutByte(Char)
                    Call AddLit(Char)
                    LitCount = LitCount + 1
                    IData.pos = IData.pos + 1
                End If
            End If
            If MatchL = MinMatch And IData.pos - MatchP > 4096 Then MatchL = 1
            If MatchL < MinMatch Then
                Char = IData.data(IData.pos)
                If LitCount = 127 Then Udata.data(HoldPos) = LitCount: LitCount = 0
                If LitCount = 0 Then HoldPos = PutByte(0)
                Call InsertSlot(IData.pos)                  'preserve for next searches
                Call PutByte(Char)                          'store literal
                Call AddLit(Char)                           'add to counter
                LitCount = LitCount + 1
                IData.pos = IData.pos + 1
            Else
                If LitCount > 0 Then Udata.data(HoldPos) = LitCount: LitCount = 0
                PutByte 128         'set byte >127 so treebuilder will see a length/distance
                MatchP = IData.pos - MatchP
                Call PutByte(Fix(MatchP / 256))             'store distance value
                Call PutByte(MatchP And 255)
                Call AddDis(MatchP)                         'add to counter
                If MatchL > 257 And Deflate64 Then
                    Call PutByte(258 - 3)                   'send that bigger lenght is comming
                    Call PutByte(Fix((MatchL - 3) / 256))   'send bigger lenght
                    Call PutByte((MatchL - 3) And 255)
                    Call AddLen(258)                        'add to counter
                Else
                    Call PutByte(MatchL - 3)                'store lenght byte
                    Call AddLen(MatchL)                     'add to counter
                End If
                Do While MatchL > 0
                    Call InsertSlot(IData.pos)              'preserve for next searches
                    If IData.pos = FileLenght - 2 Then IData.pos = IData.pos + MatchL: Exit Do
                    IData.pos = IData.pos + 1
                    MatchL = MatchL - 1
                Loop
            End If
        Loop
        If IData.pos <= FileLenght Then
            For X = IData.pos To FileLenght
                If LitCount = 127 Then Udata.data(HoldPos) = LitCount: LitCount = 0
                If LitCount = 0 Then HoldPos = PutByte(0)
                Call PutByte(CInt(IData.data(X)))
                Call AddLit(CInt(IData.data(X)))
                LitCount = LitCount + 1
            Next
        End If
        If LitCount > 0 Then Udata.data(HoldPos) = LitCount
    End If
'send EOF code
    Call PutByte(0)
    LL(256) = 1         'update the counter to complete the huffman table
'redim to the correct bounderies
    ReDim Preserve Udata.data(Udata.pos - 1)
    If CompLevel <> 0 Then
        BestStore = FileSize + 5
        BestStatic = Calc_Static
        BestDynamic = Calc_Dynamic
        Bestcase = 0
        If BestStatic < BestStore Then Bestcase = 1
        If Bestcase = 1 Then
            If BestDynamic < BestStatic Then Bestcase = 2
        Else
            If BestDynamic < BestStore Then Bestcase = 2
        End If
    Else
        Bestcase = 0
    End If
    If Bestcase = 0 Then
        Udata.pos = 0
        Udata.Bit = 0       'don't need this anymore
        Call putbits(1, 1)      'set last block type
        Call putbits(Bestcase, 2)      'Set compression type
'Store without compression
        Do While CBP <> 0       'line up till next byte
            Call putbits(0, 1)
        Loop
        StLen = FileLenght + 1
        Call PutByte(StLen And &HFF)
        Call PutByte(Fix(StLen / 256) And &HFF)
        StLen = (Not StLen) And &HFFFF&
        Call PutByte(StLen And &HFF)
        Call PutByte(Fix(StLen / 256) And &HFF)
        For X = 0 To UBound(IData.data)
            PutByte (IData.data(X))
        Next
    Else
        IData.data = Udata.data      'Put compressed stream back so whe can use it again
        IData.pos = 0
        Udata.pos = 0              'reset the output buffer but leaf data in it
        Call putbits(1, 1)      'set last block type
        Call putbits(Bestcase, 2)      'Set compression type
        If Bestcase = 1 Then
'Whe only created a dynamic tree but where using a static tree so first whe
'have to create a tree for it
            Erase LitLen
            Erase Distance
            For X = 0 To 143: LL(X) = 8: Next
            For X = 144 To 255: LL(X) = 9: Next
            For X = 256 To 279: LL(X) = 7: Next
            For X = 280 To 287: LL(X) = 8: Next
            For X = 0 To 31: DC(X) = 5: Next
            Call Create_Codes(LitLen, LL, 287)
            Call Create_Codes(Distance, DC, 31)
            NumDist = 31
        Else
'Now whe have to store the dynamic tree first
            Call putbits(MaxLitLen + 1 - 257, 5)
            Call putbits(MaxDist + 1 - 1, 5)
            Call putbits(Max_BL + 1 - 4, 4)
            For X = 0 To Max_BL
                Call putbits(Bl_Tree(Border(X)).Lenght, 3)
            Next
            Call Send_Tree(LitLen, MaxLitLen)
            Call Send_Tree(Distance, MaxDist)
        End If
        Do
            Header = IData.data(IData.pos)
            IData.pos = IData.pos + 1
            If Header < 128 Then
                If Header = 0 Then
                    Call PutBitsRev(LitLen(256).Code, LitLen(256).Lenght)
                    Exit Do                 'ready
                End If
                For X = 1 To Header
                    tb = IData.data(IData.pos)
                    IData.pos = IData.pos + 1
                    Call PutBitsRev(LitLen(tb).Code, LitLen(tb).Lenght)
                Next
            Else                        'Length / distance pairs
                MatchP = CLng(IData.data(IData.pos)) * 256 + IData.data(IData.pos + 1)
                IData.pos = IData.pos + 2
                MatchL = IData.data(IData.pos) + 3
                IData.pos = IData.pos + 1
'store lenght value and extra bits value
                X = 28
                Do While Lens(X) > MatchL
                    X = X - 1
                Loop
                Call PutBitsRev(LitLen(X + 257).Code, LitLen(X + 257).Lenght)
                Call putbits(MatchL - Lens(X), Lext(X))
'Store lenght as deflate 64 is used
                If MatchL = 258 And Deflate64 Then
                    MatchL = CLng(IData.data(IData.pos)) * 256 + IData.data(IData.pos + 1)
                    IData.pos = IData.pos + 2
                    Call putbits(MatchL, 16)
                End If
'store distance value and extra bits value
                X = NumDist
                Do While Dist(X) > MatchP
                    X = X - 1
                Loop
                Call PutBitsRev(Distance(X).Code, Distance(X).Lenght)
                Call putbits(MatchP - Dist(X), Dext(X))
            End If
        Loop
        Call Flush_Buffer                   'save the last bits
    End If
    ReDim Preserve Udata.data(Udata.pos - 1)
    ByteArray = Udata.data
    Erase Udata.data
    Erase Hash
    Erase Hash1
    Erase HashSlot1
    Erase hashSlot2
    Erase Slot
End Sub

'this sub is used to find starting positions for the chars to find
'it will return the starting adress of the match wich is good enough
Private Sub FindLongestMatch(Winpos As Long, FindMatch As Integer, NewSearch As Boolean)
    Dim HVal As Long
    Dim X As Integer
    Dim STpos As Long
    Dim Slen As Long
    MatchL = MinMatch - 1
'Check if the next chars are already in the hashtree
    HVal = Hash1(IData.data(Winpos))
    If HVal = 0 Then Exit Sub
    Do
        If HashSlot1(HVal).ToHash = 0 Then Exit Sub
        If HashSlot1(HVal).Char = IData.data(Winpos + 1) Then
            HVal = HashSlot1(HVal).ToSlot
            Exit Do
        End If
        HVal = HashSlot1(HVal).ToHash
    Loop
    Do
        If hashSlot2(HVal).ToHash = 0 Then Exit Sub
        If hashSlot2(HVal).Char = IData.data(Winpos + 2) Then
            HVal = hashSlot2(HVal).ToSlot
            Exit Do
        End If
        HVal = hashSlot2(HVal).ToHash
    Loop
'chars are in the hashtree so find the longest match
    If NewSearch Then Hash(HVal).Pointer = Hash(HVal).FirstSlot
    For X = 1 To Hash(HVal).NumSlots
        If Hash(HVal).Pointer < 1 Then Exit Sub
        STpos = Slot(Hash(HVal).Pointer).Winpos
        If Winpos - STpos > MaxHistore Then Exit Sub
        Slen = 2        'First 3 are determed by the HashSlots
        Do While IData.data(STpos + Slen) = IData.data(Winpos + Slen)
            If Winpos + Slen = FileLenght Then Exit Do
            If Slen = MaxMatch Then Exit Do
            Slen = Slen + 1
        Loop
        Hash(HVal).Pointer = Slot(Hash(HVal).Pointer).Next
        If MatchL < Slen Then MatchL = Slen: MatchP = STpos
        If MatchL >= FindMatch Then Exit Sub
    Next
End Sub

'this sub is used to store starting positions of strings to search for
Private Sub InsertSlot(Winpos As Long)
    Dim HVal As Long
    Dim NFS As Long
    Dim NuSlot As Long
'Calculate or create a hashposition where to find the chars
    HVal = Hash1(IData.data(Winpos))
    If HVal = 0 Then
        HVal = GetFreeHashSlot1
        Hash1(IData.data(Winpos)) = HVal
    End If
    Do
        If HashSlot1(HVal).ToHash = 0 Then
            NFS = GetFreeHashSlot1
            HashSlot1(HVal).ToHash = NFS
            NFS = GetFreeHashSlot2
            HashSlot1(HVal).ToSlot = NFS
            HashSlot1(HVal).Char = IData.data(Winpos + 1)
        End If
        If HashSlot1(HVal).Char = IData.data(Winpos + 1) Then
            HVal = HashSlot1(HVal).ToSlot
            Exit Do
        End If
        HVal = HashSlot1(HVal).ToHash
    Loop
    Do
        If hashSlot2(HVal).ToHash = 0 Then
            NFS = GetFreeHashSlot2
            hashSlot2(HVal).ToHash = NFS
            NFS = GetFreeHash
            hashSlot2(HVal).ToSlot = NFS
            hashSlot2(HVal).Char = IData.data(Winpos + 2)
        End If
        If hashSlot2(HVal).Char = IData.data(Winpos + 2) Then
            HVal = hashSlot2(HVal).ToSlot
            Exit Do
        End If
        HVal = hashSlot2(HVal).ToHash
    Loop
'Position found so put it in the hashtable
    If Hash(HVal).NumSlots = 0 Then             'Insert the first slot
        NFS = GetFreeSlot
        Hash(HVal).FirstSlot = NFS
        Hash(HVal).LastSlot = NFS
        Hash(HVal).NumSlots = 1
        Slot(NFS).Winpos = Winpos
        Slot(NFS).Next = -1
        Exit Sub
    End If
    If Hash(HVal).NumSlots = MaxChain Then      'Rotate the slots
        NFS = Hash(HVal).LastSlot
        Hash(HVal).LastSlot = Slot(NFS).Prev
        Slot(Hash(HVal).LastSlot).Next = -1
        NuSlot = Hash(HVal).FirstSlot
        Slot(NuSlot).Prev = NFS
        Slot(NFS).Prev = -1
        Slot(NFS).Next = NuSlot
        Slot(NFS).Winpos = Winpos
        Hash(HVal).FirstSlot = NFS
    Else                                        'Add a new slot
        NuSlot = Hash(HVal).FirstSlot
        NFS = GetFreeSlot
        Hash(HVal).FirstSlot = NFS
        Slot(NFS).Next = NuSlot
        Slot(NFS).Winpos = Winpos
        Slot(NuSlot).Prev = NFS
        Hash(HVal).NumSlots = Hash(HVal).NumSlots + 1
    End If
End Sub

Private Function GetFreeSlot()
    NxtFreeSlot = NxtFreeSlot + 1
    If NxtFreeSlot > UBound(Slot) Then ReDim Preserve Slot(NxtFreeSlot + 1000)
    GetFreeSlot = NxtFreeSlot
End Function

Private Function GetFreeHash()
    NxtFreeHash = NxtFreeHash + 1
    If NxtFreeHash > UBound(Hash) Then ReDim Preserve Hash(NxtFreeHash + 1000)
    GetFreeHash = NxtFreeHash
End Function

Private Function GetFreeHashSlot1()
    NxtFreeHashSlot1 = NxtFreeHashSlot1 + 1
    If NxtFreeHashSlot1 > UBound(HashSlot1) Then ReDim Preserve HashSlot1(NxtFreeHashSlot1 + 1000)
    GetFreeHashSlot1 = NxtFreeHashSlot1
End Function

Private Function GetFreeHashSlot2()
    NxtFreeHashSlot2 = NxtFreeHashSlot2 + 1
    If NxtFreeHashSlot2 > UBound(hashSlot2) Then ReDim Preserve hashSlot2(NxtFreeHashSlot2 + 1000)
    GetFreeHashSlot2 = NxtFreeHashSlot2
End Function

Private Sub INIT(Level As Integer, Deflate64 As Boolean)
    Dim Temp()
    Temp = Array(0, 0, 0, 0, _
                 4, 4, 8, 4, _
                 4, 5, 16, 8, _
                 4, 6, 32, 32, _
                 4, 4, 16, 16, _
                 8, 16, 32, 32, _
                 8, 16, 128, 128, _
                 8, 32, 128, 256, _
                 32, 128, 258, 1024, _
                 32, 258, 258, 4096)        'Different levels
    GoodMatch = Temp(Level * 4)
    LazyMatch = Temp(Level * 4 + 1)
    NiceMatch = Temp(Level * 4 + 2)
    MaxChain = Temp(Level * 4 + 3)
    MaxHistore = 32767
    If Deflate64 Then MaxHistore = 65535
    MaxMatch = 258
    If Deflate64 Then MaxMatch = 65538
    NumDist = 29
    If Deflate64 Then NumDist = 31
    ReDim Slot(5 * MaxChain)
    ReDim Hash(100000)
    ReDim Hash1(255)
    ReDim HashSlot1(2000)
    ReDim hashSlot2(2000)
    Erase IData.data
    ReDim Udata.data(5000)
    IData.Bit = 0
    IData.pos = 0
    Udata.Bit = 0
    Udata.pos = 0
    NxtFreeSlot = 0
    NxtFreeHash = 0
    NxtFreeHashSlot1 = 0
    NxtFreeHashSlot2 = 0
End Sub

Private Function PutByte(Char As Integer) As Long
    If Udata.pos > UBound(Udata.data) Then ReDim Preserve Udata.data(Udata.pos + 1000)
    Udata.data(Udata.pos) = Char
    PutByte = Udata.pos
    Udata.pos = Udata.pos + 1
End Function

Private Sub putbits(Value As Long, Numbits As Long)
    CurVal = CurVal + Value * 2 ^ CBP
    CBP = CBP + Numbits
    Do While CBP > 7
        Call PutByte(CurVal And 255)
        CurVal = Fix(CurVal / 256)
        CBP = CBP - 8
    Loop
End Sub

Private Sub PutBitsRev(Value As Long, Numbits As Long)
    CurVal = CurVal + Bit_Reverse(Value, Numbits) * 2 ^ CBP
    CBP = CBP + Numbits
    Do While CBP > 7
        Call PutByte(CurVal And 255)
        CurVal = Fix(CurVal / 256)
        CBP = CBP - 8
    Loop
End Sub

Private Sub Flush_Buffer()
    If CBP > 0 Then Call PutByte(CurVal And 255)
End Sub

Private Sub AddLit(Char As Integer)
    LL(Char) = LL(Char) + 1
End Sub

Private Sub AddLen(Lenght As Long)
    Dim X As Integer
    X = 28
    Do While Lens(X) > Lenght
        X = X - 1
    Loop
    LL(X + 257) = LL(X + 257) + 1
End Sub

Private Sub AddDis(Dis As Long)
    Dim X As Integer
    X = NumDist
    Do While Dist(X) > Dis
        X = X - 1
    Loop
    DC(X) = DC(X) + 1
End Sub

Private Function Calc_Static() As Long
    Dim X As Long
    Dim Temp As Long
    For X = 0 To 143
        Temp = Temp + LL(X) * 8
    Next
    For X = 144 To 255
        Temp = Temp + LL(X) * 9
    Next
    Temp = Temp + LL(256) * 7
    For X = 257 To 279
        Temp = Temp + LL(X) * 7 + LL(X) * Lext(X - 257)
    Next
    For X = 280 To 287
        Temp = Temp + LL(X) * 8 + LL(X) * Lext(X - 257)
    Next
    For X = 0 To 31
        Temp = Temp + DC(X) * 5 + DC(X) * Dext(X)
    Next
    Calc_Static = Fix((Temp + 3) / 8) + 1
End Function

'This sub calculates the vertual dynamic lenght of the file
Private Function Calc_Dynamic() As Long
    Dim X As Long
    Dim Temp As Long
    Dim Freq(18) As Long
'Create the bits for the huffman tree for literal/length and distance tables
    Calc_Dynamic = 2 ^ 30       'set maximum size possible in case of error
    If OptimizeHuffman(LitLen, LL, 287, 15) <> 0 Then Exit Function
    If OptimizeHuffman(Distance, DC, NumDist, 15) <> 0 Then Exit Function
    For X = 0 To 256
        Temp = Temp + LL(X) * LitLen(X).Lenght
    Next
    For X = 257 To 287
        Temp = Temp + LL(X) * LitLen(X).Lenght + LL(X) * Lext(X - 257)
    Next
    For X = 0 To NumDist
        Temp = Temp + DC(X) * Distance(X).Lenght + DC(X) * Dext(X)
    Next
    ReDim Bl_Tree(18)
'Collect the lenghts from the literal/lenght tree in th BL_Tree
    For MaxLitLen = 287 To 0 Step -1
        If LitLen(MaxLitLen).Lenght <> 0 Then Exit For
    Next
    Call Scan_Tree(LitLen, 287)
'Collect the lenght from the distance tree in the BL_Tree
    For MaxDist = NumDist To 0 Step -1
        If Distance(MaxDist).Lenght <> 0 Then Exit For
    Next
    Call Scan_Tree(Distance, NumDist)
'get a frequentie count from the BL_TREE
    For X = 0 To 18
        Freq(X) = Bl_Tree(X).Code
    Next
    ReDim Bl_Tree(18)
'Build the tree lenght codes for the BL_Tree
    Call OptimizeHuffman(Bl_Tree, Freq, 18, 7)
    For Max_BL = 18 To 0 Step -1
        If Bl_Tree(Border(Max_BL)).Lenght <> 0 Then Exit For
    Next
    Temp = Temp + Max_BL * 3        '3 bits for every length code
    For X = 0 To 15
        Temp = Temp + Freq(X) * Bl_Tree(X).Lenght
    Next
    Temp = Temp + Freq(16) * (Bl_Tree(16).Lenght + 2)
    Temp = Temp + Freq(17) * (Bl_Tree(17).Lenght + 3)
    Temp = Temp + Freq(18) * (Bl_Tree(18).Lenght + 7)
    Calc_Dynamic = Fix((Temp + 3) / 8) + 1
End Function

'This sub scans the trees so a total filelenght can be calculated
Private Sub Scan_Tree(tree() As CodesType, maxcode As Long)
    Dim N As Integer
    Dim Curlen As Integer, Nextlen As Integer, PrevLen As Integer
    Dim Count As Integer
    Dim Max_Count As Integer
    Dim Min_Count As Integer
    Do While tree(maxcode).Lenght = 0   'find largest codes of non zero freq
        maxcode = maxcode - 1
    Loop
    PrevLen = -1                'last emmitted lenght
    Nextlen = tree(0).Lenght
    Max_Count = 7               'max repeat count
    Min_Count = 4               'min repeat count
    If Nextlen = 0 Then Max_Count = 138: Min_Count = 3

    For N = 0 To maxcode
        Curlen = Nextlen
        If N <> maxcode Then Nextlen = tree(N + 1).Lenght Else Nextlen = -1
        Count = Count + 1
        If (Count < Max_Count) And (Curlen = Nextlen) Then
'Do nothing
        Else
            If Count < Min_Count Then
                Bl_Tree(Curlen).Code = Bl_Tree(Curlen).Code + Count
            ElseIf Curlen <> 0 Then
                If Curlen <> PrevLen Then Bl_Tree(Curlen).Code = Bl_Tree(Curlen).Code + 1
                Bl_Tree(16).Code = Bl_Tree(16).Code + 1
            ElseIf Count <= 10 Then
                Bl_Tree(17).Code = Bl_Tree(17).Code + 1
            Else
                Bl_Tree(18).Code = Bl_Tree(18).Code + 1
            End If
            Count = 0
            PrevLen = Curlen
            If Nextlen = 0 Then
                Max_Count = 138
                Min_Count = 3
            ElseIf Curlen = Nextlen Then
                Max_Count = 6
                Min_Count = 3
            Else
                Max_Count = 7
                Min_Count = 4
            End If
        End If
    Next
End Sub

'Send a literal or distance tree in compressed form, using the codes in bl_tree.
Private Sub Send_Tree(tree() As CodesType, maxcode As Long)
    Dim N As Integer
    Dim Curlen As Integer, Nextlen As Integer, PrevLen As Integer
    Dim Count As Integer
    Dim Max_Count As Integer
    Dim Min_Count As Integer
    Do While tree(maxcode).Lenght = 0   'find largest codes of non zero freq
        maxcode = maxcode - 1
    Loop
    PrevLen = -1                'last emmitted lenght
    Nextlen = tree(0).Lenght
    Max_Count = 7               'max repeat count
    Min_Count = 4               'min repeat count
    If Nextlen = 0 Then Max_Count = 138: Min_Count = 3

    For N = 0 To maxcode
        Curlen = Nextlen
        If N <> maxcode Then Nextlen = tree(N + 1).Lenght Else Nextlen = -1
        Count = Count + 1
        If (Count < Max_Count) And (Curlen = Nextlen) Then
'Do nothing
        Else
            If Count < Min_Count Then
                Do While Count > 0
                    Call PutBitsRev(Bl_Tree(Curlen).Code, Bl_Tree(Curlen).Lenght)
                    Count = Count - 1
                Loop
            ElseIf Curlen <> 0 Then
                If Curlen <> PrevLen Then
                    Call PutBitsRev(Bl_Tree(Curlen).Code, Bl_Tree(Curlen).Lenght)
                    Count = Count - 1
                End If
                Call PutBitsRev(Bl_Tree(16).Code, Bl_Tree(16).Lenght)
                Call putbits(Count - 3, 2)
            ElseIf Count <= 10 Then
                Call PutBitsRev(Bl_Tree(17).Code, Bl_Tree(17).Lenght)
                Call putbits(Count - 3, 3)
            Else
                Call PutBitsRev(Bl_Tree(18).Code, Bl_Tree(18).Lenght)
                Call putbits(Count - 11, 7)
            End If
            Count = 0
            PrevLen = Curlen
            If Nextlen = 0 Then
                Max_Count = 138
                Min_Count = 3
            ElseIf Curlen = Nextlen Then
                Max_Count = 6
                Min_Count = 3
            Else
                Max_Count = 7
                Min_Count = 4
            End If
        End If
    Next
End Sub

Public Sub init_Trees()
    Dim X As Long
    Dim Temp()
    Erase LitLen
    Erase Distance
    Erase Bl_Tree
    Erase DC
    Erase LL
    'Create the read order array
    Temp() = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    For X = 0 To UBound(Temp): Border(X) = Temp(X): Next
    'Create the Start lenghts array
    Temp() = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    For X = 0 To UBound(Temp): Lens(X) = Temp(X): Next
    'Create the Extra lenght bits array
    Temp() = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    For X = 0 To UBound(Temp): Lext(X) = Temp(X): Next
    'Create the distance code array
    Temp() = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
    For X = 0 To UBound(Temp): Dist(X) = Temp(X): Next
    'Create the extra bits distance codes
    Temp() = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
    For X = 0 To UBound(Temp): Dext(X) = Temp(X): Next
    CBP = 0 'reset current bitposition
    CurVal = 0
End Sub

Public Function Bit_Reverse(ByVal Value As Long, ByVal Numbits As Long)
    Do While Numbits > 0
        Bit_Reverse = Bit_Reverse * 2 + (Value And 1)
        Numbits = Numbits - 1
        Value = Fix(Value / 2)
    Loop
End Function

'================================================================================
'                 H U F F M A N   T A B L E   G E N E R A T I O N
'================================================================================
Private Function OptimizeHuffman(Codes() As CodesType, Freq() As Long, NumCodes As Long, MaxBits As Long) As Integer
'Generate optimized values for BITS and CODES in a HUFFMANTABLE
'based on symbol frequency counts.  freq must be dimensioned freq(0-286)
'and contain counts of symbols 0-285.  freq is destroyed in this procedure.
'This code is created by John Korejwa <korejwa@tiac.net> and edited for use
'in Deflate by Marco v/d Berg
    Dim i              As Long
    Dim j              As Long
    Dim k              As Long
    Dim N              As Long
    Dim V1             As Long
    Dim v2             As Long
    Dim others()    As Long
    Dim codesize()  As Long
    Dim bits()      As Long
    Dim swp            As Long
    Dim swp2           As Long
    Dim freq2() As Long
    Dim Code As Long
    Dim next_code() As Long
    Dim LN As Long
    Dim CodLenPos() As Integer
    Dim OverFlow As Boolean
    ReDim next_code(MaxBits)
    ReDim others(NumCodes)
    ReDim codesize(NumCodes)
    ReDim bits(MaxBits + 5)
    ReDim CodLenPos(MaxBits + 5, NumCodes + 1)
    
    freq2 = Freq        'make copy from origenals
    v2 = 0
    For i = 0 To NumCodes  'Initialize others to -1, (this value terminates chain of indicies)
        others(i) = -1
        If Freq(i) <> 0 Then v2 = v2 + 1
    Next i
    If v2 = 0 Then OptimizeHuffman = -1: Exit Function  'no data
'check if whe have more than 1 count otherwise whe get a bitlenght 0
    If v2 = 1 Then
        For i = 0 To NumCodes
            If freq2(i) = 0 Then
                freq2(i) = 1
                Exit For
            End If
        Next
    End If

   'Generate codesize()   [find huffman code sizes]
    Do 'do loop for (#non-zero-frequencies - 1) times
        V1 = -1                            'find highest v1 for      least value of freq(v1)>0
        v2 = -1                            'find highest v2 for next least value of freq(v2)>0
        swp = 2147483647 'Max Long variable
        swp2 = 2147483647
        For i = 0 To NumCodes
            If freq2(i) <> 0 Then
                If (freq2(i) <= swp2) Then
                    If (freq2(i) <= swp) Then
                        swp2 = swp
                        v2 = V1
                        swp = freq2(i)
                        V1 = i
                    Else
                        swp2 = freq2(i)
                        v2 = i
                    End If
                End If
            End If
        Next i
        If v2 = -1 Then
            freq2(V1) = 0 'all elements in freq are now set to zero
            Exit Do      'done
        End If
        freq2(V1) = freq2(V1) + freq2(v2)     'merge the two branches
        freq2(v2) = 0
        codesize(V1) = codesize(V1) + 1    'Increment all codesizes in v1's branch
        While (others(V1) >= 0)
            V1 = others(V1)
            codesize(V1) = codesize(V1) + 1
        Wend
        others(V1) = v2                    'chain v2 onto v1's branch
        codesize(v2) = codesize(v2) + 1    'Increment all codesizes in v2's branch
        While (others(v2) >= 0)
            v2 = others(v2)
            codesize(v2) = codesize(v2) + 1
        Wend
    Loop

   'Count BITS  [find the number of codes of each size]
    N = 0
    For i = 0 To NumCodes
        If codesize(i) <> 0 Then
            bits(codesize(i)) = bits(codesize(i)) + 1
            CodLenPos(codesize(i), bits(codesize(i))) = i   'keep track where as codesize is
            CodLenPos(codesize(i), 0) = bits(codesize(i))   'Number of this bitlenght
            If N < codesize(i) Then N = codesize(i)    'Keep track of largest codesize
        End If
    Next i

    i = N
    While i > MaxBits
        While bits(i) > 0
            For j = i - 2 To 1 Step -1        'Since symbols are paired for the longest Huffman
                If bits(j) > 0 Then Exit For  'code, the symbols are removed from this length
            Next j                            'category two at a time.  The prefix for the pair
            bits(i) = bits(i) - 2             '(which is one bit shorter) is allocated to one
            bits(i - 1) = bits(i - 1) + 1     'of the pair;  then, (skipping the BITS entry for
            bits(j + 1) = bits(j + 1) + 2     'that prefix length) a code word from the next
            bits(j) = bits(j) - 1             'shortest non-zero BITS entry is converted into
        Wend                                  'a prefix for two code words one bit longer.
        i = i - 1
        OverFlow = True
    Wend

    If OverFlow Then
'reputt the bitcount array into the codesize array at its proper order
        V1 = 0
        j = 0
        Do While bits(V1) = 0
            V1 = V1 + 1
        Loop
        v2 = bits(V1)
        Do While V1 <= MaxBits
            For i = 1 To CodLenPos(j, 0)
                If codesize(CodLenPos(j, i)) <> 0 Then
                    codesize(CodLenPos(j, i)) = V1
                    v2 = v2 - 1
                    If v2 = 0 Then
                        V1 = V1 + 1
                        Do While bits(V1) = 0 And V1 <= MaxBits
                            V1 = V1 + 1
                        Loop
                        v2 = bits(V1)
                    End If
                End If
            Next
            j = j + 1
        Loop
    End If

    Code = 0
    For i = 1 To MaxBits
        Code = (Code + bits(i - 1)) * 2
        next_code(i) = Code
    Next
    For i = 0 To NumCodes
        LN = codesize(i)
        If LN <> 0 Then
            Codes(i).Lenght = LN
            Codes(i).Code = next_code(LN)
            next_code(LN) = next_code(LN) + 1
        End If
    Next
End Function

'This sub is used in case of static tree couse whe don't have frequencies there
Private Function Create_Codes(tree() As CodesType, Lenghts() As Long, NumCodes As Long) As Integer
    Dim bits(16) As Long
    Dim next_code(16) As Long
    Dim Code As Long
    Dim LN As Long
    Dim X As Long
    Dim MaxBits As Integer
    Dim Minbits As Integer
'retrieve the bitlenght count and minimum and maximum bitlenghts
    Minbits = 16
    For X = 0 To NumCodes
        bits(Lenghts(X)) = bits(Lenghts(X)) + 1
        If Lenghts(X) > MaxBits Then MaxBits = Lenghts(X)
        If Lenghts(X) < Minbits And Lenghts(X) > 0 Then Minbits = Lenghts(X)
If Lenghts(X) < 7 Then
    X = X
End If
    Next
    LN = 1
    For X = 1 To MaxBits
        LN = LN + LN
        LN = LN - bits(X)
        If LN < 0 Then Create_Codes = LN: Exit Function 'Over subscribe, Return negative
    Next
    Create_Codes = LN

    Code = 0
    bits(0) = 0
    For X = 1 To MaxBits
        Code = (Code + bits(X - 1)) * 2
        next_code(X) = Code
    Next
    For X = 0 To NumCodes
        LN = Lenghts(X)
        If LN <> 0 Then
            tree(X).Lenght = LN
            tree(X).Code = next_code(LN)
            next_code(LN) = next_code(LN) + 1
        End If
    Next
End Function




