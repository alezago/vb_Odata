Attribute VB_Name = "libOnedrive"
'Function for converting a OneDrive URL to the corresponding local path
'Algorithmically shortened code from here:
'https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
'Author: Guido Witt-D�rring
Public Function GetLocalPath$(ByVal path$, Optional ByVal returnAll As Boolean = False, Optional ByVal preferredMountPointOwner$ = "", Optional ByVal rebuildCache As Boolean = False)
#If Mac Then
Const dr& = 70, ck$ = ".849C9593-D756-4E56-8D6E-42412F2A707B", ew As Boolean = True, ab$ = "/"
#Else
Const ab$ = "\", ew As Boolean = False
#End If
Const ax$ = "GetLocalPath", ex& = 53, fr& = 7, fs& = 457, ey& = 325
Static ac As Collection, ez As Date
If Not Left$(path, 8) = "https://" Then GetLocalPath = path: Exit Function
Dim r$, i$, b$, d
Dim ds$: ds = LCase$(preferredMountPointOwner)
If Not ac Is Nothing And Not rebuildCache Then
Dim bp As Collection: Set bp = New Collection
For Each d In ac
i = d(0): r = d(1)
If InStr(1, path, r, 1) = 1 Then bp.Add Key:=d(2), Item:=Replace(Replace(path, r, i, , 1), "/", ab)
Next d
If bp.count > 0 Then
If returnAll Then
For Each d In bp: b = b & "//" & d: Next d
GetLocalPath = Mid$(b, 3): Exit Function
End If
On Error Resume Next: GetLocalPath = bp(ds): On Error GoTo 0
If GetLocalPath <> "" Then Exit Function
GetLocalPath = bp(1): Exit Function
End If
GetLocalPath = path
End If
Dim bg As Collection: Set bg = New Collection
Dim ay, du$
#If Mac Then
Dim cl$, dv As Boolean
b = Environ("HOME")
du = b & "/Library/Application Support/Microsoft/Office/CLP/"
b = Left$(b, InStrRev(b, "/Library/Containers/", , 0))
bg.Add b & "Library/Containers/com.microsoft.OneDrive-mac/Data/Library/Application Support/OneDrive/settings/"
bg.Add b & "Library/Application Support/OneDrive/settings/"
cl = b & "Library/CloudStorage/"
#Else
bg.Add Environ("LOCALAPPDATA") & "\Microsoft\OneDrive\settings\"
du = Environ("LOCALAPPDATA") & "\Microsoft\Office\CLP\"
#End If
Dim a&
#If Mac Then
Dim az() As Variant: ReDim az(1 To bg.count * 11 + 1)
For Each ay In bg
For a = a + 1 To a + 9
az(a) = ay & "Business" & a Mod 11
Next a
az(a) = ay: a = a + 1
az(a) = ay & "Personal"
Next ay
az(a + 1) = cl
Dim dw As Boolean
dw = GetSetting("GetLocalPath", "AccessRequestInfoMsg", "Displayed", "False") = "True"
If Not dw Then MsgBox "The current VBA Project requires access to the OneDrive settings files to translate a OneDrive URL to the local path of the locally synchronized file/folder on your Mac. Because these files are located outside of Excels sandbox, file-access must be granted explicitly. Please approve the access requests following this message.", vbInformation
If Not GrantAccessToMultipleFiles(az) Then Err.Raise dr, ax
#End If
Dim db As Collection: Set db = New Collection
For Each ay In bg
Dim h$: h = Dir(ay, 16)
Do Until h = ""
If h = "Personal" Or h Like "Business#" Then db.Add Item:=ay & h & ab
h = Dir(, 16)
Loop
Next ay
If Not ac Is Nothing Or ew Then
Dim bf As Collection: Set bf = New Collection
Dim g
For Each g In db
Dim t$: t = IIf(g Like "*" & ab & "Personal" & ab, "????????????*", "????????-????-????-????-????????????")
Dim p$: p = Dir(g, vbNormal)
Do Until p = ""
If p Like t & ".ini" Or p Like t & ".dat" Or p Like "ClientPolicy*.ini" Or StrComp(p, "GroupFolders.ini", 1) = 0 Or StrComp(p, "global.ini", 1) = 0 Or StrComp(p, "SyncEngineDatabase.db", 1) = 0 Then bf.Add Item:=g & p
p = Dir
Loop
Next g
End If
If Not ac Is Nothing And Not rebuildCache Then
Dim au
For Each au In bf
If FileDateTime(au) > ez Then rebuildCache = True: Exit For
Next au
If Not rebuildCache Then Exit Function
End If
Dim f&, am$, e() As Byte, j&, q&, bs&, av() As Byte, cn$, n() As Byte, ao$, ak() As Byte, ba() As Byte, bt$, aw&, y&, dz&, ea&
ez = Now()
#If Mac Then
Dim z As Collection: Set z = New Collection
h = Dir(cl, 16)
Do Until h = ""
If h Like "OneDrive*" Then
dv = True
g = cl & h & ab
au = cl & h & ab & ck
z.Add Item:=g
bf.Add Item:=g
bf.Add Item:=au
End If
h = Dir(, 16)
Loop
If ac Is Nothing Then
Dim dc
If bf.count > 0 Then
ReDim dc(1 To bf.count)
For a = 1 To UBound(dc): dc(a) = bf(a): Next a
If Not GrantAccessToMultipleFiles(dc) Then Err.Raise dr, ax
End If
End If
If dv Then
For a = z.count To 1 Step -1
Dim bu&: bu = 0
On Error Resume Next
bu = GetAttr(z(a) & ck)
Dim bv As Boolean: bv = False
If Err.Number = 0 Then bv = Not CBool(bu And 16)
On Error GoTo 0
If Not bv Then
h = Dir(z(a), 16)
Do Until h = ""
If Not h Like ".Trash*" And h <> "Icon" Then
z.Add z(a) & h & ab
z.Add z(a) & h & ab & ck, z(a) & h & ab
End If
h = Dir(, 16)
Loop
z.Remove a
End If
Next a
If z.count > 0 Then
ReDim az(1 To z.count)
For a = 1 To z.count: az(a) = z(a): Next a
If Not GrantAccessToMultipleFiles(az) Then Err.Raise dr, ax
End If
On Error Resume Next
For a = z.count To 1 Step -1
z.Remove z(a)
Next a
On Error GoTo 0
Dim eb As Collection
Set eb = New Collection
For Each g In z
bu = 0
On Error Resume Next
bu = GetAttr(g & ck)
bv = False
If Err.Number = 0 Then bv = Not CBool(bu And 16)
On Error GoTo 0
If bv Then
f = FreeFile(): b = "": au = g & ck
Dim ec As Boolean: ec = False
On Error GoTo ReadFailed
Open au For Binary Access Read As #f
ReDim e(0 To LOF(f)): Get f, , e: b = e
ec = True
ReadFailed: On Error GoTo -1
Close #f: f = 0
On Error GoTo 0
If ec Then
av = b
If LenB(b) > 0 Then
ReDim n(0 To LenB(b) * 2 - 1): q = 0
For j = LBound(av) To UBound(av)
n(q) = av(j): q = q + 2
Next j
b = n
Else: b = ""
End If
Else
au = MacScript("return path to startup disk as string") & Replace(Mid$(au, 2), ab, ":")
b = MacScript("return read file """ & au & """ as string")
End If
If InStr(1, b, """guid"" : """, 0) Then
b = Split(b, """guid"" : """)(1)
am = Left$(b, InStr(1, b, """", 0) - 1)
eb.Add Key:=am, Item:=VBA.Array(am, Left$(g, Len(g) - 1))
Else
Debug.Print "Warning, empty syncIDFile encountered!"
End If
End If
Next g
End If
If Not dw Then SaveSetting "GetLocalPath", "AccessRequestInfoMsg", "Displayed", "True"
#End If
Dim c, w$(), s&, co$, bk$, dd$, cp$, bl$, aa$, al$, at$, bz$, fx$, ca As Boolean, cb$, cc$, de$, fc$, fd$, ag$, fe$
Dim ff$: ff = ChrB$(2)
Dim ed As String * 4: MidB$(ed, 1) = ChrB$(1)
Dim ee$: ee = ChrB$(0)
#If Mac Then
Const ef$ = vbNullChar & vbNullChar
#Else
Const ef$ = vbNullChar
#End If
Dim cq As Collection, fi As Date
Set cq = New Collection
Set ac = New Collection
For Each g In db
h = Mid$(g, InStrRev(g, ab, Len(g) - 1, 0) + 1)
h = Left$(h, Len(h) - 1)
If Dir(g & "global.ini", vbNormal) = "" Then GoTo NextFolder
f = FreeFile()
Open g & "global.ini" For Binary Access Read As #f
ReDim e(0 To LOF(f)): Get f, , e
Close #f: f = 0
#If Mac Then
bt = e: GoSub DecodeUTF8
e = ao
#End If
For Each c In Split(e, vbNewLine)
If c Like "cid = *" Then t = Mid$(c, 7): Exit For
Next c
If t = "" Then GoTo NextFolder
If (Dir(g & t & ".ini") = "" Or (Dir(g & "SyncEngineDatabase.db") = "" And Dir(g & t & ".dat") = "")) Then GoTo NextFolder
If h Like "Business#" Then
bz = Replace(space$(32), " ", "[a-f0-9]") & "*"
ElseIf h = "Personal" Then
bz = Replace(space$(12), " ", "[A-F0-9]") & "*!###*"
End If
p = Dir(du, vbNormal)
Do Until p = ""
a = InStrRev(p, t, , 1)
If a > 1 And t <> "" Then bl = LCase$(Left$(p, a - 2)): Exit Do
p = Dir
Loop
#If Mac Then
On Error Resume Next
fi = cq(h)
ca = (Err.Number = 0)
On Error GoTo 0
If ca Then
If FileDateTime(g & t & ".ini") < fi Then
GoTo NextFolder
Else
For a = ac.count To 1 Step -1
If ac(a)(5) = h Then
ac.Remove a
End If
Next a
cq.Remove h
cq.Add Key:=h, Item:=FileDateTime(g & t & ".ini")
End If
Else
cq.Add Key:=h, Item:=FileDateTime(g & t & ".ini")
End If
#End If
Dim bb As Collection: Set bb = New Collection
p = Dir(g, vbNormal)
Do Until p = ""
If p Like "ClientPolicy*.ini" Then
f = FreeFile()
Open g & p For Binary Access Read As #f
ReDim e(0 To LOF(f)): Get f, , e
Close #f: f = 0
#If Mac Then
bt = e: GoSub DecodeUTF8
e = ao
#End If
bb.Add Key:=p, Item:=New Collection
For Each c In Split(e, vbNewLine)
If InStr(1, c, " = ", 0) Then
bk = Left$(c, InStr(1, c, " = ", 0) - 1)
b = Mid$(c, InStr(1, c, " = ", 0) + 3)
Select Case bk
Case "DavUrlNamespace"
bb(p).Add Key:=bk, Item:=b
Case "SiteID", "IrmLibraryId", "WebID"
b = Replace(LCase$(b), "-", "")
If Len(b) > 3 Then b = Mid$(b, 2, Len(b) - 2)
bb(p).Add Key:=bk, Item:=b
End Select
End If
Next c
End If
p = Dir
Loop
Dim x As Collection: Set x = Nothing
If Dir(g & t & ".dat") = "" Then GoTo Continue
Const fz& = 1000
Const cs& = 255
Dim bc&: bc = -1
Try: On Error GoTo Catch
Set x = New Collection
Dim ct&: ct = 1
Dim cu As Date: cu = FileDateTime(g & t & ".dat")
a = 0
Do
If FileDateTime(g & t & ".dat") > cu Then GoTo Try
f = FreeFile
Open g & t & ".dat" For Binary Access Read As #f
Dim dg&: dg = LOF(f)
If bc = -1 Then bc = dg
ReDim e(0 To bc + fz)
Get f, ct, e: b = e
Dim cv&: cv = LenB(b)
Close #f: f = 0
ct = ct + bc
For d = 16 To 8 Step -8
a = InStrB(d + 1, b, ed, 0)
Do While a > d And a < cv - 168
If StrComp(MidB$(b, a - d, 1), ff, 0) = 0 Then
a = a + 8: s = InStrB(a, b, ee, 0) - a
If s < 0 Then s = 0
If s > 39 Then s = 39
#If Mac Then
cn = MidB$(b, a, s)
GoSub DecodeANSI: al = ao
#Else
al = StrConv(MidB$(b, a, s), 64)
#End If
a = a + 39: s = InStrB(a, b, ee, 0) - a
If s < 0 Then s = 0
If s > 39 Then s = 39
#If Mac Then
cn = MidB$(b, a, s)
GoSub DecodeANSI: aa = ao
#Else
aa = StrConv(MidB$(b, a, s), 64)
#End If
a = a + 121
s = InStr(-Int(-(a - 1) / 2) + 1, b, ef, 0) * 2 - a - 1
If s > cs * 2 Then s = cs * 2
If s < 0 Then s = 0
If al Like bz And aa Like bz Then
#If Mac Then
Do While s Mod 4 > 0
If s > cs * 4 Then Exit Do
s = InStr(-Int(-(a + s) / 2) + 1, b, ef, 0) * 2 - a - 1
Loop
If s > cs * 4 Then s = cs * 4
ak = MidB$(b, a, s)
ReDim n(LBound(ak) To UBound(ak))
j = LBound(ak): q = LBound(ak)
Do While j < UBound(ak)
If ak(j + 2) + ak(j + 3) = 0 Then
n(q) = ak(j)
n(q + 1) = ak(j + 1)
q = q + 2
Else
If ak(j + 3) <> 0 Then Err.Raise ey, ax
y = ak(j + 2) * &H10000 + ak(j + 1) * &H100& + ak(j)
bs = y - &H10000
ea = &HD800& Or (bs \ &H400&)
dz = &HDC00& Or (bs And &H3FF)
n(q) = ea And &HFF&
n(q + 1) = ea \ &H100&
n(q + 2) = dz And &HFF&
n(q + 3) = dz \ &H100&
q = q + 4
End If
j = j + 4
Loop
If q > LBound(n) Then
ReDim Preserve n(LBound(n) To q - 1)
at = n
Else: at = ""
End If
#Else
at = MidB$(b, a, s)
#End If
x.Add VBA.Array(aa, at), al
End If
End If
a = InStrB(a + 1, b, ed, 0)
Loop
If x.count > 0 Then Exit For
Next d
Loop Until ct >= dg Or bc >= dg
GoTo Continue
Catch:
Select Case Err.Number
Case fs
x.Remove al
Resume
Case Is <> fr: Err.Raise Err, ax
End Select
If bc > &HFFFFF Then bc = bc / 2: Resume Try
Err.Raise Err, ax
Continue:
On Error GoTo 0
If Not x Is Nothing Then GoTo SkipDbFile
f = FreeFile()
Open g & "SyncEngineDatabase.db" For Binary Access Read As #f
cv = LOF(f)
If cv = 0 Then GoTo CloseFile
Dim eg$: eg = ChrW$(&H808)
Const gd& = 8, ge& = -3, fl As Byte = 9, fm& = 6, fn& = &H16, gf& = &H15, ce& = -16, cf& = -15, eh& = &H100000
Dim bm&, cg&, bd&, ah(1 To 4) As Byte, an$, dk$, ei&, ej&, ek&, dl&, el As Byte, em As Byte, en As Boolean, eo&
cu = 0
ReDim e(1 To eh)
Do
a = 0
If FileDateTime(g & "SyncEngineDatabase.db") > cu Then
Set x = New Collection
Dim dm As Collection: Set dm = New Collection
cu = FileDateTime(g & "SyncEngineDatabase.db")
bm = 1
an = ""
End If
If LenB(an) > 0 Then
at = MidB$(b, ei, ej)
End If
Get f, bm, e
b = e
a = InStrB(1 - ce, b, eg, 0)
dl = 0
Do While a > 0
If a + ce - 2 > dl And LenB(an) > 0 Then
If dl > 0 Then
at = MidB$(b, ei, ej)
End If
bt = at: GoSub DecodeUTF8
at = ao
On Error Resume Next
x.Add VBA.Array(dk, at), an
If Err.Number <> 0 Then
If dm(an) < em Then
If x(an)(1) <> at Or x(an)(0) <> dk Then
x.Remove an
dm.Remove an
x.Add VBA.Array(dk, at), an
End If
End If
End If
dm.Add em, an
On Error GoTo 0
an = ""
End If
If e(a + ge) <> gd Then GoTo NextSig
en = True
eo = 0
If e(a + cf) = gf Then
j = a + cf
ElseIf e(a + ce) = fn Then
j = a + ce
en = False
ElseIf e(a + cf) <= fl Then
j = a + cf
ElseIf e(a + cf) = fn Then
j = a + cf
eo = 1
Else
GoTo NextSig
End If
el = e(j)
cg = fm
For q = 1 To 4
If q = 1 And el <= fl Then
ah(q) = e(j + 2)
Else
ah(q) = e(j + q)
End If
If ah(q) < 37 Or ah(q) Mod 2 = 0 Then GoTo NextSig
ah(q) = (ah(q) - 13) / 2
cg = cg + ah(q)
Next q
If en Then
bd = e(j + 5)
If bd < 15 Or bd Mod 2 = 0 Then GoTo NextSig
bd = (bd - 13) / 2
Else
bd = (e(j + 5) - 128) * 64 + (e(j + 6) - 13) / 2
If bd < 1 Or e(j + 6) Mod 2 = 0 Then GoTo NextSig
End If
cg = cg + bd
ek = a + cg - 1
If ek > eh Then
a = a - 1
Exit Do
End If
j = a + fm + eo
#If Mac Then
cn = MidB$(b, j, ah(1))
GoSub DecodeANSI: al = ao
#Else
al = StrConv(MidB$(b, j, ah(1)), 64)
#End If
j = j + ah(1)
aa = StrConv(MidB$(b, j, ah(2)), 64)
#If Mac Then
cn = MidB$(b, j, ah(2))
GoSub DecodeANSI: aa = ao
#Else
aa = StrConv(MidB$(b, j, ah(2)), 64)
#End If
If al Like bz And aa Like bz Then
ei = j + ah(2) + ah(3) + ah(4)
ej = bd
an = Left$(al, 32)
dk = Left$(aa, 32)
em = el
dl = ek
End If
NextSig:
a = InStrB(a + 1, b, eg, 0)
Loop
If a = 0 Then
bm = bm + eh + ce
Else
bm = bm + a + ce
End If
Loop Until bm > cv
CloseFile:
Close #f
SkipDbFile:
f = FreeFile()
Open g & t & ".ini" For Binary Access Read As #f
ReDim e(0 To LOF(f)): Get f, , e
Close #f: f = 0
#If Mac Then
bt = e: GoSub DecodeUTF8:
e = ao
#End If
Dim ep As Collection: Set ep = New Collection
Dim eq
eq = VBA.Array("libraryScope", "libraryFolder", "AddedScope")
Dim dn As Collection: Set dn = New Collection
For Each d In eq
dn.Add New Collection, CStr(d)
Next d
For Each c In Split(e, vbNewLine)
If InStr(1, c, " = ", 0) = 0 Then Exit For
bk = Left$(c, InStr(1, c, " = ", 0) - 1)
Select Case bk: Case "libraryScope", "libraryFolder", "AddedScope"
dn(bk).Add c, Split(c, " ", 4, 0)(2)
End Select
Next c
For Each d In eq
Dim dp As Collection: Set dp = dn(d)
a = 0
Do Until dp.count = 0
On Error Resume Next
c = "": c = dp(CStr(a))
On Error GoTo 0
If c <> "" Then
ep.Add c
dp.Remove CStr(a)
End If
a = a + 1
Loop
Next d
If h Like "Business#" Then
Dim er As Collection: Set er = New Collection
dd = ""
For Each c In ep
r = "": i = "": w = Split(c, """")
Select Case Left$(c, InStr(1, c, " = ", 0) - 1)
Case "libraryScope"
i = w(9)
ag = i: am = Split(w(10), " ")(2)
co = Split(c, " ")(2)
fx = w(3): w = Split(w(8), " ")
cb = w(1): de = w(2): cc = w(3)
If Split(c, " ", 4, 0)(2) = "0" Then
dd = i: p = "ClientPolicy.ini"
fd = am: fe = ag
Else: p = "ClientPolicy_" & cc & cb & ".ini"
End If
On Error Resume Next
r = bb(p)("DavUrlNamespace")
On Error GoTo 0
If r = "" Then
For Each d In bb
If d("SiteID") = cb And d("WebID") = de And d("IrmLibraryId") = cc Then
r = d("DavUrlNamespace"): Exit For
End If
Next d
End If
If r = "" Then Err.Raise ex, ax
er.Add VBA.Array(co, r), co
If Not i = "" Then ac.Add VBA.Array(i, r, bl, am, ag, h), Key:=i
Case "libraryFolder"
co = Split(c, " ")(3)
i = w(1): ag = i
am = Split(w(4), " ")(1)
b = "": aa = Left$(Split(c, " ")(4), 32)
Do
On Error Resume Next: x aa
ca = (Err.Number = 0): On Error GoTo 0
If Not ca Then Exit Do
b = x(aa)(1) & "/" & b
aa = x(aa)(0)
Loop
r = er(co)(1) & b
ac.Add VBA.Array(i, r, bl, am, ag, h), i
Case "AddedScope"
If dd = "" Then Err.Raise ey, ax
cp = w(5): If cp = " " Then cp = ""
w = Split(w(4), " "): cb = w(1)
de = w(2): cc = w(3): fc = w(4)
p = "ClientPolicy_" & cc & cb & fc & ".ini"
On Error Resume Next
r = bb(p)("DavUrlNamespace") & cp
On Error GoTo 0
If r = "" Then
For Each d In bb
If d("SiteID") = cb And d("WebID") = de And d("IrmLibraryId") = cc Then
r = d("DavUrlNamespace") & cp
Exit For
End If
Next d
End If
If r = "" Then Err.Raise ex, ax
b = "": aa = Left$(Split(c, " ")(3), 32)
Do
On Error Resume Next: x aa
ca = (Err.Number = 0): On Error GoTo 0
If Not ca Then Exit Do
b = x(aa)(1) & ab & b
aa = x(aa)(0)
Loop
i = dd & ab & b
ac.Add VBA.Array(i, r, bl, fd, fe, h), i
Case Else: Exit For
End Select
Next c
ElseIf h = "Personal" Then
For Each c In Split(e, vbNewLine)
If c Like "library = *" Then
w = Split(c, """"): i = w(3)
ag = i: am = Split(w(4), " ")(2)
Exit For
End If
Next c
On Error Resume Next
r = bb("ClientPolicy.ini")("DavUrlNamespace")
On Error GoTo 0
If i = "" Or r = "" Or t = "" Then GoTo NextFolder
ac.Add VBA.Array(i, r & "/" & t, bl, am, ag, h), Key:=i
If Dir(g & "GroupFolders.ini") = "" Then GoTo NextFolder
t = "": f = FreeFile()
Open g & "GroupFolders.ini" For Binary Access Read As #f
ReDim e(0 To LOF(f)): Get f, , e
Close #f: f = 0
#If Mac Then
bt = e: GoSub DecodeUTF8
e = ao
#End If
For Each c In Split(e, vbNewLine)
If c Like "*_BaseUri = *" And t = "" Then
t = LCase$(Mid$(c, InStrRev(c, "/", , 0) + 1, InStrRev(c, "!", , 0) - InStrRev(c, "/", , 0) - 1))
al = Left$(c, InStr(1, c, "_", 0) - 1)
ElseIf t <> "" Then
ac.Add VBA.Array(i & ab & x(al)(1), r & "/" & t & "/" & Mid$(c, Len(al) + 9), bl, am, ag, h), Key:=i & ab & x(al)(1)
t = "": al = ""
End If
Next c
End If
NextFolder:
t = "": b = "": bl = ""
Next g
Dim ch As Collection: Set ch = New Collection
For Each d In ac
i = d(0): r = d(1): ag = d(4)
If Right$(r, 1) = "/" Then r = Left$(r, Len(r) - 1)
If Right$(i, 1) = ab Then i = Left$(i, Len(i) - 1)
If Right$(ag, 1) = ab Then ag = Left$(ag, Len(ag) - 1)
ch.Add VBA.Array(i, r, d(2), d(3), ag), i
Next d
Set ac = ch
#If Mac Then
If dv Then
Set ch = New Collection
For Each d In ac
i = d(0): am = d(3): ag = d(4)
i = Replace(i, ag, eb(am)(1), , 1)
ch.Add VBA.Array(i, d(1), d(2)), i
Next d
Set ac = ch
End If
#End If
GetLocalPath = GetLocalPath(path, returnAll, ds, False): Exit Function
Exit Function
DecodeUTF8:
Const ci As Boolean = False
Dim u&, o&, bn&
Static cj(0 To 255) As Byte
Static fp&(2 To 4)
Static dq&(2 To 4)
If cj(0) = 0 Then
For u = &H0& To &H7F&: cj(u) = 1: Next u
For u = &HC2& To &HDF&: cj(u) = 2: Next u
For u = &HE0& To &HEF&: cj(u) = 3: Next u
For u = &HF0& To &HF4&: cj(u) = 4: Next u
For u = 2 To 4: fp(u) = (2 ^ (7 - u) - 1): Next u
dq(2) = &H80&: dq(3) = &H800&: dq(4) = &H10000
End If
Dim es As Byte
ba = bt
ReDim n(0 To (UBound(ba) - LBound(ba) + 1) * 2)
o = 0
u = LBound(ba)
Do While u <= UBound(ba)
y = ba(u)
aw = cj(y)
If aw = 0 Then
If ci Then Err.Raise 5
GoTo insertErrChar
ElseIf aw = 1 Then
n(o) = y
o = o + 2
ElseIf u + aw - 1 > UBound(ba) Then
If ci Then Err.Raise 5
GoTo insertErrChar
Else
y = ba(u) And fp(aw)
For bn = 1 To aw - 1
es = ba(u + bn)
If (es And &HC0&) = &H80& Then
y = (y * &H40&) + (es And &H3F)
Else
If ci Then Err.Raise 5
GoTo insertErrChar
End If
Next bn
If y < dq(aw) Then
If ci Then Err.Raise 5
GoTo insertErrChar
ElseIf y < &HD800& Then
n(o) = CByte(y And &HFF&)
n(o + 1) = CByte(y \ &H100&)
o = o + 2
ElseIf y < &HE000& Then
If ci Then Err.Raise 5
GoTo insertErrChar
ElseIf y < &H10000 Then
If y = &HFEFF& Then GoTo nextCp
n(o) = y And &HFF&
n(o + 1) = y \ &H100&
o = o + 2
ElseIf y < &H110000 Then
bs = y - &H10000
Dim et&: et = &HDC00& Or (bs And &H3FF)
Dim eu&: eu = &HD800& Or (bs \ &H400&)
n(o) = eu And &HFF&
n(o + 1) = eu \ &H100&
n(o + 2) = et And &HFF&
n(o + 3) = et \ &H100&
o = o + 4
Else
If ci Then Err.Raise 5
insertErrChar: n(o) = &HFD
n(o + 1) = &HFF
o = o + 2
If aw = 0 Then aw = 1
End If
End If
nextCp: u = u + aw
Loop
ao = MidB$(n, 1, o)
Return
DecodeANSI:
av = cn
o = UBound(av) - LBound(av) + 1
If o > 0 Then
ReDim n(0 To o * 2 - 1): bn = 0
For o = LBound(av) To UBound(av)
n(bn) = av(o): bn = bn + 2
Next o
ao = n
Else
ao = ""
End If
Return
End Function
