Attribute VB_Name = "Module1"
'  GeoHash Routines for VBA 2018
'
'  Copyright (C) 2018 by David Mugo (author) and Martin Stobbs (co-author)
'  Distributed uner the MIT License
'
'  Permission is hereby granted, free of charge, to any person obtaining a copy
'  of this software and associated documentation files (the "Software"), to deal
'  in the Software without restriction, including without limitation the rights
'  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'  copies of the Software, and to permit persons to whom the Software is
'  furnished to do so, subject to the following conditions:
'
'  The above copyright notice and this permission notice shall be included in
'  all copies or substantial portions of the Software.
'
'  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'  THE SOFTWARE.
'  ------------------------------------------------------------------------------
'
'   This code is a direct derivation from:
'      GeoHash Routines for Javascript 2008 (C) David Troy at
'      https://github.com/davetroy/geohash-js
'      and
'      Geohash .Net Library (C) 2011 by Sharon Lourduraj at
'      https://github.com/sharonjl/geohash-net
'
'   Included are basic test sub-procedures that demonstrate function calls.'
'   Copy and paste code below into a new module in Excel Visual Basic Editor

Private Const Base32 As String = "0123456789bcdefghjkmnpqrstuvwxyz"

Private Bits() As Variant
Private Neighbors(2, 4) As String
Private Borders(2, 4) As String

Private Num As Integer

' Set values for Bits, Neighbors and Borders arrays
Sub InitializeBitsNeighboursBorders()
    Bits = Array(16, 8, 4, 2, 1)

    Neighbors(0, 0) = "p0r21436x8zb9dcf5h7kjnmqesgutwvy"
    Neighbors(0, 1) = "bc01fg45238967deuvhjyznpkmstqrwx"
    Neighbors(0, 2) = "14365h7k9dcfesgujnmqp0r2twvyx8zb"
    Neighbors(0, 3) = "238967debc01fg45kmstqrwxuvhjyznp"
    
    Neighbors(1, 0) = "bc01fg45238967deuvhjyznpkmstqrwx"
    Neighbors(1, 1) = "p0r21436x8zb9dcf5h7kjnmqesgutwvy"
    Neighbors(1, 2) = "238967debc01fg45kmstqrwxuvhjyznp"
    Neighbors(1, 3) = "14365h7k9dcfesgujnmqp0r2twvyx8zb"
    
    Borders(0, 0) = "prxz"
    Borders(0, 1) = "bcfguvyz"
    Borders(0, 2) = "028b"
    Borders(0, 3) = "0145hjnp"
    
    Borders(1, 0) = "bcfguvyz"
    Borders(1, 1) = "prxz"
    Borders(1, 2) = "0145hjnp"
    Borders(1, 3) = "028b"
End Sub

' Enum alternative for direction
Function direction(myVar As Variant) As Integer
  Select Case myVar
    Case "N", 0: direction = 0
    Case "E", 1: direction = 1
    Case "S", 2: direction = 2
    Case "W", 3: direction = 3
    Case Else: direction = 0
  End Select
End Function

' Calculates geohash for neighbouring grids
' Defaults to 'N', if 'direct' parameter is invalid
Function CalculateAdjacent(ByVal hash As String, ByVal direct As Variant) As String
    Dim lastChr As String
    Dim typ As Integer
    Dim dir As Byte
    Dim nHash As String
    
    InitializeBitsNeighboursBorders
    
    hash = LCase(hash)
    
    typ = Len(hash) Mod 2
    
    lastChr = Mid(hash, Len(hash), 1)

    dir = direction(direct)
    
    nHash = Mid(hash, 1, Len(hash) - 1)
    
    ' VBA InStr starts at 1, C# IndexOf starts at 0
    If InStr(Borders(typ, dir), lastChr) - 1 <> -1 Then
        nHash = CalculateAdjacent(nHash, direct)
    End If

    CalculateAdjacent = nHash & Mid(Base32, InStr(Neighbors(typ, dir), lastChr), 1)
    
End Function

' Tests CalculateAdjacent function
Sub TestAdjacent()
    Debug.Print CalculateAdjacent("sbrgzyrg", "K")
End Sub

Private Sub RefineInterval(interval() As Variant, ByVal cd As Integer, ByVal mask As Integer)
    If (cd And mask) <> 0 Then
        interval(0) = (interval(0) + interval(1)) / 2
    Else
        interval(1) = (interval(0) + interval(1)) / 2
    End If
End Sub

' Decodes geohash to obtain Latitude and Longitude in an array
Private Function Decode(ByVal geohash As String) As Variant()
    Dim even As Boolean
    Dim lat() As Variant
    Dim lon() As Variant
    Dim cd As Integer
    Dim mask As Integer
    
    Dim c As String
    Dim j As Integer
    Dim i As Integer
    
    InitializeBitsNeighboursBorders
    
    even = True
    lat = Array(-90, 90)
    lon = Array(-180, 180)

    For i = 1 To Len(geohash)
        c = Mid(geohash, i, 1)
        cd = InStr(Base32, c) - 1
        For j = 0 To 4
            mask = Bits(j)
            If even = True Then
                RefineInterval lon, cd, mask
            Else
                RefineInterval lat, cd, mask
            End If

            even = Not even
        Next
    Next

    Decode = Array((lat(0) + lat(1)) / 2, (lon(0) + lon(1)) / 2)
End Function

' Decodes geohash to obtain Latitude
Function DecodeLat(ByVal geohash As String) As Variant
    DecodeLat = Decode(geohash)(0)
End Function

' Decodes geohash to obtain Longitude
Function DecodeLon(ByVal geohash As String) As Variant
    DecodeLon = Decode(geohash)(1)
End Function

' Tests Decode functions
Sub TestDecode()
    Debug.Print DecodeLat("t025bn2")  ' returns  2.1..
    Debug.Print DecodeLon("t025bn2")  ' returns 45.0...
End Sub

' Encodes Latitude, Longitude & Precision to obtain geohash
Function Encode(ByVal latitude As Double, ByVal longitude As Double, Optional ByVal precision As Integer = 12) As String
    Dim even As Boolean
    Dim bit As Integer
    Dim ch As Integer
    Dim geohash As String
    Dim lat() As Variant
    Dim lon() As Variant
    
    InitializeBitsNeighboursBorders
    
    even = True
    bit = 0
    ch = 0
    geohash = ""
    lat = Array(-90, 90)
    lon = Array(-180, 180)
    
    If precision < 1 Or precision > 20 Then precision = 12
    
    Do While Len(geohash) < precision
        Dim midd As Double
        If even Then
            midd = (lon(0) + lon(1)) / 2
            If longitude > midd Then
                ch = ch Or Bits(bit)
                lon(0) = midd
            Else
                lon(1) = midd
            End If
        Else
            midd = (lat(0) + lat(1)) / 2
            If latitude > midd Then
                ch = ch Or Bits(bit)
                lat(0) = midd
            Else
                lat(1) = midd
            End If
        End If

        even = Not even

        
        If bit < 4 Then
            bit = bit + 1

        Else
            geohash = geohash & Mid$(Base32, ch + 1, 1)
            bit = 0
            ch = 0
        End If
    Loop

    Encode = geohash
    
End Function

' Tests Encode function
Sub TestEncode()
    Debug.Print Encode(2.1, 45.00000000001, 7) ' returns t025bn2
End Sub
