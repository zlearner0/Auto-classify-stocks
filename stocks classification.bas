Attribute VB_Name = "Module1"
Option Explicit
Sub SectorsSplit()


Dim tableWidth, tableLength As Long

Dim tableArray, headersArray, sectorTickersArray(), industryTickersArray(), tempTableArray(), separatorArray(), avrgArray(1) _
, ratiosArray(), boldRowArray As Variant

Dim x, y, z, m, n, i, j, k, u, v, w, r, colNo1, colNo2, colNo3, colNo4, colNo5, colNo6, colNo7, colNo8, colNo9, colNo10, colNo11 _
, industriesNo, columnNumber, no1, no2 As Long

Dim TempTxt1, TempTxt2, header, tempNo1, tempNo2, columnLetter, c, columnLetter1, columnLetter2 As String

Dim isNewSector As Boolean

'headers filter criteria array
'headersArray = Array("Sector", "Industry", "PE1")

'create array for all stocks table
tableLength = AllStocks.Range("A500000").End(xlUp).Row
tableWidth = AllStocks.Range("ZZ1").End(xlToLeft).Column

tableArray = AllStocks.Range(Cells(1, 1).Address, Cells(tableLength, tableWidth).Address)



'getting the column no. of the "Sector" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "SECTOR" Then
    colNo1 = x
    Exit For
  End If
Next x
'getting the column no. of the "Industry" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "INDUSTRY" Then
    colNo2 = x
    Exit For
  End If
Next x
'getting the column no. of the "PE1" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "PE1" Then
    colNo3 = x
    Exit For
  End If
Next x
'getting the column no. of the "EPS0" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "EPS0" Then
    colNo4 = x
    Exit For
  End If
Next x
'getting the column no. of the "EPS1" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "EPS1" Then
    colNo5 = x
    Exit For
  End If
Next x
'getting the column no. of the "EPS2" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "EPS2" Then
    colNo6 = x
    Exit For
  End If
Next x
'getting the column no. of the "EG1" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "EG1" Then
    colNo7 = x
    Exit For
  End If
Next x
'getting the column no. of the "EG2" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "EG2" Then
    colNo8 = x
    Exit For
  End If
Next x
'getting the column no. of the "PE2" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "PE2" Then
    colNo9 = x
    Exit For
  End If
Next x
'getting the column no. of the "PEG1" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "PEG1" Then
    colNo10 = x
    Exit For
  End If
Next x
'getting the column no. of the "PEG2" in the tableArray
For x = 1 To UBound(tableArray, 2)
  If Trim(UCase(tableArray(1, x))) = "PEG2" Then
    colNo11 = x
    Exit For
  End If
Next x





'refine EG% from netative to positive and greater than 100% to be 100% and vice versa for all table
For x = LBound(tableArray, 1) + 1 To UBound(tableArray, 1)
 If tableArray(x, colNo4) <> "" And tableArray(x, colNo5) <> "" And tableArray(x, colNo6) <> "" Then
 
  If tableArray(x, colNo4) < 0 And tableArray(x, colNo5) > 0 And tableArray(x, colNo7) > 1 Then
    tableArray(x, colNo7) = 1
  End If
  If tableArray(x, colNo5) < 0 And tableArray(x, colNo6) > 0 And tableArray(x, colNo8) > 1 Then
    tableArray(x, colNo8) = 1
  End If
  If tableArray(x, colNo4) > 0 And tableArray(x, colNo5) < 0 And tableArray(x, colNo7) < -1 Then
    tableArray(x, colNo7) = -1
  End If
  If tableArray(x, colNo5) > 0 And tableArray(x, colNo6) < 0 And tableArray(x, colNo8) < -1 Then
    tableArray(x, colNo8) = -1
  End If
  
 End If
Next x


'set in order sectors
For x = LBound(tableArray, 1) + 1 To UBound(tableArray, 1)
    For y = x To UBound(tableArray, 1)
      If UCase(Trim(tableArray(y, colNo1))) < UCase(Trim(tableArray(x, colNo1))) Then
         For z = LBound(tableArray, 2) To UBound(tableArray, 2)
            If IsError(tableArray(x, z)) Then
              tableArray(x, z) = 0
            End If
            If IsError(tableArray(y, z)) Then
              tableArray(y, z) = 0
            End If
            
            TempTxt1 = tableArray(x, z)
            TempTxt2 = tableArray(y, z)
            tableArray(x, z) = TempTxt2
            tableArray(y, z) = TempTxt1
          Next z
      End If
  Next y
Next x

'record no. of tickers per sector, so we can loop through each sector separately
n = 1
ReDim Preserve sectorTickersArray(n)
sectorTickersArray(0) = 2
For x = LBound(tableArray, 1) + 1 To UBound(tableArray, 1)
 If x = UBound(tableArray, 1) Then
           m = x
           sectorTickersArray(n) = x
     Else: m = x + 1
     End If
 If tableArray(x, colNo1) <> tableArray(m, colNo1) Then
    sectorTickersArray(n) = x
    n = n + 1
    ReDim Preserve sectorTickersArray(n)
 End If
Next x

'set in order industries using sectors tickers no
For i = LBound(sectorTickersArray) To UBound(sectorTickersArray) - 1
  For x = sectorTickersArray(i) To sectorTickersArray(i + 1)
   If x = UBound(tableArray, 1) Then
           m = x
     Else: m = x + 1
     End If
  
    For y = x To sectorTickersArray(i + 1)
     If UCase(Trim(tableArray(x, colNo1))) = UCase(Trim(tableArray(m, colNo1))) Then
     If UCase(Trim(tableArray(y, colNo2))) < UCase(Trim(tableArray(x, colNo2))) Then
      
         For z = LBound(tableArray, 2) To UBound(tableArray, 2)
            If IsError(tableArray(x, z)) Then
              tableArray(x, z) = 0
            End If
            If IsError(tableArray(y, z)) Then
              tableArray(y, z) = 0
            End If
            
            TempTxt1 = tableArray(x, z)
            TempTxt2 = tableArray(y, z)
            tableArray(x, z) = TempTxt2
            tableArray(y, z) = TempTxt1
          Next z
       End If
      End If
  Next y
 Next x
Next i

'record tickers descending order per industry per sector in the array elements, so we can set in order tickers upon PE1
n = 1
ReDim Preserve industryTickersArray(n)
industryTickersArray(0) = 2
For x = LBound(tableArray, 1) + 1 To UBound(tableArray, 1)
 If x = UBound(tableArray, 1) Then
           m = x
           industryTickersArray(n) = x
     Else: m = x + 1
     End If
 If tableArray(x, colNo2) <> tableArray(m, colNo2) Then
    industryTickersArray(n) = x
    n = n + 1
    ReDim Preserve industryTickersArray(n)
 End If
Next x


'rearrange the whole tableArray upon PE1
For i = LBound(industryTickersArray) To UBound(industryTickersArray) - 1
  For x = industryTickersArray(i) To industryTickersArray(i + 1)
     If x = UBound(tableArray, 1) Then
           m = x
     Else: m = x + 1
     End If
  
    For y = x To industryTickersArray(i + 1)
     If UCase(Trim(tableArray(x, colNo2))) = UCase(Trim(tableArray(m, colNo2))) Then
     If tableArray(y, colNo3) > tableArray(x, colNo3) Then
      
         For z = LBound(tableArray, 2) To UBound(tableArray, 2)
            If IsError(tableArray(x, z)) Then
              tableArray(x, z) = 0
            End If
            If IsError(tableArray(y, z)) Then
              tableArray(y, z) = 0
            End If
            TempTxt1 = tableArray(x, z)
            TempTxt2 = tableArray(y, z)
            tableArray(x, z) = TempTxt2
            tableArray(y, z) = TempTxt1
          Next z
       End If
      End If
  Next y
 Next x
Next i





'create "separatorArray" to hold the end row of each sector after adding 4 rows as 2 separator and 1average 1 header
'where first col of the array is the end row of sector and the second col no of rows per sector
industriesNo = 0
n = 1
ReDim separatorArray(UBound(sectorTickersArray) - 1, 1)
For i = LBound(industryTickersArray) + 1 To UBound(industryTickersArray)
     
     If industryTickersArray(i) < sectorTickersArray(n) Then
       industriesNo = industriesNo + 1
       
     ElseIf industryTickersArray(i) = sectorTickersArray(n) Then
     industriesNo = industriesNo + 1
     separatorArray(n - 1, 0) = sectorTickersArray(n) + industriesNo * 4
      
      If n = 1 Then
      separatorArray(n - 1, 1) = separatorArray(n - 1, 0) - 2
      Else
      separatorArray(n - 1, 1) = separatorArray(n - 1, 0) - separatorArray(n - 2, 0) - 1
      End If
      
        n = n + 1
     End If
      
Next i








'------------------------create array to carry individual sector-----------------------------------------------
    isNewSector = False
    x = 0  ' reset x value from above loops
    w = 1  'the row no. of temp array
    u = 0  'to tackle the effect of repeated loop as We loop i.e. from 2 to 5 then 5 to 8, so 5 here is repeated
    n = 0  ' starting point for the separatorArray to give first array length for the tempTableArray
For i = LBound(industryTickersArray) To UBound(industryTickersArray) - 1

    
    
    'give sign if we start new sector
    For k = LBound(sectorTickersArray) + 1 To UBound(sectorTickersArray)
          If sectorTickersArray(k) = x - 1 Then
            isNewSector = True
            Exit For
          End If
     Next k
     
     ' we declare tempTableArray no of rows per sector
     If i = 0 Or isNewSector = True Then
     
     
     
     'add sheet and assign the sector data to it
      If i <> 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = tempTableArray(1, 3)
        Sheets(tempTableArray(1, 3)).Range(Cells(1, 1).Address, Cells(UBound(tempTableArray, 1), 1 + UBound(tempTableArray, 2)).Address) = tempTableArray
      ' assign % format for columns of EG1 & EG2
        columnLetter1 = Split(Cells(1, colNo7).Address, "$")(1)
        columnLetter2 = Split(Cells(1, colNo8).Address, "$")(1)
        Sheets(tempTableArray(1, 3)).Range(columnLetter1 & ":" & columnLetter2).NumberFormat = "0.0%"
        'format the header row with bold and border line
        tableLength = Sheets(tempTableArray(1, 3)).Range("A500000").End(xlUp).Row
        boldRowArray = Sheets(tempTableArray(1, 3)).Range("A1:A" & tableLength)
        For r = LBound(boldRowArray) To UBound(boldRowArray)
          If UCase(Trim(boldRowArray(r, 1))) = "COMPANY NAME" Then
            Sheets(tempTableArray(1, 3)).Range(r & ":" & r).Font.Bold = True
            Sheets(tempTableArray(1, 3)).Range(r & ":" & r).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
          End If
        Next r
        'auto fit columns width
         Sheets(tempTableArray(1, 3)).Columns("A:ZZ").AutoFit
      
      End If
      
      
      Erase tempTableArray
      ReDim Preserve tempTableArray(separatorArray(n, 1), UBound(tableArray, 2) - 1)
      n = n + 1
      w = 1
      isNewSector = False
     End If
     
     'adding the header of the table to the temp array
      For v = LBound(tableArray, 2) To UBound(tableArray, 2)
        tempTableArray(w - 1, v - 1) = tableArray(1, v)
      Next v
   
  avrgArray(0) = w  'the first limit of average for an industry
  For x = industryTickersArray(i) + u To industryTickersArray(i + 1)
  u = 1
    For y = LBound(tableArray, 2) To UBound(tableArray, 2)
    
        
        
         If IsNumeric(tableArray(x, y)) Then
            tempTableArray(w, y - 1) = Round(tableArray(x, y), 2)
         Else
             tempTableArray(w, y - 1) = tableArray(x, y)
         End If
        
        
    Next y
 w = w + 1
 Next x
 
   
    avrgArray(1) = w  'end limit of average

    'create array for the required ratios to be averaged
    tempTableArray(w + 1, 0) = "Average"
    ratiosArray() = Array(colNo4, colNo5, colNo6, colNo7, colNo8, colNo3, colNo9, colNo10, colNo11)
     
     'filling the tempTableArray with the averages
    For r = LBound(ratiosArray) To UBound(ratiosArray)
        columnLetter = Split(Cells(1, ratiosArray(r)).Address, "$")(1)
        tempTableArray(w + 1, ratiosArray(r) - 1) = "=ROUND(AVERAGE(" & columnLetter & avrgArray(0) + 1 & ":" & columnLetter & avrgArray(1) & "),2)"
    Next r
 
 w = w + 4
Next i


Sheets.Add(After:=Sheets(Sheets.Count)).Name = tempTableArray(1, 3)
Sheets(tempTableArray(1, 3)).Range(Cells(1, 1).Address, Cells(UBound(tempTableArray, 1), 1 + UBound(tempTableArray, 2)).Address) = tempTableArray
' assign % format for columns of EG1 & EG2
columnLetter1 = Split(Cells(1, colNo7).Address, "$")(1)
columnLetter2 = Split(Cells(1, colNo8).Address, "$")(1)
Sheets(tempTableArray(1, 3)).Range(columnLetter1 & ":" & columnLetter2).NumberFormat = "0.0%"
'format the header row with bold and border line
tableLength = Sheets(tempTableArray(1, 3)).Range("A500000").End(xlUp).Row
        boldRowArray = Sheets(tempTableArray(1, 3)).Range("A1:A" & tableLength)
        For r = LBound(boldRowArray) To UBound(boldRowArray)
          If UCase(Trim(boldRowArray(r, 1))) = "COMPANY NAME" Then
            Sheets(tempTableArray(1, 3)).Range(r & ":" & r).Font.Bold = True
            Sheets(tempTableArray(1, 3)).Range(r & ":" & r).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
          End If
        Next r
'auto fit columns width
Sheets(tempTableArray(1, 3)).Columns("A:ZZ").AutoFit

'mytest.Range(Cells(1, 1).Address, Cells(UBound(tempTableArray, 1), UBound(tempTableArray, 2)).Address) = tempTableArray



End Sub

