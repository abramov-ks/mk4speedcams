Option Explicit
Dim ErrorText As String

Private Sub CommandButton1_Click()
    ' determine map type

    Call IDXURLwriter
End Sub


Sub IDXURLwriter()
    Dim POSWGScol As String
    Dim FieldTitleRow As Integer

    Dim fs As Object
    Dim URLfile As Object
    Dim IDXfile As Object
    Dim URLoutputLine As String
    Dim IDXoutputLine As String

    ' true = E90/91 Professional
    ' false = Mk4, NAI01, NAVI02, NAVI03 etc
    Dim Professional As Boolean ' true = E90/91 Professional

    Dim text As String

    Dim CelRefCategoryName As String
    Dim CategoryName As String
    Dim OutputPath As String
    Dim CelRefLanguage As String
    Dim POIallUpperCase As Boolean

    Dim f As Integer
    Dim i As Long
    Dim n As Integer
    Dim q As Integer
    Dim r As Long
    Dim z As Integer
    Dim maxFields As Integer
    Dim fieldLength As Integer
    Dim LastDataRow As Long
    Dim LastDataCol As String

    Dim startCol As String
    Dim colLetter As String
    Dim myRange As String
    Dim myCell As String

    Dim ErrorTitle As String

    On Error GoTo errorTrap

    ' Get critical data values
    ErrorTitle = "Error locating critical data cells"
    CelRefCategoryName = findCell("CATEGORY NAME", 1)
    CelRefLanguage = findCell("LANGUAGE", 1)
    OutputPath = findCell("DIRECTORY", 1)
    POIallUpperCase = Range(findCell("POI DATA DISPLAY", 1)).Value = "ALL CAPITALS"
    ' find starting row -
    text = findCell("LATITUDE", 0)
    FieldTitleRow = Val(Right$(text, Len(text) - 1))
    POSWGScol = Left$(findCell("POSWGS", 0), 1)

    ' determine if we need to make High or Professional output files
    Professional = True
    text = findCell("CREATE DATA FOR", 1)
    text = Range(text).Value
    Select Case UCase(text)
        Case "1-SERIES E87":     Professional = True
        Case "3-SERIES E46":     Professional = False
        Case "3-SERIES E90/E91": Professional = True
        Case "5-series E38":     Professional = False
        Case "5-series E60/61":  Professional = True
        Case "7-series E38":     Professional = False
        Case "7-series E65/66":  Professional = False
        Case "X3 E83":           Professional = False
        Case "X5 E53":           Professional = False
        Case "X5 E70":           Professional = True
        Case "Z4":               Professional = False
        Case Else:               Professional = True
    End Select

    ' Get the data extremes
    ' find last data rows -
    ' look down column A (Latitude) from FieldTitleRow until the An cell is empty
    For LastDataRow = (FieldTitleRow + 1) To 65536
        If Range("A" & LastDataRow).Value = "" Then
            LastDataRow = LastDataRow - 1
            Exit For
        End If
    Next


    ' find last data column
    For n = (Asc("A") - 1) To Asc("Z") Step 1
        For i = Asc("A") To Asc("Z") Step 2
            ' ensure first column digit is blank
            If n = Asc("A") - 1 Then colLetter = Chr(i) Else colLetter = Chr(n) & Chr(i)
            ' if no value we exit
            If Range(colLetter & FieldTitleRow).Value = "" Then
                If n = Asc("A") - 1 Then LastDataCol = Chr(i - 1) Else LastDataCol = Chr(n) & Chr(i - 1)
                Exit For
            End If
        Next
        ' if no value we exit
        If Range(colLetter & FieldTitleRow).Value = "" Then Exit For
    Next


    ' sort the data before we do anything else
    ' We sort on Longitude, not on POSWGS, to ensure proper sorting
    ErrorTitle = "Error sorting POI data"
    myRange = "A" & CStr(FieldTitleRow) & ":" & LastDataCol & CStr(LastDataRow) ' the data range
    myCell = Left$(findCell("LONGITUDE", 0), 1) & CStr((FieldTitleRow + 1)) ' the column to sort by
    Range(myRange).Sort Key1:=Range(myCell), Order1:=xlAscending, header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal



    ' now write the URL file
    ' all lines are text, terminating in CRLF

    ' fields are fixed width, padded with 0x00

    ' asc2idxurl_mk4.exe creates following files:
    ' first line terminates in crlf
    ' second line has 0x00 at end, then crlf
    ' data lines terminate in crlf

    ' original USA DVD has files in following format
    ' first line terminates in crlf YY
    ' second line has 0x00 at end, then crlf YY
    ' data lines terminate in crlf YY


    ' open file output
    ' set output filename
    ' first get cell containing IDX file number, 3 rows down from CelRefCategoryName cell
    ' URL is always 1 digit higher than IDX
    ErrorTitle = "Error creating output files"
    text = Range(findCell("DIRECTORY FOR", 1)).Value
    If Right$(text, 1) <> "\" Then text = text & "\" ' ensure ends in \
    z = Range(findCell("IDX OUTPUT", 1)).Value
    IDXoutputLine = text & Right$(String$(4, "0") & z, 4) & ".IDX"
    URLoutputLine = text & Right$(String$(4, "0") & z + 1, 4) & ".URL"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set IDXfile = fs.CreateTextFile(IDXoutputLine, True)
    Set URLfile = fs.CreateTextFile(URLoutputLine, True)

    ' write URL and IDX headers
    ' first get category name and output path from sheet
    ErrorTitle = "Error writing IDX & URL headers"
    CategoryName = UCase$(Range(CelRefCategoryName).Value)
    'OutputPath = Range(OutputPath).Value
    'OutputPath = Replace(OutputPath, "\", "/") ' convert dos to unix dirs
    'If Right(OutputPath, 1) <> "/" Then OutputPath = OutputPath & "/"
    ' now write the two headers
    '
    ' URL header
    ' Example URL header format:
    ' High:
    ' OUTPUTURL-ENG/SE/SF_OUTPUT.HTM
    ' Professional:
    ' CATURL-eng/sf_41018.htm_
    ' adapt the text, but I think this is info only, so not critical
    OutputPath = Range(findCell("PATH FOR", 1)).Value
    If Professional Then ' NO
        ' nz map shows / so replace \ with /
        OutputPath = Replace(OutputPath, "\", "/") ' convert dos to unix dirs
    Else
        ' other CarinDB maps use \
        OutputPath = Replace(OutputPath, "/", "\") ' convert dos to unix dirs
    End If
    URLoutputLine = CategoryName & "URL-" & OutputPath & "SF_" & CategoryName & ".HTM"
    URLfile.WriteLine (URLoutputLine)
    '
    ' IDX header
    ' Example IDX header format showing associated SF file:
    ' High:
    ' Glambda- CATIDX-ENG/SF_40003.HTM_  (ex USA disc)
    ' Glambda- 99IDX-ENG/SF_99.HTM  (made by asc2idxurl_mk4.exe)
    ' Professional:
    ' Gphi- CATIDX-eng/sf_41018.htm
    If Professional Then
        IDXoutputLine = "Gphi- " & CategoryName & "IDX-" & OutputPath & "SF_" & CategoryName & ".HTM"
        ' nz map shows / so replace \ with /
        OutputPath = Replace(OutputPath, "\", "/") ' convert dos to unix dirs
    Else
        IDXoutputLine = "Glambda- " & CategoryName & "IDX-" & OutputPath & "SF_" & CategoryName & ".HTM"
        ' other CarinDB maps use \
        OutputPath = Replace(OutputPath, "/", "\") ' convert dos to unix dirs
    End If
    IDXfile.WriteLine (IDXoutputLine)



    ' write URL category titles and string lengths
    ' Example URL format:
    ' POSWGS:S:21|NAME:S:74|STREET:S:24|HOUSENUMBER:S:4|CITY:S:27|ZIP:S:6|PHONE:S:15|FNM:S:0|NT:S:1|FACPAGE:S:66|_:S:1
    ' Example IDX format:
    ' ID:I:2|POS:P:8|SELNAME:S:34
    ' write field headers and count number of fields
    ' this will loop from col A to col ZZ
    ' each cell has column type "A" or "AA" etc
    ' n= higher value digit
    ' i = lower value digit
    ' eg: cell AC, n=A, i=C
    startCol = POSWGScol
    maxFields = 0
    URLoutputLine = ""
    ' set start of IDX header. POS is always 8 bytes
    IDXoutputLine = "ID:I:" & Len(CStr(LastDataRow - FieldTitleRow)) & "|POS:P:8|"
    For n = (Asc("A") - 1) To Asc("Z") Step 1
        For i = Asc(startCol) To Asc("Z") Step 2


            ' ensure first column digit is blank
            If n = Asc("A") - 1 Then colLetter = Chr(i) Else colLetter = Chr(n) & Chr(i)

            ' if no value in title we exit
            If Range(colLetter & FieldTitleRow).Value = "" Then Exit For

            ' if value starts with ! then skip this column pair
            If Left$(Range(colLetter & FieldTitleRow).Value, 1) = "!" Then GoTo HeaderiLoopEnd

            ' we're good to write, append to URL text
            URLoutputLine = URLoutputLine & Range(colLetter & FieldTitleRow).Value & "|"

            ' add name to the IDX file
            If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 4)) = "NAME" Then
                ' note that the field name in the IDX file must match
                ' the name in the SF_ file
                ' so use the user-entry from the sheet
                text = Range(colLetter & FieldTitleRow).Value
                ' extract the last ":S:n"
                text = Right$(text, InStr(1, text, ":"))
                ' add ther name defined by the user
                text = UCase$(Range(Left$(CelRefCategoryName, 1) & CInt(Right$(CelRefCategoryName, 2)) + 3).Value) & text
                IDXoutputLine = IDXoutputLine & text & "|"
            End If

            ' add BRANDNAME if desired to the IDX file
            If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 9)) = "BRANDNAME" And UCase$(Range(colLetter & FieldTitleRow - 2).Value) = "YES" Then
                IDXoutputLine = IDXoutputLine & Range(colLetter & FieldTitleRow).Value & "|"
            End If
            ' add NSW1 if desired to the IDX file
            If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 4)) = "NSW1" And UCase$(Range(colLetter & FieldTitleRow - 2).Value) = "YES" Then
                IDXoutputLine = IDXoutputLine & Range(colLetter & FieldTitleRow).Value & "|"
            End If
            ' add IMPORTANCE if desired to the IDX file
            If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 10)) = "IMPORTANCE" And UCase$(Range(colLetter & FieldTitleRow - 2).Value) = "YES" Then
                IDXoutputLine = IDXoutputLine & Range(colLetter & FieldTitleRow).Value & "|"
            End If

            maxFields = maxFields + 1
HeaderiLoopEnd:
        Next i
        ' if no value in title we exit
        If Range(colLetter & FieldTitleRow).Value = "" Then Exit For
        startCol = Chr(i - 26) ' if i=Z, next=B; if i=Y, next=A
    Next n
    ' remove last | charcater
    ' keep this here, even though we don't really need to do this for the IDX file
    ' But keep it just in case we add more stull to the IDX later
    URLoutputLine = Left$(URLoutputLine, Len(URLoutputLine) - 1)
    IDXoutputLine = Left$(IDXoutputLine, Len(IDXoutputLine) - 1)
    URLfile.WriteLine (URLoutputLine)
    IDXfile.WriteLine (IDXoutputLine)
    ' headers are all done

    ' build and write IDX indexing line
    ErrorTitle = "Error writing IDX index line"
    z = LastDataRow - FieldTitleRow ' qty of rows to write
    ' start with first entry
    ' entry = 0, first 4 bytes are long of entry
    IDXoutputLine = String(4, 0) & convertWGSPOStoP(Range(POSWGScol & FieldTitleRow + 1).Value, False) & "|"
    i = 49 ' index step value, gets every i record

    'IDXoutputLine = ""
    'For r = 0 To (z - 1) Step i
    '    ' start at row 0, step in i increments
    '    IDXoutputLine = IDXoutputLine & convertLongtoStringLong(r) & convertWGSPOStoP(Range(POSWGScol & FieldTitleRow + r).Value, False) & "|"

    'Next r
    ' dynamically allocate index step value, i
    ' Examples: USA large file = every 50
    ' USA small file = every 20
    z = z ' qty of rows
    i = z \ 5 ' integer divide, minimum 5 entries
    If i < 20 Then
        i = 20  ' <100 entries
    ElseIf i < 30 Then
        i = 30
    ElseIf i < 40 Then
        i = 40
    Else
        i = 50
    End If

    For r = i To (z - i) Step i
        ' scan all rows in sheet, index every i-th entry
        ' remember: first entry = 0, so line r = entry r-1
        IDXoutputLine = IDXoutputLine & convertLongtoStringLong(r) & convertWGSPOStoP(Range(POSWGScol & FieldTitleRow + r + 1).Value, False) & "|"
    Next
    ' finish with 2nd to last entry like asc2idxurl_mk4.exe does
    ' terminate line in 0x00, crlf added when written
    ' remove last "|"
    IDXoutputLine = Left$(IDXoutputLine, Len(IDXoutputLine) - 1)
    IDXoutputLine = IDXoutputLine & Chr$(0)
    IDXfile.WriteLine (IDXoutputLine)

    ' write IDX and URL POI data to output file
    ErrorTitle = "Error writing POI data"
    z = LastDataRow - FieldTitleRow ' qty of rows to write
    For r = (FieldTitleRow + 1) To (FieldTitleRow + z) ' scan all rows in sheet
        ' this is the ROW loop
        ' loops up to LastDataRow

        ' OK, POSWGS not blank, write the data
        ' first prepare the IDX file line
        ' get the POI entry number, first entry = 0
        IDXoutputLine = Right$(String(Len(CStr(z)), " ") & CStr(r - FieldTitleRow - 1), Len(CStr(z)))
        ' now read to enter the loop
        startCol = POSWGScol
        URLoutputLine = ""
        f = 0 ' reset z, used as field counter
        For n = (Asc("A") - 1) To Asc("Z") Step 1
            ' this is the first character column loop ie "A" in "AZ"
            For i = Asc(startCol) To Asc("Z") Step 2
                ' this is the second character column loop ie "Z" in "AZ"

                ' ensure first column digit is blank
                If n = Asc("A") - 1 Then colLetter = Chr(i) Else colLetter = Chr(n) & Chr(i)

                ' if no value in title we exit
                If Range(colLetter & FieldTitleRow).Value = "" Then Exit For

                ' if value starts with ! then skip this column pair
                If Left$(Range(colLetter & FieldTitleRow).Value, 1) = "!" Then GoTo DataiLoopEnd

                ' get length of this field
                text = Range(colLetter & FieldTitleRow).Value
                text = Right$(text, Len(text) - InStrRev(text, ":"))
                fieldLength = CInt(text)
                ' get field contents
                text = Left$(Range(colLetter & r).Value & String(fieldLength, 0), fieldLength)
                If POIallUpperCase Then text = UCase$(text)

                ' build output string
                URLoutputLine = URLoutputLine & text

                ' build IDX output string
                ' add POSWGS to the IDX file
                If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 6)) = "POSWGS" Then
                    ' add the position, in Position format
                    IDXoutputLine = IDXoutputLine & convertWGSPOStoP(text, True)
                End If
                ' add NAME to the IDX file
                If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 4)) = "NAME" Then
                    ' add the name, already prepared in 'text' above
                    IDXoutputLine = IDXoutputLine & text
                End If
                ' add BRANDNAME if desired to the IDX file
                If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 9)) = "BRANDNAME" And UCase$(Range(colLetter & FieldTitleRow - 2).Value) = "YES" Then
                    IDXoutputLine = IDXoutputLine & text
                End If
                ' add NSW1 if desired to the IDX file
                If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 4)) = "NSW1" And UCase$(Range(colLetter & FieldTitleRow - 2).Value) = "YES" Then
                    IDXoutputLine = IDXoutputLine & text
                End If
                ' add IMPORTANCE if desired to the IDX file
                If UCase$(Left$(Range(colLetter & FieldTitleRow).Value, 10)) = "IMPORTANCE" And UCase$(Range(colLetter & FieldTitleRow - 2).Value) = "YES" Then
                    IDXoutputLine = IDXoutputLine & text
                End If

DataiLoopEnd:
            Next i
            ' if no value in title we exit
            If Range(colLetter & FieldTitleRow).Value = "" Then Exit For
            startCol = Chr(i - 26) ' if i=Z, next=B; if i=Y, next=A
        Next n
        ' write this line to file
        URLfile.WriteLine (URLoutputLine)
        IDXfile.WriteLine (IDXoutputLine)
    Next

    ' all done, close the files
    URLfile.Close
    IDXfile.Close

    ' all done
    text = Range(findCell("DIRECTORY FOR", 1)).Value
    If Right$(text, 1) <> "\" Then text = text & "\" ' ensure ends in \
    MsgBox "IDX and URL files successfully generated" & vbCrLf & vbCrLf & "The files have been saved to """ & text & """", vbInformation, "POI Maker"
    Exit Sub

errorTrap:
    MsgBox ErrorTitle & " - " & Err.Description, vbExclamation, "POI Maker - Error"
End Sub


Function convertLongtoStringLong(l As Long) As String
    Dim s As String
    Dim o As String
    Dim i As Integer
    ' takes a long integer l
    ' returns a string
    ' eg: long 00 00 00 00 returns chr$(0) & chr$(0) & chr$(0) & chr$(0)
    s = Right$(String(8, "0") & Hex$(l), 8)
    ' loop string, taking value of byte (2 chars) at a time
    For i = 1 To (Len(s) - 1) Step 2
        o = o & Chr$(Val("&H" & Mid$(s, i, 2)))
    Next
    ' return value
    convertLongtoStringLong = o
End Function


Function convertWGSPOStoP(WGSPOS As String, LatLong As Boolean) As String
    ' converts WGSPOS as string to a Position value
    ' WGSPOS = WGS position as per nav
    ' eg: "1098612666,-256145450"
    ' LatLong:
    ' TRUE  = returns longitude & latitude as a 8-byte string
    ' FALSE = returns just longitude as a 4-byte string
    WGSPOS = "277391010,303789431"
    LatLong = True

    Dim i As Integer
    Dim s As String
    Dim o As String
    i = InStr(1, WGSPOS, ",") - 1 ' find loc of comma

    ' convert to a string
    ' longitude is first record, followed by latitude
    ' pad with 00's to be sure string length is 8 characters
    s = Right$(String(8, "0") & Hex$(Left$(WGSPOS, i)), 8)
    If LatLong Then ' need to add latitude
        s = s & Right$(String(8, "0") & Hex$(Right$(WGSPOS, Len(WGSPOS) - i - 1)), 8)
    End If

    ' loop string, taking value of byte (2 chars) at a time
    For i = 1 To (Len(s) - 1) Step 2
        o = o & Chr$(Val("&H" & Mid$(s, i, 2)))
    Next
    ' return value
    convertWGSPOStoP = o
End Function

Function findCell(searchText As String, colOffset As Integer) As String
    Dim r As Integer ' row
    Dim c As Integer ' column
    Dim x As Boolean
    ' looks down the first 100 rows
    ' and across columns A to Z
    ' looking for the search text
    ' returns cell reference + colOffset found
    x = False
    For c = Asc("A") To Asc("Z")
        For r = 1 To 101
            If Left$(UCase$(Range(Chr$(c) & CStr(r)).Value), Len(searchText)) = searchText Then
                x = True
                Exit For
            End If
        Next
        If x Then Exit For
    Next
    If r > 100 Then
        Err.Raise vbObjectError + 1, "POI Maker", "Cannot locate '" & searchText & "' in cells A1 to Z100"
    End If
    findCell = Chr$(c + colOffset) & CStr(r)
End Function




