''
' VBA-Benchmark v0.1
' (c) Jonathan de Vries - https://github.com/jonadv/VBA-Benchmark/
'
' Benchmark VBA Code

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (stamp As Currency) As Byte
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (freq As Currency) As Byte
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (stamp As Currency) As Byte
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (freq As Currency) As Byte
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

'returns of QPC
'- as Currency -> 304462680,3775
'- as LongLong -> 3044898189059

'returns of QPF
'With a usual QPF on windows 10 (10MHz):
'- as Currency ->     1000  =      1.000
'- as LongLong -> 10000000  = 10.000.000
'---> 10 million tics per second
'---> if freq is 10MHz then:
'   1 tic = (1 / 10 000 000) * second
'   1 tic = 0.000001 seconds
'   1 tic = 0.0001 milliseconds
'   1 tic = 0.1 microseconds
'   1 tic = 100 nanoseconds

'total tics passed
'- as Currency -> (QPC2 - QPC1) * 10000
'- as LongLong -> (QPC2 - QPC1)

'tics to seconds
' = (QPC2 - QPC1) / QPF
' = TicDiff / QPF

'ticst t to seconds s = t/(s * 1)
'tics to milliseconds = t/(s * 1e3)
'tics to microseconds = t/(s * 1e6)
'tics to nanoseconds = t/(s * 1e9)

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Enum TrackID
    id_start = 255
    id_continue = 254
    id_pause = 253
    id_resizingarray = 252
End Enum

Private freq As Currency                    'frequency is the amount of tics of the QPC per second
Private stampCount As Long                  'to keep track of postition of next stamp and stampID in arrays
Private currentArraySizes As Long           'to prevent calling Ubound(arrStamps) every time
Private arrStamp() As Currency              'stores QPC stamps. (0) = start
Private arrStampID() As Byte                'stores id numbers of track calls. Byte = 0-255, so max 256 tracks, using Byte forces ID of 0 or above
Private dicStampName_ID As Dictionary       'key = custom name, value = StampID
Private Const fromCurr As Currency = 10000  'QPC and QPF downscales LongLong (actual returntype) with 10000 when return value to Currency datatype
Private time_start As Double
Private time_end As Double
Private Const overheadTestCount As Long = 100 'Overhead is tested in a loop. Lowering this ammount a lot might increase overhead because of CPU branching.


'include CTimer.Pause (stores timestamp) and .Continue (should give error when previous call wasnt a pause. Might need track ID as argument?).
'include setting time unit of output (nanos, millis, seconds, etc). calcualte default at end by total time passed
'include report to have cycle counts as well as mean, median and std
'include accurate Wait function (https://web.archive.org/web/20160324085802/http://vba.tips/2015/05/precision-time-delay/)

' ============================================= '
' Class specific Functions
' ============================================= '
' @Class_Initialize
' @Class_Terminate
Private Sub Class_Initialize()
    QueryPerformanceFrequency freq                  'frequency is set at computer boot, does not change after that
    freq = freq * fromCurr                          'scale from Currency to whole number
    Set dicStampName_ID = New Dictionary
    dicStampName_ID.CompareMode = vbBinaryCompare   'faster then vbTextCompare, but difference in capital letters will matter
    Start                                           'Start stores the first QPC, which is filtered out in Sub Report
End Sub
Private Sub Class_Terminate()
    Report                                          'print report with default settings when code is finished (to debug immediate window)
End Sub

' ============================================= '
' Public Functions
' ============================================= '
' @TrackByID        - Store QPC (cycle counts) in an array
' @TrackByName      - Same as @TrackByID but more convenient (and thus with a bit more overhead)
' @Start            - (Start) or (Reset and Restart) benchmark
' @Pause            - Convenience method to exclude pieces of code, use in combination with .Continue
' @Continue          - Use after calling .Pause to continue tracking
' @Report           - Generate report with default settings
' ReportCustomized  - Generate report with specifewied settings

Public Sub TrackByID(ByVal IDnr As Byte)
    'the fastest possible way (as in with least amount of overhead) to store
    'cpu stamps of QPC function is to store them in an array

    stampCount = stampCount + 1
    
    'store cpu stampcount in array
    QueryPerformanceCounter arrStamp(stampCount)
    
    'store id nr in seperate samesized array
    arrStampID(stampCount) = IDnr
    
    
    If stampCount = currentArraySizes Then
        'required to prevent array out of bound (it's either this
        'if-then or set the arrays to (large) fixed sizes (and still
        'get out of bound error when code is running for longer time))
    
        RedimStampArrays
        
        'redim can be time consuming so exclude this from recording.
        TrackByID TrackID.id_resizingarray
    End If
End Sub
Public Sub TrackByName(ByVal strTrackName As String)
    'intermediate/more convenient way to call track method
    'if TrackById and TrackByName are used mixed, some tracks might write to the same ID
    'reference type ByVal can save a few clock cycles https://stackoverflow.com/questions/408101/which-is-faster-byval-or-byref
    
    'when count = 0, it adds an IDnr of 0, count = 1 adds IDnr 1, etc
    If Not dicStampName_ID.Exists(strTrackName) Then dicStampName_ID.Add strTrackName, CByte(dicStampName_ID.Count)
    
    'gets IDnr and passes it as argument when calling TrackById
    TrackByID dicStampName_ID(strTrackName)

End Sub

Public Sub Start()
    Reset 're-initialize all
    time_start = CurrentTimeMillis 'accurate system timestamp in milliseconds
    TrackByID TrackID.id_start
End Sub
Public Sub Pause()
'Use in combination with .Continue to exclude code from being benched.
'Is only included in report if boExtendedReport is set to True
    TrackByID TrackID.id_pause
End Sub
Public Sub Continue()
'Use in combination with .Pause to exclude code from being benched.
'Is only included in report if boExtendedReport is set to True
    TrackByID TrackID.id_continue
End Sub
Public Sub Report()
'Calculate and output report with default settings.
'Can be called from Immediate window or while running code
    ReportArg
End Sub

Public Sub ReportCustomized(Optional ByVal boExtendedReport As Boolean = False, _
                        Optional ByVal boTransposeReport As Boolean = True, _
                        Optional ByVal OutputToRange As Range = Nothing, _
                        Optional ByVal boCorrectOverhead As Boolean = False, _
                        Optional ByVal boForceMillis As Boolean = False)
                        
    ReportArg boExtendedReport, boTransposeReport, OutputToRange, boCorrectOverhead, boForceMillis
End Sub

' ============================================= '
' Private Functions - Specific helpers
' ============================================= '
' @Reset
' @RedimStampArrays
' @ReportArg
Private Sub Reset()
'Make sure frequency is set right (in cace instance of this class is declared public static)
    QueryPerformanceFrequency freq
    freq = freq * fromCurr
    
'Set to private as public method .Start does the same
    stampCount = 0
    currentArraySizes = 0
    RedimStampArrays
    dicStampName_ID.RemoveAll
End Sub
Private Sub RedimStampArrays()
    Dim enlargementstep As Long: enlargementstep = 262144# '2^18
    'array size in memory = 20 bytes + 4 per dimension + bytes of elements. LongLong and Currency are both 8 byte per element.
    'first call of this sub sets memory usage of both arrays to (20+4+262,144*8=) 2,097,176 byte (2mb).
    'Every call enlarges it with 2 mb as well. The size of an array in memory does not impact the speed
    'of writing values to it as long as it stays in RAM. When the array becomes larger then available RAM,
    'values are written to disk memory, which is time consuming.
    currentArraySizes = currentArraySizes + enlargementstep
    If currentArraySizes = enlargementstep Then 'at start/initalisation or when timer is reset
        'erases values in arrays
        ReDim arrStamp(1 To currentArraySizes)
        ReDim arrStampID(1 To currentArraySizes)
    Else
        'keeps values in arrays
        ReDim Preserve arrStamp(1 To currentArraySizes)
        ReDim Preserve arrStampID(1 To currentArraySizes)
    End If
End Sub


Private Sub ReportArg(Optional ByVal boExtendedReport As Boolean = False, _
                        Optional ByVal boTransposeReport As Boolean = True, _
                        Optional ByVal OutputToRange As Range = Nothing, _
                        Optional ByVal boForceMillis As Boolean = False, _
                        Optional ByVal boForceNanos As Boolean = False)
                        
'dont generate report if it was generated less then 10 seconds ago (f.e. when ReportCustumized was called at end of code)
If time_end > 0 And time_end - CurrentTimeMillis > 10000 Then GoTo theEnd

If stampCount < 2 Then GoTo theEnd 'Nothing to report when only .Start (1 stamp) was called

time_end = CurrentTimeMillis 'accurate timestamp of code end (= Class_Terminate = end of code)


'calculate tic-differences (TicDiffs) per Track-call and store in evenly sized array
Dim i As Long, dFirstLast As New Dictionary
Dim arTicDiffs() As Currency: ReDim arTicDiffs(LBound(arrStamp) To stampCount)
For i = LBound(arrStamp) To stampCount 'LBound always is start-stamp
    If boExtendedReport And False Then 'Falsed out -> For later uses, when using RDTSC instead of QPC
        If Not dFirstLast.Exists(arrStampID(i) & "_fst") Then dFirstLast.Add arrStampID(i) & "_fst", arrStamp(i) * fromCurr
        dFirstLast.Item(arrStampID(i) & "_lst") = arrStamp(i) * fromCurr
    End If
    If i = LBound(arrStamp) Then arTicDiffs(i) = 0 Else arTicDiffs(i) = (arrStamp(i) - arrStamp(i - 1)) * fromCurr 'upscale to whole number
Next i
    
'seperate TicDiffs into ID-specific collection (most time consuming step in this sub)
Dim dID_colTicDiffs As New Dictionary 'key = IDnr, value = collection of time recordings (tics) per IDnr
Set dID_colTicDiffs = ticsToCollectionsInDictionaryPerID(arTicDiffs, LBound(arTicDiffs))
'example result in jsonformat: {"255":[0],"1":[156],"2":[675,766,523,764,651]}

'filter out any unwanted output here
dID_colTicDiffs.Remove TrackID.id_start & "" 'start tic value is always 0, so always filter out
If Not boExtendedReport Then
    If dID_colTicDiffs.Exists(TrackID.id_continue & "") Then dID_colTicDiffs.Remove TrackID.id_continue & ""
    If dID_colTicDiffs.Exists(TrackID.id_pause & "") Then dID_colTicDiffs.Remove TrackID.id_pause & ""
    If dID_colTicDiffs.Exists(TrackID.id_resizingarray & "") Then dID_colTicDiffs.Remove TrackID.id_resizingarray & ""
End If
'example of filtering your own defined tracks:
'dID_colTicDiffs.Remove dicStampName_ID("Initialisation")
'dicStampName_ID.Remove "Initialisation"

Dim dAll As New Dictionary, keys_ids As Variant, col_item As Variant, v As Variant
'check if TrackByName method is used and store names
'do before overhead calculations as overhead-test checks for names used
If dicStampName_ID.Count > 0 Then
    For Each v In dicStampName_ID.keys()
        dAll.Item(dicStampName_ID(v) & "_Name") = v
    Next v
End If

'UDT's in VBA can't be stored in a collection/dictionary inside a class module,
'Instead output values are stored in a dictionary with the key being the "id" concatenated with the "_Valuetype".
'After this the "Valuetype" becomes the header-name of the output table.
'This way the output only has to be added/adjusted at one place, instead of at calculation ánd report-output functions.
'Another option would be ADO Recordset, but that would require an additional Tools reference. Or just an array
'of UDT's, but that would require adjustments on 3 places: at top of the class, at calculation and
'at report formatting. In current set up, these three things are done at the same place.
Dim ticsOverhead As Double
Dim cntTic As Double, minTic As Double, maxTic As Double, sumTics As Double, avrTics As Double, elapsedTot As Double
Dim sumAllTics As Double: sumAllTics = 0
Dim cntAllTics As Double: cntAllTics = 0
Dim colTicDiffs As New Collection, key_idnr As Variant, seconds As Double
For Each key_idnr In dID_colTicDiffs.keys 'loop all identical IDnrs
    
    dAll.Item(key_idnr & "_IDnr") = key_idnr
    
    'overwrite names of the TrackID's this class uses.
    Select Case key_idnr
        Case TrackID.id_start:          dAll.Item(TrackID.id_start & "_Name") = "(Start)"           'Initialisation
        Case TrackID.id_pause:          dAll.Item(TrackID.id_pause & "_Name") = "(Before Pause)"    'Pause start
        Case TrackID.id_continue:       dAll.Item(TrackID.id_continue & "_Name") = "(Continue)"     'After Pause/Paused
        Case TrackID.id_resizingarray:  dAll.Item(TrackID.id_resizingarray & "_Name") = "(Resizing)" 'Resizing arStampID and arStamp
    End Select
    
    Set colTicDiffs = dID_colTicDiffs(key_idnr)
    cntTic = 0: minTic = 1E+15: maxTic = 0: sumTics = 0: avrTics = 0
    'break here to see the cpu tic-differences in between TrackBy calls
    For Each col_item In colTicDiffs
        v = col_item 'ammount of tics
        cntTic = cntTic + 1
        minTic = Min(minTic, v)
        maxTic = Max(maxTic, v)
        sumTics = sumTics + v
    Next col_item
    sumAllTics = sumAllTics + sumTics
    cntAllTics = cntAllTics + cntTic
    
    v = key_idnr 'IDnr
    dAll.Add v & "_Count", FormatNumber(cntTic, 0)
    dAll.Add v & "_Sum of tics", FormatNumber(sumTics, 0)
    dAll.Add v & "_Percentage", "" 'value cant be calculated yet as total sum is yet unknown, but already place in output table
    dAll.Add v & "_Time sum", secondsProperString(ticsToSeconds(sumTics), boForceMillis, boForceNanos)
    
    If Not boExtendedReport Then GoTo nextV_SkipExtendedOutput
' ----------------- Output for extended report: -----------------
    
    dAll.Add v & "_Minimum", FormatNumber(minTic)
    dAll.Add v & "_Maximum", FormatNumber(maxTic)
    dAll.Add v & "_Average", FormatNumber(sumTics / cntTic)
    dAll.Add v & "_Median", FormatNumber(MedianOfFirst_x_Elements(colTicDiffs, 1000)) 'Only from first 1000 tic measurements
'    dAll.Add v & "_ElapsedTot", (dFirstLast(v & "_lst") - dFirstLast(v & "_fst"))
    
'overhead
'Standard TrackID's (fe id_pause) test to False here as there isnt a string name in
'dicStampName_ID for them (even though they are already added to dAll with a name).
    If dicStampName_ID.Exists(dAll(v & "_Name")) Then 'if TrackByName used
        ticsOverhead = OverheadPerTrackCall(v, "ByNameMin")
        dAll.Add v & "_Overhead Min", FormatNumber(ticsOverhead, 0)
        dAll.Add v & "_Overhead Avr", FormatNumber(OverheadPerTrackCall(dAll(v & "_Name"), "ByNameAvr"))
    Else
        ticsOverhead = OverheadPerTrackCall(v, "ByIDMin")
        dAll.Add v & "_Overhead Min", FormatNumber(ticsOverhead, 0)
        dAll.Add v & "_Overhead Avr", FormatNumber(OverheadPerTrackCall(v, "ByIDAvr"))
    End If
    
'corrected values
    dAll.Add v & "_Sum (cor)", FormatNumber(sumTics - (ticsOverhead * cntTic), 0)
    If cntTic > 1 Then
        dAll.Add v & "_Avr (cor)", FormatNumber(sumTics / cntTic - ticsOverhead, 2)
        dAll.Add v & "_Min (cor)", FormatNumber(minTic - ticsOverhead, 0)
        dAll.Add v & "_Max (cor)", FormatNumber(maxTic - ticsOverhead, 0)
    End If

'time values
    dAll.Add v & "_Time avr", secondsProperString(ticsToSeconds(avrTics - ticsOverhead * 0), boForceMillis, boForceNanos)
nextV_SkipExtendedOutput:
Next key_idnr

'restores global variables, does nothing if not called before
v = OverheadPerTrackCall(v, "restore")

'calculate percentage per ID, now that sumAllTics is known
For Each key_idnr In dID_colTicDiffs.keys 'all identical IDnrs
    v = key_idnr
    dAll.Item(v & "_Percentage") = FormatPercent(dAll.Item(v & "_Sum of tics") / sumAllTics)
Next key_idnr

'calculate totals
dAll.Item("TOTAL" & "_Name") = "TOTAL"
dAll.Item("TOTAL" & "_Count") = FormatNumber(cntAllTics, 0)
dAll.Item("TOTAL" & "_Sum of tics") = FormatNumber(sumAllTics, 0)
dAll.Item("TOTAL" & "_Percentage") = FormatPercent(dAll.Item("TOTAL" & "_Sum of tics") / sumAllTics)
dAll.Item("TOTAL" & "_Time sum") = secondsProperString(ticsToSeconds(sumAllTics), boForceMillis, boForceNanos)

If boExtendedReport Then
    dAll.Item("TOTAL" & "_Average") = Round(sumAllTics / cntAllTics, 0)
End If

'dAll now holds all the values for the report. key = IDnr_ValueType, value = value

'add unique headers for output table
Dim dHeaders As New Dictionary
dHeaders.Add "IDnr", 1 'makes sure IDnr is first/most left column
For Each v In dAll.keys
    dHeaders.Item(RIGHT_AfterLastCharsOf(v, "_")) = 0
Next v

Dim arrReport() As Variant
Dim header As Variant, id As String, c As Long, r As Long: c = 0: r = 0 'column, row
ReDim arrReport(1 To dHeaders.Count, 1 To dID_colTicDiffs.Count + 1 + 1) 'arrReport(headers, datarows + headerrow + totalsrow)

'Debug.Print JsonConverter.ConvertToJson(dAll)
For i = -1 To 256 'Byte range is 0-255, minimum ID = 0, nr -1 for headers, nr 256 for TOTAL, sorted order (id_pause etc as last)
    Select Case i
        Case -1:                'headers
        Case 256:               id = "TOTAL"
        Case Else
            id = i & ""
            If Not dID_colTicDiffs.Exists(id) Then GoTo nextI
    End Select
    r = r + 1
    c = 0
    
    For Each header In dHeaders.keys()
        c = c + 1
        If r = 1 Then
            arrReport(c, r) = header
        Else
            If dAll.Exists(id & "_" & header) Then
                arrReport(c, r) = dAll.Item(id & "_" & header)
            Else
                arrReport(c, r) = ""
            End If
        End If
    Next header
nextI:
Next i


'check if table has more columns then rows, if so transpose
'If dHeaders.Count > UBound(arrReport, 2) Or dHeaders.Count > 8 Then arrReport = Transpose2DArray(arrReport)
'If dHeaders.Count > 8 Or boTransposeReport Then arrReport = Transpose2DArray(arrReport)
Array2DToImmediate (arrReport)

theEnd:
Debug.Print "Total time since Start:    " & secondsProperString((time_end - time_start) / 1000)
Debug.Print "Time to calculate report:  " & secondsProperString((CurrentTimeMillis - time_end) / 1000)
Debug.Print "Max precision:             " & secondsProperString(Precision)
Debug.Print ""
End Sub

' ============================================= '
' Private Functions - Specific Helpers
' ============================================= '
' @OverheadPerTrackCall
' @OverheadPerQPCcall
' @ticsToCollectionsInDictionaryPerID
' @ticsToSeconds
' @secondsProperString
' @MaxAccuracy

Private Function OverheadPerTrackCall(NameOrID As Variant, action As String) As Double
'calculates the overhead in amount of tics to call methods TrackByID and TrackByName.
'As these two methods adjust (values in) global variables, these global variables
'are used to calculate the overhead. They are first copied and stored as Static, which
'prevents the stamp-arrays from being copied every time an ID or Name is tested.

Dim frst_loop As Long: frst_loop = 1
Dim last_loop As Long: last_loop = frst_loop + Max(overheadTestCount, 1)

'create global arrays only once/statically and keep them in memory in between calls
'Static to keep them alive in between function calls/as long as code is running.
Static stampCountTemp As Long
Static arrStampTemp() As Currency
Static arrStampIDTemp() As Byte

'copy global variables to temps only once
If stampCountTemp = 0 Then 'only 0 when initialized
    stampCountTemp = stampCount
    arrStampTemp = arrStamp
    arrStampIDTemp = arrStampID
End If

stampCount = 0

Dim i As Long, id As Byte, name As String
Select Case action
    Case "ByIDAvr", "ByIDMin"
        id = CByte(NameOrID)
        For i = frst_loop To last_loop
            TrackByID id
        Next i
        
    Case "ByNameAvr", "ByNameMin"
        name = NameOrID
        For i = frst_loop To last_loop
            TrackByName name
        Next i
        
    Case Else '"restore"
    
End Select

Select Case action
    Case "ByIDAvr", "ByNameAvr" 'average
        OverheadPerTrackCall = (arrStamp(last_loop) - arrStamp(frst_loop)) * fromCurr / last_loop
        Exit Function
        
    Case "ByIDMin", "ByNameMin" 'minimum
        Dim minval As Double
        minval = 1E+15
        For i = frst_loop To last_loop - 1
            minval = Min(minval, CDbl((arrStamp(i + 1) - arrStamp(i)) * fromCurr))
        Next i
        OverheadPerTrackCall = minval
        Exit Function
        
    Case Else '"restore"
        If stampCountTemp = 0 Then 'check if OverheadPerTrackCall was used before (could be if not boExtendedReport)
            OverheadPerTrackCall = 0
            Exit Function
        End If
        'restore global variables, erase static ones
        stampCount = stampCountTemp
        arrStamp = arrStampTemp
        arrStampID = arrStampIDTemp
        stampCountTemp = 0
        ReDim arrStampTemp(0) 'erase/free memory
        ReDim arrStampIDTemp(0)
        OverheadPerTrackCall = OverheadPerQPCcall
        Exit Function
        
End Select

End Function

Private Function OverheadPerQPCcall() As Double
'calculates (average) time it takes to call QPC function directly.
'Does not include overhead of TrackByID or TrackByName (to look up IDnr from dictionary).

Dim arr() As Currency: ReDim arr(1 To overheadTestCount)
Dim i As Long

For i = LBound(arr) To UBound(arr)
    QueryPerformanceCounter arr(i)
Next i
OverheadPerQPCcall = (arr(UBound(arr)) - arr(LBound(arr))) * fromCurr / overheadTestCount
End Function
Private Function Precision() As Double
'Tick Interval = 1/(Performance Frequency) = Resolution
Dim resolution As Double
resolution = 1 / freq

'ElapsedTime = Ticks * Tick Interval = AccessTime
Dim accesTime As Double
accesTime = OverheadPerQPCcall * resolution

Precision = Max(resolution, accesTime)

End Function

Private Function ticsToCollectionsInDictionaryPerID(ByRef arTdifs() As Currency, ByVal lb As Long) As Dictionary
  Set ticsToCollectionsInDictionaryPerID = New Dictionary
  
  Dim offset As Long
  For offset = 0 To stampCount - 1
'    If dFilteredIDnrs.Exists(arrStampID(lb + offset)) Then GoTo next_item:
    
    On Error GoTo new_item
    ticsToCollectionsInDictionaryPerID.Item(LTrim$(str$(arrStampID(lb + offset)))).Add arTdifs(lb + offset)
    On Error GoTo 0
'next_item:
  Next
  
  Exit Function
  
new_item:
  Set ticsToCollectionsInDictionaryPerID.Item(LTrim$(str$(arrStampID(lb + offset)))) = New Collection
  Resume
End Function

Private Function ticsToSeconds(ByVal tics As Currency) As Double
    If Int(tics) <> tics Or Int(freq) <> freq Then err.Raise 9999999, , "QPC or QPF returns with datatype As Currency downscales the returns with 10 000. Upscale both returns before calling this funciton."
    ticsToSeconds = tics / freq 'time in seconds
End Function
Private Function secondsProperString(ByVal t As Double, _
                Optional ByVal boForceMilliSeconds As Boolean = False, _
                Optional ByVal boForceNanoSeconds As Boolean = False) As String
                
If boForceNanoSeconds Then boForceMilliSeconds = False
Dim res As String

If t >= 3599 And Not boForceMilliSeconds And Not boForceNanoSeconds Then    'more then 1 hour
    res = VBA.Format$(t / 3600 / 24, "HH:mm:ss")
    
ElseIf t > 599 And Not boForceNanoSeconds And Not boForceMilliSeconds Then  'more then 10 minutes
    res = VBA.Format$(t / 3600 / 24, "mm:ss")
    
ElseIf t > 120 And Not boForceNanoSeconds And Not boForceMilliSeconds Then  '2 minutes
    res = Round(t, 1) & " s"
    
ElseIf t > 10 And Not boForceNanoSeconds And Not boForceMilliSeconds Then   '10 seconds
    res = Round(t, 1) & " s"
    
ElseIf t > 0.9 And Not boForceNanoSeconds And Not boForceMilliSeconds Then  '0.9 second
    res = Round(t, 2) & " s"
    
ElseIf t > (10 / 1000#) And Not boForceNanoSeconds Or boForceMilliSeconds Then   'millisecond (1 ms = 10-E3 s)
    res = Round(t * 1000#, IIf(boForceMilliSeconds, 2, 0)) & " ms"
    
ElseIf t > (1 / 1000#) And Not boForceNanoSeconds Then                         'millisecond (1 ms = 10-E3 s)
    res = Round(t * 1000#, 2) & " ms"
    
ElseIf t > (10 / 1000000#) And Not boForceNanoSeconds Then                       'microsecond (1 us = 10-E6 s)
    res = Round(t * 1000000#) & " us"
    
ElseIf t > (10 / 1000000000#) Or boForceNanoSeconds Then                         'nanosecond (1 ns = 10-E9 s)
    res = Round(t * 1000000000#) & " ns"

'Any value below this is probably below the maximum preciscion of the QPC function (and likely cause of overhead correction).
'max precision: 1 / frequency * QPC overhead.

ElseIf t > (10 / 1000000000000#) Then res = Round(t * 1000000000000#) & " ps"   'picosecond (1 ps = 10-E12 s)
ElseIf t > (10 / 1E+15) Then res = Round(t * 1E+15) & " fs"                     'femtosecond (1 fs = 10-E15 s)
ElseIf t > (10 / 1E+18) Then res = Round(t * 1E+18) & " as"                     'attosecond (1 as = 10-E18 s)
ElseIf t > (10 / 1E+21) Then res = Round(t * 1E+21) & " zs"                     'zeptosecond (1 as = 10-E21 s) -> shortest time ever measured was 247 zeptoseconds
ElseIf t > (10 / 1E+24) Then res = Round(t * 1E+24) & " ys"                     'yoctosecond (1 as = 10-E24 s)
'"For Decimal expressions, any fractional value less than 1E-28 might be lost." (.net docs)
'https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/operators/comparison-operators
ElseIf t < 0 Then
    res = "<0"
    'happens when overhead correction is larger then actual tics passed (to lower this chance use minimum overhead instead of average).
    'output extended report to see corrected timevalues
ElseIf t = 0 Then
    res = "0"
    'should only happen with trackid.id_start
Else
    res = "-?-"
    'really fast pc or just err?
    Debug.Assert False
End If
secondsProperString = res
End Function


' ============================================= '
' Private Functions - General Helpers
' ============================================= '
' @Min
' @Max
' @MedianOfFirst_x_Elements
' @CurrentTimeMillis
' @RIGHT_AfterLastCharsOf
' @Array2DToImmediate

Private Function Min(ByVal x As Double, ByVal y As Double) As Double
    If x < y Then Min = x Else Min = y
End Function
Private Function Max(ByVal x As Double, ByVal y As Double) As Double
    If x > y Then Max = x Else Max = y
End Function

Private Function MedianOfFirst_x_Elements(col As Collection, x As Long) As Double 'MedianFromCollection
'puts specified amount of values of collection into an array, quicksorts
'the array and then takes out the Median value.
    Dim c  As Long: c = IIf(x > col.Count, col.Count, x) 'sorting large collection is time consuming so take minimum
    Dim ar() As Variant
    ReDim ar(1 To c)
    Dim i As Long
    For i = 1 To c  'col.count
        ar(i) = col(i)
    Next i
    QuickSortArray ar, LBound(ar), UBound(ar)
    MedianOfFirst_x_Elements = ar((LBound(ar) + UBound(ar)) \ 2) 'backslash rounds nr
End Function
Private Function QuickSortArray(ByRef vArray As Variant, inLow As Long, inHi As Long) 'recursive
'https://stackoverflow.com/a/152325/6544310
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray((inLow + inHi) \ 2)
    
    While (tmpLow <= tmpHi)
       While (vArray(tmpLow) < pivot And tmpLow < inHi)
          tmpLow = tmpLow + 1
       Wend
    
       While (pivot < vArray(tmpHi) And tmpHi > inLow)
          tmpHi = tmpHi - 1
       Wend
    
       If (tmpLow <= tmpHi) Then
          tmpSwap = vArray(tmpLow)
          vArray(tmpLow) = vArray(tmpHi)
          vArray(tmpHi) = tmpSwap
          tmpLow = tmpLow + 1
          tmpHi = tmpHi - 1
       End If
    Wend
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Function
Private Function CurrentTimeMillis() As Double
    ' Returns the milliseconds from 1970/01/01 00:00:00.0 to system UTC
    Dim st As SYSTEMTIME
    GetSystemTime st
    Dim t_Start, t_Now
    t_Start = DateSerial(1970, 1, 1)
    t_Now = DateSerial(st.wYear, st.wMonth, st.wDay) + _
        TimeSerial(st.wHour, st.wMinute, st.wSecond)
    CurrentTimeMillis = DateDiff("s", t_Start, t_Now) * 1000 + st.wMilliseconds
End Function

Private Function RIGHT_AfterLastCharsOf(ByVal strLeft As String, ByVal chars As String) As String
Dim s() As String
s = Split(strLeft, chars, -1, vbBinaryCompare)
RIGHT_AfterLastCharsOf = s(UBound(s))
End Function


Private Sub Array2DToImmediate(ByVal arr As Variant)
'Prints a 2D-array of values to a table (with same sized column widhts) in the immmediate window

'Each character in the Immediate window of Visual Basic (CTRL+G to show) has the same pixel width,
'thus giving the option to output a proper looking 2D-array (a table with variable string lenghts).

'settings
Dim spaces_between_collumns As Long: spaces_between_collumns = 2
Dim boOutlineRight As Boolean: boOutlineRight = True
Dim NrOfColsToNotOutlineRight As Long: NrOfColsToNotOutlineRight = 2 'IDnr and Name

Dim i As Long, j As Long
Dim arrMaxLenPerCol() As Long
ReDim arrMaxLenPerCol(UBound(arr, 1))
For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(arr, 2) To UBound(arr, 2)
            'determine max stringlength per column
            arrMaxLenPerCol(i) = IIf(Len(arr(i, j)) > arrMaxLenPerCol(i), Len(arr(i, j)), arrMaxLenPerCol(i))
    Next j
Next i

Dim str As String
ReDim arrValLen(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
For j = LBound(arr, 2) To UBound(arr, 2)
    For i = LBound(arr, 1) To UBound(arr, 1)
        If boOutlineRight And i > NrOfColsToNotOutlineRight Then 'except for 1st column
            On Error Resume Next
            str = str & space$((arrMaxLenPerCol(i) - Len(arr(i, j)) + spaces_between_collumns) * 1) & arr(i, j)
        Else
            On Error Resume Next
            str = str & arr(i, j) & space$((arrMaxLenPerCol(i) - Len(arr(i, j)) + spaces_between_collumns) * 1)
        End If
    Next i
    'capacity of Immediate window is about 200 lines of max 1021 characters per line.
    str = Left(str, 1020) & vbNewLine
Next j

'capacity of Immediate window is about 200 lines of max 1021 characters per line.
Debug.Print Left$(str, (190 * 1021#))
End Sub

Private Function Transpose2DArray(arr() As Variant) As Variant()
Dim arTemp() As Variant, c As Long, r As Long
ReDim arTemp(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
For r = LBound(arTemp, 1) To UBound(arTemp, 1)
    For c = LBound(arTemp, 2) To UBound(arTemp, 2)
        arTemp(r, c) = arr(c, r)
    Next c
Next r
Transpose2DArray = arTemp
End Function


