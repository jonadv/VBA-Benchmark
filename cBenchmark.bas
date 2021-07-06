''
' VBA-Benchmark v0.1
' Jonathan de Vries - https://github.com/jonadv/VBA-Benchmark/
'
' Benchmark VBA Code

Option Explicit

#If VBA7 Then   '64 bit
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (stamp As Currency) As Byte
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (freq As Currency) As Byte
#Else           '32 bit
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (stamp As Currency) As Byte
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (freq As Currency) As Byte
#End If
   
'About the chosen datatype: The fastest possible way (as in with least amount of overhead)
'to store qpc stamps is to store them in an array as datatype Currency (or as LongLong).
'Storing the stamps as a UDT (fe LARGE_INTEGER.lowpart and .highpart) takes much longer,
'as basically two 'seperate' primary datavalues have to be stored. Surprisingly there is no difference
'in time at all in using either LongLong or Currency, however, datatype LongLong is not availabe
'on 32-bit machines. Reverting of Currency-returns is required, but that is done after benchmarking
'finished so it will not effect the benchmark results.

'used definitions:
'QPC            QueryPerformanceCounter
'stamp          returnvalue from QPC, which is an accurate 'time'-stamp since computer has been boot
'QPF            QueryPerformanceFrequency
'frequency      the amount of QPC-cycles per second, nowadays usually 10MHz on Windows 10 but can differ per machine
'tic            difference between two QPC time stamps
'RDTSC          Read Time Stamp Counter, an even more accurate way to benchmark code, but for VBA it would require a custommade .dll

Private freq As Currency                    'frequency is the amount tics per second
Private stampCount As Long                  'to keep track of postition of next stamp and stampID in arrays
Private currentArraySizes As Long           'to prevent calling Ubound(arrStamps) every time
Private arrStamp() As Currency              'stores QPC stamps
Private arrStampID() As Byte                'stores id numbers of track calls. Byte = 0-255, so max 256 tracks, using Byte forces ID of 0 or above
Private dicStampName_ID As Dictionary       'key = custom name, value = StampID
Private Const fromCurr As Currency = 10000  'QPC and QPF downscale LongLong (actual returntype) with 10000 when they return a value with datatype Currency
Private stamp_ReportEnd As Currency         'is set at end of Report calculation and prevents printing the report when it was less then x amount of seconds ago
Private Const overheadTestCount As Long = 100 'Overhead is tested in a loop. Lowering this ammount a lot might increase overhead because of CPU branching.

Private Enum TrackID
    id_start = 255
    id_continue = 254
    id_pause = 253
    id_resizingarray = 252
End Enum

'include setting time unit of output (nanos, millis, seconds, etc). calcualte default at end by total time passed

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
' @TrackByName      - Same as @TrackByTheID but more convenient (and thus with a bit more overhead)
' @TrackByTheID     - Store QPC (cycle counts) in an array
' @Start            - (Start) or (Reset and Restart) benchmark
' @Pause            - Convenience method to exclude pieces of code, use in combination with .Continue
' @Continue         - Use after calling .Pause to continue tracking
' @Report           - Generate report
' @Sleep            - timeout code, alternative for Application.Wait
' @Wait             - same as method Sleep

Public Sub TrackByName(ByVal strTrackName As String)
    'intermediate/more convenient way to call track method (but a few cycles slower)
    'if TrackByTheID and TrackByName are used mixed, some tracks might write to the same ID
    'reference type ByVal can save a few clock cycles https://stackoverflow.com/questions/408101/which-is-faster-byval-or-byref
    
    'when count = 0, it adds an IDnr of 0, count = 1 adds IDnr 1, etc
    If Not dicStampName_ID.Exists(strTrackName) Then dicStampName_ID.Add strTrackName, CByte(dicStampName_ID.Count)
    
    'gets IDnr and passes it as argument when calling TrackByTheID
    TrackByTheID dicStampName_ID(strTrackName)

End Sub
Public Sub TrackByTheID(ByVal IDnr As Byte)
    'if it runs into an error here, you probably tried to pass a string data type
    
    'sub was called TrackByID before, but then intellisense shows it as first option/above TrackByName
    'when only typing 'tr'. This way typing 'tr' + tab should be enough.
    
    stampCount = stampCount + 1
    
    'store cpu stampcount in array
    QueryPerformanceCounter arrStamp(stampCount)
    
    'store id nr in seperate samesized array
    arrStampID(stampCount) = IDnr
    
    'check array sizes for next stamp
    If stampCount = currentArraySizes Then
        'required to prevent array out of bound (it's either this
        'if-then or set the arrays to (large) fixed sizes (and still
        'get out of bound error when code is running for longer time))
    
        RedimStampArrays
        
        'redim can be time consuming so exclude this from recording.
        TrackByTheID TrackID.id_resizingarray
    End If
End Sub

Public Sub Start()
    Reset 're-initialize all
    TrackByTheID TrackID.id_start
End Sub
Public Sub Pause()
'Use in combination with .Continue to exclude code from being tracked
'Is only included in output of report if boExtendedReport is set to True
    TrackByTheID TrackID.id_pause
End Sub
Public Sub Continue()
'Use in combination with .Pause to exclude code from being tracked
'Is only included in output of report if boExtendedReport is set to True
    TrackByTheID TrackID.id_continue
End Sub
Public Sub Report(Optional ByVal boExtendedReport As Boolean = False, _
                    Optional ByVal boTransposeReport As Boolean = False, _
                    Optional ByVal OutputToRange As Range = Nothing, _
                    Optional ByVal boCorrectOverhead As Boolean = False, _
                    Optional ByVal boForceMillis As Boolean = False)
'Public method of report function. Calculates and outputs report with default settings to debug window (ctrl + G to show).
'Can be called from Immediate window or in break mode/while running code
    ReportArg boExtendedReport, boTransposeReport, OutputToRange, boCorrectOverhead, boForceMillis
End Sub

Public Sub Sleep(seconds As Double, Optional boDoEventsWhileSleeping As Boolean = True)
'Same as Application.Wait function, but more accurate and easier to use.
'a.Sleep 2      <- VS ->     Application.Wait Now + TimeValue("0:00:02")
'set boDoEventsWhileSleeping to false for even more accuracy
Dim startstamp As Currency, restamp As Currency
QueryPerformanceCounter startstamp
Do While ticsToSeconds(stampsToTics(startstamp, restamp)) <= seconds
    If boDoEventsWhileSleeping Then DoEvents
    QueryPerformanceCounter restamp
Loop
End Sub
Public Sub Wait(seconds As Double, Optional boDoEventsWhileWaiting As Boolean = True)
'For when you're used to methodname Wait
    Sleep seconds, boDoEventsWhileWaiting
End Sub

' ============================================= '
' Private Functions - Specific bench helpers
' ============================================= '
' @Reset                - reset/re-initialise all variables
' @RedimStampArrays     - enlarge stamp arrays
' @ReportArg            - calculate and write report

Private Sub Reset()
'Sub is private as public method .Start does the same
    QueryPerformanceFrequency freq      'Make sure frequency is set right (in case an instance of this class is declared public static)
    freq = freq * fromCurr
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
                        Optional ByVal boTransposeReport As Boolean = False, _
                        Optional ByVal OutputToRange As Range = Nothing, _
                        Optional ByVal boForceMillis As Boolean = False, _
                        Optional ByVal boForceNanos As Boolean = False)

'dont generate report if it was generated less then 1 seconds ago (f.e. when ReportCustom
'was called at end of code, then ignore print call from Class_Terminate)
Dim stamp_ReportStart As Currency
QueryPerformanceCounter stamp_ReportStart
If stamp_ReportEnd > 0 Then If ticsToSeconds(stampsToTics(stamp_ReportEnd, stamp_ReportStart)) <= 1 Then Exit Sub

'Nothing to report when only .Start (1 stamp) was called
If stampCount < 2 Then GoTo theEnd

'Start report with dimensions
Dim i As Long                           'index number at various places
Dim v As Variant                        'used in various loops over dictionary/collections

Dim arTicDiffs() As Currency            'array to hold the differences between two stamps, same sized array as arrStamp and arrStampID
Dim dID_colTicDiffs As New Dictionary   'key = IDnr, value = collection of time recordings (tics) per IDnr
Dim key_idnr As Variant                 'used to loop through dID_colTicDiffs
Dim colTicDiffs As New Collection       'collection of TicDiffs coming out of dID_colTicDiffs
Dim col_item As Variant                 'used to loop through tic(difss) in colTicDiffs

Dim cntTic As Double                    'tic-values for report
Dim sumTics As Double
Dim minTic As Double
Dim maxTic As Double
Dim avrTics As Double
Dim cntAllTics As Double
Dim sumAllTics As Double
Dim ticsOverhead As Double

Dim dAll As New Dictionary              'temp to hold the values of the output report. key = IDnr concatenated with the ValueType
Dim dHeaders As New Dictionary          'dict to filter out unique ValueTypes out of dAll
Dim header As Variant                   'loop through keys in dHeaders
Dim arrReport() As Variant              'holds report values as (2D) table
Dim col As Long, row As Long            'index numbers used for looping in arrReport
Dim strID As String                     'IDnr of stamp as string


'calculate tic-differences (TicDiffs) per Track-call and store in evenly sized array
ReDim arTicDiffs(LBound(arrStamp) To stampCount)
For i = LBound(arrStamp) To stampCount 'LBound always is start-stamp
    arTicDiffs(i) = stampsToTics_fromArrays(i - 1, i)
Next i
    
'seperate TicDiffs into ID-specific collection (most time consuming step in this sub)
Set dID_colTicDiffs = ticsToCollectionsInDictionaryPerID(arTicDiffs, LBound(arTicDiffs))

'filter out any unwanted output here
dID_colTicDiffs.Remove TrackID.id_start & "" 'start tic value is always 0, so always filter out
If Not boExtendedReport Then
    If dID_colTicDiffs.Exists(TrackID.id_continue & "") Then dID_colTicDiffs.Remove TrackID.id_continue & ""
    If dID_colTicDiffs.Exists(TrackID.id_pause & "") Then dID_colTicDiffs.Remove TrackID.id_pause & ""
    If dID_colTicDiffs.Exists(TrackID.id_resizingarray & "") Then dID_colTicDiffs.Remove TrackID.id_resizingarray & ""
End If
'example of filtering your own defined tracks, both steps required:
'dID_colTicDiffs.Remove dicStampName_ID("Initialisation")
'dicStampName_ID.Remove "Initialisation"

'check if TrackByName method is used and store names
'If TrackByName is not used, name-column won't be printed, so print Totals-name in IDnr column
If dicStampName_ID.Count > 0 Then
    For Each v In dicStampName_ID.Keys()
        dAll.item(dicStampName_ID(v) & "_Name") = v
    Next v
    dAll.item("TOTAL" & "_Name") = "TOTAL"
Else
    dAll.item("TOTAL_IDnr") = "TOTAL"
End If

'UDT's in VBA can't be stored in a collection/dictionary inside a class module.
'Instead output values are stored in a dictionary with the key being the "id" concatenated with the "_Valuetype".
'After this the "_Valuetype" becomes the header-name of the output table.
'This way the output only has to be added/adjusted at one place, instead of at calculation ánd report-output functions.
'Other options, like ADO recordset or an array of UDT's, would require to adjust the reportcode in 3 places:
'at decleration of the UDT, at calculation (sum, count, etc) and at report formatting/creating the table.
'In current set up, these three things are done at the same place.

' -------------------------------------------------------------------------
' ----------------------- Start calculating report ------------------------
' -------------------------------------------------------------------------

cntAllTics = 0: sumAllTics = 0
For Each key_idnr In dID_colTicDiffs.Keys 'loop all identical IDnrs
    
    dAll.item(key_idnr & "_IDnr") = key_idnr
    
    'overwrite names of the TrackID's this class uses.
    Select Case key_idnr
        Case TrackID.id_start:          dAll.item(TrackID.id_start & "_Name") = "(Start)"           'Initialisation
        Case TrackID.id_pause:          dAll.item(TrackID.id_pause & "_Name") = "(Before Pause)"    'Pause start
        Case TrackID.id_continue:       dAll.item(TrackID.id_continue & "_Name") = "(Continue)"     'After Pause/Paused
        Case TrackID.id_resizingarray:  dAll.item(TrackID.id_resizingarray & "_Name") = "(Resizing)" 'Resizing arStampID and arStamp
    End Select
    
    Set colTicDiffs = dID_colTicDiffs(key_idnr)
    cntTic = 0: minTic = 1E+15: maxTic = 0: sumTics = 0: avrTics = 0
    'break here to see the cpu tic-differences in between TrackBy calls
    For Each col_item In colTicDiffs 'col_item = collection of (ammount of) tics
        cntTic = cntTic + 1
        minTic = Min(minTic, col_item)
        maxTic = Max(maxTic, col_item)
        sumTics = sumTics + col_item
    Next col_item
    sumAllTics = sumAllTics + sumTics
    cntAllTics = cntAllTics + cntTic
    
    v = key_idnr 'IDnr
    dAll.Add v & "_Count", FormatNumber(cntTic, 0)
    dAll.Add v & "_Sum of tics", FormatNumber(sumTics, 0)
    dAll.Add v & "_Percentage", "" 'value cant be calculated yet as total sum is yet unknown, but already place in output table
    dAll.Add v & "_Time sum", secondsProperString(ticsToSeconds(sumTics), boForceMillis, boForceNanos)
    
    If Not boExtendedReport Then GoTo nextV_SkipExtendedOutput
' ---------------------- Output for extended report: ----------------------
    
    dAll.Add v & "_Minimum", minTic
    dAll.Add v & "_Maximum", maxTic
    dAll.Add v & "_Average", FormatNumber(sumTics / cntTic)
    dAll.Add v & "_Median", FormatNumber(MedianOfFirst_x_Elements(colTicDiffs, 1000)) 'Only from first 1000 tic measurements
    
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

'restores statically stored stamp arrays, does nothing if not called before
v = OverheadPerTrackCall(v, "restore")

'calculate percentage per ID, now that sumAllTics is known
For Each key_idnr In dID_colTicDiffs.Keys 'all identical IDnrs
    v = key_idnr
    dAll.item(v & "_Percentage") = FormatPercent(dAll.item(v & "_Sum of tics") / sumAllTics)
Next key_idnr

'calculate totals
dAll.item("TOTAL" & "_Count") = FormatNumber(cntAllTics, 0)
dAll.item("TOTAL" & "_Sum of tics") = FormatNumber(sumAllTics, 0)
If sumAllTics > 0 Then dAll.item("TOTAL" & "_Percentage") = FormatPercent(dAll.item("TOTAL" & "_Sum of tics") / sumAllTics)
dAll.item("TOTAL" & "_Time sum") = secondsProperString(ticsToSeconds(sumAllTics), boForceMillis, boForceNanos)

If boExtendedReport Then
    If cntAllTics > 0 Then dAll.item("TOTAL" & "_Average") = Round(sumAllTics / cntAllTics, 0)
End If

'dAll now holds all the values for the report. key = IDnr_ValueType, value = value

' -------------------------------------------------------------------------
' ---------------------- End of calculating report ------------------------
' -------------------------------------------------------------------------

'add unique headers for output table
dHeaders.Add "IDnr", 1 'makes sure IDnr is first/most left column
For Each v In dAll.Keys
    dHeaders.item(RIGHT_AfterLastCharsOf(v, "_")) = 0
Next v

col = 0: row = 0 'column, row
ReDim arrReport(1 To dHeaders.Count, 1 To dID_colTicDiffs.Count + 1 + 1) 'arrReport(headers, datarows + headerrow + totalsrow)
'loop all possible ID numbers and store values of dAll in arrReport
'Byte range is 0-255, minimum ID = 0, nr -1 for headers, nr 256 for TOTAL, sorted order (id_pause etc as last)
For i = -1 To 256
    Select Case i
        Case -1:                'headers
        Case 256:               strID = "TOTAL"
        Case Else
            strID = i & ""
            If Not dID_colTicDiffs.Exists(strID) Then GoTo nextI
    End Select
    row = row + 1
    col = 0
    
    For Each header In dHeaders.Keys()
        col = col + 1
        If row = 1 Then
            arrReport(col, row) = header
        Else
            If dAll.Exists(strID & "_" & header) Then
                arrReport(col, row) = dAll.item(strID & "_" & header)
            End If
        End If
    Next header
nextI:
Next i

If boTransposeReport Then arrReport = Transpose2DArray(arrReport)

Array2DToImmediate (arrReport)

theEnd:
QueryPerformanceCounter stamp_ReportEnd
Debug.Print "Total time recorded:             " & secondsProperString(ticsToSeconds(stampsToTics_fromArrays(LBound(arrStamp), stampCount)))
If boExtendedReport Then Debug.Print "Time to calculate report stamps: " & secondsProperString(ticsToSeconds(stampsToTics(stamp_ReportStart, stamp_ReportEnd)))
If boExtendedReport Then Debug.Print "Max precision:                   " & secondsProperString(Precision, , True)
Debug.Print ""

End Sub

' ============================================= '
' Private Functions - Specific Report Helpers
' ============================================= '
' @OverheadPerTrackCall                 - overhead of QPC including TrackBy-methods
' @OverheadPerQPCcall                   - overhead of only the QPC function
' @Precision                            - returns maximum precision of this class in seconds
' @ticsToCollectionsInDictionaryPerID   - group stamps from global stamparray into seperate (per tracked ID) collections
' @stampsToTics_fromArrays              - retrieve tics from arrays and return difference
' @stampsToTics                         - returns difference between to stamps
' @ticsToSeconds                        - convert qpc-tics to seconds
' @secondsProperString                  - convert seconds to appropriate readable text

Private Function OverheadPerTrackCall(NameOrID As Variant, Action As String) As Double
'calculates the overhead in amount of tics to call methods TrackByTheID and TrackByName.
'As these two methods adjust (values in) global variables, these global variables
'are also used to calculate the overhead. They are first copied and stored as Static, which
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
Select Case Action
    Case "ByIDAvr", "ByIDMin"
        id = CByte(NameOrID)
        For i = frst_loop To last_loop
            TrackByTheID id
        Next i
        
    Case "ByNameAvr", "ByNameMin"
        name = NameOrID
        For i = frst_loop To last_loop
            TrackByName name
        Next i
        
    Case Else '"restore"
    
End Select

Select Case Action
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
'Does not include overhead of TrackByTheID (calling the function
'itself and the if-block within) or TrackByName (the time it takes
'to look up the IDnr from dictionary).

Dim arr() As Currency: ReDim arr(1 To overheadTestCount)
Dim i As Long

For i = LBound(arr) To UBound(arr)
    QueryPerformanceCounter arr(i)
Next i
OverheadPerQPCcall = (arr(UBound(arr)) - arr(LBound(arr))) * fromCurr / overheadTestCount
End Function
Private Function Precision() As Double
'returns maximum available precision of this benchmark class on the machine it runs in (full) seconds.
'As described in microsoft docs https://docs.microsoft.com/en-us/windows/win32/sysinfo/acquiring-high-resolution-time-stamps#low-level-hardware-clock-characteristics

'Tick Interval = 1/(Performance Frequency) = Resolution
Dim resolution As Double
resolution = 1 / freq

'ElapsedTime = Ticks * Tick Interval = AccessTime
Dim accesTime As Double
accesTime = OverheadPerQPCcall * resolution

Precision = Max(resolution, accesTime)

End Function

Private Function ticsToCollectionsInDictionaryPerID(ByRef arTdifs() As Currency, ByVal lb As Long) As Dictionary
'Groups the global stamp-array into seperate collections per ID
'Returns a dictionary where key = TrackID, value = collection of tics
'example result in jsonformat: {"255":[0],"1":[156],"2":[675,766,523,764,651]}

  Set ticsToCollectionsInDictionaryPerID = New Dictionary
  
  Dim offset As Long
  For offset = 0 To stampCount - 1
    On Error GoTo new_item
    ticsToCollectionsInDictionaryPerID.item(LTrim$(str$(arrStampID(lb + offset)))).Add arTdifs(lb + offset)
    On Error GoTo 0
  Next
  
  Exit Function
  
new_item:
  Set ticsToCollectionsInDictionaryPerID.item(LTrim$(str$(arrStampID(lb + offset)))) = New Collection
  Resume
End Function

Private Function stampsToTics_fromArrays(ByVal stampNrBefore As Long, ByVal stampNrAfter As Long) As Currency
'Gets stamps from arrays and return difference in tics between them
If stampNrBefore < LBound(arrStamp) Then
    stampsToTics_fromArrays = 0
Else
    stampsToTics_fromArrays = stampsToTics(arrStamp(stampNrBefore), arrStamp(stampNrAfter))
End If
End Function

Private Function stampsToTics(ByVal stampBefore As Currency, ByVal stampAfter As Currency) As Currency
'Calculates the difference in between two QPC stamps and upscales them from Currency to whole numbers

'example returns of QPC:
'- as Currency -> 304462680,3775    --> needs upscaling by 10 000
'- as LongLong -> 3044626803775
'--->
'- as Currency -> (QPC2 - QPC1) * 10000
'- as LongLong -> (QPC2 - QPC1)

'example returns of QPF (is system specific, but commonly 10Mhz on Windows 10)
'with a usual QPF on windows 10 (10MHz):
'- as Currency ->     1000  =      1 000
'- as LongLong -> 10000000  = 10 000 000

'---> if freq is 10MHz then:
'---> 10 million tics per second
'   1 tic = (1 / 10 000 000) seconds
'   1 tic = 0.0000001 seconds
'   1 tic = 0.0001 milliseconds
'   1 tic = 0.1 microseconds
'   1 tic = 100 nanoseconds

'tics t to seconds s = t/(s * 1)
'tics to milliseconds = t/(s * 1e3)
'tics to microseconds = t/(s * 1e6)
'tics to nanoseconds = t/(s * 1e9)

stampsToTics = (stampAfter - stampBefore) * fromCurr
End Function
Private Function ticsToSeconds(ByVal tics As Currency) As Double
'returns time in full seconds
    If Int(tics) <> tics Or Int(freq) <> freq Then err.Raise 9999999, , "QPC or QPF returns with datatype As Currency downscales the returns with 10 000. Upscale both returns before calling this funciton."
    'Int(freq) is actually not a proper check to see if it has been upscaled, as it is often also a round number when downscaled (10mhz)
    ticsToSeconds = tics / freq
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
    
ElseIf t > 120 And Not boForceNanoSeconds And Not boForceMilliSeconds Then  '>2 minutes
    res = Round(t, 1) & " s"
    
ElseIf t > 10 And Not boForceNanoSeconds And Not boForceMilliSeconds Then   '>10 seconds
    res = Round(t, 1) & " s"
    
ElseIf t > 0.9 And Not boForceNanoSeconds And Not boForceMilliSeconds Then  '>0.9 second
    res = Round(t, 2) & " s"
    
ElseIf t > (10 / 1000#) And Not boForceNanoSeconds Or boForceMilliSeconds Then  'millisecond (1 ms = 10-E3 s)
    res = Round(t * 1000#, IIf(boForceMilliSeconds, 2, 0)) & " ms"
    
ElseIf t > (1 / 1000#) And Not boForceNanoSeconds Then                          'millisecond (1 ms = 10-E3 s)
    res = Round(t * 1000#, 2) & " ms"
    
ElseIf t > (10 / 1000000#) And Not boForceNanoSeconds Then                      'microsecond (1 us = 10-E6 s)
    res = Round(t * 1000000#) & " us"
    
ElseIf t > (10 / 1000000000#) Or boForceNanoSeconds Then                        'nanosecond (1 ns = 10-E9 s)
    res = Round(t * 1000000000#) & " ns"

'Any value below this is probably below the maximum precision of the QPC function (and likely cause of overhead correction).
'max precision = 1 / frequency * QPC overhead.

ElseIf t > (10 / 1000000000000#) Then res = Round(t * 1000000000000#) & " ps"   'picosecond (1 ps = 10-E12 s)
ElseIf t > (10 / 1E+15) Then res = Round(t * 1E+15) & " fs"                     'femtosecond (1 fs = 10-E15 s)
ElseIf t > (10 / 1E+18) Then res = Round(t * 1E+18) & " as"                     'attosecond (1 as = 10-E18 s)
ElseIf t > (10 / 1E+21) Then res = Round(t * 1E+21) & " zs"                     'zeptosecond (1 as = 10-E21 s) -> shortest time ever measured was 247 zeptoseconds :)
ElseIf t > (10 / 1E+24) Then res = Round(t * 1E+24) & " ys"                     'yoctosecond (1 as = 10-E24 s)
'"For Decimal expressions, any fractional value less than 1E-28 might be lost." (.net docs)

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
' Private Functions - General Report Helpers
' ============================================= '
' @Min                          - minimum of two double-values
' @Max                          - maximum of two double-values
' @MedianOfFirst_x_Elements     - median of a part of a collection
' @QuickSortArray               - quick sort an array
' @RIGHT_AfterLastCharsOf       - last part of string
' @Array2DToImmediate           - print array to console
' @Transpose2DArray             - flip 2D-array 90 degrees

Private Function Min(ByVal x As Double, ByVal y As Double) As Double
    If x < y Then Min = x Else Min = y
End Function
Private Function Max(ByVal x As Double, ByVal y As Double) As Double
    If x > y Then Max = x Else Max = y
End Function

Private Function MedianOfFirst_x_Elements(col As Collection, x As Long) As Double 'MedianFromCollection
'puts specified amount (x) of values of collection into an array, quicksorts
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

Private Function RIGHT_AfterLastCharsOf(ByVal strLeft As String, ByVal chars As String) As String
'returns the part of the string that is most right to given char(s)
Dim s() As String
s = Split(strLeft, chars, -1, vbBinaryCompare)
RIGHT_AfterLastCharsOf = s(UBound(s))
End Function


Private Sub Array2DToImmediate(ByVal arr As Variant)
'Prints a 2D-array of values to a table (with same sized column widhts) in the immmediate window

'Each character in the Immediate window of VB Editor (CTRL+G to show) has the same pixel width,
'thus giving the option to output a proper looking 2D-array (a table with variable string lenghts).

'settings
Dim spaces_between_collumns As Long: spaces_between_collumns = 2
Dim NrOfColsToOutlineLeft As Long: NrOfColsToOutlineLeft = 2    'all cols are outlined right, except for first x (2 here, so IDnr and Name)
Dim maxLength As Long: maxLength = 198 * 1021&                  'capacity of Immediate window is about 200 lines of 1021 characters per line.
Dim i As Long, j As Long
Dim arrMaxLenPerCol() As Long
Dim str As String

'determine max stringlength per column
ReDim arrMaxLenPerCol(UBound(arr, 1))
For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(arr, 2) To UBound(arr, 2)
        arrMaxLenPerCol(i) = IIf(Len(arr(i, j)) > arrMaxLenPerCol(i), Len(arr(i, j)), arrMaxLenPerCol(i))
    Next j
Next i

'build table
For j = LBound(arr, 2) To UBound(arr, 2)
    For i = LBound(arr, 1) To UBound(arr, 1)
        'outline left --> value & spaces & column_spaces
        If i < NrOfColsToOutlineLeft Then
            On Error Resume Next
            str = str & arr(i, j) & space$((arrMaxLenPerCol(i) - Len(arr(i, j)) + spaces_between_collumns) * 1)
        
        'last column to outline left --> value & spaces
        ElseIf i = NrOfColsToOutlineLeft Then
            On Error Resume Next
            str = str & arr(i, j) & space$((arrMaxLenPerCol(i) - Len(arr(i, j))) * 1)
                    
        'outline right --> spaces & column_spaces & value
        Else 'i > NrOfColsToOutlineLeft Then
            On Error Resume Next
            str = str & space$((arrMaxLenPerCol(i) - Len(arr(i, j)) + spaces_between_collumns) * 1) & arr(i, j)
        End If
    Next i
    str = str & vbNewLine
    If Len(str) > maxLength Then GoTo theEnd
Next j

theEnd:
'capacity of Immediate window is about 200 lines of 1021 characters per line.
If Len(str) > maxLength Then str = Left(str, maxLength) & vbNewLine & " - Table to large for Immediate window"
Debug.Print str
End Sub

Private Function Transpose2DArray(arr() As Variant) As Variant()
Dim arTemp() As Variant
Dim c As Long
Dim r As Long
ReDim arTemp(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
For r = LBound(arTemp, 1) To UBound(arTemp, 1)
    For c = LBound(arTemp, 2) To UBound(arTemp, 2)
        arTemp(r, c) = arr(c, r)
    Next c
Next r
Transpose2DArray = arTemp
End Function



