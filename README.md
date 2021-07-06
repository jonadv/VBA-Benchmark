**# VBA-Benchmark**


[![release](https://img.shields.io/github/release/jonadv/VBA-Benchmark.svg?style=flat&logo=github)](https://github.com/jonadv/VBA-Benchmark/releases/latest) [![last-commit](https://img.shields.io/github/last-commit/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark/commits/master) [![downloads](https://img.shields.io/github/downloads/jonadv/VBA-Benchmark/total.svg?style=flat)](https://somsubhra.com/github-release-stats/?username=jonadv&repository=VBA-Benchmark) [![code-size](https://img.shields.io/github/languages/code-size/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark) [![language](https://img.shields.io/github/languages/top/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark/search?l=vba) [![license](https://img.shields.io/github/license/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark/blob/master/LICENSE) [![gitter](https://img.shields.io/gitter/room/jonadv/VBA-Benchmark.svg?style=flat&logo=gitter)](https://gitter.im/jonadv)

A one-class module to conveniently time your code. Reports automatically. 
Includes a convinience method to use instead of Application.Wait. 

**Advantages**
- Accuracy to nanoseconds
- No delay of console logging/using Debug.Print
- Convenient to use
- Identify parts of code
- Re-use same name for pieces of code all over the place
- Time code spread over several modules
- Up to 256 different tracks

**How to use**

- Copy paste all code into a class module named cBenchmark
- Create an instance of it at the start of your own code (fe `Dim bm As New cBenchmark`)
- Or declare class instance as global to time multi-module code (write `Set bm = New cBenchmark` at start of code)
- Write `bm.TrackByName "Description of codepart"` in between all of your code 
- Open Immediate window and run you're code
- When you're code finishes, the report is printed automatically 
- Or, when instance is declared global, write `Set bm = Nothing` before last `End Sub`
- The report will show each identically used name, the time it took and the time percentage of each track


**Example:**

```
Sub testCBenchmark()
    Dim bm As New cBenchmark
    Dim i As Long
bm.TrackByName "Initialisations"

    bm.Sleep 0.05    'wait 50 milliseconds/simulating code running
bm.TrackByName "Slept"

    For i = 1 To 1000000
        i = i * 1
    Next i
bm.TrackByName "Finished loop"

    bm.Sleep 0.05    'wait 50 milliseconds/simulating code running
bm.TrackByName "Slept"

    Application.Wait Now + TimeValue("0:00:01")
bm.TrackByName "Waited"
End Sub
```

**Prints:**

```
IDnr  Name             Count  Sum of tics  Percentage  Time sum
0     Initialisations      1          191       0,01%     19 us
1     Slept                2    1.005.608      26,95%    101 ms
2     Finished loop        1       79.548       2,13%   7,95 ms
3     Waited               1    2.646.483      70,92%    265 ms
      TOTAL                5    3.731.830     100,00%    373 ms

Total time recorded:             373 ms
```

**How it works**

Everytime `TrackByName` is called a 'CPU-timestamp' is stored. After you're code finishes, stamps are grouped and written to a report. 


**Function overview**
 | Scope | Method Name | Description | Return value |
 | ----- | ----------- | ----------- | ------------ |
 | Class specific Functions | Class_Initialize | initialise varialbes and set first stamp | 	
 | Class specific Functions | Class_Terminate | calculates and writes report to debug | 	
 | Public Functions | TrackByName | Same as @TrackByTheID but more convenient (and thus with a bit more overhead) | 	
 | Public Functions | TrackByTheID | Store QPC (cycle counts) in an array | 	
 | Public Functions | Start | (Start) or (Reset and Restart) benchmark | 	
 | Public Functions | Pause | Convenience method to exclude pieces of code, use in combination with .Continue | 	
 | Public Functions | Continue | Use after calling .Pause to continue tracking | 	
 | Public Functions | Report | Generate report | 	
 | Public Functions | Sleep | timeout code, alternative for Application.Wait | 	
 | Public Functions | Wait | same as method Sleep | 	
 | Private Functions - Specific bench helpers | Reset | reset/re-initialise all variables | 	
 | Private Functions - Specific bench helpers | RedimStampArrays | enlarge stamp arrays | 	
 | Private Functions - Specific bench helpers | ReportArg | calculate and write report | 	
 | Private Functions - Specific Report Helpers | OverheadPerTrackCall | overhead of QPC including TrackBy-methods | 	
 | Private Functions - Specific Report Helpers | OverheadPerQPCcall | overhead of only the QPC function | 	
 | Private Functions - Specific Report Helpers | Precision | returns maximum precision of this class in seconds | 	
 | Private Functions - Specific Report Helpers | ticsToCollectionsInDictionaryPerID | group stamps from global stamparray into seperate (per tracked ID) collections | 	
 | Private Functions - Specific Report Helpers | stampsToTics_fromArrays | retrieve tics from arrays and return difference | 	
 | Private Functions - Specific Report Helpers | stampsToTics | returns difference between to stamps | 	
 | Private Functions - Specific Report Helpers | ticsToSeconds | convert qpc-tics to seconds | 	
 | Private Functions - Specific Report Helpers | secondsProperString | convert seconds to appropriate readable text | 	
 | Private Functions - General Report Helpers | Min | minimum of two double-values | 	
 | Private Functions - General Report Helpers | Max | maximum of two double-values | 	
 | Private Functions - General Report Helpers | MedianOfFirst_x_Elements | median of a part of a collection | 	
 | Private Functions - General Report Helpers | QuickSortArray | quick sort an array | 	
 | Private Functions - General Report Helpers | RIGHT_AfterLastCharsOf | last part of string | 	
 | Private Functions - General Report Helpers | Array2DToImmediate | print array to console | 	
 | Private Functions - General Report Helpers | Transpose2DArray | flip 2D-array 90 degrees | 	
![image](https://user-images.githubusercontent.com/10421216/124674544-3dae0f00-debb-11eb-8261-cbec18b6963c.png)
































_**dev notes**

- This should be the fastest and most accurate possible way to time VBA code, as in, with as little overhead as possible.
- When the code is running, the only thing this class does is storing the timestamps. Any processing is delayed untill after code finished running.
- The only better method then using the QueryPerformanceCounter would be to read the TSC directly with RDTSC, which requires a custom made .dll - and a bit more complexity.
- Using a UDT instead of Currency or LongLong as a datatype of the QPC function might be faster then returning a LongLong or Currency, but storing/handling that returnvalue in VBA is far from faster then just a LongLong or Currency value. This mmight be because VBA needs to store two seperate values (lowpart and highpart), instead of just a big one. As LongLong is only available on 64-bit systems and there seemed to be no differnce in speeds whatsoever (between these two datatypes), Currency is the best option.



_
