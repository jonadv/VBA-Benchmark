# VBA-Benchmark


[![release](https://img.shields.io/github/release/jonadv/VBA-Benchmark.svg?style=flat&logo=github)](https://github.com/jonadv/VBA-Benchmark/releases/latest) [![last-commit](https://img.shields.io/github/last-commit/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark/commits/master) [![downloads](https://img.shields.io/github/downloads/jonadv/VBA-Benchmark/total.svg?style=flat)](https://somsubhra.com/github-release-stats/?username=jonadv&repository=VBA-Benchmark) [![code-size](https://img.shields.io/github/languages/code-size/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark) [![language](https://img.shields.io/github/languages/top/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark/search?l=vba) [![license](https://img.shields.io/github/license/jonadv/VBA-Benchmark.svg?style=flat)](https://github.com/jonadv/VBA-Benchmark/blob/master/LICENSE) [![gitter](https://img.shields.io/gitter/room/jonadv/VBA-Benchmark.svg?style=flat&logo=gitter)](https://gitter.im/jonadv)

A one-class module to conveniently time your code. Reports automatically. 
Includes a convinience method to use instead of Application.Wait. 

How to use
- Copy paste all code into a class module named cBenchmark
- Create an instance of it at the start of your own code (fe `Dim bm As New cBenchmark`) 
- Write `bm.TrackByName "Description of codepart"` in between all of your code 
- Open Immediate window and run you're code
- When you're code finishes, the report is printed automatically (or write `Set bm = Nothing` to make sure)
- the report will show each identically used name and time it took

Example:

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

Will print: 
`
IDnr  Name             Count  Sum of tics  Percentage  Time sum
0     Initialisations      1          191       0,01%     19 us
1     Slept                2    1.005.608      26,95%    101 ms
2     Finished loop        1       79.548       2,13%   7,95 ms
3     Waited               1    2.646.483      70,92%    265 ms
      TOTAL                5    3.731.830     100,00%    373 ms

Total time recorded:             373 ms`

How it works
Everytime `TrackByName` is called a 'CPU-timestamp' is stored. After you're code finishes, stamps are grouped and the time running (per uniquely given name) is calculated. 


Notes
- This is the fastest and most accurate possible way to time VBA code, as in, with as little overhead as possible.
- When the code is running, the only thing this class does is storing the timestamps. Any processing is delayed untill after code finished running.
- The only better method then using the QueryPerformanceCounter would be to read the TSC directly with RDTSC, which requires a custom made .dll - and a bit more complexity.
- Using a UDT instead of Currency or LongLong as a datatype of the QPC function might be faster then returning a LongLong or Currency, but storing/handling that returnvalue in VBA is far from faster then just a LongLong or Currency value. This mmight be because VBA needs to store two seperate values (lowpart and highpart), instead of just a big one. As LongLong is only available on 64-bit systems and there seemed to be no differnce in speeds whatsoever (between these two datatypes), Currency is the best option.



