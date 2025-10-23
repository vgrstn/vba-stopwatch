# vba-stopwatch

High-resolution stopwatch for VBA using the Windows Performance Counter API.  
Provides nanosecond-level precision on supported CPUs, works on both 32- and 64-bit VBA hosts,  
and includes a **global predeclared instance** for immediate use — just call `Stopwatch.Start`.

---

## 📦 Features

- High-resolution timing via `QueryPerformanceCounter` and `QueryPerformanceFrequency`
- Works on both **x86** and **x64** architectures
- **Predeclared default instance** — usable without `New` (e.g., `Stopwatch.Start`)
- Supports **multiple independent instances** if you need concurrent timers
- Extremely lightweight — uses native 8-byte counters (`Currency` as `LARGE_INTEGER`)
- MIT-licensed and [Rubberduck](https://rubberduckvba.com/)-compatible

---

## ⚙️ Public Interface

| Member        | Type       | Description |
|----------------|------------|-------------|
| `Start()`      | `Sub`      | Starts or resumes the stopwatch. |
| `Pause()`      | `Function` | Pauses and returns seconds since last start. |
| `Halt()`       | `Function` | Stops and returns total seconds since first start. |
| `Reset()`      | `Sub`      | Resets all counters. |
| `Interval()`   | `Function` | Returns seconds since last start without stopping. |
| `Elapsed()`    | `Function` | Returns total elapsed seconds (default member). |
| `Running`      | `Property` | Returns `True` if stopwatch is currently running. |

---

## 🚀 Quick Start (Predeclared Instance)

No `New` keyword is needed — the class is **predeclared**:

```vb
' The predeclared instance is always available as "Stopwatch"
Stopwatch.Start
Call SomeProcedure
Debug.Print "Elapsed:", Stopwatch.Halt, "seconds"

---

## 🧪 Benchmark Example
Dim i As Long, total As Double

Stopwatch.Start
For i = 1 To 1000000
    total = total + Sqr(i)
Next
Debug.Print "Elapsed time:", Stopwatch.Halt, "seconds"

Example output:

Elapsed time: 0.092133 seconds

---

## 🔗 References

Microsoft Docs – QueryPerformanceCounter
Rubberduck VBA Add-in
