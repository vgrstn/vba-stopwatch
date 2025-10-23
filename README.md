# Stopwatch Class (VBA)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Office)-blue)
![Compatibility](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

High-resolution stopwatch for VBA using the Windows Performance Counter API.  
Provides **nanosecond-level precision** on supported CPUs, works on both 32- and 64-bit VBA hosts,  
and includes a **global predeclared instance** for immediate use — just call `Stopwatch.Start`.

---

## 📦 Features

- High-resolution timing via `QueryPerformanceCounter` and `QueryPerformanceFrequency`
- Works on both **x86** and **x64** architectures
- **Predeclared default instance** — usable anywhere without `New` (`Stopwatch.Start`)
- Supports **multiple independent instances** for concurrent timing
- Extremely lightweight — uses native 8-byte counters (`Currency` as `LARGE_INTEGER`)
- MIT-licensed and [Rubberduck-compatible](https://rubberduckvba.com/)
- Safe, dependency-free, and fully deterministic within Windows timer precision

---

## ⚙️ Public Interface

| Member        | Type       | Description |
|----------------|------------|-------------|
| `Start()`      | `Sub`      | Starts or resumes the stopwatch. |
| `Pause()`      | `Function` | Pauses and returns seconds since last `Start`. |
| `Halt()`       | `Function` | Stops and returns total seconds since first start. |
| `Reset()`      | `Sub`      | Resets all counters. |
| `Interval()`   | `Function` | Returns seconds since last `Start` without stopping. |
| `Elapsed()`    | `Function` | Returns total elapsed seconds (default member). |
| `Running`      | `Property` | Returns `True` if stopwatch is currently running. |

---

## ⚡ Performance Notes

- Typical overhead: **< 0.5 µs per call** on modern CPUs.  
- Precision is limited by Windows’ scheduler (~0.1 ms on modern systems).  
- For reliable benchmarking, repeat measurements and compute averages.  
- `Currency` is used as a safe 8-byte wrapper for `LARGE_INTEGER` in both x86 and x64 environments.  
- The implementation overhead is **negligible compared to process jitter**.

---

### API References

| API Function | Library | Description |
|---------------|----------|-------------|
| `QueryPerformanceFrequency` | `kernel32.dll` | Retrieves the frequency of the high-resolution performance counter. |
| `QueryPerformanceCounter`   | `kernel32.dll` | Retrieves the current value of the high-resolution performance counter. |

---

## 🚀 Quick Start (Predeclared Instance)

No `New` keyword required — the class is **predeclared** (`@PredeclaredId`).  
That means a **global instance** named `Stopwatch` is available as soon as the class is imported.

```vb
' Basic usage with the global predeclared instance
Stopwatch.Start
SomeProcedure
Debug.Print "Elapsed:", Stopwatch.Halt, "seconds"
```

## 🧪 Example – Benchmark Loop

A minimal example showing how to benchmark a code block or algorithm using the predeclared stopwatch instance.

```vb
Dim i As Long, total As Double

Stopwatch.Start
For i = 1 To 1000000
    total = total + Sqr(i)
Next
Debug.Print "Elapsed time:", Stopwatch.Halt, "seconds"
```
