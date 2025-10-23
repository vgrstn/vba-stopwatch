VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Stopwatch Class."
'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

'@IgnoreModule ProcedureNotUsed, VariableNotAssigned, UnassignedVariableUsage
'@PredeclaredId
'@Folder("Class")
'@ModuleDescription "Stopwatch Class."

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Stopwatch Class
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.10.10
'
' Methods
' Start             Starts or resumes the stopwatch
' Reset             Resets the stopwatch.
' Pause             Pauses and returns the time measured since the last start
' Halt              Halts and returns the time measured since the first start
' Interval          Returns the time measured since the last start
' Elapsed           Returns the time measured since the first start (default)
'
' Properties
' Running           Returns True if the stopwatch is measuring or False if not
'
' The Stopwatch Class is predeclared which creates a global default instance
' with the same name as the Class making the object immediately available. If
' needed multiple stopwatches can be instantiated and used simultaneously and
' independently by instantiating multiple Stopwatch objects.
'
' The implementation is based on two API functions which require the use of a
' UNIX LARGE_INTEGER (8-byte) variable. An 8-byte VB Currency type is available
' for both x32 and x64 architectures. This works perfectly because the measured
' time in seconds is obtained as a quotient of two VB Currency values.
'
' It is not necessary to check for the successful execution of the API's.
' see https://learn.microsoft.com/en-us/windows/win32/api/profileapi
'
' The cpu time taken for the measurement itself is insignificant compared to the
' variation caused by the preemptive multitasking by the Windows OS. To obtain a
' reliable result, the actual measurement should be repeated thousands of times
' whilst limiting the interference by other processes.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private API declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32.dll" ( _
    ByRef lpPFreq As Currency _
) As Boolean

' https://docs.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32.dll" ( _
    ByRef lpPCount As Currency _
) As Boolean


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type TPRIVATE
    TicksPerSecond As Currency
    StartTicks As Currency
    AccumulatedTicks As Currency
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    QueryPerformanceFrequency this.TicksPerSecond
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Starts or resumes the stopwatch."
Public Sub Start()
Attribute Start.VB_Description = "Starts or resumes the stopwatch."
    
    If this.StartTicks = 0 Then
        QueryPerformanceCounter this.StartTicks
    End If
    
End Sub


'@Description "Resets the stopwatch."
Public Sub Reset()
Attribute Reset.VB_Description = "Resets the stopwatch."
    
    this.StartTicks = 0
    this.AccumulatedTicks = 0
    
End Sub


'@Description "Pauses the stopwatch and returns the time measured since the last start."
Public Function Pause() As Double
Attribute Pause.VB_Description = "Pauses the stopwatch and returns the time measured since the last start."
        
    If this.StartTicks <> 0 Then
        Dim Ticks As Currency: QueryPerformanceCounter Ticks
        this.AccumulatedTicks = this.AccumulatedTicks + Ticks - this.StartTicks
        Pause = (Ticks - this.StartTicks) / this.TicksPerSecond
        this.StartTicks = 0
    End If
    
End Function
 

'@Description "Halts the stopwatch and returns the total time measured since the first start."
Public Function Halt() As Double
Attribute Halt.VB_Description = "Halts the stopwatch and returns the total time measured since the first start."
    
    If this.StartTicks <> 0 Then
        Dim Ticks As Currency: QueryPerformanceCounter Ticks
        this.AccumulatedTicks = this.AccumulatedTicks + Ticks - this.StartTicks
        this.StartTicks = 0
    End If
    Halt = this.AccumulatedTicks / this.TicksPerSecond
    this.AccumulatedTicks = 0

End Function


'@Description "Returns the time measured since the last start of the stopwatch."
Public Function Interval() As Double
Attribute Interval.VB_Description = "Returns the time measured since the last start of the stopwatch."

    If this.StartTicks <> 0 Then
        Dim Ticks As Currency: QueryPerformanceCounter Ticks
        Interval = (Ticks - this.StartTicks) / this.TicksPerSecond
    End If
    
End Function


'@DefaultMember
'@Description "Returns the time measured since the first start of the stopwatch."
Public Function Elapsed() As Double
Attribute Elapsed.VB_Description = "Returns the time measured since the first start of the stopwatch."
Attribute Elapsed.VB_UserMemId = 0
    
    If this.StartTicks <> 0 Then
        Dim Ticks As Currency: QueryPerformanceCounter Ticks
        Elapsed = (this.AccumulatedTicks + (Ticks - this.StartTicks)) / this.TicksPerSecond
    Else
        Elapsed = this.AccumulatedTicks / this.TicksPerSecond
    End If

End Function


'@Description "Returns True if the stopwatch is measuring or False if not."
Public Property Get Running() As Boolean
Attribute Running.VB_Description = "Returns True if the stopwatch is measuring or False if not."

    Running = (this.StartTicks <> 0)

End Property
