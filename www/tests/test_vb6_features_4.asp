<%
Option Explicit

Class EventSource
    Event OnStep(count)
    
    Sub Run()
        Dim i
        For i = 1 To 3
            RaiseEvent OnStep(i)
        Next
    End Sub
End Class

Class EventSink
    Private m_log
    Private WithEvents m_src
    
    Public Property Get Log()
        Log = m_log
    End Property
    
    Sub Class_Initialize()
        m_log = ""
        Set m_src = New EventSource
    End Sub
    
    Sub m_src_OnStep(count)
        m_log = m_log & count & ","
    End Sub
    
    Sub Start()
        m_src.Run()
    End Sub
End Class

Dim sink
Set sink = New EventSink
sink.Start()

Response.Write "Result: " & sink.Log
' Expected: Result: 1,2,3,
%>
