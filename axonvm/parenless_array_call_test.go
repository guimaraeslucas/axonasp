/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 */
package axonvm

import (
	"testing"
)

// TestParenlessMethodCallVariant1_ConcatenatedExpression verifies that parenless method calls
// on array elements with concatenated expression arguments (e.g. m_Items(0).Show prefix & "-")
// parse and execute without syntax or runtime errors.
func TestParenlessMethodCallVariant1_ConcatenatedExpression(t *testing.T) {
	source := `<%
Class ItemClass
    Public Sub Show(val)
        Response.Write "SHOW:" & val
    End Sub
End Class

Class Holder
    Private m_Items()
    Private Sub Class_Initialize(): ReDim m_Items(1): End Sub
    Public Function Add(item): Set m_Items(0) = item: End Function
    Public Function Run(prefix)
        m_Items(0).Show prefix & "-"
    End Function
End Class

Dim item, h
Set item = New ItemClass
Set h = New Holder
h.Add item
h.Run "HELLO"
%>`
	out := runASPSourceForTest(t, source)
	if out != "SHOW:HELLO-" {
		t.Fatalf("expected 'SHOW:HELLO-' got %q", out)
	}
}

// TestParenlessMethodCallVariant2_MultipleArguments verifies that parenless method calls
// on array elements with multiple simple arguments (e.g. m_Arr(0).Receive msg, fromUser)
// correctly bind both arguments without evaluating to Empty.
func TestParenlessMethodCallVariant2_MultipleArguments(t *testing.T) {
	source := `<%
Class Receiver
    Public Sub Receive(msg, fromUser)
        Response.Write "REC:" & msg & " FROM " & fromUser
    End Sub
End Class

Class B
    Private m_Arr()
    Private Sub Class_Initialize(): ReDim m_Arr(1): End Sub
    Public Function Store(obj): Set m_Arr(0) = obj: End Function
    Public Function Relay(msg, fromUser)
        m_Arr(0).Receive msg, fromUser
    End Function
End Class

Dim r, b
Set r = New Receiver
Set b = New B
b.Store r
b.Relay "HELLO", "ALICE"
%>`
	out := runASPSourceForTest(t, source)
	if out != "REC:HELLO FROM ALICE" {
		t.Fatalf("expected 'REC:HELLO FROM ALICE' got %q", out)
	}
}

// TestParenlessMethodCallVariant3_SingleArgument verifies that parenless method calls
// on array elements with a single argument inside a loop (e.g. m_Obs(i).Update news)
// pass the argument correctly without dropping it silently.
func TestParenlessMethodCallVariant3_SingleArgument(t *testing.T) {
	source := `<%
Class Observer
    Public Sub Update(news)
        Response.Write "UPD:" & news & ";"
    End Sub
End Class

Class Pub
    Private m_Obs()
    Private m_Count
    Private Sub Class_Initialize(): m_Count = 0: ReDim m_Obs(10): End Sub
    Public Function Add(o): Set m_Obs(m_Count) = o: m_Count = m_Count + 1: End Function
    Public Function Notify(news)
        Dim i
        For i = 0 To m_Count - 1
            m_Obs(i).Update news
        Next
    End Function
End Class

Dim obs1, obs2, pub
Set obs1 = New Observer
Set obs2 = New Observer
Set pub = New Pub
pub.Add obs1
pub.Add obs2
pub.Notify "BREAKING NEWS"
%>`
	out := runASPSourceForTest(t, source)
	if out != "UPD:BREAKING NEWS;UPD:BREAKING NEWS;" {
		t.Fatalf("expected 'UPD:BREAKING NEWS;UPD:BREAKING NEWS;' got %q", out)
	}
}

// TestParenlessMethodCallChainedAccess verifies multi-level chained property and method calls
// on array elements (e.g. m_Arr(0).SubObj.Update news).
func TestParenlessMethodCallChainedAccess(t *testing.T) {
	source := `<%
Class SubClass
    Public Sub Update(news)
        Response.Write "SUB:" & news
    End Sub
End Class

Class MainClass
    Public SubObj
    Private Sub Class_Initialize()
        Set SubObj = New SubClass
    End Sub
End Class

Class Container
    Private m_Arr()
    Private Sub Class_Initialize()
        ReDim m_Arr(1)
        Set m_Arr(0) = New MainClass
    End Sub
    Public Function Run(news)
        m_Arr(0).SubObj.Update news
    End Function
End Class

Dim c
Set c = New Container
c.Run "DATA"
%>`
	out := runASPSourceForTest(t, source)
	if out != "SUB:DATA" {
		t.Fatalf("expected 'SUB:DATA' got %q", out)
	}
}
