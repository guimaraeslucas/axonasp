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

func TestInterfaceNilMapPanic(t *testing.T) {
	code := `<%
Option Explicit

Class ICloneable
    Public Function Clone As ICloneable
    End Function
End Class

Class MyResume
    Implements ICloneable
    Public Name As String
    Public Age As Integer
    Public Skills

    Public Function ICloneable_Clone As ICloneable
        Dim copy As MyResume
        Set copy = New MyResume
        copy.Name = Me.Name
        copy.Age = Me.Age
        
        Dim ub As Integer, i As Integer, arr
        ub = UBound(Me.Skills)
        ReDim arr(ub)
        For i = 0 To ub
            arr(i) = Me.Skills(i)
        Next
        copy.Skills = arr
        Set ICloneable_Clone = copy
    End Function
End Class

' Demo Execution
Dim r1 As MyResume
Dim r2 As ICloneable
Dim r2Copy As MyResume
Set r1 = New MyResume
r1.Name = "张三"
r1.Age = 25
r1.Skills = Array("VBScript", "HTML")

Set r2 = r1.Clone()         ' Interface method return → interface variable
Set r2Copy = r2             ' Interface variable → concrete class variable
r2Copy.Name = "李四"         ' ← CRASH: assignment to entry in nil map
r2Copy.Skills(0) = "JavaScript"

Response.Write(r1.Name & " " & r1.Skills(0))
Response.Write(r2Copy.Name & " " & r2Copy.Skills(0))
%>`

	output, err := runVBScriptTest(code)
	if err != nil {
		t.Fatalf("Execution failed: %v", err)
	}

	expected := "张三 VBScript李四 JavaScript"
	if output != expected {
		t.Errorf("Expected output %q, got %q", expected, output)
	}
}
