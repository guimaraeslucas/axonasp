# Work with ADODB and ADODB.Stream in AxonASP

## Overview
AxonASP provides ADODB compatibility objects for Connection, Recordset, Command, Field, Parameters, Errors collections, and ADODB.Stream operations.

## Syntax
```asp
Set conn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set cmd = Server.CreateObject("ADODB.Command")
Set stm = Server.CreateObject("ADODB.Stream")
`````

## Parameters and Arguments
- Connection methods: Open, Close, Execute, BeginTrans, CommitTrans, RollbackTrans, OpenSchema, Cancel.
- Connection properties: ConnectionString, ConnectionTimeout, CommandTimeout, CursorLocation, Provider, DefaultDatabase, Mode, IsolationLevel, State, Version, Errors.
- Recordset methods: Open, Close, MoveFirst, MoveLast, MoveNext, MovePrevious, Move, Find, GetRows, GetString, AddNew, Update, UpdateBatch, CancelUpdate, CancelBatch, Delete, Requery, Resync, Clone, NextRecordset, Supports, Save, Seek.
- Recordset properties: EOF, BOF, Fields, RecordCount, PageCount, PageSize, AbsolutePage, AbsolutePosition, Bookmark, Filter, Sort, Status, Source, ActiveCommand, EditMode.
- Command methods: Execute, CreateParameter, Cancel, Parameters.Refresh.
- Command properties: ActiveConnection, CommandText, CommandTimeout, CommandType, Prepared, Parameters.
- Field methods: AppendChunk, GetChunk.
- Field properties: Name, Type, Value, OriginalValue, UnderlyingValue, ActualSize, DefinedSize, Attributes, Precision, NumericScale, Status.
- Errors collection methods/properties: Clear, Item, Count.
- Error properties: Number, Source, Description, SQLState.
- ADODB.Stream methods: Open, Close, Read, ReadText, Write, WriteText, SaveToFile, LoadFromFile, CopyTo, SetEOS, SkipLine, Flush.
- ADODB.Stream properties: Type, State, Mode, Position, Size, EOS, Charset, LineSeparator.

## Return Values
ADODB calls return compatibility values following AxonASP dispatch behavior. Collection and object-returning members return native object handles.

## Remarks
- ADODBOLE.Connection is supported as a specific compatibility path.
- Unsupported provider features may return compatibility errors.
- Use explicit Open and Close sequences to control connection and stream lifetime.

## Code Example
```asp
<%
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("./data.mdb")
conn.Open

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM users", conn
Do While Not rs.EOF
  Response.Write rs.Fields("name").Value & "<br>"
  rs.MoveNext
Loop

rs.Close
conn.Close
%>
`````
