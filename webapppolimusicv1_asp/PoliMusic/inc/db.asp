<%
' =============================================================================
'  PoliMusic - Capa de acceso a datos (ADO)
'  Todas las consultas usan parametros para evitar inyeccion SQL.
' =============================================================================

' Abre la conexion global probando las cadenas de conexion en orden.
Sub DB_Open()
    Dim i, cn, lastErr

    If IsObject(g_Conn) Then
        If Not (g_Conn Is Nothing) Then
            If g_Conn.State = 1 Then Exit Sub
        End If
    End If

    lastErr = ""
    For i = 0 To UBound(g_ConnStrings)
        On Error Resume Next
        Set cn = Server.CreateObject("ADODB.Connection")
        If Err.Number = 0 Then
            cn.ConnectionTimeout = 10
            cn.CommandTimeout = 45
            cn.Open g_ConnStrings(i)
        End If
        If Err.Number = 0 Then
            On Error GoTo 0
            Set g_Conn = cn
            Exit Sub
        End If
        lastErr = Err.Description
        Err.Clear
        On Error GoTo 0
        Set cn = Nothing
    Next

    Response.Write "<div style=""font-family:Segoe UI,Arial,sans-serif;max-width:720px;margin:60px auto;"">"
    Response.Write "<h2 style=""color:#c0392b"">No se pudo conectar a la base de datos</h2>"
    Response.Write "<p>Verifique que el servicio SQL Server este iniciado y que el login <b>" & DB_USER & "</b> exista en la base <b>" & DB_NAME & "</b>.</p>"
    Response.Write "<pre style=""background:#f4f6fa;padding:12px;border-radius:8px;white-space:pre-wrap"">" & Server.HTMLEncode(lastErr) & "</pre>"
    Response.Write "<p style=""color:#666"">Última cadena probada: " & Server.HTMLEncode(g_ConnStrings(UBound(g_ConnStrings))) & "</p>"
    Response.Write "</div>"
    Response.End
End Sub

Sub DB_Close()
    On Error Resume Next
    If IsObject(g_Conn) Then
        If Not (g_Conn Is Nothing) Then
            If g_Conn.State = 1 Then g_Conn.Close
        End If
    End If
    Set g_Conn = Nothing
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
'  Construye un parametro de tipo nvarchar (SQL Server convierte solo).
' -----------------------------------------------------------------------------
Function MakeParam(cmd, idx, val)
    Dim nm, s
    nm = "@p" & CStr(idx)
    If IsNull(val) Or IsEmpty(val) Then
        Set MakeParam = cmd.CreateParameter(nm, 202, 1, 1, Null)
    Else
        s = CStr(val)
        If Len(s) = 0 Then
            Set MakeParam = cmd.CreateParameter(nm, 202, 1, 1, "")
        ElseIf Len(s) > 4000 Then
            Set MakeParam = cmd.CreateParameter(nm, 202, 1, 4000, Left(s, 4000))
        Else
            Set MakeParam = cmd.CreateParameter(nm, 202, 1, Len(s), s)
        End If
    End If
End Function

' -----------------------------------------------------------------------------
'  Crea un ADODB.Command listo para ejecutar.
'
'  El proveedor MSOLEDBSQL solo admite parametros POSICIONALES ("?") cuando se
'  usa adCmdText: los parametros con nombre (@p0, @p1...) provocan el error
'  "Must declare the scalar variable @p0". Por eso aqui se reescriben las
'  consultas sobre la marcha: cada @pN se sustituye por "?" y se anade el valor
'  correspondiente en el mismo orden. Si un parametro aparece varias veces
'  (p.ej. @p1 en la consulta de login), su valor se repite tantas veces como
'  marcadores haya. Los "@pN" que esten dentro de un literal de cadena se
'  respetan y no se tocan.
' -----------------------------------------------------------------------------
Function MakeCmd(sql, prms)
    Dim cmd, i, L, ch, outS, j, numS, inLit, idxList, parts, k

    L = Len(sql)
    outS = ""
    idxList = ""
    i = 1
    inLit = False

    Do While i <= L
        ch = Mid(sql, i, 1)

        If ch = "'" Then
            outS = outS & ch
            If inLit Then
                ' Dentro de un literal: "''" es una comilla escapada.
                If i < L Then
                    If Mid(sql, i + 1, 1) = "'" Then
                        outS = outS & "'"
                        i = i + 2
                    Else
                        inLit = False
                        i = i + 1
                    End If
                Else
                    inLit = False
                    i = i + 1
                End If
            Else
                inLit = True
                i = i + 1
            End If

        ElseIf (Not inLit) And ch = "@" And LCase(Mid(sql, i + 1, 1)) = "p" Then
            j = i + 2
            numS = ""
            Do While j <= L
                If InStr("0123456789", Mid(sql, j, 1)) > 0 Then
                    numS = numS & Mid(sql, j, 1)
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            If Len(numS) > 0 Then
                outS = outS & "?"
                idxList = idxList & "|" & numS
                i = j
            Else
                outS = outS & ch
                i = i + 1
            End If

        Else
            outS = outS & ch
            i = i + 1
        End If
    Loop

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = g_Conn
    cmd.CommandType = 1                 ' adCmdText
    cmd.CommandTimeout = 45
    cmd.CommandText = outS

    If IsArray(prms) Then
        If Len(idxList) > 0 Then
            parts = Split(Mid(idxList, 2), "|")
            For k = 0 To UBound(parts)
                ' Defensa: si la consulta usa un @pN para el que no se ha pasado
                ' valor, se envia cadena vacia en lugar de provocar un error
                ' "Subscript out of range" que quedaria oculto por On Error.
                If CLng(parts(k)) <= UBound(prms) Then
                    cmd.Parameters.Append MakeParam(cmd, k, prms(CLng(parts(k))))
                Else
                    cmd.Parameters.Append MakeParam(cmd, k, "")
                End If
            Next
        End If
    End If

    Set MakeCmd = cmd
End Function

' -----------------------------------------------------------------------------
'  Devuelve un Recordset (usar para SELECT).
' -----------------------------------------------------------------------------
Function DB_RS(sql, prms)
    Dim cmd, rs
    DB_Open
    On Error Resume Next
    Set cmd = MakeCmd(sql, prms)
    Set rs = cmd.Execute
    If Err.Number = 0 Then
        ' Si el primer resultado no devuelve filas, ADO lo entrega cerrado:
        ' avanzar hasta el primer Recordset realmente abierto.
        Do While Not (rs Is Nothing)
            If rs.State = 1 Then Exit Do          ' adStateOpen
            Set rs = rs.NextRecordset
            If Err.Number <> 0 Then Exit Do
        Loop
    End If
    If Err.Number <> 0 Then
        Response.Write "<div style=""font-family:Consolas,monospace;background:#fff3f3;border:1px solid #f0b7b7;padding:14px;border-radius:8px;margin:16px 0"">"
        Response.Write "<b>Error SQL:</b> " & Server.HTMLEncode(Err.Description) & "<br><br><b>Consulta:</b><br>"
        Response.Write Server.HTMLEncode(sql) & "</div>"
        Err.Clear
        Set rs = Nothing
    End If
    On Error GoTo 0
    Set cmd = Nothing
    Set DB_RS = rs
End Function

' -----------------------------------------------------------------------------
'  Ejecuta una sentencia sin devolver filas (INSERT / UPDATE / DELETE).
' -----------------------------------------------------------------------------
Sub DB_Run(sql, prms)
    Dim cmd
    DB_Open
    On Error Resume Next
    Set cmd = MakeCmd(sql, prms)
    cmd.Execute , , 128                 ' adExecuteNoRecords
    If Err.Number <> 0 Then
        Response.Write "<div style=""font-family:Consolas,monospace;background:#fff3f3;border:1px solid #f0b7b7;padding:14px;border-radius:8px;margin:16px 0"">"
        Response.Write "<b>Error SQL:</b> " & Server.HTMLEncode(Err.Description) & "<br><br><b>Sentencia:</b><br>"
        Response.Write Server.HTMLEncode(sql) & "</div>"
        Err.Clear
    End If
    On Error GoTo 0
    Set cmd = Nothing
End Sub

' -----------------------------------------------------------------------------
'  Ejecuta un INSERT y devuelve el ID generado (SCOPE_IDENTITY).
'
'  El lote "INSERT ...; SELECT SCOPE_IDENTITY()" produce DOS resultados: el
'  primero (el INSERT) no devuelve filas y ADO lo entrega como un Recordset
'  CERRADO. Por eso hay que avanzar con NextRecordset hasta el resultado que si
'  tiene filas; si no, se obtiene el error 800a0e78 ("la operacion no esta
'  permitida si el objeto esta cerrado").
' -----------------------------------------------------------------------------
Function DB_Insert(sql, prms)
    Dim rs, id, cmd
    id = 0
    DB_Open
    On Error Resume Next
    Set cmd = MakeCmd(sql & "; SELECT CAST(SCOPE_IDENTITY() AS int) AS NEW_ID;", prms)
    Set rs = cmd.Execute
    Do While Not (rs Is Nothing)
        If rs.State = 1 Then Exit Do          ' adStateOpen
        Set rs = rs.NextRecordset
        If Err.Number <> 0 Then Exit Do
    Loop
    If Err.Number <> 0 Then
        Response.Write "<div style=""font-family:Consolas,monospace;background:#fff3f3;border:1px solid #f0b7b7;padding:14px;border-radius:8px;margin:16px 0"">"
        Response.Write "<b>Error SQL:</b> " & Server.HTMLEncode(Err.Description) & "<br><br><b>Sentencia:</b><br>"
        Response.Write Server.HTMLEncode(sql) & "</div>"
        Err.Clear
    ElseIf Not (rs Is Nothing) Then
        If rs.State = 1 Then
            If Not rs.EOF Then id = ToInt(rs("NEW_ID"))
            rs.Close
        End If
    End If
    On Error GoTo 0
    Set rs = Nothing
    Set cmd = Nothing
    DB_Insert = id
End Function

' -----------------------------------------------------------------------------
'  Devuelve el primer campo de la primera fila (o un valor por defecto).
' -----------------------------------------------------------------------------
Function DB_Scalar(sql, prms, fallback)
    Dim rs, v
    v = fallback
    Set rs = DB_RS(sql, prms)
    If IsObject(rs) Then
        If Not (rs Is Nothing) Then
            If Not rs.EOF Then
                If Not IsNull(rs(0)) Then v = rs(0)
            End If
            rs.Close
        End If
    End If
    Set rs = Nothing
    DB_Scalar = v
End Function

' -----------------------------------------------------------------------------
'  Devuelve el numero de filas de una consulta.
' -----------------------------------------------------------------------------
Function DB_Count(sql, prms)
    Dim rs, v
    v = 0
    Set rs = DB_RS(sql, prms)
    If IsObject(rs) Then
        If Not (rs Is Nothing) Then
            If Not rs.EOF Then
                If Not IsNull(rs(0)) Then v = CLng(rs(0))
            End If
            rs.Close
        End If
    End If
    Set rs = Nothing
    DB_Count = v
End Function

' -----------------------------------------------------------------------------
'  Lee un campo de un Recordset tratando los NULL como cadena vacia.
' -----------------------------------------------------------------------------
Function F(rs, fieldName)
    Dim v
    v = ""
    If IsObject(rs) Then
        If Not (rs Is Nothing) Then
            If Not rs.EOF Then
                If Not IsNull(rs(fieldName)) Then v = rs(fieldName)
            End If
        End If
    End If
    F = v
End Function
%>
