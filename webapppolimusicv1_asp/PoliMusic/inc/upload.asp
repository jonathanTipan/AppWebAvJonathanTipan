<%
' =============================================================================
'  PoliMusic - Lectura de formularios y subida de archivos
'
'  La clase clsForm unifica los dos formatos que puede enviar un formulario:
'    - application/x-www-form-urlencoded  (sin archivos)
'    - multipart/form-data                (con archivos)
'  De este modo las paginas siempre usan la misma interfaz:
'      Set frm = New clsForm : frm.Load
'      valor = frm.Val("campo")
'      If frm.HasFile("archivo") Then ruta = frm.SaveFile(...)
' =============================================================================

' -----------------------------------------------------------------------------
'  Conversiones entre cadenas normales y cadenas de bytes
' -----------------------------------------------------------------------------
Function StrToBytes(s)
    Dim i, r
    r = ""
    For i = 1 To Len(s)
        r = r & ChrB(AscW(Mid(s, i, 1)) And 255)
    Next
    StrToBytes = r
End Function

Function BytesToStr(b)
    Dim i, n, r
    n = LenB(b)
    r = ""
    For i = 1 To n
        r = r & Chr(AscB(MidB(b, i, 1)))
    Next
    BytesToStr = r
End Function

' Decodifica una cadena de bytes UTF-8 a texto Unicode.
Function Utf8ToStr(b)
    Dim i, n, c, c2, c3, c4, code, r
    n = LenB(b)
    i = 1
    r = ""
    Do While i <= n
        c = AscB(MidB(b, i, 1))
        If c < 128 Then
            r = r & Chr(c)
            i = i + 1
        ElseIf c >= 192 And c < 224 Then
            c2 = AscB(MidB(b, i + 1, 1))
            r = r & ChrW(((c And 31) * 64) + (c2 And 63))
            i = i + 2
        ElseIf c >= 224 And c < 240 Then
            c2 = AscB(MidB(b, i + 1, 1))
            c3 = AscB(MidB(b, i + 2, 1))
            r = r & ChrW(((c And 15) * 4096) + ((c2 And 63) * 64) + (c3 And 63))
            i = i + 3
        Else
            c2 = AscB(MidB(b, i + 1, 1))
            c3 = AscB(MidB(b, i + 2, 1))
            c4 = AscB(MidB(b, i + 3, 1))
            code = ((c And 7) * 262144) + ((c2 And 63) * 4096) + ((c3 And 63) * 64) + (c4 And 63)
            code = code - 65536
            r = r & ChrW(55296 + (code \ 1024)) & ChrW(56320 + (code Mod 1024))
            i = i + 4
        End If
    Loop
    Utf8ToStr = r
End Function

' -----------------------------------------------------------------------------
'  clsForm
' -----------------------------------------------------------------------------
Class clsForm
    Public Form         ' Scripting.Dictionary con los campos de texto
    Public Files        ' Scripting.Dictionary con los archivos subidos
    Public Errors       ' Mensajes de error acumulados

    Private m_data
    Private m_boundary
    Private m_isMultipart

    Private Sub Class_Initialize()
        Set Form = Server.CreateObject("Scripting.Dictionary")
        Set Files = Server.CreateObject("Scripting.Dictionary")
        Form.CompareMode = 1
        Files.CompareMode = 1
        Errors = ""
        m_data = ""
        m_boundary = ""
        m_isMultipart = False
        Randomize
    End Sub

    Public Sub Load()
        m_isMultipart = (InStr(LCase(N(Request.ServerVariables("CONTENT_TYPE"))), "multipart/form-data") > 0)
        If m_isMultipart Then
            ParseMultipart
        Else
            LoadUrlEncoded
        End If
    End Sub

    Public Property Get IsMultipart()
        IsMultipart = m_isMultipart
    End Property

    Public Function Val(fieldName)
        If Form.Exists(fieldName) Then
            Val = N(Form(fieldName))
        Else
            Val = ""
        End If
    End Function

    Public Function IntVal(fieldName, defaultValue)
        IntVal = ToIntDef(Val(fieldName), defaultValue)
    End Function

    Public Function HasFile(fieldName)
        Dim fi
        HasFile = False
        If Files.Exists(fieldName) Then
            Set fi = Files(fieldName)
            HasFile = (fi("size") > 0)
        End If
    End Function

    Public Function FileName(fieldName)
        Dim fi
        FileName = ""
        If Files.Exists(fieldName) Then
            Set fi = Files(fieldName)
            FileName = fi("filename")
        End If
    End Function

    ' Guarda el archivo subido en destFolder y devuelve la ruta relativa que se
    ' almacena en la base de datos (por ejemplo "uploads/songFiles/x.mp3").
    ' Devuelve "" si no habia archivo o si hubo un error.
    Public Function SaveFile(fieldName, destFolder, prefix, allowedExts, maxBytes)
        Dim fi, ext, fname, dest, phys, fso, stream, ok, p

        SaveFile = ""
        If Not Files.Exists(fieldName) Then Exit Function
        Set fi = Files(fieldName)
        If fi("size") <= 0 Then Exit Function

        If fi("size") > maxBytes Then
            Errors = Errors & "El archivo " & fi("filename") & " supera el límite de " & _
                     FormatNumber(maxBytes / 1048576, 0) & " MB. "
            Exit Function
        End If

        ext = LCase(GetExtension(fi("filename")))
        If Len(ext) = 0 Then ext = ".dat"
        If InStr(allowedExts, "|" & ext & "|") = 0 Then
            Errors = Errors & "Tipo de archivo no permitido: " & fi("filename") & ". "
            Exit Function
        End If

        fname = prefix & "_" & StampNow() & "_" & CStr(Int(Rnd() * 9000) + 1000) & ext

        On Error Resume Next
        Set fso = FSObj()
        phys = Server.MapPath(destFolder)
        If Not fso.FolderExists(phys) Then fso.CreateFolder phys
        If Len(destFolder) > 0 Then
            If Right(destFolder, 1) <> "/" Then destFolder = destFolder & "/"
        End If
        dest = phys & "\" & fname

        Set stream = Server.CreateObject("ADODB.Stream")
        If Err.Number = 0 Then
            stream.Type = 1
            stream.Open
            stream.Write PartBytes(fi)
            stream.SaveToFile dest, 2
            stream.Close
        End If
        If Err.Number <> 0 Then
            Errors = Errors & "No se pudo guardar " & fi("filename") & ": " & Err.Description & ". "
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

        Set stream = Nothing
        SaveFile = destFolder & fname
    End Function

    ' -------------------------------------------------------------------------
    Private Sub LoadUrlEncoded()
        Dim k
        For Each k In Request.Form
            Form(CStr(k)) = N(Request.Form(k))
        Next
    End Sub

    Private Sub ParseMultipart()
        Dim ct, bp, delim, dashdash, crlf, pos, partStart, nextPos, bodyLen, body

        ct = N(Request.ServerVariables("CONTENT_TYPE"))
        bp = InStr(LCase(ct), "boundary=")
        If bp = 0 Then
            Errors = "No se encontró el boundary del formulario."
            Exit Sub
        End If

        m_boundary = Trim(Mid(ct, bp + 9))
        If Left(m_boundary, 1) = """" Then
            m_boundary = Mid(m_boundary, 2)
            bp = InStr(m_boundary, """")
            If bp > 0 Then m_boundary = Left(m_boundary, bp - 1)
        Else
            bp = InStr(m_boundary, ";")
            If bp > 0 Then m_boundary = Left(m_boundary, bp - 1)
        End If
        m_boundary = Trim(m_boundary)

        If Request.TotalBytes <= 0 Then
            Errors = "La petición no contiene datos."
            Exit Sub
        End If

        m_data = Request.BinaryRead(Request.TotalBytes)

        crlf = ChrB(13) & ChrB(10)
        dashdash = StrToBytes("--")
        delim = dashdash & StrToBytes(m_boundary)

        pos = InStrB(1, m_data, delim)
        Do While pos > 0
            partStart = pos + LenB(delim)
            If partStart > LenB(m_data) Then Exit Do
            If MidB(m_data, partStart, 2) = dashdash Then Exit Do
            If MidB(m_data, partStart, 2) = crlf Then partStart = partStart + 2

            nextPos = InStrB(partStart, m_data, delim)
            If nextPos = 0 Then Exit Do

            bodyLen = (nextPos - 2) - partStart
            If bodyLen < 0 Then bodyLen = 0
            body = MidB(m_data, partStart, bodyLen)

            HandlePart body, partStart
            pos = nextPos
        Loop
    End Sub

    Private Sub HandlePart(body, bodyOffset)
        Dim sep, hdrStr, lines, i, line, contentBytes
        Dim fieldName, fileName, contentType, d, contentLen

        sep = InStrB(body, ChrB(13) & ChrB(10) & ChrB(13) & ChrB(10))
        If sep = 0 Then Exit Sub

        hdrStr = BytesToStr(MidB(body, 1, sep - 1))

        contentLen = LenB(body) - sep - 3
        If contentLen < 0 Then contentLen = 0
        contentBytes = MidB(body, sep + 4, contentLen)

        fieldName = ""
        fileName = ""
        contentType = ""

        lines = Split(hdrStr, vbCrLf)
        For i = 0 To UBound(lines)
            line = Trim(lines(i))
            If LCase(Left(line, 20)) = "content-disposition:" Then
                fieldName = ExtractAttr(line, "name")
                fileName = ExtractAttr(line, "filename")
            ElseIf LCase(Left(line, 13)) = "content-type:" Then
                contentType = Trim(Mid(line, 14))
            End If
        Next

        If Len(fieldName) = 0 Then Exit Sub

        If Len(fileName) > 0 Then
            ' IMPORTANTE: no se guarda el contenido como cadena de bytes, porque
            ' MidB devuelve un String y ADODB.Stream.Write solo acepta un array de
            ' bytes (error 3001). Se guardan las coordenadas del fragmento dentro
            ' de m_data y SaveFile extrae el array de bytes real cuando lo necesita.
            Set d = Server.CreateObject("Scripting.Dictionary")
            d("filename") = StripPath(fileName)
            d("contenttype") = contentType
            d("size") = contentLen
            d("start") = bodyOffset + sep + 3      ' posicion 1-based en m_data
            d("len") = contentLen
            Set Files(fieldName) = d
        Else
            Form(fieldName) = Utf8ToStr(contentBytes)
        End If
    End Sub

    ' Extrae un array de bytes real a partir de las coordenadas guardadas.
    ' ADODB.Stream.Write solo acepta Byte(), nunca un String.
    Private Function PartBytes(fi)
        Dim st
        Set st = Server.CreateObject("ADODB.Stream")
        st.Type = 1
        st.Open
        st.Write m_data
        st.Position = CLng(fi("start")) - 1
        PartBytes = st.Read(CLng(fi("len")))
        st.Close
        Set st = Nothing
    End Function

    ' Busca atributo="valor" dentro de una cabecera, evitando confundir
    ' "name=" con "filename=".
    Private Function ExtractAttr(s, attr)
        Dim ls, la, p, q, prev
        ExtractAttr = ""
        ls = LCase(s)
        la = LCase(attr) & "="
        p = 1
        Do
            p = InStr(p, ls, la)
            If p = 0 Then Exit Do
            If p = 1 Then Exit Do
            prev = Mid(ls, p - 1, 1)
            If prev = " " Or prev = ";" Then Exit Do
            p = p + Len(la)
        Loop
        If p = 0 Then Exit Function
        p = p + Len(la)
        If Mid(s, p, 1) = """" Then
            p = p + 1
            q = InStr(p, s, """")
            If q = 0 Then Exit Function
            ExtractAttr = Mid(s, p, q - p)
        Else
            q = InStr(p, s, ";")
            If q = 0 Then q = Len(s) + 1
            ExtractAttr = Trim(Mid(s, p, q - p))
        End If
    End Function
End Class

' -----------------------------------------------------------------------------
'  Helpers de nombres de archivo
' -----------------------------------------------------------------------------
Function StripPath(s)
    Dim t, i
    t = Replace(N(s), "\", "/")
    i = InStrRev(t, "/")
    If i > 0 Then t = Mid(t, i + 1)
    t = Replace(t, """", "")
    StripPath = t
End Function

Function GetExtension(s)
    Dim t, i
    t = StripPath(s)
    i = InStrRev(t, ".")
    If i > 0 Then GetExtension = Mid(t, i) Else GetExtension = ""
End Function

Function StampNow()
    Dim d
    d = Now()
    StampNow = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & _
               Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function
%>
