<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Contador de reproducciones
'  Se invoca desde el reproductor (sendBeacon) cada vez que empieza una pista.
' =============================================================================
Dim pd_id
Auth_Load

Response.ContentType = "text/plain"
Response.CharSet = "utf-8"

pd_id = ToInt(QStr("id"))

If g_IsLogged And pd_id > 0 Then
    DB_Run "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = @p0", Array(pd_id)
    Response.Write "ok"
Else
    Response.Write "skip"
End If
%>
