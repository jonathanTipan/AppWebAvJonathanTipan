<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Punto de entrada
'  Redirige a la pantalla principal si hay sesion, o al login si no la hay.
' =============================================================================
Auth_Load

If g_IsLogged Then
    Go AppUrl("home.asp")
Else
    Go AppUrl("login.asp")
End If
%>
