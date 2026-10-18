<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Cierre de sesion
' =============================================================================
LogoutUser
Response.Redirect AppUrl("login.asp?bye=1")
%>
