<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<%
'==========================
'Author : 허정호
'Create Date : 20201117
'Desc : ASP 설정 코드
'==========================
Response.CharSet = "UTF-8"
Response.CodePage = "65001"
Response.ContentType = "text/html;charset=UTF-8"
Response.CodePage = "65001"

'Response.CharSet = "EUC-KR"
'Response.CodePage = "949"
'Response.ContentType = "text/html;charset=euc-kr"
'Response.CodePage = "949"

'no-cache 설정
Response.Expires = 0
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Cache-Control","no-cache,must-revalidate"
%>