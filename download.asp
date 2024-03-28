<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
  path = Request("path")
  att_file = request("att_file")

  ie_version = Request.ServerVariables("HTTP_USER_AGENT")
  Response.Clear
  Response.ContentType = "application/octet-stream"
  Response.AddHeader "Content-Disposition", "attachment;filename="&att_file
  Response.AddHeader "Content-Transfer-Encoding", "binary"
  Response.AddHeader "Pragma", "no-cache"
  Response.AddHeader "Expires", "0" 


  ' 스트림을 선언

  Set objStream = Server.CreateObject("ADODB.Stream")
  objStream.Open

  objStream.Type = 1
  objStream.LoadFromFile server.mappath(path)&"\"&att_file '파일경로
  strFile = objStream.Read

  Response.BinaryWrite strFile
  Set objStream = Nothing 

%>
