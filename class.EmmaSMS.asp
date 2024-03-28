<%

Class EmmaSMS

	Private Param
	Public ErrMsg
	Public LastPoint

	'***************************************************
	' 클래스를 시작되면 개체를 생성하는 이벤트 핸들러 
	'***************************************************
	Private Sub Class_Initialize
		Set Param = server.CreateObject("Scripting.Dictionary")
	End Sub


	'***************************************************
	' 파일을 불러들여 템플릿을 정의하는 함수
	'***************************************************
	Public Function login(id, pass)
		Param.Item("Id") = id
		Param.Item("Pass") = pass
	End Function


	'***************************************************
	' 파일을 불러들여 템플릿을 정의하는 함수
	'***************************************************
	Public Function send(sms_to, sms_from, sms_message, sms_date)
		Param.Item("To") = sms_to
		Param.Item("From") = sms_from
		Param.Item("Message") = Server.UrlEncode(sms_message)
		Param.Item("Date") = sms_date

'	Response.write "sms_to: " & sms_to & " sms_from: "& sms_from& " sms_msg: "&sms_msg& " sms_date: " & sms_date

		Dim args(1), Answer
		Set args(0) = Param

		RPC_URL = "http://whoisweb.net/emma/API/EmmaSend_ASP.php"
		Set Answer = xmlRPC(RPC_URL, "EmmaSend", args)
		
		If Answer("Code") <> "00" then
			ErrMsg = Answer("CodeMsg")
			send = false
'			Response.write ErrMsg
		Else 
			LastPoint = Answer("LastPoint")
			send = true
		End If

		Set args(0) = Nothing
		Set Answer = Nothing

	End Function


	'***************************************************
	' 클래스를 소멸하면서 개체를 소멸하는 이벤트 핸들러 
	'***************************************************
	Private Sub Class_Terminate 
		Set Param = Nothing  
	End Sub 


End Class

%>