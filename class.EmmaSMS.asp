<%

Class EmmaSMS

	Private Param
	Public ErrMsg
	Public LastPoint

	'***************************************************
	' Ŭ������ ���۵Ǹ� ��ü�� �����ϴ� �̺�Ʈ �ڵ鷯 
	'***************************************************
	Private Sub Class_Initialize
		Set Param = server.CreateObject("Scripting.Dictionary")
	End Sub


	'***************************************************
	' ������ �ҷ��鿩 ���ø��� �����ϴ� �Լ�
	'***************************************************
	Public Function login(id, pass)
		Param.Item("Id") = id
		Param.Item("Pass") = pass
	End Function


	'***************************************************
	' ������ �ҷ��鿩 ���ø��� �����ϴ� �Լ�
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
	' Ŭ������ �Ҹ��ϸ鼭 ��ü�� �Ҹ��ϴ� �̺�Ʈ �ڵ鷯 
	'***************************************************
	Private Sub Class_Terminate 
		Set Param = Nothing  
	End Sub 


End Class

%>