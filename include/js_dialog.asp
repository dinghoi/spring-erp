<%
'/***************************************************************
'   작성자     : 조형렬 (lyoul@k-net.or.kr)
'   최종완료일 : 2001.12.03
'   파  일     : js_dialog.asp
'   설  명     : 간단한 자바스크립트 다이알로그 박스를 위한 함수
'****************************************************************/

' 메세지 출력
	Sub js_back (msg)

%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			history.back();
		</script>
<%
	End Sub


	' 메세지 출력
	Sub js_msg (msg)

%>
		<script language="javascript">
			alert ("<%=msg%>") ;
		</script>
<%
	End Sub


	' 창닫기
	Sub js_close ()

%>
		<script language="javascript">
			window.close() ;
		</script>
<%
	End Sub

	' 메세지 출력 후 되돌아가기
	Sub js_msg_back (msg)

%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			history.back () ;
		</script>
<%
	End Sub

	' 메세지 출력 후 리다이렉트
	Sub js_msg_redirect (msg, url)
%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			location.href = "<%=url%>" ;
		</script>
<%
	End Sub

	' 메세지 출력 없이 바로 리다이렉트
	Sub js_redirect (url)
%>
		<script language="javascript">
			location.href = "<%=url%>" ;
		</script>
<%
	End Sub

	' 메세지 출력 후 윈도우 닫기
	Sub js_msg_close (msg)
%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			window.close () ;
		</script>
<%
	End Sub
   '팝업창닫고 reload
	Sub js_prefresh_msg_redirect (msg, url)
%>
		<script language="javascript">
			opener.location.reload () ;
			alert ("<%=msg%>") ;
			//location.href = "<%=url%>" ;
		  window.close () ;
		</script>
<%
	End sub
	'팝업창 닫고 redirect
	Sub js_msg_close_redirect (msg, url)
%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			opener.location.href = "<%=url%>" ;
		  window.close () ;
		</script>
<%
	End sub

	'팝업창 닫고 opener redirect(메시지 없음)
	Sub js_nomsg_close_redirect (url)
%>
		<script language="javascript">
			opener.location.href = "<%=url%>" ;
	  	    window.close () ;
		</script>
<%
	End sub

	Sub js_prefresh_msg_close (msg)
%>
		<script language="javascript">
		   opener.location.reload () ;
		   alert ("<%=msg%>")
		   window.close () ;
		</script>
<%
	End Sub

	Sub js_prefresh_msg_reload (msg, url)
%>
		<script language="javascript">
			opener.location.reload () ;
			alert ("<%=msg%>") ;
			location.href = "<%=url%>" ;
		</script>
<%
	End Sub

	Sub js_prefresh_msg_back (msg)
%>
		<script language="javascript">
		   opener.location.reload () ;
		   alert ("<%=msg%>")
		   history.back () ;
		</script>
<%
	End Sub

	Sub js_msg_replace (msg, url)
%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			location.replace ("<%=url%>") ;
		</script>
<%
	End Sub

	Sub js_replace (url)
%>
		<script language="javascript">
			location.replace ("<%=url%>") ;
		</script>
<%
	End Sub

	Sub js_replace_with_size (url)
%>
		<script language="javascript">
			var			url = "<%=url%>" ;

			if ( url.indexOf (url, "?") == -1 )
				url = url + "?"
			else
				url = url + "&"

			location.replace (url + "awidth=" +
					screen.availWidth + "&aheight=" +
					screen.availHeight + "&wtop=" +
					window.screenTop + "&wleft=" +
					window.screenLeft + "&wwidth=" +
					screen.width + "&wheight=" +
					screen.height) ;
		</script>
<%
	End Sub

	Sub js_msg_replace_with_size (msg, url)
%>
		<script language="javascript">
			var			url = "<%=url%>" ;

			alert ("<%=msg%>") ;
			if ( url.indexOf (url, "?") == -1 )
				url = url + "?"
			else
				url = url + "&"

			location.replace (url + "awidth=" +
					screen.availWidth + "&aheight=" +
					screen.availHeight + "&wtop=" +
					window.screenTop + "&wleft=" +
					window.screenLeft + "&wwidth=" +
					screen.width + "&wheight=" +
					screen.height) ;
		</script>
<%
	End Sub
%>
