<%
'/***************************************************************
'   �ۼ���     : ������ (lyoul@k-net.or.kr)
'   �����Ϸ��� : 2001.12.03
'   ��  ��     : js_dialog.asp
'   ��  ��     : ������ �ڹٽ�ũ��Ʈ ���̾˷α� �ڽ��� ���� �Լ�
'****************************************************************/

' �޼��� ���
	Sub js_back (msg)

%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			history.back();
		</script>
<%
	End Sub


	' �޼��� ���
	Sub js_msg (msg)

%>
		<script language="javascript">
			alert ("<%=msg%>") ;
		</script>
<%
	End Sub


	' â�ݱ�
	Sub js_close ()

%>
		<script language="javascript">
			window.close() ;
		</script>
<%
	End Sub

	' �޼��� ��� �� �ǵ��ư���
	Sub js_msg_back (msg)

%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			history.back () ;
		</script>
<%
	End Sub

	' �޼��� ��� �� �����̷�Ʈ
	Sub js_msg_redirect (msg, url)
%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			location.href = "<%=url%>" ;
		</script>
<%
	End Sub

	' �޼��� ��� ���� �ٷ� �����̷�Ʈ
	Sub js_redirect (url)
%>
		<script language="javascript">
			location.href = "<%=url%>" ;
		</script>
<%
	End Sub

	' �޼��� ��� �� ������ �ݱ�
	Sub js_msg_close (msg)
%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			window.close () ;
		</script>
<%
	End Sub
   '�˾�â�ݰ� reload
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
	'�˾�â �ݰ� redirect
	Sub js_msg_close_redirect (msg, url)
%>
		<script language="javascript">
			alert ("<%=msg%>") ;
			opener.location.href = "<%=url%>" ;
		  window.close () ;
		</script>
<%
	End sub

	'�˾�â �ݰ� opener redirect(�޽��� ����)
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
