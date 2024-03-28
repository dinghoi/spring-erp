<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_nologin.asp" -->
<!--#include virtual="/common/func.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
	<link href="css/login.css" rel="stylesheet" type="text/css">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<script type="text/javascript">
		$(document).ready(function(){
			$("#btn_submit").on("click", function(){
				window.open("", "popup_window", "width=1500, height=600, scrollbars=yes, resizable=yes");
				$("#hanbiro").submit();
			});

			//var params = { "id" : "lhs0806" };
			var params = { "id" : "jungho_heo" };

			$.ajax({
				url: "http://gw.k-one.co.kr/ngw/approval/sso/token"
				,async: false
				,type: 'get'
				,data: params
				,dataType: "jsonp" // crossdomain
				,success: function(data){
					//console.log(data.data.token);
					console.log(data)

					$("input[name=token]").val( data.data.token );
				}
			});
		});
	</script>
	<title>NKP시스템 한비로 그룹웨어 연동</title>
</head>
<body topmargin="0" leftmargin="0">
	한비로 test
	<form name="hanbiro" id="hanbiro" action="http://gw.k-one.co.kr/ngw/approval/sso/write_form" target="popup_window">
		<input type="text" name="token" value=""><br>
		<input type="text" name="callback" value="http://intra.k-won.co.kr/hanbiro_callback.asp"><br>
		<input type="text" name="formname" value="기본양식"><br>

		<textarea rows="4" cols="50" name="content" form="hanbiro">
			<html>
			<table >
				<tr>
					<td>1</td>
					<td>2</td>
					<td>3</td>
				</tr>
				<tr>
					<td>4</td>
					<td>5</td>
					<td>6</td>
				</tr>
			</table>
			</html>
		</textarea>
		<br>
		<button type="button" id="btn_submit">전송</button>
	</form>
</body>
</html>
