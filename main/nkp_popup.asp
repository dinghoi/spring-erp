<!--#include virtual="/common/inc_top.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
	<TITLE>popup</TITLE>
	<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=euc-kr">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<style type="text/css">
	<!--
	.style1 {font-size: 12px}

	body{
		background-color:#FFFFFF;
		margin-left:0px;
		margin-top:0px;
	}

	table{
		width:400px;
		border:0;
		padding:0;
		border-spacing:0;
	}

	.pop_img{
		width:635px;
		height:603px;
	}

	.td_win_close{
		 width:585px;
		 height:25px;
		 vertical-align:middle;
	}

	.div_win_close{
		text-align:right;
	}

	.td_chkBox{
		width:50px;
		height:25px;
		text-align:center;
		vertical-align:middle;
	}
	-->
	</style>

	<script type="text/javascript">
	function setCookie(cname, cvalue, exdays){
		var d = new Date();
		d.setTime(d.getTime() + (exdays*24*60*60*1000));

		var expires = "expires="+ d.toUTCString();
		document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
	}

	// '오늘만 이 창을 열지 않음' 클릭
	function closePop(){
		//if(document.forms[0].todayPop.checked){
			setCookie('nkp_popup', 'nkp_popup', 1);
		//}

		self.close();
	}

	//function closewin(){
	//  var expire = new Date();
	//  expire.setDate(expire.getDate() - 1);
	//  document.cookie = "ww2=1; expires=" + expire.toGMTString()+ "; path=/";
	//
	//  self.close();
	//}

	</script>
</head>
<body>
<!-- ImageReady Slices (popup.psd) -->
 	<table>
    <tr>
		<td colspan="2">
			<img src="/image/nkp_popup.gif" class="pop_img">
		</td>
    </tr>
    <tr>
		<td class="td_win_close">
			<div class="div_win_close">
				<span class="style1"><strong>오늘만 이 창을 열지 않음</strong></span>
			</div>
		</td>
		<td class="td_chkBox">
			<input name="todayPop" type="checkbox" id="todayPop" onClick="closePop();" value="checkbox">
		</td>
    </tr>
  </table>
<!-- End ImageReady Slices -->
</body>
</html>