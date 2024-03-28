<!--#include virtual="/common/inc_top.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>팝업 이미지 업로드</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<script type="text/javascript">
		function frmcheck(){
			if (formcheck(document.frm) && chkfrm()){
				document.frm.submit ();
			}
		}

		function chkfrm(){
			//if(document.frm.up_image.value == ""){
			if($('#up_image').val() == ""){
				alert('업로드 파일을 선택해 주세요.');
				frm.up_image.focus();
				return false;
			}

//			file_name = document.frm.up_image.value;
//			file_type = file_name.slice(file_name.lastindexof(".")).tolowercase();

			if(document.frm.up_pass.value != "123456"){
				alert('비밀번호를 확인해 주세요.');
				frm.up_pass.focus();
				return false;
			}

			if(!confirm('업로드 하시겠습니까?')) return false;
			else return true;
		}
	</script>
</head>
<body>
	<div id="container">
		<h3 class="tit">팝업창 이미지 업로드</h3>
		<form action="/board/popup_file_up_ok.asp" method="post" name="frm" enctype="multipart/form-data">
		<div class="gView">
			<table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="30%" >
					<col width="*" >
				</colgroup>
				<tbody>
					<tr>
						<th class="first">업로드파일</th>
						<td class="left"><input name="up_image" type="file" id="up_image" size="30"></td>
					</tr>
					<tr>
						<th class="first">비밀번호</th>
						<td class="left"><input name="up_pass" type="password" id="up_pass" style="width:150px"></td>
					</tr>
				</tbody>
			</table>
		</div>
		<br>
		<div align="center">
			<span class="btnType01"><input type="button" value="업로드" onclick="frmcheck();"></span>
			<span class="btnType01"><input type="button" value="취소" onclick="close_win();"></span>
		</div>
		</form>
	</div>
</body>
</html>