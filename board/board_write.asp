<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim u_type, board_seq, board_gubun, condi, condi_value
Dim page, ck_sw, rs
Dim reg_id, reg_name, board_title, board_body, att_file
Dim pass, ins_gubun, gubun_name, title_line
Dim rsBoard

u_type = Request("u_type")
board_seq = Request("board_seq")
board_gubun = Request("board_gubun")
condi = Request("condi")
condi_value = Request("condi_value")
page = Request("page")
ck_sw = Request("ck_sw")

reg_id = user_id
reg_name = user_name
board_title = ""
board_body = ""
att_file = ""
pass = ""

If u_type = "U" Then
	ins_gubun = "변경"
Else
  	ins_gubun = "등록"
End If

If board_gubun = "1" Then
	gubun_name = "사내공지"
ElseIf board_gubun = "2" Then
  	gubun_name = "사내게시판"
ElseIf board_gubun = "3" Then
  	gubun_name = "A/S공지"
ElseIf board_gubun = "4" Then
  	gubun_name = "자료실"
Else
  	gubun_name = "전체게시판"
End If

title_line = "게시판 " & ins_gubun

If u_type = "U" Then
	'Sql="select * from board where board_seq="&board_seq
	objBuilder.Append "SELECT board_seq, board_gubun, board_title, board_body, "
	objBuilder.Append "	pass, att_file, reg_id, reg_name "
	objBuilder.Append "FROM board WHERE board_seq = " & board_seq

	Set rsBoard = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	board_seq = rsBoard("board_seq")
	board_gubun = rsBoard("board_gubun")
	board_title = rsBoard("board_title")
	board_body = rsBoard("board_body")
	pass = rsBoard("pass")
	att_file = rsBoard("att_file")
	reg_id = rsBoard("reg_id")
	reg_name = rsBoard("reg_name")

	rsBoard.Close() : Set rsBoard = Nothing
	DBConn.Close() : Set DBConn = Nothing
End If
%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">-->
<!DOCTYPE html>
<html lang="ko">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>NKP 시스템</title>
	<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<!--<script type="text/javascript" src="/java/js_window.js"></script>-->

	<script type="text/javascript">
		function getPageCode(){
			return "1 1";
		}

		function goAction(){
		   window.close();
		}

		function frmcheck(){
			if(chkfrm()){
				document.frm.submit();
			}
		}

		function chkfrm(){
			const browser = isBrowserCheck();
			var board_body;

			//브라우저 별 내용 값 호출 방식 설정
			if(browser === 'Mozilla'){
				board_body = $('textarea[name=board_body]').val();
			}else{
				board_body = FCKeditorAPI.GetInstance('board_body').GetXHTML(true);
			}

			k = 0;

			for (j=0;j<3;j++){
				if (eval("document.frm.board_gubun[" + j + "].checked")){
					k = k + 1
				}
			}

			if (k==0){
				alert ("게시판 위치를 체크하세요");
				return false;
			}

			if(document.frm.board_title.value ==""){
				alert('제목을 입력하세요');
				frm.board_title.focus();
				return false;
			}

			if(isEmpty(board_body)){
				alert('내용을 입력하세요');

				//브라우저 별 내용 값 호출 방식 설정
				if(browser === 'Mozilla'){
					$('textarea[name=board_body]').focus();
				}else{
					FCKeditorAPI.GetInstance('board_body').Focus();
				}
				return false;
			}

			if(document.frm.pass.value ==""){
				alert('비밀번호를 입력하세요');
				frm.pass.focus();
				return false;
			}

			if(!confirm('등록 하시겠습니까?')) return false;
			else return true;

		}

		//사용 용도 모름[허정호_20210330]
		function file_browse(){
			document.frm.att_file.click();
			document.frm.text1.value=document.frm.att_file.value;
		}
	</script>
</head>
<body>
	<div id="wrap">
		<div id="container">
			<h3 class="tit"><%=title_line%></h3>
			<form method="post" name="frm" action="/board/board_write_ok.asp" enctype="multipart/form-data">

			<div class="gView">
				<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
					<tbody>
						<tr>
						  <th scope="row" width="15%">게시판위치</th>
							<td class="left">
							  <input type="radio" name="board_gubun" value="1" <%If board_gubun = "1" Then %>checked<%End If %> style="width:30px">
							  사내공지
							  <input type="radio" name="board_gubun" value="3" <%If board_gubun = "3" Then %>checked<%End If %> style="width:30px">
							  A/S공지
							  <input type="radio" name="board_gubun" value="4" <%If board_gubun = "4" Then %>checked<%End If %> style="width:30px">
							자료실 </td>
						</tr>
						<tr>
						  <th scope="row" width="15%">작성자</th>
							<td class="left"><%=reg_name%>&nbsp;(&nbsp;<%=reg_id%>&nbsp;)</td>
						</tr>
						<tr>
							<th scope="row">제목</th>
						<td class="left"><input type="text" name="board_title" value="<%=board_title%>" onkeyup="checklength(this,100)" title="제목" style="width:800px;"/></td>
						</tr>
						<tr>
							<th scope="row">내용</th>
							<td>&nbsp;
							<!--#include virtual="/fckeditor/fckeditor.asp" -->
							<%
							 Dim oFCKeditor, sBasePath
							 sBasePath = "/FCKeditor/"  '<- 절대 경로

							 Set oFCKeditor = New FCKeditor
							 oFCKeditor.BasePath = sBasePath

							 oFCKeditor.Width = "98%"
							 oFCKeditor.Height = "300" '<- 높이변경
							 oFCKeditor.ToolbarSet = "Default" '<-메뉴스타일 변경
							 oFCKeditor.Value = board_body '<-초기 텍스트 변경

							 oFCKeditor.Config("UseBROnCarriageReturn") = true '<-Enter키 사용시 br적용 여부

							' oFCKeditor.Config("SkinPath") = sBasePath + "editor/skins/office2003/"   '<-  스킨 변경

							 oFCKeditor.Create "board_body"
							%>
							</td>
						</tr>
						<tr>
							<th scope="row">첨부파일</th>
							<td class="left">
							<p><%=att_file%>
							  <label>
								<input name="v_att_file" type="hidden" id="v_att_file" value="<%=att_file%>">
							  </label>
							</p>
							<input type="file" name= "att_file"  size="70" accept="image/gif">
							* 첨부파일은 1개만 가능하며 최대용량은 8MB ( 필요시 압축해서 첨부)
							</td>
						</tr>
						<tr>
							<th scope="row">패스워드</th>
							<td class="left"><input name="pass" type="text" title="패스워드" style="width:80px;" maxlength="4" notnull errname="패스워드"/>
							  * 수정 또는 삭제시 반드시 필요합니다.
							 </td>
						</tr>
					</tbody>
				</table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
				<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
			</div>
			<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
			<input type="hidden" name="board_seq" value="<%=board_seq%>" ID="Hidden1">
			<input type="hidden" name="condi" value="<%=condi%>" ID="Hidden1">
			<input type="hidden" name="condi_value" value="<%=condi_value%>" ID="Hidden1">
			<input type="hidden" name="page" value="<%=page%>" ID="Hidden1">
			<input type="hidden" name="ck_sw" value="<%=ck_sw%>" ID="Hidden1">
		</form>
	</div>
</div>
</body>
</html>

