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
	ins_gubun = "����"
Else
  	ins_gubun = "���"
End If

If board_gubun = "1" Then
	gubun_name = "�系����"
ElseIf board_gubun = "2" Then
  	gubun_name = "�系�Խ���"
ElseIf board_gubun = "3" Then
  	gubun_name = "A/S����"
ElseIf board_gubun = "4" Then
  	gubun_name = "�ڷ��"
Else
  	gubun_name = "��ü�Խ���"
End If

title_line = "�Խ��� " & ins_gubun

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
	<title>NKP �ý���</title>
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

			//������ �� ���� �� ȣ�� ��� ����
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
				alert ("�Խ��� ��ġ�� üũ�ϼ���");
				return false;
			}

			if(document.frm.board_title.value ==""){
				alert('������ �Է��ϼ���');
				frm.board_title.focus();
				return false;
			}

			if(isEmpty(board_body)){
				alert('������ �Է��ϼ���');

				//������ �� ���� �� ȣ�� ��� ����
				if(browser === 'Mozilla'){
					$('textarea[name=board_body]').focus();
				}else{
					FCKeditorAPI.GetInstance('board_body').Focus();
				}
				return false;
			}

			if(document.frm.pass.value ==""){
				alert('��й�ȣ�� �Է��ϼ���');
				frm.pass.focus();
				return false;
			}

			if(!confirm('��� �Ͻðڽ��ϱ�?')) return false;
			else return true;

		}

		//��� �뵵 ��[����ȣ_20210330]
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
						  <th scope="row" width="15%">�Խ�����ġ</th>
							<td class="left">
							  <input type="radio" name="board_gubun" value="1" <%If board_gubun = "1" Then %>checked<%End If %> style="width:30px">
							  �系����
							  <input type="radio" name="board_gubun" value="3" <%If board_gubun = "3" Then %>checked<%End If %> style="width:30px">
							  A/S����
							  <input type="radio" name="board_gubun" value="4" <%If board_gubun = "4" Then %>checked<%End If %> style="width:30px">
							�ڷ�� </td>
						</tr>
						<tr>
						  <th scope="row" width="15%">�ۼ���</th>
							<td class="left"><%=reg_name%>&nbsp;(&nbsp;<%=reg_id%>&nbsp;)</td>
						</tr>
						<tr>
							<th scope="row">����</th>
						<td class="left"><input type="text" name="board_title" value="<%=board_title%>" onkeyup="checklength(this,100)" title="����" style="width:800px;"/></td>
						</tr>
						<tr>
							<th scope="row">����</th>
							<td>&nbsp;
							<!--#include virtual="/fckeditor/fckeditor.asp" -->
							<%
							 Dim oFCKeditor, sBasePath
							 sBasePath = "/FCKeditor/"  '<- ���� ���

							 Set oFCKeditor = New FCKeditor
							 oFCKeditor.BasePath = sBasePath

							 oFCKeditor.Width = "98%"
							 oFCKeditor.Height = "300" '<- ���̺���
							 oFCKeditor.ToolbarSet = "Default" '<-�޴���Ÿ�� ����
							 oFCKeditor.Value = board_body '<-�ʱ� �ؽ�Ʈ ����

							 oFCKeditor.Config("UseBROnCarriageReturn") = true '<-EnterŰ ���� br���� ����

							' oFCKeditor.Config("SkinPath") = sBasePath + "editor/skins/office2003/"   '<-  ��Ų ����

							 oFCKeditor.Create "board_body"
							%>
							</td>
						</tr>
						<tr>
							<th scope="row">÷������</th>
							<td class="left">
							<p><%=att_file%>
							  <label>
								<input name="v_att_file" type="hidden" id="v_att_file" value="<%=att_file%>">
							  </label>
							</p>
							<input type="file" name= "att_file"  size="70" accept="image/gif">
							* ÷�������� 1���� �����ϸ� �ִ�뷮�� 8MB ( �ʿ�� �����ؼ� ÷��)
							</td>
						</tr>
						<tr>
							<th scope="row">�н�����</th>
							<td class="left"><input name="pass" type="text" title="�н�����" style="width:80px;" maxlength="4" notnull errname="�н�����"/>
							  * ���� �Ǵ� ������ �ݵ�� �ʿ��մϴ�.
							 </td>
						</tr>
					</tbody>
				</table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
				<span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
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

