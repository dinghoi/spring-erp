<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/asmg_dbcon.asp" -->
<!--#include virtual="/include/asmg_user.asp" -->
<%
u_type = request("u_type")
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

if u_type = "U" then
	ins_gubun = "����"
  else
  	ins_gubun = "���"
end if

if board_gubun = "1" then
	gubun_name = "�系����"
  elseif board_gubun = "2" then
  	gubun_name = "�系�Խ���"
  elseif board_gubun = "3" then
  	gubun_name = "A/S����"
  elseif board_gubun = "4" then
  	gubun_name = "�ڷ��"
  else
  	gubun_name = "��ü�Խ���"  
end if
title_line = gubun_name + " " + ins_gubun


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_gubun = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if u_type = "U" then
	Sql="select * from board where board_seq="&board_seq
	Set rs=DbConn.Execute(Sql)
	board_seq = rs("board_seq")
	board_gubun = rs("board_gubun")
	board_title = rs("board_title")
	board_body = rs("board_body")
	pass = rs("pass")
	att_file = rs("att_file")
	reg_id = rs("reg_id")
	reg_name = rs("reg_name")
	rs.close()
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html> 
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				a=confirm('����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
			}
			function file_browse() 
     		{ 
           		document.frm.att_file.click(); 
           		document.frm.text1.value=document.frm.att_file.value;  
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<div id="container">				
				<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="board_write_ok.asp" enctype="multipart/form-data">
			
				<div class="gView">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<tbody>
							<tr>
							  <th scope="row" width="15%">��ġ</th>
								<td class="left"><%=gubun_name%></td>
							</tr>
							<tr>
							  <th scope="row" width="15%">�ۼ���</th>
								<td class="left"><%=reg_name%>&nbsp;(&nbsp;<%=reg_id%>&nbsp;)</td>
							</tr>
							<tr>
								<th scope="row">����</th>
						    <td class="left"><input type="text" name="board_title" value="<%=board_title%>" onkeyup="checklength(this,100)" notnull errname="����" title="����" style="width:800px;"/></td>
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
									<input type="file" name= "att_file"  size="70"> 
									<input type="text" name="text1" size="70">
									<img src="/image/but_ser.jpg" onclick="file_browse()" style="cursor:pointer">
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
            <div align=center>
                <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:history.go(-1);"></span>
            </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="board_seq" value="<%=board_seq%>" ID="Hidden1">
                <input type="hidden" name="board_gubun" value="<%=board_gubun%>" ID="Hidden1">
                <input type="hidden" name="condi" value="<%=condi%>" ID="Hidden1">
                <input type="hidden" name="condi_value" value="<%=condi_value%>" ID="Hidden1">
                <input type="hidden" name="page" value="<%=page%>" ID="Hidden1">
                <input type="hidden" name="ck_sw" value="<%=ck_sw%>" ID="Hidden1">
            </form>
		</div>				
	</div>        				
	</body>
</html>

