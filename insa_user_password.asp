<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	owner_view=Request.form("owner_view")
	view_condi = request.form("view_condi")
  else
	owner_view=request("owner_view")
	view_condi = request("view_condi")
end if

if view_condi = "" then
	view_condi = user_id
	owner_view = "T"
	ck_sw = "n"
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if view_condi <> "" then
     if owner_view = "C" then  
	          sql = "select * from memb where user_name like '%"+view_condi+"%'"
         else
	          sql = "select * from memb where user_id = '"+view_condi+"'"
     end if
	 Rs.Open Sql, Dbconn, 1
end if

title_line = "����� ��й�ȣ Ȯ�� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
	    </script>
		<script type="text/javascript">		
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("������ �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
			
			function user_password_modify(val) {

            if (!confirm("����� ��й�ȣ�� �ʱ�ȭ �Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
			document.frm.view_condi1.value = document.getElementById(val).value;
			
            document.frm.action = "insa_user_password_ok.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_user_password.asp?ck_sw=<%="n"%>" method="post" name="frm">
                <fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">���
                                </label>
							<strong>���� : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>                
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">�����</th>
								<td class="left"><%=rs("user_name")%>(<%=rs("user_id")%>)</td>
							</tr>
							<tr>
								<th class="first">������й�ȣ</th>
								<td class="left"><%=rs("pass")%>&nbsp;</td>
							</tr>
							<tr>
								<th class="first">�Ҽ�ȸ��</th>
								<td class="left"><%=rs("emp_company")%>&nbsp;</td>
							</tr>
							<tr>
								<th class="first">�Ҽ�</th>
								<td class="left"><%=rs("org_name")%>(<%=rs("team")%>)&nbsp;</td>
							</tr>
							<tr>
								<th class="first">�ڵ�����ȣ</th>
								<td class="left"><%=rs("hp")%>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="user_password_modify('view_condi');return false;" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="view_condi1" value="<%=view_condi%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

