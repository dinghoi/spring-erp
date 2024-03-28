<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim rsEdu, title_line

title_line = "���� ����"

If m_seq = "" Or m_name = "" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('ȸ���⺻���� ��� �� �̿� �����մϴ�.');"
	Response.Write "	location.href='/member/member_add.asp';"
	Response.Write "</script>"

	Response.End
End If

objBuilder.Append "SELECT edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date, edu_comment, "
objBuilder.Append "	edu_seq "
objBuilder.Append "FROM member_edu "
objBuilder.Append "WHERE m_seq = '"&m_seq&"' "
objBuilder.Append "ORDER BY m_seq, edu_seq ASC;"

Set rsEdu = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ�������</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			//���� ��� �˾�
			function eduAddPopup(){
				var url = '/member/member_edu_add.asp';
				var pop_name = '�������� ���';
				var features = 'scrollbars=yes,width=750,height=350';

				pop_Window(url, pop_name, features);
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_individual_edu.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
                        <dd>
                            <p>
						    <strong>���� : </strong>
							<label>
								<input type="text" name="m_name" id="m_name" value="<%=m_name%>" readonly="true" style="width:150px; text-align:left"/>
							</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="5%" >
						</colgroup>
						<thead>
                            <tr>
                              <th colspan="3">����&nbsp;������</th>
                              <th colspan="2">�������</th>
                              <th colspan="2">����&nbsp;������No.</th>
                              <th colspan="2">����&nbsp;�Ⱓ</th>
                              <th colspan="3">����&nbsp;�ֿ�&nbsp;����</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsEdu.EOF Or rsEdu.BOF Then
							Response.Write "<tr><td colspan='11' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						Else
							Do Until rsEdu.EOF
						%>
							<tr>
                              <td colspan="3"><%=rsEdu("edu_name")%>&nbsp;</td>
                              <td colspan="2"><%=rsEdu("edu_office")%>&nbsp;</td>
                              <td colspan="2"><%=rsEdu("edu_finish_no")%>&nbsp;</td>
                              <td colspan="2"><%=rsEdu("edu_start_date")%>��<%=rsEdu("edu_end_date")%>&nbsp;</td>
                              <td colspan="3"><%=rsEdu("edu_comment")%>&nbsp;</td>
							</tr>
						<%
								rsEdu.MoveNext()
							Loop
						End If
						rsEdu.Close() : Set rsEdu = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="eduAddPopup();" class="btnType04">���� ���</a>
					</div>
                    </td>
			      </tr>
				</table>
			</form>
		</div>
	</div>
	</body>
</html>