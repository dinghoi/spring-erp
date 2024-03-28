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
Dim rsFamily, arrFamily, title_line

If m_seq = "" Or m_name = "" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('ȸ���⺻���� ��� �� �̿� �����մϴ�.');"
	Response.Write "	location.href='/member/member_add.asp';"
	Response.Write "</script>"

	Response.End
End If

objBuilder.Append "SELECT f_seq, f_rel, f_name, f_birthday, "
objBuilder.Append "	f_birthday_id, f_job, f_person1, f_person2, f_tel_ddd, "
objBuilder.Append "	f_tel_no1, f_tel_no2, f_live "
objBuilder.Append "FROM member_family "
objBuilder.Append "WHERE m_seq = '"&m_seq&"' "

Set rsFamily = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "���� ����"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>ȸ������</title>
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
			//function familyAddPopup(id, name){
			function familyAddPopup(){
				var url = '/member/member_family_add.asp';
				var pop_name = '�������� ���';
				//var param = '?m_seq='+id+'&m_name='+name;
				var features = 'scrollbars=yes,width=750,height=450';

				//url += param;

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
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
                        <dd>
                            <p>
                            <strong>���� : </strong>
                            <label>
								<input type="text" name="m_name" size="10" value="<%=m_name%>" readonly />
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
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th colspan="2">����</th>
                                <th>����</th>
                                <th>�������</th>
                                <th colspan="2">����</th>
                                <th colspan="2">��ȭ��ȣ</th>
                                <th colspan="2">�ֹι�ȣ</th>
                                <th>���ſ���</th>
                                <th>No.</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						Dim i, f_seq, f_rel, f_name, f_birthday, f_birthday_id, f_job, f_person1
						Dim f_person2, f_tel_ddd, f_tel_no1, f_tel_no2, f_live

						If Not rsFamily.EOF Then
							arrFamily = rsFamily.getRows()

							For i = LBound(arrFamily) To UBound(arrFamily, 2)
								f_seq = arrFamily(0, i)
								f_rel = arrFamily(1, i)
								f_name = arrFamily(2, i)
								f_birthday  = arrFamily(3, i)
								f_birthday_id = arrFamily(4, i)
								f_job = arrFamily(5, i)
								f_person1 = arrFamily(6, i)
								f_person2 = f_toString(arrFamily(7, i), "")
								f_tel_ddd  = arrFamily(8, i)
								f_tel_no1 = arrFamily(9, i)
								f_tel_no2 = arrFamily(10, i)
								f_live = arrFamily(11, i)

								If f_person2 <> "" Then
									f_person2 = "*******"
								End If
						%>
							<tr>
                              <td colspan="2"><%=f_rel%>&nbsp;</td>
                              <td ><%=f_name%>&nbsp;</td>
                              <td><%=f_birthday%>&nbsp;(<%=f_birthday_id%>)&nbsp;</td>
                              <td colspan="2"><%=f_job%>&nbsp;</td>
                              <td colspan="2"><%=f_tel_ddd%>-<%=f_tel_no1%>-<%=f_tel_no2%>&nbsp;</td>
                              <td colspan="2"><%=f_person1%>-<%=f_person2%>&nbsp;</td>
                              <td><%=f_live%>&nbsp;</td>
                              <td class="right"><%=f_seq%></td>
							</tr>
						<%
							Next
						Else
							Response.Write "<tr><td colspan='12' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If

						rsFamily.Close() : Set rsFamily = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="familyAddPopup();" class="btnType04">�������</a>
					</div>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>