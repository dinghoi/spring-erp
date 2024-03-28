<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim page, page_cnt, be_pg, start_page, pgsize, stpage
Dim rsCount, rsMember, total_record, total_page, title_line
Dim rsEmp, rsOrg, pg_url, view_condi

'Dim curr_date, view_sort, orderSql, whereSql, pg_cnt

m_name = f_toString(f_Request("m_name"), "")
page = f_Request("page")
page_cnt = f_Request("page_cnt")
'pg_cnt = CInt(f_Request("pg_cnt"))
'view_sort = f_Request("view_sort")
'view_condi = f_Request("view_condi")

be_pg = "/insa/insa_approve_mg.asp"
'curr_date = DateValue(Mid(CStr(Now()), 1, 10))

'If view_condi = "" Then
'	view_condi = "���̿�"
'End If

' ȭ�� �� ������
pgsize = 10

If page = "" Then
	page = 1
	start_page = 1
End If
stpage = Int((page - 1) * pgsize)

If m_name <> "" Then
	view_condi = "AND m_name LIKE '%"&m_name&"%' "
End If

'pg_url = "&view_condi="&view_condi

objBuilder.Append "SELECT COUNT(*) FROM member_info "
objBuilder.Append "WHERE m_approve_yn = 'N' "&view_condi


Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount
rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT m_seq, m_name, m_ename, m_birthday, m_person1, m_person2, m_sex, "
objBuilder.Append "	m_hp_ddd, m_hp_no1, m_hp_no2, m_emergency_tel, m_last_edu, m_image, m_reg_date "
objBuilder.Append "FROM member_info "
objBuilder.Append "WHERE m_approve_yn = 'N' AND m_del_yn NOT IN ('Y', 'U') "&view_condi
objBuilder.Append "LIMIT "& stpage & "," &pgsize

Set rsMember = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "�ű�ä��"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>�λ���� �ý���</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>
	<script type="text/javascript">
		function getPageCode(){
			return "1 1";
		}

		function frmcheck(){
			if(formcheck(document.frm) && chkfrm()){
				document.frm.submit();
			}
		}

		function chkfrm(){
			if(document.frm.m_name.value == ""){
				alert ("������ �Է��� �ּ���.");
				return false;
			}
			return true;
		}

		//ä�� ���
		function member_del(seq){
			if(!confirm('��� ó���Ͻðڽ��ϱ�?')){
				return false;
			}

			var params = {"m_seq" : seq};

			$.ajax({
 					 url: "/insa/insa_approve_del.asp"
					,type: 'post'
					,data: params
					,dataType: "json"
					,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
					,beforeSend: function(jqXHR){
							jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
						}
					//,success:function(data, status, request){
					,success: function(data){
						var result = data.result;

						if( result=="succ"){
							alert("���������� ��ҵǾ����ϴ�.");
							location.reload();
						}else if( result=="invalid" ){
							alert("������ ��ȣ�� ��Ȯ���� �ʽ��ϴ�.");
						}else if(result=="fail"){
							alert("��� ó���� �����߽��ϴ�.");
						}
					}
					,error: function(jqXHR, status, errorThrown){
						alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
					}
				});
		}
	</script>
</head>
<body>
	<div id="wrap">
		<!--#include virtual = "/include/insa_header.asp" -->
		<!--#include virtual = "/include/insa_sub_menu1.asp" -->
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/insa/insa_approve_mg.asp" method="post" name="frm">

			<fieldset class="srch">
				<legend>��ȸ����</legend>
				<dl>
					<dt>������ �˻���</dt>
					<dd>
						<p>
							<strong>���� : </strong>
							<label>
								<input type="text" name="m_name" id="m_name" value="<%=m_name%>" style="width:100px; text-align:left">
							</label>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"></a>
						</p>
					</dd>
				</dl>
			</fieldset>

			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="5%">
						<col width="10%">
						<col width="5%">
						<col width="5%">
						<col width="*">
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">��  ��</th>
							<th scope="col">�����̸�</th>
							<th scope="col">�������</th>
							<th scope="col">�ֹι�ȣ</th>
							<th scope="col">����</th>
							<th scope="col">�޴��� ��ȣ</th>
							<th scope="col">��󿬶�ó</th>
							<th scope="col">�����з�</th>
							<th scope="col">��������</th>
							<th scope="col">�������</th>
							<th scope="col">����</th>
							<th scope="col">���</th>
						</tr>
					</thead>
				<tbody>
					<%
					If rsMember.EOF Or rsMember.BOF Then
						Response.Write "<tr><td colspan='11' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
					Else
						Do Until rsMember.EOF
					%>
					<tr>
						<td class="first"><%=rsMember("m_name")%></td>
						<td><%=rsMember("m_ename")%>&nbsp;</td>
						<td><%=rsMember("m_birthday")%>&nbsp;</td>
						<td>
							<%=rsMember("m_person1")%>-<%=rsMember("m_person2")%>&nbsp;
						</td>
						<td><%=rsMember("m_sex")%>&nbsp;</td>
						<td>
							<%=rsMember("m_hp_ddd")%>-<%=rsMember("m_hp_no1")%>-<%=rsMember("m_hp_no2")%>&nbsp;
						</td>
						<td><%=rsMember("m_emergency_tel")%>&nbsp;</td>
						<td><%=rsMember("m_last_edu")%>&nbsp;</td>
						<td>
						<%
						If f_toString(rsMember("m_image"), "") = "" Then
							Response.Write "N"
						Else
							Response.Write "Y"
						End If
						%>&nbsp;
						</td>
						<td><%=Mid(rsMember("m_reg_date"), 1, 10)%>&nbsp;</td>
						<%
						If insa_grade = "0" Then
						%>
						<td>
							<a href="#" onClick="pop_Window('/insa/insa_approve_in.asp?m_seq=<%=rsMember("m_seq")%>','ä�� ����','scrollbars=yes,width=1250,height=600')">���</a>
						</td>
						<td>
							<a href="#" onClick="member_del('<%=rsMember("m_seq")%>');">���</a>
						</td>
						<%Else %>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						<%End If %>
					</tr>
					<%
							rsMember.MoveNext()
						Loop
					End If
					rsMember.Close() : Set rsMember = Nothing
					DBConn.Close() : Set DBConn = Nothing
					%>
					</tbody>
				</table>
			</div>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr>
				<td>
				<%
				'Page Navi
				Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
				%>
				</td>
			  </tr>
			  </table>
		</form>
	</div>
</div>
</body>
</html>