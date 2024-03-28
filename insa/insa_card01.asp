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
Dim rsEmp, i, title_line, arrTemp

emp_no = f_Request("emp_no")

title_line = " �λ���ī��-��Ÿ���� "

objBuilder.Append "CALL USP_PERSON_CARD_ETC_INFO('"&emp_no&"')"
Call Rs_Open(rsEmp, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsEmp.EOF Then
	arrTemp = rsEmp.getRows()
End If

Call Rs_Close(rsEmp)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>�λ� ���� �ý���</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<script type="text/javascript">
	/*
		function insert_off(){
			document.getElementById('into_tab').style.display = 'none';
		}
	*/

	//�λ� ���� �˾�[����ȣ_20210819]
	function insaPopView(id, type){
		var url, win_name, features;
		var param = '?emp_no='+id;

		switch(type){
			case 'fam':
				url = '/insa/insa_family_view.asp';
				win_name = '���� ����';
				features = 'scrollbars=yes,width=800,height=400';
				break;
			case 'app':
				url = '/insa/insa_appoint_view.asp';
				win_name = '�߷� ����';
				features = 'scrollbars=yes,width=1200,height=400';
				break;
			case 'edu':
				url = '/insa/insa_edu_view.asp';
				win_name = '���� ����';
				features = 'scrollbars=yes,width=800,height=400';
				break;
			/*default :
				url = '/insa/insa_card01.asp';
				win_name = '�λ��� ��Ÿ����';
				features = 'scrollbars=yes,width=1300,height=750';*/
		}

		url += param;
		pop_Window(url, win_name, features);
	}
	</script>
</head>
<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">-->
<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
<body>
	<div id="wrap">
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<!--<form action="insa_card01.asp" method="post" name="frm">-->
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableWrite">
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
						<col width="9%" >
					</colgroup>
					<tbody>
					<%
					Dim emp_birthday, emp_grade_date, emp_org_baldate, emp_sawo_date, emp_stay_name
					Dim emp_stay_code, emp_type, emp_bonbu, emp_saupbu, emp_team
					Dim emp_reside_place, emp_yuncha_date, emp_sawo_id, emp_disabled
					Dim emp_disab_grade, emp_hobby, emp_birthday_id, emp_family_gugun, emp_family_dong
					Dim emp_family_addr, emp_emergency_tel, org_company, org_bonbu, org_saupbu
					Dim org_team, org_reside_place, emp_family_sido, emp_end_date
					Dim sawo_id, stay_name, stay_sido, stay_gugun, stay_dong
					Dim stay_addr, emp_disabled_yn, emp_org_code

					If IsArray(arrTemp) Then
						emp_stay_name = arrTemp(0, 0)
						emp_stay_code = arrTemp(1, 0)
						emp_type = arrTemp(2, 0)
						emp_company = arrTemp(3, 0)
						emp_bonbu = arrTemp(4, 0)
						emp_saupbu = arrTemp(5, 0)
						emp_team = arrTemp(6, 0)
						emp_reside_place = arrTemp(7, 0)
						emp_yuncha_date = arrTemp(8, 0)
						emp_sawo_id = arrTemp(9, 0)
						emp_disabled = arrTemp(10, 0)
						emp_disab_grade = arrTemp(11, 0)
						emp_hobby = arrTemp(12, 0)
						emp_birthday_id = arrTemp(13, 0)
						emp_family_gugun = arrTemp(14, 0)
						emp_family_dong = arrTemp(15, 0)
						emp_family_addr = arrTemp(16, 0)
						emp_emergency_tel = arrTemp(17, 0)
						org_company = arrTemp(18, 0)
						org_bonbu = arrTemp(19, 0)
						org_saupbu = arrTemp(20, 0)
						org_team = arrTemp(21, 0)
						org_name = arrTemp(22, 0)
						org_reside_place = arrTemp(23, 0)
						emp_family_sido = arrTemp(24, 0)
						emp_end_date = arrTemp(25, 0)
						emp_birthday = arrTemp(26, 0)
						emp_grade_date = arrTemp(27, 0)
						emp_org_baldate = arrTemp(28, 0)
						emp_sawo_date = arrTemp(29, 0)
						stay_name = arrTemp(30, 0)
						stay_sido = arrTemp(31, 0)
						stay_gugun = arrTemp(32, 0)
						stay_dong = arrTemp(33, 0)
						stay_addr = arrTemp(34, 0)
						emp_disabled_yn = arrTemp(35, 0)
						emp_org_code = arrTemp(36, 0)
					End If

					If emp_sawo_id = "Y" Then
						sawo_id = "����"
					Else
						sawo_id = "����"
					End If
					%>
						<tr>
							<th colspan="2" class="first">��������</th>
							<td class="left"><%=emp_type%>&nbsp;</td>
							<th>����</th>
							<td colspan="6" class="left"><%Call EmpOrgCodeSelect(emp_org_code)%>&nbsp;</td>
							<th>���������</th>
							<td class="left"><%=emp_yuncha_date%>&nbsp;</td>
						</tr>
						<tr>
							<th colspan="2" class="first">�������Կ���</th>
							<td class="left"><%=sawo_id%>&nbsp;</td>
							<th>����������</th>
							<td class="left"><%=emp_sawo_date%>&nbsp;</td>
							<th>��ֿ���</th>
							<td colspan="2" class="left">
							<%
							If emp_disabled_yn = "Y" Then
								Response.Write emp_disabled & "-" & emp_disab_grade
							Else
								Response.Write emp_disabled
							End If
							%>
							&nbsp;</td>
							<th>���</th>
							<td class="left"><%=emp_hobby%>&nbsp;</td>
							<th>�������</th>
							<td class="left"><%=emp_birthday%>(<%=emp_birthday_id%>)&nbsp;</td>
						</tr>
						<tr>
							<th colspan="2" class="first">����(�ּ�)</th>
							<td colspan="8" class="left"><%=emp_family_sido%>&nbsp;<%=emp_family_gugun%>&nbsp;<%=emp_family_dong%>&nbsp;<%=emp_family_addr%></td>
							<th>��󿬶���</th>
							<td class="left"><%=emp_emergency_tel%>&nbsp;</td>
						</tr>
						<tr>
							<th colspan="2" class="first">�Ǳٹ�����</th>
							<td colspan="2" class="left"><%=emp_stay_name%>&nbsp;</td>
							<th >�Ǳٹ��� �ּ�</th>
							<td colspan="5" class="left"><%=stay_sido%>&nbsp;<%=stay_gugun%>&nbsp;<%=stay_dong%>&nbsp;<%=stay_addr%>&nbsp;</td>
							<th >������</th>
							<td class="left"><%=emp_end_date%>&nbsp;</td>
						</tr>
						<tr>
							<th colspan="10" class="left">�� ���� ���� ��</th>
							<td colspan="2" class="right">&nbsp;
							<a href="#" class="btnType03" onClick="insaPopView('<%=emp_no%>','fam');">�� ���� ������</a>
						</tr>
						<tr>
							<th colspan="3">����</th>
							<th colspan="2">����</th>
							<th colspan="2">�������</th>
							<th colspan="2">����</th>
							<th colspan="2">�ֹι�ȣ</th>
							<th>���ſ���</th>
						</tr>
						<%
						Dim rsFamily, arrFamily
						Dim family_rel, family_name, family_birthday, family_birthday_id, family_job
						Dim family_person1, family_person2, family_live

						'objBuilder.Append "CALL USP_PERSON_INSA_CARD_FAMILY_SEL('"&emp_no&"')"
						objBuilder.Append "CALL USP_PERSON_CARD_FAMILY_INFO('"&emp_no&"')"

						Call Rs_Open(rsFamily, DBConn, objBuilder.ToString())
						objBuilder.Clear()

						If Not rsFamily.EOF Then
							arrFamily = rsFamily.getRows()
						End If

						Call Rs_Close(rsFamily)

						If IsArray(arrFamily) Then
							For i = LBound(arrFamily) To UBound(arrFamily, 2)
								family_rel = arrFamily(0, i)
								family_name = arrFamily(1, i)
								family_birthday = arrFamily(2, i)
								family_birthday_id = arrFamily(3, i)
								family_job = arrFamily(4, i)
								family_person1 = arrFamily(5, i)
								family_person2 = arrFamily(6, i)
								family_live = arrFamily(7, i)
						%>
						<tr>
							<td colspan="3" class="left"><%=family_rel%>&nbsp;</td>
							<td colspan="2" class="left"><%=family_name%>&nbsp;</td>
							<td colspan="2" class="left"><%=family_birthday%>(<%=family_birthday_id%>)&nbsp;</td>
							<td colspan="2" class="left"><%=family_job%>&nbsp;</td>
							<td colspan="2" class="left"><%=family_person1%>-<%=family_person2%>&nbsp;</td>
							<td class="left"><%=family_live%>&nbsp;</td>
						</tr>
						<%
							Next
						Else
						%>
						<tr>
							<td colspan="12" style="height:30px;">��ȸ�� ������ �����ϴ�.</td>
						</tr>
						<%
						End If
						%>
						<tr>
							<th colspan="10" class="left">�� �߷� ���� ��</th>
							<td colspan="2" class="right">&nbsp;
							<a href="#" class="btnType03" onClick="insaPopView('<%=emp_no%>','app');">�� �߷� ������</a>
							</td>
						</tr>
						<tr>
							<th rowspan="2" colspan="2" class="first">�߷���</th>
							<th rowspan="2" scope="col">�߷ɱ���</th>
							<th rowspan="2" scope="col">�߷�����</th>
							<th colspan="3" scope="col">�߷���</th>
							<th colspan="5" scope="col">�߷���</th>
						</tr>
						<tr>
							<th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">ȸ��</th>
							<th scope="col">�Ҽ�</th>
							<th scope="col">����/å</th>
							<th scope="col">ȸ��</th>
							<th scope="col">�Ҽ�</th>
							<th scope="col">����/å</th>
							<th colspan="2" scope="col">���</th>
						</tr>
						<%
						Dim rsAppoint, arrAppoint
						Dim app_date, app_id, app_id_type, app_to_company, app_to_orgcode
						Dim app_to_org, app_to_grade, app_to_job, app_to_position, app_to_enddate
						Dim app_be_company, app_be_orgcode, app_be_org, app_be_grade, app_be_job
						Dim app_be_position, app_be_enddate, app_start_date, app_finish_date, app_reward
						Dim app_comment

						objBuilder.Append "CALL USP_PERSON_CARD_APPOINT_INFO('"&emp_no&"')"

						Call Rs_Open(rsAppoint, DBConn, objBuilder.ToString())
						objBuilder.Clear()

						If Not rsAppoint.EOF Then
							arrAppoint = rsAppoint.getRows()
						End If

						Call Rs_Close(rsAppoint)

						If IsArray(arrAppoint) Then
							For i = LBound(arrAppoint) To UBound(arrAppoint, 2)
								app_date = arrAppoint(0, i)
								app_id = arrAppoint(1, i)
								app_id_type = arrAppoint(2, i)
								app_to_company = arrAppoint(3, i)
								app_to_orgcode = arrAppoint(4, i)
								app_to_org = arrAppoint(5, i)
								app_to_grade = arrAppoint(6, i)
								'app_to_job = arrAppoint(7, i)
								app_to_position = arrAppoint(8, i)
								app_to_enddate = arrAppoint(9, i)
								app_be_company = arrAppoint(10, i)
								app_be_orgcode = arrAppoint(11, i)
								app_be_org = arrAppoint(12, i)
								app_be_grade = arrAppoint(13, i)
								'app_be_job = arrAppoint(14, i)
								app_be_position = arrAppoint(15, i)
								app_be_enddate = arrAppoint(16, i)
								app_start_date = arrAppoint(17, i)
								app_finish_date = arrAppoint(18, i)
								app_reward = arrAppoint(19, i)
								app_comment = arrAppoint(20, i)
						%>
						<tr>
							<td colspan="2" class="left"><%=app_date%>&nbsp;</td>
							<td class="left"><%=app_id%>&nbsp;</td>
							<td class="left"><%=app_id_type%>&nbsp;</td>
							<td class="left"><%=app_to_company%>&nbsp;</td>
							<td class="left"><%=app_to_orgcode%>)<%=app_to_org%>&nbsp;</td>
							<td class="left"><%=app_to_grade%>-<%=app_to_position%>&nbsp;</td>
							<td class="left"><%=app_be_company%>&nbsp;</td>
							<td class="left"><%=app_be_orgcode%>)<%=app_be_org%>&nbsp;</td>
							<td class="left"><%=app_be_grade%>-<%=app_be_position%>&nbsp;</td>
							<td colspan="2" class="left"><%=app_start_date%>-<%=app_finish_date%><%=app_be_enddate%>&nbsp;<%=app_reward%>&nbsp;<%=app_comment%>&nbsp;</td>
						</tr>
						<%
							Next
						Else
						%>
						<tr>
							<td colspan="12" style="height:30px;">��ȸ�� ������ �����ϴ�.</td>
						</tr>
						<%
						End If
						%>
						 <tr>
							<th colspan="10" class="left">�� ���� ���� ��</th>
							<td colspan="2" class="right">&nbsp;
							<a href="#" class="btnType03" onClick="insaPopView('<%=emp_no%>','edu');">�� ���� ������</a>
							</td>
						 </tr>
						<tr>
							<th colspan="3">����&nbsp;������</th>
							<th colspan="2">�������</th>
							<th>����&nbsp;������No.</th>
							<th colspan="2">����&nbsp;�Ⱓ</th>
							<th colspan="4">����&nbsp;�ֿ�&nbsp;����</th>
						</tr>
						<%
						Dim rsEdu, arrEdu
						Dim edu_name, edu_office, edu_finish_no, edu_start_date, edu_end_date
						Dim edu_comment

						objBuilder.Append "CALL USP_PERSON_CARD_EDU_INFO('"&emp_no&"')"

						Call Rs_Open(rsEdu, DBConn, objBuilder.ToString())
						objBuilder.Clear()

						If Not rsEdu.EOF Then
							arrEdu = rsEdu.getRows()
						End If

						Call Rs_Close(rsEdu)

						If IsArray(arrEdu) Then
							For i = LBound(arrEdu) To UBound(arrEdu, 2)
								edu_name = arrEdu(0, i)
								edu_office = arrEdu(1, i)
								edu_finish_no = arrEdu(2, i)
								edu_start_date = arrEdu(3, i)
								edu_end_date = arrEdu(4, i)
								edu_comment = arrEdu(5, i)
						%>
						<tr>
							<td colspan="3" class="left"><%=edu_name%>&nbsp;</td>
							<td colspan="2"class="left"><%=edu_office%>&nbsp;</td>
							<td class="left"><%=edu_finish_no%>&nbsp;</td>
							<td colspan="2" class="left"><%=edu_start_date%> - <%=edu_end_date%>&nbsp;</td>
							<td colspan="4" class="left"><%=edu_comment%>&nbsp;</td>
						</tr>
						<%
							Next
						Else
						%>
						<tr>
							<td colspan="12" style="height:30px;">��ȸ�� ������ �����ϴ�.</td>
						</tr>
						<%
						End If
						%>
						<tr>
							<th colspan="12" class="left">�� ���� �ɷ� ��</th>
						</tr>
						<tr>
							<th colspan="3">���б���</th>
							<th colspan="2">��������</th>
							<th colspan="2">����</th>
							<th colspan="2">�޼�</th>
							<th colspan="3">�����</th>
						</tr>
						<%
						Dim rsLang, arrLang
						Dim lang_id, lang_id_type, lang_point, lang_grade, lang_get_date

						objBuilder.Append "CALL USP_PERSON_LANGUAGE_INFO('"&emp_no&"')"
						Call Rs_Open(rsLang, DBConn, objBuilder.ToString())
						objBuilder.Clear()

						If Not rsLang.EOF Then
							arrLang = rsLang.getRows()
						End If

						If IsArray(arrLang) Then
							For i = LBound(arrLang) To UBound(arrLang, 2)
								lang_id = arrLang(0, i)
								lang_id_type = arrLang(1, i)
								lang_point = arrLang(2, i)
								lang_grade = arrLang(3, i)
								lang_get_date = arrLang(4, i)
						%>
						<tr>
							<td colspan="3" class="left"><%=lang_id%>&nbsp;</td>
							<td colspan="2" class="left"><%=lang_id_type%>&nbsp;</td>
							<td colspan="2" class="left"><%=lang_point%>&nbsp;</td>
							<td colspan="2" class="left"><%=lang_grade%>&nbsp;</td>
							<td colspan="3" class="left"><%=lang_get_date%>&nbsp;</td>
						</tr>
						<%
							Next
						Else
						%>
						<tr>
							<td colspan="12" style="height:30px;">��ȸ�� ������ �����ϴ�.</td>
						</tr>
						<%
						End If

						DBConn.Close() : Set DBConn = Nothing
						%>
				  </tbody>
				</table>
				<br>
				<div align="right">
					<a href="#" class="btnType04" onclick="close_win();" >�ݱ�</a>&nbsp;&nbsp;
				</div>
				<br>
		<!--</form>-->
	</div>
</div>
</body>
</html>