<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
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

objBuilder.Append "CALL USP_PERSON_FAMILY_LIST('"&user_id&"')"

Call Rs_Open(rsFamily, DBConn, objBuilder.ToString())
objBuilder.Clear()

title_line = "가족 사항"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
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

			//가족 등록 팝업[허정호_20210830]
			function personFamilyInit(id, name){
				var url = '/person/insa_family_add.aspp';
				var pop_name = '가족사항 등록';
				var param = '?emp_no='+id;
				var features = 'scrollbars=yes,width=1250,height=650';

				url += param;

				pop_Window(url, pop_name, features);
			}
		</script>

		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>사번 : </strong>
								<label>
									<input type="text" name="emp_no" size="10" value="<%=user_id%>" class="no-input" readonly/>
								</label>
                            <strong>성명 : </strong>
                                <label>
									<input type="text" name="emp_name" size="10" value="<%=user_name%>" class="no-input" readonly/>
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
                            <col width="5%" >
						</colgroup>
						<thead>
                            <tr>
                                <th colspan="2">관계</th>
                                <th>성명</th>
                                <th>생년월일</th>
                                <th colspan="2">직업</th>
                                <th colspan="2">전화번호</th>
                                <th colspan="2">주민번호</th>
                                <th>동거여부</th>
                                <th>No.</th>
                                <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						Dim i, family_empno, family_seq, family_rel, family_name
						Dim family_birthday, family_birthday_id, family_job, family_person1
						Dim family_person2, family_tel_ddd, family_tel_no1, family_tel_no2
						Dim family_live

						If Not rsFamily.EOF Then
							arrFamily = rsFamily.getRows()

							For i = LBound(arrFamily) To UBound(arrFamily, 2)
								family_empno = arrFamily(0, i)
								family_seq = arrFamily(1, i)
								family_rel = arrFamily(2, i)
								family_name = arrFamily(3, i)
								family_birthday  = arrFamily(4, i)
								family_birthday_id = arrFamily(5, i)
								family_job = arrFamily(6, i)
								family_person1 = arrFamily(7, i)
								family_person2 = arrFamily(8, i)
								family_tel_ddd  = arrFamily(9, i)
								family_tel_no1 = arrFamily(10, i)
								family_tel_no2 = arrFamily(11, i)
								family_live = arrFamily(12, i)

								If f_toString(family_person2, "") <> "" Then
									family_person2 = "*******"
								End If
						%>
							<tr>
                              <td colspan="2"><%=family_rel%>&nbsp;</td>
                              <td ><%=family_name%>&nbsp;</td>
                              <td><%=family_birthday%>&nbsp;(<%=family_birthday_id%>)&nbsp;</td>
                              <td colspan="2"><%=family_job%>&nbsp;</td>
                              <td colspan="2"><%=family_tel_ddd%>-<%=family_tel_no1%>-<%=family_tel_no2%>&nbsp;</td>
                              <td colspan="2"><%=family_person1%>-<%=family_person2%>&nbsp;</td>
                              <td><%=family_live%>&nbsp;</td>
                              <td class="right"><%=family_seq%></td>
							  <td>
								<a href="#" onClick="pop_Window('/person/insa_family_add.asp?family_empno=<%=family_empno%>&family_seq=<%=family_seq%>&emp_name=<%=in_name%>&u_type=U','insa_family_add_pop','scrollbars=yes,width=750,height=450')">수정</a>
							  </td>
							</tr>
						<%
							Next
						Else
							Response.Write "<tr><td colspan='13' style='font-weight:bold;height:30px;'>조회된 내역이 없습니다.</td></tr>"
						End If

						Call Rs_Close(rsFamily)
						DBConn.Close : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
						<a href="#" onClick="pop_Window('/person/insa_family_add.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=450')" class="btnType04">가족등록</a>
					</div>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>