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
Dim view_condi, owner_view, title_line
Dim rsLang, lang_empname
Dim rs_emp, emp_name, emp_bonbu, emp_saupbu, emp_team
Dim emp_org_code, emp_org_name, lang_empno, lang_seq, rsEmp
Dim lang_yn

view_condi = f_Request("view_condi")
owner_view = f_Request("owner_view")

If view_condi = "" Then
	owner_view = "T"
End If

title_line = " 어학능력 사항 "

objBuilder.Append "SELECT emlt.lang_empno, emlt.lang_id, emlt.lang_id_type, emlt.lang_point, "
objBuilder.Append "	emlt.lang_grade, emlt.lang_get_date, emlt.lang_seq, "
objBuilder.Append "	emtt.emp_name, emtt.emp_org_code, eomt.org_name "
objBuilder.Append "FROM emp_language AS emlt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emlt.lang_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

If owner_view = "C" Then
	objBuilder.Append "WHERE emtt.emp_name LIKE '%" & view_condi & "%' "
Else
	objBuilder.Append "WHERE emlt.lang_empno = '"&view_condi&"' "
End If
objBuilder.Append "ORDER BY emlt.lang_empno, emlt.lang_seq ASC;"

Set rsLang = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
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
				if(document.frm.view_condi.value == ""){
					alert ("조건을 입력하시기 바랍니다");
					return false;
				}
				return true;
			}

			function language_del(val, val2, val3, val4){
				if (!confirm("정말 삭제하시겠습니까 ?")) return;

				var frm = document.frm;

				document.frm.lang_empno.value = val;
				document.frm.lang_seq.value = val2;
				document.frm.lang_empname.value = val3;
				document.frm.owner_view.value = val4;

				document.frm.action = "/insa/insa_language_del.asp";
				document.frm.submit();
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_language_mg.asp" method="post" name="frm">
					<input type="hidden" name="lang_empno" value="<%=lang_empno%>"/>
					<input type="hidden" name="lang_seq" value="<%=lang_seq%>"/>
					<input type="hidden" name="lang_empname" value="<%=lang_empname%>"/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
                                <label>
									<input type="radio" name="owner_view" value="T" <%If owner_view = "T" Then %>checked<%End If %> style="width:25px;"/>사번
									<input type="radio" name="owner_view" value="C" <%If owner_view = "C" Then %>checked<%End If %> style="width:25px;"/>성명
                                </label>
								<strong>조건 : </strong>
								<label>
        							<input type="text" name="view_condi" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left;"/>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
                            <col width="6%" >
                            <col width="11%" >
                            <col width="*" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>사번</th>
                                <th>성명</th>
                                <th>소속</th>
                                <th>어학구분</th>
                                <th>어학종류</th>
                                <th>점수</th>
                                <th>급수</th>
                                <th>취득일</th>
                                <th>어학</th>
                                <th>수정</th>
                                <th>비고</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsLang.EOF Or rsLang.BOF Then
							lang_yn = "N"	'언어사항 등록 여부
							Response.Write "<tr><td colspan='11' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsLang.EOF
								lang_empno = rsLang("lang_empno")
								emp_name = rsLang("emp_name")
								emp_org_code = rsLang("emp_org_code")
								emp_org_name = rsLang("org_name")
						%>
							<tr>
								<td><%=rsLang("lang_empno")%>&nbsp;</td>
								<td><%=emp_name%>&nbsp;</td>
								<td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
								<td><%=rsLang("lang_id")%>&nbsp;</td>
								<td><%=rsLang("lang_id_type")%>&nbsp;</td>
								<td><%=rsLang("lang_point")%>&nbsp;</td>
								<td><%=rsLang("lang_grade")%>&nbsp;</td>
								<td><%=rsLang("lang_get_date")%>&nbsp;</td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_language_add.asp?lang_empno=<%=rsLang("lang_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')">등록</a>
								</td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_language_add.asp?lang_empno=<%=rsLang("lang_empno")%>&lang_seq=<%=rsLang("lang_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=U','insa_language_add_pop','scrollbars=yes,width=750,height=300')">수정</a>
								</td>
								<%If insa_grade = "0" Then %>
								<td>
									<a href="#" onClick="language_del('<%=rsLang("lang_empno")%>', '<%=rsLang("lang_seq")%>', '<%=emp_name%>', '<%=owner_view%>');return false;">삭제</a>
								</td>
								<%End If%>
							</tr>
						<%
								rsLang.MoveNext()
							Loop
							rsLang.close() : Set rsLang = Nothing
						End If
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
					<%
					If owner_view = "T" And f_toString(view_condi, "") <> "" And lang_yn = "N" Then
						objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no = '"&view_condi&"';"

						Set rsEmp = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsEmp.EOF Then
				    %>
                    <a href="#" onClick="pop_Window('/insa/insa_language_add.asp?lang_empno=<%=view_condi%>&emp_name=<%=rsEmp("emp_name")%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">어학사항 등록</a>
					<%
						End If
						rsEmp.Close() : Set rsEmp = Nothing
					End If
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </div>
                    </td>
			      </tr>
				  </table>

			</form>
		</div>
	</div>
	</body>
</html>