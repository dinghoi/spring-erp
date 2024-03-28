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
Dim view_condi, owner_view, ck_sw, title_line
Dim edu_empno, edu_seq, edu_empname
Dim rsEdu, edu_yn
Dim rs_emp, emp_name, emp_bonbu, emp_saupbu, emp_team
Dim emp_org_code, emp_org_name, deu_empno, task_memo, view_memo
Dim rsEmp

view_condi = f_Request("view_condi")
owner_view = f_Request("owner_view")

title_line = " 교육 사항 "

If view_condi = "" Then
	owner_view = "T"
End If

objBuilder.Append "SELECT emet.edu_empno, emet.edu_comment, emet.edu_name, emet.edu_office, emet.edu_finish_no, "
objBuilder.Append "	emet.edu_start_date, emet.edu_end_date, emet.edu_seq, "
objBuilder.Append "	emtt.emp_name, emtt.emp_org_code, eomt.org_name "
objBuilder.Append "FROM emp_edu AS emet "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emet.edu_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

If owner_view = "C" Then
	objBuilder.Append "WHERE emtt.emp_name LIKE '%"&view_condi&"%' "
Else
	objBuilder.Append "WHERE emet.edu_empno = '"&view_condi&"' "
End If

objBuilder.Append "ORDER BY emet.edu_empno, emet.edu_seq ASC;"

Set rsEdu = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사 관리 시스템</title>
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

			function edu_del(val, val2, val3, val4){

				if (!confirm("정말 삭제하시겠습니까 ?")) return;

				var frm = document.frm;

				document.frm.edu_empno.value = val;
				document.frm.edu_seq.value = val2;
				document.frm.edu_empname.value = val3;
				document.frm.owner_view.value = val4;

				document.frm.action = "/insa/insa_edu_del.asp";
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
				<form action="/insa/insa_edu_mg.asp" method="post" name="frm">
					<input type="hidden" name="edu_empno" value="<%=edu_empno%>"/>
					<input type="hidden" name="edu_seq" value="<%=edu_seq%>"/>
					<input type="hidden" name="edu_empname" value="<%=edu_empname%>"/>
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
                            <col width="15%" >
							<col width="15%" >
							<col width="12%" >
							<col width="12%" >
                            <col width="*" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>사번</th>
                                <th>성명</th>
                                <th>소속</th>
                                <th>교육&nbsp;과정명</th>
                                <th>교육기관</th>
                                <th>교육&nbsp;수료증No.</th>
                                <th>교육&nbsp;기간</th>
                                <th>교육&nbsp;주요&nbsp;내용</th>
                                <th>교육</th>
                                <th>수정</th>
                                <th>비고</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsEdu.EOF Or rsEdu.BOF Then
							edu_yn = "N"	'교육사항 등록 여부
							Response.Write "<tr><td colspan='11' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsEdu.EOF
								edu_empno = rsEdu("edu_empno")
								emp_name = rsEdu("emp_name")
                                emp_org_code = rsEdu("emp_org_code")
                                emp_org_name = rsEdu("org_name")
							    task_memo = Replace(rsEdu("edu_comment"), Chr(34), Chr(39))
								view_memo = task_memo

								If Len(task_memo) > 10 Then
							    	view_memo = Mid(task_memo, 1, 10)&".."
								End If
						%>
							<tr>
								<td><%=rsEdu("edu_empno")%>&nbsp;</td>
								<td><%=emp_name%>&nbsp;</td>
								<td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
								<td><%=rsEdu("edu_name")%>&nbsp;</td>
								<td><%=rsEdu("edu_office")%>&nbsp;</td>
								<td><%=rsEdu("edu_finish_no")%>&nbsp;</td>
								<td><%=rsEdu("edu_start_date")%>∼<%=rsEdu("edu_end_date")%>&nbsp;</td>
								<td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_edu_add.asp?edu_empno=<%=rsEdu("edu_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%=""%>','insa_edu_add_pop','scrollbars=yes,width=750,height=350')">등록</a>
								</td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_edu_add.asp?edu_empno=<%=rsEdu("edu_empno")%>&edu_seq=<%=rsEdu("edu_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%="U"%>','insa_edu_add_pop','scrollbars=yes,width=750,height=350')">수정</a>
								</td>
								<%If insa_grade = "0" Then %>
								<td>
									<a href="#" onClick="edu_del('<%=rsEdu("edu_empno")%>', '<%=rsEdu("edu_seq")%>', '<%=emp_name%>', '<%=owner_view%>');return false;">삭제</a>
								</td>
								<%End If %>
							</tr>
						<%
								rsEdu.MoveNext()
							Loop
							rsEdu.Close() : Set rsEdu = Nothing
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
					If owner_view = "T" And f_toString(view_condi, "") <> "" And edu_yn = "N" Then
						objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no = '"&view_condi&"';"

						Set rsEmp = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsEmp.EOF Then
				    %>
                    <a href="#" onClick="pop_Window('/insa/insa_edu_add.asp?edu_empno=<%=view_condi%>&emp_name=<%=rsEmp("emp_name")%>','교육사항 등록','scrollbars=yes,width=750,height=350')" class="btnType04">교육 등록</a>
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

