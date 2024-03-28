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
Dim emp_name
Dim rsMemb, title_line, ii

'gubun = request("gubun")
emp_name = Request.Form("emp_name")

'If emp_name = "" Then
	'sql = "select * from memb where (user_name = '"&emp_name&"') and (emp_no < '799999' and emp_no > '100000')"
'Else
'	sql = "select * from memb where (user_name like '%"&emp_name&"%') and (emp_no < '799999' and emp_no > '100000') ORDER BY user_name ASC"
'End If

objBuilder.Append "SELECT memt.user_name, memt.user_grade, emtt.emp_no, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_reside_company "
objBuilder.Append "FROM memb AS memt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON memt.user_id = emtt.emp_no "
objBuilder.Append "	AND emtt.emp_pay_id <> '2' "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (emtt.emp_no < '799999' AND emtt.emp_no > '100000') "

If emp_name = "" Then
	objBuilder.Append "	AND memt.user_name = '"&emp_name&"' "
Else
	objBuilder.Append "	AND memt.user_name LIKE '%"&emp_name&"%' "
End If

If emp_name <> "" And Not IsNull(emp_name) Then
	objBuilder.Append "ORDER BY memt.user_name ASC "
End If

Set rsMemb = Server.CreateObject("ADODB.RecordSet")
rsMemb.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

title_line = "직원 검색"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>직원 검색</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function emp_code(user_name,emp_no,user_grade,org_name)
			{
				opener.document.frm.emp_name.value = user_name;
				opener.document.frm.emp_no.value = emp_no;
				opener.document.frm.emp_grade.value = user_grade;
				opener.document.frm.org_name.value = org_name;
				window.close();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.emp_name.value =="") {
					alert('직원 이름을 입력하세요');
					frm.emp_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
            <h3 class="tit"><%=title_line%></h3>
            <form action="/member/memb_search.asp" method="post" name="frm">
                <fieldset class="srch">
                    <legend>조회영역</legend>
                    <dl>
                        <dd>
                            <p>
                            <strong>직원 이름을 입력하세요 </strong>
                                <label>
                                <input name="emp_name" type="text" id="emp_name" value="<%=emp_name%>" style="width:150px;text-align:left; ime-mode:active">
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
                        </dd>
                    </dl>
                </fieldset>
                <div class="gView">
                    <table cellpadding="0" cellspacing="0" class="tableList">
                        <colgroup>
                            <col width="15%" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="*" >
                        </colgroup>
                        <thead>
                            <tr>
                                <th class="first" scope="col">이 름</th>
                                <th scope="col">사원번호</th>
                                <th scope="col">직 급</th>
                                <th scope="col">부 서</th>
                            </tr>
                        </thead>
                        <tbody>
                        <%
                        ii = 0

                        Do Until rsMemb.EOF or rsMemb.BOF
                            ii = ii + 1

                            ' 개행문자가 있을 시 제거한다.
                            org_name = rsMemb("org_name")
                            org_name = Replace(org_name, Chr(13), "")
                            org_name = Replace(org_name, Chr(10), "")
                            %>
                            <tr>
                                <td class="first"><a href="#" onClick="emp_code('<%=rsMemb("user_name")%>','<%=rsMemb("emp_no")%>','<%=rsMemb("user_grade")%>','<%=org_name%>');"><%=rsMemb("user_name")%></a>
                                </td>
                                <td><%=rsMemb("emp_no")%></td>
                                <td><%=rsMemb("user_grade")%></td>
                                <td><%'=org_name%>
								<%
									Call EmpOrgInSaupbuText(rsMemb("org_company"), rsMemb("org_bonbu"), rsMemb("org_saupbu"), rsMemb("org_team"))

									If rsMemb("org_reside_company") <> "" Then
										Response.Write "(" & rsMemb("org_reside_company") & ")"
									End If
								%>
								</td>
                            </tr>
                            <%
                            rsMemb.MoveNext()
                        Loop
                        rsMemb.Close() : Set rsMemb = Nothing

                        If ii = 0 Then
                        %>
                            <tr>
                                <td class="first" colspan="4">내역이 없습니다</td>
                            </tr>
                        <%
                        End If
                        %>
                        </tbody>
                    </table>
                </div>
            </form>
		</div>
	</body>
</html>
