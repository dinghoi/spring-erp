<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim emp_name
'gubun = request("gubun")
emp_name = Request.Form("emp_name")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if emp_name = "" then
	sql = "select * from memb where (user_name = '"&emp_name&"') and (emp_no < '799999' and emp_no > '100000')"
  else
	sql = "select * from memb where (user_name like '%"&emp_name&"%') and (emp_no < '799999' and emp_no > '100000') ORDER BY user_name ASC"
end if
'Response.write sql
Rs.Open Sql, Dbconn, 1

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
            <form action="memb_search.asp" method="post" name="frm">
                <fieldset class="srch">
                    <legend>조회영역</legend>
                    <dl>
                        <dd>
                            <p>
                            <strong>직원 이름을 입력하세요 </strong>
                                <label>
                                <input name="emp_name" type="text" id="emp_name" value="<%=emp_name%>" style="width:150px;text-align:left; ime-mode:active">
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
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
                        do until rs.eof or rs.bof
                            ii = ii + 1

                            ' 개행문자가 있을 시 제거한다.
                            org_name = rs("org_name")
                            org_name = Replace(org_name,chr(13),"")
                            org_name = Replace(org_name,chr(10),"")
                            %>
                            <tr>
                                <td class="first"><a href="#" onClick="emp_code('<%=rs("user_name")%>','<%=rs("emp_no")%>','<%=rs("user_grade")%>','<%=org_name%>');"><%=rs("user_name")%></a>
                                </td>
                                <td><%=rs("emp_no")%></td>
                                <td><%=rs("user_grade")%></td>
                                <td><%=org_name%></td>
                            </tr>
                            <%
                            rs.movenext()
                        loop
                        rs.close()
                        if ii = 0 then
                            %>
                            <tr>
                                <td class="first" colspan="4">내역이 없습니다</td>
                            </tr>
                            <%
                        end if
                        %>
                        </tbody>
                    </table>
                </div>
            </form>
		</div>
	</body>
</html>
