<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
gubun = Request("gubun")

if gubun = "AS" then
    SQL = "    SELECT * FROM overtime_code           "&chr(13)&_ 
          "     WHERE use_yn = 'Y'                   "&chr(13)&_ 
          "       AND ( apply_dept = 'SM/SI'         "&chr(13)&_ 
          "          OR apply_dept = '��ü'          "&chr(13)&_ 
          "          OR apply_dept = '�ַ�ǻ����'  "&chr(13)&_ 
          "           )                              "&chr(13)&_
          "  ORDER BY overtime_code ASC              "&chr(13)  
 else
    SQL = "  SELECT * FROM overtime_code   "&chr(13)&_
          "   WHERE use_yn = 'Y'           "&chr(13)&_
          "     AND apply_dept = '����52'  "&chr(13)&_
          "ORDER BY overtime_code ASC      "&chr(13)
end if
'Response.write SQL
Rs.open SQL, Dbconn, 1

title_line = "��Ư�� ���� �˻�"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��Ư�� ���� �˻�</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function code_list(work_gubun,holi_id,sign_yn)
			{
				opener.document.frm.work_gubun.value = work_gubun;
				//opener.document.frm.holi_id.value    = holi_id;
				//opener.document.frm.sign_yn.value    = sign_yn;
                opener.fn_SelectedOvertimeCode(work_gubun) ;
                
				window.close();
			}

			function frmcheck () {
				document.frm.submit ();
			}
			
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="overtime_code_search.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="20%" >
							<col width="15%" >
							<col width="20%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�߱ٳ���</th>
								<th scope="col">��������</th>
								<th scope="col">�Ĵ�</th>
								<th scope="col">�ٹ��ð�</th>
								<th scope="col">����ǰ��</th>
							</tr>
						</thead>
						<tbody>
                        <%
                        i = 0
                        do until rs.eof or rs.bof
                            if rs("you_yn") = "Y" then
                                you_view = "����" 
                            else
                                you_view = "����"
                            end if
                            %>
                            <tr>
                                <td class="first">
                                <a href="#" onClick="code_list('<%=rs("work_gubun")%>','<%=rs("holi_id")%>','<%=rs("sign_yn")%>');"><%=rs("work_gubun")%></a>
                                </td>
                                <td><%=rs("holi_id")%></td>
                                <td><%=rs("meals_yn")%></td>
                                <td><%=rs("work_time1")%>&nbsp;</td>
                                <td><%=rs("sign_yn")%></td>
                            </tr>
                            <%
                            i = i + 1
                            rs.movenext()
                        loop
                        rs.close()
                        if i = 0 then
                            %>
                            <tr>
                                <td class="first" colspan="5">������ �����ϴ�</td>
                            </tr>
                            <%
                        end if
                        %>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
			</form>
		</div>        				
	</body>
</html>

