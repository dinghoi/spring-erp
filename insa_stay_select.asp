<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

gubun = request("gubun")
reside_code = request("reside_code")
if gubun = "" then
   gubun = Request.Form("gubun")
end if

in_name = ""
If Request.Form("in_name")  <> "" Then 
  in_name = Request.Form("in_name") 
End If

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if in_name = "" then
	first_view = "N"
	sql = "select * from emp_stay where stay_name = '" + in_name + "'"
  else
	first_view = "Y"
	Sql = "select * from emp_stay where stay_name like '%" + in_name + "%' ORDER BY stay_name ASC"
end if

rs.open SQL, DbConn, 1


title_line = "�� �Ǳٹ��� �˻� ��"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�Ǳٹ��� �˻�</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
	        function staysel(stay_code,stay_name,stay_sido,stay_gugun,stay_dong,stay_addr,gubun)
			{
				if(gubun =="stay")
					{ 
					opener.document.frm.emp_stay_code.value = stay_code;
					opener.document.frm.emp_stay_name.value = stay_name;
					opener.document.frm.stay_sido.value = stay_sido;
					opener.document.frm.stay_gugun.value = stay_gugun;
					opener.document.frm.stay_dong.value = stay_dong;
					opener.document.frm.stay_addr.value = stay_addr;
					window.close();
					opener.document.frm.stay_addr.focus();
					}
				if(gubun =="juso")
					{ 
					opener.document.frm.emp_sido.value = sido;
				    opener.document.frm.emp_gugun.value = gugun;
				    opener.document.frm.emp_dong.value = dong;
				    opener.document.frm.emp_zip.value = zip;
				    window.close();
				    opener.document.frm.emp_addr.focus();
					}					
				<%	
				'else
				'	{ 
				'	opener.document.frm.sido.value = sido;
				'   opener.document.frm.family_gugun.value = gugun;
				'   opener.document.frm.family_dong.value = dong;
				'   opener.document.frm.family_zip.value = zip;
				'    window.close();
				'    opener.document.frm.family_addr.focus();
				'	}
				%>
			}			
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.in_name.value =="") {
					alert('�ٹ������� �Է��ϼ���');
					frm.in_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_stay_select.asp?gubun=<%=gubun%>&reside_code=<%=reside_code%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>�ٹ������� �Է��ϼ��� </strong>
								<label>
        						<input name="in_name" type="text" id="in_name" value="<%=in_name%>" style="text-align:left; width:150px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�ڵ�</th>
                                <th scope="col">�ٹ�����</th>
                                <th scope="col">��ȭ��ȣ</th>
                                <th scope="col">����óȸ��</th>
                                <th scope="col">����ó��</th>
								<th scope="col">�ּ�</th>
 							</tr>
						</thead>
						<tbody>
                    	<%
						if first_view = "Y" then
							do until rs.eof or rs.bof
						%>
							<tr>
								<td class="first"><%=rs("stay_code")%></td>
                                <td>
                                <a href="#" onClick="staysel('<%=rs("stay_code")%>','<%=rs("stay_name")%>','<%=rs("stay_sido")%>','<%=rs("stay_gugun")%>','<%=rs("stay_dong")%>','<%=rs("stay_addr")%>','<%=gubun%>');"><%=rs("stay_name")%></a>
                                </td>
                                <td><%=rs("stay_tel_ddd")%>-<%=rs("stay_tel_no1")%>-<%=rs("stay_tel_no2")%></td>
                                <td><%=rs("stay_reside_company")%></td>
                                <td><%=rs("stay_org_name")%></td>
								<td><%=rs("stay_sido")%> - <%=rs("stay_gugun")%> - <%=rs("stay_dong")%> - <%=rs("stay_addr")%></td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
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

