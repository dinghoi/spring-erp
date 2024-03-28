<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'on Error resume next
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
Dim in_dong
Dim rs
Dim rs_numRows

gubun = request("gubun")

in_dong = ""
If Request.Form("in_dong")  <> "" Then
  in_dong = Request.Form("in_dong")
End If

title_line = "�� �ּ�(���ڵ�) �˻� ��"

if in_dong = " " then
       in_dong = ""
end if
if in_dong = "" then
       sql = "select * from area_mg where  mg_group = '1' and dong = '" + in_dong + "'"
   else
       Sql = "select * from area_mg where  mg_group = '1' and dong like '%" + in_dong + "%' ORDER BY dong,sido,gugun ASC"
end if

rs.open SQL, DbConn, 1

'Response.write SQL&"<br>"


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�ּ�(���ڵ�) �˻�</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function ziparea(sido,gugun,dong,zip,gubun){
				if(gubun =="family"){
					opener.document.frm.emp_family_sido.value = sido;
				    opener.document.frm.emp_family_gugun.value = gugun;
				    opener.document.frm.emp_family_dong.value = dong;
				    opener.document.frm.emp_family_zip.value = zip;
				    window.close();
				    opener.document.frm.emp_family_addr.focus();
				}

				if(gubun =="org"){
					opener.document.frm.org_sido.value = sido;
				    opener.document.frm.org_gugun.value = gugun;
				    opener.document.frm.org_dong.value = dong;
				    opener.document.frm.org_zip.value = zip;
				    window.close();
				    opener.document.frm.org_addr.focus();
				}

				if(gubun =="juso"){
					opener.document.frm.emp_sido.value = sido;
				    opener.document.frm.emp_gugun.value = gugun;
				    opener.document.frm.emp_dong.value = dong;
				    opener.document.frm.emp_zipcode.value = zip;
				    window.close();
				    opener.document.frm.emp_addr.focus();
				}

				if(gubun =="stay"){
					opener.document.frm.stay_sido.value = sido;
				    opener.document.frm.stay_gugun.value = gugun;
				    opener.document.frm.stay_dong.value = dong;
				    opener.document.frm.stay_zip.value = zip;
				    window.close();
				    opener.document.frm.stay_addr.focus();
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

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.in_dong.value ==""){
					alert('������ �Է��ϼ���');
					frm.in_dong.focus();
					return false;
				}
				{
					return true;
				}
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="zipcode_search.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
                        <dd>
                            <p>
							<strong>������ �Է��ϼ��� </strong>
								<label>
        						<input name="in_dong" type="text" id="in_dong" value="<%=in_dong%>" style="text-align:left; width:150px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="25%" >
							<col width="25%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�����ȣ</th>
								<th scope="col">�õ�</th>
								<th scope="col">����</th>
								<th scope="col">��</th>
							</tr>
						</thead>
						<tbody>
                    	<%
							v_cnt = 0
							do until rs.eof or rs.bof
							     v_cnt = v_cnt + 1
						%>
							<tr>
                                <td class="first"><%=rs("zipcode")%></td>
								<td><%=rs("sido")%></td>
								<td><%=rs("gugun")%></td>
								<td>
                                <a href="#" onClick="ziparea('<%=rs("sido")%>','<%=rs("gugun")%>','<%=rs("dong")%>','<%=rs("zipcode")%>','<%=gubun%>');"><%=rs("dong")%></a>
                                </td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
                        <% If Request.Form("in_dong")  <> "" Then
						      if v_cnt = 0 then %>
							     <td class="first" colspan="4" style=" border-top:1px solid #e3e3e3;">��ȸ ������ �����ϴ�.</td>
                        <%        else %>
								<td class="first" colspan="4" style=" border-top:1px solid #e3e3e3;">(<%=v_cnt%>)&nbsp;���� ��ȸ�Ǿ����ϴ�.</td>
                        <%    end if
						   end if %>

						</tbody>
					</table>
				</div>
				</form>
		</div>
	</body>
</html>

