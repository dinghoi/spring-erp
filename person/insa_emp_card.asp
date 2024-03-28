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
Dim u_type, emp_name, rsEmp, photo_image, emp_email, org_code
Dim org_tel_ddd, org_tel_no1, org_tel_no2, org_fax_no, org_tel_no
Dim title_line, emp_image, emp_org_name, emp_org_code, emp_job
Dim emp_position, emp_extension_no, emp_hp_ddd, emp_hp_no1, emp_hp_no2

u_type = Request.QueryString("u_type")
emp_no = Request.QueryString("emp_no")
emp_name = Request.QueryString("emp_name")

title_line = "직원 정보"

objBuilder.Append "SELECT emp_email, emp_org_name, emp_image, emp_org_code, "
objBuilder.Append "	emp_job, emp_position, emp_extension_no, emp_hp_ddd, "
objBuilder.Append "	emp_hp_no1, emp_hp_no2, "
objBuilder.Append "	eomt.org_tel_ddd, eomt.org_tel_no1, eomt.org_tel_no2 "
objBuilder.Append "FROM emp_master AS emmt "
objBuilder.Append "LEFT OUTER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emp_no = '"&emp_no&"';"

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

emp_email = rsEmp("emp_email")
emp_image = rsEmp("emp_image")
emp_org_name = rsEmp("emp_org_name")
emp_org_code = f_toString(rsEmp("emp_org_code"), "")
emp_job = rsEmp("emp_job")
emp_position = rsEmp("emp_position")
emp_extension_no = rsEmp("emp_extension_no")
emp_hp_ddd = rsEmp("emp_hp_ddd")
emp_hp_no1 = rsEmp("emp_hp_no1")
emp_hp_no2 = rsEmp("emp_hp_no2")
org_tel_ddd = f_toString(rsEmp("org_tel_ddd"), "")
org_tel_no1 = f_toString(rsEmp("org_tel_no1"), "")
org_tel_no2 = f_toString(rsEmp("org_tel_no2"), "")

rsEmp.Close() : Set rsEmp = Nothing
DBConn.Close() : Set DBConn = Nothing

photo_image = "/emp_photo/"&emp_image
emp_email = emp_email&"@k-one.co.kr"

If org_tel_ddd <> "" Then
	org_tel_no = org_tel_ddd&"-"&org_tel_no1&"-"&org_tel_no2
Else
	org_tel_no = ""
End If

org_fax_no = ""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}

			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_emp_card.asp" method="post" name="frm">
                <div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="14%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">성명</th>
                                <td class="left"><%=emp_name%>&nbsp;</td>
                            </tr>
                            <tr>
								<th class="first">직위</th>
                                <td class="left"><%=emp_job%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th class="first">직책</th>
                                <td class="left"><%=emp_position%>&nbsp;</td>
                 			</tr>
                            <tr>
                                <th class="first">소속</th>
                                <td class="left"><%=emp_org_name%>&nbsp;</td>
                 			</tr>
                            <tr>
                                <th class="first">사진</th>
                                <td class="left">
                                <img src="<%=photo_image%>" width="110" height="120" alt="">
                                </td>
                 			</tr>
                            <tr>
                                <th class="first">내선번호</th>
                                <td class="left"><%=emp_extension_no%>&nbsp;</td>
                 			</tr>
                            <tr>
                                <th class="first">팩스번호</th>
                                <td class="left"><%=org_fax_no%>&nbsp;</td>
                 			</tr>
                            <tr>
                                <th class="first">전화번호</th>
                                <td class="left"><%=org_tel_no%>&nbsp;</td>
                 			</tr>
                            <tr>
                                <th class="first">핸드폰</th>
                                <td class="left"><%=emp_hp_ddd%>-<%=emp_hp_no1%>-<%=emp_hp_no2%>&nbsp;</td>
                 			</tr>
                            <tr>
                                <th class="first">e메일</th>
                                <td class="left"><%=emp_email%>&nbsp;</td>
                 			</tr>
			           </tbody>
			        </table>
				  </div>
                   	<br>
               		<div align="right">
						<a href="#" class="btnType04" onclick="javascript:goAction()">닫기</a>&nbsp;&nbsp;
					</div>
                    <br>
        	</form>
      </div>
	</body>
</html>
