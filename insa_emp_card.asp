<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%


curr_date = mid(cstr(now()),1,10)

u_type = request("u_type")
emp_no = request("emp_no")
emp_name = request("emp_name")
be_pg = request("be_pg")
be_pg1 = "insa_emp_card.asp"
page = request("page")

view_sort = request("view_sort")
page_cnt = request("page_cnt")


Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
Set rs = DbConn.Execute(SQL)

photo_image = "/emp_photo/" + rs("emp_image")
emp_email = rs("emp_email") + "@k-won.co.kr"
org_code = rs("emp_org_code")

if rs("emp_org_code") <> "" then
	Sql="select * from emp_org_mst where org_code = '"&org_code&"'"
	Set rs_org=DbConn.Execute(Sql)

	org_fax_no = ""
	org_tel_no = rs_org("org_tel_ddd") + "-" + rs_org("org_tel_no1") + "-" + rs_org("org_tel_no2")
	else
	org_fax_no = ""
	org_tel_no = ""
end if

title_line = " 직원 정보 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
		</script>
		<script type="text/javascript">
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
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_card.asp??emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&u_type=<%=u_type%>" method="post" name="frm">
                <div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="14%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">성명</th>
                                <td class="left"><%=rs("emp_name")%>&nbsp;</td>
                            </tr>
                            <tr>
								<th class="first">직위</th>
                                <td class="left"><%=rs("emp_job")%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th class="first">직책</th>
                                <td class="left"><%=rs("emp_position")%>&nbsp;</td>
                 			</tr>   
                            <tr>
                                <th class="first">소속</th>
                                <td class="left"><%=rs("emp_org_name")%>&nbsp;</td>
                 			</tr>                            
                            <tr>
                                <th class="first">사진</th>
                                <td class="left">
                                <img src="<%=photo_image%>" width=110 height=120 alt="">
                                </td>
                 			</tr>
                            <tr>
                                <th class="first">내선번호</th>
                                <td class="left"><%=rs("emp_extension_no")%>&nbsp;</td>
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
                                <td class="left"><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%>&nbsp;</td>
                 			</tr>  
                            <tr>
                                <th class="first">e메일</th>
                                <td class="left"><%=emp_email%>&nbsp;</td>
                 			</tr>  
			           </tbody>
			        </table>
				  </div>
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>
                    <br>              
        	</form>
      </div>				
	</body>
</html>
