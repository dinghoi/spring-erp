<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
u_type = request("u_type")
card_no = request("card_no")

emp_no     = ""
card_no1   = ""
card_no2   = ""
card_no3   = ""
card_no4   = ""
card_type  = ""
emp_name   = ""
emp_grade  = ""
card_issue = "신규"
card_limit = ""
valid_thru = ""
card_memo  = ""
car_vat_sw = "C"
use_yn     = "Y"

curr_date = mid(now(),1,10)
title_line = "카드 사용자 등록"
if u_type = "U" then

	sql = "select count(*) from card_slip where card_no = '"&card_no&"'"
	Set rs_count=DbConn.Execute(Sql)
	slip_count = cint(rs_count(0)) 'Result.RecordCount

	Sql = "select * from card_owner where card_no = '"&card_no&"'"
	Set rs=DbConn.Execute(Sql)

 	emp_no        = rs("emp_no")
 	emp_name      = rs("emp_name")
	card_no       = rs("card_no")
	owner_company = rs("owner_company")
'	card_no1      = mid(rs("card_no"),1,4)
'	card_no2      = mid(rs("card_no"),6,4)
'	card_no3      = mid(rs("card_no"),11,4)
'	card_no4      = mid(rs("card_no"),16,4)
	card_type     = rs("card_type")
	card_issue    = rs("card_issue")
	card_limit    = rs("card_limit")
	valid_thru    = rs("valid_thru")
	create_date   = rs("create_date")
	start_date    = rs("start_date")
	card_memo     = rs("card_memo")
	car_vat_sw    = rs("car_vat_sw")
    use_yn        = rs("use_yn")
    pl_yn         = rs("pl_yn")
	reg_id        = rs("reg_id")
	reg_date      = mid(rs("reg_date"),1,10)
	reg_name      = rs("reg_name")
	mod_id        = rs("mod_id")
	mod_date      = rs("mod_date")
	mod_name      = rs("mod_name")
    rs.close()

	del_sw = "N"

	title_line = "카드 사용자 내역 수정"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {  $( "#datepicker" ).datepicker();
							$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker" ).datepicker("setDate", "<%=create_date%>" );
			});
			$(function() {  $( "#datepicker1" ).datepicker();
							$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker1" ).datepicker("setDate", "<%=start_date%>" );
			});
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmdel () {
				if (chkdel()) {
					document.frm.del_sw.value = "Y";
					document.frm.submit ();
				}
			}
			function chkdel() {
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

            function chkfrm()
            {
				if(document.frm.card_type.value =="") {
					alert('카드종류 선택하세요');
					frm.card_type.focus();
					return false;}
				if(document.frm.card_no.value =="") {
					alert('카드번호를 입력하세요');
					frm.card_no.focus();
					return false;}
				if(document.frm.owner_company.value =="") {
					alert('소유회사를 선택하세요');
					frm.owner_company.focus();
					return false;}
				if(document.frm.emp_no.value =="") {
					alert('사원조회을 하세요');
					frm.emp_no.focus();
					return false;}
//				if(document.frm.card_limit.value =="") {
//					alert('사용한도를 입력 하세요');
//					frm.card_limit.focus();
//					return false;}
				if(document.frm.valid_thru.value =="") {
					alert('유효기간을 입력 하세요');
					frm.valid_thru.focus();
					return false;}
//				if(document.frm.create_date.value =="") {
//					alert('발급일을 입력 하세요');
//					frm.create_date.focus();
//					return false;}
				if(document.frm.start_date.value =="") {
					alert('개시일을 입력 하세요');
					frm.start_date.focus();
					return false;}
				if(document.frm.create_date.value > document.frm.curr_date.value) {
					alert('발급일이 현재일보다 클수가 없습니다.');
					frm.create_date.focus();
					return false;}
				if(document.frm.start_date.value > document.frm.curr_date.value) {
					alert('개시일이 현재일보다 클수가 없습니다.');
					frm.start_date.focus();
					return false;}
//				if(document.frm.card_memo.value =="") {
//					alert('비고를 입력하세요');
//					frm.card_memo.focus();
//					return false;}

                <%
                if (u_type = "U") then
                    %>
                    return true;
                    <%
                else
                    %>
                    var retVal = false;
                    $.ajax({
                            url: "card_owner_add_ajax.asp"
                            ,async: false
                            ,type: 'post'
                            ,data:  { "card_no" : document.frm.card_no.value }
                            ,dataType: "json"
                            ,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
                            ,beforeSend: function(jqXHR){
                                jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
                            }
                            ,success: function(data)
                            {
                                var result = data.result;

                                if( result=="succ")
                                {
                                    if (data.total_record > 0)
                                    {
                                        alert("이미 등록된 카드번호입니다.");
                                        retVal = false ;
                                    }
                                    else
                                    {
                                        retVal = confirm('입력하시겠습니까?') ;
                                    }
                                }
                                else if(result=="error")
                                {
                                    alert("호출이 실패했습니다.");
                                    retVal = false ;
                                }
                            }
                            ,error: function(jqXHR, status, errorThrown){
                                alert("에러가 발생하였습니다.\n상태코드 : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
                            }
                    });

                    return retVal ;
                    <%
                end if
                %>
			}

		</script>
	</head>
	<body>
		<div id="container">
			<h3 class="tit"><%=title_line%></h3>
			<form action="card_owner_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">카드종류</th>
				        <td class="left">
                        <select name="card_type" id="card_type" style="width:150px">
				          <option value="">선택</option>
                            <%
                            Sql="select * from etc_code where etc_type = '44' order by etc_name asc"
                            rs_etc.Open Sql, Dbconn, 1
                            do until rs_etc.eof
                                %>
                                <option value='<%=rs_etc("etc_name")%>' <%If card_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                <%
                                rs_etc.movenext()
                            loop
                            rs_etc.close()
                            %>
			            </select>
                        </td>
				        <th>카드번호</th>
				        <td class="left">
                            <% if slip_count = 0 then	%>
                                <input name="card_no" type="text" style="width:150px" maxlength="19" onKeyUp="checkNum(this);" value="<%=card_no%>">
                                &nbsp;<strong>반드시 '-' 포함</strong>
                            <% else	%>
                                <%=card_no%>
                                <input type="hidden" name="card_no" value="<%=card_no%>">
                            <% end if	%>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">소유회사</th>
				        <td class="left">
                        <select name="owner_company" id="owner_company" style="width:150px">
                            <option value="">선택</option>
                            <%
                            ' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
                            Sql = "SELECT org_company FROM emp_org_mst "
							Sql = Sql & "WHERE (org_end_date IS NULL OR org_end_date = '0000-00-00') "
							Sql = Sql & "AND org_level = '회사' ORDER BY org_company ASC"

                            rs_etc.Open Sql, Dbconn, 1
                            do until rs_etc.eof
                                %>
                                <option value='<%=rs_etc("org_company")%>' <%If owner_company = rs_etc("org_company") then %>selected<% end if %>><%=rs_etc("org_company")%></option>
                                <%
                                rs_etc.movenext()
                            loop
                            rs_etc.close()
                            %>
			            </select>
                        </td>
				        <th><span class="first">사용자</span></th>
				        <td class="left"><% if (u_type <> "U") or (u_type = "U" and reg_date = curr_date) or (emp_no = "" or isnull(emp_no)) then	%>
                          <input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=emp_name%>">
                          <input name="emp_grade" type="text" id="emp_grade" style="width:60px" value="<%=emp_grade%>">
                          <a href="#" onClick="pop_Window('/member/memb_search.asp','memb_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">사원조회</a>
                          <%   else	%>
                          <%=emp_name%>&nbsp;<%=emp_grade%>
                        <% end if	%></td>
			          </tr>
				      <tr>
				        <th class="first">신규/재발행</th>
				        <td class="left"><label>
                            <input type="radio" name="card_issue" value="신규" <% if card_issue="신규" then %>checked<% end if %> style="width:30px">
                            신규
                            <input type="radio" name="card_issue" value="재발행" <% if card_issue ="재발행" then %>checked<% end if %> style="width:30px">
                            재발행 </label>
                        </td>
				        <th><span class="first">한도</span></th>
				        <td class="left"><input name="card_limit" type="text" id="card_limit" style="width:150px" value="<%=card_limit%>"></td>
			          </tr>
				      <tr>
				        <th class="first">유효기간</th>
				        <td class="left"><input name="valid_thru" type="text" id="valid_thru" maxlength="6" size="6" onKeyUp="checkNum(this);" value="<%=valid_thru%>"></td>
				        <th>발급일</th>
				        <td class="left"><input name="create_date" type="text" value="<%=create_date%>" style="width:80px;text-align:center" id="datepicker"></td>
			          </tr>
				      <tr>
				        <th class="first">사용개시일</th>
				        <td class="left"><input name="start_date" type="text" value="<%=start_date%>" style="width:80px;text-align:center" id="datepicker1"></td>
				        <th>차량유지비</th>
				        <td class="left"><label>
                            <input type="radio" name="car_vat_sw" value="Y" <% if car_vat_sw= "Y" then %>checked<% end if %> style="width:30px">
                            과세
                            <input type="radio" name="car_vat_sw" value="N" <% if car_vat_sw ="N" then %>checked<% end if %> style="width:30px">
                            비과세
                            <input type="radio" name="car_vat_sw" value="C" <% if car_vat_sw ="C" then %>checked<% end if %> style="width:30px">
                            경우에 따라 </label>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">사용유무</th>
				        <td class="left"><label>
                            <input type="radio" name="use_yn" value="Y" <% if use_yn="Y" then %>checked<% end if %> style="width:30px">
                            사용
                            <input type="radio" name="use_yn" value="N" <% if use_yn ="N" then %>checked<% end if %> style="width:30px">
                            미사용 </label>
                        </td>
				        <th>비고</th>
				        <td class="left"><input name="card_memo" type="text" id="card_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=card_memo%>"></td>
			          </tr>
				      <tr>
				        <th class="first">손익유무</th>
				        <td class="left" colspan="3">
                        <label>
                            <input type="radio" name="pl_yn" value="Y" <% if pl_yn="Y" then %>checked<% end if %> style="width:30px">
                            적용
                            <input type="radio" name="pl_yn" value="N" <% if pl_yn ="N" then %>checked<% end if %> style="width:30px">
                            제외
                        </label>
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                    <%	if slip_count = 0 then	%>
                        <span class="btnType01"><input type="button" value="삭제" onclick="javascript:frmdel();" ID="Button1" NAME="Button1"></span>
                    <%	end if	%>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="del_sw" value="<%=del_sw%>" ID="Hidden1">
                <input type="hidden" name="curr_date" value="<%=curr_date%>" ID="Hidden1">
                <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                <input type="hidden" name="org_name" value="<%=org_name%>" ID="Hidden1">
                <input type="hidden" name="reg_date" value="<%=reg_date%>" ID="Hidden1">
                <input type="hidden" name="mod_id" value="<%=mod_id%>" ID="Hidden1">
                <input type="hidden" name="mod_name" value="<%=mod_name%>" ID="Hidden1">
                <input type="hidden" name="mod_date" value="<%=mod_date%>" ID="Hidden1">
			</form>
		</div>
	</body>
</html>

