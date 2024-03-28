<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
draft_no = request("draft_no")
draft_man = request("draft_man")
rever_yymm = request("pmg_yymm")
give_date = request("pmg_date")
company = request("company")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_alb = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_alba_mst where draft_no = '"&draft_no&"'"
Set Rs_alb = DbConn.Execute(SQL)
if not Rs_alb.eof then
        company = Rs_alb("company")
		draft_man = Rs_alb("draft_man")
		draft_tax_id = Rs_alb("draft_tax_id")
	    cost_company = Rs_alb("cost_company")
	    bonbu = Rs_alb("bonbu")
	    saupbu = Rs_alb("saupbu")
	    team = Rs_alb("team")
	    org_name = Rs_alb("org_name")
	    sign_no = Rs_alb("sign_no")
    	bank_name = Rs_alb("bank_name")
		account_no = Rs_alb("account_no")
		account_name = Rs_alb("account_name")
		bank_code = Rs_alb("bank_code")
   else
        bank_name = ""
		account_no = ""
		account_name = ""
		bank_code = ""
end if
Rs_alb.close()

title_line = "월 사업소득 상세 내역"

	sql = "select * from pay_alba_cost where rever_yymm = '"&rever_yymm&"' and draft_no = '"&draft_no&"' and company = '"&company&"'"
	set rs = dbconn.execute(sql)

    draft_no = rs("draft_no")
	rever_yymm = rs("rever_yymm")
	draft_no = rs("draft_no")
    company = rs("company")
	give_date = rs("give_date")
	draft_tax_id = rs("draft_tax_id")
	draft_man = rs("draft_man")
	org_name = rs("org_name")
	cost_company = rs("cost_company")
	sign_no = rs("sign_no")
	work_comment = rs("work_comment")
	bank_name = rs("bank_name")
	account_no = rs("account_no")
	account_name = rs("account_name")
	
	alba_cnt = int(rs("alba_cnt"))
	alba_work = int(rs("alba_work"))
	alba_pay = int(rs("alba_pay"))
	alba_trans = int(rs("alba_trans"))
	alba_meals = int(rs("alba_meals"))
	alba_other = int(rs("alba_other"))
	de_other = int(rs("de_other"))
	alba_give_total = int(rs("alba_give_total"))
	curr_pay = int(rs("pay_amount"))
	tax_amt1 = int(rs("tax_amt1"))
	tax_amt2 = int(rs("tax_amt2"))
	tax_hap = tax_amt1 + tax_amt2

	rs.close()

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=ins_last_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=last_check_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=car_year%>" );
			});	  
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
				if(document.frm.draft_no.value =="") {
					alert('등록번호을 입력하세요');
					frm.draft_no.focus();
					return false;}
				{
					return true;
				}
			}

        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_person_view.asp?draft_no=<%=draft_no%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">등록번호</th>
								<td class="left"><%=draft_no%></td>
								<th >성명</th>
								<td class="left" ><%=draft_man%>&nbsp;&nbsp;-&nbsp;&nbsp;<%=draft_tax_id%></td>
							</tr>
                            <tr>
								<th class="first">귀속년월</th>
								<td class="left" ><%=rever_yymm%></td>
                                <th >지급일</th>
								<td class="left"><%=give_date%></td>
							</tr>             
							<tr>
								<th class="first">소속</th>
								<td class="left"><%=company%>&nbsp;&nbsp;<%=org_name%>&nbsp;</td>
								<th>계좌번호</th>
								<td class="left"><%=account_no%>(<%=bank_name%>-<%=account_name%>)&nbsp;</td>
							</tr>
                        	<tr>
								<th class="first">비용회사</th>
								<td class="left"><%=cost_company%>&nbsp;</td>
								<th>전자결재No.</th>
								<td class="left"><%=sign_no%>&nbsp;</td>
							</tr>
							<tr>
								<th class="first" style="background:#F5FFFA">잡급</th>
								<td class="left">
                                <input name="alba_pay" type="text" value="<%=formatnumber(alba_pay,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
								<th style="background:#F8F8FF">소득세</th>
                                <td class="left">
								<input name="tax_amt1" type="text" value="<%=formatnumber(tax_amt1,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">교통비</th>
								<td class="left">
                                <input name="alba_trans" type="text" value="<%=formatnumber(alba_trans,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
								<th style="background:#F8F8FF">지방소득세</th>
                                <td class="left">
								<input name="tax_amt2" type="text" value="<%=formatnumber(tax_amt2,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">식대</th>
								<td class="left">
                                <input name="alba_meals" type="text" value="<%=formatnumber(alba_meals,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
								<th style="background:#F8F8FF">&nbsp;</th>
                                <td class="left">&nbsp;
								<input name="de_other" type="hidden" value="<%=formatnumber(de_other,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">기타지급</th>
								<td class="left">
                                <input name="alba_other" type="text" value="<%=formatnumber(alba_other,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
								<th style="background:#F8F8FF">세액계</th>
                                <td class="left">
								<input name="tax_tot" type="text" value="<%=formatnumber(tax_hap,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
							</tr>
                             <tr>
								<th class="first" style="background:#F5FFFA">지급총액</th>
								<td class="left">
                                <input name="give_tot" type="text" value="<%=formatnumber(alba_give_total,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
								<th style="background:#F8F8FF">차인지급액</th>
                                <td class="left">
								<input name="curr_pay" type="text" value="<%=formatnumber(curr_pay,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F8F8FF">작업일수</th>
								<td class="left">
                                <input name="alba_cnt" type="text" value="<%=formatnumber(alba_cnt,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
								<th style="background:#F8F8FF">작업량</th>
                                <td class="left">
								<input name="alba_work" type="text" value="<%=formatnumber(alba_work,0)%>" style="width:100px;text-align:right" readonly="true">&nbsp;</td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F8F8FF">작업내용</th>
								<td colspan="3" class="left">
                                <input name="work_comment" type="text" value="<%=work_comment%>" style="width:550px" readonly="true"></td>
							</tr>   
                      </tbody>
					</table>
				</div>
                <br>
                <br>
                <div align=center>
                    <span class="btnType01">
                    <a href="#" onClick="pop_Window('insa_pay_albacost_print.asp?draft_no=<%=draft_no%>&draft_man=<%=draft_man%>&rever_yymm=<%=rever_yymm%>&give_date=<%=give_date%>&company=<%=company%>','insa_albacost_report','scrollbars=yes,width=750,height=500')"><input type="button" value="출력" ID="Button1" NAME="Button1"></a>
			        </span>
                    <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="company" value="<%=company%>" ID="Hidden1">
                <input type="hidden" name="org_name" value="<%=org_name%>" ID="Hidden1">
                <input type="hidden" name="bonbu" value="<%=bonbu%>" ID="Hidden1">
                <input type="hidden" name="saupbu" value="<%=saupbu%>" ID="Hidden1">
                <input type="hidden" name="team" value="<%=team%>" ID="Hidden1">
                <input type="hidden" name="cost_company" value="<%=cost_company%>" ID="Hidden1">
                <input type="hidden" name="sign_no" value="<%=sign_no%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

