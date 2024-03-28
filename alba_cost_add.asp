<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")

draft_no = request("draft_no")
draft_man = request("draft_man")
rever_yymm = request("rever_yymm")
give_date = request("give_date")
'cost_company = request("cost_company")
'bonbu = request("bonbu")
'saupbu = request("saupbu")
'team = request("team")
'org_name = request("org_name")
'sign_no = request("sign_no")

company = emp_company
work_comment = ""
cost_company = ""

alba_cnt = 0
alba_work = 0
alba_pay = 0
alba_trans = 0
alba_meals = 0
alba_other = 0
alba_other2 = 0
de_other = 0
tax_amt1 = 0
tax_amt2 = 0
alba_give_total = 0
curr_pay = 0

to_date = mid(cstr(now()),1,10)
'	give_date = to_date '지급일

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_alb = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "아르바이트 비용 입력"

if u_type = "U" then

	sql = "select * from pay_alba_cost where rever_yymm = '"&rever_yymm&"' and draft_no = '"&draft_no&"' and give_date = '"&give_date&"'"
	set rs = dbconn.execute(sql)

    rever_yymm = rs("rever_yymm")
	draft_no = rs("draft_no")
    company = rs("company")
	give_date = rs("give_date")
	draft_tax_id = rs("draft_tax_id")
	draft_man = rs("draft_man")

	bonbu = rs("bonbu")
	saupbu = rs("saupbu")
	team = rs("team")
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

	title_line = "아르바이트 비용 수정"
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=give_date%>" );
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
				if(document.frm.draft_no.value =="" ) {
					alert('등록번호를 입력하세요');
					frm.draft_no.focus();
					return false;}
				if(document.frm.give_date.value =="" ) {
					alert('지급일을 입력하세요');
					frm.give_date.focus();
					return false;}
				if(document.frm.cost_company.value =="" ) {
					alert('비용사용 고객사를 선택하세요');
					frm.cost_company.focus();
					return false;}
//				if(document.frm.sign_no.value =="") {
//					alert('전자결재번호를 입력하세요');
//					frm.sign_no.focus();
//					return false;}
				if(document.frm.alba_pay.value =="" || document.frm.alba_pay.value ==0) {
					alert('잡급 금액을 입력하세요');
					frm.alba_cost.focus();
					return false;}
				if(document.frm.alba_cnt.value =="" || document.frm.alba_cnt.value ==0) {
					alert('작업일수를 입력하세요');
					frm.alba_cnt.focus();
					return false;}
//				if(document.frm.alba_work.value =="" || document.frm.alba_work.value ==0) {
//					alert('작업량을 입력하세요');
//					frm.alba_work.focus();
//					return false;}
				if(document.frm.work_comment.value =="" ) {
					alert('작업내용을 입력하세요');
					frm.work_comment.focus();
					return false;}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function give_cal(txtObj){
				al_pay = parseInt(document.frm.alba_pay.value.replace(/,/g,""));
				al_trans = parseInt(document.frm.alba_trans.value.replace(/,/g,""));
				al_meals = parseInt(document.frm.alba_meals.value.replace(/,/g,""));
				al_other = parseInt(document.frm.alba_other.value.replace(/,/g,""));

//				give_hap = al_pay + al_trans + al_meals + al_other;
				give_hap = al_pay + al_trans + al_meals + al_other;

				tax_1 = al_pay * (3 / 100)
				tax_1 = parseInt(tax_1);
				tax_1 = (parseInt(tax_1 / 10)) * 10;

				tax_2 = al_pay * (0.3 / 100)
				tax_2 = parseInt(tax_2);
				tax_2 = (parseInt(tax_2 / 10)) * 10;

				tax_hap = tax_1 + tax_2;

			    curr_amt = give_hap - tax_hap;

				give_hap = String(give_hap);
				num_len = give_hap.length;
				sil_len = num_len;
				give_hap = String(give_hap);
				if (give_hap.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) give_hap = give_hap.substr(0,num_len -3) + "," + give_hap.substr(num_len -3,3);
				if (sil_len > 6) give_hap = give_hap.substr(0,num_len -6) + "," + give_hap.substr(num_len -6,3) + "," + give_hap.substr(num_len -2,3);
				document.frm.give_tot.value = give_hap;

				tax_1 = String(tax_1);
				num_len = tax_1.length;
				sil_len = num_len;
				tax_1 = String(tax_1);
				if (tax_1.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_1 = tax_1.substr(0,num_len -3) + "," + tax_1.substr(num_len -3,3);
				if (sil_len > 6) tax_1 = tax_1.substr(0,num_len -6) + "," + tax_1.substr(num_len -6,3) + "," + tax_1.substr(num_len -2,3);
				document.frm.tax_amt1.value = tax_1;

				tax_2 = String(tax_2);
				num_len = tax_2.length;
				sil_len = num_len;
				tax_2 = String(tax_2);
				if (tax_2.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_2 = tax_2.substr(0,num_len -3) + "," + tax_2.substr(num_len -3,3);
				if (sil_len > 6) tax_2 = tax_2.substr(0,num_len -6) + "," + tax_2.substr(num_len -6,3) + "," + tax_2.substr(num_len -2,3);
				document.frm.tax_amt2.value = tax_2;

				tax_hap = String(tax_hap);
				num_len = tax_hap.length;
				sil_len = num_len;
				tax_hap = String(tax_hap);
				if (tax_hap.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_hap = tax_hap.substr(0,num_len -3) + "," + tax_hap.substr(num_len -3,3);
				if (sil_len > 6) tax_hap = tax_hap.substr(0,num_len -6) + "," + tax_hap.substr(num_len -6,3) + "," + tax_hap.substr(num_len -2,3);
				document.frm.tax_tot.value = tax_hap;

                curr_amt = String(curr_amt);
				num_len = curr_amt.length;
				sil_len = num_len;
				curr_amt = String(curr_amt);
				if (curr_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) curr_amt = curr_amt.substr(0,num_len -3) + "," + curr_amt.substr(num_len -3,3);
				if (sil_len > 6) curr_amt = curr_amt.substr(0,num_len -6) + "," + curr_amt.substr(num_len -6,3) + "," + curr_amt.substr(num_len -2,3);
				document.frm.curr_pay.value = curr_amt;

				if (txtObj.value.length >= 2) {
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}
				if (txtObj.value.length<5) {
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
				}
				var num = txtObj.value;
				if (num == "--" ||  num == "." ) num = "";
				if (num != "" ) {
					temp=new String(num);
					if(temp.length<1) return "";

					// 음수처리
					if(temp.substr(0,1)=="-") minus="-";
					else minus="";

					// 소수점이하처리
					dpoint=temp.search(/\./);

					if(dpoint>0)
					{
					// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
					dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
					temp=temp.substr(0,dpoint);
					}else dpointVa="";

					// 숫자이외문자 삭제
					temp=temp.replace(/\D/g,"");
					zero=temp.search(/[1-9]/);

					if(zero==-1) return "";
					else if(zero!=0) temp=temp.substr(zero);

					if(temp.length<4) return minus+temp+dpointVa;
					buf="";
					while (true)
					{
					if(temp.length<3) { buf=temp+buf; break; }

					buf=","+temp.substr(temp.length-3)+buf;
					temp=temp.substr(0, temp.length-3);
					}
					if(buf.substr(0,1)==",") buf=buf.substr(1);

					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;
				}else txtObj.value = "0";
			}

			function taxx_cal(txtObj){
				al_pay = parseInt(document.frm.alba_pay.value.replace(/,/g,""));
				al_trans = parseInt(document.frm.alba_trans.value.replace(/,/g,""));
				al_meals = parseInt(document.frm.alba_meals.value.replace(/,/g,""));
				al_other = parseInt(document.frm.alba_other.value.replace(/,/g,""));

				tax_1 = parseInt(document.frm.tax_amt1.value.replace(/,/g,""));
				tax_2 = parseInt(document.frm.tax_amt2.value.replace(/,/g,""));

//				give_hap = al_pay + al_trans + al_meals + al_other;
				give_hap = al_pay + al_trans + al_meals + al_other;


				tax_hap = tax_1 + tax_2;

			    curr_amt = give_hap - tax_hap;

				give_hap = String(give_hap);
				num_len = give_hap.length;
				sil_len = num_len;
				give_hap = String(give_hap);
				if (give_hap.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) give_hap = give_hap.substr(0,num_len -3) + "," + give_hap.substr(num_len -3,3);
				if (sil_len > 6) give_hap = give_hap.substr(0,num_len -6) + "," + give_hap.substr(num_len -6,3) + "," + give_hap.substr(num_len -2,3);
				document.frm.give_tot.value = give_hap;

				tax_1 = String(tax_1);
				num_len = tax_1.length;
				sil_len = num_len;
				tax_1 = String(tax_1);
				if (tax_1.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_1 = tax_1.substr(0,num_len -3) + "," + tax_1.substr(num_len -3,3);
				if (sil_len > 6) tax_1 = tax_1.substr(0,num_len -6) + "," + tax_1.substr(num_len -6,3) + "," + tax_1.substr(num_len -2,3);
				document.frm.tax_amt1.value = tax_1;

				tax_2 = String(tax_2);
				num_len = tax_2.length;
				sil_len = num_len;
				tax_2 = String(tax_2);
				if (tax_2.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_2 = tax_2.substr(0,num_len -3) + "," + tax_2.substr(num_len -3,3);
				if (sil_len > 6) tax_2 = tax_2.substr(0,num_len -6) + "," + tax_2.substr(num_len -6,3) + "," + tax_2.substr(num_len -2,3);
				document.frm.tax_amt2.value = tax_2;

				tax_hap = String(tax_hap);
				num_len = tax_hap.length;
				sil_len = num_len;
				tax_hap = String(tax_hap);
				if (tax_hap.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) tax_hap = tax_hap.substr(0,num_len -3) + "," + tax_hap.substr(num_len -3,3);
				if (sil_len > 6) tax_hap = tax_hap.substr(0,num_len -6) + "," + tax_hap.substr(num_len -6,3) + "," + tax_hap.substr(num_len -2,3);
				document.frm.tax_tot.value = tax_hap;

                curr_amt = String(curr_amt);
				num_len = curr_amt.length;
				sil_len = num_len;
				curr_amt = String(curr_amt);
				if (curr_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) curr_amt = curr_amt.substr(0,num_len -3) + "," + curr_amt.substr(num_len -3,3);
				if (sil_len > 6) curr_amt = curr_amt.substr(0,num_len -6) + "," + curr_amt.substr(num_len -6,3) + "," + curr_amt.substr(num_len -2,3);
				document.frm.curr_pay.value = curr_amt;

				if (txtObj.value.length >= 2) {
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}
				if (txtObj.value.length<5) {
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
				}
				var num = txtObj.value;
				if (num == "--" ||  num == "." ) num = "";
				if (num != "" ) {
					temp=new String(num);
					if(temp.length<1) return "";

					// 음수처리
					if(temp.substr(0,1)=="-") minus="-";
					else minus="";

					// 소수점이하처리
					dpoint=temp.search(/\./);

					if(dpoint>0)
					{
					// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
					dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
					temp=temp.substr(0,dpoint);
					}else dpointVa="";

					// 숫자이외문자 삭제
					temp=temp.replace(/\D/g,"");
					zero=temp.search(/[1-9]/);

					if(zero==-1) return "";
					else if(zero!=0) temp=temp.substr(zero);

					if(temp.length<4) return minus+temp+dpointVa;
					buf="";
					while (true)
					{
					if(temp.length<3) { buf=temp+buf; break; }

					buf=","+temp.substr(temp.length-3)+buf;
					temp=temp.substr(0, temp.length-3);
					}
					if(buf.substr(0,1)==",") buf=buf.substr(1);

					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;
				}else txtObj.value = "0";
			}

			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U')
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}

			function delcheck()
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.method = "post";
					document.frm.action = "alba_cost_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="alba_cost_add_save.asp" method="post" name="frm">
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
								<td class="left"><input name="draft_no" type="text" id="draft_no" style="width:90px" readonly="true" value="<%=draft_no%>">
								<a href="#" class="btnType03" onClick="pop_Window('alba_person_search.asp?gubun=<%="1"%>','alba_search','scrollbars=yes,width=600,height=400')">찾기</a>
            					</td>
								<th >성명</th>
								<td class="left" >
                                <input name="draft_man" type="text" value="<%=draft_man%>" style="width:90px" readonly="true">
                                -
                                <input name="draft_tax_id" type="text" value="<%=draft_tax_id%>" style="width:90px" readonly="true">
                                </td>
							</tr>
                            <tr>
								<th class="first">귀속년월</th>
								<td class="left" ><input name="rever_yymm" type="text" value="<%=rever_yymm%>" style="width:70px" readonly="true"></td>
                                <th >지급일</th>
								<td class="left"><input name="give_date" type="text" value="<%=give_date%>" style="width:80px;text-align:center" id="datepicker" readonly="true"></td>
							</tr>
							<tr>
								<th class="first">사용조직</th>
								<td class="left" >
							<%' if cost_grade = "0" or saupbu = "경영지원실" then	%>
                                  <input name="org_name" type="text" value="<%=org_name%>" readonly="true" style="width:150px">
                                  <a href="#" onClick="pop_Window('/insa/org_search.asp','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조직조회</a>
                            <% '  else	%>
                                <%'=org_name%>
                                <!--<input name="org_name" type="hidden" value="<%'=org_name%>">-->
                            <%' end if	%>
                                  <input name="emp_company" type="hidden" value="<%=company%>">
                                  <input name="bonbu" type="hidden" value="<%=bonbu%>">
                                  <input name="saupbu" type="hidden" value="<%=saupbu%>">
                                  <input name="team" type="hidden" value="<%=team%>">
                                  <input name="reside_place" type="hidden" value="<%=reside_place%>">
                                  <input name="reside_company" type="hidden" value="<%=reside_company%>">
                                </td>
								<th>계좌번호</th>
								<td class="left">
                                <input name="bank_name" type="text" value="<%=bank_name%>" style="width:60px" readonly="true">
                                -
                                <input name="account_no" type="text" value="<%=account_no%>" style="width:90px" readonly="true">
                                -
                                <input name="account_name" type="text" value="<%=account_name%>" style="width:50px" readonly="true">
                                </td>
							</tr>
                        	<tr>
								<th class="first">고객사</th>
								<td class="left"><input name="cost_company" type="text" id="cost_company" style="width:150px" readonly="true" value="<%=cost_company%>">
								<a href="#" class="btnType03" onClick="pop_Window('insa_trade_search.asp?gubun=<%="2"%>','tradesearch','scrollbars=yes,width=600,height=400')">조회</a>
                                </td>
								<th>전자결재No.</th>
								<td class="left"><input name="sign_no" type="text" id="sign_no" style="width:40px" onKeyUp="checkNum(this);" value="<%=sign_no%>" maxlength="4">
*숫자4자리만 입력 가능</td>
							</tr>
							<tr>
								<th class="first" style="background:#F5FFFA">잡급</th>
								<td class="left">
                                <input name="alba_pay" type="text" value="<%=formatnumber(alba_pay,0)%>" style="width:100px;text-align:right" onKeyUp="give_cal(this);"></td>
								<th style="background:#F8F8FF">소득세</th>
                                <td class="left">
								<input name="tax_amt1" type="text" value="<%=formatnumber(tax_amt1,0)%>" style="width:100px;text-align:right" onKeyUp="taxx_cal(this);"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">교통비</th>
								<td class="left">
                                <input name="alba_trans" type="text" value="<%=formatnumber(alba_trans,0)%>" style="width:100px;text-align:right" onKeyUp="give_cal(this);"></td>
								<th style="background:#F8F8FF">지방소득세</th>
                                <td class="left">
								<input name="tax_amt2" type="text" value="<%=formatnumber(tax_amt2,0)%>" style="width:100px;text-align:right" onKeyUp="taxx_cal(this);"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">식대</th>
								<td class="left">
                                <input name="alba_meals" type="text" value="<%=formatnumber(alba_meals,0)%>" style="width:100px;text-align:right" onKeyUp="give_cal(this);"></td>
								<th style="background:#F8F8FF">&nbsp;</th>
                                <td class="left">&nbsp;
								<input name="de_other" type="hidden" value="<%=formatnumber(de_other,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#F5FFFA">기타지급</th>
								<td class="left">
                                <input name="alba_other" type="text" value="<%=formatnumber(alba_other,0)%>" style="width:100px;text-align:right" onKeyUp="give_cal(this);"></td>
								<th style="background:#F8F8FF">세액계</th>
                                <td class="left">
								<input name="tax_tot" type="text" value="<%=formatnumber(tax_hap,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                             <tr>
								<th class="first" style="background:#F5FFFA">지급총액</th>
								<td class="left">
                                <input name="give_tot" type="text" value="<%=formatnumber(alba_give_total,0)%>" style="width:100px;text-align:right" readonly="true"></td>
								<th style="background:#F8F8FF">차인지급액</th>
                                <td class="left">
								<input name="curr_pay" type="text" value="<%=formatnumber(curr_pay,0)%>" style="width:100px;text-align:right" readonly="true"></td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F8F8FF">작업일수</th>
								<td class="left">
                                <input name="alba_cnt" type="text" value="<%=formatnumber(alba_cnt,0)%>" style="width:100px;text-align:right"></td>
								<th style="background:#F8F8FF">작업량</th>
                                <td class="left">
								<input name="alba_work" type="text" value="<%=formatnumber(alba_work,0)%>" style="width:100px;text-align:right"></td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F8F8FF">작업내용</th>
								<td colspan="3" class="left">
                                <input name="work_comment" type="text" value="<%=work_comment%>" style="width:550px" onKeyUp="checklength(this,100)"></td>
							</tr>

                      </tbody>
					</table>
                    <h3 class="stit" style="font-size:12px;">※ 금액 수정시 "0" 인경우 반드시 "0"을 입력하시기 바랍니다!</h3>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
            <% if u_type = "U" then	%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();"></span>
			<% end if	%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">

                <input type="hidden" name="old_rever_yymm" value="<%=rever_yymm%>">
                <input type="hidden" name="old_draft_no" value="<%=draft_no%>">
				<input type="hidden" name="old_give_date" value="<%=give_date%>">
			</form>
		</div>
	</body>
</html>

