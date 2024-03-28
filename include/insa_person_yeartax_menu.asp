				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				
				Set Dbconn=Server.CreateObject("ADODB.Connection")
                Set rs_yyyy = Server.CreateObject("ADODB.Recordset")
                dbconn.open DbConnect
				
				sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
                rs_yyyy.Open Sql, Dbconn, 1
                if not rs_yyyy.eof then
                       y_final =  rs_yyyy("y_final") 
                   else	   
	                   y_final =  ""
                end if
				rs_yyyy.close()	
				%>                
                    <a href="insa_pay_yeartax_before.asp?y_final=<%=y_final%>" class="btnType01">이전근무지</a>
                    <a href="insa_pay_yeartax_mg.asp?y_final=<%=y_final%>" class="btnType01">소득자정보</a>
                    <a href="insa_pay_yeartax_person.asp?y_final=<%=y_final%>.asp" class="btnType01">인적공제명세</a>
                    <a href="insa_pay_yeartax_annuity.asp?y_final=<%=y_final%>.asp" class="btnType01">연금보험</a>
                    <a href="insa_pay_yeartax_insurance.asp?y_final=<%=y_final%>.asp" class="btnType01">특별공제</a>
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="연금저축"%>&y_final=<%=y_final%>" class="btnType01">그밖의 공제</a>
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="신용카드"%>&y_final=<%=y_final%>" class="btnType01">신용카드공제</a>
                    <a href="insa_pay_yeartax_other.asp?y_final=<%=y_final%>" class="btnType01">기타공제</a>
                    <a href="insa_pay_yeartax_deduction.asp?y_final=<%=y_final%>" class="btnType01">세액감면/공제</a>
                  <% '  <a href="insa_pay_yeartax_final_submit.asp" class="btnType01">최종제출</a> %>
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_final_submit.asp?u_type=<%="U"%>','insa_user_password_pop','scrollbars=yes,width=750,height=350')" class="btnType01">최종제출</a>
                    <a href="insa_pay_yeartax_medical_report.asp" class="btnType01">소득공제출력등</a>
                  <% '  <a href="insa_pay_yeartax_tax_report.asp" class="btnType01">소득공제출력등</a> %>
                    <a href="insa_pay_yeartax_wonchen_report.asp" class="btnType01">원천징수영수증</a>
				</div>
