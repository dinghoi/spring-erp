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
                    <a href="insa_pay_yeartax_before.asp?y_final=<%=y_final%>" class="btnType01">�����ٹ���</a>
                    <a href="insa_pay_yeartax_mg.asp?y_final=<%=y_final%>" class="btnType01">�ҵ�������</a>
                    <a href="insa_pay_yeartax_person.asp?y_final=<%=y_final%>.asp" class="btnType01">����������</a>
                    <a href="insa_pay_yeartax_annuity.asp?y_final=<%=y_final%>.asp" class="btnType01">���ݺ���</a>
                    <a href="insa_pay_yeartax_insurance.asp?y_final=<%=y_final%>.asp" class="btnType01">Ư������</a>
                    <a href="insa_pay_yeartax_saving.asp?s_id=<%="��������"%>&y_final=<%=y_final%>" class="btnType01">�׹��� ����</a>
                    <a href="insa_pay_yeartax_credit.asp?c_id=<%="�ſ�ī��"%>&y_final=<%=y_final%>" class="btnType01">�ſ�ī�����</a>
                    <a href="insa_pay_yeartax_other.asp?y_final=<%=y_final%>" class="btnType01">��Ÿ����</a>
                    <a href="insa_pay_yeartax_deduction.asp?y_final=<%=y_final%>" class="btnType01">���װ���/����</a>
                  <% '  <a href="insa_pay_yeartax_final_submit.asp" class="btnType01">��������</a> %>
                    <a href="#" onClick="pop_Window('insa_pay_yeartax_final_submit.asp?u_type=<%="U"%>','insa_user_password_pop','scrollbars=yes,width=750,height=350')" class="btnType01">��������</a>
                    <a href="insa_pay_yeartax_medical_report.asp" class="btnType01">�ҵ������µ�</a>
                  <% '  <a href="insa_pay_yeartax_tax_report.asp" class="btnType01">�ҵ������µ�</a> %>
                    <a href="insa_pay_yeartax_wonchen_report.asp" class="btnType01">��õ¡��������</a>
				</div>
