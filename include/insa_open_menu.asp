				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				%>
					<a href="insa_open_emp_add.asp" class="btnType01">�λ�⺻���</a>
                    <a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=300')" class="btnType01">�������׵��</a>
                    
					<a href="#" onClick="pop_Window('insa_school_add.asp?sch_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_school_add_pop','scrollbars=yes,width=750,height=300')" class="btnType01">�зµ��</a>
					<a href="#" onClick="pop_Window('insa_career_add.asp?career_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_career_add_pop','scrollbars=yes,width=750,height=300')" class="btnType01">��µ��</a>
                    <a href="#" onClick="pop_Window('insa_individual_qual_add.asp?qual_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_qual_add_pop','scrollbars=yes,width=750,height=300')" class="btnType01">�ڰ� ���</a>
                    <a href="#" onClick="pop_Window('insa_edu_add.asp?edu_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_edu_add_pop','scrollbars=yes,width=750,height=300')" class="btnType01">���� ���</a>
                    <a href="#" onClick="pop_Window('insa_language_add.asp?lang_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_language_add_pop','scrollbars=yes,width=750,height=300')" class="btnType01">���л��� ���</a>
 				</div>
