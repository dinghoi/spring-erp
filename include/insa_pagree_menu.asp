				<div class="btnRight">
                <%
				in_name = request.cookies("srvmg_user")("coo_user_name")
                in_empno = request.cookies("srvmg_user")("coo_user_id")
				%>
					<a href="insa_system_popup.asp" class="btnType01">입사서약서</a>
                    <a href="insa_year_income_vow.asp" class="btnType01">연봉근로계약서</a>
                    <a href="insa_system_popup.asp" class="btnType01">소프트웨어사용서약서</a>
                    <a href="insa_system_popup.asp" class="btnType01">보안서약서</a>
                    <%
					'<a href="insa_join_company_vow1.asp" class="btnType01">입사서약서</a>
                    '<a href="insa_year_income_vow1.asp" class="btnType01">연봉근로계약서</a>
                    '<a href="insa_soft_vow.asp" class="btnType01">소프트웨어사용서약서</a>
                    '<a href="insa_security_vow.asp" class="btnType01">보안서약서</a>
					%>
 				</div>
