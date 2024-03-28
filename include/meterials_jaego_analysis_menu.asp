				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>  
					<a href="meterials_system_popup.asp" class="btnType01">재고변동 추이</a>
                    <a href="meterials_system_popup.asp" class="btnType01">재고평가</a>
                    <a href="meterials_system_popup.asp" class="btnType01">품목별재고 비고분석</a>
				</div>
