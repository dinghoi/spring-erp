			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/></h1>
				<h2><img src="/image/cost_title.gif" alt="비용 관리 시스템" width="189" height="25"/></h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>님</strong> 안녕하세요
					</p>
				</div>	
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="person_cost_report.asp">비용등록관리</a></li>    
                        <li class="dep1"><a href="emp_cost_mg.asp">비용현황관리</a></li>                        
                        <li class="dep1">                       
                    <%	if cost_grade = "0" or (bonbu = "ITO 사업본부" and position = "사업부장") or (bonbu = "ITO 사업본부" and position = "본부장") then	%>						
                        <a href="saupbu_profit_loss_total.asp">손익현황</a>
					<%	  else	%>
                        <a>손익현황</a>
					<%	end if	%>
						</li>
                        <li class="dep1">                       
                    <%	if cost_grade = "0" and (user_id = "100539" or user_id = "100031" or user_id = "100041" or user_id = "900001" or user_id = "100359") then	%>						
                        <a href="cost_grade_mg.asp">비용코드관리</a>
					<%	  else	%>
                        <a>비용코드관리</a>
					<%	end if	%>
						</li>
                    </ul>
                </div>
			</div>
