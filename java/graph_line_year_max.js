			var chart;
			var chart2;
			var chart3;
			var chart4;
			$(document).ready(function() {
				s_tab10 = parseInt(document.frm.s_tab10.value.replace(/,/g,""));
				s_tab11 = parseInt(document.frm.s_tab11.value.replace(/,/g,""));
				s_tab12 = parseInt(document.frm.s_tab12.value.replace(/,/g,""));
				s_tab13 = parseInt(document.frm.s_tab13.value.replace(/,/g,""));
				s_tab14 = parseInt(document.frm.s_tab14.value.replace(/,/g,""));
				s_tab15 = parseInt(document.frm.s_tab15.value.replace(/,/g,""));
				s_tab20 = parseInt(document.frm.s_tab20.value.replace(/,/g,""));
				s_tab21 = parseInt(document.frm.s_tab21.value.replace(/,/g,""));
				s_tab22 = parseInt(document.frm.s_tab22.value.replace(/,/g,""));
				s_tab23 = parseInt(document.frm.s_tab23.value.replace(/,/g,""));
				s_tab24 = parseInt(document.frm.s_tab24.value.replace(/,/g,""));
				s_tab25 = parseInt(document.frm.s_tab25.value.replace(/,/g,""));
				s_tab30 = parseInt(document.frm.s_tab30.value.replace(/,/g,""));
				s_tab31 = parseInt(document.frm.s_tab31.value.replace(/,/g,""));
				s_tab32 = parseInt(document.frm.s_tab32.value.replace(/,/g,""));
				s_tab33 = parseInt(document.frm.s_tab33.value.replace(/,/g,""));
				s_tab34 = parseInt(document.frm.s_tab34.value.replace(/,/g,""));
				s_tab35 = parseInt(document.frm.s_tab35.value.replace(/,/g,""));
				s_tab40 = parseInt(document.frm.s_tab40.value.replace(/,/g,""));
				s_tab41 = parseInt(document.frm.s_tab41.value.replace(/,/g,""));
				s_tab42 = parseInt(document.frm.s_tab42.value.replace(/,/g,""));
				s_tab43 = parseInt(document.frm.s_tab43.value.replace(/,/g,""));
				s_tab44 = parseInt(document.frm.s_tab44.value.replace(/,/g,""));
				s_tab45 = parseInt(document.frm.s_tab45.value.replace(/,/g,""));
				s_tab50 = parseInt(document.frm.s_tab50.value.replace(/,/g,""));
				s_tab51 = parseInt(document.frm.s_tab51.value.replace(/,/g,""));
				s_tab52 = parseInt(document.frm.s_tab52.value.replace(/,/g,""));
				s_tab53 = parseInt(document.frm.s_tab53.value.replace(/,/g,""));
				s_tab54 = parseInt(document.frm.s_tab54.value.replace(/,/g,""));
				s_tab55 = parseInt(document.frm.s_tab55.value.replace(/,/g,""));
				s_tab60 = parseInt(document.frm.s_tab60.value.replace(/,/g,""));
				s_tab61 = parseInt(document.frm.s_tab61.value.replace(/,/g,""));
				s_tab62 = parseInt(document.frm.s_tab62.value.replace(/,/g,""));
				s_tab63 = parseInt(document.frm.s_tab63.value.replace(/,/g,""));
				s_tab64 = parseInt(document.frm.s_tab64.value.replace(/,/g,""));
				s_tab65 = parseInt(document.frm.s_tab65.value.replace(/,/g,""));
				s_tab70 = parseInt(document.frm.s_tab70.value.replace(/,/g,""));
				s_tab71 = parseInt(document.frm.s_tab71.value.replace(/,/g,""));
				s_tab72 = parseInt(document.frm.s_tab72.value.replace(/,/g,""));
				s_tab73 = parseInt(document.frm.s_tab73.value.replace(/,/g,""));
				s_tab74 = parseInt(document.frm.s_tab74.value.replace(/,/g,""));
				s_tab75 = parseInt(document.frm.s_tab75.value.replace(/,/g,""));
				s_tab80 = parseInt(document.frm.s_tab80.value.replace(/,/g,""));
				s_tab81 = parseInt(document.frm.s_tab81.value.replace(/,/g,""));
				s_tab82 = parseInt(document.frm.s_tab82.value.replace(/,/g,""));
				s_tab83 = parseInt(document.frm.s_tab83.value.replace(/,/g,""));
				s_tab84 = parseInt(document.frm.s_tab84.value.replace(/,/g,""));
				s_tab85 = parseInt(document.frm.s_tab85.value.replace(/,/g,""));
				s_tab90 = parseInt(document.frm.s_tab90.value.replace(/,/g,""));
				s_tab91 = parseInt(document.frm.s_tab91.value.replace(/,/g,""));
				s_tab92 = parseInt(document.frm.s_tab92.value.replace(/,/g,""));
				s_tab93 = parseInt(document.frm.s_tab93.value.replace(/,/g,""));
				s_tab94 = parseInt(document.frm.s_tab94.value.replace(/,/g,""));
				s_tab95 = parseInt(document.frm.s_tab95.value.replace(/,/g,""));
				s_tab100 = parseInt(document.frm.s_tab100.value.replace(/,/g,""));
				s_tab101 = parseInt(document.frm.s_tab101.value.replace(/,/g,""));
				s_tab102 = parseInt(document.frm.s_tab102.value.replace(/,/g,""));
				s_tab103 = parseInt(document.frm.s_tab103.value.replace(/,/g,""));
				s_tab104 = parseInt(document.frm.s_tab104.value.replace(/,/g,""));
				s_tab105 = parseInt(document.frm.s_tab105.value.replace(/,/g,""));
				s_tab110 = parseInt(document.frm.s_tab110.value.replace(/,/g,""));
				s_tab111 = parseInt(document.frm.s_tab111.value.replace(/,/g,""));
				s_tab112 = parseInt(document.frm.s_tab112.value.replace(/,/g,""));
				s_tab113 = parseInt(document.frm.s_tab113.value.replace(/,/g,""));
				s_tab114 = parseInt(document.frm.s_tab114.value.replace(/,/g,""));
				s_tab115 = parseInt(document.frm.s_tab115.value.replace(/,/g,""));
				s_tab120 = parseInt(document.frm.s_tab120.value.replace(/,/g,""));
				s_tab121 = parseInt(document.frm.s_tab121.value.replace(/,/g,""));
				s_tab122 = parseInt(document.frm.s_tab122.value.replace(/,/g,""));
				s_tab123 = parseInt(document.frm.s_tab123.value.replace(/,/g,""));
				s_tab124 = parseInt(document.frm.s_tab124.value.replace(/,/g,""));
				s_tab125 = parseInt(document.frm.s_tab125.value.replace(/,/g,""));

				emp_tab10 = parseInt(document.frm.emp_tab10.value.replace(/,/g,""));
				emp_tab11 = parseInt(document.frm.emp_tab11.value.replace(/,/g,""));
				emp_tab12 = parseInt(document.frm.emp_tab12.value.replace(/,/g,""));
				emp_tab13 = parseInt(document.frm.emp_tab13.value.replace(/,/g,""));
				emp_tab14 = parseInt(document.frm.emp_tab14.value.replace(/,/g,""));
				emp_tab15 = parseInt(document.frm.emp_tab15.value.replace(/,/g,""));
				emp_tab20 = parseInt(document.frm.emp_tab20.value.replace(/,/g,""));
				emp_tab21 = parseInt(document.frm.emp_tab21.value.replace(/,/g,""));
				emp_tab22 = parseInt(document.frm.emp_tab22.value.replace(/,/g,""));
				emp_tab23 = parseInt(document.frm.emp_tab23.value.replace(/,/g,""));
				emp_tab24 = parseInt(document.frm.emp_tab24.value.replace(/,/g,""));
				emp_tab25 = parseInt(document.frm.emp_tab25.value.replace(/,/g,""));
				emp_tab30 = parseInt(document.frm.emp_tab30.value.replace(/,/g,""));
				emp_tab31 = parseInt(document.frm.emp_tab31.value.replace(/,/g,""));
				emp_tab32 = parseInt(document.frm.emp_tab32.value.replace(/,/g,""));
				emp_tab33 = parseInt(document.frm.emp_tab33.value.replace(/,/g,""));
				emp_tab34 = parseInt(document.frm.emp_tab34.value.replace(/,/g,""));
				emp_tab35 = parseInt(document.frm.emp_tab35.value.replace(/,/g,""));
				emp_tab40 = parseInt(document.frm.emp_tab40.value.replace(/,/g,""));
				emp_tab41 = parseInt(document.frm.emp_tab41.value.replace(/,/g,""));
				emp_tab42 = parseInt(document.frm.emp_tab42.value.replace(/,/g,""));
				emp_tab43 = parseInt(document.frm.emp_tab43.value.replace(/,/g,""));
				emp_tab44 = parseInt(document.frm.emp_tab44.value.replace(/,/g,""));
				emp_tab45 = parseInt(document.frm.emp_tab45.value.replace(/,/g,""));
				emp_tab50 = parseInt(document.frm.emp_tab50.value.replace(/,/g,""));
				emp_tab51 = parseInt(document.frm.emp_tab51.value.replace(/,/g,""));
				emp_tab52 = parseInt(document.frm.emp_tab52.value.replace(/,/g,""));
				emp_tab53 = parseInt(document.frm.emp_tab53.value.replace(/,/g,""));
				emp_tab54 = parseInt(document.frm.emp_tab54.value.replace(/,/g,""));
				emp_tab55 = parseInt(document.frm.emp_tab55.value.replace(/,/g,""));
				emp_tab60 = parseInt(document.frm.emp_tab60.value.replace(/,/g,""));
				emp_tab61 = parseInt(document.frm.emp_tab61.value.replace(/,/g,""));
				emp_tab62 = parseInt(document.frm.emp_tab62.value.replace(/,/g,""));
				emp_tab63 = parseInt(document.frm.emp_tab63.value.replace(/,/g,""));
				emp_tab64 = parseInt(document.frm.emp_tab64.value.replace(/,/g,""));
				emp_tab65 = parseInt(document.frm.emp_tab65.value.replace(/,/g,""));
				emp_tab70 = parseInt(document.frm.emp_tab70.value.replace(/,/g,""));
				emp_tab71 = parseInt(document.frm.emp_tab71.value.replace(/,/g,""));
				emp_tab72 = parseInt(document.frm.emp_tab72.value.replace(/,/g,""));
				emp_tab73 = parseInt(document.frm.emp_tab73.value.replace(/,/g,""));
				emp_tab74 = parseInt(document.frm.emp_tab74.value.replace(/,/g,""));
				emp_tab75 = parseInt(document.frm.emp_tab75.value.replace(/,/g,""));
				emp_tab80 = parseInt(document.frm.emp_tab80.value.replace(/,/g,""));
				emp_tab81 = parseInt(document.frm.emp_tab81.value.replace(/,/g,""));
				emp_tab82 = parseInt(document.frm.emp_tab82.value.replace(/,/g,""));
				emp_tab83 = parseInt(document.frm.emp_tab83.value.replace(/,/g,""));
				emp_tab84 = parseInt(document.frm.emp_tab84.value.replace(/,/g,""));
				emp_tab85 = parseInt(document.frm.emp_tab85.value.replace(/,/g,""));
				emp_tab90 = parseInt(document.frm.emp_tab90.value.replace(/,/g,""));
				emp_tab91 = parseInt(document.frm.emp_tab91.value.replace(/,/g,""));
				emp_tab92 = parseInt(document.frm.emp_tab92.value.replace(/,/g,""));
				emp_tab93 = parseInt(document.frm.emp_tab93.value.replace(/,/g,""));
				emp_tab94 = parseInt(document.frm.emp_tab94.value.replace(/,/g,""));
				emp_tab95 = parseInt(document.frm.emp_tab95.value.replace(/,/g,""));
				emp_tab100 = parseInt(document.frm.emp_tab100.value.replace(/,/g,""));
				emp_tab101 = parseInt(document.frm.emp_tab101.value.replace(/,/g,""));
				emp_tab102 = parseInt(document.frm.emp_tab102.value.replace(/,/g,""));
				emp_tab103 = parseInt(document.frm.emp_tab103.value.replace(/,/g,""));
				emp_tab104 = parseInt(document.frm.emp_tab104.value.replace(/,/g,""));
				emp_tab105 = parseInt(document.frm.emp_tab105.value.replace(/,/g,""));
				emp_tab110 = parseInt(document.frm.emp_tab110.value.replace(/,/g,""));
				emp_tab111 = parseInt(document.frm.emp_tab111.value.replace(/,/g,""));
				emp_tab112 = parseInt(document.frm.emp_tab112.value.replace(/,/g,""));
				emp_tab113 = parseInt(document.frm.emp_tab113.value.replace(/,/g,""));
				emp_tab114 = parseInt(document.frm.emp_tab114.value.replace(/,/g,""));
				emp_tab115 = parseInt(document.frm.emp_tab115.value.replace(/,/g,""));
				emp_tab120 = parseInt(document.frm.emp_tab120.value.replace(/,/g,""));
				emp_tab121 = parseInt(document.frm.emp_tab121.value.replace(/,/g,""));
				emp_tab122 = parseInt(document.frm.emp_tab122.value.replace(/,/g,""));
				emp_tab123 = parseInt(document.frm.emp_tab123.value.replace(/,/g,""));
				emp_tab124 = parseInt(document.frm.emp_tab124.value.replace(/,/g,""));
				emp_tab125 = parseInt(document.frm.emp_tab125.value.replace(/,/g,""));

				p_tab10 = parseInt(document.frm.p_tab10.value.replace(/,/g,""));
				p_tab11 = parseInt(document.frm.p_tab11.value.replace(/,/g,""));
				p_tab12 = parseInt(document.frm.p_tab12.value.replace(/,/g,""));
				p_tab13 = parseInt(document.frm.p_tab13.value.replace(/,/g,""));
				p_tab14 = parseInt(document.frm.p_tab14.value.replace(/,/g,""));
				p_tab15 = parseInt(document.frm.p_tab15.value.replace(/,/g,""));
				p_tab20 = parseInt(document.frm.p_tab20.value.replace(/,/g,""));
				p_tab21 = parseInt(document.frm.p_tab21.value.replace(/,/g,""));
				p_tab22 = parseInt(document.frm.p_tab22.value.replace(/,/g,""));
				p_tab23 = parseInt(document.frm.p_tab23.value.replace(/,/g,""));
				p_tab24 = parseInt(document.frm.p_tab24.value.replace(/,/g,""));
				p_tab25 = parseInt(document.frm.p_tab25.value.replace(/,/g,""));
				p_tab30 = parseInt(document.frm.p_tab30.value.replace(/,/g,""));
				p_tab31 = parseInt(document.frm.p_tab31.value.replace(/,/g,""));
				p_tab32 = parseInt(document.frm.p_tab32.value.replace(/,/g,""));
				p_tab33 = parseInt(document.frm.p_tab33.value.replace(/,/g,""));
				p_tab34 = parseInt(document.frm.p_tab34.value.replace(/,/g,""));
				p_tab35 = parseInt(document.frm.p_tab35.value.replace(/,/g,""));
				p_tab40 = parseInt(document.frm.p_tab40.value.replace(/,/g,""));
				p_tab41 = parseInt(document.frm.p_tab41.value.replace(/,/g,""));
				p_tab42 = parseInt(document.frm.p_tab42.value.replace(/,/g,""));
				p_tab43 = parseInt(document.frm.p_tab43.value.replace(/,/g,""));
				p_tab44 = parseInt(document.frm.p_tab44.value.replace(/,/g,""));
				p_tab45 = parseInt(document.frm.p_tab45.value.replace(/,/g,""));
				p_tab50 = parseInt(document.frm.p_tab50.value.replace(/,/g,""));
				p_tab51 = parseInt(document.frm.p_tab51.value.replace(/,/g,""));
				p_tab52 = parseInt(document.frm.p_tab52.value.replace(/,/g,""));
				p_tab53 = parseInt(document.frm.p_tab53.value.replace(/,/g,""));
				p_tab54 = parseInt(document.frm.p_tab54.value.replace(/,/g,""));
				p_tab55 = parseInt(document.frm.p_tab55.value.replace(/,/g,""));
				p_tab60 = parseInt(document.frm.p_tab60.value.replace(/,/g,""));
				p_tab61 = parseInt(document.frm.p_tab61.value.replace(/,/g,""));
				p_tab62 = parseInt(document.frm.p_tab62.value.replace(/,/g,""));
				p_tab63 = parseInt(document.frm.p_tab63.value.replace(/,/g,""));
				p_tab64 = parseInt(document.frm.p_tab64.value.replace(/,/g,""));
				p_tab65 = parseInt(document.frm.p_tab65.value.replace(/,/g,""));
				p_tab70 = parseInt(document.frm.p_tab70.value.replace(/,/g,""));
				p_tab71 = parseInt(document.frm.p_tab71.value.replace(/,/g,""));
				p_tab72 = parseInt(document.frm.p_tab72.value.replace(/,/g,""));
				p_tab73 = parseInt(document.frm.p_tab73.value.replace(/,/g,""));
				p_tab74 = parseInt(document.frm.p_tab74.value.replace(/,/g,""));
				p_tab75 = parseInt(document.frm.p_tab75.value.replace(/,/g,""));
				p_tab80 = parseInt(document.frm.p_tab80.value.replace(/,/g,""));
				p_tab81 = parseInt(document.frm.p_tab81.value.replace(/,/g,""));
				p_tab82 = parseInt(document.frm.p_tab82.value.replace(/,/g,""));
				p_tab83 = parseInt(document.frm.p_tab83.value.replace(/,/g,""));
				p_tab84 = parseInt(document.frm.p_tab84.value.replace(/,/g,""));
				p_tab85 = parseInt(document.frm.p_tab85.value.replace(/,/g,""));
				p_tab90 = parseInt(document.frm.p_tab90.value.replace(/,/g,""));
				p_tab91 = parseInt(document.frm.p_tab91.value.replace(/,/g,""));
				p_tab92 = parseInt(document.frm.p_tab92.value.replace(/,/g,""));
				p_tab93 = parseInt(document.frm.p_tab93.value.replace(/,/g,""));
				p_tab94 = parseInt(document.frm.p_tab94.value.replace(/,/g,""));
				p_tab95 = parseInt(document.frm.p_tab95.value.replace(/,/g,""));
				p_tab100 = parseInt(document.frm.p_tab100.value.replace(/,/g,""));
				p_tab101 = parseInt(document.frm.p_tab101.value.replace(/,/g,""));
				p_tab102 = parseInt(document.frm.p_tab102.value.replace(/,/g,""));
				p_tab103 = parseInt(document.frm.p_tab103.value.replace(/,/g,""));
				p_tab104 = parseInt(document.frm.p_tab104.value.replace(/,/g,""));
				p_tab105 = parseInt(document.frm.p_tab105.value.replace(/,/g,""));
				p_tab110 = parseInt(document.frm.p_tab110.value.replace(/,/g,""));
				p_tab111 = parseInt(document.frm.p_tab111.value.replace(/,/g,""));
				p_tab112 = parseInt(document.frm.p_tab112.value.replace(/,/g,""));
				p_tab113 = parseInt(document.frm.p_tab113.value.replace(/,/g,""));
				p_tab114 = parseInt(document.frm.p_tab114.value.replace(/,/g,""));
				p_tab115 = parseInt(document.frm.p_tab115.value.replace(/,/g,""));
				p_tab120 = parseInt(document.frm.p_tab120.value.replace(/,/g,""));
				p_tab121 = parseInt(document.frm.p_tab121.value.replace(/,/g,""));
				p_tab122 = parseInt(document.frm.p_tab122.value.replace(/,/g,""));
				p_tab123 = parseInt(document.frm.p_tab123.value.replace(/,/g,""));
				p_tab124 = parseInt(document.frm.p_tab124.value.replace(/,/g,""));
				p_tab125 = parseInt(document.frm.p_tab125.value.replace(/,/g,""));

				b_tab11 = parseInt(document.frm.b_tab11.value.replace(/,/g,""));
				b_tab12 = parseInt(document.frm.b_tab12.value.replace(/,/g,""));
				b_tab13 = parseInt(document.frm.b_tab13.value.replace(/,/g,""));
				b_tab21 = parseInt(document.frm.b_tab21.value.replace(/,/g,""));
				b_tab22 = parseInt(document.frm.b_tab22.value.replace(/,/g,""));
				b_tab23 = parseInt(document.frm.b_tab23.value.replace(/,/g,""));
				b_tab31 = parseInt(document.frm.b_tab31.value.replace(/,/g,""));
				b_tab32 = parseInt(document.frm.b_tab32.value.replace(/,/g,""));
				b_tab33 = parseInt(document.frm.b_tab33.value.replace(/,/g,""));
				b_tab41 = parseInt(document.frm.b_tab41.value.replace(/,/g,""));
				b_tab42 = parseInt(document.frm.b_tab42.value.replace(/,/g,""));
				b_tab43 = parseInt(document.frm.b_tab43.value.replace(/,/g,""));
				b_tab51 = parseInt(document.frm.b_tab51.value.replace(/,/g,""));
				b_tab52 = parseInt(document.frm.b_tab52.value.replace(/,/g,""));
				b_tab53 = parseInt(document.frm.b_tab53.value.replace(/,/g,""));
				b_tab61 = parseInt(document.frm.b_tab61.value.replace(/,/g,""));
				b_tab62 = parseInt(document.frm.b_tab62.value.replace(/,/g,""));
				b_tab63 = parseInt(document.frm.b_tab63.value.replace(/,/g,""));
				b_tab71 = parseInt(document.frm.b_tab71.value.replace(/,/g,""));
				b_tab72 = parseInt(document.frm.b_tab72.value.replace(/,/g,""));
				b_tab73 = parseInt(document.frm.b_tab73.value.replace(/,/g,""));
				b_tab81 = parseInt(document.frm.b_tab81.value.replace(/,/g,""));
				b_tab82 = parseInt(document.frm.b_tab82.value.replace(/,/g,""));
				b_tab83 = parseInt(document.frm.b_tab83.value.replace(/,/g,""));
				b_tab91 = parseInt(document.frm.b_tab91.value.replace(/,/g,""));
				b_tab92 = parseInt(document.frm.b_tab92.value.replace(/,/g,""));
				b_tab93 = parseInt(document.frm.b_tab93.value.replace(/,/g,""));
				b_tab101 = parseInt(document.frm.b_tab101.value.replace(/,/g,""));
				b_tab102 = parseInt(document.frm.b_tab102.value.replace(/,/g,""));
				b_tab103 = parseInt(document.frm.b_tab103.value.replace(/,/g,""));
				b_tab111 = parseInt(document.frm.b_tab111.value.replace(/,/g,""));
				b_tab112 = parseInt(document.frm.b_tab112.value.replace(/,/g,""));
				b_tab113 = parseInt(document.frm.b_tab113.value.replace(/,/g,""));
				b_tab121 = parseInt(document.frm.b_tab121.value.replace(/,/g,""));
				b_tab122 = parseInt(document.frm.b_tab122.value.replace(/,/g,""));
				b_tab123 = parseInt(document.frm.b_tab123.value.replace(/,/g,""));

				view_year = document.frm.view_year.value;
				title_text = view_year + "년 회사별 비용 사용 현황"
				title_text2 = view_year + "년 회사별 인력 현황"
				title_text3 = view_year + "년 회사별 급여 현황"
				title_text4 = view_year + "년 매출 및 손익 현황"

			if (document.frm.view_id.value == "1") {
				chart = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view',
						defaultSeriesType: 'line',
						marginRight: 130,
						marginBottom: 60
					},
					 credits: {
								enabled: false
					},
					title: {
						text: title_text,
						x: -20 //center
					},
					xAxis: {
						categories: ['1월', '2월', '3월', '4월', '5월', '6월','7월', '8월', '9월', '10월', '11월', '12월']
					},
					yAxis: {
						title: {
							text: '사용금액 (백만원)'
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					tooltip: {
						formatter: function() {
				                return '<b>'+ this.series.name +'</b><br/>'+
//								this.x +': '+ this.y +'점';
								this.y +'백만원';
						}
					},
					legend: {
						layout: 'vertical',
						align: 'right',
						verticalAlign: 'top',
						x: 10,
						y: 200,
						borderWidth: 0
					},
					series: [{
						name: '총괄',
						data: [s_tab10, s_tab20, s_tab30, s_tab40, s_tab50, s_tab60, s_tab70, s_tab80, s_tab90, s_tab100, s_tab110, s_tab120]
					}, {
						name: '케이원정보통신',
						data: [s_tab11, s_tab21, s_tab31, s_tab41, s_tab51, s_tab61, s_tab71, s_tab81, s_tab91, s_tab101, s_tab111, s_tab121]
					}, {
						name: '휴디스',
						data: [s_tab12, s_tab22, s_tab32, s_tab42, s_tab52, s_tab62, s_tab72, s_tab82, s_tab92, s_tab102, s_tab112, s_tab122]
					}, {
						name: '케이네트웍스',
						data: [s_tab13, s_tab23, s_tab33, s_tab43, s_tab53, s_tab63, s_tab73, s_tab83, s_tab93, s_tab103, s_tab113, s_tab123]
					}, {
						name: '코리아디엔씨',
						data: [s_tab14, s_tab24, s_tab34, s_tab44, s_tab54, s_tab64, s_tab74, s_tab84, s_tab94, s_tab104, s_tab114, s_tab124]
					}]
				});								
			}
			if (document.frm.view_id.value == "2") {
				chart2 = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view2',
						defaultSeriesType: 'line',
						marginRight: 130,
						marginBottom: 60
					},
					 credits: {
								enabled: false
					},
					title: {
						text: title_text2,
						x: -20 //center
					},
					xAxis: {
						categories: ['1월', '2월', '3월', '4월', '5월', '6월','7월', '8월', '9월', '10월', '11월', '12월']
					},
					yAxis: {
						title: {
							text: '인원수 (명)'
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					tooltip: {
						formatter: function() {
				                return '<b>'+ this.series.name +'</b><br/>'+
//								this.x +': '+ this.y +'점';
								this.y +'명';
						}
					},
					legend: {
						layout: 'vertical',
						align: 'right',
						verticalAlign: 'top',
						x: 10,
						y: 200,
						borderWidth: 0
					},
					series: [{
						name: '총괄',
						data: [emp_tab10, emp_tab20, emp_tab30, emp_tab40, emp_tab50, emp_tab60, emp_tab70, emp_tab80, emp_tab90, emp_tab100, emp_tab110, emp_tab120]
					}, {
						name: '케이원정보통신',
						data: [emp_tab11, emp_tab21, emp_tab31, emp_tab41, emp_tab51, emp_tab61, emp_tab71, emp_tab81, emp_tab91, emp_tab101, emp_tab111, emp_tab121]
					}, {
						name: '휴디스',
						data: [emp_tab12, emp_tab22, emp_tab32, emp_tab42, emp_tab52, emp_tab62, emp_tab72, emp_tab82, emp_tab92, emp_tab102, emp_tab112, emp_tab122]
					}, {
						name: '케이네트웍스',
						data: [emp_tab13, emp_tab23, emp_tab33, emp_tab43, emp_tab53, emp_tab63, emp_tab73, emp_tab83, emp_tab93, emp_tab103, emp_tab113, emp_tab123]
					}, {
						name: '코리아디엔씨',
						data: [emp_tab14, emp_tab24, emp_tab34, emp_tab44, emp_tab54, emp_tab64, emp_tab74, emp_tab84, emp_tab94, emp_tab104, emp_tab114, emp_tab124]
					}]
				});								
			}

			if (document.frm.view_id.value == "3") {
				chart3 = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view3',
						defaultSeriesType: 'line',
						marginRight: 130,
						marginBottom: 60
					},
					 credits: {
								enabled: false
					},
					title: {
						text: title_text3,
						x: -20 //center
					},
					xAxis: {
						categories: ['1월', '2월', '3월', '4월', '5월', '6월','7월', '8월', '9월', '10월', '11월', '12월']
					},
					yAxis: {
						title: {
							text: '급여 (백만원)'
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					tooltip: {
						formatter: function() {
				                return '<b>'+ this.series.name +'</b><br/>'+
//								this.x +': '+ this.y +'점';
								this.y +'백만원';
						}
					},
					legend: {
						layout: 'vertical',
						align: 'right',
						verticalAlign: 'top',
						x: 10,
						y: 200,
						borderWidth: 0
					},
					series: [{
						name: '총괄',
						data: [p_tab10, p_tab20, p_tab30, p_tab40, p_tab50, p_tab60, p_tab70, p_tab80, p_tab90, p_tab100, p_tab110, p_tab120]
					}, {
						name: '케이원정보통신',
						data: [p_tab11, p_tab21, p_tab31, p_tab41, p_tab51, p_tab61, p_tab71, p_tab81, p_tab91, p_tab101, p_tab111, p_tab121]
					}, {
						name: '휴디스',
						data: [p_tab12, p_tab22, p_tab32, p_tab42, p_tab52, p_tab62, p_tab72, p_tab82, p_tab92, p_tab102, p_tab112, p_tab122]
					}, {
						name: '케이네트웍스',
						data: [p_tab13, p_tab23, p_tab33, p_tab43, p_tab53, p_tab63, p_tab73, p_tab83, p_tab93, p_tab103, p_tab113, p_tab123]
					}, {
						name: '코리아디엔씨',
						data: [p_tab14, p_tab24, p_tab34, p_tab44, p_tab54, p_tab64, p_tab74, p_tab84, p_tab94, p_tab104, p_tab114, p_tab124]
					}]
				});								
			}

			if (document.frm.view_id.value == "4") {
				chart4 = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view4',
						defaultSeriesType: 'line',
						marginRight: 130,
						marginBottom: 60
					},
					 credits: {
								enabled: false
					},
					title: {
						text: title_text4,
						x: -20 //center
					},
					xAxis: {
						categories: ['1월', '2월', '3월', '4월', '5월', '6월','7월', '8월', '9월', '10월', '11월', '12월']
					},
					yAxis: {
						title: {
							text: '금액 (백만원)'
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					tooltip: {
						formatter: function() {
				                return '<b>'+ this.series.name +'</b><br/>'+
//								this.x +': '+ this.y +'점';
								this.y +'백만원';
						}
					},
					legend: {
						layout: 'vertical',
						align: 'right',
						verticalAlign: 'top',
						x: 10,
						y: 200,
						borderWidth: 0
					},

					series: [{
						name: '매출',
						data: [b_tab11, b_tab21, b_tab31, b_tab41, b_tab51, b_tab61, b_tab71, b_tab81, b_tab91, b_tab101, b_tab111, b_tab121]
					}, {
						name: '비용',
						data: [b_tab12, b_tab22, b_tab32, b_tab42, b_tab52, b_tab62, b_tab72, b_tab82, b_tab92, b_tab102, b_tab112, b_tab122]
					}, {
						name: '손익',
						data: [b_tab13, b_tab23, b_tab33, b_tab43, b_tab53, b_tab63, b_tab73, b_tab83, b_tab93, b_tab103, b_tab113, b_tab123]
					}]
				});								
			}

			});
