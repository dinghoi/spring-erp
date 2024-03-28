			var chart;
			$(document).ready(function() {
				j = "3";
				var mon_tab = new Array();
				mon_tab[0] = document.frm.mon_tab0.value;
				mon_tab[1] = document.frm.mon_tab1.value;
				mon_tab[2] = document.frm.mon_tab2.value;
				mon_tab[3] = document.frm.mon_tab3.value;
				mon_tab[4] = document.frm.mon_tab4.value;
				mon_tab[5] = document.frm.mon_tab5.value;
				mon_tab[6] = document.frm.mon_tab6.value;
				mon_tab[7] = document.frm.mon_tab7.value;
				mon_tab[8] = document.frm.mon_tab8.value;
				mon_tab[9] = document.frm.mon_tab9.value;
				mon_tab[10] = document.frm.mon_tab10.value;
				mon_tab[11] = document.frm.mon_tab11.value;
				mon_tab[12] = document.frm.mon_tab12.value;
				s_tab00 = parseFloat(document.frm.s_tab00.value.replace(/,/g,""));
				s_tab01 = parseFloat(document.frm.s_tab01.value.replace(/,/g,""));
				s_tab02 = parseFloat(document.frm.s_tab02.value.replace(/,/g,""));
				s_tab03 = parseFloat(document.frm.s_tab03.value.replace(/,/g,""));
				s_tab10 = parseFloat(document.frm.s_tab10.value.replace(/,/g,""));
				s_tab11 = parseFloat(document.frm.s_tab11.value.replace(/,/g,""));
				s_tab12 = parseFloat(document.frm.s_tab12.value.replace(/,/g,""));
				s_tab13 = parseFloat(document.frm.s_tab13.value.replace(/,/g,""));
				s_tab20 = parseFloat(document.frm.s_tab20.value.replace(/,/g,""));
				s_tab21 = parseFloat(document.frm.s_tab21.value.replace(/,/g,""));
				s_tab22 = parseFloat(document.frm.s_tab22.value.replace(/,/g,""));
				s_tab23 = parseFloat(document.frm.s_tab23.value.replace(/,/g,""));
				s_tab30 = parseFloat(document.frm.s_tab30.value.replace(/,/g,""));
				s_tab31 = parseFloat(document.frm.s_tab31.value.replace(/,/g,""));
				s_tab32 = parseFloat(document.frm.s_tab32.value.replace(/,/g,""));
				s_tab33 = parseFloat(document.frm.s_tab33.value.replace(/,/g,""));
				s_tab40 = parseFloat(document.frm.s_tab40.value.replace(/,/g,""));
				s_tab41 = parseFloat(document.frm.s_tab41.value.replace(/,/g,""));
				s_tab42 = parseFloat(document.frm.s_tab42.value.replace(/,/g,""));
				s_tab43 = parseFloat(document.frm.s_tab43.value.replace(/,/g,""));
				s_tab50 = parseFloat(document.frm.s_tab50.value.replace(/,/g,""));
				s_tab51 = parseFloat(document.frm.s_tab51.value.replace(/,/g,""));
				s_tab52 = parseFloat(document.frm.s_tab52.value.replace(/,/g,""));
				s_tab53 = parseFloat(document.frm.s_tab53.value.replace(/,/g,""));
				s_tab60 = parseFloat(document.frm.s_tab60.value.replace(/,/g,""));
				s_tab61 = parseFloat(document.frm.s_tab61.value.replace(/,/g,""));
				s_tab62 = parseFloat(document.frm.s_tab62.value.replace(/,/g,""));
				s_tab63 = parseFloat(document.frm.s_tab63.value.replace(/,/g,""));
				s_tab70 = parseFloat(document.frm.s_tab70.value.replace(/,/g,""));
				s_tab71 = parseFloat(document.frm.s_tab71.value.replace(/,/g,""));
				s_tab72 = parseFloat(document.frm.s_tab72.value.replace(/,/g,""));
				s_tab73 = parseFloat(document.frm.s_tab73.value.replace(/,/g,""));
				s_tab80 = parseFloat(document.frm.s_tab80.value.replace(/,/g,""));
				s_tab81 = parseFloat(document.frm.s_tab81.value.replace(/,/g,""));
				s_tab82 = parseFloat(document.frm.s_tab82.value.replace(/,/g,""));
				s_tab83 = parseFloat(document.frm.s_tab83.value.replace(/,/g,""));
				s_tab90 = parseFloat(document.frm.s_tab90.value.replace(/,/g,""));
				s_tab91 = parseFloat(document.frm.s_tab91.value.replace(/,/g,""));
				s_tab92 = parseFloat(document.frm.s_tab92.value.replace(/,/g,""));
				s_tab93 = parseFloat(document.frm.s_tab93.value.replace(/,/g,""));
				s_tab100 = parseFloat(document.frm.s_tab100.value.replace(/,/g,""));
				s_tab101 = parseFloat(document.frm.s_tab101.value.replace(/,/g,""));
				s_tab102 = parseFloat(document.frm.s_tab102.value.replace(/,/g,""));
				s_tab103 = parseFloat(document.frm.s_tab103.value.replace(/,/g,""));
				s_tab110 = parseFloat(document.frm.s_tab110.value.replace(/,/g,""));
				s_tab111 = parseFloat(document.frm.s_tab111.value.replace(/,/g,""));
				s_tab112 = parseFloat(document.frm.s_tab112.value.replace(/,/g,""));
				s_tab113 = parseFloat(document.frm.s_tab113.value.replace(/,/g,""));
				s_tab120 = parseFloat(document.frm.s_tab120.value.replace(/,/g,""));
				s_tab121 = parseFloat(document.frm.s_tab121.value.replace(/,/g,""));
				s_tab122 = parseFloat(document.frm.s_tab122.value.replace(/,/g,""));
				s_tab123 = parseFloat(document.frm.s_tab123.value.replace(/,/g,""));
				sla_month = document.frm.sla_month.value;
				sla_yy = sla_month.substr(2,2);
				sla_mm = sla_month.substr(4,2);

				chart = new Highcharts.Chart({
					chart: {
						renderTo: 'graph_view',
						defaultSeriesType: 'column'
					},
					title: {
						text: document.frm.graph_title.value
					},
//					subtitle: {
//						text: 'Source: WorldClimate.com'
//					},
					xAxis: {
						categories: [ mon_tab[0], mon_tab[1],  mon_tab[2],  mon_tab[3],  mon_tab[4],  mon_tab[5], mon_tab[6], 
									mon_tab[7], mon_tab[8], mon_tab[9], mon_tab[10], mon_tab[11], mon_tab[12]]
					},
					yAxis: {
						min: 0,
						title: {
							text: 'SLA 점수'
						}
					},
					legend: {
						layout: 'vertical',
						backgroundColor: '#FFFFFF',
						align: 'left',
						verticalAlign: 'top',
						x: 170,
						y: 0,
						floating: true,
						shadow: true
					},
					tooltip: {
						formatter: function() {
							return ''+
								this.x +': '+ this.y +' 점';
						}
					},
					plotOptions: {
						column: {
							pointPadding: 0.2,
							borderWidth: 0
						}
					},
				        series: [{
						name: '주전산기',
						data: [s_tab00, s_tab10, s_tab20, s_tab30, s_tab40, s_tab50, s_tab60, s_tab70, s_tab80, s_tab90, s_tab100, s_tab110, s_tab120]
				
					}, {
						name: '1지역',
						data: [s_tab01, s_tab11, s_tab21, s_tab31, s_tab41, s_tab51, s_tab61, s_tab71, s_tab81, s_tab91, s_tab101, s_tab111, s_tab121]
				
					}, {
						name: '2지역',
						data: [s_tab02, s_tab12, s_tab22, s_tab32, s_tab42, s_tab52, s_tab62, s_tab72, s_tab82, s_tab92, s_tab102, s_tab112, s_tab122]
				
					}, {
						name: '3지역',
						data: [s_tab03, s_tab13, s_tab23, s_tab33, s_tab43, s_tab53, s_tab63, s_tab73, s_tab83, s_tab93, s_tab103, s_tab113, s_tab123]
				
					}]
				});
				
				
			});
