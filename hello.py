import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Dial Colors')
worksheet.write('A2', 60)
worksheet.write('A3', 60)
worksheet.write('A4', 60)
worksheet.write('A5', 60)
worksheet.write('A6', 60)
worksheet.write('A7', 60)
worksheet.write('B1', 'Needle')
worksheet.write('B2', 180)
worksheet.write('B3', '=((180/200)*(C3+100))-B4')
worksheet.write('B4', 4)
worksheet.write('B5', '=360-SUM(B2:B4)')
worksheet.write('C1', 'Chart Title')
worksheet.write('C2', '=$C$3 & "% Share"')
worksheet.write('C3', 0)

worksheet2 = workbook.add_worksheet()

def create_dial():
    data_axis = []
    data_val = []
    for x in xrange(0, 360):
        data_axis.append(x)
        data_val.append("=IF(ABS(ABS(ROW()-Sheet1!C3-2-180)-180)<10,8-(0.25*ABS(ABS(ROW()-Sheet1!C3-2-180)-180)),0)")

    worksheet2.write_column('D2', data_axis)
    worksheet2.write_column('E2', data_val)


    arrow_radar_chart = workbook.add_chart({
        'type': 'radar',
        'subtype': 'filled',
    })

    arrow_radar_chart.add_series({
        'categories': '=Sheet2!$D$2:$D$361',
        'values':     '=Sheet2!$E$2:$E$361',
        'fill': {'color': '#FFFFFF'},
    })

    arrow_radar_chart.set_legend({
        'none': True,
    })

    worksheet.insert_chart('F11', arrow_radar_chart, {
        'x_offset': -10, 
        'y_offset': 0,
    })

nps_chart = workbook.add_chart({
    'type': 'doughnut',
    'orientation': 270,
})

nps_chart.add_series({
    'values': '=Sheet1!$A2:$A$7',
    'points': [
        {'fill': {'color': '#FC7613'}},
        {'fill': {'color': '#FFD641'}},
        {'fill': {'color': '#B6C94C'}},
        {'fill': {'none': True}},
        {'fill': {'none': True}},
        {'fill': {'none': True}},
    ],
})

worksheet.insert_chart('F11', nps_chart, {
    'x_offset': 0, 
    'y_offset': 0,
})

create_dial()

workbook.close()
