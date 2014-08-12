import xlsxwriter

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Chart Title')
worksheet.write('A2', '=$A$3 & "% Share"')
worksheet.write('A3', 0)

worksheet2 = workbook.add_worksheet()
#worksheet2.hide()

worksheet2.write('A1', 'Dial Colors')
worksheet2.write('A2', 60)
worksheet2.write('A3', 60)
worksheet2.write('A4', 60)
worksheet2.write('A5', 60)
worksheet2.write('A6', 60)
worksheet2.write('A7', 60)
worksheet2.write('B1', 'Needle')
worksheet2.write('B2', 180)
worksheet2.write('B3', '=((180/200)*(Sheet1!A3+100))-B4')
worksheet2.write('B4', 4)
worksheet2.write('B5', '=360-SUM(B2:B4)')

def create_dial():
    data_axis = []
    data_val = []
    for x in xrange(0, 360):
        data_axis.append(x)
        width = 15
        decrement = 0.4
        data_val.append("=IF(ABS(ABS(ROW()-Sheet1!A3-2-180)-180)<%s,8-(%s*ABS(ABS(ROW()-Sheet1!A3-2-180)-180)),0)"
                            % (width, decrement))

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

    worksheet.insert_chart('B2', arrow_radar_chart, {
        #'x_offset': -10, 
        'x_offset': 0, 
        'y_offset': 0,
    })

nps_chart = workbook.add_chart({
    'type': 'doughnut',
    'orientation': 270,
})

nps_chart.add_series({
    'values': '=Sheet2!$A2:$A$7',
    'points': [
        {'fill': {'color': '#FC7613'}},
        {'fill': {'color': '#FFD641'}},
        {'fill': {'color': '#B6C94C'}},
        {'fill': {'none': True}},
        {'fill': {'none': True}},
        {'fill': {'none': True}},
    ],
})

worksheet.insert_chart('B2', nps_chart, {
    'x_offset': 10, 
    'y_offset': 0,
})

create_dial()

workbook.close()
