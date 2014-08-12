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
worksheet.write('C3', 100)

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

worksheet.insert_chart('C6', nps_chart)


needle_chart = workbook.add_chart({
    'type': 'pie',
    'orientation': 90,
})

needle_chart.add_series({
        'values': '=Sheet1!$B2:$B$5',
        'points': [
            {'fill': {'none': True}},
            {'fill': {'none': True}},
            #{'fill': {'color': 'green'}},
            {'fill': {'color': 'blue'}},
            #{'fill': {'color': 'brown'}},
            #{'fill': {'color': 'orange'}},
            {'fill': {'none': True}},
        ],
    })

worksheet.insert_chart('C6', needle_chart)

workbook.close()
