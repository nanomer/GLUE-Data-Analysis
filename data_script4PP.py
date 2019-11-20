#!/usr/bin/env python3

import sys
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import os
from scipy import stats

# ASSUMPTIONS:
# 1. Step size of x-axis is 1 or -1
# 2. Everything is named correctly
# 3. For every IDVG there is a corresponding IDVD file

# GLOBAL VARIABLES
workbookName = ''

def main():
    global workbookName

    # Assume only data text files in Raw Data/
    files = os.scandir('Raw Data/')

    # First make sure all files are same sample ID
    checkSameSampleID = []
    for file in files:
        filename = file.name
        sampleID = filename[filename.find('CMM'): filename.find(' ', filename.find('.'))]
        checkSameSampleID.append(sampleID)
    workbookName = checkSameSampleID[0]
    for check in checkSameSampleID:
        if check != workbookName:
            sys.exit('Files have different Sample IDs. Please check Raw Data.')

    # Create workbook, the excel file
    workbook = xlsxwriter.Workbook('Processed Data/' + workbookName + '.xlsx')
    files = os.scandir('Raw Data/')

    for file in files:
        process_file(workbook, file)

    workbook.close()

def process_file(workbook, file):
    # Grab the info from the file name to name the worksheet and set constants
    filename = file.name
    startIndex = filename.find('CMM')
    sIndex = filename.find('s', startIndex)
    cIndex = filename.find('c', startIndex)
    dIndex = filename.find('d', startIndex)
    lIndex = filename.find('L', startIndex)
    wIndex = filename.find('W', startIndex)
    kIndex = filename.find('K', startIndex)
    sampleType = filename[0:4]
    sampleNum = filename[sIndex + 1: filename.find(' ', sIndex)]
    cap = float(filename[cIndex + 1: filename.find(' ', cIndex)])
    deviceNum = filename[dIndex + 1: filename.find(' ', dIndex)]
    length = int(filename[filename.find(' ', dIndex) + 1: lIndex])
    width = int(filename[filename.find(' ', lIndex) + 1: wIndex])
    temperature = int(filename[filename.find(' ', wIndex) + 1: kIndex])
    worksheetName = 'S' + sampleNum + ' D' + deviceNum + ' ' + str(temperature) + 'K ' + str(length) + 'L ' + sampleType
    worksheet = workbook.add_worksheet(worksheetName)

    curFile = open(r'Raw Data/' + filename)

    # Find the primary information
    line = curFile.readline()
    while line.find('Measurement.Primary.Start') == -1:
        line = curFile.readline()
    priStart = int(line[line.find('\t') + 1: line.find('\n')])
    line = curFile.readline()
    priStop = int(line[line.find('\t') + 1: line.find('\n')])
    line = curFile.readline()
    priSteps = int(line[line.find('\t') + 1: line.find('\n')]) # Should be -1 or 1

    # Find the secondary information
    line = curFile.readline()
    while line.find('Measurement.Secondary.Start') == -1:
        line = curFile.readline()
    secStart = int(line[line.find('\t') + 1: line.find('\n')])
    line = curFile.readline()
    secCount = int(line[line.find('\t') + 1: line.find('\n')])
    line = curFile.readline()
    secSteps = int(line[line.find('\t') + 1: line.find('\n')])
    lowerBound = abs(secSteps) + 5

    # Skip lines until you reach data
    line = curFile.readline()
    while line.find('Ig') == -1 or line.find('Id') == -1 or line.find('Vg') == -1:
        line = curFile.readline()

    # Time to start populating w/ data (First 3 columns)
    row = 0
    col = 0  # Raw data always starts at column A (0)
    primary = 'Vg' # 4PP primary always Vg
    secondary = 'Vd'
    worksheet.write(row, col, line[0: 2])
    worksheet.write(row, col + 1, line[3: 5])
    worksheet.write(row, col + 2, line[6: 8])
    worksheet.write(row, col + 3, line[9: 11])
    worksheet.write(row, col + 4, line[12: 14])
    worksheet.write(row, col + 5, line[15:17])
    row = 1
    line = curFile.readline()

    # Skip zero Vd
    for i in range(202):
        line = curFile.readline();

    IdLin = []  # for calculating trend line manually later
    V2Lin = []
    V1Lin = []
    while line:
        nextIndex = 0;
        curVg = float(line[nextIndex: line.find('\t', nextIndex)])
        for i in range(6):
            num = float(line[nextIndex: line.find('\t', nextIndex)])
            worksheet.write(row, col + i, num)
            if abs(curVg) >= lowerBound:
                if i == 1:
                    IdLin.append(num)
                elif i == 3:
                    V1Lin.append(num)
                elif i == 4:
                    V2Lin.append(num)
            nextIndex = line.find('\t', nextIndex) + 1
        row += 1
        line = curFile.readline()
    curFile.close()

    ### Now worksheet has all the data from the file ###

    #Useful variables predefined here
    endRow = (abs(priStop) - priStart + 1) * 2  # num of rows
    lwRatio = length/width
    baseSecInterval = secSteps + secStart
    maxX = priStop
    midX = (priStart + priStop) // 2 + (priStop // 10) # Should be 60/-60, not really mid
    minX = priStart
    reverse = True
    lineXStart = lowerBound + 1
    lineXEnd = 200 - lowerBound + 2
    skipGraph = 11

    #Graph dict, starts w/ values for first graph (abs) and will change for others
    title = {'name': 'Voltage Probe Original Readings'}
    yAxis = {'name': 'Voltage Probe Reading (V)',
             'label_position': 'high',
             'num_format': '#.#0E-0#',
             'num_font': {'bold': 1},
             'name_font': {'size': 14},
             'name_layout': {'x': 0.03, 'y': 0.3},
             }
    xAxis = {'name': 'VGATE(V)',
             'reverse': reverse,
             'major_gridlines': {'visible': True},
             'min': maxX,
             'max': minX,
             'name_font': {'size': 14},
             'num_font': {'bold': 1},
             'label_position': 'low',
             }

    # Correct V1 data
    col = 6  # Original data always starts at column E (4)
    worksheet.write(0, col, 'Correct V1')

    for i in range(endRow):
        worksheet.write(i + 1, col, '=D' + str(i + 2) + '*5')

    # Correct V2 data
    col += 1
    worksheet.write(0, col, 'Correct V2')

    for i in range(endRow):
        worksheet.write(i + 1, col, '=E' + str(i + 2) + '*5')

    # V2-V1
    col += 1
    worksheet.write(0, col, 'V2-V1')

    for i in range(endRow):
        worksheet.write(i + 1, col, '=E' + str(i + 2) + '-D' + str(i + 2))

    # Correct V2-V1
    col += 1
    worksheet.write(0, col, 'Correct V2-V1')

    for i in range(endRow):
        worksheet.write(i + 1, col, '=H' + str(i + 2) + '-G' + str(i + 2))

    #G = I/V
    col += 1
    worksheet.write(0, col, 'G')

    for i in range(endRow):
        worksheet.write(i + 1, col, '=B' + str(i + 2) + '/I' + str(i + 2))

    #Correct G = I/V
    col += 1
    worksheet.write(0, col, 'Correct G')

    for i in range(endRow):
        worksheet.write(i + 1, col, '=B' + str(i + 2) + '/J' + str(i + 2))

    # dG/dVg
    col += 1
    worksheet.write(0, col, 'dG/dVg')

    for i in range(endRow - 2):
        worksheet.write_formula(i + 1, col, '=LINEST(K' + str(i + 2) + ':K' + str(i + 4) +
                                ',A' + str(i + 2) + ':A' + str(i + 4) + ')')
    # Correct dG/dVg
    col += 1
    worksheet.write(0, col, 'Correct dG/dVg')
    for i in range(endRow - 2):
        worksheet.write_formula(i + 1, col, '=LINEST(L' + str(i + 2) + ':L' + str(i + 4) +
                                ',A' + str(i + 2) + ':A' + str(i + 4) + ')')

    # 4PP Mob
    col += 1
    worksheet.write(0, col, '4PP Mob')
    for i in range(endRow - 2):
        worksheet.write(i + 1, col, '=M' + str(i + 2) + '*' + str(lwRatio) + '/' + str(cap))

    # 4PP Corrected Mob
    col += 1
    worksheet.write(0, col, '4PP Corrected Mob')
    for i in range(endRow - 2):
        worksheet.write(i + 1, col, '=N' + str(i + 2) + '*' + str(lwRatio) + '/' + str(cap))

    # Calculate trend line values
    mOrig, bOrig, mCorrect, bCorrect, xInterOrig, xInterCorrect = calc_trendline(IdLin, V1Lin, V2Lin)

    # Device Constants
    col += 2
    worksheet.write(1, col, 'Device Constants')
    worksheet.write(2, col, 'W')
    worksheet.write(2, col + 1, width)
    worksheet.write(3, col, 'L')
    worksheet.write(3, col + 1, length)
    worksheet.write(4, col, 'Cins')
    worksheet.write(4, col + 1, cap)

    # Correct Mob
    worksheet.write(6, col, 'Correct Mob')
    worksheet.write(7, col, '=' + str(lwRatio / cap * mCorrect[0]))

    ### GRAPHING ###
    col += 3
    origProbeChart = workbook.add_chart({'type': 'scatter'})
    graph(origProbeChart, title, yAxis, xAxis)
    origProbeChart.add_series({ 'values': [worksheetName, 1, 3, endRow, 3],
                                'categories': [worksheetName, 1, 0, endRow, 0],
                                'name': 'V1',
                                'name_font': {'bold': 1},
                                'line': {'dash_type': 'round_dot'},
                                'marker': {'type': 'circle'},
                                'min': maxX
                            })
    origProbeChart.add_series({ 'values': [worksheetName, 1, 4, endRow, 4],
                                'categories': [worksheetName, 1, 0, endRow, 0],
                                'name': 'V2',
                                'name_font': {'bold': 1},
                                'line': {'dash_type': 'round_dot'},
                                'marker': {'type': 'circle'},
                                'min': maxX
                            })
    worksheet.insert_chart(xl_rowcol_to_cell(1, col), origProbeChart)

    correctProbeChart = workbook.add_chart({'type': 'scatter'})
    title['name'] = 'Voltage Probe Corrected Readings'
    graph(correctProbeChart, title, yAxis, xAxis)
    correctProbeChart.add_series({ 'values': [worksheetName, 1, 6, endRow, 6],
                                'categories': [worksheetName, 1, 0, endRow, 0],
                                'name': 'Correct V1',
                                'name_font': {'bold': 1},
                                'line': {'dash_type': 'round_dot'},
                                'marker': {'type': 'circle'},
                                'min': maxX
                            })
    correctProbeChart.add_series({ 'values': [worksheetName, 1, 7, endRow, 7],
                                'categories': [worksheetName, 1, 0, endRow, 0],
                                'name': 'Correct V2',
                                'name_font': {'bold': 1},
                                'line': {'dash_type': 'round_dot'},
                                'marker': {'type': 'circle'},
                                'min': maxX
                            })
    worksheet.insert_chart(xl_rowcol_to_cell(26, col), correctProbeChart)

    col += skipGraph
    gChart = workbook.add_chart({'type': 'scatter'})
    title['name'] = 'G'
    yAxis['name'] = 'G (A/V)'
    xAxis['min'] = -100
    xAxis['max'] = lowerBound * -1
    graph(gChart, title, yAxis, xAxis)
    gChart.add_series({ 'values': [worksheetName, lineXStart, 10, lineXEnd, 10],
                        'categories': [worksheetName, lineXStart, 0, lineXEnd, 0],
                        'name': 'G',
                        'name_font': {'bold': 1},
                        'line': {'dash_type': 'round_dot'},
                        'marker': {'type': 'circle'},
                        'trendline': {'type': 'linear',
                                      'display_equation': True,
                                      'name': 'Linear (G)',
                                      },
                    })
    worksheet.insert_chart(xl_rowcol_to_cell(1, col), gChart)

    correctgChart = workbook.add_chart({'type': 'scatter'})
    title['name'] = 'Corrected G'
    graph(correctgChart, title, yAxis, xAxis)
    correctgChart.add_series({ 'values': [worksheetName, lineXStart, 11, lineXEnd, 11],
                        'categories': [worksheetName, lineXStart, 0, lineXEnd, 0],
                        'name': 'Correct G',
                        'name_font': {'bold': 1},
                        'line': {'dash_type': 'round_dot'},
                        'marker': {'type': 'circle'},
                        'trendline': {'type': 'linear',
                                      'display_equation': True,
                                      'name': 'Linear\n(Correct G)',
                                      },
                    })
    worksheet.insert_chart(xl_rowcol_to_cell(26, col), correctgChart)

    col += skipGraph
    origMobChart = workbook.add_chart({'type': 'scatter'})
    title['name'] = '4PP Original Mobility'
    yAxis['name'] = 'Mobility (cm^2/Vs)'
    xAxis['min'] = maxX
    xAxis['max'] = minX
    graph(origMobChart, title, yAxis, xAxis)
    origMobChart.add_series({'values': [worksheetName, 1, 14, endRow, 14],
                       'categories': [worksheetName, 1, 0, endRow, 0],
                       'name': '4PP\nMobility',
                       'name_font': {'bold': 1},
                       'line': {'dash_type': 'round_dot'},
                       'marker': {'type': 'circle'},
                       })
    worksheet.insert_chart(xl_rowcol_to_cell(1, col), origMobChart)

    correctMobChart = workbook.add_chart({'type': 'scatter'})
    title['name'] = '4PP Corrected Mobility'
    graph(correctMobChart, title, yAxis, xAxis)
    correctMobChart.add_series({'values': [worksheetName, 1, 15, endRow, 15],
                       'categories': [worksheetName, 1, 0, endRow, 0],
                       'name': '4PP\nCorrected\nMobility',
                       'name_font': {'bold': 1},
                       'line': {'dash_type': 'round_dot'},
                       'marker': {'type': 'circle'},
                       })
    worksheet.insert_chart(xl_rowcol_to_cell(26, col), correctMobChart)

def graph(chart, title, yAxis, xAxis):
    global workbookName
    chart.set_size({'width': 680,
                   'height': 480,
                   })
    chart.set_plotarea({'layout': {'x': 0.17,
                                  'y': 0.1,
                                  'width': 0.63,
                                  'height': 0.73
                                  }
                       })
    chart.set_legend({'font': {'bold': 1, 'size': 14}})
    chart.set_title(title)
    chart.set_y_axis(yAxis)
    chart.set_x_axis(xAxis)

def calc_trendline(id, v1, v2):
    x = []
    length = len(id)
    for i in range(length // 2):
        x.append(-(100 - length // 2 + 1) - i)
    for i in range(-100, -(100 - length // 2)):
        x.append(i)
    mOrig = []
    bOrig = []
    mCorrect = []
    bCorrect = []
    xInterOrig = []
    xInterCorrect = []

    curOrigY = []
    curCorrectY = []
    for i in range(length):
        curOrigY.append(id[i] / (v2[i] - v1[i]))
        curCorrectY.append(id[i] / (5 * (v2[i] - v1[i])))
    slopeOrig, interceptOrig, r_valueOrig, p_valueOrig, std_errOrig = stats.linregress(x, curOrigY)
    slopeCorrect, interceptCorrect, r_valueCorrect, p_valueCorrect, std_errCorrect = stats.linregress(x, curCorrectY)
    mOrig.append(slopeOrig)
    bOrig.append(interceptOrig * -1)
    mCorrect.append(slopeCorrect)
    bCorrect.append(interceptCorrect * -1)
    xInterOrig.append(interceptOrig * -1 / slopeOrig)
    xInterCorrect.append(interceptCorrect * -1 / slopeCorrect)

    return mOrig, bOrig, mCorrect, bCorrect, xInterOrig, xInterCorrect

if __name__ == '__main__':
    main()
    sys.exit(42)
