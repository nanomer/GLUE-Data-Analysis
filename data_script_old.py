import sys

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

import os

from scipy import stats

# Assume only data text files in Raw Data/
files = os.scandir('Raw Data/')

# First make sure all files are same sample ID
checkSameSampleID = []
for file in files:
    filename = file.name
    sampleID = filename[filename.find('CMM') : filename.find(' ', filename.find('.'))]
    checkSameSampleID.append(sampleID)
workbookName = checkSameSampleID[0]
for check in checkSameSampleID:
    if check != workbookName:
        sys.exit('Files have different Sample IDs. Please check Raw Data.')

# Create workbook, the excel file
workbook = xlsxwriter.Workbook('Processed Data/' + workbookName + '.xlsx')
files = os.scandir('Raw Data/')
idvdWorksheets = []
curWS = 0

# Create each worksheet from files
for file in files:
    filename = file.name
    startIndex = filename.find('CMM')
    sIndex = filename.find('s', startIndex)
    cIndex = filename.find('c', startIndex)
    dIndex = filename.find('d', startIndex)
    lIndex = filename.find('L', startIndex)
    wIndex = filename.find('W', startIndex)
    kIndex = filename.find('K', startIndex)

    sampleType = filename[0:4]
    sampleNum = filename[sIndex + 1 : filename.find(' ', sIndex)]
    cap = float(filename[cIndex + 1 : filename.find(' ', cIndex)])
    deviceNum = filename[dIndex + 1 : filename.find(' ', dIndex)]

    length = int(filename[filename.find(' ', dIndex) + 1 : lIndex])
    width = int(filename[filename.find(' ', lIndex) + 1 : wIndex])
    temperature = int(filename[filename.find(' ', wIndex) + 1 : kIndex])

    worksheetName = 'S' + sampleNum + ' D' + deviceNum + ' ' + str(temperature) + 'K ' + str(length) + 'L ' + sampleType
    worksheet = workbook.add_worksheet(worksheetName)

    curFile = open(r'Raw Data/'+filename)

    # Find the primary information
    line = curFile.readline()
    while line.find('Measurement.Primary.Start') == -1:
        line = curFile.readline()
    priStart = int(line[line.find('\t') + 1 : line.find('\n')])
    line = curFile.readline()
    priStop = int(line[line.find('\t') + 1 : line.find('\n')])
    line = curFile.readline()
    priSteps = int(line[line.find('\t') + 1 : line.find('\n')])

    # Find the secondary information
    line = curFile.readline()
    while line.find('Measurement.Secondary.Start') == -1:
        line = curFile.readline()
    secStart = int(line[line.find('\t') + 1 : line.find('\n')])
    line = curFile.readline()
    secCount = int(line[line.find('\t') + 1 : line.find('\n')])
    line = curFile.readline()
    secSteps = int(line[line.find('\t') + 1 : line.find('\n')])

    line = curFile.readline()
    while line.find('Ig') == -1 or line.find('Id') == -1 or line.find('V') == -1:
        line = curFile.readline()

    # Time to start populating w/ data (First 3 columns)
    row = 0
    col = 0 # Raw data always starts at column A (0)
    primary = line[0 : 2]
    secondary = 'Vd'
    if primary == 'Vd':
        secondary = 'Vg'
    worksheet.write(row, col, line[0 : 2])
    worksheet.write(row, col + 1, line[3 : 5])
    worksheet.write(row, col + 2, line[6 : 8])

    row = 1
    line = curFile.readline()
    y_fwd = [] # for calculating trend line manually later
    y_rvs = []
    while line:
        nextIndex = 0;
        for i in range(3):
            worksheet.write(row, col + i, float(line[nextIndex : line.find('\t', nextIndex)]))
            if i == 1:
                if (row >= 263 and row <= 303) or (row >= 465 and row <= 505) or (row >= 667 and row <= 707) or (row >= 869 and row <= 909) or (row >= 1071 and row <= 1111):
                    y_fwd.append(float(line[nextIndex : line.find('\t', nextIndex)]))
                elif (row >=  304 and row <= 344) or (row >= 506 and row <= 546) or (row >= 708 and row <= 748) or (row >= 910 and row <= 950) or (row >= 1112 and row <= 1152):
                    y_rvs.append(float(line[nextIndex : line.find('\t', nextIndex)]))
            nextIndex = line.find('\t', nextIndex) + 1
        row += 1
        line = curFile.readline()
    curFile.close()

    # Useful variables
    endRow = (abs(priStop) - priStart + 1) * 2 # num of rows
    wlRatio = width/length

    # Organize data by steps
    col = 4 # Original data always starts at column E (4)
    row = 1
    for i in range(secCount):
        worksheet.write(0, col, secondary + ' ' + str(secStart + secSteps * i))
        for j in range(endRow):
            worksheet.write(j + 1, col, '=B' + str(row + 1))
            row += 1
        col += 1

    # Absolute value
    col += 1
    absStart = col # Starting col of abs values
    row = 1
    for i in range(secCount):
        worksheet.write(0, col, 'Abs ' + secondary + ' ' + str(secStart + secSteps * i))
        for j in range(endRow):
            worksheet.write_formula(j + 1, col, '=ABS(' + xl_rowcol_to_cell(j + 1, col - secCount - 1) + ')')
        col += 1

    # Plot abs values
    col += 1
    startIDVD = col
    absChart = workbook.add_chart({'type': 'scatter'})
    absChart.set_size({'width': 500,
                       'height': 480,
                       })
    absChart.set_plotarea({'layout': {'x': 0.17,
                                      'y': 0.1,
                                      'width': 0.63,
                                      'height': 0.73
                                      }
                           })
    absChart.set_legend({'font': {'bold': 1, 'size': 14}})
    absChart.set_title({'name': workbookName + ' ' + worksheetName})
    absChart.set_y_axis({'name': 'ABS IDRAIN (A)',
                         'label_position': 'high',
                         'num_format': '#.#0E-0#',
                         'num_font': {'bold': 1},
                         'name_font': {'size': 14},
                         'name_layout': {'x': 0, 'y': 0.4},
                         })
    if primary == 'Vd':
        absChart.set_x_axis({'name': 'VDRAIN (V)',
                             'reverse': True,
                             'major_gridlines': {'visible': True},
                             'min': priStop,
                             'name_font': {'size': 14},
                             'num_font': {'bold': 1},
                             'label_position': 'low',
                             })
    else:
        absChart.set_x_axis({'name': 'VGATE (V)',
                             'reverse': True,
                             'major_gridlines': {'visible': True},
                             'min': priStop,
                             'name_font': {'size': 14},
                             'num_font': {'bold': 1},
                             'label_position': 'low',
                             })
    for i in range(1, secCount):
        absChart.add_series({ 'values': [worksheetName, 1, absStart + i, endRow, absStart + i],
                              'categories': [worksheetName, 1, 0, endRow, 0],
                              'name': str(secStart + secSteps * i),
                              'name_font': {'bold': 1},
                              'line': {'dash_type': 'round_dot'},
                              'marker': {'type': 'circle'},
                              'min': priStop,
                              })
    worksheet.insert_chart(xl_rowcol_to_cell(1, col), absChart)

    # Plot abs values w/ log base
    absLogChart = workbook.add_chart({'type': 'scatter'})
    absLogChart.set_size({'width': 500,
                       'height': 480,
                       })
    absLogChart.set_plotarea({'layout': {'x': 0.17,
                                      'y': 0.1,
                                      'width': 0.63,
                                      'height': 0.73
                                      }
                           })
    absLogChart.set_legend({'font': {'bold': 1, 'size': 14}})
    absLogChart.set_title({'name': workbookName + ' ' + worksheetName})
    absLogChart.set_y_axis({'name': 'ABS IDRAIN (A)',
                         'label_position': 'high',
                         'num_format': '#.#0E-0#',
                         'num_font': {'bold': 1},
                         'name_font': {'size': 14},
                         'name_layout': {'x': 0, 'y': 0.4},
                         'log_base': 10,
                         })
    if primary == 'Vd':
        absLogChart.set_x_axis({'name': 'VDRAIN (V)',
                             'reverse': True,
                             'major_gridlines': {'visible': True},
                             'min': priStop,
                             'name_font': {'size': 14},
                             'num_font': {'bold': 1},
                             'label_position': 'low',
                             })
    else:
        absLogChart.set_x_axis({'name': 'VGATE (V)',
                             'reverse': True,
                             'major_gridlines': {'visible': True},
                             'min': priStop,
                             'name_font': {'size': 14},
                             'num_font': {'bold': 1},
                             'label_position': 'low',
                             })
    for i in range(1, secCount):
        absLogChart.add_series({ 'values': [worksheetName, 1, absStart + i, endRow, absStart + i],
                              'categories': [worksheetName, 1, 0, endRow, 0],
                              'name': str(secStart + secSteps * i),
                              'name_font': {'bold': 1},
                              'line': {'dash_type': 'round_dot'},
                              'marker': {'type': 'circle'},
                              'min': priStop,
                              })
    worksheet.insert_chart(xl_rowcol_to_cell(26, col), absLogChart)

    if primary == 'Vg':
        # Sq root abs values
        col += 9
        sqrtStart = col;
        row = 1
        for i in range(secCount):
            worksheet.write(0, col, 'SQRT Abs ' + secondary + ' ' + str(secStart + secSteps * i))
            for j in range(endRow):
                worksheet.write_formula(j + 1, col, '=SQRT(' + xl_rowcol_to_cell(j + 1, absStart + i) + ')')
            col += 1

        # Plot Sq root FWD
        col += 1
        sqrtFwdChart = workbook.add_chart({'type': 'scatter'})
        sqrtFwdChart.set_size({'width': 540,
                               'height': 480,
                               })
        sqrtFwdChart.set_plotarea({'layout': {'x': 0.17,
                                              'y': 0.1,
                                              'width': 0.63,
                                              'height': 0.73
                                              }
                                   })
        sqrtFwdChart.set_legend({'font': {'bold': 1, 'size': 14}})
        sqrtFwdChart.set_title({'name': workbookName + ' ' + worksheetName + ' FWD VTH'})
        sqrtFwdChart.set_y_axis({'name': 'SQRT ABS IDRAIN (A)',
                                 'label_position': 'high',
                                 'num_format': '#.#0E-0#',
                                 'num_font': {'bold': 1},
                                 'name_font': {'size': 14},
                                 'name_layout': {'x': 0, 'y': 0.4},
                                 })

        sqrtFwdChart.set_x_axis({'name': 'VGATE (V)',
                                 'reverse': True,
                                 'major_gridlines': {'visible': True},
                                 'min': priStop,
                                 'name_font': {'size': 14},
                                 'num_font': {'bold': 1},
                                 'label_position': 'low',
                                })
        for i in range(1, secCount):
            sqrtFwdChart.add_series({'values': [worksheetName, 1, sqrtStart + i, endRow // 2, sqrtStart + i],
                                     'categories': [worksheetName, 1, 0, endRow // 2, 0],
                                     'name': str(secStart + secSteps * i) + ' V',
                                     'name_font': {'bold': 1},
                                     'line': {'dash_type': 'round_dot'},
                                     'marker': {'type': 'circle'},
                                     'min': priStop,
                                     })
        worksheet.insert_chart(xl_rowcol_to_cell(1, col), sqrtFwdChart)

        # Plot Sq root RVS
        sqrtRvsChart = workbook.add_chart({'type': 'scatter'})
        sqrtRvsChart.set_size({'width': 540,
                               'height': 480,
                               })
        sqrtRvsChart.set_plotarea({'layout': {'x': 0.17,
                                              'y': 0.1,
                                              'width': 0.63,
                                              'height': 0.73
                                              }
                                   })
        sqrtRvsChart.set_legend({'font': {'bold': 1, 'size': 14}})
        sqrtRvsChart.set_title({'name': workbookName + ' ' + worksheetName + ' RVS VTH'})
        sqrtRvsChart.set_y_axis({'name': 'SQRT ABS IDRAIN (A)',
                                 'label_position': 'high',
                                 'num_format': '#.#0E-0#',
                                 'num_font': {'bold': 1},
                                 'name_font': {'size': 14},
                                 'name_layout': {'x': 0, 'y': 0.4},
                                 })

        sqrtRvsChart.set_x_axis({'name': 'VGATE (V)',
                                 'reverse': True,
                                 'major_gridlines': {'visible': True},
                                 'min': priStop,
                                 'name_font': {'size': 14},
                                 'num_font': {'bold': 1},
                                 'label_position': 'low',
                                })
        for i in range(1, secCount):
            sqrtRvsChart.add_series({'values': [worksheetName, endRow // 2 + 1, sqrtStart + i, endRow, sqrtStart + i],
                                     'categories': [worksheetName, endRow // 2 + 1, 0, endRow, 0],
                                     'name': str(secStart + secSteps * i) + ' V',
                                     'name_font': {'bold': 1},
                                     'line': {'dash_type': 'round_dot'},
                                     'marker': {'type': 'circle'},
                                     'min': priStop,
                                     })
        worksheet.insert_chart(xl_rowcol_to_cell(26, col), sqrtRvsChart)

        # Trend line Fwd
        col += 9
        trendFwdChart = workbook.add_chart({'type': 'scatter'})
        trendFwdChart.set_size({'width': 540,
                               'height': 480,
                               })
        trendFwdChart.set_plotarea({'layout': {'x': 0.17,
                                              'y': 0.1,
                                              'width': 0.63,
                                              'height': 0.73
                                              }
                                   })
        trendFwdChart.set_legend({'font': {'bold': 1, 'size': 14}})
        trendFwdChart.set_title({'name': workbookName + ' ' + worksheetName + ' FWD VTH'})
        trendFwdChart.set_y_axis({'name': 'SQRT ABS IDRAIN (A)',
                                 'label_position': 'high',
                                 'num_format': '#.#0E-0#',
                                 'num_font': {'bold': 1},
                                 'name_font': {'size': 14},
                                 'name_layout': {'x': 0, 'y': 0.4},
                                 })

        trendFwdChart.set_x_axis({'name': 'VGATE (V)',
                                 'reverse': True,
                                 'major_gridlines': {'visible': True},
                                 'min': priStop,
                                 'max': priStop + 40,
                                 'name_font': {'size': 14},
                                 'num_font': {'bold': 1},
                                 'label_position': 'low',
                                })
        for i in range(1, secCount):
            trendFwdChart.add_series({'values': [worksheetName, endRow // 2 - 40, sqrtStart + i, endRow // 2, sqrtStart + i],
                                     'categories': [worksheetName, endRow // 2 - 40, 0, endRow // 2, 0],
                                     'name': str(secStart + secSteps * i) + ' V',
                                     'name_font': {'bold': 1},
                                     'line': {'dash_type': 'round_dot'},
                                     'min': priStop,
                                     'marker': {'type': 'circle'},
                                     'trendline': {'type': 'linear',
                                                   'display_equation': True,
                                                   'name': 'L' + str(secStart + secSteps * i),
                                                   },
                                     })
        worksheet.insert_chart(xl_rowcol_to_cell(1, col), trendFwdChart)

        # Trend line rvs
        trendRvsChart = workbook.add_chart({'type': 'scatter'})
        trendRvsChart.set_size({'width': 540,
                               'height': 480,
                               })
        trendRvsChart.set_plotarea({'layout': {'x': 0.17,
                                              'y': 0.1,
                                              'width': 0.63,
                                              'height': 0.73
                                              }
                                   })
        trendRvsChart.set_legend({'font': {'bold': 1, 'size': 14}})
        trendRvsChart.set_title({'name': workbookName + ' ' + worksheetName + ' RVS VTH'})
        trendRvsChart.set_y_axis({'name': 'SQRT ABS IDRAIN (A)',
                                 'label_position': 'high',
                                 'num_format': '#.#0E-0#',
                                 'num_font': {'bold': 1},
                                 'name_font': {'size': 14},
                                 'name_layout': {'x': 0, 'y': 0.4},
                                 })

        trendRvsChart.set_x_axis({'name': 'VGATE (V)',
                                 'reverse': True,
                                 'major_gridlines': {'visible': True},
                                 'min': priStop,
                                 'max': priStop + 40,
                                 'name_font': {'size': 14},
                                 'num_font': {'bold': 1},
                                 'label_position': 'low',
                                })
        for i in range(1, secCount):
            trendRvsChart.add_series({'values': [worksheetName, endRow // 2 + 1, sqrtStart + i, endRow // 2 + 41, sqrtStart + i],
                                     'categories': [worksheetName, endRow // 2 + 1, 0, endRow // 2 + 41, 0],
                                     'name': str(secStart + secSteps * i) + ' V',
                                     'name_font': {'bold': 1},
                                     'line': {'dash_type': 'round_dot'},
                                     'min': priStop,
                                     'marker': {'type': 'circle'},
                                     'trendline': {'type': 'linear',
                                                   'display_equation': True,
                                                   'name': 'L' + str(secStart + secSteps * i) ,
                                                   },
                                     })
        worksheet.insert_chart(xl_rowcol_to_cell(26, col), trendRvsChart)

        # init variables for calc x-intercept
        xFwd = []
        for i in range(41):
            xFwd.append(-60 - i)
        xRvs = []
        for i in range(41):
            xRvs.append(-100 + i)
        mFwd = []
        bFwd = []
        mRvs = []
        bRvs = []
        xInterFwd = []
        xInterRvs = []

        for num in range(5):
            curFwdY = [None] * 41
            curRvsY = [None] * 41
            for i in range(41):
                curFwdY[i] = pow(abs(y_fwd[num * 41 + i]), 0.5)
                curRvsY[i] = pow(abs(y_rvs[num * 41 + i]), 0.5)
            slopeFwd, interceptFwd, r_valueFwd, p_valueFwd, std_errFwd = stats.linregress(xFwd, curFwdY)
            slopeRvs, interceptRvs, r_valueRvs, p_valueRvs, std_errRvs = stats.linregress(xRvs, curRvsY)
            mFwd.append(slopeFwd)
            bFwd.append(interceptFwd * -1)
            mRvs.append(slopeRvs)
            bRvs.append(interceptRvs * -1)
            xInterFwd.append(interceptFwd * -1 / slopeFwd)
            xInterRvs.append(interceptRvs * -1 / slopeRvs)

        # Create intercept chart
        col += 9
        for i in range(1, secCount):
            worksheet.write(i, col, 'Vd ' + str(secStart + secSteps * i))
        col += 1
        worksheet.write(0, col, 'm FWD')
        for i in range(5):
            worksheet.write(i + 1, col, mFwd[i])
        col += 1
        worksheet.write(0, col, 'b FWD')
        for i in range(5):
            worksheet.write(i + 1, col, bFwd[i])
        col += 1
        fVth = col
        worksheet.write(0, col, 'VTH FWD')
        for i in range(1, secCount):
            worksheet.write(i, col, xInterFwd[i - 1])
        col += 1
        worksheet.write(0, col, 'm RVS')
        for i in range(5):
            worksheet.write(i + 1, col, mRvs[i])
        col += 1
        worksheet.write(0, col, 'b RVS')
        for i in range(5):
            worksheet.write(i + 1, col, bRvs[i])
        col += 1
        rVth = col
        worksheet.write(0, col, 'VTH RVS')
        for i in range(1, secCount):
            worksheet.write(i, col, xInterRvs[i - 1])

        # dId/dVg
        col += 2
        dIdStart = col
        for i in range(1, secCount):
            worksheet.write(0, col, "dId/dVg " + str(secStart + secSteps *i))
            for j in range(1, endRow - 1):
                worksheet.write_formula(j, col, '=LINEST(' + xl_rowcol_to_cell(j, 4 + i) + ':' + xl_rowcol_to_cell(j + 2, 4 + i) + ',A' + str(j + 1) + ':A' + str(j + 3) + ')')
            col += 1

        # dSQId/dVg
        col += 1
        dSQIdStart = col
        for i in range(1, secCount):
            worksheet.write(0, col, "dSQId/dVg " + str(secStart + secSteps *i))
            for j in range(1, endRow - 1):
                worksheet.write_formula(j, col, '=LINEST(' + xl_rowcol_to_cell(j, sqrtStart + i) + ':' + xl_rowcol_to_cell(j + 2, sqrtStart + i) + ',A' + str(j + 1) + ':A' + str(j + 3) + ')')
            col += 1

        # Linear Mobility
        col += 1
        worksheet.write(0, col, "Linear Mobility")
        col += 1
        linMob = col
        for i in range(1, secCount):
            worksheet.write(0, col, "lmob " + str(secStart + secSteps * i))
            for j in range(1, endRow - 1):
                worksheet.write_formula(j, col, '=(' + xl_rowcol_to_cell(j, dIdStart + i - 1) + ')/(' + str(abs(secStart + secSteps * i)) + '*' + str(wlRatio * cap) + ')')
            col += 1

        # Sat Mobility
        col += 1
        worksheet.write(0, col, "Sat Mobility")
        col += 1
        satMob = col
        for i in range(1, secCount):
            worksheet.write(0, col, "smob " + str(secStart + secSteps * i))
            for j in range(1, endRow - 1):
                worksheet.write_formula(j, col, '=(2*(' + xl_rowcol_to_cell(j, dSQIdStart + i - 1) + ')^2)/(' + str(wlRatio * cap) + ')')
            col += 1

        # Combined mobilities chart: 0-Vth is sat and Vth + 1 - -100 is lin
        col += 1
        worksheet.write(0, col, "Combo Mobility")
        col+= 1
        combMob = col
        for i in range(1, secCount):
            worksheet.write(0, col, "mob " + str(secStart + secSteps * i))
            curDivPoint = int(round(xInterRvs[i - 1])) + secStart + secSteps * i
            for j in range(1, abs(curDivPoint) + 2):
                if(j <= 99):
                    worksheet.write_formula(j, col, '=' + xl_rowcol_to_cell(j, satMob + i - 1))
            if curDivPoint < 0:
                for j in range(abs(curDivPoint) + 2, 100):
                    worksheet.write_formula(j, col, '=' + xl_rowcol_to_cell(j, linMob + i - 1))
            col += 1

        # Mobility graph
        col += 1
        mobChart = workbook.add_chart({'type': 'scatter'})
        mobChart.set_size({'width': 680,
                               'height': 480,
                               })
        mobChart.set_plotarea({'layout': {'x': 0.17,
                                              'y': 0.1,
                                              'width': 0.63,
                                              'height': 0.73
                                              }
                                   })
        mobChart.set_legend({'font': {'bold': 1, 'size': 14}})
        mobChart.set_title({'name': workbookName + ' ' + worksheetName + ' MOBILITY'})
        mobChart.set_y_axis({'name': 'Mobility (cm^2/Vs)',
                                 'label_position': 'high',
                                 'num_format': '#.#',
                                 'num_font': {'bold': 1},
                                 'name_font': {'size': 14},
                                 'name_layout': {'x': 0, 'y': 0.4},
                                 })

        mobChart.set_x_axis({'name': 'VGATE (V)',
                                 'reverse': True,
                                 'major_gridlines': {'visible': True},
                                 'min': priStop,
                                 'max': 0,
                                 'name_font': {'size': 14},
                                 'num_font': {'bold': 1},
                                 'label_position': 'low',
                                })
        for i in range(1, secCount):
            curDivPoint = abs(int(round(xInterRvs[i - 1])) + secStart + secSteps * i)
            mobChart.add_series({'values': [worksheetName, 1, combMob + i - 1, curDivPoint + 2, combMob + i - 1],
                                     'categories': [worksheetName, 1, 0, curDivPoint + 2, 0],
                                     'name': 'Sat' + str(secStart + secSteps * i) + ' V',
                                     'name_font': {'bold': 1},
                                     'line': {'dash_type': 'round_dot'},
                                     'min': priStop,
                                     'marker': {'type': 'circle'},
                                     })
            mobChart.add_series({'values': [worksheetName, curDivPoint + 2, combMob + i - 1, 99, combMob + i - 1],
                                     'categories': [worksheetName, curDivPoint + 2, 0, 99, 0],
                                     'name': 'Lin' + str(secStart + secSteps * i) + ' V',
                                     'name_font': {'bold': 1},
                                     'line': {'dash_type': 'solid'},
                                     'min': priStop,
                                     'marker': {'type': 'square'},
                                     })
        worksheet.insert_chart(xl_rowcol_to_cell(1, col), mobChart)

        # Do IDVD stuff with reverse Vth -100

        # Mob Factor
        col = startIDVD
        col += 9
        idvdWorksheets[curWS].write(0, col, "Mob Factor")
        col += 1
        factorStart = col
        for i in range(1, secCount):
            idvdWorksheets[curWS].write(0, col, "F " + str(secStart + secSteps * i))
            for j in range(1, endRow - 2):
                idvdWorksheets[curWS].write_formula(j, col, '=1/((' + str(secStart + secSteps * i) + '*A' + str(j + 1) + ')-(' + str(xInterRvs[4]) + '*A' + str(j + 1) + ')-((A' + str(j + 1) + ')^2/2))')
            col += 1

        # Linear Mobility
        col += 1
        idvdWorksheets[curWS].write(0, col, "Linear Mobility")
        col += 1
        linMob = col
        for i in range(1, secCount):
            idvdWorksheets[curWS].write(0, col, "lmob " + str(secStart + secSteps * i))
            for j in range(1, endRow - 2):
                idvdWorksheets[curWS].write_formula(j, col, '=(' + xl_rowcol_to_cell(j, absStart + i) + '*' + xl_rowcol_to_cell(j, factorStart + i - 1) + ')/(' + str(wlRatio * cap) + ')')
            col += 1

        # Sat Mobility
        col += 1
        idvdWorksheets[curWS].write(0, col, "Sat Mobility")
        col += 1
        satMob = col
        for i in range(1, secCount):
            idvdWorksheets[curWS].write(0, col, "smob " + str(secStart + secSteps * i))
            for j in range(1, endRow - 2):
                idvdWorksheets[curWS].write_formula(j, col, '=((2*' + xl_rowcol_to_cell(j, absStart + i) + ')/((' + str(wlRatio * cap) + ')*(' + str(secStart + secSteps * i) + '-' + str(xInterRvs[4]) +')^2))')
            col += 1

        # Combined mobilities chart
        col += 1
        idvdWorksheets[curWS].write(0, col, "Combo Mobility")
        col += 1
        combMob = col
        for i in range(1, secCount):
            idvdWorksheets[curWS].write(0, col, "mob " + str(secStart + secSteps * i))
            curDivPoint = int(round(secStart + secSteps * i - xInterRvs[4]))
            if curDivPoint < 0:
                for j in range(1, abs(curDivPoint) + 2):
                    idvdWorksheets[curWS].write_formula(j, col, '=' + xl_rowcol_to_cell(j, linMob + i - 1))
                curDivPoint = abs(curDivPoint)
            else:
                curDivPoint = -1
            for j in range(curDivPoint + 2, 100):
                idvdWorksheets[curWS].write_formula(j, col, '=' + xl_rowcol_to_cell(j, satMob + i - 1))
            col += 1

        # Mobility graph
        col += 1
        mobChart = workbook.add_chart({'type': 'scatter'})
        mobChart.set_size({'width': 680,
                           'height': 480,
                           })
        mobChart.set_plotarea({'layout': {'x': 0.17,
                                          'y': 0.1,
                                          'width': 0.63,
                                          'height': 0.73
                                          }
                               })
        mobChart.set_legend({'font': {'bold': 1, 'size': 14}})
        mobChart.set_title({'name': workbookName + ' ' + idvdWorksheets[curWS].get_name() + ' MOBILITY'})
        mobChart.set_y_axis({'name': 'Mobility (cm^2/Vs)',
                             'label_position': 'high',
                             'num_format': '#.#',
                             'num_font': {'bold': 1},
                             'name_font': {'size': 14},
                             'name_layout': {'x': 0, 'y': 0.4},
                             'max': 20
                             })

        mobChart.set_x_axis({'name': 'VGATE (V)',
                             'reverse': True,
                             'major_gridlines': {'visible': True},
                             'min': priStop,
                             'max': 0,
                             'name_font': {'size': 14},
                             'num_font': {'bold': 1},
                             'label_position': 'low',
                             })
        for i in range(1, secCount):
            curDivPoint = int(round(secStart + secSteps * i - xInterRvs[4]))
            if curDivPoint < 0:
                mobChart.add_series({'values': [idvdWorksheets[curWS].get_name(), 1, combMob + i - 1, abs(curDivPoint) + 2, combMob + i - 1],
                                     'categories': [idvdWorksheets[curWS].get_name(), 1, 0, abs(curDivPoint) + 2, 0],
                                     'name': 'Lin' + str(secStart + secSteps * i) + ' V',
                                     'name_font': {'bold': 1},
                                     'line': {'dash_type': 'round_dot'},
                                     'min': priStop,
                                     'marker': {'type': 'circle'},
                                     })
                curDivPoint = abs(curDivPoint)
            else:
                curDivPoint = -1
            mobChart.add_series({'values': [idvdWorksheets[curWS].get_name(), curDivPoint + 2, combMob + i - 1, 99,
                                            combMob + i - 1],
                                 'categories': [idvdWorksheets[curWS].get_name(), curDivPoint + 2, 0, 99, 0],
                                 'name': 'Sat' + str(secStart + secSteps * i) + ' V',
                                 'name_font': {'bold': 1},
                                 'line': {'dash_type': 'solid'},
                                 'min': priStop,
                                 'marker': {'type': 'square'},
                                 })
        idvdWorksheets[curWS].insert_chart(xl_rowcol_to_cell(1, col), mobChart)

        curWS += 1

    else:
        idvdWorksheets.append(worksheet)

workbook.close()
