from __future__ import print_function as pf
from flask import Flask, render_template, request, redirect
import gspread # googlesheets api python library
import  datetime as dt
import xlsxwriter
import os

app = Flask(__name__)
gc = gspread.service_account(filename='/etc/secrets/keys.json') # authenticate with Google Sheets

def column(array, column):
    return [row[column] for row in array]

def error(error):
    return render_template('error.html', error=error)

def list_reverse_index(li, x):
    for i in reversed(range(len(li))):
        if li[i] == x:
            return i
    raise ValueError("{} is not in list".format(x))

def int_to_char(int):
    char = xlsxwriter.utility.xl_col_to_name(int)
    return char

@app.route('/healthcheck')
def health_check():
    return ''


@app.route('/') # the index page
def index(): # retireve list of headers to populate name options list    
    sheetInstance = gc.open_by_key(os.getenv(SHEET_KEY)) # open sheet via it's key, found in th URL
    namesRowAsList = sheetInstance.sheet1.row_values(1) # index starts on 1. get header row. Only the top left cell of merged cells has the value of any of the merged cells.
    namesNoBlanks = sorted(filter(bool, namesRowAsList)) # remove blank items
    return render_template('index.html', names=namesNoBlanks, bool=bool) # this is where vartiables and functions are passed into the template for use in jinja


@app.route('/<string:name>', methods=['POST', 'GET'])
def name(name):
    # GET AND PREP SHEET VARS
    gcSheetObject = gc.open_by_key(os.getenv(SHEET_KEY)) # open sheet
    sheet1 = gcSheetObject.sheet1.get_all_values() # list of lists      
    headerRow = gcSheetObject.sheet1.row_values(1) # list of header row cells' contents
    userLeftColumnIndex = headerRow.index(name)
    datesColumn = gcSheetObject.sheet1.col_values(1)# sheet1() index starts on 1. This py list starts on 0. find today's row, make one if not exist. This references live page. mutable?
    todayRowIndex = None
    todaysStarting = None
    todaysEnding = None
    emptyTable = False
    sheetStartingPeriodDayDT = dt.datetime.strptime(datesColumn[2], '%a, %b %d, %Y') # all other appends must be in this format
    periodStartingDays = [sheetStartingPeriodDayDT]
    todayDT = dt.datetime.now() # THIS MUST BE FOR THE DAY, NOT THE ACCURACY .NOW GIVES
    
    def refresh_table__after_total_appending():
        sheet1 = gcSheetObject.sheet1.get_all_values()
        datesColumn = gcSheetObject.sheet1.col_values(1)
        return sheet1, datesColumn
    
    def sum_row(index):
        if "Total" not in datesColumn:
            topOfPeriodIndex = 3 # this is for gs so it isnt starting on 0
        else:
            topOfPeriodIndex = list_reverse_index(datesColumn, "Total") + 2
        botOfPeriodIndex = len(datesColumn)
        sum = '=sum(' + int_to_char(index+1) + str(topOfPeriodIndex) + ':' + int_to_char(index+1) + str(botOfPeriodIndex) + ') - sum(' + int_to_char(index) + str(topOfPeriodIndex) + ':' + int_to_char(index) + str(botOfPeriodIndex) + ')'
        return sum
    
    def append_total_row():
        gcSheetObject.sheet1.update_cell(len(datesColumn) + 1, 1, "Total") # append total to column 1
        driverIndexes = []
        for each in headerRow:
            if bool(each):
                driverIndexes.append(headerRow.index(each))
        for index in driverIndexes:
            gcSheetObject.sheet1.update_cell(len(datesColumn) + 1, index + 1, sum_row(index))
    
    # MAKE A TOTAL ROW IF ONE IS NEEDED
    try:
        if datesColumn[-1] != 'Total' and datesColumn[-1] != todayDT.strftime('%a, %b %d, %Y'): # if last row is not a total or today
            # GET LIST OF ALL NEW PERIOD DAYS BETWEEN TODAY AND THE FIRST DAY ON SHEET, INCLUSIVE
            for i in range(1, 200): # should last about 8 yrs of logging. starting on 1 because periodStartingDays is assigned a value earlier
                day = sheetStartingPeriodDayDT + dt.timedelta(14*i) # pick days 14 days apart. i starts at 0 so first day indexed is the row 3 day        
                if day.strftime('%b %d, %Y') <= todayDT.strftime('%b %d, %Y'): # this prevents it from going all the way to 200
                    periodStartingDays.append(day)
                else:
                    break
            if todayDT.strftime('%b %d, %Y') == periodStartingDays[-1].strftime('%b %d, %Y') or datesColumn[-1] == periodStartingDays[-1].strftime('%b %d, %Y'): # if today or last day is a period starting day.
                append_total_row()
                sheet1, datesColumn = refresh_table__after_total_appending()
            elif dt.datetime.strptime(datesColumn[-1], '%a, %b %d, %Y') <= periodStartingDays[-1]:
                append_total_row()
                sheet1, datesColumn = refresh_table__after_total_appending()
    except Exception as e:
        return error("Error: " + e)
    
    # SET CURRENT DAY INFO INJECT INTO SUB BOX
    try:
        todayRowIndex = datesColumn.index(dt.datetime.now().strftime('%a, %b %d, %Y'))
        #print('today exists at ' + str(todayRowIndex))
        todaysStarting = sheet1[todayRowIndex][userLeftColumnIndex]        
        #print('starting column index: ' + str(userLeftColumnIndex))
        todaysEnding = sheet1[todayRowIndex][userLeftColumnIndex + 1]
        #print('starting mileage: ' + str(todaysStarting))
        #print('ending mileage: ' + str(todaysEnding))
    except:    
        #print('todaysStarting or todaysEnding or todayRowIndex failed assignment')
        todaysStarting = 'Start'
        todaysEnding = 'End'
    if todaysStarting == '':
        todaysStarting = 'Start'
    if todaysEnding == '':
        todaysEnding = 'End'
    
    sheet1 = sheet1[2:] # keep all but the first two rows
    dates = column(sheet1, 0)
    userStartColumn = column(sheet1, userLeftColumnIndex)
    userEndColumn = column(sheet1, userLeftColumnIndex + 1)
    
    # MAKE MILEAGE SUM
    mileageSum = []
    for start, end in zip(userStartColumn, userEndColumn):  
        if bool(start) and bool(end): # do the difference if both mileages exist
                mileageSum.append(int(end) - int(start))
        else:
            mileageSum.append('')
    # MAKE TABLE
    table = []
    table.append(mileageSum)
    table.append(userEndColumn)
    table.append(userStartColumn)
    table.append(dates)
    table = list(zip(*reversed(table))) # rotate 90 right

    if len(table) == 0:
        emptyTable = True
        print('Table is empty')
    
    # SUBMISSION HANDLING
    if request.method == 'POST': # action for submit button being hit
        print('Submitting mileage...')
        startMileage = request.form['startMileage']
        endMileage = request.form['endMileage']
        print(endMileage)
        todayDateStr = dt.datetime.now().strftime('%a, %b %d, %Y')
        try:
            todayIsIn = datesColumn[-1] == todayDateStr
        except:
            todayIsIn = False
        if emptyTable and (bool(startMileage) or bool(endMileage)): # if table is empty and they have something to submit. This should never happen with the new requirement to start new sheets with the first day of the period.
            print('table is empty and they have something to submit')
            gcSheetObject.sheet1.append_row([todayDateStr], table_range='A3:A999') # append today's date to first available A row
            print("Appended today's date")
            # if there is nothing in the sheet today's row must be 3
            try:
                gcSheetObject.sheet1.update_cell(3, userLeftColumnIndex + 1, startMileage) # the +1s are becuase this indexing starts on 1
                print('Added starting mileage to row 3, column ' + str(userLeftColumnIndex + 1))
            except:
                return error("ERROR. Starting mileage may not have been entered")
            try:
                gcSheetObject.sheet1.update_cell(3, userLeftColumnIndex + 2, endMileage)
                print('Added ending mileage to row 3, column ' + str(userLeftColumnIndex + 2))
            except:
                return error("ERROR. Starting mileage may not have been entered")
        elif (not todayIsIn) and (bool(startMileage) or bool(endMileage)): # if today is not in, and they are entering something
            print('today is not in, and they are entering something')
            gcSheetObject.sheet1.append_row([todayDateStr], table_range='A3:A999') # append today's date to first available A row
            print("Appended today's date")
            datesColumn = gcSheetObject.sheet1.col_values(1)# refresh dates column after appending. Find today's row, make one if not exist
            todayRowIndex = datesColumn.index(todayDateStr)
            try:
                gcSheetObject.sheet1.update_cell(todayRowIndex + 1, userLeftColumnIndex + 1, startMileage) # the +1s are becuase this indexing starts on 1
            except:
                return error("ERROR. Starting mileage may not have been entered")
            try:
                gcSheetObject.sheet1.update_cell(todayRowIndex + 1, userLeftColumnIndex + 2, endMileage)
            except:
                return error("ERROR. Ending mileage may not have been entered")
        elif not emptyTable and todayIsIn:
            try:
                gcSheetObject.sheet1.update_cell(todayRowIndex + 1, userLeftColumnIndex + 1, startMileage) # the +1s are becuase this indexing starts on 1
            except:
                return error("ERROR. Starting mileage may not have been entered")
            try:
                gcSheetObject.sheet1.update_cell(todayRowIndex + 1, userLeftColumnIndex + 2, endMileage)
            except:
                return error("ERROR. Ending mileage may not have been entered")
        else:
            return error('No information has been entered, and there is no mileage today to be cleared')
        
        
        # APPEND SUBMISSION TO LOGGER SHEET
        submissionLog = gcSheetObject.worksheet('Submission Log')
        submission = [name, dt.datetime.now().strftime('%x, %X'), 'Starting mileage: ' + request.form['startMileage'], 'Ending mileage: ' + request.form['endMileage']]
        submissionLog.append_row(submission) # appended item must be py.list
        
        return redirect(request.referrer)
    
    
    return render_template('timelog.html', name=name, table=table, emptyTable = emptyTable, today=dt.datetime.now().strftime('%a, %b %d, %Y'), todaysStarting=todaysStarting, todaysEnding=todaysEnding)


if __name__ == '__main__': # start the site
    app.run(debug=True)

# SHEET REQUIREMENTS:
    # sheet should be cleared before TBD years of use
    # the submission records sheet must be named 'Submission Log'
    # names must always be in same column as starting mileage, not ending mileage. merged cells have their value on the furthest left cell in the merge.
    # The sheet's first row absolutly must be the first monday of a new period in exactly the regular date format even if it must be typed in, even if no data to enter.
    # black squares in top left must be empty
    # Heroku will end free plans on November 28th, 2022
# PLANNED FEATURES
    # some sort of lockout so people cant change mileage after-hours
    # add admin panel
# ISSUES
    # a number somewhere in a column without a date will cause a submission to fail indexing if there isnt already a date for 'today'
# TESTING
    # can clear time by entering nothing
    # can enter one or the other, no need for both
    # entering nothing on a day without a row does not make a new row
    # entering on a new day creates a new row
    # can enter very first day
    # can enter nothing on first day and not crash everythin
    # can you leave only a signin number yesterday and still add to today?
    # what happend if a column is cleared an not removed?