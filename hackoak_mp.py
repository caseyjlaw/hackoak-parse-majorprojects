# script to parse and visualize Oakland major projects spreadsheet

import openpyxl

classlist = [ss.lower() for ss in ['Commercial, Industrial, and Civic Projects', 'Mixed-Use Projects', 'Residential Projects']]
statelist = [ss.lower() for ss in ['Application Approved','Projects Under Construction','Under Construction','Application Submitted-Under Review','Pre-Application Discussions']]

def parse(sheetname):
    """ Parse spreadsheet
    """

    wb = openpyxl.load_workbook(sheetname)
    sheets = wb.sheetnames

    d = {}
    for sheet in sheets:
        sh = wb[sheet]
        good = 0; failed = 0
        for row in sh.rows:
            # first define class/state for following rows
            fil = filter(lambda cell: cell.value and type(cell.value) is unicode, row)   # remove Nones and non-unicode
            clfil = filter(lambda cell: cell.value.encode('ascii', 'ignore').lower().strip() in classlist, fil)  # need to convert unicode to string correctly. lowercase and remove trailing empty spaces to match to template
            stfil = filter(lambda cell: cell.value.encode('ascii', 'ignore').lower().strip() in statelist, fil)
            try:
                cl = clfil[0].value.lower()  # define current class of project
            except:
                pass
            try:
                st = stfil[0].value.lower()  # define current state of project
            except:
                pass

            try:
                key = st + ' --- ' + cl    # define keys that are available
                if not d.has_key(key):
                    d[key] = {}  # fill out dict fields
                    d[key]['dist1'] = 0
                    d[key]['dist2'] = 0
                    d[key]['dist3'] = 0
                    d[key]['dist4'] = 0
                    d[key]['dist5'] = 0
                    d[key]['dist6'] = 0
                    d[key]['dist7'] = 0
                    d[key]['total'] = 0
            except:
                continue      # some early rows don't have complete key

            if type(row[0].value) == int:   # this gives us a data row, which is indexed with an int cell
                if 'units' in row[5].value.encode('ascii', 'ignore').lower():   # if this cell refers to "units"
                    try:
                        units = getunits(row[5])
                    except:
                        print
                        print row[5].value
                        try:
                            units = int(input('*** Auto parse failed. How many units above? ***'))
                        except SyntaxError:
                            pass
                        failed += 1

                try:
                    dist = getdist(row[4])
                except:
                    print
                    print row[4].value
                    try:
                        dist = int(input('*** Auto parse failed. What district above? ***'))
                    except SyntaxError:
                        pass
                    failed += 1

                if dist and units:
                    good += 1
                    d[key]['total'] += units
                    d[key]['dist'+str(dist)] += units
        print 'Autoparsed %d rows and %d manually.' % (good, failed)
    return d

def getunits(cell):
    """ Takes description cell value
    Tries to parse it to get residential unit count.
    """

    res = cell.value.encode('ascii', 'ignore').lower()   # cast description field to string

    # get residential unit count. known to skip lots of cases
    for bullet in res.split('\n'):   # bullet cast as carriage return
        words = bullet.split(' ')
        if 'units' in words:
            for loc in range(words.index('units'), -1, -1):
                try:
                    units = int(words[loc].lstrip('n').rstrip())
                except ValueError:
                    pass
                else:
                    print 'got one:', units
                    break
#        units = int(bullet.split('residential')[0].lstrip('n').lstrip('n').rstrip().rstrip('n'))  # remove a few things from ends

    return int(units)

def getdist(cell):
    """ Parse cell for district info. Can be multiple values, but we just take first one.
    """

    if type(cell.value) == unicode:
        cell2 = cell.value.encode('ascii', 'ignore')
        try:
            dist = int(cell2.split(' and ')[0])
        except:
            dist = int(cell2.split(' & ')[0])
    else:
        dist = int(cell.value)

    return dist

def vis(d):
    """ Visualize distribution of housing of all types in each district
    """
    pass
