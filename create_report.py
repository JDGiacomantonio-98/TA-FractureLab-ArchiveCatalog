import xlrd, docx

def main():
    return generate_report(load_data())


def load_data():
    ''' Retrieve and pack data from the excel database of archive resources '''

    db = xlrd.open_workbook("Lab-of-Fractrure database - python.xlsx").sheet_by_index(0)

    keys = set()
    for r in range(1, db.nrows-1):
        keys.add(db.cell_value(r, 0))
        archive = dict.fromkeys(keys, {})

    keys = []
    for h in range(2, db.ncols):
        keys.append(db.cell_value(0, h))
        # document[db.cell_value(0, h)] = ''
    
    for f in archive.keys():
        index = {}
        for r in range(1, db.nrows-1):
            if db.cell_value(r, 0) != f:
                continue
            document = {}
            for i, h in enumerate(range(2, db.ncols)):
                document[keys[i]] = db.cell_value(r, h)
            index[int(db.cell_value(r, 1))] = document
        archive[f] = index
    
    return archive

def generate_report(arg):
    ''' Generate the requested Word file accordingly to all the authors formatting rules '''

    pass

main()