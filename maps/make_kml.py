
import xlrd

kml = '<?xml version="1.0" encoding="UTF-8"?>\n \
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">\n \
<Document>'

workbook = xlrd.open_workbook('collections.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
old_species = ''
for row in range(worksheet.nrows):
    if row != 0 and worksheet.cell_value(row, 6) == 'Chylismia':
        collection_id = worksheet.cell_value(row, 0)
        species = worksheet.cell_value(row, 6) + ' ' + worksheet.cell_value(row, 7) + ' ssp. ' + worksheet.cell_value(row, 8)
        latitude = worksheet.cell_value(row, 9)
        longitude = worksheet.cell_value(row, 10)
        if species != old_species:
            if old_species != '':
                kml = kml + '</Folder>\n'
            kml = kml + '<Folder><name>' + species + '</name>\n'

        kml = kml + '<Placemark>\n<name>' + str(collection_id) + '</name>\n'
        kml = kml + '<description>' + species + '</description>\n'
        kml = kml + '<Point><coordinates>-' + str(longitude) + ',' + str(latitude) + ',0</coordinates></Point>\n</Placemark>\n'
        old_species = species

kml = kml + '</Folder>\n</Document></kml>'

kml_file = open("chylismia_collections.kml", "w")
kml_file.write(kml)


"""
# iterate over rows
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = -1
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    print 'Row:', curr_row
    curr_cell = -1
    # iterate through cells in row
    while curr_cell < num_cells:
        curr_cell += 1
        # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
        cell_type = worksheet.cell_type(curr_row, curr_cell)
        cell_value = worksheet.cell_value(curr_row, curr_cell)
        print ' ', cell_type, ':', cell_value
"""



