# xlserializer
Class helper to read file .xls with python.

# Requirements
* Python (3.3)
* xlrd (0.9.3)

# Use

    from serializers import BaseXlSerializer

    serializer = BaseXlSerializer(filename="workbook.xlsx")
    serializer.set_sheet("NameSheet") # open sheet
    
    #prints all the names of columns
    print(serializer.column_names)
    #prints all data
    print(serializer.data)
