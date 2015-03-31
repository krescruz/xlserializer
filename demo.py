from serializers import XlSerializerBase

serializer = XlSerializerBase(path="workbook.xlsx")
serializer.set_sheet("NameSheet")

print(serializer.column_names)
print(serializer.data)