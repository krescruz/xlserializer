from serializers import XlSerializerBase

serializer = XlSerializerBase(path="workbook.xlsx", declared_columns=('NumEmp','Nombre', ))

serializer.set_sheet("NameSheet", idx_cols=7, idx_data=9)

#print(serializer.column_names)
print(serializer.data)
