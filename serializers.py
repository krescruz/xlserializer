import xlrd
from xlrd.biffh import XLRDError
import copy


class XlSerializerBase(object):	
	
	workbook = None

	def __init__(self, path=None, init_row=0, sheet_name=(), declared_columns=()):
		self.path = path
		self.init_row = init_row
		self.sheet_name = sheet_name
		self.declared_columns = declared_columns
		self._open()

	def _open(self):
		self.workbook = xlrd.open_workbook(self.path)

	@property
	def sheet_names(self):
		 return self.workbook.sheet_names()

	def set_sheet(self, sheet_name):
		try:
			self.worksheet = self.workbook.sheet_by_name(sheet_name)
			self._data = self.get_data_rows()
		except XLRDError:
			self.worksheet = None
		return self.worksheet

	@property
	def column_names(self):
		assert self.worksheet is not None, '`worksheet` sheet no valid.'
		row = self.worksheet.row(self.init_row)
		return list(column.value for column in row)

	def _get_index_col(self, name):
		return self.column_names.index(name)

	def _col_slice(self, col_index, start_rowx=0, end_rowx=None):
		cols= self.worksheet.col_slice(col_index, start_rowx, end_rowx)
		return cols

	def get_data_rows(self):
		data_cols = []
		column_names = self.declared_columns if self.declared_columns else self.column_names

		for name in column_names:
			col_index = self._get_index_col(name)
			cells = self._col_slice(col_index, start_rowx=self.init_row+1)
			data_cols.append(self.to_internal_value(cells))
		return list(zip(*data_cols))

	def to_internal_value(self, rows):
		values = []
		for cell in rows:
			internal_value = None
			if cell.ctype is xlrd.XL_CELL_DATE:
				internal_value = xlrd.xldate_as_tuple(cell.value, self.workbook.datemode)
			else:
				internal_value = cell.value

			values.append(internal_value)

		return values

	@property
	def data(self):
		return self._data

import datetime
from xlrd import Book
ser = XlSerializerBase(path="workbook.xlsx")
ser.set_sheet("Hoja1")
#print(ser.column_names
print(ser.data[0][1])
print(ser.data[0][2])