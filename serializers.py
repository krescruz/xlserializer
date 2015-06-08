import xlrd
from xlrd.biffh import XLRDError
import datetime


class BaseXlSerializer(object):	
	
	workbook = None
	_data = []

	def __init__(self, filename=None):
		self._open(filename)

	def _open(self, filename):
		self.workbook = xlrd.open_workbook(filename)

	@property
	def sheet_names(self):
		 return self.workbook.sheet_names()

	def set_sheet(self, sheet_name, declared_columns=None, init_row_header=0, init_row_data=1):
		self.declared_columns = declared_columns
		self.init_row_header = init_row_header
		self.init_row_data = init_row_data
		try:
			self.worksheet = self.workbook.sheet_by_name(sheet_name)
			self._data = self.get_data_rows()
		except XLRDError:
			self.worksheet = None
		return self.worksheet

	@property
	def column_names(self):
		assert self.worksheet is not None, '`worksheet` sheet no valid.'
		row = self.worksheet.row(self.init_row_header)
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
			cells = self._col_slice(col_index, start_rowx=self.init_row_data)
			data_cols.append(self.to_internal_value(cells))
		return list(zip(*data_cols))

	def to_internal_value(self, rows):
		values = []
		for cell in rows:
			internal_value = None
			if cell.ctype is xlrd.XL_CELL_DATE:
				xldate_as_tuple = xlrd.xldate_as_tuple(cell.value, self.workbook.datemode)
				internal_value = datetime.datetime(*xldate_as_tuple)
			else:
				internal_value = cell.value

			values.append(internal_value)

		return values

	@property
	def data(self):
		return self._data
