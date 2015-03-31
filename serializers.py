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
		data_rows = {}
		column_names = self.declared_columns if self.declared_columns else self.column_names
		for name in column_names:
			col_index = self._get_index_col(name)
			data_rows[name] = self._col_slice(col_index)
		return data_rows

	@property
	def data(self):
		return self.get_data_rows()
