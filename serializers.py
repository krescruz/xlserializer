import xlrd
from xlrd.biffh import XLRDError


class XlSerializerBase(object):	
	workbook = None
	def __init__(self, path=None, init_row=0):
		self.path = path
		self.init_row = init_row
		self.sheet_name = None
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

	@property
	def data(self):
		num_rows = self.worksheet.nrows - 1
		init_data = self.init_row + 1
		return list(
			self.worksheet.row(row)
			for row in range(init_data, num_rows)
		)
