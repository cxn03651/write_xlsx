# -*- coding: utf-8 -*-

require 'write_xlsx/workbook'

class WriteXLSX < Writexlsx::Workbook
  if RUBY_VERSION < '1.9'
    $KCODE = 'u'
  end
end
