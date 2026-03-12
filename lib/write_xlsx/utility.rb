# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'write_xlsx/utility/common'
require 'write_xlsx/utility/cell_reference'
require 'write_xlsx/utility/sheetname_quoting'
require 'write_xlsx/utility/dimensions'
require 'write_xlsx/utility/string_width'
require 'write_xlsx/utility/date_time'
require 'write_xlsx/utility/url'
require 'write_xlsx/utility/xml_primitives'
require 'write_xlsx/utility/drawing'
require 'write_xlsx/utility/chart_formatting'
require 'write_xlsx/utility/rich_text'

module Writexlsx
  module Utility
    include Common
    include CellReference
    include SheetnameQuoting
    include Dimensions
    include StringWidth
    include DateTime
    include Url
    include XmlPrimitives
    include Drawing
    include ChartFormatting
    include RichText

    ROW_MAX       = 1048576  # :nodoc:
    COL_MAX       = 16384    # :nodoc:
    STR_MAX       = 32767    # :nodoc:
    SHEETNAME_MAX = 31       # :nodoc:
  end
end
