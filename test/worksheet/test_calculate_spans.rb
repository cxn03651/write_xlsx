# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

#
# class for test_calculate_spans
#
class CalcSpansTC
  attr_reader :row, :col, :expected
  def initialize(row, col, expected)
    @row = row
    @col = col
    @expected = expected
  end
end

class TestCalculateSpans < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_calculate_spans
    [
      CalcSpansTC.new( 0, 0, [ '1:16', '17:17' ]),
      CalcSpansTC.new( 1, 0, [ '1:15', '16:17' ]),
      CalcSpansTC.new( 2, 0, [ '1:14', '15:17' ]),
      CalcSpansTC.new( 3, 0, [ '1:13', '14:17' ]),
      CalcSpansTC.new( 4, 0, [ '1:12', '13:17' ]),
      CalcSpansTC.new( 5, 0, [ '1:11', '12:17' ]),
      CalcSpansTC.new( 6, 0, [ '1:10', '11:17' ]),
      CalcSpansTC.new( 7, 0, [ '1:9',  '10:17' ]),
      CalcSpansTC.new( 8, 0, [ '1:8',   '9:17' ]),
      CalcSpansTC.new( 9, 0, [ '1:7',   '8:17' ]),
      CalcSpansTC.new(10, 0, [ '1:6',   '7:17' ]),
      CalcSpansTC.new(11, 0, [ '1:5',   '6:17' ]),
      CalcSpansTC.new(12, 0, [ '1:4',   '5:17' ]),
      CalcSpansTC.new(13, 0, [ '1:3',   '4:17' ]),
      CalcSpansTC.new(14, 0, [ '1:2',   '3:17' ]),
      CalcSpansTC.new(15, 0, [ '1:1',   '2:17' ]),
      CalcSpansTC.new(16, 0, [ nil, '1:16', '17:17' ]),
      CalcSpansTC.new(16, 1, [ nil, '2:17', '18:18' ])
    ].each do |t|
      worksheet = @workbook.add_worksheet('')
      r   = t.row
      col = t.col
      (r .. r + 16).each do |row|
        worksheet.write(row, col, 1)
        col += 1
      end
      worksheet.__send__('calculate_spans')
      result = worksheet.instance_variable_get(:@row_spans)
      expected = t.expected
      assert_equal(expected, result, "WHEN row: #{t.row}, col: #{t.col}")
    end
  end
end
