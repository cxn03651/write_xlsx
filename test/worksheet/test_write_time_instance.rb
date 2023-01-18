# -*- coding: utf-8 -*-

require 'helper'

class TestWriteTimeInstance < Minitest::Test
  def assert_nothing_raised
    yield
  rescue StandardError => e
    raise Minitest::UnexpectedError.new(e)
  end

  def test_write_time_instance_nothing_raised
    workbook  = WriteXLSX.new(StringIO.new)
    worksheet = workbook.add_worksheet

    assert_nothing_raised do
      worksheet.write('A1', Time.new)
    rescue NoMethodError => e
      assert(false, "Failed! : Worksheet#write should not raises with Time instance token.")
    ensure
      workbook.close
    end
  end
end
