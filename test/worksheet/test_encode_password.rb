# -*- coding: utf-8 -*-

require 'helper'

class TestWorksheetEncodePassword < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  ###############################################################################
  #
  # Tests for WriteXLSX::Worksheet methods.
  #
  #
  def test_encode_password
    tests.each do |test|
      password, expected = test

      assert_equal(
        expected,
        @worksheet.__send__(:encode_password, password),
        "Password '#{password}' failed."
      )
    end
  end

  def tests
    [
      ["password",                        "83AF"],
      ["This is a longer phrase",         "D14E"],
      ["0",                               "CE2A"],
      ["01",                              "CEED"],
      ["012",                             "CF7C"],
      ["0123",                            "CC4B"],
      ["01234",                           "CACA"],
      ["012345",                          "C789"],
      ["0123456",                         "DC88"],
      ["01234567",                        "EB87"],
      ["012345678",                       "9B86"],
      ["0123456789",                      "FF84"],
      ["01234567890",                     "FF86"],
      ["012345678901",                    "EF87"],
      ["0123456789012",                   "AF8A"],
      ["01234567890123",                  "EF90"],
      ["012345678901234",                 "EFA5"],
      ["0123456789012345",                "EFD0"],
      ["01234567890123456",               "EF09"],
      ["012345678901234567",              "EEB2"],
      ["0123456789012345678",             "ED33"],
      ["01234567890123456789",            "EA14"],
      ["012345678901234567890",           "E615"],
      ["0123456789012345678901",          "FE96"],
      ["01234567890123456789012",         "CC97"],
      ["012345678901234567890123",        "AA98"],
      ["0123456789012345678901234",       "FA98"],
      ["01234567890123456789012345",      "D298"],
      ["0123456789012345678901234567890", "D2D3"]
    ]
  end
end
