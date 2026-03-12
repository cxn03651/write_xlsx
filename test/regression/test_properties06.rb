# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionProperties06 < Minitest::Test
  def setup
    setup_dir_var
    @long_string = ('This is a long string. ' * 11) + 'AA'
  end

  def teardown
    @tempfile.close(true)
  end

  def test_properties06
    @xlsx = 'properties06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    company_guid = "2096f6a2-d2f7-48be-b329-b73aaa526e5d"
    site_id      = "cb46c030-1825-4e81-a295-151c039dbf02"
    action_id    = "88124cf5-1340-457d-90e1-0000a9427c99"

    workbook.set_custom_property("MSIP_Label_#{company_guid}_Enabled",     'true',                 'text')
    workbook.set_custom_property("MSIP_Label_#{company_guid}_SetDate",     '2024-01-01T12:00:00Z', 'text')
    workbook.set_custom_property("MSIP_Label_#{company_guid}_Method",      'Privileged',           'text')
    workbook.set_custom_property("MSIP_Label_#{company_guid}_Name",        'Confidential',         'text')
    workbook.set_custom_property("MSIP_Label_#{company_guid}_SiteId",      site_id,                'text')
    workbook.set_custom_property("MSIP_Label_#{company_guid}_ActionId",    action_id,              'text')
    workbook.set_custom_property("MSIP_Label_#{company_guid}_ContentBits", '2',                    'text')

    workbook.close
    compare_for_regression
  end
end
