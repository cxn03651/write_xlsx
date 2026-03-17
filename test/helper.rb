# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'bundler'

begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  warn e.message
  warn "Run `bundle install` to install missing gems"
  exit e.status_code
end
require 'minitest/autorun'

$LOAD_PATH.unshift(File.dirname(__FILE__))
$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
require 'write_xlsx'
require 'stringio'
require 'tempfile'

class Writexlsx::Workbook
  #
  # Set the default index for each format. This is mainly used for testing.
  #
  def set_default_xf_indices # :nodoc:
    # Delete the default url format.
    @formats.formats.delete_at(1)

    @formats.formats.each do |format|
      format.get_xf_index
    end
  end
end

class Minitest::Test
  def setup_dir_var
    @test_dir = File.dirname(__FILE__)
    @perl_output = File.join(@test_dir, 'perl_output')
    @regression_output = File.join(@test_dir, 'regression', 'xlsx_files')
    @io = StringIO.new(''.dup)
  end

  def expected_xlsx
    File.join(@regression_output, @xlsx)
  end

  def expected_to_array(lines)
    array = []
    lines.each_line do |line|
      str = line.chomp.sub(/^\s+/, '')
      array << str unless str == ''
    end
    array
  end

  def got_to_array(xml_str)
    xml_str.gsub(/[\r\n]/, '')
           .gsub(%r{ +/>}, ' />')
           .gsub(/>[ \t]*</, ">\t<")
           .split("\t")
  end

  def vml_str_to_array(vml_str)
    ret = ''
    vml_str.split(/[\r\n]+/).each do |vml|
      str = vml.sub(/^\s+/, '')              # Remove leading whitespace.
               .sub(/\s+$/, '')              # Remove trailing whitespace.
               .gsub("'", '"')               # Convert VMLs attribute quotes.
               .gsub(%r{([^ ])/>$}, '\1 />') # Add space before element end like XML::Writer.
               .sub(/"$/, '" ')              # Add space between attributes.
               .sub(/>$/, ">\n")             # Add newline after element end.
               .gsub("><", ">\n<")           # Split multiple elements.
      str.chomp! if str == "<x:Anchor>\n"    # Put all of Anchor on one line.
      ret += str
    end
    ret.split("\n")
  end

  def entrys(xlsx)
    result = []
    Zip::File.foreach(xlsx) { |entry| result << entry }
    result
  end

  def compare_for_regression(ignore_members = nil, ignore_elements = nil)
    store_to_tempfile
    compare_xlsx(expected_xlsx, @tempfile.path, ignore_members, ignore_elements, true)
  end

  def store_to_tempfile
    @tempfile = Tempfile.open(@xlsx)
    @tempfile.binmode
    @tempfile.write(@io.string)
    @tempfile.close
  end

  def compare_xlsx_for_regression(exp_filename, got_filename, ignore_members = nil, ignore_elements = nil)
    compare_xlsx(exp_filename, got_filename, ignore_members, ignore_elements, true)
  end

  #
  # Compare two xlsx files by normalizing and comparing each entry.
  #
  def compare_xlsx(exp_filename, got_filename, ignore_members = nil, ignore_elements = nil, regression = false)
    got_members = filtered_xlsx_entries(got_filename, ignore_members)
    exp_members = filtered_xlsx_entries(exp_filename, ignore_members)

    assert_same_member_names(exp_members, got_members)

    exp_members.each_index do |i|
      compare_xlsx_entry(
        exp_members[i],
        got_members[i],
        ignore_elements,
        regression
      )
    end
  end

  #
  # Return sorted xlsx entries and remove ignored members when requested.
  #
  def filtered_xlsx_entries(filename, ignore_members = nil)
    members = entrys(filename).sort_by(&:name)
    return members unless ignore_members

    members.reject { |member| ignore_members.include?(member.name) }
  end

  #
  # Assert that two xlsx containers have the same member names.
  #
  def assert_same_member_names(exp_members, got_members)
    assert_equal(
      exp_members.map(&:name),
      got_members.map(&:name),
      'file members differs.'
    )
  end

  #
  # Compare a single entry in two xlsx containers.
  #
  def compare_xlsx_entry(exp_member, got_member, ignore_elements, regression)
    got_xml_str = normalized_member_xml(got_member, regression)
    exp_xml_str = normalized_member_xml(exp_member, regression, expected: true)

    got_xml = xml_to_comparable_array(got_xml_str, got_member.name)
    exp_xml = xml_to_comparable_array(exp_xml_str, exp_member.name)

    if ignore_elements && ignore_elements[exp_member.name]
      got_xml = reject_ignored_elements(got_xml, ignore_elements[exp_member.name])
      exp_xml = reject_ignored_elements(exp_xml, ignore_elements[exp_member.name])
    end

    case exp_member.name
    when '[Content_Types].xml', /.rels$/
      got_xml = sort_rel_file_data(got_xml)
      exp_xml = sort_rel_file_data(exp_xml)
    end

    assert_equal(exp_xml, got_xml, "#{exp_member.name} differs.")
  end

  #
  # Read and normalize a single xlsx entry.
  #
  def normalized_member_xml(member, regression = false, expected: false)
    raw = member.get_input_stream.read
    xml = normalize_xml_empty_tags(raw)

    if member.name == 'docProps/core.xml'
      xml = normalize_core_xml(xml, regression, expected)
    end

    if member.name == 'xl/workbook.xml'
      xml = normalize_workbook_xml(xml)
    end

    if member.name =~ %r{xl/worksheets/sheet\d.xml}
      xml = normalize_worksheet_xml(xml, expected)
    end

    if member.name =~ %r{xl/charts/chart\d.xml}
      xml = normalize_chart_xml(xml)
    end

    xml
  rescue StandardError
    p ruby_19 { raw.encoding } || ruby_18 { raw }
    raise
  end

  #
  # Normalize empty tags in XML so comparisons are stable.
  #
  def normalize_xml_empty_tags(xml)
    xml.gsub(%r{(\S)/>}, '\1 />')
  end

  #
  # Remove dates and user specific data from core.xml.
  #
  def normalize_core_xml(xml, regression, expected)
    if expected && regression
      xml = xml.gsub(/ ?John/, '')
    end

    xml.gsub(/\d\d\d\d-\d\d-\d\dT\d\d:\d\d:\d\dZ/, '')
  end

  #
  # Normalize workbook.xml elements that commonly differ across environments.
  #
  def normalize_workbook_xml(xml)
    xml = xml.sub(/<workbookView[^>]*>/, '<workbookView/>')
    xml.sub(/<calcPr[^>]*>/, '<calcPr/>')
  end

  #
  # Normalize worksheet xml elements that commonly contain environment-specific
  # printer settings.
  #
  def normalize_worksheet_xml(xml, expected)
    xml = xml
          .sub('horizontalDpi="200" ', '')
          .sub('verticalDpi="200" ', '')

    if expected
      xml = xml.sub(/(<pageSetup[^>]* )r:id="rId1"/, '\1')
               .sub(%r{ +/>}, ' />')
    end

    xml
  end

  #
  # Normalize chart xml elements that often differ.
  #
  def normalize_chart_xml(xml)
    xml.sub(/<c:pageMargins[^>]*>/, '<c:pageMargins/>')
  end

  #
  # Convert xml text into the comparable array form used by existing tests.
  #
  def xml_to_comparable_array(xml, member_name)
    if member_name =~ /.vml$/
      vml_str_to_array(xml)
    else
      got_to_array(xml)
    end
  end

  #
  # Remove ignored XML lines for test-specific comparisons.
  #
  def reject_ignored_elements(xml_array, ignored_patterns)
    regex = Regexp.new(ignored_patterns.join('|'))
    xml_array.reject { |s| s =~ regex }
  end

  def sort_rel_file_data(xml_array)
    header = xml_array.shift
    tail   = xml_array.pop
    xml_array.sort.unshift(header).push(tail)
  end

  #
  # Build worksheet XML from a block and return it as a string.
  #
  def worksheet_xml_string
    workbook  = WriteXLSX.new(StringIO.new(''.dup))
    worksheet = workbook.add_worksheet

    yield(workbook, worksheet)

    worksheet.assemble_xml_file
    worksheet.instance_variable_get(:@writer).string
  end

  #
  # Compare worksheet XML strings using the same normalization style as the
  # existing xlsx regression helpers.
  #
  def compare_worksheet_xml(expected_xml, actual_xml)
    exp_xml = got_to_array(expected_xml)
    got_xml = got_to_array(actual_xml)

    assert_equal(exp_xml, got_xml, 'worksheet xml differs.')
  end

  #
  # Assert that a worksheet XML string includes all lines in an expected
  # XML fragment after normalization.
  #
  def assert_worksheet_xml_includes(actual_xml, expected_fragment)
    got_xml = got_to_array(actual_xml)
    exp_xml = got_to_array(expected_fragment)

    exp_xml.each do |line|
      assert_includes(got_xml, line, "worksheet xml does not include: #{line}")
    end
  end

  #
  # Assert that a worksheet XML string does not include any lines in an
  # unexpected XML fragment after normalization.
  #
  def refute_worksheet_xml_includes(actual_xml, unexpected_fragment)
    got_xml = got_to_array(actual_xml)
    exp_xml = got_to_array(unexpected_fragment)

    exp_xml.each do |line|
      refute_includes(got_xml, line, "worksheet xml unexpectedly includes: #{line}")
    end
  end

  #
  # Extract the first conditionalFormatting element from worksheet XML.
  #
  def extract_conditional_formatting_xml(xml)
    xml[/<conditionalFormatting\b.*?<\/conditionalFormatting>/m]
  end
end
