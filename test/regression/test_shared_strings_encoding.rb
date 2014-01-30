# -*- coding: iso-8859-8 -*-

require 'helper'

class TestRegressionSharedStringsEncoding < Test::Unit::TestCase
  def setup
    setup_dir_var
    @f = File.open(File.join(@test_dir, 'regression', 'klt.csv'), "r:iso-8859-8")
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_shared_strings_encoding
    sizes = [10, 10, 18, 10, 18, 18, 12, 16, 16, 25,
             18, 36, 60, 16, 16, 16, 18, 16, 16, 22,
             14, 14, 14, 14, 14, 14, 14, 14, 14, 14,
             14, 14, 14, 14, 14, 14]

#    input_file = ARGV[0]
#    output_file = ARGV[1]

    # coding: ISO-8859-8
#    Encoding.default_external = Encoding.find("ISO-8859-8")

    # define an array to hold the Sikum records

    # loop through each record in the csv file, adding
    # each record to our array.
    #f = File.open(input_file, "r")
    @xlsx = 'shared_strings_encoding.xlsx'
    workbook = WriteXLSX.new(@xlsx)
    heading1 = workbook.add_format(:align => 'center', :bold => 1 , :size => 12 , :color => 'blue', :bg_color => 27)
    heading2 = workbook.add_format(:align => 'center', :bold => 1 , :size => 12 , :color => 'blue', :bg_color => 'silver')
    format1 = workbook.add_format(:align => 'right', :color => 'blue', :bg_color => 9 )
    format11 = workbook.add_format(:align => 'right', :color => 'blue', :bg_color => 9 )
    format11.set_num_format('0.00')
    format3 = workbook.add_format(:align => 'right', :color => 'red', :bg_color => 9 )
    worksheet = workbook.add_worksheet
    worksheet.right_to_left

    i=0

    @f.each_line { |line|
      words = line.split(',')
      if i == 0 then
        for j in 0..words.size-1

          word = words[j].to_s.gsub(/"/," ")
          word = word.gsub(/\n/, "" )
          word = word.gsub(/\r/, "" )

#          worksheet.merge_range( 'A1:L1', word.force_encoding('iso-8859-8').encode("UTF-8"), heading1)
          worksheet.merge_range( 'A1:L1', word, heading1)
        end
      elsif i == 1 then
        for j in 0..words.size-1

          word = words[j].to_s.gsub(/"/," ")
          word = word.gsub(/\n/, "" )
          word = word.gsub(/\r/, "" )

          worksheet.set_column(j,0,sizes[j])
#          worksheet.write( i, j, word.force_encoding('iso-8859-8').encode("UTF-8"), heading2)
          worksheet.write( i, j, word, heading2)
        end
      else

        heara = words[11].to_s
        for j in 0..words.size-1

          jj=j
          word = words[j].to_s.gsub("_"," ")
          word = word.gsub(/\n/, "" )
          word = word.gsub(/\r/, "" )
          word = word.lstrip.rstrip

          if heara.include?('abcdefghijklml') then
#            worksheet.write( i, jj, word.force_encoding('iso-8859-8').encode("UTF-8"), format3)
            worksheet.write( i, jj, word, format3)
          else
#            worksheet.write( i, jj, word.force_encoding('iso-8859-8').encode("UTF-8"), format1)
            worksheet.write( i, jj, word, format1)
          end
        end
      end
      i = i+1
    }

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end

  def input_file
    <<EOS
רשימת חיוויים שהגיעו מהיצרן
תאריך חזרה,שעת חזרה,שם יצרן,מסמך יצרן,מספר זהות,שם,מספר פוליסה,מטפל,צוות,מסמך,תהליך,הערה
01/11/2012,09:04:20,_מגדל                                             ,      131, 39977777,ממממממממממש         ,00171044973,                    ,                    ,דף הערות                 ,                         ,תהליך הפקה/שינוי לא נמצא במערכת
01/11/2012,09:04:24,_מגדל                                             ,      131, 36888865,ההההההה             ,00171081635,נעמה ודאוקר         ,לילי לילעזר         ,דף הערות                 ,                         ,
EOS
  end
end
