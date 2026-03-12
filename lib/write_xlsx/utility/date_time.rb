# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module DateTime
      #
      # convert_date_time(date_time_string)
      #
      # The function takes a date and time in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format
      # and converts it to a decimal number representing a valid Excel date.
      #
      def convert_date_time(date_time_string)       # :nodoc:
        date_time = date_time_string.to_s.sub(/^\s+/, '').sub(/\s+$/, '').sub(/Z$/, '')

        # Check for invalid date char.
        return nil if date_time =~ /[^0-9T:\-.Z]/

        # Check for "T" after date or before time.
        return nil unless date_time =~ /\dT|T\d/

        days      = 0 # Number of days since epoch
        seconds   = 0 # Time expressed as fraction of 24h hours in seconds

        # Split into date and time.
        date, time = date_time.split("T")

        # We allow the time portion of the input DateTime to be optional.
        if time
          # Match hh:mm:ss.sss+ where the seconds are optional
          if time =~ /^(\d\d):(\d\d)(:(\d\d(\.\d+)?))?/
            hour   = ::Regexp.last_match(1).to_i
            min    = ::Regexp.last_match(2).to_i
            sec    = ::Regexp.last_match(4).to_f
          else
            return nil # Not a valid time format.
          end

          # Some boundary checks
          return nil if hour >= 24
          return nil if min  >= 60
          return nil if sec  >= 60

          # Excel expresses seconds as a fraction of the number in 24 hours.
          seconds = ((hour * 60 * 60) + (min * 60) + sec) / (24.0 * 60 * 60)
        end

        # We allow the date portion of the input DateTime to be optional.
        return seconds if date == ''

        # Match date as yyyy-mm-dd.
        if date =~ /^(\d\d\d\d)-(\d\d)-(\d\d)$/
          year   = ::Regexp.last_match(1).to_i
          month  = ::Regexp.last_match(2).to_i
          day    = ::Regexp.last_match(3).to_i
        else
          return nil  # Not a valid date format.
        end

        # Set the epoch as 1900 or 1904. Defaults to 1900.
        # Special cases for Excel.
        unless date_1904?
          return      seconds if date == '1899-12-31' # Excel 1900 epoch
          return      seconds if date == '1900-01-00' # Excel 1900 epoch
          return 60 + seconds if date == '1900-02-29' # Excel false leapday
        end

        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        #
        epoch   = date_1904? ? 1904 : 1900
        offset  = date_1904? ? 4 : 0
        norm    = 300
        range   = year - epoch

        # Set month days and check for leap year.
        mdays   = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        leap    = 0
        leap    = 1  if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
        mdays[1] = 29 if leap != 0

        # Some boundary checks
        return nil if year  < epoch || year  > 9999
        return nil if month < 1     || month > 12
        return nil if day   < 1     || day   > mdays[month - 1]

        # Accumulate the number of days since the epoch.
        days = day                               # Add days for current month
        (0..(month - 2)).each do |m|
          days += mdays[m]                      # Add days for past months
        end
        days += range * 365                      # Add days for past years
        days += (range / 4)                      # Add leapdays
        days -= ((range + offset) / 100)         # Subtract 100 year leapdays
        days += ((range + offset + norm) / 400)  # Add 400 year leapdays
        days -= leap                             # Already counted above

        # Adjust for Excel erroneously treating 1900 as a leap year.
        days += 1 if !date_1904? && days > 59

        date_time = sprintf("%0.10f", days + seconds)
        date_time = date_time.sub(/\.?0+$/, '') if date_time =~ /\./
        if date_time =~ /\./
          date_time.to_f
        else
          date_time.to_i
        end
      end
    end
  end
end
