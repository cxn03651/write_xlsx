# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module Common
      PERL_TRUE_VALUES   = [false, nil, 0, "0", "", [], {}].freeze

      #
      # return perl's boolean result
      #
      def ptrue?(value)
        !PERL_TRUE_VALUES.include?(value)
      end

      def check_parameter(params, valid_keys, method)
        invalids = params.keys - valid_keys
        unless invalids.empty?
          raise WriteXLSXOptionParameterError,
                "Unknown parameter '#{invalids.join(", ")}' in #{method}."
        end
        true
      end

      def absolute_char(absolute)
        absolute ? '$' : ''
      end

      def float_to_str(float)
        return '' unless float

        if float == float.to_i
          float.to_i.to_s
        else
          float.to_s
        end
      end

      def put_deprecate_message(method)
        warn("Warning: calling deprecated method #{method}. This method will be removed in a future release.")
      end
    end
  end
end
