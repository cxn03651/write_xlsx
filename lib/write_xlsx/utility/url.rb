# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module Url
      def escape_url(url)
        unless url =~ /%[0-9a-fA-F]{2}/
          # Escape the URL escape symbol.
          url = url.gsub("%", "%25")

          # Escape whitespae in URL.
          url = url.gsub(/[\s\x00]/, '%20')

          # Escape other special characters in URL.
          re = /(["<>\[\]`^{}])/
          while re =~ url
            match = $LAST_MATCH_INFO[1]
            url = url.sub(re, sprintf("%%%x", match.ord))
          end
        end

        url
      end
    end
  end
end
