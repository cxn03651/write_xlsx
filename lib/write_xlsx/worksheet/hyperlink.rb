# -*- encoding: utf-8 -*-

module Writexlsx
  class Worksheet
    class Hyperlink   # :nodoc:
      include Writexlsx::Utility

      attr_reader :url, :link_type, :str, :url_str
      attr_accessor :tip, :display

      def initialize(url, str = nil)
        link_type = 1

        # Remove the URI scheme from internal links.
        if url =~ /^internal:/
          url = url.sub(/^internal:/, '')
          link_type = 2
          # Remove the URI scheme from external links.
        elsif url =~ /^external:/
          url = url.sub(/^external:/, '')
          link_type = 3
        end

        # The displayed string defaults to the url string.
        str ||= url.dup

        # For external links change the directory separator from Unix to Dos.
        if link_type == 3
          url = url.gsub(%r|/|, '\\')
          str.gsub!(%r|/|, '\\')
        end

        # Strip the mailto header.
        str.sub!(/^mailto:/, '')

        # Copy string for use in hyperlink elements.
        url_str = str.dup

        # External links to URLs and to other Excel workbooks have slightly
        # different characteristics that we have to account for.
        if link_type == 1
          # Escape URL unless it looks already escaped.
          unless url =~ /%[0-9a-fA-F]{2}/
            # Escape the URL escape symbol.
            url = url.gsub(/%/, "%25")

            # Escape whitespae in URL.
            url = url.gsub(/[\s\x00]/, '%20')

            # Escape other special characters in URL.
            re = /(["<>\[\]`^{}])/
            while re =~ url
              match = $~[1]
              url = url.sub(re, sprintf("%%%x", match.ord))
            end
          end

          # Ordinary URL style external links don't have a "location" string.
          url_str = nil
        elsif link_type == 3
          # External Workbook links need to be modified into the right format.
          # The URL will look something like 'c:\temp\file.xlsx#Sheet!A1'.
          # We need the part to the left of the # as the URL and the part to
          # the right as the "location" string (if it exists).
          url, url_str = url.split(/#/)

          # Add the file:/// URI to the url if non-local.
          if url =~ %r![:]! ||        # Windows style "C:/" link.
              url =~ %r!^\\\\!        # Network share.
            url = "file:///#{url}"
          end

          # Convert a ./dir/file.xlsx link to dir/file.xlsx.
          url = url.sub(%r!^.\\!, '')

          # Treat as a default external link now that the data has been modified.
          link_type = 1
        end

        # Excel limits escaped URL to 255 characters.
        if url.bytesize > 255
          raise "URL '#{url}' > 255 characters, it exceeds Excel's limit for URLS."
        end
        @url       = url
        @link_type = link_type
        @str       = str
        @url_str   = url_str
      end

      def write_external_attributes(row, col, id)
        ref = xl_rowcol_to_cell(row, col)

        attributes = ['ref', ref]
        attributes += r_id_attributes(id)

        attributes << 'location' << url_str if url_str
        attributes << 'display'  << display if display
        attributes << 'tooltip'  << tip     if tip
        attributes
      end

      def write_internal_attributes(row, col)
        ref = xl_rowcol_to_cell(row, col)

        attributes = ['ref', ref, 'location', url]

        attributes << 'tooltip' << tip if tip
        attributes << 'display' << str
      end
    end
  end
end
