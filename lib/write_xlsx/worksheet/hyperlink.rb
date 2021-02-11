# -*- encoding: utf-8 -*-

module Writexlsx
  class Worksheet
    class Hyperlink   # :nodoc:
      include Writexlsx::Utility

      attr_reader :str, :tip

      MAXIMUM_URLS_SIZE = 255

      def self.factory(url, str = nil, tip = nil)
        if url =~ /^internal:(.+)/
          InternalHyperlink.new($~[1], str, tip)
        elsif url =~ /^external:(.+)/
          ExternalHyperlink.new($~[1], str, tip)
        else
          new(url, str, tip)
        end
      end

      def initialize(url, str, tip)
        # The displayed string defaults to the url string.
        str ||= url.dup

        # Strip the mailto header.
        normalized_str = str.sub(/^mailto:/, '')

        # Split url into the link and optional anchor/location.
        url, @url_str = url.split(/#/, 2)

        # Escape URL unless it looks already escaped.
        url = escape_url(url)

        # Excel limits the escaped URL and location/anchor to 255 characters.
        if url.bytesize > 255 || (!@url_str.nil? && @url_str.bytesize > 255)
          raise "Ignoring URL '#{url}' where link or anchor > 255 characters since it exceeds Excel's limit for URLS. See LIMITATIONS section of the Excel::Writer::XLSX documentation."
        end

        @url       = url
        @str       = normalized_str
        @tip       = tip
      end

      def attributes(row, col, id)
        ref = xl_rowcol_to_cell(row, col)

        attr = [ ['ref', ref] ]
        attr << r_id_attributes(id)

        attr << ['location', @url_str] if @url_str
        attr << ['display',  @display] if @display
        attr << ['tooltip',  @tip]     if @tip
        attr
      end

      def external_hyper_link
        ['/hyperlink', @url, 'External']
      end

      def display_on
        @display = @url_str
      end
    end

    class InternalHyperlink < Hyperlink
      undef external_hyper_link

      def initialize(url, str, tip)
        @url = url
        # The displayed string defaults to the url string.
        str ||= @url.dup

        # Strip the mailto header.
        @str = str.sub(/^mailto:/, '')

        # Copy string for use in hyperlink elements.
        @url_str = @str.dup

        # Excel limits escaped URL to 255 characters.
        if @url.bytesize > MAXIMUM_URLS_SIZE
          raise "URL '#{@url}' > #{MAXIMUM_URLS_SIZE} characters, it exceeds Excel's limit for URLS."
        end

        @tip = tip
      end

      def attributes(row, col, dummy = nil)
        attr = [
                ['ref', xl_rowcol_to_cell(row, col)],
                ['location', @url]
               ]

        attr << ['tooltip', @tip] if @tip
        attr << ['display', @str]
      end
    end

    class ExternalHyperlink < Hyperlink
      def initialize(url, str, tip)
        # The displayed string defaults to the url string.
        str ||= url.dup

        # For external links change the directory separator from Unix to Dos.
        url = url.gsub(%r|/|, '\\')
        str.gsub!(%r|/|, '\\')

        # Strip the mailto header.
        str.sub!(/^mailto:/, '')

        # Split url into the link and optional anchor/location.
        url, url_str = url.split(/#/, 2)

        # Escape URL unless it looks already escaped.
        url = escape_url(url)

        # Add the file:/// URI to the url if non-local.
        if url =~ %r![:]! ||        # Windows style "C:/" link.
            url =~ %r!^\\\\!        # Network share.
          url = "file:///#{url}"
        end

        # Convert a ./dir/file.xlsx link to dir/file.xlsx.
        url = url.sub(%r!^.\\!, '')
        @url_str   = url_str

        # Excel limits the escaped URL and location/anchor to 255 characters.
        if url.bytesize > 255 || (!@url_str.nil? && @url_str.bytesize > 255)
          raise "Ignoring URL '#{url}' where link or anchor > 255 characters since it exceeds Excel's limit for URLS. See LIMITATIONS section of the Excel::Writer::XLSX documentation."
        end

        @url       = url
        @str       = str
        @tip       = tip
      end
    end
  end
end
