# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  module Utility
    module XmlPrimitives
      #
      # Write the <color> element.
      #
      def write_color(name, value, writer = @writer) # :nodoc:
        attributes = [[name, value]]

        writer.empty_tag('color', attributes)
      end

      def write_xml_declaration
        @writer.xml_decl
        yield
        @writer.crlf
        @writer.close
      end

      def r_id_attributes(id)
        ['r:id', "rId#{id}"]
      end

      def xml_str
        @writer.string
      end
    end
  end
end
