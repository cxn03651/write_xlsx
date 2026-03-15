# -*- encoding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class FontStyle
      def initialize(format)
        @format = format
      end

      def index
        @format.state.font.index
      end

      def index=(value)
        @format.state.font.index = value
        @format.send(:sync_font_ivars_from_state)
      end

      def name
        @format.state.font.name
      end

      def name=(value)
        @format.state.font.name = value
        @format.send(:sync_font_ivars_from_state)
      end

      def size
        @format.state.font.size
      end

      def size=(value)
        @format.state.font.size = value
        @format.send(:sync_font_ivars_from_state)
      end

      def bold
        @format.state.font.bold
      end

      def bold=(value)
        @format.state.font.bold = value
        @format.send(:sync_font_ivars_from_state)
      end

      def italic
        @format.state.font.italic
      end

      def italic=(value)
        @format.state.font.italic = value
        @format.send(:sync_font_ivars_from_state)
      end

      def color
        @format.state.font.color
      end

      def color=(value)
        @format.state.font.color = value
        @format.send(:sync_font_ivars_from_state)
      end

      def color_indexed
        @format.state.font.color_indexed
      end

      def color_indexed=(value)
        @format.state.font.color_indexed = value
        @format.send(:sync_font_ivars_from_state)
      end

      def underline
        @format.state.font.underline
      end

      def underline=(value)
        @format.state.font.underline = value
        @format.send(:sync_font_ivars_from_state)
      end

      def strikeout
        @format.state.font.strikeout
      end

      def strikeout=(value)
        @format.state.font.strikeout = value
        @format.send(:sync_font_ivars_from_state)
      end

      def outline
        @format.state.font.outline
      end

      def outline=(value)
        @format.state.font.outline = value
        @format.send(:sync_font_ivars_from_state)
      end

      def shadow
        @format.state.font.shadow
      end

      def shadow=(value)
        @format.state.font.shadow = value
        @format.send(:sync_font_ivars_from_state)
      end

      def script
        @format.state.font.script
      end

      def script=(value)
        @format.state.font.script = value
        @format.send(:sync_font_ivars_from_state)
      end

      def family
        @format.state.font.family
      end

      def family=(value)
        @format.state.font.family = value
        @format.send(:sync_font_ivars_from_state)
      end

      def charset
        @format.state.font.charset
      end

      def charset=(value)
        @format.state.font.charset = value
        @format.send(:sync_font_ivars_from_state)
      end

      def scheme
        @format.state.font.scheme
      end

      def scheme=(value)
        @format.state.font.scheme = value
        @format.send(:sync_font_ivars_from_state)
      end

      def condense
        @format.state.font.condense
      end

      def condense=(value)
        @format.state.font.condense = value
        @format.send(:sync_font_ivars_from_state)
      end

      def extend
        @format.state.font.extend
      end

      def extend=(value)
        @format.state.font.extend = value
        @format.send(:sync_font_ivars_from_state)
      end

      def theme
        @format.state.font.theme
      end

      def theme=(value)
        @format.state.font.theme = value
        @format.send(:sync_font_ivars_from_state)
      end

      def hyperlink
        @format.state.font.hyperlink
      end

      def hyperlink=(value)
        @format.state.font.hyperlink = value
        @format.send(:sync_font_ivars_from_state)
      end

      def color_indexed
        @format.state.font.color_indexed
      end

      def color_indexed=(value)
        @format.state.font.color_indexed = value
        @format.send(:sync_font_ivars_from_state)
      end
    end
  end
end
