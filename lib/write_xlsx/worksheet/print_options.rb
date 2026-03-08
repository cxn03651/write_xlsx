# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Print and display options extracted from Worksheet to slim the main class.
    module PrintOptions
      #
      # Set the page orientation as portrait.
      # The default worksheet orientation is portrait, so you won't generally
      # need to call this method.
      #
      def set_portrait
        @page_setup.orientation        = true
        @page_setup.page_setup_changed = true
      end

      #
      # Set the page orientation as landscape.
      #
      def set_landscape
        @page_setup.orientation         = false
        @page_setup.page_setup_changed  = true
      end

      #
      # This method is used to display the worksheet in "Page View/Layout" mode.
      #
      def set_page_view(flag = 1)
        @page_view = flag
      end

      #
      # set_pagebreak_view
      #
      # Set the page view mode.
      #
      def set_pagebreak_view
        @page_view = 2
      end

      #
      # Set the colour of the worksheet tab.
      #
      def tab_color=(color)
        @tab_color = Colors.new.color(color)
      end

      # This method is deprecated. use tab_color=().
      def set_tab_color(color)
        put_deprecate_message("#{self}.set_tab_color")
        self.tab_color = color
      end

      #
      # Store the horizontal page breaks on a worksheet.
      #
      def set_h_pagebreaks(*args)
        breaks = args.collect do |brk|
          Array(brk)
        end.flatten
        @page_setup.hbreaks += breaks
      end

      #
      # Store the vertical page breaks on a worksheet.
      #
      def set_v_pagebreaks(*args)
        @page_setup.vbreaks += args
      end
    end
  end
end
