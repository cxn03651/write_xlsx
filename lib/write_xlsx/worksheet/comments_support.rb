# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module CommentsSupport
      #
      # This method is used to make all cell comments visible when a worksheet
      # is opened.
      #
      def show_comments(visible = true)
        @comments_visible = visible
      end

      def comments_visible? # :nodoc:
        !!@comments_visible
      end

      #
      # This method is used to set the default author of all cell comments.
      #
      def comments_author=(author)
        @comments_author = author || ''
      end

      # This method is deprecated. use comments_author=().
      def set_comments_author(author)
        put_deprecate_message("#{self}.set_comments_author")
        self.comments_author = author
      end

      def sorted_comments # :nodoc:
        @comments.sorted_comments
      end

      def num_comments_block
        @comments.size / 1024
      end

      def has_comments? # :nodoc:
        !@comments.empty?
      end

      def has_vml?  # :nodoc:
        @has_vml
      end

      def has_header_vml?  # :nodoc:
        !(header_images.empty? && footer_images.empty?)
      end

      def buttons_data  # :nodoc:
        @buttons_array
      end

      def header_images_data  # :nodoc:
        @header_images_array
      end
    end
  end
end
