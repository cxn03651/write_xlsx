#!/usr/bin/env ruby
# -*- encoding: utf-8 -*-

require 'zip/zipfilesystem'
require 'fileutils'

  # src  zip filename
  # dest  destination directory
  # options :fs_encoding=[UTF-8,Shift_JIS,EUC-JP]
  def extract_vba_project(src, dest, options = {})
    FileUtils.makedirs(dest)
    Zip::ZipInputStream.open(src) do |is|
      loop do
        entry = is.get_next_entry()
        break if entry.nil?()
        if entry.name == 'xl/vbaProject.bin'
          path = File.join(dest, 'vbaProject.bin')
          File.open(path, File::CREAT|File::WRONLY|File::BINARY) do |w|
            w.puts(is.read())
          end
          # The mod data on vbaProject.bin isn't generally set correctly in the
          # xlsm/zip file. This can cause issues on Windows so reset it to the
          # current data.
          ctime = Time.now
          File::utime(ctime, ctime, path)
          break
        end
      end
    end
  end

  # main

  extract_vba_project('add_vba_project.xlsm', './')
