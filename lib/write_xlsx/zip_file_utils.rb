# -*- coding: utf-8 -*-
#
# from http://d.hatena.ne.jp/alunko/20071021
#
require 'kconv'
require 'zip/zipfilesystem'
require 'fileutils'

module ZipFileUtils

  # src  file or directory
  # dest  zip filename
  # options :fs_encoding=[UTF-8,Shift_JIS,EUC-JP]
  def self.zip(src, dest, options = {})
    src = File.expand_path(src)
    dest = File.expand_path(dest)
    File.unlink(dest) if File.exist?(dest)
    Zip::ZipFile.open(dest, Zip::ZipFile::CREATE) {|zf|
      if(File.file?(src))
        zf.add(encode_path(File.basename(src), options[:fs_encoding]), src)
        break
      else
        each_dir_for(src){ |path|
          if File.file?(path)
            zf.add(encode_path(relative(path, src), options[:fs_encoding]), path)
          elsif File.directory?(path)
            zf.mkdir(encode_path(relative(path, src), options[:fs_encoding]))
          end
        }
      end
    }
    FileUtils.chmod(0644, dest)
  end

  # src  zip filename
  # dest  destination directory
  # options :fs_encoding=[UTF-8,Shift_JIS,EUC-JP]
  def self.unzip(src, dest, options = {})
    FileUtils.makedirs(dest)
    Zip::ZipInputStream.open(src) do |is|
      loop do
        entry = is.get_next_entry()
        break unless entry
        dir = File.dirname(entry.name)
        FileUtils.makedirs(dest+ '/' + dir)
        path = encode_path(dest + '/' + entry.name, options[:fs_encoding])
        if(entry.file?())
          File.open(path, File::CREAT|File::WRONLY|File::BINARY) do |w|
            w.puts(is.read())
          end
        else
          FileUtils.makedirs(path)
        end
      end
    end
  end

  private

  def self.each_dir_for(dir_path, &block)
    dir = Dir.open(dir_path)
    each_file_for(dir_path){ |file_path|
      yield(file_path)
    }
  end

  def self.each_file_for(path, &block)
    if File.file?(path)
      yield(path)
      return true
    end
    dir = Dir.open(path)
    file_exist = false
    dir.each(){ |file|
      next if file == '.' || file == '..'
      file_exist = true if each_file_for(path + "/" + file, &block)
    }
    yield(path) unless file_exist
    return file_exist
  end

  def self.relative(path, base_dir)
    path[base_dir.length() + 1 .. path.length()] if path.index(base_dir) == 0
  end

  def self.encode_path(path, encode_s)
    return path unless encode_s
    case(encode_s)
    when('UTF-8')
      return path.toutf8()
    when('Shift_JIS')
      return path.tosjis()
    when('EUC-JP')
      return path.toeuc()
    else
      return path
    end
  end
end
