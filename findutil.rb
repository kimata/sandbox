#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require "pathname"

class FindUtil
  def self.find(path, opt, &block)
    path = Pathname.new(path)
    find_impl(path, 1, norm_opt(opt, path), &block)
  end

  def self.find_impl(path, level, opt, &block)
    path.children(true).each {|entry|
      if (entry.directory?) then
        next if opt.has_key?(:dir_filter) && !opt[:dir_filter].call(entry.basename.to_s)
        find_impl(entry, level, opt, &block)
      else
        next if opt.has_key?(:file_filter) && !opt[:file_filter].call(entry.basename.to_s)
        block.call(entry.to_s, entry.relative_path_from(opt[:base_path]))
      end
    }
  end

  def self.norm_opt(opt, path)
    norm_opt = opt.clone
    norm_opt[:base_path] = path
    norm_opt[:max_leve]  = 1000 unless opt.has_key?(:max_leve)
    return norm_opt
  end
end

if __FILE__ == $PROGRAM_NAME
  path = '.'

  FindUtil.find(path, {
                  max_level: 1,
                  file_filter: lambda {|name| /^.*\.rb$/.match(name) },
                  dir_filter: lambda {|name| !/old/.match(name) },
                }) {|path, rel_path|
    puts "path: %-30s, relpath:%s\n" % [path, rel_path]
  }
end




