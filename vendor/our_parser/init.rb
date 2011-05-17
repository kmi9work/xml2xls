# Include hook code here
require 'rubygems'
#require 'ruby-debug'
require './lib/our_parser.rb'
#Debugger.start
xls = OurParser.new(File.read(ARGV[0]), ARGV[1])
puts xls.out_xml.to_s