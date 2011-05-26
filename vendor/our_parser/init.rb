# Include hook code here
require 'rubygems'
#require 'ruby-debug'
require './lib/our_parser.rb'
#Debugger.start
xls = OurParser.new(File.read(ARGV[0]), ARGV[1])
# puts '<?xml version="1.0" encoding="UTF-8"?>'
# puts '<?mso-application progid="Excel.Sheet"?>'
# puts xls.out_xml
xls.out_xml.write("ss.xls")