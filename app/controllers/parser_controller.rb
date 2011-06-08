class ParserController < ApplicationController
  def index
  end
  def process_xml
    str = params[:file].read
    xls = OurParser.new(str, params[:template], File.basename(params[:file].original_filename, ".xml"))
    if xls.is_sep?
      xls.out_xml.each do |h_xml|
        h_xml.each do |filename, xml| 
          xml.write("#{Rails.root}/tmp/#{filename}")
        end
      end
    else
      xls.out_xml.write("#{Rails.root}/tmp/#{File.basename(params[:file].original_filename, ".xml")}.xls")
    end
  end
end
