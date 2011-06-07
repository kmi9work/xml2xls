class ParserController < ApplicationController
  def index
  end
  def process_xml
    str = params[:file].read
    xls = OurParser.new(str, params[:template], File.basename(params[:file].path, ".xml"))
    if xls.is_sep?
      xls.out_xml.flatten.each_with_index do |xml, i|
        xml.write("#{Rails.root}/tmp/ss#{i}.xls")
      end
    else
      xls.out_xml.write("#{Rails.root}/tmp/ss.xls")
    end
    
  end
end
