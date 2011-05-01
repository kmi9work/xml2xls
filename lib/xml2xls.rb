module Xml2xls
  require 'rubygems'
  require 'xml/smart'
  require 'xml/xslt'
  def xsl_transform file_path, template_path
    xslt = XML::XSLT.new()

    xslt.xml = XML::Smart.open file_path
    xslt.xsl = XML::Smart.open template_path

    out = xslt.serve()
    return out;
  end
  def xml_to_xls xml_file
    
  end
  
  def get_template_suffix prefix
    index = prefix.index("_")
    return "" unless index
    return prefix[(index+1)..-1]
  end
  
  def process_xml_sep_node reader, template_name, file_name
    prefix = reader.prefix
    suffix = get_template_suffix(prefix)
    index = suffix.index("_")
    throw "suffix" unless index  #make exceptions
    country = suffix[(index + 1)..-1].downcase
    xsl_path = "xslt" # or use FileManager
    template_path = File.join(xsl_path, template_name + "_sep.xsl")
    return nil unless File.file? template_path
    xls_folder = "xls" # or use FileManager
    file_path = File.join(xls_folder, file_name) # or use FileManager.get_new_file_path
    node_xml = reader.read_outer_xml
    return nil if node_xml.strip.empty?
    
  end
  
  def process_to_xls xml_path, template_name
    log_message
    xls_folder_path = "xslt" # or use FileManager
    if File.file? File.join(xls_folder_path, template_name + "_sep.xsl")
      reader = XML::Reader.file(xml_path)
      
      while reader.move_to_element("Item")
        process_xml_sep_node(reader, template_name, xml_path)
      end
      #transform xml only with _sep.xsl
    elsif File.file? File.join(xls_folder_path, template_name + "_.xsl") or File.file? File.join(xls_folder_path, template_name + "_h.xsl") or File.file? File.join(xls_folder_path, template_name + "_gen.xsl")
      #transform xml with all others xsl's
    else
      #transform xml via txt styles
    end
  end
end



include Xml2xls
# require 'spreadsheet'
# book = Spreadsheet.open "_ololo.xls"
# sheet = book.worksheet 0
# puts sheet.row(0).format(0).inspect.gsub(", ", "\n")

puts xsl_transform ARGV[0], ARGV[1]
