#only Ruby 1.9
=begin
class _Xml2xl
  require 'rubygems'
  require 'libxml-ruby'
  require 'xml/xslt'
  
  FILEMANAGER = {
    "xslt" => "xslt",
    "xls" => "xls"
  }
  
  CONFIG = {
    "cache_enabled" => true
  }
  CONSTANTS = {
    :BaseItem => "BaseItem",
    :Assortment => "Assortment",
    :PackagingItem => "PackagingItem",
    :BaseItemVersion => "BaseItemVersion",
    :AssortmentVersion => "AssortmentVersion",
    :PackagingItemVersion => "PackagingItemVersion",
    :AddActionCode => "add",
    :PriorAddActionCode => "priod add",
    :DeleteActionCode => "del",
    :ChangeActionCode => "chg",
    :CorrectionActionCode => "cor",
    :XmlNamespace => "http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItem{0}",
    :Changed => "changed",
    :Deleted => "deleted",
    :Added => "added"
  }
  
  @xml_tpl_header = "
<?xml version=\"1.0\" encoding=\"utf-8\"?>
<?mso-application progid=\"Excel.Sheet\"?>
<s:Workbook xmlns:s=\"urn:schemas-microsoft-com:office:spreadsheet\" 
xmlns:x=\"urn:schemas-microsoft-com:office:excel\" 
xmlns:o=\"urn:schemas-microsoft-com:office:office\" 
xmlns:sinfos=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/TradeItemMessage\" 
xmlns:fnf_fnd_at=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_AT\" 
xmlns:fnf_fnd_be=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_BE\" 
xmlns:fnf_fnd_ch=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_CH\" 
xmlns:fnf_fnd_cz=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_CZ\" 
xmlns:fnf_fnd_de=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_DE\" 
xmlns:fnf_fnd_dk=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_DK\" 
xmlns:fnf_fnd_ee=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_EE\" 
xmlns:fnf_fnd_es=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_ES\" 
xmlns:fnf_fnd_fi=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_FI\" 
xmlns:fnf_fnd_fr=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_FR\" 
xmlns:fnf_fnd_gb=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_GB\" 
xmlns:fnf_fnd_hu=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_HU\" 
xmlns:fnf_fnd_ie=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_IE\" 
xmlns:fnf_fnd_it=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_IT\" 
xmlns:fnf_fnd_nl=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_NL\" 
xmlns:fnf_fnd_pl=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_PL\" 
xmlns:fnf_fnd_pt=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_PT\" 
xmlns:fnf_fnd_ro=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_RO\" 
xmlns:fnf_fnd_ru=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_RU\" 
xmlns:fnf_fnd_se=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_SE\" 
xmlns:fnf_fnd_ua=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_FND_UA\" 
xmlns:fnf_rap_at=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_AT\" 
xmlns:fnf_rap_de=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_DE\" 
xmlns:fnf_rap_dk=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_DK\" 
xmlns:fnf_rap_ee=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_EE\" 
xmlns:fnf_rap_fi=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_FI\" 
xmlns:fnf_rap_pl=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_PL\" 
xmlns:fnf_rap_ru=\"http://schemas.sinfos.de/TradeItemMessages/1.2.0/FNF/TradeItemFNF_RAP_RU\" 
>{0}<s:Worksheet s:Name=\"Items\"><s:Table {1}>"
  
  @xml_table_tpl_footer = "</s:Table>"
  @xml_work_sheet_tpl_footer = "</s:Worksheet>"
  @xml_workbook_tpl_footer = "</s:Workbook>"
  @map_to_regex = Regexp.new "<mapTo list=\"(?<list>[^\"]+)\" firstcell=\"(?<first>[^\"]+)\" secondcell=\"(?<second>[^\"]+)\">(?<code>[^<]+)</mapTo>"
  @convert_regex = Regexp.new "<mapTo list=\"(?<list>[^\"]+)\" tocode=\"(?<tocode>[^\"]+)\" firstcell=\"(?<first>[^\"]+)\" secondcell=\"(?<second>[^\"]+)\" thirdcell=\"(?<third>[^\"]+)\">(?<code>[^<]+)</mapTo>"
  @map_substitution_string = "<mapTo list=\"{0}\" firstcell=\"{2}\" secondcell=\"{3}\">{1}</mapTo>"
  @convert_substitution_string = "<mapTo list=\"{0}\" tocode=\"{2}\" firstcell=\"{3}\" secondcell=\"{4}\" thirdcell=\"{5}\">{1}</mapTo>"
  @regex_tpls = {}
  @group =["(?<code>\\D+)", "(?<code>\\d+)", "(?<code>\\D+)", "(?<code>\\D{1,3})"]
  @tpl_text = nil
  @pi_tpl_text = nil
  @main_transforms = {}
  @item_transforms = {}
  @item_transform_templates = {}
  
  
  def get_template template
    return @main_transforms[template] if @main_transforms[template]
    transform = XML::XSLT.new
    xsl_folder_path = FILEMANAGER["xslt"]
    xsl_file_path = File.join(xsl_folder_path, template + ".xsl")
    throw "Couldn`t find file with xslt" unless File.file? xsl_file_path #make exceptions
    transform.xsl = xsl_file_path
    @main_transforms[template] = transform
  end
  
  def get_item_template template_path
    return @item_transform_templates[template_path] if @item_transform_templates[template_path]
    f = File.open template_path
    @item_transform_templates[template_path] = f.read
    f.close
    return @item_transform_templates[template_path]
  end
  
  def get_item_transform file_path, country, suffix
    key = country + ";" + suffix + ";" + file_path
    return @item_transforms[key] if @item_transforms[key]
    xsl = get_item_template(file_path)
    xsl.format!(country.upcase, suffix.downcase, suffix.upcase)
    transform = XML::XSLT.new
    transform.xsl = xsl
    @item_transforms[key] = transform
  end
  
  def get_template_suffix prefix
    index = prefix.index("_")
    return "" unless index
    return prefix[(index+1)..-1]
  end
  
  def get_bi_template_part hierarchy_dictionary, bigtin, suffix
    lower = suffix.downcase
    HIERARCHY = "-PI"
    @str = ""
    pi_str = get_pi_text(hierarchy_dictionary, bigtin, HIERARCHY, lower)
    return @tpl_text.format(lower, bigtin, hierarchy_dictionary[bigtin], pi_str)
  end
  
  def get_pi_text hierarchy_dictionary, parent_gtin, prefix, lower
    hierarchy_dictionary.each do |key, value|
      if value == parent_gtin
        @str += @pi_tpl_text.format(lower, key, prefix, "")
        get_pi_text(hierarchy_dictionary, key, '-' + prefix, lower)
      end
    end
  end
  
  def get_reader xml
    return LibXML::XML::Reader.string(xml)
  end
  
  
  def process_xml_node reader, template_name
    #returns string, code
    prefix = reader.prefix
    suffix = get_template_suffix(prefix)
    index = suffix.index("_")
    throw "suffix" unless index #make exceptions
    country = suffix[(index+1)..-1].downcase
    return "","" if !suffix or suffix.empty?
    transform = XML::XSLT.new
    xsl_file_path = FILEMANAGER["xslt"]
    file_path = File.join(xsl_file_path, template_name + "_.xsl")
    h_xsl_file_path = File.join(xsl_file_path, template_name + "_h.xsl")
    xsl_file_path = File.join(xsl_file_path, template_name + "_gen.xsl")
    node_xml = reader.read_outer_xml
    return "","" if node_xml.strip.empty?
    if File.file? file_path
      transform = get_item_transform(file_path, country, suffix)
    elsif File.file? h_xsl_file_path
      bi_str = ""
      ba_str = ""
      hierarchy_dictionary, dict = build_hierarchy(get_reader(node_xml), suffix)
      hierarchy_dictionary.each do |key, value|
        bi_str += get_bi_template_part(hierarchy_dictionary, key, suffix) if value == "BI"
        ba_str += get_bi_template_part(hierarchy_dictionary, key, suffix) if value == "BA"
      end
      xsl = get_item_template(h_xsl_file_path)
      xsl.format!(country.upcase, suffix.downcase, suffix.upcase, bi_str, ba_str)
      transform.xsl = xsl
    else
      return "","" unless File.file? xsl_file_path
      transform = get_item_transform(xsl_file_path, country, suffix)
    end
    transform.xml = node_xml
    out = transform.serve
    code = suffix
    return out, code
  end
  
  def process_to_xls xml_path, template_name
    time = Time.now
    cache_enabled = CONFIG["cache_enabled"]
    #get_xml(xml_path) if (cache_enabled and(false))#smth with account and cache
    xsl_folder_path = FILEMANAGER["xslt"]
    if File.file? File.join(xsl_folder_path, template_name + "_sep.xsl")
      reader = LibXML::XML::Reader.file(xml_path)
      files = []
      while reader.read
        next if (reader.local_name != "Item")
        files << process_xml_sep_node(reader, template_name, xml_path)
      end
      reader.close
      #zip, delete, write log, return zip
    else
      transform = get_template(template_name)
      transform.xml = "<empty></empty>"
      xml = transform.serve
      size = 0
      if File.file? File.join(xls_folder_path, template_name + "_.xsl") or File.file? File.join(xls_folder_path, template_name + "_h.xsl") or File.file? File.join(xls_folder_path, template_name + "_gen.xsl")
        xls_folder = FILEMANAGER["xls"]
        file_name = xml_path
        xls_path = File.join(xls_folder, file_name)
        out_stream = File.open(xls_path, "w+")
        MARKER = "<root />"
        index = xml.index("<root />")
        unless index
          out_xml = xml
        else
          out_stream.print(xml[0...index])
          size = File.size(xml_path)
          reader = LibXML::XML::Reader.file(xml_path)
          while reader.read
            local_name = reader.local_name
            next if local_name != "Item"
            xml_text, code = process_xml_node(reader, template_name)
            next if xml_text.empty?
            xml_text = process_text(xml_text)
            out_stream.print(xml_text)
          end
          reader.close
          out_stream.print(xml[(index + MARKER.size + 1)..-1])
        end
        out_stream.close
      else
        reader = LibXML::XML::Reader.string(xml)
        while reader.read
          break if reader.local_name == "root"
        end
        columns_text = reader.read_inner_xml
        size = File.size(xml_path)
        xml_reader = LibXML::XML::Reader.file(xml_path)
        xls_folder = FILEMANAGER["xls"]
        file_name = xml_path
        xls_path = File.join(xls_folder, file_name)
        style_file_path = File.join(xls_folder_path, template_name + "_style.txt")
        style = ""
        if File.file? style_file_path
          style = File.open(style_file_path, "r").read
        end
        tbl_attributes_file_path = File.join(xsl_folder_path, template_name + "_tbl_attr.txt")
        tbl_attributes = ""
        stream = File.open(xls_path, "w+")
        stream.print @xml_tpl_header.format(style, tbl_attributes)
        stream.print columns_text
        while xml_reader.read
          next if xml_reader.local_name != "Item"
          xml_text = process_xml_node(xml_reader, template_name)
          next if xml_text.empty?
          xml_text = process_text(xml_text)
          stream.print xml_text
        end
        xml_reader.close
        stream.print @xml_table_tpl_footer
        stream.print @xml_work_sheet_tpl_footer
        stream.print @xml_workbook_tpl_footer
        stream.close
      end
      return xls_path
      #////////////////
    end
  rescue
    #make exceptions
  end
  
  def process_xml_sep_node reader, template_name, file_name
    prefix = reader.prefix
    suffix = get_template_suffix(prefix)
    index = suffix.index("_")
    throw "suffix" unless index  #make exceptions
    country = suffix[(index + 1)..-1].downcase
    xsl_path = FILEMANAGER["xslt"]
    template_path = File.join(xsl_path, template_name + "_sep.xsl")
    return nil unless File.file? template_path
    xls_folder = FILEMANAGER["xls"]
    file_path = File.join(xls_folder, file_name) # or use FileManager.get_new_file_path
    node_xml = reader.read_outer_xml
    return nil if node_xml.strip.empty?
    xsl = get_item_template(template_path)
    result = []
    leaves, pi_statuses = get_leaf_pi(get_reader(node_xml), suffix)
    leaves.each do |key, value|
      value.each do |pi|
        new_file_name = key + "_" + pi + "_" + file_name
        new_file_name += "_" + pi_statuses[pi].downcase if pi_statuses[pi]
        new_file_path = File.join(xls_folder, new_file_name)
        #FileManager.GetNewFilePath(new_file_path);
        xml_str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        xml_str += "<? mso-application progid=\"Excel.Sheet\"?>\n"
        transform = XML::XSLT.new
        transform.xsl = xsl.format(country.upcase, suffix.downcase, suffix.upcase, key, pi)
        transform.xml = node_xml
        xml_str += transform.serve
        stream = File.open(new_file_path, "w+")
        stream.print xml_str
        stream.close
        result << new_file_path
      end
    end
    return result
  end
  
  def process_text text
    temp = text.dup
    count = 0
    @regex_tpls.each do |key, value|
      regex = Regexp.new(value.format(@group[count]))
      temp.scan(regex).each do |match|
        s = match[2].strip
        folder_path = FILEMANAGER["xslt"]
        file_path = File.join(folder_path, key)
        
      end
      
    end
  end
  
# //////  
  def xsl_transform file_path, template_path
    xslt = XML::XSLT.new

    xslt.xml = LibXML::XML::Smart.open file_path
    xslt.xsl = LibXML::XML::Smart.open template_path

    out = xslt.serve()
    return out;
  end
  def xml_to_xls xml_file
    
  end
  
  def build_hierarchy reader, suffix
    dict = {}
    added_deleted_pi = {}
    while reader.read
      local_name = reader.local_name
      gtin = nil
      action = nil
      case local_name
      when CONSTANTS[:BaseItemVersion]
        if (node = reader.find_descendant("GTIN"))
          dict[node.content.strip] = "BI"
        end
      when CONSTANTS[:AssortmentVersion]
        if (node = reader.find_descendant("GTIN"))
          dict[node.content.strip] = "BA"
        end
      when CONSTANTS[:PackagingItemVersion]
        if (node = reader.find_descendant("GTIN"))
          gtin = node.content.strip
          if (node = reader.find_descendant("ActionRequest"))
            action = node.content.strip
          end
          if (node = reader.find_descendant("GTINofNextLowerPackagingItem"))
            dict[gtin] = node.content.strip
          end
        end
      end
      if local_name == CONSTANTS[:PackagingItemVersion]
        if (node = reader.find_descendant("ActionRequest"))
          action = node.content.strip
        end
        if gtin and action
          if action == CONSTANTS[:AddActionCode].upcase or action == CONSTANTS[:DeleteActionCode].upcase
            added_deleted_pi[gtin] = action
          end
        end
      end
    end
    reader.close
    return added_deleted_pi, dict
  end
  
  def process_to_xls xml_path, template_name
    log_message
    xls_folder_path = "xslt" # or use FileManager
    if File.file? File.join(xls_folder_path, template_name + "_sep.xsl")
      reader = LibXML::XML::Reader.file(xml_path)
      
      while reader.read
        next if reader.local_name != "Item"
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
=end