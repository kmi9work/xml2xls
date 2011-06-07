#only Ruby 1.9
#require 'ruby-debug'
require 'nokogiri'
require 'spreadsheet'
require 'mapping'
require 'core_ext'

class OurParser  
  attr_reader :out_xml
  
  def is_sep?
    @is_sep
  end
  
  def initialize xml, template_name, filename
    @filename = file_name
    set_up_vars
    xsl_folder_path = FILEMANAGER["xslt"]
    if File.file? File.join(xsl_folder_path, template_name + "_sep.xsl")
      processed_nodes = sep_transform(xml, template_name)
      @out_xml = processed_nodes.map{|pnode| pnode.map{|n| make_stylesheet(n.to_s)}}
      #make a stylesheet from each element.
    else
      @is_sep = false
      xslt = get_xslt(template_name)
      empty_xml = Nokogiri::XML::Document.new 
      processed_xml = xslt.transform(empty_xml).to_s.strip
      #transform empty xml with "#{template_name}.xsl" xslt
      if File.file? File.join(xsl_folder_path, template_name + "_.xsl") or File.file? File.join(xsl_folder_path, template_name + "_h.xsl") or File.file? File.join(xsl_folder_path, template_name + "_gen.xsl")
        out_xml = suffix_transform(processed_xml, xml, template_name)
        out_xml.gsub!('<?xml version="1.0"?>', "")
      else
        out_xml = style_transform(processed_xml, xml, template_name)
      end
      @out_xml = make_stylesheet(out_xml)
    end
  rescue Exception => ex
    puts ex.message + "OurParser.initialize"
    puts ex.backtrace
    @out_xml = nil
  end

  protected
  
  FILEMANAGER = {
    "xslt" => File.join(Rails.root, "xslt"),
    "xls" => File.join(Rails.root, "tmp")
  }
  
  def sep_transform xml, template_name
    #transform to separate files.
    @is_sep = true
    reader = get_reader(xml)
    processed_nodes = []
    while reader.read
      next if (reader.local_name != "Item")
      processed_nodes << process_xml_sep_node(reader, template_name, xml)
    end
    #processed_nodes -- array of [processed <Item>(xml nodes) arrays of [leaves from get_leaf_pi]] 
    return processed_nodes
  end
  
  def suffix_transform processed_xml, xml, template_name
    #transform with suffixes _, _h, _gen
    index = processed_xml.index(marker = "<root/>")
    index ||= processed_xml.index(marker = "<root />")
    unless index
      out_xml = processed_xml
    else
      out_xml = processed_xml[0...index]
      reader = get_reader(xml)
      while reader.read
        next if (local_name = reader.local_name) != "Item"
        xml_text = process_xml_node(reader, template_name)
        next if xml_text.empty?
        xml_text = process_text(xml_text)
        out_xml += xml_text
      end
    end
    #out_xml = xml before '<root>' + processed xml node <Item> + xml after '<root>'
    return out_xml + processed_xml[(index + marker.size + 1)..-1]
  end
  
  def style_transform processed_xml, xml, template_name
    #transform with style txt's (_style.txt or _tbl_attr.txt)
    reader = get_reader(processed_xml)
    while reader.read
      break if reader.local_name == "root"
    end
    columns_text = reader.inner_xml
    xml_reader = get_reader(xml)
    style_file_path = File.join(xsl_folder_path, template_name + "_style.txt")
    style = File.read(style_file_path) if File.file? style_file_path
    tbl_attributes_file_path = File.join(xsl_folder_path, template_name + "_tbl_attr.txt")
    tbl_attributes = File.read(tbl_attributes_file_path) if File.file? tbl_attributes_file_path
    
    out_xml = @xml_tpl_header.format(style, tbl_attributes)
    #format you may find in ext_core.
    out_xml += columns_text
    while xml_reader.read
      next if xml_reader.local_name != "Item"
      xml_text = process_xml_node(xml_reader, template_name)
      next if xml_text.empty?
      xml_text = process_text(xml_text)
      out_xml += xml_text
    end
    out_xml += @xml_table_tpl_footer
    out_xml += @xml_work_sheet_tpl_footer
    out_xml += @xml_workbook_tpl_footer
  end
  
  def process_xml_node reader, template_name 
    prefix = reader.prefix
    suffix = get_template_suffix(prefix)
    index = suffix.index("_")
    raise OurParserError.new("Wrong suffix: #{suffix}") unless index
    country = suffix[(index + 1)..-1].downcase
    return "" if !suffix or suffix.empty?
    #xslt = XML::XSLT.new
    xsl_file_path = FILEMANAGER["xslt"]
    file_path = File.join(xsl_file_path, template_name + "_.xsl")
    h_xsl_file_path = File.join(xsl_file_path, template_name + "_h.xsl")
    xsl_file_path = File.join(xsl_file_path, template_name + "_gen.xsl")
    node_xml = reader.outer_xml
    return "" if node_xml.strip.empty?
    if File.file? file_path
      xslt = get_item_xslt(file_path, country, suffix) 
    elsif File.file? h_xsl_file_path
      bi_str = ""
      ba_str = ""
      hierarchy_dictionary, dict = build_hierarchy(get_reader(node_xml), suffix) 
      hierarchy_dictionary.each do |key, value|
        bi_str += get_bi_template_part(hierarchy_dictionary, key, suffix) if value == "BI" 
        ba_str += get_bi_template_part(hierarchy_dictionary, key, suffix) if value == "BA"
      end
      xsl = get_item_template(h_xsl_file_path).format(country.upcase, suffix.downcase, suffix.upcase, bi_str, ba_str)
      xslt = Nokogiri::XSLT xsl
    else
      return "" unless File.file? xsl_file_path
      xslt = get_item_xslt(xsl_file_path, country, suffix)
    end
    doc = Nokogiri::XML node_xml
    out = xslt.transform(doc)
    return out.to_s
  end

  

  def process_xml_sep_node reader, template_name, xml
    prefix = reader.prefix
    suffix = get_template_suffix(prefix)
    index = suffix.index("_")
    raise OurParserError.new("Wrong suffix: #{suffix}") unless index
    country = suffix[(index + 1)..-1].downcase
    xsl_path = FILEMANAGER["xslt"]
    template_path = File.join(xsl_path, template_name + "_sep.xsl")
    return nil unless File.file? template_path
    node_xml = reader.outer_xml
    return nil if node_xml.strip.empty?
    xsl = get_item_template(template_path)
    processed_xmls = {}
    leaves, pi_statuses = get_leaf_pi(get_reader(node_xml), suffix)
    leaves.each do |key, value|
      value.each do |pi|
        filename = File.join(FILEMANAGER[:xls], "#{key}_#{pi}_#{@filename}")
        doc = Nokogiri::XML(node_xml)
        xslt = Nokogiri::XSLT(xsl.format(country.upcase, suffix.downcase, suffix.upcase, key, pi))
        processed_xml = xslt.transform(doc).to_s.gsub('<?xml version="1.0"?>','')
        processed_xmls[filename] = processed_xml
      end
    end
    return processed_xmls
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
        mapping = Mapping.read_xml1!(file_path, 0, 1)
        value = ""
        value = mapping[s] if mapping[s]
        new_val = match[0].gsub(">" + s + "<", ">" + value + "<")
        temp.gsub!(match[0], new_val)
      end
      count += 1
    end
    temp.scan(@map_to_regex).each do |match|
      s = match[1].strip
      code = match[4].strip
      first = match[2].strip
      second = match[3].strip
      folder_path = FILEMANAGER["xslt"]
      file_path = File.join(folder_path, s + "_.xml")
      mapping = Mapping.read_xml2!(file_path, first, second)
      value = ""
      value = mapping[code] if mapping[code]
      temp.gsub!(@map_substitution_string.format(s, code, first, second), value) 
    end
    temp.scan(@convert_regex).each do |match|
      s = match[1].strip
      strs = match[6].strip.split(':').delete_if{|i| i.empty?}
      next if strs.size < 2
      code = strs[0]
      value = strs[1].to_f
      tocode = match[2].strip
      first = match[3].strip
      second = match[4].strip
      third = match[5].strip
      if code != tocode
        folder_path = FILEMANAGER["xslt"]
        file_path = File.join(folder_path, s + "_.xml")
        mapping = Mapping.read_xml3!(file_path, first, second, third)
        coeff = ""
        coeff = mapping[code + ";" + tocode] if mapping[code + ";" + tocode]
        next if coeff.empty?
        value *= coeff.to_f
      end
      temp.gsub!(@convert_substitution_string.format(s, match[6].strip, tocode, first, second, third), value.to_s)
    end
    return temp
  end
  
  def set_up_vars
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
    @map_to_regex = Regexp.new "(<mapTo list=\"([^\"]+)\" firstcell=\"([^\"]+)\" secondcell=\"([^\"]+)\">([^<]+)</mapTo>)" #take in brackets whole regexp's 'cause "aaaa".scan(/((a)a(a))/) => [[_"aaa"_, "a", "a"]]
    @convert_regex = Regexp.new "(<mapTo list=\"([^\"]+)\" tocode=\"([^\"]+)\" firstcell=\"([^\"]+)\" secondcell=\"([^\"]+)\" thirdcell=\"([^\"]+)\">([^<]+)</mapTo>)"
    @map_substitution_string = "<mapTo list=\"{0}\" firstcell=\"{2}\" secondcell=\"{3}\">{1}</mapTo>"
    @convert_substitution_string = "<mapTo list=\"{0}\" tocode=\"{2}\" firstcell=\"{3}\" secondcell=\"{4}\" thirdcell=\"{5}\">{1}</mapTo>"

    @group =["(\\D+)", "(\\d+)", "(\\D+)", "(\\D{1,3})"]
    @main_xslts = {}
    @item_xslts = {}
    @item_xslt_templates = {}
    @regex_tpls = {
      "contries.xml"    => "<s:Cell mapTo=\"countries\" (s:StyleID=\"[^>]+\")?><s:Data s:Type=\"String\">{0}</s:Data></s:Cell>",
      "nds.xml"         => "<s:Cell mapTo=\"nds\" (s:StyleID=\"[^>]+\")?><s:Data s:Type=\"String\">{0}</s:Data></s:Cell>",
      "measurement.xml" => "<s:Cell mapTo=\"measurement\" (s:StyleID=\"[^>]+\")?><s:Data s:Type=\"String\">{0}</s:Data></s:Cell>",
      "packaging.xml"   => "<s:Cell mapTo=\"packaging\" (s:StyleID=\"[^>]+\")?><s:Data s:Type=\"String\">{0}</s:Data></s:Cell>"
    }
    xsl_folder_path = FILEMANAGER["xslt"]
    tpl_path = File.join(xsl_folder_path, "Tpl.txt")
    pi_tpl_path = File.join(xsl_folder_path, "PIVersionTpl.txt")
    @tpl_text = File.read(tpl_path)
    @pi_tpl_text = File.read(pi_tpl_path)
    @closest_colors = {}
    @is_sep = nil
  end

  def get_xslt template_name
    return @main_xslts[template_name] if @main_xslts[template_name]
    xsl_folder_path = FILEMANAGER["xslt"]
    xsl_file_path = File.join(xsl_folder_path, template_name + ".xsl")
    raise OurParserError.new("Couldn`t find file with xslt #{xsl_file_path}") unless File.file? xsl_file_path
    xslt = Nokogiri::XSLT(File.read(xsl_file_path))
    @main_xslts[template_name] = xslt
  end

  def get_item_template template_path
    return @item_xslt_templates[template_path] if @item_xslt_templates[template_path]
    @item_xslt_templates[template_path] = File.read(template_path)
  end

  def get_item_xslt file_path, country, suffix
    key = country + ";" + suffix + ";" + file_path
    return @item_xslts[key] if @item_xslts[key]
    xsl = get_item_template(file_path).format(country.upcase, suffix.downcase, suffix.upcase)
    xslt = Nokogiri::XSLT xsl
    @item_xslts[key] = xslt
  end

  def get_template_suffix prefix
    index = prefix.index("_")
    return "" unless index
    return prefix[(index+1)..-1]
  end

  def get_bi_template_part hierarchy_dictionary, bigtin, suffix
    lower = suffix.downcase
    hierarchy = "-PI"
    @str = "" #only for get_pi_text
    get_pi_text(hierarchy_dictionary, bigtin, hierarchy, lower)
    return @tpl_text.format(lower, bigtin, hierarchy_dictionary[bigtin], @str)
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
    return Nokogiri::XML::Reader(xml.to_s)
  end


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

  def build_hierarchy reader, suffix
    #build hierarchy by GTIN's
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
    return dict, added_deleted_pi
  end  

  def get_leaf_pi (reader, suffix)
    dict, added_deleted_pi = build_hierarchy(reader, suffix)
    result = {}
    dict.each do |key, value|
      if is_basic_item(dict, key)
        @leaves = [] #only for check_for_leaves
        check_for_leaves(dict, key, added_deleted_pi)
        result[key] = @leaves
      end
    end
    return result, added_deleted_pi
  end

  def check_for_leaves dict, gtin, adddel
    @leaves ||= []
    children = get_all_children(dict, gtin)
    @leaves << gtin if children.empty? and !is_basic_item(dict, gtin)
    status = adddel[gtin]
    children.each do |child|
      adddel[child] = status if status
      check_for_leaves(dict, child, adddel)
    end
  end

  def is_basic_item dict, gtin
    dict[gtin] == "BI" or dict[gtin] == "BA"
  end

  def get_all_children dict, gtin
    children = []
    dict.each do |key, value|
      children << key if value == gtin
    end
    return children
  end

  def make_stylesheet doc
    book = Spreadsheet::Workbook.new
    xml = (Nokogiri::XML doc).child
    styles = {}
    xml.children.each do |main|
      if main.name == "Styles"
        styles = process_styles(styles, main.children)
      elsif main.name == "Worksheet"
        #process Worksheet
        sheet = book.create_worksheet(:name => main["Name"])
        main.children.each do |worksheet|
          if worksheet.name == "Table"
            #Process Table
            col_index = -1
            row_index = -1
            worksheet.children.each do |col_row|
              if col_row.name == "Column"
                #Process Column
                if col_row["Index"]
                  col_index = col_row["Index"].strip.to_i 
                else
                  col_index += 1
                end
                hidden = col_row["Hidden"].strip.to_bool if col_row["Hidden"]
                width = col_row["Width"].strip.to_f if col_row["Width"]
                if col_row["Span"] # repeat styles span times
                  (col_row["Span"].strip.to_i).times do |i|
                    col = Spreadsheet::Column.new(col_index + i, nil)
                    col.hidden = hidden
                    col.width = width
                    sheet.columns << col
                  end
                else
                  column = Spreadsheet::Column.new(col_index, nil)
                  column.hidden = hidden
                  column.width = width
                  sheet.columns << column
                end
                #/Process Column
              elsif col_row.name == "Row"
                #Process Row
                if col_row["Index"]
                  row_index = col_row["Index"].strip.to_i 
                else
                  row_index += 1
                end
                sheet.row(row_index).height = col_row["Height"].strip.to_i if col_row["Height"]
                cell_index = -1
                sheet.row(row_index).set_format(0, styles["Default"])
                col_row.children.each do |cell|
                  #Process Cell
                  next unless cell.name == "Cell"
                  prev_cell_index = cell_index
                  if cell["Index"]
                    cell_index = cell["Index"].strip.to_i 
                  else
                    cell_index += 1
                  end
                  if cell_index > prev_cell_index + 1
                    #set default style to empty cells
                    i = prev_cell_index + 1
                    while i < cell_index
                      sheet.row(row_index).set_format(i, styles["Default"]) 
                      i += 1
                    end
                  end
                  if cell["StyleID"] and styles[cell["StyleID"]]
                    sheet.row(row_index).set_format(cell_index, styles[cell["StyleID"]]) 
                  else
                    sheet.row(row_index).set_format(cell_index, styles["Default"]) 
                  end
                  cell.children.each do |data|
                    if data["Type"] == "String"
                      sheet.row(row_index).push data.text
                    elsif data["Type"] == "Number"
                      sheet.row(row_index).push data.text.strip.to_i
                    end
                  end
                  #/Process Cell
                end
                #/Process Row
              end
            end
            #/Process Table
          end
        end
        #/process Worksheet
      end
    end
    return book
  end
  
  COLORS = {
    "#00FFFF" => "aqua",
    "#000000" => "black",
    "#0000FF" => "blue",
    "#00FFFF" => "cyan",
    "#A52A2A" => "brown",
    "#FF00FF" => "fuchsia",
    "#808080" => "gray",
    "#808080" => "grey",
    "#008000" => "green",
    "#00FF00" => "lime",
    "#FF00FF" => "magenta",
    "#000080" => "navy",
    "#FFA500" => "orange",
    "#800080" => "purple",
    "#FF0000" => "red",
    "#C0C0C0" => "silver",
    "#FFFFFF" => "white",
    "#FFFF00" => "yellow"
  }
  
  def closest_color str
    return @closest_colors[str] if @closest_colors[str]
    r1 = str[1..2].to_i(16)
    g1 = str[3..4].to_i(16)
    b1 = str[5..6].to_i(16)
    min = nil
    min_hex = nil
    COLORS.each_key do |hex|
      r = hex[1..2].to_i(16)
      g = hex[3..4].to_i(16)
      b = hex[5..6].to_i(16)
      distance = Math.sqrt((r-r1)**2 + (g-g1)**2 + (b-b1)**2)
      min ||= distance
      min_hex ||= hex
      if distance < min
        min = distance
        min_hex = hex
      end
    end
    @closest_colors[str] = COLORS[min_hex]
  end
  
  def make_color str 
    #because gem spreadsheet doesn't recognize hex.
    return COLORS[str] if COLORS[str]
    return closest_color(str)
  end

  def process_styles(styles, xml_styles)
    xml_styles.each do |xml_style|
      id = xml_style["ID"]
      next unless id
      styles[id] = styles[xml_style["Parent"]].dup if xml_style["Parent"] and styles[xml_style["Parent"]]
      styles[id] ||= Spreadsheet::Format.new
      xml_style.children.each do |param|
        case param.name
        when "Alignment"
          param.each do |key, value|
            case key
            when "Vertical"
              styles[id].vertical_align = value
            when "Horizontal"
              styles[id].horizontal_align = value
            end
          end
        when "Borders"
          param.each do |border|
            next unless border.name == "Border"
            border.each do |key, value|
              case key
              when "Bottom"
                styles[id].bottom = value.to_bool
              when "Left"
                styles[id].left = value.to_bool
              when "Right"
                styles[id].right = value.to_bool
              when "Top"
                styles[id].top = value.to_bool
              end
            end
          end
        when "Font"
          name = param["FontName"] ? param["FontName"] : "unnamed"
          font = Spreadsheet::Font.new name
          param.each do |key, value|
            case key
            when "Family"
              font.family = value 
            when "Color"
              font.color = make_color(value)
            when "Bold"
              font.bold = true
            when "CharSet"
            when "FontName"
            when "Size"
              font.size = value.to_i
            when "Italic"
              font.italic = value.to_bool
            when "Underline"
              font.underline = value.to_bool
            end
          end
          styles[id].font = font
        when "Interior"
          param.each do |key, value|
            case key
            when "Color"
              styles[id].pattern_fg_color = make_color(value)
              styles[id].pattern = 1
            when "Pattern"
            when "PatternColor"
            end
          end
        when "NumberFormat"
          param.each do |key, value|
            case key
            when "Format"
              styles[id].number_format = value
            end
          end  
        end
      end
    end
    return styles
  end

end




