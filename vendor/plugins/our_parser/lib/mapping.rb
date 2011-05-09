require 'nokogiri'
class Mapping < Hash
  @xml_docs = {}
  @mappings = {}
  def get_doc file_name
    return @xml_docs[file_name] if @xml_docs[file_name]
    doc = Nokogiri::XML File.read(file_name)
    @xml_docs[file_name] = doc
  end
  
  def get_mapping1 file_name, code_column_index, value_column_index
    key = code_column_index + ";" + value_column_index + ";" + file_name
    return @mappings[key] if @mappings[key]
    mapping = {}
    doc = get_doc(file_name)
    nodes = doc.xpath("/Table/Row")
    nodes = doc.xpath("/mapping/row") if nodes.empty?
    return mapping if nodes.empty?
    nodes.each do |row_node|
      code = row_node.children[code_column_index].text
      next if code.empty?
      code = code.strip
      value = row_node.children[value_column_index].text
      next if value.empty?
      value = value.strip
      mapping[code] = value unless code.empty? and value.empty?
    end
    @mappings[key] = mapping
  end
  
  def get_mapping2 file_name, code_column, value_column
    key = code_column + ";" + value_column + ";" + file_name
    return @mappings[key] if @mappings[key]
    mapping = {}
    doc = get_doc(file_name)
    nodes = doc.xpath("/mapping/row")
    return mapping if nodes.empty?
    nodes.each do |row_node|
      child = row_node.children.filter(code_column).first
      next unless child
      code = child.text
      next if code.empty?
      code = code.strip
      child = row_node.children.filter(value_column).first
      next unless child
      value = child.text
      next if value.empty?
      value = value.strip
      mapping[code] = value unless code.empty? and value.empty?
    end
    @mappings[key] = mapping
  end
  
  def get_mapping3 file_name, code_column, code_column2, value_column
    key = code_column + ";" + value_column + ";" + file_name
    return @mappings[key] if @mappings[key]
    mapping = {}
    doc = get_doc(file_name)
    nodes = doc.xpath("/mapping/row")
    return mapping if nodes.empty?
    nodes.each do |row_node|
      child = row_node.children.filter(code_column).first
      next unless child
      code = child.text
      next if code.empty?
      code = code.strip
      child = row_node.children.filter(code_column2).first
      next unless child
      next if child.text.empty?
      code += ";" + child.text.strip
      child = row_node.children.filter(value_column).first
      next unless child
      value = child.text
      next if value.empty?
      value = value.strip
      mapping[code] = value unless code.empty? and value.empty?
    end
    @mappings[key] = mapping
  end
  
  def read_xml1! file_name, code_column_index, value_column_index
    self = {}
    get_mapping1(file_name, code_column_index, value_column_index).each do |key, value|
      self[key] = value
    end
  end
  
  def read_xml2! file_name, code_column, value_column
    self = {}
    get_mapping2(file_name, code_column, value_column).each do |key, value|
      self[key] = value
    end
  end
  
  def read_xml3! file_name, code_column, code_column2, value_column
    self = {}
    get_mapping3(file_name, code_column, code_column2, value_column).each do |key, value|
      self[key] = value
    end
  end
end