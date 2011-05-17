class String
  def format *args
    str = self.dup
    args.each_with_index do |a, i|
      str.gsub!("{#{i}}", a)
    end
    str
  end
  def format! *args
    args.each_with_index do |a, i|
      self.gsub!("{#{i}}", a)
    end
    self
  end
  def to_bool
    !!(self =~ /^(t|1|(true))$/i)
  end
end

module Nokogiri
  module XML
    class Reader
      def find_descendant str
        doc = Nokogiri::XML::DocumentFragment.parse self.outer_xml
        find_in_node doc.children, str
      end
      protected
      def find_in_node nodeset, str
        nodeset.each do |n|
          return n if n.name == str
          node = find_in_node n.children, str
          return node if node
        end
        return nil
      end
    end
  end
end
