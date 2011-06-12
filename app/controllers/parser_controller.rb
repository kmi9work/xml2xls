class ParserController < ApplicationController
  def index
  end
  def process_xml
    require 'zip/zip'
    require 'zip/zipfilesystem'
    str = params[:file].read
    name = File.basename(params[:file].original_filename, ".xml")
    xls = OurParser.new(str, params[:template], name)
    prefix = "#{Rails.root}/tmp/xls/"
    if xls.is_sep?
      dir_name = make_dir(prefix + name, 0)
      t = Tempfile.new("zipfile_to_#{request.remote_ip}")
      Zip::ZipOutputStream.open(t.path) do |zos|
        xls.out_xls.each do |h_xml|
          h_xml.each do |filename, xls| 
            filename += ".xls"
            xls.write(File.join(dir_name, filename))
            puts "entry name:#{File.join(name, filename)}"
            zos.put_next_entry(File.join(name, filename))
            puts "zip name: #{File.join(dir_name, filename)}"
            zos.print IO.read(File.join(dir_name, filename))
          end
        end
      end
      send_file t.path, :type => 'application/zip', :disposition => 'attachment', :filename => "#{File.basename dir_name}.zip"
    else
      fname = prefix + "#{name}.xls"
      xls.out_xls.write(fname)
      send_file fname, :filename => "#{name}.xls"
    end
  end
  
  protected
  
  def make_dir name, n
    if File.directory? name
      if File.directory? name + "_#{n}"
        make_dir(name, n + 1) 
      else
        Dir.mkdir name + "_#{n}"
        return name + "_#{n}"
      end
    else
      Dir.mkdir name
      return name
    end
  end
end
