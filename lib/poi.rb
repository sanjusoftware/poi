require 'poi/version'
require 'rjb'

module Poi

  def initialize_poi
    dir = File.join(File.dirname(File.dirname(__FILE__)), 'jars')
    if File.exist?(dir)
      jardir = File.join(File.dirname(File.dirname(__FILE__)), 'jars', '**', '*.jar')
    else
      jardir = File.join('.','jars', '**', '*.jar')
    end
    Rjb::load(classpath = Dir.glob(jardir).join(':'), jvmargs=['-Djava.awt.headless=true'])
  end

  def poi_output_file(file)
    Rjb::import('java.io.FileOutputStream').new(file)
  end

  def poi_byte_array_output_stream
    Rjb::import('java.io.ByteArrayOutputStream').new
  end

  def poi_input_file(file)
    Rjb::import('java.io.FileInputStream').new(file)
  end

end