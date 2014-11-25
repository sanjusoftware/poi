require 'spec_helper'

describe 'POI Initialization' do

  include Poi
  it 'should initialize poi and return a valid object' do

    initialize_poi

    expect(Rjb::import('org.apache.poi.hssf.usermodel.HSSFWorkbook').new).not_to be_nil
  end

  it 'should create a FileOutputStream object' do
    expect(Rjb::import('java.io.FileOutputStream').new('foo').java_methods).to eq(poi_output_file('foo').java_methods)
    File.delete 'foo'
  end

  it 'should create a byteArrayOutputStream object' do
    expect(Rjb::import('java.io.ByteArrayOutputStream').new.java_methods).to eq(poi_byte_array_output_stream.java_methods)
  end

  it 'should create a input file object' do
    file = File.join(File.dirname(File.dirname(__FILE__)), 'spec', 'data', 'image001.jpg')
    expect(Rjb::import('java.io.FileInputStream').new(file).java_methods).to eq(poi_input_file(file).java_methods)
  end

end