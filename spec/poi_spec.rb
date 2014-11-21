require 'spec_helper'

describe 'POI Initialization' do

  include Poi
  it 'should initialize poi and return a valid object' do

    initialize_poi

    expect(Rjb::import('org.apache.poi.hssf.usermodel.HSSFWorkbook').new).not_to be_nil
  end

end