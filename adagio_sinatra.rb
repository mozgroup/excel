require File.dirname(__FILE__) + '/config/boot'
require 'sinatra'

get '/' do
  xls_path = "c:\\RadRails\\workspace\\LeoIngwer\\php\\jewelpacServer\\TEST_AREXPORT.xls"
  require 'win32ole'
  begin
    excel = WIN32OLE.new('Excel.Application')
    sheet = excel.Workbooks.Open(xls_path).Worksheets('Sheet1')
    sheet.Range('A1:A3').columns.each { |col| col.cells.each { |cell| puts cell['Value'] } }
  rescue  Exception => e
      puts "XMLRPC error: create Adagio DBF"
      puts e.message
      puts e.backtrace.inspect
  end
end

get '/hello/:name' do
  # matches "GET /hello/foo" and "GET /hello/bar"
  # params[:name] is 'foo' or 'bar'
  "Hello #{params[:name]}!"
end