require File.dirname(__FILE__) + '/config/boot'

require "sinatra"
require 'win32ole'

get '/' do
  begin
    xls_path = "C:\\Ruby187\\rails_projects\\excel-clone\\TEST_AREXPORT.xls"
    excel = WIN32OLE.new('Excel.Application')
    sheet = excel.Workbooks.Open(xls_path).Worksheets(1)
    sheet.Range('A1:A3').columns.each { |col| col.cells.each { |cell| puts cell['Value'] } }
    excel.quit
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