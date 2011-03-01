require File.dirname(__FILE__) + '/config/boot'

require "sinatra"
require 'win32ole'
#require 'csv'
require 'json'
require 'rest-client'
require 'mysql'
#require 'rubygems'
require 'active_record'

get '/import_tracking_info' do
  csv_filename = 'j:\\zerion\\Linxship Mailbox\\FTPDaily.csv'
  #csv_filename = 'c:\\RadRails\\workspace\\LeoIngwer\\FTPDaily.csv'
  csv_array = Array.new
  
  CSV.open(csv_filename, 'r', ',') do |row|
    ship_date = row[0]
    order_number = row[1]
    tracking_number = row[2]
    invoice_number = row[3]
    csv_array << ['ship_date' => ship_date, 'order_number' => order_number, 'tracking_number' => tracking_number, 'invoice_number' => invoice_number]
  end
      
  return csv_array
end

get '/invoice_results' do
content_type :json 
  invoices = Array.new
  begin
    #xls_path = "j:\\zerion\\fromAdagio\\LI\\A"
    xls_path = "C:\\Ruby187\\rails_projects\\excel-clone\\A.xls"
    excel = WIN32OLE.new('Excel.Application')
    sheet = excel.Workbooks.Open(xls_path).Worksheets(1)
    sheet.UsedRange.rows.each { |row|
      purchase_order = row.cells(31)
      invoice_number = row.cells(32)
      line_item = [purchase_order['Value'], invoice_number['Value']]
      invoices << line_item }
  rescue  Exception => e
    puts "XMLRPC error: create Adagio DBF"
    puts e.message
    puts e.backtrace.inspect
  ensure
    excel.quit
  end
    
  return invoices.to_json
end

get '/' do
  begin
    xls_path = "C:\\Ruby187\\rails_projects\\excel-clone\\TEST_AREXPORT.xls"
    excel = WIN32OLE.new('Excel.Application')
    sheet = excel.Workbooks.Open(xls_path).Worksheets(1)
    sheet.Range('A1:A3').columns.each { |col| col.cells.each { |cell| puts cell['Value'] } }
  rescue  Exception => e
    puts "XMLRPC error: create Adagio DBF"
    puts e.message
    puts e.backtrace.inspect
  ensure
    excel.quit
  end
end

class TimesheetItem < ActiveRecord::Base
  set_table_name 'timesheetItem'
end

get '/hello/:name' do
  # matches "GET /hello/foo" and "GET /hello/bar"
  # params[:name] is 'foo' or 'bar'
  
  TimesheetItem.establish_connection(
    :adapter => 'mysql',
    :database => 'timesheet',
    :username => 'apache',
    :password => 'ItisJustTest',
    :host => '192.168.12.5'
  ) 
  @day = Date.today.day
  @day = @day - 15 if @day > 15
  @day = @day - 1 if @day > 1
  start_date = '2011-02-16'
  @timesheet_items = TimesheetItem.find_by_sql("select 
     day#{@day}Hour as hour1, day#{@day + 1}Hour as hour2,
     day#{@day + 2}Hour as hour3, day#{@day + 3}Hour as hour4,
     day#{@day + 4}Hour as hour5, day#{@day + 5}Hour as hour6,
     day#{@day + 6}Hour as hour7, day#{@day + 7}Hour as hour8,
     day#{@day + 8}Hour as hour9, day#{@day + 9}Hour as hour10,
     day#{@day + 10}Hour as hour11, day#{@day + 11}Hour as hour12,
     day#{@day + 12}Hour as hour13, day#{@day + 13}Hour as hour14,
     day#{@day + 14}Hour as hour15,
     client.name as client_name, project.name as project_name, 
     task.name as task_name, employee.userName as employee_name, parent.name as parent_name,
     day1Hour + day2Hour + day3Hour + day4Hour + day5Hour + day6Hour + day7Hour + day8Hour + day9Hour + day10Hour + 
     day11Hour + day12Hour + day13Hour + day14Hour + day15Hour + day16Hour as total_hours
     from client, project, timesheetItem, timesheetInterval, employee, task
     left join task parent on task.pid = parent.id
   where client.id = project.clientID
     and project.id = task.projectID
     and task.id = timesheetItem.taskID
     and timesheetItem.employeeID = employee.id
     and timesheetItem.intervalID = timesheetInterval.id
     and timesheetInterval.startDate = '#{start_date}'
     and (task.name = 'Maintainenance & Update' or task.name = 'FISMA Compliance')
   order by employee.userName")
  @timesheet_items.each{|timesheet_item|
    puts "#{timesheet_item.hour1} #{timesheet_item.hour2} #{timesheet_item.hour3} #{timesheet_item.hour4}"
    puts "#{timesheet_item.hour5} #{timesheet_item.hour6} #{timesheet_item.hour7} #{timesheet_item.hour8}"
    puts "#{timesheet_item.hour9} #{timesheet_item.hour10} #{timesheet_item.hour11} #{timesheet_item.hour12}"
    puts "#{timesheet_item.hour13} #{timesheet_item.hour14} #{timesheet_item.hour15}"
    puts "#{timesheet_item.total_hours}hrs #{timesheet_item.task_name} #{timesheet_item.employee_name}" } 
  
  jdata = "Hello #{params[:name]}!"
  jdata = "<foo><name>do</name></foo>"
  jdata = ["start_date" => start_date]
  jdata << @timesheet_items
  resource = RestClient::Resource.new 'http://192.168.12.121:3000/tests/index'
  #resource.put {:data => jdata}, :content_type => 'application/xml', :accept => :html
  #resource.put {:data => jdata, :content_type => :json, :accept => :html}
  RestClient.post 'http://192.168.12.121:3000/tests/index', jdata.to_json, {:content_type => :json, :accept => :html}
  #resource.get :params => {:id => 3, 'foo' => 'bar'}, :content_type => :xml, :accept => :html
end

class Crack
  def initialize(app)
    @app = app
  end

  def call(env)
    puts '------------------------------------'
    response = @app.call(env)
    ActiveRecord::Base.clear_active_connections!
    puts '..................................'
    response
  end
end

use Crack