
require File.dirname(__FILE__) + '/config/boot'
require 'spreadsheet'
require 'xmlrpc/server'
require 'csv'
s = XMLRPC::Server.new(8013, 'localhost')

class InvoiceHandler
  def import_tracking_info(uid, pw)
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
  
  def getInvoiceResults(uid, pw)
    invoices = Array.new
    excel_filename = "j:\\zerion\\fromAdagio\\LI\\A"
    #excel_filename = "c:\\RadRails\\workspace\\LeoIngwer\\php\\jewelpacServer\\fromAdagio\\LI\\A.xls"
    book = Spreadsheet.open excel_filename
    sheet1 = book.worksheet 0
    sheet1.each 1 do |row|
      purchase_order  = row[31]
      invoice_number  = row[32]
      
      line_item = [purchase_order, invoice_number]
      invoices << line_item
    end
    
    return invoices
  end
  
  def createLIInvoiceRequests(uid, pw, order_line_items)     
    excel_filename = "j:\\zerion\\toAdagio\\LI\\A"
    #excel_filename = "c:\\RadRails\\workspace\\LeoIngwer\\php\\jewelpacServer\\toAdagio\\LI\\A"
    worksheet_name = "Sheet1"
    
    book = Spreadsheet::Workbook.new
    sheet1 = book.create_worksheet :name => worksheet_name
    
    columns = ['Hdr-Header type', 'Hdr-New invoice', 'Hdr-Customer update type', 'Hdr-Customer code', 'Hdr-Customer name 1', 
'Hdr-Cust addr1/street1', 'Hdr-Cust addr2/street2', 'Hdr-Cust addr3/street3', 'Hdr-Cust address 4', 'Hdr-Customer zip', 
'Hdr-Customer tel', 'Hdr-Customer contact', 'Hdr-Salesperson', 'Hdr-Territory', 'Hdr-Tax exempt 1', 
'Hdr-Tax exempt 2', 'Hdr-Tax group', 'Hdr-Description 1', 'Hdr-Ship to code', 'Hdr-Ship to name 1', 
'Hdr-Ship to name 2', 'Hdr-Ship to addr1/street1', 'Hdr-Ship to addr2/street2', 'Hdr-Ship to addr3/street3', 'Hdr-Ship to address4', 
'Hdr-Ship to zip', 'Hdr-Ship to contact', 'Hdr-Ship to telephone', 'Hdr-Ship via', 'Hdr-Ship to loc', 
'Hdr-Purch. order', 'Hdr-Invoice no', 'Hdr-Invoice date', 'Hdr-Reference', 'Hdr-Header opt 1', 
'Hdr-Header opt 2', 'Hdr-Inv orig', 'Text-Type', 'Text-Code', 'Text-Line 1', 
'Text-Line 2', 'Text-Line 3', 'Text-Line 4', 'Text-Line 5', 'Text-Line 6', 
'Text-Line 7', 'Text-Line 8', 'Text-Line 9', 'Text-Line 10', 'Item type', 
'Item code', 'Item date', 'Reference', 'Opt. field 1', 'Opt. field 2', 
'Item text 1', 'Item text 2', 'Item text 3', 'Item text 4', 'Item text 5', 
'Item text 6', 'Item text 7', 'Item text 8', 'Item text 9', 'Item text 10', 
'Qty Ordered', 'Qty Shipped', 'Qty Back ordered', 'Unit cost', 'Unit price', 
'Extended price', 'Price adjustment %', 'Total before tax', 'Item description']
    
    columns.each{|column|
      sheet1.row(0).push column
    }
    order_line_items.each_with_index{|order_line_item, index|
 
      string_values = order_line_item[0]
      numeric_values = order_line_item[1]
      
      string_values.each{|value|
        sheet1.row(index + 1).push value
        puts ".......value#{value}"        
      }
      
      numeric_values.each{|value|
        sheet1.row(index + 1).push value
        puts ".......value#{value}"        
      }
    }
    book.write "#{excel_filename}.xls"
    true
  end
  
  def createTHTHInvoiceRequests(uid, pw, order_line_items)     
    excel_filename = "j:\\zerion\\toAdagio\\THTH\\A"
    #excel_filename = "c:\\RadRails\\workspace\\LeoIngwer\\php\\jewelpacServer\\toAdagio\\THTH\\A"
    worksheet_name = "Sheet1"
    
    book = Spreadsheet::Workbook.new
    sheet1 = book.create_worksheet :name => worksheet_name
    
    columns = ['Hdr-Header type', 'Hdr-New invoice', 'Hdr-Customer update type', 'Hdr-Customer code', 'Hdr-Customer name 1', 
'Hdr-Cust addr1/street1', 'Hdr-Cust addr2/street2', 'Hdr-Cust addr3/street3', 'Hdr-Cust address 4', 'Hdr-Customer zip', 
'Hdr-Customer tel', 'Hdr-Customer contact', 'Hdr-Salesperson', 'Hdr-Territory', 'Hdr-Tax exempt 1', 
'Hdr-Tax exempt 2', 'Hdr-Tax group', 'Hdr-Description 1', 'Hdr-Ship to code', 'Hdr-Ship to name 1', 
'Hdr-Ship to name 2', 'Hdr-Ship to addr1/street1', 'Hdr-Ship to addr2/street2', 'Hdr-Ship to addr3/street3', 'Hdr-Ship to address4', 
'Hdr-Ship to zip', 'Hdr-Ship to contact', 'Hdr-Ship to telephone', 'Hdr-Ship via', 'Hdr-Ship to loc', 
'Hdr-Purch. order', 'Hdr-Invoice no', 'Hdr-Invoice date', 'Hdr-Reference', 'Hdr-Header opt 1', 
'Hdr-Header opt 2', 'Hdr-Inv orig', 'Text-Type', 'Text-Code', 'Text-Line 1', 
'Text-Line 2', 'Text-Line 3', 'Text-Line 4', 'Text-Line 5', 'Text-Line 6', 
'Text-Line 7', 'Text-Line 8', 'Text-Line 9', 'Text-Line 10', 'Item type', 
'Item code', 'Item date', 'Reference', 'Opt. field 1', 'Opt. field 2', 
'Item text 1', 'Item text 2', 'Item text 3', 'Item text 4', 'Item text 5', 
'Item text 6', 'Item text 7', 'Item text 8', 'Item text 9', 'Item text 10', 
'Qty Ordered', 'Qty Shipped', 'Qty Back ordered', 'Unit cost', 'Unit price', 
'Extended price', 'Price adjustment %', 'Total before tax', 'Item description']
    
    columns.each{|column|
      sheet1.row(0).push column
    }
    order_line_items.each_with_index{|order_line_item, index|
 
      string_values = order_line_item[0]
      numeric_values = order_line_item[1]
      
      string_values.each{|value|
        sheet1.row(index + 1).push value
        puts ".......value#{value}"        
      }
      
      numeric_values.each{|value|
        sheet1.row(index + 1).push value
        puts ".......value#{value}"        
      }
    }
    book.write "#{excel_filename}.xls"
    true
  end
end

class AccountManagementHandler
def test
  require 'win32ole'
  begin
  excel = WIN32OLE.new('Excel.Application')
  sheet = excel.Workbooks.Open('c:\\RadRails\\workspace\\LeoIngwer\\php\\jewelpacServer\\TEST_AREXPORT1.xls').Worksheets('Sheet1')
sheet.Range('A1:A3').columns.each { |col| col.cells.each { |cell| puts cell['Value'] } }
  rescue  Exception => e
      puts "XMLRPC error: create Adagio DBF"
      puts e.message
      puts e.backtrace.inspect
    end
    return true
end

  def getPendingAdagioCustomers(uid, pw)
    customers = Array.new
    #excel_filename = "j:\\zerion\\fromAdagio\\LI\\A"
    excel_filename = "c:\\RadRails\\workspace\\LeoIngwer\\php\\jewelpacServer\\TEST_AREXPORT.xls"
    book = Spreadsheet.open excel_filename
    sheet1 = book.worksheet 0
    sheet1.each 1 do |row|
    puts '.............'
      company_name  = row[3]
      code  = row[2]
      terms  = row[6]
      credit_limit  = row[4]
      tax_id  = row[22]
      created_at  = row[30]
      updated_at  = row[32]
      jbt  = row[31]
      email  = row[40]
      sale_account = ['company_name' => company_name, 'code' => code, 'terms' => terms, 
      'credit_limit' => credit_limit, 'tax_id' => tax_id, 'created_at' => created_at, 
      'updated_at' => updated_at, 'jbt' => jbt, 'email' => email]
      
      first_name  = row[43]
      phone  = row[17]
      fax  = row[18]
      address  = row[10]
      address2  = row[11]
      city  = row[15]
      state  = row[16]
      zipcode  = row[14]
      billing_addr = ['company_name' => company_name, 'code' => code, 'terms' => terms, 
      'credit_limit' => credit_limit, 'tax_id' => tax_id, 'created_at' => created_at, 
      'updated_at' => updated_at, 'jbt' => jbt, 'email' => email]
      
      location  = 'S01'
      store_number  = row[3]
      phone  = row[17]
      fax  = row[18]
      address  = row[10]
      address2  = row[11]
      city  = row[15]
      state  = row[16]
      zipcode  = row[14]
      email  = row[40]
      shipping_addr = ['company_name' => company_name, 'code' => code, 'terms' => terms, 
      'credit_limit' => credit_limit, 'tax_id' => tax_id, 'created_at' => created_at, 
      'updated_at' => updated_at, 'jbt' => jbt, 'email' => email]
      
      first_name  = row[20]
      phone  = row[17]
      fax  = row[18]
      name  = row[20]
      email  = row[42]
      sale_account_contact = ['company_name' => company_name, 'code' => code, 'terms' => terms, 
      'credit_limit' => credit_limit, 'tax_id' => tax_id, 'created_at' => created_at, 
      'updated_at' => updated_at, 'jbt' => jbt, 'email' => email]
      
      tax_status = row[24];
      on_hold = row[5];
      customer_type = row[26];
      account_set = row[0];
      active = row[44];
      territory = row[28];
      
      customers << ['sale_account' => sale_account, 'billing_addr' => billing_addr, 
      'shipping_addr' => shipping_addr, 'sale_account_contact' => sale_account_contact,
      'tax_status' => tax_status, 'on_hold' => on_hold, 'customer_type' => customer_type, 
      'account_set' => account_set, 'active' => active, 'territory' => territory]
    end
    
    return customers
  end
  
  def getPendingAdagioLastMaints(uid, pw)
  end
  
  def getLast5WebDealers(uid, pw)
  end
  
  def confirmNewCust(uid, pw)
  end
  
  def confirmLastMaint(uid, pw)
  end
end

s.add_handler("account_management", AccountManagementHandler.new)
s.add_handler("adagio", InvoiceHandler.new)
s.serve
