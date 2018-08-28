require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe "Spreadsheetx" do
  
  it "opens xlsx files successfully" do
    
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
    
  end
  
  it "allow accessing worksheets" do
    
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
    
    workbook.worksheets.length.should == 2
    workbook.worksheets.last.name.should == 'Test'
    
  end
  
  it "allow accessing row counts" do
    
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)

    workbook.worksheets.last.row_count.should == 8
    
  end
  
  it "can be saved" do
  
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
  
    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_out.xlsx"
    workbook.save(new_xlsx_file)
  
  end
  
  it "can convert an address of a cell to a cell name" do
    
    SpreadsheetX::Worksheet.cell_id(1, 1).should == 'A1'
    SpreadsheetX::Worksheet.cell_id(2, 1).should == 'B1'
    SpreadsheetX::Worksheet.cell_id(27, 9).should == 'AA9'
    SpreadsheetX::Worksheet.cell_id(26, 4).should == 'Z4'
    SpreadsheetX::Worksheet.cell_id(820, 496).should == 'AEN496'
    
  end
  
  it "allows cell values to be updated" do
  
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
  
    workbook.worksheets.last.update_cell(1, 1, 9)
    workbook.worksheets.last.update_cell(1, 2, 'A')
    workbook.worksheets.last.update_cell(1, 3, nil)

    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_changed_out.xlsx"
    workbook.save(new_xlsx_file)
  
  end

  it "allows cells to be added" do
  
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
  
    workbook.worksheets.last.update_cell(9, 9, 9)
    workbook.worksheets.last.update_cell(9, 10, 'A')

    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_added_out.xlsx"
    workbook.save(new_xlsx_file)
  
  end

  it "handles large numbers of rows and cols" do
    
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
  
    500.times do |row|
      6.times do |col|
        random_string = (0...30).map{65.+(rand(25)).chr}.join
        # ump the row because there is no row 0
        workbook.worksheets.last.update_cell((col+1), (row+1), random_string)
      end
    end

    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_large_data.xlsx"
    workbook.save(new_xlsx_file)
  
  end

  it "handles various types of content" do
  
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
  
    workbook.worksheets.last.update_cell(9, 9, Time.now)
    workbook.worksheets.last.update_cell(1, 4, 'A string')
    workbook.worksheets.last.update_cell(9, 10, 'A string')
    workbook.worksheets.last.update_cell(9, 11, 10.3)
    workbook.worksheets.last.update_cell(9, 12, 53)
    workbook.worksheets.last.update_cell(9, 13, nil)

    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_various_content.xlsx"
    workbook.save(new_xlsx_file)
  
  end

  it "can read and return a list of number formats currently in the document" do
  
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
  
    workbook.formats.count.should == 3
    workbook.formats.first.id.to_i.should > 0
    puts workbook.formats.first.format.should == '[$-F400]h:mm:ss\ AM/PM'

    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_various_content.xlsx"
    workbook.save(new_xlsx_file)
  
  end
  

  it "can set formats on cells" do
  
    # a valid xlsx file used for testing
    empty_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec.xlsx"
    workbook = SpreadsheetX.open(empty_xlsx_file)
  
    date_format = workbook.formats.first
    workbook.worksheets.last.update_cell(1, 8, Time.now, date_format)

    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_cell_format.xlsx"
    workbook.save(new_xlsx_file)
  
  end
  
  it "allow accessing worksheet.drawing" do

    # a valid xlsx file used for testing
    target_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_drawing.xlsx"
    workbook = SpreadsheetX.open(target_xlsx_file)

#    workbook.worksheets.length.should == 2
#    workbook.worksheets.last.name.should == 'Test'

  end

  it "allows text box values to be updated" do

    # a valid xlsx file used for testing
    target_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_drawing.xlsx"
    workbook = SpreadsheetX.open(target_xlsx_file)

    workbook.worksheets.first.update_textbox("text box 1", "cahnged_box")

    new_xlsx_file = "#{File.dirname(__FILE__)}/../templates/spec_drawing_output.xlsx"
    workbook.save(new_xlsx_file)

  end

end
