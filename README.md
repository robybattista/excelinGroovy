# excelinGroovy
http://vladimir.orany.cz/excel-in-groovy/

Getting Started
This workshop will learn you how to use Groovy Spreadsheet Builder API to write and read OfficeOpenXML Documents (XSLX). Some of the topics needed for finishing the exercises are written directly in this document but you can always follow Groovy Spreadsheet Builder Documentation for the ones not covered or for deeper insight.

Software Requirements
To get started you need to have installed following:

Java Development Kit 8

OfficeOpenXML Documents (XSLX) compatible application

Although you should be able to display the resulting XLSX file in any compatible application the workshop exercises were so far only tested in Microsoft Excel.
Project is built using Gradle build tool but it’s using Gradle Wrapper so you are not required to install it manually nor you don’t have to be afraid of version clash.

You can use IDE of your choice but IntelliJ IDEA provides so far the best Gradle and Groovy integration even in the free Community version and it will be used in live coding sessions during the workshop.

To get started download exercise project archive file and unzip it.

Business Model
In most of the exercises are using sample sales data from https://www.kaggle.com/kyanyoga/sample-sales-data.

Inside the exercises, the data are represented using following type structure to mock existing application.

Business Model
Test data are stored locally in following CSV file:

excel-in-groovy/eig.model.test/src/main/resources/eig/model/test/testdata.csv
Code Structure
├── eig.model
│   └── src
│       └── main
│           └── groovy
│               └── eig
│                   └── model 
│                       ├── Customer.groovy
│                       ├── Deal.groovy
│                       ├── Order.groovy
│                       ├── OrderLine.groovy
│                       ├── OrderStatus.groovy
│                       └── Product.groovy
├── eig.model.test
│   └── src
│       └── main
│           ├── groovy
│           │   └── eig
│           │       └── model
│           │           └── test 
│           │               ├── TestData.groovy
│           │               └── TestFiles.groovy
│           └── resources
│               └── eig
│                   └── model
│                       └── test 
│                           ├── servicehours.xlsx
│                           ├── testdata.csv
│                           ├── testdata.xlsx
│                           └── testorder.xlsx
└── eig.tasks
    └── src
        ├── main
        │   └── groovy
        │       └── eig
        │           └── tasks
        │               └── ExcelIntegration.groovy 
        └── test
            └── groovy
                └── eig
                    └── tasks 
                        ├── _01_SalesReportSpec.groovy
                        ├── _02_ExecutiveReportSpec.groovy
                        ├── _03_ExportOrderAsExcelSpec.groovy
                        ├── _04_FillDataSpec.groovy
                        ├── _05_HappySpec.groovy
                        ├── _06_ImportOrdersSpec.groovy
                        ├── _07_ImportProductsSpec.groovy
                        └── _08_ImportServiceHoursSpec.groovy
Business Model
Test Utilities
Test Data
Implementation class
Spock Specifications
Each exercise is represented as Spock specification. Each of the specification tests one method in ExcelIntegration class. This is where you are going to write your own code.

Writing Spreadsheets
First 5 exercises will teach you how to write rich XLSX files.

1. Sales Report
Sales Report
Sale manager wants to see all the orders in spreadsheet so she can easily filter particular customers and products. Export all the orders in a spreadsheet, add automatic filter and freeze the first row and column so it’s easily navigable. Make sure the texts displayed fit into the colums and numbers and dates are formatted properly.
Implement method buildSalesReport in ExcelIntegration class to make _01_SalesReportSpec passing. See existing file for reference.

Builder Basics
Spreadsheet builder follows the logical nested structure of the spreadsheet where there is several sheets in the workbook, each of the sheets having several rows and each of the rows containing several cells. The row and column indicies are held internally so there is no need to address rows and cells by their indicies unless you want to skip some rows or columns.

import eig.model.Order
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetBuilder
import org.modelcatalogue.spreadsheet.builder.api.SpreadsheetDefinition
import org.modelcatalogue.spreadsheet.builder.poi.PoiSpreadsheetBuilder

class ExcelIntegration {
    static SpreadsheetDefinition buildSalesReport(Map<Integer, Order> orders) {
        SpreadsheetBuilder builder = PoiSpreadsheetBuilder.INSTANCE         
        return builder.build {                                              
            sheet('Data') {                                                 
                filter auto                                                 
                freeze('A', 1)                                              
                row {                                                       
                    cell {                                                  
                        value 'Order Number'                                
                        width auto                                          
                    }
                    // ...
                }
                // ...
            }
        }
    }
    // ...
}
Obtaining an instance of SpreadsheetBuilder based on Apache POI
Building new spreadsheet and returning SpreadsheetDefinition which can be later written into file or stream.
Creating new sheet called Data
Enabling the automatic filter for all the data in the current sheet
Freezing the first row and the column A
Creating new row
Creating new cell
Assigning Order Number string into the cell
Letting the column to fit into the width of the longest text present in the column
Formatting Values
Format can be specified to any cell within the style definition. See the data formats section of the documentation for more information on topic.

cell {
    value line.sales
    style {
        format '#.00'                                                       
    }
}
cell {
    value Date.from(order.date.toInstant(ZoneOffset.of('+1')))              
    style {
        format 'dd.mm.yyyy hh:mm'                                           
    }
}
Fixing the number of digits of the sales to two digits after the decimal point
Converting the Java 8 LocalDateTime to Date
Printing the date in given format day.month.year hour:minute
Solution
You can verify your solution with the reference one here.

2. Executive Report
Executive Report
Executive officer wants to review states of all orders. Export all the orders to the spreadsheet and highlight every order which hasn’t been shipped yet. For example orders with status resolved should be highlighted with light green color. Don’t repeat the values common to whole order but merge the cells instead. Display the prices with two fixed digits after the decimal point and prefix them with € sign. Render the dates properly. The headlines should be centered, bold and they should have been written using bigger font as rest of the data.
Implement method buildExecutiveReport in ExcelIntegration class to make _02_ExecutiveReportSpec passing. See existing file for reference.

Merging Cells
You can define colspan and rowspan for every cell. Using colspan will also shift he next column pointer so you will automatically continue after the merged cells. On the other hand rowspan will not change the inner next row index so you will continue on the row bellow current one. This fits most of the use cases where the columns are known and data are entered into rows dynamically.

row {                                                                       
    cell {
        value "Grouped"
        rowspan 5                                                           
        colspan 2                                                           
    }
    cell 'Next Cell'                                                        
}
row {                                                                       
    cell('C') {                                                             
        value 'Render this under the "Next Cell"'
    }
}
Create new row 1
Set the rowspan of the cell - the cell will render through the rows 1 to 5
Set the colspan of the cell - the cell will render through the columns A and B
Creates new cell - the cell will be automatically placed into column C
Creates new row 2 - no rows are skipped
Creates new cell in column C to not to collide with the merged cell A1
Named Styles
There is limited number of styles which can be declared within the spreadsheet. For files containing only couple of cell this is not a problem but if you have thousands of rows you can reach the limits easily. It is a good practise to use named styles instead. Named styles are defined in the top level of the builder code and then can be applied to either whole row or a single cell.

See the data formats section of the documentation for more information on topic.

builder.build {
    style('light-green') {                                                  
        foreground lightGreen                                               
    }

    style('dollar') {
        format '$ #.00'
    }

    style('header') {
        font {                                                              
            make bold
            size 72
        }
    }
    style('top-left') {
        align top left                                                      
    }

    sheet {
        row {
            style 'light-green'                                             
            cell {
                value 3.2
                style 'dollar'                                              
            }
        }
    }
}
Declaring the style light-green
Setting the foreground color of the cell (solid fill is applied automatically)
Declaring bold font with size of 72 points
Aligning cell content to top left corner
Using the named style light-green for whole row
Using the names style dollar for single row
Solution
You can verify your solution with the reference one here.

3. Order
Order
Shipping department needs to print an order to each package. Export single order to spreadsheet so it can be printed to A4 paper in portrait orientation. Use formulas whenever possible instead of computed values so the order can be updated manually if needed (i.g. there are not enought items on stock). For EMEA territory print the currency values with € sign and make the outer border double and inner dashed. For the rest of the world print EUR instead and make the border thick and thin respectively.
Implement method buildOrder in ExcelIntegration class to make _03_ExportOrderAsExcelSpec passing. See existing file for EMEA region and the rest of the world for reference.

Page Settings
Each sheet can declare the paper size and orientation:

sheet {
    page {
        paper a5                                                            
        orientation landscape                                               
    }
}
Set the paper size to A5
Set the orientation to landscape
Stylesheets
You can externalize the style definition to separate class so it can be swap easily.

enum SpanishStyles implements Stylesheet {                                  

    INSTANCE

    @Override void declareStyles(CanDefineStyle stylable) {
        stylable.with {                                                     
            style 'border-thick-top', {
                border top, {
                    style thick
                    color black
                }
            }
        }
    }

}
To create a stylesheet, implement Stylesheet interface (you can create enum to make it singleton)
You usually declare styles within the with block so you don’t have to repeat stylable. calls
Then you can apply the styles on top level of the builder code:

builder.build {
    apply SpanishStyles.INSTANCE                                            
}
Apply particular stylesheet on current spreadsheet (this usually happens withing if-else block
Named Cells and Formulas
You can declare names for cells so you can later refer them in formulas more easily.

cell {
    value line.price
    name "price_01"                                                         
}
cell {
    value line.quantity
    name "qty_01"
}
cell {
    formula "#{price_01 * #{qty_01}"                                        
    name "total_01"
}
Declare name of the cell which needs to be unique within the whole spreadsheet
Use the declared name in the formula with #{name}
Solution
You can verify your solution with the reference one here.

4. Sales Charts
Sales Charts
Create two charts displaying cumulated sales by territory and by product line.
Implement method fillData in ExcelIntegration class to make _04_FillDataSpec passing. See existing file for reference.

Beyond the Builder (Charts etc.)
Apache POI as current underlying implementation is not capable of working with charts in spreadsheet files thus the only way how to generate file with chart is to use a template file which refers to data area which is filled with data using the API. The key point is to keep the data area dynamic. For example charts' Y-Axis Area and and Label Area can be pointed to named range which was created using the OFFSET function.

You create new dynamic named range using menu Insert - Name - Define. The formula will be similar to following one:

=OFFSET(Sheet1!$A$1,0,0,Sheet1!$D$1,2)
First parameter is the the very first cell (top left).

Second and third parameter are row and column offset and we can keep them zero all the time.

Third parameter is number of rows to expand the range - this is the number of items printed and it needs to be stored in the spreadsheet somewhere and referenced here in the formula.

Fourth parameter is the column width of the range, e.g. 2 for a range having just two columns.

Syntax of functions and name of the functions varies in different language mutations of MS Excel. Excel function name translation page may be handy to figure out the proper name of the OFFSET function in your locale. In locales where comma (',') is used for decimal point you may need to replaces commas in the formula with semicolons ;.
Named range cannot be used for chart data for the pie chart but two dynamic named ranges can be used for Y-Axis and Labels areas.

Once you have your template spreadsheet file ready with graphs and dynamic ranges you can simply pass it to the build method:

builder.build(templateFile) {                                               
    sheet('Data') {                                                         
        row(2) {                                                            
            cell 'Cars'
            cell 1234.56
        }
    }
}
Existing template is passed into the build method
Either existing sheet Data is matched or new one is created
Existing cells are rewritten or new ones are created in if not present in the template file yet
Solution
You can verify your solution with the reference one here.

5. Pixel Art
Pixel Art
Draw following smiley inside 1 cm grid.

smiley
Implement method drawSmiley in ExcelIntegration class to make _05_HappySpec passing. See existing file for reference.

Using the Definition References
You can obtain the reference to definition objects as the first parameter of definition closure. These objects can be used in situations when another closure would shadow the delegate scope of the closure or when you want to refactor code and extract part of the builder calls into separate methods.

row { RowDefinition rd ->
    10.times {
        rd.cell it
    }
}
Solution
You can verify your solution with the reference one here.

Reading Spreadsheets
Last 3 exercises will show you how you can read files using spreadsheet criteria.

6. Data Migration
Data Migration
As a developer you are asked to import legacy data into new system. You only have data available in the form of spreadsheet (Sales Report). You don’t have to take care about Customers and Products duplicities (these would be handeled by underlying persistence store).
Implement method loadOrders in ExcelIntegration class to make _06_ImportOrdersSpec passing. The input data is similar to Sales Report generated in exercise 1.

Criteria Basics
Spreadsheet criteria follow the very same nested structure as the builder. For example to find a cell with the value Order Number you can write following piece of code:

SpreadsheetCriteria criteria = PoiSpreadsheetCriteria.FACTORY.forStream(inputStream)    
Cell orderNumberHeader = criteria.query {                                               
    sheet {
        row {
            cell {
                value 'Order Number'                                                    
}   }   }   }.cell                                                                      
Integer cellColumn = orderNumberHeader?.column                                          
There is only one implementation at the moment which is based on Apache POI
Create new criteria query
Query for cell will value Order Number within any column, row or sheet
Return just a single cell or null if not found
Return the numeric index of the column (1-based)
There are some handy predicates such as range you can add to method calls such as row. Row and column numbers always starts with 1 as this is the number you can see in the spreadsheet.

SpreadsheetCriteriaResult orderNumbers = criteria.query {
    sheet {
        row(range(2, 5)) {                                                              
            cell('B')                                                                   
}   }   }
Return only rows from 2 to 5
Return only cells in the column B
Data Rows
If you have typical data dump spreadsheet with headers in one row and data in the rest of the rows you can wrap existing Row object into DataRow which allows to retrieve cells using the subscript operator [].

DataRow dataRow = DataRow.create(orderRow, headerRow)
Object customerName = dataRow['Customer Name'].value
Reading Cell Values
When you get the value property of the Cell the return type will correspond the current type of the cell stored in the spreadsheet (e.g. String or Number). This is problematic especially with temporal data as these are stored as numbers as well. To read the values in appropriate type you can use read method of the cell which accepts single parameter which is the desired type.

Cell dateCell = dataRow['Order Date']
Date date = dateCell?.read(Date)
Solution
You can verify your solution with the reference one here.

7. Order Received
Order Received
Your company is ordering goods from ACME Corp. regulary. Order summary spreadsheet is always sent by email when the order is dispatched. You are asked to import the new product quantities into your internal system for each order.
Implement method loadProducts in ExcelIntegration class to make _07_ImportProductsSpec passing. The input data is similar to Order generated in exercise 3.

Navigating the Result
You can easily navigate next or previous sheets using next and previous properties of the sheet.

You can easily navigate next or previous rows using above and bellow properties of the row.

You can easily navigate the cells around selected cell using above, aboveRight, right, bellowRight, bellow, bellowLeft, left and aboveLeft properties of the cell.

Cell qtyHeadline = criteria.query {
    sheet {
        row {
            cell {
                value 'Qty'
            }
        }
    }
}
Integer firstQty = qty.bellow.read(Integer)
Solution
You can verify your solution with the reference one here.

8. LADS Challenge
LADS Challenge
LADS is a company which is accredited to manage all the time tables for public transport in your country. They are obligated by law to make the data available to public. To keep the monopoly they publish the data as spreadsheets containing service hours for each stop. You are asked to read all the departures from the Airport stop for bus line 200 on the working day.

Service Hours for Bus 200
Implement method loadDeparturesFromTheAirportWD in ExcelIntegration class to make _08_ImportServiceHoursSpec passing. The input data can be found at excel-in-groovy/eig.model.test/src/main/resources/eig/model/test/servicehours.xlsx.

Solution
You can verify your solution with the reference one here.

Summary
You have just finished Excel in Groovy workshop which explained you how to use Groovy Spreadsheet Builder to create rich spreadsheet files as well as how to use spreadsheet criteria to query existing ones. Please, leave a feedback in Excel in Groovy Feedback Form.
