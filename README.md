# ExcelR.DotNetCore 

Simple dotnet core library contains  to create/read xlsx/csv files in an easy way.

# Why ExcelR
With the help of ExcelR you can
- Read or write file from/to disk or stream
- Write you list of objects to file 
- Read data from file to your model
- Control the color/font of column write to the xlsx file
- Set diffrent heading style
- Can read custom property name with help of excelprop attribute


# Examples:-
* Lets we have test class and some sample data as follow
  
   ```
     public class TestModel
    {
        [ExcelRProp(Name = "First Name")]
        public string FirstName { get; set; }


        [ExcelRProp(ColTextColor = "Red", Name = "Last Name")]
        public string LastName { get; set; }

        [ExcelRProp(SkipExport = true)]
        public bool IsMale { get; set; }

        [ExcelRProp(HeadTextColor = "Blue" ,Name = "Date Of Birth")]
        public DateTime? Dob { get; set; }

    }
    var sampleData =   var list = new List<TestModel>
            {
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "Braat", LastName = "Lee"},
                new TestModel {IsMale = true,  FirstName = "Flintop"},
                new TestModel {IsMale = true, Dob = DateTime.Now.AddDays(15), FirstName = "Michel"},
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "Michel", LastName = "John"},
                new TestModel {IsMale = false, FirstName = "john", LastName = "Cena"}
            };
    ```
## Write and save data to xlsx file

#### Method1:-
```
sampleData.ToExcel().Save(filePath);
   ```

#### Method2:-


* Get worksheet and write data to sheet as follow
   ```
         var sheet = ExcelExporter.GetWorkSheet();//you can pass custom sheet 
         name 
        //File data in the sheet
         sheet.Write(dataToWrite);
    ```
* Save sheet to stream or disk
   ```
    var stream = sheet.ToStream();
    sheet.Save(filePath);
   ```

## Write and save data to csv file
 ```
sampleData.ToCsv(filePath);
 ```
   
## Read data from xlsx file or stream
* Get worksheet from file or stream
   ```
   var workSheet=ExcelImporter.GetWorkSheet(filePath)
                     Or
    var workSheet=ExcelImporter.GetWorkSheet(stream)
   ```
 * Read data from sheet
   ```
   var data= ExcelImporter.Read<TestModel>(sheet);
   ```
## Read data from csv file
```
 var data = CsvHelper.ReadFromFile<TestModel>(sourceFilePath);
```
   
## Manually creating xlsx from complex models

   ```
        var sheet = ExcelExporter.GetWorkSheet();//you can pass custom sheet name 
         var rowNo=0;
         //create header row
         var headerRow = sheet.CreateRow(rowNo++,Style.H1);
         //Set header values
         headerRow.SetValue(0,"String property")
         //Create data rows and fill data
         foreach(var item in sampleData){
         var dataRow = sheet.CreateRow(rowNo++,Style.H1);
         dataRow.SetValue(0,item.StringProp);
         }
         
         //Save to  file
         sheet.Woorkbook.Save(filePath);
         //Output to stream
          sheet.Woorkbook.ToStream();
         
   ```
 # Facing any issue [Log it here](https://github.com/tech-farmz/ExcelR.DotNetCore/issues/new)