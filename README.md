# Excel-To-Sql-Server-Convertor

Push database data from Excel to Sql Server using a console application

## Usage:

1 . Create an Excel file with the data that you want to insert into the database
 
 - Each worksheet represents a table, with the table name as the sheet name 
 - The columns in each worksheet represent the columns in the table
 - The rows represent the data in the table

![image](https://user-images.githubusercontent.com/40519064/98439678-729c7300-212e-11eb-873d-7c3a69b47847.png)

2 . Add the location of the file in the program

```string filePath = @"C:\_MyDotNetApplications\ExcelToSqlServerConvertor\Test.xlsx";```

3 . Create a database in Sql Server and add the connection string to your program

```xml
<connectionStrings>
  <add name="ExcelToSqlServerConvertorTestDb" connectionString="data source=.; database=ExcelToSqlServerConvertorTestDb; integrated security=SSPI" providerName="System.Data.SqlClient"/>
</connectionStrings>
```

4 . You can retrieve the connection string using configuration manager

```string connectionStr = ConfigurationManager.ConnectionStrings["ExcelToSqlServerConvertorTestDb"].ToString();```

5 . Ensure there are already tables in your database

 - The tables need to correspond to the worksheet schema in your Excel file
 
![image](https://user-images.githubusercontent.com/40519064/98439750-f0f91500-212e-11eb-8508-cbeb6c42d96f.png)
 
6 . Before running the program, delete any existing data from your tables in the database

## Example result of running the application:

![image](https://user-images.githubusercontent.com/40519064/98439795-39183780-212f-11eb-89f3-1e434dab0820.png)

## TODO:

1. Create a GUI version of application with configurable parameters i.e. connection string, file location
2. Add code to delete any existing table data before insertion
 
