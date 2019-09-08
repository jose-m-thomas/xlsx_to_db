# xlsx_to_db

Java Library to Convert Excel document to Database format quickly for data processing, 

Sample code
-----------
```java
File = getExcelFile();
Connection conn = getDbConnection();
Xlsx2DbConverter converter = new Xlsx2DbConverter(file, conn);
converter.convert(true, null);
//
Now you can see all the tables populated with excel data 

```
