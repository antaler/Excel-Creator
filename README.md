# *PREVIEW* Excel-Creator

Easily generate an excel file
```xml
<dependency>
  <groupId>com.antonioalejandro</groupId>
  <artifactId>Excel-Creator</artifactId>
  <version>1.0</version>
</dependency>
```

## How Implement Interface

Implement this interface ``` IExcelObject ``` with a method:

- ``` obtainFields() ```: return a List of ExcelData Object in the same order that introduce the headers values.  

### ExcelData Object

This Object have a overloaded constructor. The values that can accept are strings, Integers, Long, Double, Date and boolean. The boolean data can insert custom string values (true and false are default values) like an example with gender.

```java
@ExcelItem
public class Example {
    
    @ExcelColum(order = 1)
    private String name;
    
    @ExcelColum(order = 2)
    private String lastname;
    
    @ExcelColum(order = 3)
    private Integer age;

    @ExcelColum(order = 4, trueValue= "Man", falseValue="Woman")
    private boolean isMan;
}
```

## How create a Excel File

### Create a ExcelBook Object

The constructor are overloaded. You can put the sheetname only, the sheetname and headers or sheetname,headers, and the list object(excel data).

```java
public static void main(String[] args) {
    List<String> headers = Arrays.asList("Name","LastName", "Age","Gender");
    //the data
    List<Example> data = Arrays.asList(new Example("Name 1","Lastname 1", 10, true),new Example("Name 2","Lastname 2", 20, false));
    ExcelBook<Example> excelBook = new ExcelBook<Example>("Example");
    // Set headers values
    excelBook.setHeaders(headers);
    // Set Data
    excelBook.setData(data);
    // you can set color to headers row use java.awt.Color object
    excelBook.setHeaderColor(new Color(125,125,125));
    // you can set color to data cells
    excelBook.setDataColor(new Color(255,255,255));
    // you can remove cells borders
    excelBook.setBlankSheet();
    //  if the excel contain dates is Data you can set a format. the formats are specified in ExcelCellDateFormat enum
    excelBook.setFormatDate(ExcelCellDateFormat.NUMBER_SHORT_WITH_TIME);
    // you can set a imagen like logo. you add image like byte[] or Input Stream
    excelBook.addLogo(/*byte[] or inputStream*/);
    // when you are to prepare to send file o write file in pc. you have two options
    // write file in yout pc
    try{
        excelBook.write("path in your computer"/*Example: /home/user/Desktop/ejemplo(name file without extension)*/);
    }catch(IOException e){}
    //or recive a byte[]
    try{
        byte[] bytes = excelbook.prepareToSend();
    }catch(IOException e){}
    // you have to close the Excelbook to create safely
    excelBook.close(); // throws IOException
}
```

## Create your own ExcelBook

The class ExcelBook is a predifiend implementation. So you can extends ExcelBookAbstract to create your own custom class.
