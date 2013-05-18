projectforge-excel
==================

Excel export package (convenient usage of POI) for exporting MS-Excel sheet with a few lines of code or modifiing existing MS-Excel sheets.

Under construction...

## Creating Excel sheets manually
```java
ExportWorkbook workbook = new ExportWorkbook();
ExportSheet sheet = workbook.addSheet("Test");
sheet.getContentProvider().setColWidths(20, 20);
sheet.addRow().setValues("Type", "result");
sheet.addRow().setValues("String", "This is a text.");
sheet.addRow().setValues("int", 1234);
sheet.addRow().setValues("BigDecimal", new BigDecimal("1042.38"));
sheet.addRow().setValues("Date", new Date());
final File file = new File("test-excel.xls");
workbook.write(new FileOutputStream(file));
```

## Using log4j or other logging frameworks
ProjectForge's excel module uses the Java standard logging as default. If you need log4j
you may initialize it before using continuous-db:
```java
  Logger.setLoggerBridge(new LoggerBridgeLog4j()); // Before the first log message
```
You may use any other logging framework if you implement the LoggerBridge yourself.
