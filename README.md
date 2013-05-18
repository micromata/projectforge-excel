projectforge-excel
==================

Excel export package (convenient usage of POI) for exporting MS-Excel sheet with a few lines of code or modifiing existing MS-Excel sheets.

Under construction...

## Using log4j or other logging frameworks
ProjectForge's excel module uses the Java standard logging as default. If you need log4j
you may initialize it before using continuous-db:
```java
  Logger.setLoggerBridge(new LoggerBridgeLog4j()); // Before the first log message
```
You may use any other logging framework if you implement the LoggerBridge yourself.
