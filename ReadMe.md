## List of files ##

**jacob-1.14.3.jar** ,**jacob-1.14.3-x86.dll** - The two files jacob.dll  and  jacob.jar allows you to call COM  components from Java. It uses JNI to make native calls to the COM libraries.

**log4j-1.2.15.jar** - Log4j is an open source project allows the developer to control which log statements are output with arbitrary granularity

**agsTestUtils-QCPlugIn.jar** - JAR file, a wrapper that contains methods to access the OTA API exposed by HP QC. The COM objects of HP QC are accessed through a Java-COM bridge JACOB. The wrapper implements methods to establish connection, create test sets, test run instance, run the tests, update the results, attach the files to test/defects and creates links between tests and defects

**SampleScript.java** - Sample implementation of the QC Plugin JAR for reference. _Note - This file is for reference and is not intended to be used as a solution._


## Details ##

The two files jacob.dll  and  jacob.jar allows you to call COM  components from Java. It uses JNI to make native calls to the COM libraries.

jacob.jar  - must be placed in the classpath

jacob.dll  - must be registered and available in system32 folder

log4j-1.2.15.jar  -  must be placed in the classpath

agsTestUtils-QCPlugIn.jar - must be placed in the classpath