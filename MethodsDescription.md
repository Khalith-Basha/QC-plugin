# Introduction #

This page details the methods implemented in the plugin and sample code snippets accessing the same.


## Method Details and sample code snippet ##

### _establishConnection_ ###

The function _establishConnection_   is used to establish the connection with Quality Center by passing QC URL as parameter. This function will return the QC connect instance as Dispatch type.

Below is the snippet of _establishConnection_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/establishConnection.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/establishConnection.png)

### _login_ ###

After getting the QC connection instance, user should login to QC using valid user credentials.  The function _login_ is used to login to QC by passing the QC connection instance and valid username, password. The _login_ function will return Boolean status as True on successful login and False on unsuccessful login.

Below is the snippet of _login_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/login.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/login.png)

### _**connectToProject**_ ###

On successful login, now user should connect to a domain and project in the corresponding domain.  To connect particular project in domain the function _connectToProject_ with QC connection instance, domain name and project name should be passed as parameters. The _connectToProject_ function will return Boolean status as True on successful connection to specified project and False on unsuccessful connection.

Below is the snippet of _connectToProject_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/connectProject.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/connectProject.png)

Once the connection is established for the specified project successfully, using the QC connection instance we can perform the actions like getting test sets, creating test run instance and etc.

### _**getTestSet**_ ###

The function _getTestSet_ is used to get a particular test set form the specified path. It returns the test set instance as Dispatch type. Also we have function _getAllTestSets_  to get all the test sets form the given folder path.

Below is the snippet of _getTestSet_ code:


![https://ags-qcplugin.googlecode.com/svn/trunk/images/getTestSet.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/getTestSet.png)

### _**getAllTestSets**_ ###

The function _getAllTestSets_  is used to get all the test sets for the given folder path. This function return the collection object (Dispatch)

Below is the snippet of _getAllTestSets_ code:


![https://ags-qcplugin.googlecode.com/svn/trunk/images/getAllTestSets.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/getAllTestSets.png)

### _**getTests**_ ###

To  get  all tests form the  test set, the function  _getTests_  with test set as parameter . This function will return all the tests form the test set provided.

Below is the snippet of _getTests_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/getTests.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/getTests.png)

### _**createTestRunInstance**_ ###

After getting all test sets from the given folder path, next step is to create the test run instances for each test . To create the test run instance  use  _createTestRunInstance_ function by passing the test instance. This function will return the test run instance as Dispatch object.

Below is the snippet of _createTestRunInstance_ code

![https://ags-qcplugin.googlecode.com/svn/trunk/images/getTestRunInstance.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/getTestRunInstance.png)

### _**getTestSteps**_ ###

To get the test steps form  the generate test run  use _getTestSteps_  function by passing the test run instance. This function will return the list of steps for the given run instance.

Below is the snippet of _getTestSteps_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/getTestSteps.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/getTestSteps.png)

### _**setTestStepStatus**_ ###

The function _setTestStepStatus_  will update the test results of each test step, the result of test step can be passed as passed ,Failed  and etc,. Also can update the test step actual result.

Below is the snippet of _setTestStepStatus_ code

![https://ags-qcplugin.googlecode.com/svn/trunk/images/setTestStepStatus.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/setTestStepStatus.png)

### _**attachFile**_ ###

To attach a file to a test / test step level  the function  _attachFile_  with attachment path and source object should be passed. This  function will return the status as True on successfully attaching the file to provided source object else  False .

Below is the snippet of _attachFile_ code at **step level**

![https://ags-qcplugin.googlecode.com/svn/trunk/images/fileAttachStepLevel.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/fileAttachStepLevel.png)

Below is the snippet of _attachFile_ code at **test run level**

![https://ags-qcplugin.googlecode.com/svn/trunk/images/fileAttachTestRunLevel.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/fileAttachTestRunLevel.png)

### _**createBugFactory**_ ###

For creating new defects we need bug factory object, to create bug factory call _createBugFactory_  function by passing the QC connection object. This function will return the bug factory as Dispatch object.

Below is the snippet of _createBugFactory_ code

![https://ags-qcplugin.googlecode.com/svn/trunk/images/createBugFactory.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/createBugFactory.png)

### _**logDefect**_ ###

After creating  the bug factory . using the bug factory new defects can be logged in Quality Center. To log new defects use  _logDefect_  function by passing the bug factory  instance , QC user ID, summary, description, priority and severity  of the defect. This function will return the  defect instance as Dispatch type.

Below is the snippet of _logDefect_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/logDefect.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/logDefect.png)

### _**logDefectWithAttachment**_ ###

To log a new defect with attachment use the function _logDefectWithAttachment_. Pass the bug factory  instance ,QC user ID, summary, description, priority ,severity and path of the attachment. This function will return the  defect instance as Dispatch type.

Below is the snippet of _logDefectWithAttachment_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/logDefectWithAttachment.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/logDefectWithAttachment.png)

### _**linkDefect**_ ###

A defect can be linked at run/step level , the function _linkDefect_  by passing the defect instance and the source object. This function will return the status of the linkage.

Below is the snippet of _linkDefect_ code:

![https://ags-qcplugin.googlecode.com/svn/trunk/images/linkDefect.png](https://ags-qcplugin.googlecode.com/svn/trunk/images/linkDefect.png)