
/*
 * Copyright 2012 Alliance Global Services, Inc. All rights reserved.
 *
 * Licensed under the General Public License, Version 3.0 (the "License") you may not use this file except in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.gnu.org/licenses/gpl-3.0.txt
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
 *
 * Class: QCPlugIn
 * 
 * Purpose: This class contains the implementation of  functions like establishing the connections with QC, login to QC and connect to specified project ,creating the test run instance and update the test step results, attach the file to test and defect, create the new defect and link the defect to corresponding test.
 */
package com.agstestutils.qcplugin;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.*;
import org.apache.log4j.Logger;


public class QCPlugIn {
	
	static Logger logger = Logger.getLogger(QCPlugIn.class);
	
	/**
	 * This method is used to establish the QC connection. 
	 * @param url is the URL of QC server  
	 * @return On establishing successful connection with QC  the connection instance will be returned
	 * @author Srikanth Kasam
	 **/ 
	public ActiveXComponent establishConnection(String   url){
		
		logger.debug("Starting connection with Quality Center");		 
		ActiveXComponent QCObj = null;		
		// Establishing the connection with Quality Center 
		try {	
			  logger.info("Connecting Quality center URL :"+ url);
			  QCObj = new ActiveXComponent("TDApiOle80.TDConnection");
			  Dispatch.call(QCObj,"initConnectionEx",url);			  
			  logger.debug("Connection established with Quality Center");
			  
		  } catch (ComFailException e) { 
			  logger.error("Unable to establish the connection with Quality Center for URL: " +url +
			  			".Please check the URL and pass valid URL.:: "+ e.getMessage());			  
		  }
		  catch(ComException e)  {
			  logger.error("Unable to establish the connection with Quality Center for URL: " +url +
			  			".Please check the URL and pass valid URL.:: "+ e.getMessage());			  
		  }	
		  
		return QCObj;
	}
	
	/** 
	 * This method is used to Login to QC server by passing valid QC username and QC password
	 * @param objQCconn is the object of QC connection  
	 * @param userName is the valid user name to login 
	 * @param password is the valid password to login 
	 * @return  The status of login is returned as True / False 
	 * @author Srikanth Kasam
	 **/
	public boolean login(ActiveXComponent objQCconn, String userName, String password) {
		
		boolean loginStatus = false;
		logger.debug("Logging into Quality Center");		
		// Logging into Quality Center
		try {
		logger.info("Calling login function with Username :"+userName +" and Password :" + password);
		Dispatch.call(objQCconn,"login",userName,password);
		
		if (objQCconn.getPropertyAsBoolean("LoggedIn") && (Dispatch.call(objQCconn,"UserName").toString()).equalsIgnoreCase(userName)) {	
			logger.info("Successfully Logged into Quality Center");
			loginStatus = true;
		 }
		else {
			 logger.info("Failed to login into Quality Center using  Username :"+userName +" and Password :" + password);
			 loginStatus = false;
			 logger.debug("Login failed using Username :"+userName +" and Password :" + password);
		  }
		logger.debug("Status of login to QC is "+ loginStatus);
		
		  } catch (ComFailException e) { 
			  logger.error("Invalid QC login credentials. Please check the username and " +
		  				" password and try to login again ::"+e.getMessage());
		  } catch(ComException e)  {
				  logger.error("Invalid QC login credentials. Please check the username and " +
				  				" password and try to login again ::"+e.getMessage());
			  
		  }	
		
		return loginStatus;	
	}
	
	/** 
	 * This method is used for connecting the  specific project in QC 
	 * @param objQCconn is the object of QC connection
	 * @param domain is the name of the Domain  
	 * @param project is the name of the Project to connect		 
	 * @return The status of connection to project is returned as True / False
	 * @author Srikanth Kasam 	 
	 **/		
	public boolean connectToProject(ActiveXComponent objQCconn,String domain, String project){
		
		boolean projectConnStatus = false;
		logger.debug("Connecting the domain and project");
		// Connecting to project in the domain 
		try	{	
			logger.info("Connecting to project : " +project +" in domain :"+domain);
			Dispatch.call(objQCconn,"connect",domain , project);
		
		if (objQCconn.getPropertyAsBoolean("connected") && (Dispatch.call(objQCconn,"ProjectName").toString()).equalsIgnoreCase(project))	{
			logger.info("Successfully connected to project: "+ project);
			projectConnStatus = true;
		 }
		 else {	
			 logger.info("Failed to connect the project : " +project);	
			 projectConnStatus = false;
		  }
		logger.debug("Status of project connection is :" +projectConnStatus);
		  } catch (ComFailException e) { 
			  logger.error("Failed to connect the Domain :" + domain+ " and Project :"+project +
				  		 " . Please check the domain and project. ::" +e.getMessage());
		  } catch(ComException e)	 {
			  logger.error("Failed to connect the Domain :" + domain+ " and Project :"+project +
					  		 " . Please check the domain and project. ::" +e.getMessage());			  
		  }
		  		
		return projectConnStatus;	
	}
	
	/**
	 * This method is used to generate the test set for the specified folder path and name
	 * @param folderPath is the folder path of the test set  
	 * @param testSetName is the test set name	  
	 * @return  The Test set from the specified folder path  is returned
	 * @author Srikanth Kasam 
	 **/		
	public Dispatch getTestSet(ActiveXComponent objQCconn, String folderPath, String testSetName ) {
		
		Dispatch theTestSet = null;	
		Dispatch tsList = null;
		 // Getting the test set from the given folder path 
		logger.debug("Getting the test set :"+ testSetName+ " from the folder path : "+ folderPath );
		  try {
			// Getting the test set tree manager from the test set factory
			  logger.info("Creating the test set factory");
			  Dispatch.call(objQCconn, "testsetFactory").toDispatch();
			  logger.info("Getting the test set tree manager form the test set factory");
			  Dispatch tsTreeMgr = Dispatch.call(objQCconn,"TestSetTreeManager").toDispatch();
			  logger.info(" Successfully retrived the test set tree manager form the test set factory ");
			
			  //Use TestSetTreeManager.NodeByPath  to get the test set folder
			  logger.info("Getting the test set folders from the test set tree manager");
			  Dispatch tsFolder = Dispatch.call(tsTreeMgr, "NodeByPath", folderPath).toDispatch();
			  logger.info("Successfully got the test set folders ");
			  
			  logger.info ("Finding the test set folder :"+testSetName+ " form the test set tree");
			  tsList = Dispatch.call(tsFolder,"FindTestSets",testSetName).toDispatch();			  
			  theTestSet =  Dispatch.call(tsList, "Item","1").toDispatch();		  
			  logger.info("Found the test set folder :"+ Dispatch.get(theTestSet,"name").toString());
			  				 
		}
		  catch (ComFailException e) { 
			  logger.error("Unable to find the test set folder :" + testSetName+ " in the given folder paht: "+folderPath +
				  		 " . Please verify the test set name and folder path. ::" +e.getMessage());
		  } 
		  catch (ComException e) {
			  logger.error("Unable to find the test set folder :" + testSetName+ " in the given folder paht: "+folderPath +
				  		 " . Please verify the test set name and folder path. ::" +e.getMessage());		  
		}
		 
		return theTestSet;
	}
	
		/**
		 * This method is used to retrieve all test sets from the specified folder path.
		 * @param objQCconn is the QC Connection instance 
		 * @param folderPath is the folder path of the test set  		  
		 * @return Returns a collection of Test sets from the specified folder path
		 * @author Srikanth Kasam  
		 **/		
		public Dispatch getAllTestSets(ActiveXComponent objQCconn, String folderPath ) {
			
			Dispatch testSetsList = null;
			// Getting all test sets form the given folder path 
			  try {
				  // Get the test set tree manager from the test set factory
				  logger.info("Creating the test set factory");
				  Dispatch.call(objQCconn, "testsetFactory").toDispatch();
				  logger.info("Getting the test set tree manager form the test set factory");
				  Dispatch tsTreeMgr = Dispatch.call(objQCconn,"TestSetTreeManager").toDispatch();
				  logger.info(" Successfully retrived the test set tree manager form the test set factory ");
				  
				  //Use TestSetTreeManager.NodeByPath  to get the test set folder
				  logger.info("Getting the test set folders from the test set tree manager");
				  Dispatch tsFolder = Dispatch.call(tsTreeMgr, "NodeByPath", folderPath).toDispatch();				  
				  logger.info("Successfully got the test set folders ");
				  
				  //To Get all  test set to Variant	
				  logger.info ("Creating the test set folder factory");
			  	  Dispatch testSetFolderFact = Dispatch.call(tsFolder, "TestSetFactory").toDispatch();  			  
				  testSetsList = Dispatch.call(testSetFolderFact,"NewList"," ").toDispatch();
				  logger.info("Got the test sets from the given path :" + folderPath);
				  
				  //Our result is a collection, so we need to work though the collection.
					EnumVariant enumVariant = new EnumVariant(testSetsList);
					Dispatch folder = null;
						while (enumVariant.hasMoreElements()) { 
							folder = enumVariant.nextElement().toDispatch();
							logger.info("Test set folder name :: " + Dispatch.call(folder, "Name").toString());
						}
					logger.debug("Got list of all test sets from specifed folder path");
			} catch (ComFailException e) { 
				  logger.error("Unable to get all test set folders from the path :" + folderPath +
					  		 " . Please verify the folder path. ::" +e.getMessage());
			}catch (ComException e) {
				  logger.error("Unable to get all test set folders from the path :" + folderPath +
					  		 " . Please verify the folder path. ::" +e.getMessage());		  
			}
			 
			return testSetsList;
		}
		
	/** 
	 *  This method is used to get all the tests from the test set provided.
	 * 	If needed this function should be iterated through all the test sets  
	 *  which are returned from getAllTestSets to get all the tests form the test set.
	 * @param theTestSet is the  Test Set instance		  
	 * @return  The tests object from the specified Test sets
	 * @author Srikanth Kasam   
	 **/
	public Dispatch getTests(Dispatch theTestSet) {
		
		Dispatch testList = null;	
		Dispatch TSTestFact = null;
		//Get the test instances from the test set
		logger.debug("Starting retriving the tests for the test set :"+Dispatch.get(theTestSet,"name").toString());
		  		 
		try {
			logger.info("Getting the tests factory from the test set :"+Dispatch.get(theTestSet,"name").toString());					  
		  	TSTestFact = Dispatch.call(theTestSet, "TSTestFactory").toDispatch();
		  	logger.info("Successfully created the tests factory form the test set");
		  	
			// Getting the tests list
		  	logger.info("Getting the test list form the test set : " + Dispatch.get(theTestSet,"name").toString());
			//Dispatch tsFilter = Dispatch.call(TSTestFact,"Filter").toDispatch();
			//Dispatch.call(tsFilter, "Filter","TC_CYCLE_ID").putInt(Dispatch.call(theTestSet,"ID").changeType((short) 3).toInt());
			testList = Dispatch.call(TSTestFact,"NewList"," ").toDispatch();
			logger.info("Successfully retrived the list of test cases form the test set :"+ Dispatch.get(theTestSet,"name").toString());
			
			//Our result is a collection, so we need to work though the.
			EnumVariant enumVariant = new EnumVariant(testList);
			Dispatch test = null;
				while (enumVariant.hasMoreElements()) { 
					test = enumVariant.nextElement().toDispatch();
					logger.info("Test details : Test ID :"+Dispatch.call(test,"ID")+" || Test Name :"+
					Dispatch.call(test,"Name"));					
				}
			logger.debug("Got tests list form the given test set : " +Dispatch.get(theTestSet,"name").toString());
				
		}catch (ComFailException e) { 
			  logger.error("Unable to get the tests from the test set :" + Dispatch.get(theTestSet,"name").toString()+
				  		 " . Please verify the test set ::" +e.getMessage());
		}catch (ComException e) {
			  logger.error("Unable to get the tests from the test set :" + Dispatch.get(theTestSet,"name").toString()+
				  		 " . Please verify the test set ::" +e.getMessage());
		}	
		
		return testList;
	}
	
	/** 
	 * This method is used to create the test run instance and copies the design steps for 
	 * the specific test instance provided. This function should be iterated through all the
	 * list of tests returned form getTests function.
	 * @param objTest is the instance of Test		  
	 * @return The test run instance of specified Test  
	 * @author Srikanth Kasam	 
	 **/
	public Dispatch createTestRunInstance(Dispatch objTest){
		
		Dispatch newRun = null;
		// Creating the test run instance
		logger.debug("Creating the test run for test :" +Dispatch.call(objTest,"Name"));
		try {
			//Properties of Test instance
			logger.info("Creating the test run for test : " +" Test ID : "+Dispatch.call(objTest,"ID")+
					"|| Test Name :" +Dispatch.call(objTest,"Name")+" || Test Type: "+Dispatch.call(objTest,"Type"));
			
			logger.info("Creating the run factory ");
			Dispatch RunF = Dispatch.call(objTest, "RunFactory").toDispatch();
			
			//Creating the test run
			logger.info("Creating the test run ");
			newRun = Dispatch.call(RunF,"AddItem","Null").toDispatch();
			Dispatch.put(newRun,"Status","No Run");
			Dispatch.put(newRun,"Name","TRun_" +System.currentTimeMillis());
			
			//Coping the test run design steps
			Dispatch.call(newRun,"CopyDesignSteps");
			Dispatch.call(newRun,"Post");			
			logger.info("Successfully created test run : " +" Test Run ID : "+Dispatch.call(newRun,"ID")+
					"|| Test Run Name :" +Dispatch.call(newRun,"Name"));
			
		}catch (ComFailException e) { 
			  logger.error("Unable to create the test run for the test  : Test ID : "+Dispatch.call(objTest,"ID")+
						"|| Test Name :" +Dispatch.call(objTest,"Name")+" . Please verify the test::" +e.getMessage());
		}catch (ComException e) {
				logger.error("Unable to create the test run for the test  : Test ID : "+Dispatch.call(objTest,"ID")+
						"|| Test Name :" +Dispatch.call(objTest,"Name")+" . Please verify the test::" +e.getMessage());
		}	
						
		return newRun;
	}
	/** 
	 *	This method is used to get the test steps factory for the specified test run 
	 *	@param newRun is the instance of test run		  
	 * 	@return The Steps of specified test run
	 * 	@author Srikanth Kasam 
	 **/
	public Dispatch getTestSteps(Dispatch newRun) {	
		
		Dispatch stepsList = null;
		//Getting the test steps
		logger.debug("Getting the test steps form test run : Test Run Name :" +Dispatch.call(newRun,"Name"));
		try {
			logger.info("Creating steps factory for the current run");
			Dispatch runStepF = Dispatch.call(newRun ,"StepFactory").toDispatch();
			stepsList = Dispatch.call(runStepF,"NewList","").toDispatch();
			logger.info("Successfully retrived the steps from the run: "+Dispatch.call(newRun,"Name"));
			
		}catch (ComFailException e) { 
				logger.error("Unable to get test run steps for the run  : Test Run ID : "+Dispatch.call(newRun,"ID")+
					"|| Test Run Name :" +Dispatch.call(newRun,"Name")+ " . Please verify the run.::" +e.getMessage());
		}catch (ComException e) {
				logger.error("Unable to get test run steps for the run  : Test Run ID : "+Dispatch.call(newRun,"ID")+
					"|| Test Run Name :" +Dispatch.call(newRun,"Name")+ " . Please verify the run.::" +e.getMessage());
		}	
		
		return stepsList;
	}
				
	/** 
	 * This method is used to update the specific  test step  result with status  and 
	 * actual result.This function should be iterated through all the list of test steps 
	 * returned form getTestSteps function. 
	 * @param theStep is the instance of test step
	 * @param status is the status of test run
	 * @param actualResult is the actual result of test run
	 * @return The status of updating the test step as True / False
	 * @author Srikanth Kasam
	 **/
	public boolean setTestStepStatus(Dispatch theStep, String status, String actualResult) {
		
		boolean stepUpdateStatus = false;
		// To set the status of the test step
		logger.debug("Updating the step result with  Status:" +status + " , Actual Result :" +actualResult);
		try {
			logger.info("Updating the step status to :"+status);			
			Dispatch.put(theStep,"status" , status);
			logger.info("Updating the actual result of the step to ::" +actualResult);
			Variant[] parm1 = { new Variant("ST_ACTUAL")};
		    Variant value1 = new Variant(actualResult);
		    QCPlugIn.setIndexedProperty(theStep, "Field", parm1, value1);
		    logger.info("Posting the step results to QC ");
		    Dispatch.call(theStep, "Post");
		    logger.info("Result updated succefully in QC");
		    stepUpdateStatus = true;
		    logger.debug("Status of the test step is updated successfully in Quality Center.");
		    
		}catch (ComFailException e) { 
				logger.error("Unable to update step result. Please verify the test step.::" +e.getMessage());
		}catch (ComException e) {
				logger.error("Unable to update step result. Please verify the test step.::" +e.getMessage());
		}
		
		return stepUpdateStatus;
	}
					
	/** 
	 * This method is used to attach the specified file to source object (may be a test /test run / test step) 
	 * @param attachmentPath is the path of the attachment
	 * @param objTest is the instance of Test	 		 
	 * @return The status of attachment as True / False 
	 * @author Srikanth Kasam
	 **/
	public boolean attachFile(String attachmentPath, Dispatch objTest){
		
		boolean attachFileStatus = false; 
		//Attaching file to test instance
			
		String attachmentFileName =  attachmentPath.substring(attachmentPath.lastIndexOf("\\")+ 1,attachmentPath.length());
	    String attachmentFilePath =  attachmentPath.substring(0,attachmentPath.lastIndexOf("\\"));
	    
	    System.out.println("attachment File Name:  " + attachmentFileName);
	    System.out.println("attachment File Path:  " + attachmentFilePath);
	    
	    logger.debug("Attaching the file : "+ attachmentFileName+ " from path :" +attachmentFilePath);
		//Use Bug.Attachments to get the attachment factory for the defect
	   
	   		try { 
	   			//Use Attachments to get the attachment factory
	   			logger.info("Getting attachment factory ");	   			
			    Dispatch attachFact = Dispatch.call(objTest,"attachments").toDispatch();
			    //Add a new extended storage object,an attachment
			    logger.info("Creating attachment object for file :"+ attachmentFileName);
			    Dispatch attachObj  = Dispatch.call(attachFact,"AddItem",attachmentFileName).toDispatch();	        
			    //Modify the attachment description			    
			    Dispatch.put(attachObj,"Description", "File Attachment");	        
			    //Update the attachment record in the project database			    
			    Dispatch.call(attachObj,"Post");
			    //Get the attachment extended storage object
			    logger.info("Creating attachment extended storage object");
			    Dispatch ExStrg = Dispatch.call(attachObj, "AttachmentStorage").toDispatch();	        
			    //Specify the location of the file to upload.
			    logger.info("Locating the file path :" + attachmentFileName);
			    Dispatch.put(ExStrg,"ClientPath",attachmentFilePath);	    
			    //Use IExtendedStorage.Save to upload the file
			    logger.info("Uploading the attachment file to QC");
			    Dispatch.call(ExStrg,"Save", attachmentFileName, "True");	        
			    attachFileStatus = true;
			    logger.debug("File : "+attachmentFileName +" is attached  from file path : "+ attachmentFilePath+" with " +Dispatch.call(objTest,"Name").toString() );
	        
	   		}catch (ComFailException e) { 
				logger.error("Unable to upload the file :"+attachmentFileName+"  from file path : " +attachmentFilePath+ 
						" Please verify the file path provided and file existence ::" +e.getMessage());
	   		}catch (ComException e) {
	   			logger.error("Unable to upload the file :"+attachmentFileName+"  from file path : " +attachmentFilePath+ 
						" Please verify the file path provided and file existence ::" +e.getMessage());
	   		}	   		
	        return attachFileStatus;	
	}	
	/** 
	 * This method is used to create the bug factory, this bug factory can be used for creating new bugs.  
	 * @param objQCconn is the QC connection object  		 
	 * @return the bug factory instance
	 * @author Srikanth Kasam
	 **/
	public Dispatch createBugFactory(ActiveXComponent objQCconn) {
		
		Dispatch bugF = null;
		//Creating the bug factory for logging the new defects
		logger.debug("Creating bug factory");
		try{
			bugF = Dispatch.call(objQCconn ,"bugFactory").toDispatch();
			logger.info("Bug factroy is created");
			
		}catch (ComFailException e) { 
			logger.error("Unable to create the bug factory. Please check the QC connetion instance and try again ::" +e.getMessage());
   		}catch (ComException e) {
   			logger.error("Unable to create the bug factory. Please check the QC connetion instance and try again ::" +e.getMessage());
   		}
		
      return bugF;			
		}	
	/** 
	 * This method is used to create a new defect with provided details without any attachment.  
	 * @param bugF is the instance of bug factory
	 * @param detectedBy is the logged in QC user name  
	 * @param description is the defect description
	 * @param summary is the defect title
	 * @param severity is the defect severity
	 * @param priority is the defect priority		 
	 * @return The Defect instance 
	 * @author Srikanth Kasam
	 **/
	public Dispatch logDefect(Dispatch bugF, String detectedBy, String summary, String description, String severity, String priority){
		 
		Dispatch theBug = null;
		
		// Creating the new defect
		logger.debug("Creating a new defect with details :: Summary : " + summary + " Description : " 
					+ description+" Severity : "+ severity +" Priority : "+ priority);
		
		//Get the current the system date value to assign it to Detection Date property
		Calendar currentDate = Calendar.getInstance(); 
		SimpleDateFormat formatter =  new SimpleDateFormat("MM/dd/yyyy");
		String dateNow = formatter.format(currentDate.getTime());
		
		   try {			   
			   	logger.info("Creating the new defect instance");
			    theBug = Dispatch.call(bugF , "AddItem","New Defect").toDispatch();
			    Dispatch.put(theBug,"Summary",summary); 
			    Dispatch.put(theBug,"status", "New");
			    Dispatch.put(theBug,"Priority", priority);
			    Dispatch.put(theBug,"DetectedBy", detectedBy);
			   
			   // To Set the value of fields of object	    
			    Variant[] parm1 = { new Variant("BG_SEVERITY")};
			    Variant value1 = new Variant(severity);
			    QCPlugIn.setIndexedProperty(theBug, "Field", parm1, value1);
			    
			    Variant[] parm2 = { new Variant("BG_DETECTION_DATE")};
			    Variant value2 = new Variant(dateNow);
			    QCPlugIn.setIndexedProperty(theBug, "Field", parm2, value2);
			    
			    Variant[] parm3 = { new Variant("BG_DESCRIPTION")};
			    Variant value3 = new Variant(description);
			    QCPlugIn.setIndexedProperty(theBug, "Field", parm3, value3);
			    
			    Dispatch.call(theBug,"Post");
			    logger.debug("New defect is created and defect ID is ::" +Dispatch.call(theBug,"ID"));
			    
		   }catch (ComFailException e) { 
				logger.error("Unable to create the new defect. Please re-check the values provided. ::" +e.getMessage());
	   		}catch (ComException e) {
	   			logger.error("Unable to create the new defect. Please re-check the values provided. ::" +e.getMessage());
	   		}
		  
		   return theBug;
		}
	
	/** 
	 * This method is used to create a new defect with provided details and with attachment. 
	 * @param bugF is the bug factory instance
	 * @param detectedBy is the logged in QC user name 
	 * @param description is the defect description
	 * @param summary is the defect title
	 * @param severity is the defect severity
	 * @param priority is the defect priority	
	 * @param attachmentPath is the attachment path		 
	 * @return The Defect instance 
	 * @author Srikanth Kasam
	 **/
	public Dispatch logDefectWithAttachment(Dispatch bugF, String detectedBy, String summary, String description, String severity, 
									String priority, String attachmentPath){
		
		 	Dispatch theBug = null;
		 	String attachmentFileName = null;
		 	String attachmentFilePath = null;
		 	// Creating the new defect
			logger.debug("Creating a new defect with details :: Summary : " + summary + " Description : " 
					+ description+" Severity : "+ severity +" Priority : "+ priority+ " and attachment form" +
							" path : " + attachmentPath);
			
		 	//Get the current the system date value to assign it to Detection Date property
			Calendar currentDate = Calendar.getInstance(); 
			SimpleDateFormat formatter =  new SimpleDateFormat("MM/dd/yyyy");
			String dateNow = formatter.format(currentDate.getTime());
			
		   try {
			   	logger.info("Creating the new defect instance");			   	
			   	theBug = Dispatch.call(bugF , "AddItem","New Defect").toDispatch();
			    Dispatch.put(theBug,"Summary",summary); 
			    Dispatch.put(theBug,"status", "New");
			    Dispatch.put(theBug,"Priority", priority);
			    Dispatch.put(theBug,"DetectedBy", detectedBy);
			    
			   // To Set the value of fields of object	    
			    Variant[] parm1 = { new Variant("BG_SEVERITY")};
			    Variant value1 = new Variant(severity);
			    QCPlugIn.setIndexedProperty(theBug, "Field", parm1, value1);
			    
			    Variant[] parm2 = { new Variant("BG_DETECTION_DATE")};
			    Variant value2 = new Variant(dateNow);
			    QCPlugIn.setIndexedProperty(theBug, "Field", parm2, value2);
			    
			    Variant[] parm3 = { new Variant("BG_DESCRIPTION")};
			    Variant value3 = new Variant(description);
			    QCPlugIn.setIndexedProperty(theBug, "Field", parm3, value3);
			    
			    Dispatch.call(theBug,"Post"); 
			    logger.info("Created new defect successfully in QC ");
			    
			    //Use Bug.Attachments to get the attachment factory for the defect
			    logger.info("Attaching the file from : "+ attachmentPath + " to defect ID : " +Dispatch.call(theBug,"ID"));
			    			    
			     attachmentFileName =  attachmentPath.substring(attachmentPath.lastIndexOf("\\")+ 1 , attachmentPath.length());
			     attachmentFilePath =  attachmentPath.substring(0 , attachmentPath.lastIndexOf("\\"));
			   
			    	// Use Bug.Attachments to get the bug attachment factory
			    	logger.info("Getting bug attachment factory ");	 
			    	Dispatch attachFact = Dispatch.call(theBug,"attachments").toDispatch();		    
			    	//Add a new extended storage object,an attachment	
			    	logger.info("Creating attachment object for file :"+ attachmentFileName);
			    	Dispatch attachObj  = Dispatch.call(attachFact,"AddItem" , attachmentFileName).toDispatch();		        
			    	//Modify the attachment description			    	
			        Dispatch.put(attachObj,"Description", "Bug Sample Attachment");		        
			        //Update the attachment record in the project database			        
			        Dispatch.call(attachObj,"Post");		        
			        //Get the bug attachment extended storage object
			        logger.info("Creating attachment extended storage object");
			        Dispatch ExStrg = Dispatch.call(attachObj, "AttachmentStorage").toDispatch();		        
			        //Specify the location of the file to upload.
			        logger.info("Locating the file path :" + attachmentFileName);
			        Dispatch.put(ExStrg,"ClientPath", attachmentFilePath);		    
			        //Use IExtendedStorage.Save to upload the file
			        logger.info("Uploading the file to QC ");
			        Dispatch.call(ExStrg,"Save", attachmentFileName , "True");
			        logger.debug("New defect is created  with attachment the defect ID is ::" +Dispatch.call(theBug,"ID"));
			        
		   } catch (ComFailException e) { 
				logger.error("Unable to  create new defect with attachment. Please verify the file "+attachmentFileName
							+" +exist at file path  : " +attachmentFilePath+" ::" +e.getMessage());
	   		}catch (ComException e) {
	   			logger.error("Unable to  create new defect with attachment. Please verify the file "+attachmentFileName
						+" +exist at file path  : " +attachmentFilePath+" ::" +e.getMessage());
	   		}
		   
		return theBug;		   
	}
	
	/** 
	 * This method is used to link the specified bug to corresponding test or test step.  
	 * @param theBug is the bug instance
	 * @param linkObject is the instance to link (either Test/ Test Step) 		 
	 * @return The status of attachment as True / False
	 * @author Srikanth Kasam
	 **/
	
	public boolean linkDefect(Dispatch theBug, Dispatch linkObject) {
		
		boolean linkStatus = false;
		// Linking the defect with either Test/ Test Step
		logger.debug("Linking the defect ID : "+Dispatch.call(theBug,"ID"));
		
		Dispatch ilink = linkObject;
		try {
			logger.info(" Creating the bug link factory ");
			Dispatch linkF = Dispatch.call(ilink,"BugLinkFactory").toDispatch();
			logger.info("Linking the defect ");
			Dispatch link = Dispatch.call(linkF ,"AddItem" ,theBug).toDispatch();
			Dispatch.call(link, "Post");
			logger.info(" Defect is linked with test successfully");
			linkStatus = true;
			
		} catch (ComFailException e) { 
			logger.error("Unable to  link the defect with test /test step. Please verify " +
					"the defect ID: "+Dispatch.call(theBug,"ID")+" existance. ::" +e.getMessage());
		}catch (ComException e) {
			logger.error("Unable to  link the defect with test /test step. Please verify " +
					"the defect ID: "+Dispatch.call(theBug,"ID")+" existance. ::" +e.getMessage());
		}
		   return linkStatus;
	}				
	/** 
	 * This method is used to create a new Test Set (FOR IMPLEMENTATION PHASE II)
	 * @param objQCconn is the object of QC connection
	 * @param folderPath is the folder path
	 * @param testSetName is the test set name 		 
	 * @return The Test Set instance
	 * @author Srikanth Kasam
	 **/
	//public Dispatch createTestSet(ActiveXComponent objQCconn, String folderPath, String testSetName) {	
		
	//}
	/** 
	 * This method is associate the Tests under a given Test to the specified TestSet
	 * @param objTest is the Test instance
	 * @param objTestSet is the test set instance 		 
	 * @returns the TestSet instance
	 * @author Srikanth Kasam	 
	 **/
	//public TestSet associateTestToTestSet(Test objTest, TestSet objTestSet){
	
	//} 
	
	public static void setIndexedProperty(Dispatch activex, String name, Variant[] indexes, Variant value) 
	{ 
		Variant[] variants = new Variant[indexes.length + 1]; 		 
		for(int i=0; i<indexes.length; i++) { 
		 variants[i] = indexes[i]; 
		} 
		variants[variants.length-1] = value; 
		Dispatch.invoke(activex, name, Dispatch.Put, variants, new int[variants.length]); 
	}
	
}


