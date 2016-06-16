Data Import Tool
================

Generating Excel Templates and Importing Bulk Data - Using simple servlet and Apache POI.

Build Status
============

Travis

[![Build
Status](https://travis-ci.org/avikganguly01/DataImportTool.png?branch=master)](https://travis-ci.org/avikganguly01/DataImportTool)

Setup Instructions
==================

1. Before you run the application, you need to have gradle installed and create a file dataimport.properties directly under your home directory. It should have the following 4 parameters:-

  mifos.endpoint=https://demo.openmf.org/fineract-provider/api/v1/  
  mifos.user.id=mifos  
  mifos.password=password  
  mifos.tenant.id=default  
  
  sample file is at https://github.com/openMF/DataImportTool/blob/develop/dataimport.properties

2. Use the command "gradle clean tomcatRunWar" to run the application and access it at localhost:8070/DataImportTool.

3. If you are hosting the data import tool in the cloud, you need to ssh into the system to create the dataimport.properties file.

Note :- Default gradlew config will allow you to remote debug on port 8006.

Troubleshooting
===============

1. If you are hosting both this tool and the backend in the same system, you can change the debug port in gradlew.bat under mifosng-provider to listen in on a different port instead of 8005:-  set DEFAULT_JVM_OPTS=-Xdebug -Xrunjdwp:transport=dt_socket,address=8006,server=y,suspend=n

2. If you accidentally run out of heap size when running both in the same system, make sure your \_JAVA_OPTIONS in Environment variables is set to -Xms512m -Xmx512m -XX:MaxPermSize=512m and it is getting picked up by gradle.



To Dos
======

1. Transaction support for group loans and group savings -> (Blocker) Can't find an endpoint which returns all loan accounts or all loan accounts associated with groups. /loans returns only individual loans.
2. Better workbook populator unit tests which will use FormulaEvaluator to evaluate if the data validation formulas and in-cell formulas embedded as Strings are not broken due to shifting of columns.
3. Minor improvements to group related features once the release is stable (like sync repayments with meetings).

Dev Setup
=========
1. Eclipsify using command gradle clean cleanEclipse eclipse
2. Import into project into workspace
When opened, in the package explorer of the java perspective, right-click import->General->Existing projects into workspace. In the dialog that opens, specify the root directory option by browsing to and selecting the DataImportTool directory.
