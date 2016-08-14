This utility imports pages and references from one or two Excel.xlsx files into Ardoq.

See [data.xslx](./src/main/resources/data.xlsx?raw=true) for an example Excel format we can use for importing.

##Execution
You must configure your authentication token and set the environment variable ```ardoqToken``` in your environment, or specify it in the properties file.

After you have downloaded this project or [the Überjar](./build/ardoq-excel-import-1.0.jar?raw=true) you can execute it with the following command:

```java -Dfile.encoding=UTF-8 -classpath "build/ardoq-excel-import-1.0.jar" com.ardoq.ExcelImport ./default.properties```

##Configuration

The default.properties file shows how to configure it:
```=ini
#The authentication token to use (Generate a new one via your profile - Click on your name in Ardoq, Profile and Prefs -> API and Tokens)
#ardoqToken=

#The Ardoq host (default is https://app.ardoq.com
#ardoqHost=https://app.ardoq.com

#Name of the workspace to synchronize to
workspaceName=Excel import
#The model name you wish to use
modelName=Application service

#The organization name if you have an Enterprise account
#organization=mw

#Deletes components and references not found in spreadsheet if set to YES
deleteMissing=YES

#ComponentFile is the excel file with components
componentFile=./src/main/resources/data.xlsx

#component sheet is the name of the spreadsheet with compopnents to load
componentSheet=Application list

#The name of the columns that you wish to map to a page type in the model.
compMapping_System=Application
compMapping_Application=Service

#The name of the column that has the descriptions
compDescriptionColumn=General description

#The name of the columns you wish to map to a field in the model. You can have as many column/field mappings as you want.
#Please escape whitespace, colon and equal characters with \ (backslash)

fieldColMapping_TAM\ Mapping=tam_mapping
fieldColMapping_Criticality=criticality
fieldColMapping_Category=category


#Reference File is the excel file with references (can be the same as component file)
#If you do not wish to import references, please comment it out with #
referenceFile=./src/main/resources/data.xlsx

#Component page separator for source and target page references.
referenceComponentSeparator=::

#Referencesheet is the name of the spreadsheet that has the references
referenceSheet=References
#Which row the references start from (set it to 1 if you have header row)
referenceStartFromRow=1
#Which column has the source reference in the format ParentPage::ChildSourcePage
referenceSourceColumn=0
#The column that contains the link type, set it to -1 if no column is available
referenceLinkTypeColumn=1
#The default link type from the model that you wish to use, if no value is present in LinkTypeColumn, or you do not have any
referenceDefaultLinkType=Synchronous
#The column that contains the references, you can have many target pages in one column, with comma separated values.
referenceStartFromColumn=2
```
