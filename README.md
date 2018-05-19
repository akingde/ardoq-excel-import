This utility imports data from Excel into Ardoq.

See [./examples/business_process/business_process.xlsm](./examples/business_process/business_process.xlsm?raw=true) for an example Excel format we can use for importing.
The file also contains a VBA script for automatically generating references that can be used as a starting point.

## Execution
You must configure your authentication token and set the environment variable `ardoqToken` and `organization` in your environment, or specify it in the properties file.

After you have downloaded this project you can execute the example with the following command (remember to configure your api-token and organization first):

```java -Dfile.encoding=UTF-8 -classpath "build/ardoq-excel-import-1.2.1.jar" com.ardoq.ExcelImport ./examples/business_process/business_process.properties```

If you only need the binary you can download [the Überjar](./build/ardoq-excel-import-1.2.1.jar?raw=true) 

## Example

The included example file [./examples/business_process/business_process.xlsm](./examples/business_process/business_process.xlsm?raw=true) show how you can import a small process into the default Business Process template.

![Model](examples/business_process/img/Model.png)

After importing the example file, Ardoq will automatically generated visualizations, for example this Swimlane view. Note how we can color the steps differently depending on the step Complexity.

![Swimlane](examples/business_process/img/Swimlane.png)

The table view can be used to see the imported field values.
![Table view](examples/business_process/img/Table.png)

The Treemap view can dynamically size steps according to field values (in this example the duration of each step).

![Treemap view](examples/business_process/img/Treemap.png)

Showing existing links and quickly adding new ones can be done using the Dependency Matrix view

![Dependency Matrix view](examples/business_process/img/DependencyMatrix.png)


## Cross workspace reference configuration
If you need to create references across workspaces, that can be done running multiple imports.
First you need to import the workspace and components that will be references from another workspace.
See an example of configurations and example Excel spread sheet in ./examples/cross_workspace_references

## Configuration

The default.properties file shows how to configure it:
```=ini
#The authentication token to use (Generate a new one via your profile - Click on your name in Ardoq, Profile and Prefs -> API and Tokens)
ardoqToken=<your api token>

#The Ardoq host (default is https://app.ardoq.com
ardoqHost=https://app.ardoq.com

#Name of the workspace to synchronize to
workspaceName=Business Process via Excel

#The model name you wish to use. (You can find the available models here https://app.ardoq.com/api/model?includeCommon=true?org=<your organization label>)
modelName=Business Process

#The organization label (You'll find this in the API-token tab under your settings)
organization=<your organization label>

#Deletes components and references not found in spreadsheet if set to YES
deleteMissing=NO

#ComponentFile is the excel file with components
componentFile=./examples/business_process/business_process.xlsm

#component sheet is the name of the spreadsheet with components to load
componentSheet=Process steps

#The name of the columns that you wish to map to a page type in the model (whitespaces must be escaped with \).
compMapping_Business\ process=Business process
#If you have multiple leaf types in your model, you can use a separate column to specify the type
dynamicCompMapping_Type=Steps

#The name of the column that has the descriptions
compDescriptionColumn=General description

#The name of the columns you wish to map to a field in the model. You can have as many column/field mappings as you want.
#Please escape whitespace, colon and equal characters with \ (backslash)

fieldColMapping_Complexity=complexity
fieldColMapping_Avg.\ Time\ spent\ (minutes)=time_per_step

#Reference File is the excel file with references (can be the same as component file)
#If you do not wish to import references, please comment it out with #
referenceFile=./examples/business_process/business_process.xlsm

#Component page separator for source and target page references.
referenceComponentSeparator=::

#If you wish to link to a different _PRE-EXISTING_ workspace specify name here
#NB! Component's must exist already, and will not be auto created.
#If you need that, then you must a seperate configuration that maps the models and fields, and run it seperately.
#targetWorkspaceName=excelImportTarget 

#Referencesheet is the name of the spreadsheet that has the references
referenceSheet=Links (auto generated)
#Which row the references start from (set it to 1 if you have header row)
referenceStartFromRow=1
#Which column has the source reference in the format ParentPage::ChildSourcePage
referenceSourceColumn=0
#The column that contains the link type, set it to -1 if no column is available
referenceLinkTypeColumn=1
#The default link type from the model that you wish to use, if no value is present in LinkTypeColumn, or you do not have any
referenceDefaultLinkType=Next step
#The column that contains the references, you can have many target pages in one column, with comma separated values.
referenceStartFromColumn=2
```
