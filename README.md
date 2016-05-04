# TranslationManager

TranslationManager is a tool for those working with multilingual applications to aid in the management of translated data.

### Goal
The idea behind this tool is to make managing translations easier. It assumes the use of resource (.resx) files to store the localized text for the application and excel files to pass between translators. When complete, this application will be able to generate the excel sheets to send to translators and be able to read the excel sheets back into the application. If new resource keys are found, they are added. If the keys already exist, the value is updated.

### Features
- Importing an excel file

**Upcoming Additions:**
- Exporting to an excel file
- Import results
- Ability to delete resource file contents before import
- Save settings preferences

### Importing Usage
1. Choose a resource file and an excel file
2. Set the correct settings
  1. Worksheet Index
    - The one-based index of the worksheet within the excel file that contains the data you would like to import
  2. Skip First Row?
    - If the first row of the excel file are column headers, you can skip it by selecting YES
  3. Data Label Column
    - The column that holds the resource keys
  4. Translated Text Column
    - The column that holds the resource text
3. Click import

A message will display whether or not the import was successful.


### Dependencies
[EPPlus](http://epplus.codeplex.com/)


### Example
See the TestingFiles folder in this repository. Change any of the data in the excel file and run it through the import process to see how the resource file gets updated.
