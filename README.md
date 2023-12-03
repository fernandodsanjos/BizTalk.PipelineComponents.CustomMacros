
# BizTalk.PipelineComponents.CustomMacros
Adds custom SendPort macros, works for both Filename and folder<br/>

## Properties
|Name|Description|
|--|--|
|Must Exist|Target folder path must exist.|
|Context Required|An error is thrown if the specified Context cannot be found|
|SSOAffiliate|Use SSO Application instead of host user|

## Macros
If the PipelineComponent is used in any stage in a sendport pipeline, the following custom macros will be available

**%IF(#,#,#)%**<br/> If the expression (first parameter) is matched in any way the second parameter will be returned or else the last parameter (else) will be returned<br/>
Example: outbound transport location is **c:\Volvo\Cars\V40\cars.vsc** and the macro is **Cars_%IF(XC40,SUV,Regular)%.xml** then **Cars_Regular.xml** will be returned.<br/>

**%TimeStamp(#,#,..)%** <br/> Used to modify timestamp much like the _Nearest Macro_ except with TimeStamp macro you can modify **multiple** placeholders and the **complete modified** timestamp will be returned.<br/>
Example: outbound transport location is **\\\\share\\DC\\certificates_%TimeStamp([0-9]{8}([0-9]{2})[0-9]{2}([0-9]{2}),5,4)%.xml** then macro will return **certificates_20230808003032.xml** where hour and seconds are modified before returning the complete modified timestamp.<br/>

**%Nearest(#,#)%** <br/> Used to modify timestamp in the filename as sometimes you need to make sure two files has the same timestamp even though they are created in different processes.<br/>Consider two files named certificates_2023080813**3030**.csv and competences_2023080813**3133**.csv and where specification says the timestamp must be the same for both files.
<br/>Using the following file pattern %FilePattern([a-z]+_[0-9]{8})%**%Nearest(([0-9]{2})[0-9]{2}[.],30)%00.csv** the files will be renamed to certificates_202308081**33000**.csv and competences_2023080813**3000**.csv
<br/>First part uses **FilePattern** macro to extract the beginning of thte filename. **Nearest** macro picks out the minut part of the datetime and calculates it to the nearest 30 min window (second parameter). 
In this sample seconds have been hardcoded to 00.<br/><br/>


**%ParentFolder%**<br/> Returnes the parent folder to original source folder.<br/>

**%FolderRange(#,#)%**<br/> Return a folder or folderpath part from incomming file. <br/>
Example if original folderpath is **\\\\server\folder1\folder2\folder3\filename.txt**<br/>
Then **%FolderRange(2,2)%** returns  folder1\folder2<br/>If **%FolderRange(2)%** is used then folder2 is returned.<br/>

![Example](/MacroFolder.JPG?raw=true "Example")

**%OriginalPath%**<br/> Folder structure from incomming file, volume or drive letter excluded. 
If inboundlocation is **D:\Integrations\Incomming\file.txt** then the macro returns **Integrations\Incomming**.<br/>

**%OriginalFolder%**<br/> Returns original folder name. 
If inboundlocation is **D:\Integrations\Incomming\file.txt** then the macro returns **Incomming**.<br/>

**%FileNameOnly%**<br/>    Like %SourceFileName% but without the file extension.<br/>

**%AddDays(#,#)%**<br/> Lets you add or subtract day(s) for DateTime.Now. Works like DateTime.Now.AddDays(#).ToString(#)<br/>
                  Examples: _%AddDays(1,dd)%_, add one day to current datetime<br/>
                           _%AddDays(-2,dd)%_, substract two days to current datetime<br/>

**%AddMonths(#,#)%**<br/> Lets you add or subtract month(s) for DateTime.Now. Works like DateTime.Now.AddMonths(#).ToString(#)<br/>
                  Examples: _%AddMonths(1,mm)%_, add one month to current datetime<br/>
                           _%AddMonths(-2,mm)%_, substract two months to current datetime<br/>
			   
**%AddYears(#,#)%**<br/> Lets you add or subtract day(s) for DateTime.Now. Works like DateTime.Now.AddYears(#).ToString(#)<br/>
                  Examples: _%AddYears(1,yyyy)%_, add one year to current datetime<br/>
                           _%AddYears(-2,yyyy)%_, substract two years to current datetime<br/>
			   
**%UTCAddDays(#,#)%**<br/> Lets you add or subtract day(s) for DateTime.UtcNow. Works like DateTime.UtcNow.AddDays(#).ToString(#)<br/>
                  Examples: _%UTCAddDays(1,dd)%_, add one day to current utc datetime<br/>
                           _%UTCAddDays(-2,dd)%_, substract one day to current utc datetime<br/>

**%UTCAddMonths(#,#)%**<br/> Lets you add or subtract month(s) for DateTime.UtcNow. Works like DateTime.UtcNow.AddMonths(#).ToString(#)<br/>
                  Examples: _%UTCAddMonths(1,mm)%_, add one month to current utc datetime<br/>
                           _%UTCAddMonths(-2,mm)%_, substract two months to current utc datetime<br/>
			   
**%UTCAddYears(#,#)%**<br/> Lets you add or subtract day(s) for DateTime.UtcNow. Works like DateTime.UtcNow.AddYears(#).ToString(#)<br/>
                  Examples: _%UTCAddYears(1,yyyy)%_, add one year to current datetime<br/>
                           _%UTCAddYears(-2,yyyy)%_, substract two years to current utc datetime<br/>

**%UTCDateTime(#)%**<br/>    Lets you format the current date time as you want it. Works more or less like DateTime.UtcNow.ToString()<br/>
                  Examples: _%UTCDateTime(d)%_, current day number<br/>
                            _%UTCDateTime(yyyy)%_, current year<br/>
                            _%UTCDateTime(MM)%_, current month<br/>
			    
**%DateTime(#)%**<br/>    Lets you format the current date time as you want it. Works more or less like DateTime.Now.ToString(#)<br/>
                  Examples: _%DateTime(d)%_, current day number<br/>
                            _%DateTime(yyyy)%_, current year<br/>
                            _%DateTime(MM)%_, current month<br/>
			    
**%FileDateTime(#)%**<br/> Works like the %DateTime% macro but it uses the FileCreationTime instead.<br/>

**%Context(#)%**<br/>     Use any value from standard context values<br/>
                  Example _%Context(BTS.InterchangeID)%_ returns the InterchageID of the message.<br/><br/>
**Available standard Context**
  - SBMessaging: http://schemas.microsoft.com/BizTalk/2012/Adapter/BrokeredMessage-properties
  - POP3: http://schemas.microsoft.com/BizTalk/2003/pop3-properties
  - MSMQT: http://schemas.microsoft.com/BizTalk/2003/msmqt-properties
  - ErrorReport: http://schemas.microsoft.com/BizTalk/2005/error-report
  - EdiOverride: http://schemas.microsoft.com/BizTalk/2006/edi-properties
  - EDI: http://schemas.microsoft.com/Edi/PropertySchema
  - EdiIntAS: http://schemas.microsoft.com/BizTalk/2006/as2-properties
  - BTF2: http://schemas.microsoft.com/BizTalk/2003/btf2-properties
  - BTS: http://schemas.microsoft.com/BizTalk/2003/system-properties
  - FILE: http://schemas.microsoft.com/BizTalk/2003/file-properties
  - MessageTracking: http://schemas.microsoft.com/BizTalk/2003/file-properties
  - AzureStorage: http://schemas.microsoft.com/BizTalk/Adapter/AzureStorage-properties

**Custom Context**<br/>Custom context can also be applied by using the prefix CST. A search will be done through all none MS context properties until a match is made.<br/>
                **Distinguished Fields**<br/>
                  It is possible to use distinguished fields by adding a filename friendly xpath variant. This is a better alternative then using unnecessary promoted properties.<br/>
                  Example: 
                  If the distinguished xpath looks like bellow
/*[local-name()='Ledger' and namespace-uri()='http://ICC.Company.Schemas']/*[local-name()='AccountingDate' and namespace-uri()='']
Filename friendly xpath would be _\~Ledger~AccountingDate_
Final macro would then be _%Context(\~Ledger~AccountingDate)%_

				  
**%Folder%**<br/>          Adds a backslash \. This makes it possible to choose any folder bellow the root folder specified by URI.<br/>
                  Example: _%DateTime(yyyy)%%Folder%%DateTime(MM)%%Folder%%DateTime(dd)%%Folder%message.xml_
                           If the URI is c:\integrations\int0010\ and the current date is 2017-05-04 the final folder
                           and filename will be c:\integrations\int0010\2017\05\04\message.xml<br/>
Folder structure will be created if it does not exist.
			   
**%ReceivedDateTime(##)%**<br/> Works like the %DateTime% macro but it uses the AdapterReceiveCompleteTime instead. 
				To use this macro tracking must be enabled.<br/> 
**%Root%**<br/> Returns the root node of a message<br/> 
**%DateTimeFormat([context],[dateformat])%**<br/>
Example: <br/>%DateTimeFormat(FILE.FileCreated,yyyy)%<br/>%DateTimeFormat(\~Invoice\~InvoiceDate,yyyy)% <br/>
 Works like the %DateTime% macro but it works on any context property , promoted or not or  distinguished, that contains any value that could be interpreted as a date/time format<br/>
 Date format yyyyMMdd is also allowed<br/>

**%FilePattern([regex])%** Runs a regex expression on sourcefilename.<br/>
Example: If incomming filename is  P123789.txt and sendport filename is Q%FilePattern([0-9]+)%.txt, Then outgoing filename would result in Q123789.txt<br/>
The component also handle groups in RegEx expression. For example if the expression is "%FilePattern(\_([0-9]{3})\_)%.pdf" and the filename is "WAP_123_12345.xml" then the result would be "123.pdf".<br/>
_Only the first RegEx group will be returned._

**%FileExtension%** Returns the fileextension from sourcefilename.<br/>
Example: If incomming filename is  P123789.txt then .txt will be returned.

