# BizTalk.PipelineComponents.CustomMacros
Adds custom SendPort macros<br/>

If the PipelineComponent is used in any stage in a sendport pipeline, the following custom macros will be available
One of the goals was to not use pipeline properties, just plug-and-play.<br/>

**%FileNameOnly%**<br/>    Like %SourceFileName% but without the file extension.<br/>

**%DateTime(#)%**<br/>    Lets you format the current date time as you want it. Works more or less like DateTime.Now.ToString(#)<br/>
                  Examples: _%DateTime(%d)%_, current day number<br/>
                            _%DateTime(yyyy)%_, current year<br/>
                            _%DateTime(MM)%_, current month<br/>
			    
**%FileDateTime(#)%**<br/> Works like the %DateTime% macro but it uses the FileCreationTime instead.<br/>

**%Context(#)%**<br/>     Use any value from standard context values<br/>
                  Example _%Context(BTS.InterchangeID)%_ returns the InterchageID of the message.<br/>
				  Custom context can also be applied by using the prefix CST. A search will be done through all none MS context properties until a match is made.<br/>
                *Distinguished fields*
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

