# BizTalk.PipelineComponents.CustomMacros
Adds custom SendPort macros

If the PipelineComponent is used in any stage in a sendport pipeline, the following custom macros will be available

%FileNameOnly%    like %SourceFileName% but without the file extension.
%DateTime(##)%    Lets you format the current date time as you want it. Works more or less like DateTime.Now.ToString(##)
                  Examples: %DateTime(%d)%, current day number
                            %DateTime(yyyy)%, current year
                            %DateTime(MM)%, current month
%Context(##)%     Use any value from standard context values
                  Example %Context(BTS.InterchangeID)% returns the InterchageID of the message.
%Folder%          Adds a backslash \. This enables you choose any folder bellow the root folder specified by URI.
                  Example: %DateTime(yyyy)%%Folder%%DateTime(MM)%%Folder%%DateTime(dd)%%Folder%message.xml
                           If the URI is c:\integrations\int0010\ and the current date is 2017-05-04 the final folder
                           and filename will be c:\integrations\int0010\2017\05\04\message.xml
