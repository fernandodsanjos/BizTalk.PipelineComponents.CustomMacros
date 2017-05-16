# BizTalk.PipelineComponents.CustomMacros
Adds custom SendPort macros

If the PipelineComponent is used in any stage in a sendport pipeline, the following custom macros will be available
One of the goals was to not use pipeline properties, just plug-and-play.

%FileNameOnly%    like %SourceFileName% but without the file extension.
%DateTime(##)%    Lets you format the current date time as you want it. Works more or less like DateTime.Now.ToString(##)
                  Examples: %DateTime(%d)%, current day number
                            %DateTime(yyyy)%, current year
                            %DateTime(MM)%, current month
%FileDateTime(##)% Works like the %DateTime% macro but it uses the FileCreationTime instead.
%Context(##)%     Use any value from standard context values
                  Example %Context(BTS.InterchangeID)% returns the InterchageID of the message.
				  Custom context can also be applied by using the prefix CST. A search will be done through all none MS context until
				  a match is made. The only thing one must make sure is that there must not exist two (2) custom context properties with the same name.
%Folder%          Adds a backslash \. This makes it possible to choose any folder bellow the root folder specified by URI.
                  Example: %DateTime(yyyy)%%Folder%%DateTime(MM)%%Folder%%DateTime(dd)%%Folder%message.xml
                           If the URI is c:\integrations\int0010\ and the current date is 2017-05-04 the final folder
                           and filename will be c:\integrations\int0010\2017\05\04\message.xml
