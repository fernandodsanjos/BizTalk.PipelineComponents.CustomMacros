using System;
using System.Collections.Generic;
using System.Resources;
using System.Drawing;
using System.Collections;
using System.Collections.Specialized;
using System.Reflection;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using System.IO;
using Microsoft.BizTalk.Streaming;
using Microsoft.BizTalk.Message.Interop;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.ScalableTransformation;
using Microsoft.XLANGs.RuntimeTypes;
using System.ComponentModel.DataAnnotations;
using System.Globalization;

namespace BizTalkComponents.PipelineComponents
{
    /// <summary>
    ///  Transforms original message stream using streaming scalable transformation via provided map specification.
    /// </summary>
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [ComponentCategory(CategoryTypes.CATID_Any)]
    [System.Runtime.InteropServices.Guid("45A34C0D-8D73-45fd-960D-DB365CD56371")]
    public partial class CustomMacros : IBaseComponent
    {
        Dictionary<string,string> propertys = new Dictionary<string,string>();

        readonly string ns_Tracking = "http://schemas.microsoft.com/BizTalk/2003/messagetracking-properties";
        readonly string ns_BTS = "http://schemas.microsoft.com/BizTalk/2003/system-properties";
        readonly string ns_FILE = "http://schemas.microsoft.com/BizTalk/2003/file-properties";
        //2019-03-14 Updated FilePattern so that it can handle captures (...)
        readonly string macros = @"%FilePattern\([^%]+\)%|%DateTime\([%YyMmDdTHhSs:\-fzZkK\s]*\)%|%ReceivedDateTime\([%YyMmDdTHhSs:\-fzZkK\s]*\)%|%FileDateTime\([%YyMmDdTHhSs:\-fzZkK\s]*\)%|%FileNameOnly%|%Context\([^)]*\)%|%DateTimeFormat\([~a-zA-Z0-9\. ]+[,]{1}[-dfFghHKmMsStyz: ]+\)%";


        public CustomMacros()
        {
            propertys = new Dictionary<string,string>();

            propertys.Add("SBMessaging","http://schemas.microsoft.com/BizTalk/2012/Adapter/BrokeredMessage-properties");
            propertys.Add("POP3","http://schemas.microsoft.com/BizTalk/2003/pop3-properties");
            propertys.Add("MSMQT","http://schemas.microsoft.com/BizTalk/2003/msmqt-properties");
            propertys.Add("ErrorReport","http://schemas.microsoft.com/BizTalk/2005/error-report");
            propertys.Add("EdiOverride","http://schemas.microsoft.com/BizTalk/2006/edi-properties");
            propertys.Add("EDI","http://schemas.microsoft.com/Edi/PropertySchema");
            propertys.Add("EdiIntAS","http://schemas.microsoft.com/BizTalk/2006/as2-properties");
            propertys.Add("BTF2","http://schemas.microsoft.com/BizTalk/2003/btf2-properties");
            propertys.Add("BTS","http://schemas.microsoft.com/BizTalk/2003/system-properties");
            propertys.Add("FILE","http://schemas.microsoft.com/BizTalk/2003/file-properties");
            propertys.Add("MessageTracking", "http://schemas.microsoft.com/BizTalk/2003/file-properties");
            
        }

        #region IComponent Members

        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            CultureInfo ci = CultureInfo.InvariantCulture;
            object val = null;
            
            //%Folder% = Adds a backslash
            //Does the filemask have any custom macros
            string transport = (string)pInMsg.Context.Read("OutboundTransportLocation", ns_BTS);

            MatchCollection collection = Regex.Matches(transport, macros);

            if (collection.Count > 0)
                pInMsg.Context.Promote("IsDynamicSend", ns_BTS, (object)true);
           
            foreach (Match match in collection)
            {
                 if (match.Value.StartsWith("%DateTime("))
                 {
                     transport = CheckDateTime(transport, match);
                 }
                 //AdapterReceiveCompleteTime
                 else if (match.Value.StartsWith("%ReceivedDateTime("))
                 {
                     //Trackig must be enabled
                     val = pInMsg.Context.Read("AdapterReceiveCompleteTime", ns_Tracking);
                     if (val == null)
                     {
                         transport = transport.Replace(match.Value, "");
                         continue;
                     }
                         

                     //DateTime dt = DateTime.ParseExact((string)val, "yyyy-MM-dd hh:mm:ss", ci);

                     transport = CheckContextDateTime(transport, match, (DateTime)val, "ReceivedDateTime");
                 }
                 else if (match.Value.StartsWith("%FileDateTime("))
                 {
                     val = pInMsg.Context.Read("FileCreationTime", ns_FILE);

                     if (val == null)
                     {
                         transport = transport.Replace(match.Value, "");
                         continue;
                     }

                     //DateTime dt = DateTime.ParseExact((string)val, "yyyy-MM-dd hh:mm:ss", ci);
                     transport = CheckContextDateTime(transport, match, (DateTime)val, "FileDateTime");
                     //transport = CheckFileDateTime(transport, match,(DateTime)val);
                 }
                 else if (match.Value == "%FileNameOnly%")
                 {
                     val = pInMsg.Context.Read("ReceivedFileName", ns_FILE);

                     if (val == null || String.IsNullOrWhiteSpace((string)val))
                     {
                         transport = transport.Replace(match.Value, "");
                         continue;
                     }

                     string receivedFileName = Path.GetFileNameWithoutExtension((string)val);
                     
                     transport = transport.Replace("%FileNameOnly%", receivedFileName);
                 }
                 else if (match.Value.StartsWith("%FilePattern("))
                 {
                     val = pInMsg.Context.Read("ReceivedFileName", ns_FILE);

                     if (val == null || String.IsNullOrWhiteSpace((string)val))
                     {
                         transport = transport.Replace(match.Value, "");
                         continue;
                     }

                     string receivedFileName = Path.GetFileNameWithoutExtension((string)val);

                     transport = GetFileParts(transport,receivedFileName, match);
                 }
                 else if (match.Value.StartsWith("%Context("))
                 {
                     transport = GetContext(transport, match, pInMsg);
                 }
                else if (match.Value.StartsWith("%DateTimeFormat("))
                {
                    transport = DateTimeFormat(pInMsg, transport, match);
                }
                    
            }
                   
            

            if (transport.Contains("%Folder%"))
            {
                transport = transport.Replace("%Folder%", @"\");

                string transportType = (string)pInMsg.Context.Read("OutboundTransportType", ns_BTS);

                if(transportType == "FILE")
                {
                    string dirname = Path.GetDirectoryName(transport);
                    //Will create the new directory structure if it does not exist, only applies to FOLDER
                    System.IO.Directory.CreateDirectory(dirname);
                }
                
            }

            //2018-03-12 Added %Root% Message root node
            if (transport.Contains("%Root%"))
            {
                BTS.MessageType msg = new BTS.MessageType();

                string msgType = (string)pInMsg.Context.Read(msg.Name.Name, msg.Name.Namespace);

   

                transport = transport.Replace("%Root%", msgType.Substring(msgType.IndexOf('#') + 1));
            }

            pInMsg.Context.Promote("OutboundTransportLocation", ns_BTS, (object)transport);

          
          //  pInMsg.BodyPart.Data = stm;

            return pInMsg;
        }



        #endregion

        #region Private methods

        private string DateTimeFormat(IBaseMessage msg, string input, Match match)
        {
            string retpart = String.Empty;
            object context = null;

            string[] parts = match.Value.Split(',');

            parts[0] = parts[0].Replace("%DateTimeFormat(", "").Trim();//context part
            string format = parts[1] = parts[1].Replace(")%", "").Trim();//datetime format

            context = GetContextValue(parts[0], msg);

            try
            {
                if (context is DateTime)
                {
                    retpart = ((DateTime)context).ToString(format);
                }
                else if (Regex.Match((string)context, @"^\d{8}$").Success)

                {
                    retpart = DateTime.ParseExact((string)context, "yyyyMMdd", null).ToString(format);
                }
                else
                    retpart = DateTime.Parse((string)context).ToString(format);
            }
            catch (FormatException)
            {

               //do nothing
            }
            finally
            {

                 input = input.Replace(match.Value, retpart);
               
            }

            return input;

        }

        object CustomMacro(IBaseMessage msg,string ctxName)
        {
            object val = null;

                for (int i = 0; i < msg.Context.CountProperties; i++)
                {
                    string name = String.Empty;
                    string ns = String.Empty;

                    object context_val = msg.Context.ReadAt(i, out name, out ns);

                    System.Diagnostics.Trace.WriteLine(String.Format("Context name {0}, ns {1} value {2}", name, ns, context_val));

                    if (ns.StartsWith("http://schemas.microsoft.com/BizTalk/") == false && ctxName == name)
                    {
                        val = context_val;
                        break;
                    }
                       

                }

                return val; 
            
        }

        static string GetFileParts(string input, string fileName, Match match)
        {

            Match innerMatch = Regex.Match(match.Value, @"%FilePattern\((.*)\)%");
            string innerValue = String.Empty;


            //Get RegEx pattern
            foreach (Group item in innerMatch.Groups)
            {
                if (item.ToString() != match.Value)
                {
                    innerValue = item.Value.Trim();
                    break;
                }

            }

            //Get first match
            innerMatch = Regex.Match(fileName, innerValue);

            innerValue = String.Empty;

            //If count > 0 then capturing group has been used
            if (innerMatch.Groups.Count == 1)
            {
                innerValue = innerMatch.Groups[0].Value.Trim();
            }
            else if (innerMatch.Groups.Count > 1)
            {
                //A capture is used, pick the first one
                innerValue = innerMatch.Groups[1].Value.Trim();
            }



            return input.Replace(match.Value, innerValue);

        }

        //Moved retrival of context value in it's own method
        object GetContextValue(string context, IBaseMessage msg)
        {
            object val = null;
            if (context.StartsWith("~"))
            {
                val = SearchDistinguishedFields(context, msg);
            }
            else
            {
                string[] ctx = context.Split('.');

                string ns = String.Empty;
                propertys.TryGetValue(ctx[0], out ns);


                if (String.IsNullOrWhiteSpace(ns) && ctx[0] != "CST")
                    return val;

                if (ctx[0] == "CST")
                {
                    val = CustomMacro(msg, ctx[1]);
                }
                else
                {

                     val = msg.Context.Read(ctx[1], ns);
                    
                    
                }
            }

            return val;

        }
        string GetContext(string input, Match match, IBaseMessage msg)
        {
            //format %Context(BTS.InterchangeID)%
            Match innerMatch = Regex.Match(match.Value, @"%Context\((.*)\)%");
            string innerValue = String.Empty;
            object val = null;
           

            foreach (Group item in innerMatch.Groups)
            {
                if (item.Value != match.Value)
                {
                    innerValue = item.Value.Trim();
                    break;
                }

            }

            //Added possibility to use distinguished fields
            val = GetContextValue(innerValue, msg);

            //2019-02-13 FIX for invalid conversion
            //2019-02-25 BUG FIX for null value
            if (val != null && !(val is string))
                val = val.ToString();

            if (val == null || String.IsNullOrWhiteSpace((string)val))
            { 
                input =  input.Replace(match.Value, ""); 
            }
            else
            {
                input = input.Replace(match.Value, (string)val); 

            }


            return input;
        }

        string CheckDateTime(string input,Match match)
        {
            
            //include a percent ("%") format specifier before the single custom date and time specifier. like d,m e.t.c
            //bellow creates two regex groups, one of them contains the values in the middle (.*)
                Match innerMatch = Regex.Match(match.Value, @"%DateTime\((.*)\)%");
                string innerValue = String.Empty;

                 
                foreach (Group item in innerMatch.Groups)
                {
                    if (item.Value != match.Value)
                    { 
                        innerValue = item.Value;
                        break;
                    }
                        
                }

                try
                {
                    string dateformat = DateTime.Now.ToString(innerValue);
                    input = input.Replace(match.Value, dateformat);
                }
                catch (System.FormatException)
                {

                    throw new Exception("Invalid dateformat " + innerMatch.Value);
                }

                return input;
                
        }

        string SearchDistinguishedFields(string input, IBaseMessage msg)
        {
            object val = null;

            for (int i = 0; i < msg.Context.CountProperties; i++)
            {
                string name = String.Empty;
                string ns = String.Empty;

                object context_val = msg.Context.ReadAt(i, out name, out ns);
                
                if (ns.EndsWith("btsDistinguishedFields") == true && DistinguishedString(name).EndsWith(input))
                {
                    val = context_val;
                    break;
                }
                
            }

            return (string)val;
        }
        string DistinguishedString(string contextname)
        {

            MatchCollection col = Regex.Matches(contextname, "'([a-zA-Z]+)'");
            StringBuilder builder = new StringBuilder(col.Count);

            foreach (Match item in col)
            {
                //First group contains single qoutes, second group value inside single qoutes ( )
                builder.AppendFormat("~{0}", item.Groups[1].Value);

            }

            return builder.ToString();
        }

        string CheckContextDateTime(string input, Match match, DateTime fileDateTime,string macro)
        {

            //include a percent ("%") format specifier before the single custom date and time specifier. like d,m e.t.c
            //bellow creates two regex groups, one of them contains the values in the middle (.*)
            Match innerMatch = Regex.Match(match.Value, String.Format(@"%{0}\((.*)\)%",macro));
            string innerValue = String.Empty;


            foreach (Group item in innerMatch.Groups)
            {
                if (item.Value != match.Value)
                {
                    innerValue = item.Value;
                    break;
                }

            }

            try
            {
                string dateformat = fileDateTime.ToString(innerValue);
                input = input.Replace(match.Value, dateformat);
            }
            catch (System.FormatException)
            {

                throw new Exception(String.Format("Invalid dateformat {0} for macro {1}" + innerMatch.Value,macro));
            }

            return input;

        }

      
        #endregion

        /// <summary>
        /// Loads configuration property for component.
        /// </summary>
        /// <param name="pb">Configuration property bag.</param>
        /// <param name="errlog">Error status (not used in this code).</param>
        public void Load(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, Int32 errlog)
        {
            return;

        }

        /// <summary>
        /// Saves current component configuration into the property bag.
        /// </summary>
        /// <param name="pb">Configuration property bag.</param>
        /// <param name="fClearDirty">Not used.</param>
        /// <param name="fSaveAllProperties">Not used.</param>
        public void Save(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, bool fClearDirty, bool fSaveAllProperties)
        {
            return;

        }
       
    }


    
}
