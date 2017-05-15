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
using BizTalkComponents.Utils;
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
        

        readonly string ns_BTS = "http://schemas.microsoft.com/BizTalk/2003/system-properties";
        readonly string ns_FILE = "http://schemas.microsoft.com/BizTalk/2003/file-properties";
        readonly string macros = @"%DateTime\([%YyMmDdTHhSs:\-fzZkK\s]*\)%|%FileDateTime\([%YyMmDdTHhSs:\-fzZkK\s]*\)%|%FileNameOnly%|%Context\(\D*\.\D*\)%";


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
                 else if (match.Value.StartsWith("%FileDateTime("))
                 {
                     val = pInMsg.Context.Read("FileCreationTime", ns_FILE);

                     if (val == null)
                         continue;

                     //DateTime dt = DateTime.ParseExact((string)val, "yyyy-MM-dd hh:mm:ss", ci);

                     transport = CheckFileDateTime(transport, match,(DateTime)val);
                 }
                 else if (match.Value == "%FileNameOnly%")
                 {
                     val = pInMsg.Context.Read("ReceivedFileName", ns_FILE);

                     if (val == null || String.IsNullOrWhiteSpace((string)val))
                         continue;

                     string receivedFileName = Path.GetFileNameWithoutExtension((string)val);
                     
                     transport = transport.Replace("%FileNameOnly%", receivedFileName);
                 }
                 else if (match.Value.StartsWith("%Context("))
                 {
                     transport = GetContext(transport, match, pInMsg);
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

            pInMsg.Context.Promote("OutboundTransportLocation", ns_BTS, (object)transport);

            return pInMsg;
        }



        #endregion

        #region Private methods

        object InitCustomMacros(IBaseMessage msg,string ctxName)
        {
            object val = null;

                for (int i = 0; i < msg.Context.CountProperties; i++)
                {
                    string name = String.Empty;
                    string ns = String.Empty;

                    val = msg.Context.ReadAt(i, out name, out ns);

                    if (ns.StartsWith("http://schemas.microsoft.com/BizTalk/") == false && ctxName == name)
                    {
                        break;
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

            string[] context = innerValue.Split('.');

            string ns = String.Empty;
            propertys.TryGetValue(context[0], out ns);
             

            if (String.IsNullOrWhiteSpace(ns) && context[0] != "CST")
                return input.Replace(match.Value, "");

            if(context[0] == "CST")
            {
                val = InitCustomMacros(msg, context[1]);
            }
            else 
            { 
                val = msg.Context.Read(context[1], ns); 
            }
            

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


        string CheckFileDateTime(string input, Match match,DateTime fileDateTime)
        {

            //include a percent ("%") format specifier before the single custom date and time specifier. like d,m e.t.c
            //bellow creates two regex groups, one of them contains the values in the middle (.*)
            Match innerMatch = Regex.Match(match.Value, @"%FileDateTime\((.*)\)%");
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

                throw new Exception("Invalid dateformat " + innerMatch.Value);
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
