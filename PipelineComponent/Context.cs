using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BizTalkComponents.PipelineComponents
{
    public class Context
    {
        private Dictionary<string, string> lst = new Dictionary<string, string>();
        private const string systemPropertiesNamespace = "http://schemas.microsoft.com/BizTalk/2003/system-properties";
        //This signature does not work in the map => Read(string name,string ns = systemPropertiesNamespace)


        public string Read(string name)
        {

            return this.Read(name, systemPropertiesNamespace);
        }

        public string Read(string name, string ns)
        {

            string property = String.Empty;

            lst.TryGetValue(String.Format("{0}#{1}", ns, name), out property);

            return property == null ? String.Empty : property;
        }

        public void Add(string name, string value, string ns = systemPropertiesNamespace)
        {
            if (lst.ContainsKey(String.Format("{0}#{1}", ns, name)) == false)
            {
                lst.Add(String.Format("{0}#{1}", ns, name), value);
            }
        }
    }
}
