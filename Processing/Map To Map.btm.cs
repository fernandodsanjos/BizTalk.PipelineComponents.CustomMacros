namespace BizTalkComponents.PipelineComponents {
    
    
    [Microsoft.XLANGs.BaseTypes.SchemaReference(@"BizTalkComponents.PipelineComponents.Schema_Transform_Source", typeof(global::BizTalkComponents.PipelineComponents.Schema_Transform_Source))]
    [Microsoft.XLANGs.BaseTypes.SchemaReference(@"BizTalkComponents.PipelineComponents.Schema_Transform_Target", typeof(global::BizTalkComponents.PipelineComponents.Schema_Transform_Target))]
    public sealed class XSLTransform_Map_To_Map : global::Microsoft.XLANGs.BaseTypes.TransformBase {
        
        private const string _strMap = @"<?xml version=""1.0"" encoding=""UTF-16""?>
<xsl:stylesheet xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" xmlns:msxsl=""urn:schemas-microsoft-com:xslt"" xmlns:var=""http://schemas.microsoft.com/BizTalk/2003/var"" exclude-result-prefixes=""msxsl var"" version=""1.0"" xmlns:ns0=""http://XSLTransform.Schema.Schema"">
  <xsl:output omit-xml-declaration=""yes"" method=""xml"" version=""1.0"" />
  <xsl:template match=""/"">
    <xsl:apply-templates select=""/ns0:MotherOf_x0020_ALLRoots"" />
  </xsl:template>
  <xsl:template match=""/ns0:MotherOf_x0020_ALLRoots"">
    <ns0:Targhet>
      <Record>
        <ID>
          <xsl:value-of select=""Record/ID/text()"" />
        </ID>
        <MessageID>
          <xsl:choose>
            <xsl:when test=""function-available('msxsl:Read')"">

              <xsl:value-of select=""msxsl:Read('MessageID')"" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>ms xsl</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </MessageID>
      </Record>
    </ns0:Targhet>
  </xsl:template>
</xsl:stylesheet>";
        
        private const string _strArgList = @"<ExtensionObjects />";
        
        private const string _strSrcSchemasList0 = @"BizTalkComponents.PipelineComponents.Schema_Transform_Source";
        
        private const global::BizTalkComponents.PipelineComponents.Schema_Transform_Source _srcSchemaTypeReference0 = null;
        
        private const string _strTrgSchemasList0 = @"BizTalkComponents.PipelineComponents.Schema_Transform_Target";
        
        private const global::BizTalkComponents.PipelineComponents.Schema_Transform_Target _trgSchemaTypeReference0 = null;
        
        public override string XmlContent {
            get {
                return _strMap;
            }
        }
        
        public override string XsltArgumentListContent {
            get {
                return _strArgList;
            }
        }
        
        public override string[] SourceSchemas {
            get {
                string[] _SrcSchemas = new string [1];
                _SrcSchemas[0] = @"BizTalkComponents.PipelineComponents.Schema_Transform_Source";
                return _SrcSchemas;
            }
        }
        
        public override string[] TargetSchemas {
            get {
                string[] _TrgSchemas = new string [1];
                _TrgSchemas[0] = @"BizTalkComponents.PipelineComponents.Schema_Transform_Target";
                return _TrgSchemas;
            }
        }
    }
}
