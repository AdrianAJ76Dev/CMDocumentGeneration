// Standard libraries
using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Collections.Generic;

// Xml libraries
using System.Linq;
using System.Xml.Linq;
using System.Xml.Serialization;

// Open XmL libraries
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Validation;

/*  11.15.2020
*   The class defining Repeated Sections and Repeating Section Items
*   ONLY Starts exists in Word 2013*/
using Wrd2013=DocumentFormat.OpenXml.Office2013.Word;


namespace CMDocumentGeneration.Models
{
    public class CMDocument{
        //  CMDocuments "has a" classes that will delegate functionality and be inherited by derived classes
        /*  Contact It is in sole source letters and contracts
        *   This should be the common information that ALWAYS
        *   comes from the Salesforce exported JSON
        */
        /*  Every document generated needs a namespace and a namespace prefix */
        protected string templateName;
        protected string fileName;
        public string FileName{
            get{return fileName;}
        }

        protected CMDocument(){
            fileName = ".docx";
        }
        
        /* 05.15.2021 - Word templates are now "configured" in a JSON file 
        * because I want to define specific behavior for each template
        * and default values for each template. This is part of an effort
        * to make this web api that creates Word documents more flexible.
        */
        
        

        /*
        protected string ReadJSONFileForTemplate(string fileName, string templateName){
            const string TEMPLATE_NAME = "TemplateName";
            string template;
            MemoryStream msTemplates = new MemoryStream();
            
            msTemplates=AzureResources.GetJSONFile(fileName);

            ReadOnlySpan<byte> jsonReadOnlySpan = msTemplates.ToArray();
            var reader = new Utf8JsonReader(jsonReadOnlySpan);
            while (reader.Read()){
                JsonTokenType tokenType = reader.TokenType;
                switch (tokenType)
                {
                    case JsonTokenType.PropertyName:
                        if (reader.ValueTextEquals(Encoding.UTF8.GetBytes(TEMPLATE_NAME)))
                        {
                            reader.Read();
                            if (reader.GetString()==templateName)
                            {
                                template=;
                            }
                        }
                        break;
                    default:
                        break;
                }
            }


            return template;
        }
        */


        protected class customXML{
        /*  09.04.2020
        *   The "Word" component classes
        *   Custom XmL
        *   This is the common merge data that will populate every Word Document
        *   They are like "Document Variables" that I've used in my macros and
        *   that I've used in all of the code I wrote while at Micro-Modeling Associates
        *   I learned this from Jeffery Jones and Frank Kristeller.
        */

            protected string xmlNamespace;
            protected string xmlElementName;
            protected string fileName = ".xml";
            protected string dataBindingXPath;
            public string XMLElementName{
                get{return xmlElementName;}
            }
            public string XMLNS{
                get{return xmlNamespace;}
            }
            public string FileName{
                get{return fileName;}
                set{fileName=value;}
            }
            // Create an XmlAttributes to override the default root element.
            /*  move all Xml classes into SerializedDataToXml because it's not used
            *   anywhere else and create properties that get the element name and 
            *   namespaces. Plus make it able to take ANY POCO coming from a form 
            */
            protected XmlRootAttribute attXmlRoot = new XmlRootAttribute();
            protected XmlAttributes attXmlAttributes = new XmlAttributes();
            protected XmlAttributeOverrides attXmlAttributeOverrides = new XmlAttributeOverrides();
            protected XmlSerializerNamespaces CustomNamespaces = new XmlSerializerNamespaces();
            public customXML(){
                //  Create the XmL that will be merged into the created document
            }
           
            public void SerializeDataToXml(ClientInfo primaryContact)
            {
                attXmlRoot.ElementName=xmlElementName;
                attXmlRoot.Namespace=xmlNamespace;
                attXmlAttributes.XmlRoot=attXmlRoot;
                attXmlAttributeOverrides.Add(typeof(ClientInfo), attXmlAttributes);
                CustomNamespaces.Add(xmlElementName, xmlNamespace);
                MemoryStream ms = new MemoryStream();
                XmlSerializer XmlDoc = new XmlSerializer(typeof(ClientInfo), attXmlAttributeOverrides);
                XmlDoc.Serialize(ms, primaryContact, CustomNamespaces);            
                ms.Position = 0;
                AzureResources.SaveCustomXmlFile(ms,fileName);
                ms.Close();
            }

            public void InsertCustomXmlData(string xmlfilename, string ID, string fileName, out string datastoreID)
            {
                MemoryStream ms = AzureResources.GetCustomXmlFile(xmlfilename);
                //  Add custom xml or "The Basic Contact Data"
                MemoryStream generatedDocument=AzureResources.GetGeneratedDocument(fileName);
                using(WordprocessingDocument newDocument = WordprocessingDocument.Open(generatedDocument, true))
                {
                    MainDocumentPart mdpNewDocument = newDocument.MainDocumentPart;

                    /*  09.20.2020 
                    *   In order to link to content controls, the custom XML part MUST have a custom XML Propertie part
                    */

                    /* 11.11.2020 Rethinking creating a NEW custom xml part everytime this routine is called.
                    *   I don't really see the need
                    *   Write it so it's only created once and THEN reference
                    */
                    CustomXmlPart cxml=mdpNewDocument.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                    
                    CustomXmlPropertiesPart cxmlProps = cxml.AddNewPart<CustomXmlPropertiesPart>();                   
                    DataStoreItem dsi = new DataStoreItem();
                    SchemaReference sr = new SchemaReference();
                    dsi.ItemId=Guid.NewGuid().ToString("B");
                    datastoreID=dsi.ItemId;
                    sr.Uri=ID;
                    dsi.SchemaReferences=new SchemaReferences();
                    dsi.SchemaReferences.AppendChild<SchemaReference>(sr);
                    cxmlProps.DataStoreItem=dsi;
                    cxmlProps.DataStoreItem.Save();

                    ms.Position = 0;
                    cxml.FeedData(ms);
                    newDocument.Save();
                    newDocument.Close();
                }
                AzureResources.SaveGeneratedDocument(generatedDocument, fileName);
                ms.Close();
                generatedDocument.Close();           
            }
        }   
        public class autoText{
            protected string name;
            protected string templateName;
            protected WordprocessingDocument wrdTemplate;
            protected GlossaryDocumentPart gdPart;
            protected GlossaryDocument gd;
            protected DocPart dpAutoText;
            protected DocPartBody dpbAutoText;
            protected MemoryStream msAutoText;
            protected string strAutoText;
            protected OpenXmlElement ccAutoText;

            public OpenXmlElement RetrieveAutoText(string templateName, string autoTextName)
            {
                MemoryStream wordTemplate=AzureResources.GetWordTemplate(templateName);
                using (WordprocessingDocument wrdTemplate = WordprocessingDocument.Open(wordTemplate, true))
                {
                    MainDocumentPart mdpWrdTmpl = wrdTemplate.MainDocumentPart;
                    gdPart=mdpWrdTmpl.GlossaryDocumentPart;
                    gd=gdPart.GlossaryDocument;

                    /*  09.23.2020
                        Find the right AutoText Entry using LINQ
                    */
                    var dpatxt = (from dpAutoText in gd.DocParts
                                where dpAutoText.GetFirstChild<DocPartProperties>().DocPartName.Val.Value == autoTextName
                                select dpAutoText).FirstOrDefault();

                    dpbAutoText=dpatxt.GetFirstChild<DocPartBody>();
                    /*  09.23.2020
                    *   Pull out the SdtElement/SdtBlock that contains the xml for the autotext in the DocPartBody
                    *   The AutoText defaults to being put on a page and is contained in a content control already.
                    */
                    ccAutoText=dpbAutoText.GetFirstChild<SdtElement>();
                    return ccAutoText.CloneNode(true);
                }
            }

            public void InsertAutoText(string fileName, OpenXmlElement currAutoText){
                MemoryStream generatedDocument=AzureResources.GetGeneratedDocument(fileName);
                using(WordprocessingDocument newDocument = WordprocessingDocument.Open(generatedDocument, true))
                {
                    MainDocumentPart mdpNewDocument = newDocument.MainDocumentPart;
                    Body body = mdpNewDocument.Document.Body;
                    var finalSectionBreak=body.LastChild;
                    finalSectionBreak.InsertBeforeSelf<OpenXmlElement>(currAutoText);
                    newDocument.Save();
                }
                AzureResources.SaveGeneratedDocument(generatedDocument, fileName);           
            }
        }
        protected class contentControl{
            protected string wordFileName;
            public contentControl(string wordDocumentFileName){
                wordFileName=wordDocumentFileName;
            }
                
            /*  09.18.2020
            *   This should be the LAST piece of the process
            *   The document is complete when the content controls are linked
            */
            public void BindContentControls(string xmlfilename, string wrdFileName, string ns, string prefix, string datastoreID, string contentControlName){
                /*  Content Controls */
                List<SdtElement> ccAll;                   

                /*  09.18.2020
                *   This brings back ALL of the content controls
                *   Use SdtElement instead of SdtBlock for content controls
                *   SdtBlock ONLY retrieves content controls that have a paragraph
                */
                //  Add custom xml or "The Basic Contact Data"
                MemoryStream generatedDocument=AzureResources.GetGeneratedDocument(wrdFileName);
                using(WordprocessingDocument newDocument = WordprocessingDocument.Open(generatedDocument, true))
                {
                    MainDocumentPart mdp = newDocument.MainDocumentPart;
                    /*  11.12.2020 
                    *   Get ALL the Content Controls in the document
                    *   ccAll=mdp.Document.Body.Elements<SdtElement>().ToList();
                    */                    
                    ccAll=mdp.Document.Descendants<SdtElement>().ToList();

                    
                    /*  11.12.2020
                    *   Out of ALL the Content Controls
                    *   ONLY get the Content Controls that MATCH the Prefix:
                    *   agrmt-m, agrmt-pc, agrmt-q 
                    */
                    IEnumerable<SdtElement> ccCurrSet = 
                        from cc in ccAll
                        where cc.Descendants<Tag>().FirstOrDefault().Val.Value.Contains(prefix) // Can I just store the prefix in the tag?
                        select cc;

                    /*  11.12.2020
                    *   IF we're looking at QUOTE Content Controls
                    *   there may be levels and a repeating Content Control
                    */
                    //  11.18.2020 Doing the Quote/Budget Schedule
                    if (prefix=="agrmt-q")
                    {
                        
                        /*  11.13.2020
                        *   Get the content Control containing the repeating rows
                        */
                        //  11.17.2020 Get number of items in quote and cycle through
                        IEnumerable<SdtElement> ccLineItemsContainer =
                            from cc in ccAll
                            where cc.Descendants<SdtAlias>().FirstOrDefault<SdtAlias>()
                                    .Val.Value.Contains(contentControlName)
                            select cc;
                        
                        DoBinding(ccLineItemsContainer, datastoreID, ns, prefix, contentControlName);

                        //  11.18.2020 Select all the content controls in that repeating row
                        IEnumerable<SdtElement> ccRepeatingRowValues =
                            from cc in ccAll
                            where cc.Descendants<Tag>().FirstOrDefault<Tag>()
                                    .Val.Value.Contains("/rr")
                            select cc;
                        // Code above uses tag of content control to find repeating rows so DON'T CHANGE THE TAGS in the content controls!!!

                        DoBinding(ccRepeatingRowValues, datastoreID, ns, prefix, contentControlName);
                    }                        
                    else
                        DoBinding(ccCurrSet, datastoreID, ns, prefix, contentControlName);
                    
                   newDocument.Save();
                   newDocument.Close();
                }

                AzureResources.SaveGeneratedDocument(generatedDocument, wrdFileName); 
                generatedDocument.Close();
           }

            private void DoBinding(IEnumerable<SdtElement> specificContentControls, string datastoreID, string ns, string prefix, string contentControlName){
                SdtAlias ccName;
                SdtProperties ccProps;
                DataBinding ccDataBinding;

                /*  11.12.2020
                *   Constants for XPath
                */
                int customXmlElementCount=1;
                foreach (SdtElement wrdCC in specificContentControls)
                {
                    ccProps=wrdCC.GetFirstChild<SdtProperties>();
                    ccName=ccProps.GetFirstChild<SdtAlias>();
                    if (ccName != null)                                                     // If the name is null, the element isn't a content control
                    {
                        //  11.5.2020 XPath query for databinding stored here 
                        //ccTag=wrdCC.Descendants<Tag>().FirstOrDefault<Tag>();               
                        // I don't think I'm using this object anywhere.
                                                                                                                
                        ccDataBinding = new DataBinding();
                        ccDataBinding.PrefixMappings=string.Format("xmlns:ns='{0}'", ns);

                        //  11.18.2020 Am I working with the Quote Container Content Control
                        if (prefix=="agrmt-q")
                        {
                            ccDataBinding.PrefixMappings=string.Format("xmlns:ns0='{0}'", ns);

                            switch (ccName.Val.Value)
                            {
                                case "LineItems":
                                    //  "//ns:agrmt-q[1]/ns:LineItems[1]/ns:LineItem" />"
                                    ccDataBinding.XPath=string.Format("//ns0:{0}[1]/ns0:LineItems[1]/ns0:LineItem", prefix);
                                    break;

                                case "DistrictCostItems":
                                    ccDataBinding.XPath=string.Format("//ns0:{0}[1]/ns0:DistrictCostItems[1]/ns0:DistrictCostItem", prefix);
                                    break;

                                default:
                                    //  "//ns:agrmt-q/ns:LineItems/ns:LineItem/ns:ProductName[1]"
                                    ccDataBinding.XPath=string.Format("//ns0:{0}[1]/ns0:{3}[1]/ns0:{4}[{2}]/ns0:{1}[{2}]", prefix, ccName.Val.Value, customXmlElementCount, contentControlName,contentControlName.TrimEnd('s'));
                                    break;
                            }

                     }
                        else
                        {
                            //  "//ns0:agrmt-m[1]/ns0:ContractNumber[1]" 11.13.2020
                            ccDataBinding.XPath=string.Format("//ns:{0}/ns:{1}[1]", prefix, ccName.Val.Value);
                        }
                        
                       ccDataBinding.StoreItemId=datastoreID;                              //  "{E45CD94D-6275-426C-A007-762425B85F33}"
                       ccProps.Append(ccDataBinding);                       
                    }
                }
            }
        }
        public class ClientInfo{
            public string FirstName{get;set;}
            public string LastName{get;set;}
            public string Title{get;set;}
            public string Institution{get;set;}
            public string MailingStreet{get;set;}
            public string MailingCity{get;set;}
            public string MailingState{get;set;}
            public string MailingPostalCode{get;set;}
            public string Phone{get;set;}
            public string Email{get;set;}
         }        
        protected class CBSignatory{
            string FirstName { get; set;}
            string LastName { get; set;}
            string Title { get; set;}
        }
        protected void CreateWordDocument(string templateName, string fileName)
        {
            /*  "wordDocument" is actually the template used to create the document. This functionality is 
            *   very important because I want the user to have the ability to create as many templates
            *   as they want to fit whatever the business need is.
            */
            MemoryStream wordDocument=AzureResources.GetWordTemplate(templateName);
            using (WordprocessingDocument newdocument = WordprocessingDocument.Open(wordDocument, true))
            {
                /*  newdocument is STILL a template and Word will not open it so change it to a document! 
                *   This is ALSO changing "wordDocument" to a document. ImPORTANT: The WordprocessingDocument
                *   AND the memoryStream IS the SAmE!!*/
                newdocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                newdocument.Save();
            }
            /*  fileName is the name of the document i.e
            *   Adrian-Jones-Cm-Sole-Source-Letter.docx or 
            *   Adrian-Jones-Cm-SpringBoard-materials-Contract.docx */
            AzureResources.SaveGeneratedDocument(wordDocument, fileName);           
            wordDocument.Close();
        }
    }
}     