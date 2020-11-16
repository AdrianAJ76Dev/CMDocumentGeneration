// Standard libraries
using System;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;

// Xml libraries
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

// Open XML libraries
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.CustomXmlDataProperties;

namespace CMDocumentGeneration.Models
{
    // A sole source letter "is-a" CMDocument
    public class CMSoleSourceLetter : CMDocument
    {        
        public class Contact : ClientInfo{
            public string Signatory {get;set;}
            public enum InstitutionType : byte {K12=0,HED=1}
            public InstitutionType instType=InstitutionType.K12;
        }
       
        private class XmlPrimaryContact : customXML{
            public XmlPrimaryContact(){
               xmlElementName="ssl";
               xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/SSL/Contact";
           }           
            public void SerializeDataToXml(Contact PrimaryContact){
                attXmlRoot.ElementName=xmlElementName;
                attXmlRoot.Namespace=xmlNamespace;
                attXmlAttributes.XmlRoot=attXmlRoot;
                attXmlAttributeOverrides.Add(typeof(Contact), attXmlAttributes);
                CustomNamespaces.Add(xmlElementName, xmlNamespace);
                MemoryStream ms = new MemoryStream();
                XmlSerializer XmlDoc = new XmlSerializer(typeof(Contact), attXmlAttributeOverrides);
                XmlDoc.Serialize(ms, PrimaryContact, CustomNamespaces);            
                ms.Position = 0;
                AzureResources.SaveCustomXmlFile(ms,fileName);               
            }
        }

        private XmlPrimaryContact xmlSSLPrimaryContact;
       
 
        /*  This holds an autotext name*/
        public CMSoleSourceLetter(){
            xmlSSLPrimaryContact=new XmlPrimaryContact();
        }

        public void Generate(Contact PrimaryContact){
            /*  Create the basic Contracts Management document based on a Contracts Management template 
            *   Currently the templates will be stored in Azure. Later, I will add the ability to grab
            *   the templates from SharePoint or Amazon Web Services
            *   09.04.2020
            *   First - Create just the Word document based on any template found 
            */
            templateName="Sole-Source-Letter-v6.dotx"; // Name of Word temaplate            
            fileName="CM-Sole-Source-Letter-"+PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".docx"; // Name of Word document
            xmlSSLPrimaryContact.FileName="CM-"+xmlSSLPrimaryContact.XMLElementName.ToUpper()+"-"+PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
                  
            CreateWordDocument(templateName,fileName);

            MemoryStream generatedDocument=AzureResources.GetGeneratedDocument(fileName);
            //OpenXmlElement AutoText;
            using(WordprocessingDocument newDocument = WordprocessingDocument.Open(generatedDocument, true))
            {
                MainDocumentPart mdpNewDocument = newDocument.MainDocumentPart;
                string linkID;
                xmlSSLPrimaryContact.SerializeDataToXml(PrimaryContact);
                xmlSSLPrimaryContact.InsertCustomXmlData(xmlSSLPrimaryContact.FileName,xmlSSLPrimaryContact.XMLNS,fileName, out linkID);
            }
        }

    }
}