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

// 10.27.2020 JSON library
using System.Text.Json;

namespace CMDocumentGeneration.Models
{
   /*   Agreement/Amendment
    *   ex. "http//www.collegeboard.org/sdp/contractsmanagement/Agreement"
    *   ex. "http//www.collegeboard.org/sdp/contractsmanagement/Amendment"
    */
    // A contract "is-a" CMDocument
    public class CMContract : CMDocument
    {        
        /*  09.10.2020
        *   Common data for a contracts management contract
        *   Special to the College Board expressed as 
        *   enumerations
        */
        public enum ContractType{agreement, amendment};
        /*  The product type becomes part of the name */
        public enum productType:Int16 {
            SpringBoard=0, 
            PreAP=1, 
            EnrollmentPlanningServiceUnlimited=2, 
            StudentSearchService=3,
            SegmentAnalysisService=4, 
            PowerFAIDSInitial=5,
            PowerFAIDSUpgrade=6,
            PowerFAIDSMaintenance=7,
            NetPartnerInitial=8,
            NetPartnerSupport=9,
            Profile=10,
            InterestInMyCollege=11
        };
        public enum instType:Int16
        {
            K12=0,
            HED=1
        };

        public class MainContract{
            public instType InstitutionType{get;set;}
            public string ContractNumber{get;set;}
            public string AccountName{get;set;}
            public string CreateDate{get;set;}
            public string ContractStartDate{get;set;}
            public string ContractEndDate{get;set;}
            public string Term{get;set;}
            public string ImplementationYears{get;set;}
            public string ClientSignatory{get;set;}
            public string ClientTitle{get;set;}
            public string Signatory{get;set;}
            public string SignatoryTitle{get;set;}
        }
        public class Amendment{}
        public class Rider : autoText{
            public string ProductName{get;set;}
            public string Program{get;set;}
            public string AutoTextName{get;set;}
        }
        public class Quote : autoText{
            public string AutoTextQuoteName{get;set;}
            public List<LineItem> LineItems{get;set;}
        }
        
        public class LineItem
        {
            public string ProductName{get;set;}
            public string StartDate{get;set;}
            public string EndDate{get;set;}
            public string Quantity{get;set;}
            public string TotalCost{get;set;}
        }
        public class autoTextSettings{
            // reads set values in a table to determine what is AutoText and what is custom XmL
            public productType contractRiderID{get;set;}
            public string AutoTextName{get;set;}
            public string Product{get;set;}
        }

        private class XmlMainContract : customXML{
            public XmlMainContract(){
                xmlElementName="agrmt-m";
                xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Agreement";
            }
            public void SerializeDataToXml(MainContract Agreement){
                attXmlRoot.ElementName=xmlElementName;
                attXmlRoot.Namespace=xmlNamespace;
                attXmlAttributes.XmlRoot=attXmlRoot;
                attXmlAttributeOverrides.Add(typeof(MainContract), attXmlAttributes);
                CustomNamespaces.Add(xmlElementName, xmlNamespace);
                MemoryStream ms = new MemoryStream();
                XmlSerializer XmlDoc = new XmlSerializer(typeof(MainContract), attXmlAttributeOverrides);
                XmlDoc.Serialize(ms, Agreement, CustomNamespaces);            
                ms.Position = 0;
                AzureResources.SaveCustomXmlFile(ms,fileName);               
            }
        }
        private class XmlAmendment : customXML{
            public XmlAmendment(){
                xmlElementName="agrmt-ma";
                xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Amendment";
            }
        }
        private class XmlRider : customXML{
            /*  11.11.2020
            *   This is for any data merged into the rider i.e. custom Xml that goes in the rider
            */
            public XmlRider(){
                xmlElementName="agrmt-r";
                xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Agreement/Rider";
            }
        }
        private class XmlQuote : customXML{
            public XmlQuote(){
               xmlElementName="agrmt-q";
               xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Agreement/Quote";
            }

            public void SerializeDataToXml(Quote AgreementQuote){
                attXmlRoot.ElementName=xmlElementName;
                attXmlRoot.Namespace=xmlNamespace;
                attXmlAttributes.XmlRoot=attXmlRoot;
                attXmlAttributeOverrides.Add(typeof(Quote), attXmlAttributes);
                CustomNamespaces.Add(xmlElementName, xmlNamespace);
                MemoryStream ms = new MemoryStream();
                XmlSerializer XmlDoc = new XmlSerializer(typeof(Quote), attXmlAttributeOverrides);
                XmlDoc.Serialize(ms, AgreementQuote, CustomNamespaces);            
                ms.Position = 0;
                AzureResources.SaveCustomXmlFile(ms,fileName);
            }               
        }

        private class XmlPrimaryContact : customXML{
            public XmlPrimaryContact(){
               xmlElementName="agrmt-pc";
               xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Agreement/Contact/Primary";
            }
        }
        
        private contentControl cc;
        private XmlMainContract xmlMainContract = new XmlMainContract();
        private XmlPrimaryContact xmlPrimaryContact = new XmlPrimaryContact();
        private XmlQuote xmlAgreementQuote = new XmlQuote();

        /*  IF using nested classes MUST use AUTOMATIC PROPERTIES to make them work with Model Binding!!! */
        public MainContract Agreement{get;set;}
        public ClientInfo PrimaryContact{get;set;}
        public List<Rider> AgreementRiders{get;set;} 
        public List<autoTextSettings> SFProductTranslateToRider{get;set;}
        public Quote AgreementQuote{get;set;}
        
        public void Generate(CMContract cmNewContract)
        {
            /*  Create the basic Contracts Management document based on a Contracts Management template 
            *   Currently the templates will be stored in Azure. Later, I will add the ability to grab
            *   the templates from SharePoint or Amazon Web Services
            *   09.04.2020
            *   First - Create just the Word document based on any template found 
            */


            /*  09.27.2020
            *   Pick which template to use: Either for K12 Contracts or Higher Ed. Contracts 
            */
            switch (Agreement.InstitutionType)
            {
                case instType.K12:
                    templateName="K12 Template.dotx";            
                    break;
                case instType.HED:
                    //templateName="HED Template v2.dotx"; // Without content control in the autotext
                    templateName="HED Template.dotx"; // With content control in the autotext
                    break;
            }
            
            fileName="CM-Contract-"
                +PrimaryContact.FirstName+"-"
                +PrimaryContact.LastName+"-"
                +Agreement.ContractNumber
                +".docx"; // Name of Word document

            CreateWordDocument(templateName,fileName);

            // Choose the rider
            MemoryStream msAutoTextNamesJSON = AzureResources.GetJSONFile("AutoText.JSON");
            ReadOnlySpan<byte> jsonReadOnlySpan = msAutoTextNamesJSON.ToArray();
            SFProductTranslateToRider=JsonSerializer.Deserialize<List<autoTextSettings>>(jsonReadOnlySpan);
            msAutoTextNamesJSON.Close();

            List<autoTextSettings> AutoTextFound = new List<autoTextSettings>();
            foreach (var currRider in AgreementRiders)
            {
                foreach (var autoTextName in SFProductTranslateToRider)
                {
                    if (currRider.ProductName.Contains(autoTextName.Product))
                    {
                        AutoTextFound.Add(autoTextName);
                    }
                }
            }

            OpenXmlElement AutoText;
            //  11.06.2020 Open the generated Word document once and insert all that needs to be inserted!
            // 11.06.2020 Insert Riders by name sent from Salesforce using product names
            foreach (var currRider in AutoTextFound)
            {
                switch (currRider.contractRiderID)
                {
                    case productType.SpringBoard:
                        Rider SpringBoard = new Rider();
                        AutoText = SpringBoard.RetrieveAutoText(templateName,currRider.AutoTextName);
                        SpringBoard.InsertAutoText(fileName,AutoText);
                        break;

                        case productType.PreAP:
                            Rider PreAP = new Rider();
                            AutoText = PreAP.RetrieveAutoText(templateName,currRider.AutoTextName);
                            PreAP.InsertAutoText(fileName,AutoText);
                            break;

                        case productType.EnrollmentPlanningServiceUnlimited:
                            Rider EPSUnlimited = new Rider();
                            AutoText = EPSUnlimited.RetrieveAutoText(templateName,currRider.AutoTextName);
                            EPSUnlimited.InsertAutoText(fileName, AutoText);
                            break;

                        case productType.StudentSearchService:
                            Rider SSSUnlimited = new Rider();
                            AutoText = SSSUnlimited.RetrieveAutoText(templateName,currRider.AutoTextName);
                            SSSUnlimited.InsertAutoText(fileName, AutoText);
                            break;

                        case productType.SegmentAnalysisService:
                            Rider SASUnlimited = new Rider();
                            AutoText = SASUnlimited.RetrieveAutoText(templateName,currRider.AutoTextName);
                            SASUnlimited.InsertAutoText(fileName,AutoText);
                            break;

                        case productType.PowerFAIDSInitial:
                            break;

                        case productType.Profile:
                            break;
                }
            }

            //  Put in the quote           
            OpenXmlElement qAutoText=AgreementQuote.RetrieveAutoText(templateName,AgreementQuote.AutoTextQuoteName);
            AgreementQuote.InsertAutoText(fileName,qAutoText);

            //  Add the custom xml or "Do the Merge"
            
            string linkID;
            cc=new contentControl(fileName);
            
            xmlMainContract.FileName="CM-Contract-"+xmlMainContract.XMLElementName.ToUpper()+"-";
            xmlMainContract.FileName+=PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
            xmlMainContract.SerializeDataToXml(cmNewContract.Agreement);
            xmlMainContract.InsertCustomXmlData(xmlMainContract.FileName,xmlMainContract.XMLNS,fileName,out linkID);
            
            cc.BindContentControls(xmlMainContract.FileName, fileName, xmlMainContract.XMLNS, xmlMainContract.XMLElementName, linkID);

            xmlPrimaryContact.FileName="CM-Contract-"+xmlPrimaryContact.XMLElementName.ToUpper()+"-";
            xmlPrimaryContact.FileName+=PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
            xmlPrimaryContact.SerializeDataToXml(cmNewContract.PrimaryContact);
            xmlPrimaryContact.InsertCustomXmlData(xmlPrimaryContact.FileName,xmlPrimaryContact.XMLNS,fileName,out linkID);

            cc.BindContentControls(xmlPrimaryContact.FileName,fileName,xmlPrimaryContact.XMLNS, xmlPrimaryContact.XMLElementName, linkID);

            xmlAgreementQuote.FileName="CM-Contract-"+xmlAgreementQuote.XMLElementName.ToUpper()+"-";
            xmlAgreementQuote.FileName+=PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
            xmlAgreementQuote.SerializeDataToXml(cmNewContract.AgreementQuote);
            xmlAgreementQuote.InsertCustomXmlData(xmlAgreementQuote.FileName,xmlAgreementQuote.XMLNS,fileName,out linkID);

            
            cc.BindContentControls(xmlAgreementQuote.FileName,fileName,xmlAgreementQuote.XMLNS, xmlAgreementQuote.XMLElementName, linkID);
        }
    }
}