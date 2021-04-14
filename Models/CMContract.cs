using System;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Reflection;

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
        public enum productType:byte {
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
            InterestInMyCollege=11,
            None=99
        };
        public enum instType:byte
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
            public string TotalPrice{get;set;}
            public int Term{get;set;}
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
            public List<LineItem> LineItems{get;set;}
        }

        public class SBComplexQuote{
            public List<SummaryItem> Summary{get;set;}
            public List<PaymentScheduleItem> PaymentSchedule{get;set;}
            public List<DistrictCostItem> DistrictCostItems{get;set;}
            public List<DistrictSavingsItem> DistrictSavings{get;set;}
            public void AddProcessedQuote(){
            /* 1.20.2021
            * 
            */
        }

        }
        
        // 1.21.2021 Classes that describe SB Quote JSON
        public class LineItem
        {
            public string ProductName{get;set;}
            public string StartDate{get;set;}
            public string EndDate{get;set;}
            public string Quantity{get;set;}
            public string TotalCost{get;set;}
        }
        public class SummaryItem
        {
            public string TypeOfCost{get;set;}
            public float Fees{get;set;}
            public float CostSavings{get;set;}
            public float DistrictCost{get;set;}
        }
        public class PaymentScheduleItem
        {
            public int YearNo{get;set;}
            public string Year{get;set;}
            public float Total{get;set;}
        }       
        public class DistrictCostItem
        {
            public string Product{get;set;}
            public int Quantity{get;set;}
            public float UnitPrice{get;set;}
            public float TotalDiscountAmount{get;set;}
            public float TotalPrice{get;set;}
        }
        public class DistrictSavingsItem
        {
            public string Product{get;set;}
            public int Quantity{get;set;}
            public float UnitPrice{get;set;}
            public float TotalDiscountAmount{get;set;}
            public float TotalPrice{get;set;}
        }

        // This is an idea of what's basically in all quotes. This is the data that is in every quote from Salesforce
        public class SFQuote
        {
            public float subTotal{get;set;}
            public float discount{get;set;}
            public float shippingHandling{get;set;}
            public float grandTotal{get;set;}
            public string product{get;set;}
            public int quantity{get;set;}
            public DateTime startDate{get;set;}
            public DateTime endDate{get;set;}
            public float catalogUnitPrice{get;set;}
            public float unitPriceAdjustment{get;set;}
            public float unitPrice{get;set;}
            public float totalDiscountAmount{get;set;}
            public float totalDiscountPercentage{get;set;}
            public float totalPrice{get;set;}
        }

        
        public class autoTextSettings{
            // reads set values in a table to determine what is AutoText and what is custom XmL
            public productType contractRiderID{get;set;}
            public string Product{get;set;}
            public string AutoTextRider{get;set;}
            public string AutoTextQuote{get;set;}
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
        // 1.19.2021 --- SERIOUSLY consider if this can be added as another class Quote inherits from!
        private class XmlQuote : customXML{
            public XmlQuote(){
               xmlElementName="agrmt-q";
               xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Agreement/Quote";
            }

            public string xpathContentControlName{get;set;}

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

        private class XmlInvoiceBilling : customXML{
            public XmlInvoiceBilling(){
               xmlElementName="agrmt-ibc";
               xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Agreement/Contact/InvoiceBilling";
            }
        }

        private class XmlTechnicalSupport : customXML{
            public XmlTechnicalSupport(){
               xmlElementName="agrmt-ts";
               xmlNamespace="http//www.collegeboard.org/sdp/contractsmanagement/Agreement/Contact/TechnicalSupport";
            }
        }

        
        private contentControl cc;
        private XmlMainContract xmlMainContract = new XmlMainContract();
        private XmlPrimaryContact xmlPrimaryContact = new XmlPrimaryContact();
        private XmlInvoiceBilling xmlInvoiceBilling = new XmlInvoiceBilling();
        private XmlTechnicalSupport xmlTechnicalSupport = new XmlTechnicalSupport();
        private XmlQuote xmlAgreementQuote = new XmlQuote();

        /*  IF using nested classes MUST use AUTOMATIC PROPERTIES to make them work with Model Binding!!! */
        public MainContract Agreement{get;set;}
        public ClientInfo PrimaryContact{get;set;}
        public ClientInfo InvoiceBilling{get;set;}
        public ClientInfo TechnicalSupport{get;set;}
        public List<Rider> AgreementRiders{get;set;} 
        public List<autoTextSettings> SFProductTranslateToRider{get;set;}
        public Quote AgreementQuote{get;set;}

        public SBComplexQuote MultiYearSBQuote{get;set;}
        
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
                    templateName="K12-Template.dotx";            
                    break;
                case instType.HED:
                    //templateName="HED Template v2.dotx"; // Without content control in the autotext
                    templateName="HED-Template.dotx"; // With content control in the autotext
                    break;
            }
            
            // Choose the rider
            MemoryStream msAutoTextNamesJSON = AzureResources.GetJSONFile("AutoText.JSON");
            ReadOnlySpan<byte> jsonReadOnlySpan = msAutoTextNamesJSON.ToArray();
            SFProductTranslateToRider=JsonSerializer.Deserialize<List<autoTextSettings>>(jsonReadOnlySpan);
            msAutoTextNamesJSON.Close();

            List<autoTextSettings> AutoTextFound = new List<autoTextSettings>();
            string productsInFileName = string.Empty;
            foreach (var currRider in AgreementRiders)
            {
                foreach (var autoTextName in SFProductTranslateToRider)
                {
                    if (currRider.ProductName.Contains(autoTextName.Product))
                    {
                        AutoTextFound.Add(autoTextName);
                        productsInFileName+=autoTextName.AutoTextRider;
                        productsInFileName+="-";
                    }
                }
            }

            productsInFileName=productsInFileName.Remove(productsInFileName.Length-1,1);
            fileName="CM-Contract-"
                +PrimaryContact.FirstName+"-"
                +PrimaryContact.LastName+"-"
                +productsInFileName
                +Agreement.ContractNumber
                +".docx"; // Name of Word document

            CreateWordDocument(templateName,fileName);

            OpenXmlElement AutoText;
            int quoteAutoTextIndex=0;
            productType chosenRider=productType.EnrollmentPlanningServiceUnlimited;

            //  11.06.2020 Open the generated Word document once and insert all that needs to be inserted!
            // 11.06.2020 Insert Riders by name sent from Salesforce using product names
            foreach (var currRider in AutoTextFound)
            {
                switch (currRider.contractRiderID)
                {
                    case productType.SpringBoard:
                        Rider SpringBoard = new Rider();
                        AutoText = SpringBoard.RetrieveAutoText(templateName,currRider.AutoTextRider);
                        SpringBoard.InsertAutoText(fileName,AutoText);
                        quoteAutoTextIndex=(int)productType.SpringBoard;
                        chosenRider = productType.SpringBoard;
                        break;

                    case productType.PreAP:
                        Rider PreAP = new Rider();
                        AutoText = PreAP.RetrieveAutoText(templateName,currRider.AutoTextRider);
                        PreAP.InsertAutoText(fileName,AutoText);
                        quoteAutoTextIndex=(int)productType.PreAP;
                        chosenRider=productType.PreAP;
                        break;

                    case productType.EnrollmentPlanningServiceUnlimited:
                        Rider EPSUnlimited = new Rider();
                        AutoText = EPSUnlimited.RetrieveAutoText(templateName,currRider.AutoTextRider);
                        EPSUnlimited.InsertAutoText(fileName, AutoText);
                        quoteAutoTextIndex=(int)productType.EnrollmentPlanningServiceUnlimited;
                        chosenRider=productType.EnrollmentPlanningServiceUnlimited;
                        break;

                    case productType.StudentSearchService:
                        Rider SSSUnlimited = new Rider();
                        AutoText = SSSUnlimited.RetrieveAutoText(templateName,currRider.AutoTextRider);
                        SSSUnlimited.InsertAutoText(fileName, AutoText);
                        quoteAutoTextIndex=(int)productType.StudentSearchService;
                        chosenRider=productType.StudentSearchService;
                        break;

                    case productType.SegmentAnalysisService:
                        Rider SASUnlimited = new Rider();
                        AutoText = SASUnlimited.RetrieveAutoText(templateName,currRider.AutoTextRider);
                        SASUnlimited.InsertAutoText(fileName,AutoText);
                        quoteAutoTextIndex=(int)productType.SegmentAnalysisService;
                        chosenRider=productType.SegmentAnalysisService;
                        break;

                    case productType.PowerFAIDSInitial:
                        quoteAutoTextIndex=(int)productType.PowerFAIDSInitial;
                        chosenRider=productType.PowerFAIDSInitial;
                        break;

                    case productType.Profile:
                        quoteAutoTextIndex=(int)productType.Profile;
                        chosenRider=productType.Profile;
                        break;

                    default:
                        chosenRider=productType.None;
                        break;
                }
            }

            //  Put in the quote
            OpenXmlElement qAutoText=AgreementQuote.RetrieveAutoText(templateName,SFProductTranslateToRider[quoteAutoTextIndex].AutoTextQuote); // 12.09.2020 Need to think about a better way of doing this because there's one autotext entry for the quote that's the same no matter how many riders there are.
            AgreementQuote.InsertAutoText(fileName,qAutoText);
            //  Add the custom xml or "Do the Merge"
            
            string linkID;
            cc=new contentControl(fileName);

            /* 1.20.2021 This can be removed because processing happens in the new quote object: SBComplexQuote/Instance: MultiYearSBQuote
            * but for now LEAVE this in until I am able to create a web api from python, pandas and flask!
            */
            xmlAgreementQuote.xpathContentControlName="LineItems";
            
            xmlMainContract.FileName="CM-Contract-"+xmlMainContract.XMLElementName.ToUpper()+"-";
            xmlMainContract.FileName+=PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
            xmlMainContract.SerializeDataToXml(cmNewContract.Agreement);
            xmlMainContract.InsertCustomXmlData(xmlMainContract.FileName, xmlMainContract.XMLNS, fileName, out linkID);
            
            cc.BindContentControls(xmlMainContract.FileName, fileName, xmlMainContract.XMLNS, xmlMainContract.XMLElementName, linkID, xmlAgreementQuote.xpathContentControlName);

            // Primary Contact
            xmlPrimaryContact.FileName="CM-Contract-"+xmlPrimaryContact.XMLElementName.ToUpper()+"-";
            xmlPrimaryContact.FileName+=PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
            xmlPrimaryContact.SerializeDataToXml(cmNewContract.PrimaryContact);
            xmlPrimaryContact.InsertCustomXmlData(xmlPrimaryContact.FileName, xmlPrimaryContact.XMLNS, fileName, out linkID);

            cc.BindContentControls(xmlPrimaryContact.FileName, fileName, xmlPrimaryContact.XMLNS, xmlPrimaryContact.XMLElementName, linkID, xmlAgreementQuote.xpathContentControlName);
            
            // Tech Support Contact
            xmlTechnicalSupport.FileName="CM-Contract-"+xmlPrimaryContact.XMLElementName.ToUpper()+"-";
            xmlTechnicalSupport.FileName+=TechnicalSupport.FirstName+"-"+TechnicalSupport.LastName+".xml";
            xmlTechnicalSupport.SerializeDataToXml(cmNewContract.TechnicalSupport);
            xmlTechnicalSupport.InsertCustomXmlData(xmlTechnicalSupport.FileName, xmlTechnicalSupport.XMLNS, fileName, out linkID);

            cc.BindContentControls(xmlTechnicalSupport.FileName, fileName, xmlTechnicalSupport.XMLNS, xmlTechnicalSupport.XMLElementName, linkID, xmlAgreementQuote.xpathContentControlName);
            
            // Invoice & Billing Contact
            xmlInvoiceBilling.FileName="CM-Contract-"+xmlInvoiceBilling.XMLElementName.ToUpper()+"-";
            xmlInvoiceBilling.FileName+=InvoiceBilling.FirstName+"-"+InvoiceBilling.LastName+".xml";
            xmlInvoiceBilling.SerializeDataToXml(cmNewContract.InvoiceBilling);
            xmlInvoiceBilling.InsertCustomXmlData(xmlInvoiceBilling.FileName, xmlInvoiceBilling.XMLNS, fileName, out linkID);

            cc.BindContentControls(xmlInvoiceBilling.FileName, fileName, xmlInvoiceBilling.XMLNS, xmlInvoiceBilling.XMLElementName, linkID, xmlAgreementQuote.xpathContentControlName);


            // 1.20.2021 Determine whether to create a STANDARD QUOTE or a MULTI-YEAR SPRINGBOARD QUOTE
            if (chosenRider == productType.SpringBoard && Agreement.Term > 1)
            {
                /* 1.20.2010 send JSON Quote from Salesforce to Python/Pandas restful web api */
            }

            // 1.19.2021 Should this be in the quote object? Should the quote inherit both AutoText AND CustomXml? Think about it.
            xmlAgreementQuote.FileName="CM-Contract-"+xmlAgreementQuote.XMLElementName.ToUpper()+"-";
            xmlAgreementQuote.FileName+=PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
            xmlAgreementQuote.SerializeDataToXml(cmNewContract.AgreementQuote);
            xmlAgreementQuote.InsertCustomXmlData(xmlAgreementQuote.FileName, xmlAgreementQuote.XMLNS, fileName, out linkID);
            cc.BindContentControls(xmlAgreementQuote.FileName, fileName, xmlAgreementQuote.XMLNS, xmlAgreementQuote.XMLElementName, linkID, xmlAgreementQuote.xpathContentControlName);
        }
    }
}