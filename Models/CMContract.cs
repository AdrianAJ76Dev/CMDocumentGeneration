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
    *   
    *   05.15.2021
    *   Possibly read data from header, afterall they too are Key-Values
    *   just like the JSON in the body, AND use THOSE values to determine
    *   WHAT template to use and HOW to use it!
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
            Nexus6HSEnglishTeacher=12,
            Nexus6HSGymTeacher=13,
            Nexus6HSMathTeacher=14,
            Nexus6SDProfessor=15,
            Nexus6SDTeachingAssistant=16,
            Nexus6ResidentialAssistant=17,
            Nexus6TeachingAssistant=18,
            Nexus8Professor=19,
            Nexus8TeachingAssistant=20,
            Nexus9Professor=21,
            Nexus9TeachingAssistantBeta=22,
            None=99
        };
        
        /* 05.05.2021 - Revising this to read a JSON file (or maybe later database) and get the active templates 
        *  I would love to be able to add/remove templates without recompiling this code
        */
        public enum instType:byte
        {
            K12=0,
            HED=1
        };


        public class WordTemplate
        {
            // *.dotx file used to create the document
            public string TemplateName{get;set;}
            public bool HasRiders{get;set;}
            public bool HasQuote{get;set;}
        }

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
            // 05.05.2021 Added to retrieve the associated Word Template
            public string PAbbrv{get;set;}
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
        
        /* 05.06.2021 - Salesforce user now selects the template to use. (template name exported in JSON)
        * This is done currently in the old system via button click on the contract record 
        */
        public string ContractTemplate{get;set;}

        /* 05.16.2021 Word template object that determines how the final contract is constructed */
        public List<WordTemplate> ActiveWordTemplates{get;set;}
        public WordTemplate WordTemplateSelected {get;set;}

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
            *   09.27.2020
            */

            /* 05.16.2021 Retrieve the selected Word Template (comes from the request Header) */
            const string FILE_NAME_TEMPLATES = "Templates-Contract.JSON";
            MemoryStream msTemplates=AzureResources.GetJSONFile(FILE_NAME_TEMPLATES);
            ReadOnlySpan<byte> jsonReadOnlySpanTemplates = msTemplates.ToArray();
            ActiveWordTemplates = JsonSerializer.Deserialize<List<WordTemplate>>(jsonReadOnlySpanTemplates);
            msTemplates.Close();

            MemoryStream msAutoTextNamesJSON = AzureResources.GetJSONFile("AutoText.JSON");
            ReadOnlySpan<byte> jsonReadOnlySpanRiders = msAutoTextNamesJSON.ToArray();
            // 05.18.2021 Consider changing the name of the type autoTextSettings and the name of SFProductTranslateToRider
            SFProductTranslateToRider=JsonSerializer.Deserialize<List<autoTextSettings>>(jsonReadOnlySpanRiders);
            msAutoTextNamesJSON.Close();
            string productsInFileName = string.Empty;

            /* Get the selected template from the list */
            var queryTemplates = 
                    from selTemplate in ActiveWordTemplates 
                    where selTemplate.TemplateName == ContractTemplate
                    select selTemplate;

            WordTemplateSelected = queryTemplates.FirstOrDefault();

            /* 05.18.2021 Start with the arrays from the JSON in the request
            *  and the JSON defined for this Web api */
            var SFProducts = (from ar in AgreementRiders
                select ar.ProductName);
            
            var queryProducts = (from pt in SFProductTranslateToRider
                where SFProducts.Contains(pt.Product)
                select pt);

            foreach (var productAbbreviation in queryProducts)
            {
                productsInFileName+=productAbbreviation.PAbbrv;
                productsInFileName+="-";
            }

            /* 05.19.2021 Think about creating a Custom Field on the ScratchOrgInfo object that handles Contract #s*/
            productsInFileName=productsInFileName.Remove(productsInFileName.Length-1,1);
            fileName="CM-Contract-"
                +PrimaryContact.FirstName+"-"
                +PrimaryContact.LastName+"-"
                +productsInFileName
                +Agreement.ContractNumber
                +".docx"; // Name of Word document

            // 05.06.2021 - Getting Word template from JSON now to facialitate adding future templates to the systeme
            CreateWordDocument(ContractTemplate,fileName);

            OpenXmlElement AutoText;
            productType chosenRider=productType.EnrollmentPlanningServiceUnlimited;

            // 05.17.2021 Not all contracts have riders
            string linkID;
            cc=new contentControl(fileName);

            if (WordTemplateSelected.HasRiders)
            {
                /*  Choose the rider
                *  05.05.2021 AND the template
                *  Really think of changing the name from AutoText.JSON to something else that would include the template value
                *  How do I determine if there are no riders? Maybe I should think about how I read the JSON (maybe no riders there)
                */

                /* 05.10.2021 Not all contracts will have AutoText 
                *  This is the point where I check if the AutoText JSON
                *  contained a member defining a rider and/or a memeber
                *  defining a quote. So and/or rider & quote can be either
                *  or one, nil AutoTextRider=null/AutoTextQuote=null
                */

                /*
                List<autoTextSettings> AutoTextFound = new List<autoTextSettings>();
                foreach (var currRider in AgreementRiders)
                {
                    foreach (var autoTextName in SFProductTranslateToRider)
                    {
                        /* 05.07.2021 Seriously think of redoing this code because it
                        * depends on there being an EXACT match in the AutoText.JSON
                        * The riders didn't get inserted because of ™ v. (TM) 
                        * NOTE: Only the AutoText.JSON needs to be corrected to match
                        * whatever comes out of salesforce
                        */
                        /*
                        if (currRider.ProductName.Contains(autoTextName.Product))
                        {
                            AutoTextFound.Add(autoTextName);
                            productsInFileName+=autoTextName.PAbbrv;
                            productsInFileName+="-";
                        }
                    }
                }
                */

                // 11.06.2020 Open the generated Word document once and insert all that needs to be inserted!
                // 11.06.2020 Insert Riders by name sent from Salesforce using product names
                // 05.18.2021 Consider using LINQ here
                foreach (var currRider in queryProducts)
                {
                    switch (currRider.contractRiderID)
                    {
                        case productType.SpringBoard:
                            Rider SpringBoard = new Rider();
                            AutoText = SpringBoard.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            SpringBoard.InsertAutoText(fileName,AutoText);
                            chosenRider = productType.SpringBoard;
                            break;

                        case productType.PreAP:
                            Rider PreAP = new Rider();
                            AutoText = PreAP.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            PreAP.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.PreAP;
                            break;

                        case productType.EnrollmentPlanningServiceUnlimited:
                            Rider EPSUnlimited = new Rider();
                            AutoText = EPSUnlimited.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            EPSUnlimited.InsertAutoText(fileName, AutoText);
                            chosenRider=productType.EnrollmentPlanningServiceUnlimited;
                            break;

                        case productType.StudentSearchService:
                            Rider SSSUnlimited = new Rider();
                            AutoText = SSSUnlimited.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            SSSUnlimited.InsertAutoText(fileName, AutoText);
                            chosenRider=productType.StudentSearchService;
                            break;

                        case productType.SegmentAnalysisService:
                            Rider SASUnlimited = new Rider();
                            AutoText = SASUnlimited.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            SASUnlimited.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.PowerFAIDSInitial:
                            chosenRider=productType.PowerFAIDSInitial;
                            break;

                        case productType.Profile:
                            chosenRider=productType.Profile;
                            break;

                        case productType.Nexus6HSEnglishTeacher:
                            Rider Nexus6HSEnglishTeacher = new Rider();
                            AutoText = Nexus6HSEnglishTeacher.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus6HSEnglishTeacher.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus6HSGymTeacher:
                            Rider Nexus6HSGymTeacher = new Rider();
                            AutoText = Nexus6HSGymTeacher.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus6HSGymTeacher.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus6HSMathTeacher:
                            Rider Nexus6HSMathTeacher = new Rider();
                            AutoText = Nexus6HSMathTeacher.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus6HSMathTeacher.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus6ResidentialAssistant:
                            Rider Nexus6ResidentialAssistant = new Rider();
                            AutoText = Nexus6ResidentialAssistant.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus6ResidentialAssistant.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus6SDProfessor:
                            Rider Nexus6SDProfessor = new Rider();
                            AutoText = Nexus6SDProfessor.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus6SDProfessor.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus6SDTeachingAssistant:
                            Rider Nexus6SDTeachingAssistant = new Rider();
                            AutoText = Nexus6SDTeachingAssistant.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus6SDTeachingAssistant.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus6TeachingAssistant:
                            Rider Nexus6TeachingAssistant = new Rider();
                            AutoText = Nexus6TeachingAssistant.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus6TeachingAssistant.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus8Professor:
                            Rider Nexus8Professor = new Rider();
                            AutoText = Nexus8Professor.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus8Professor.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus8TeachingAssistant:
                            Rider Nexus8TeachingAssistant = new Rider();
                            AutoText = Nexus8TeachingAssistant.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus8TeachingAssistant.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus9Professor:
                            Rider Nexus9Professor = new Rider();
                            AutoText = Nexus9Professor.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus9Professor.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        case productType.Nexus9TeachingAssistantBeta:
                            Rider Nexus9TeachingAssistantBeta = new Rider();
                            AutoText = Nexus9TeachingAssistantBeta.RetrieveAutoText(ContractTemplate,currRider.AutoTextRider);
                            Nexus9TeachingAssistantBeta.InsertAutoText(fileName,AutoText);
                            chosenRider=productType.SegmentAnalysisService;
                            break;

                        default:
                            chosenRider=productType.None;
                            break;
                    }
                }
            }

            if (WordTemplateSelected.HasQuote)
            {
                //  Put in the quote
                // 05.10.2021 There may not be a quote or, in general, how to gracefully fail when no autotext is found
                OpenXmlElement qAutoText=AgreementQuote.RetrieveAutoText(ContractTemplate,SFProductTranslateToRider[(int)chosenRider].AutoTextQuote); // 12.09.2020 Need to think about a better way of doing this because there's one autotext entry for the quote that's the same no matter how many riders there are.
                AgreementQuote.InsertAutoText(fileName,qAutoText);
                //  Add the custom xml or "Do the Merge"
                // 1.19.2021 Should this be in the quote object? Should the quote inherit both AutoText AND CustomXml? Think about it.
                /* 1.20.2021 This can be removed because processing happens in the new quote object: SBComplexQuote/Instance: MultiYearSBQuote
                * but for now LEAVE this in until I am able to create a web api from python, pandas and flask!
                */
                xmlAgreementQuote.xpathContentControlName="LineItems";
                xmlAgreementQuote.FileName="CM-Contract-"+xmlAgreementQuote.XMLElementName.ToUpper()+"-";
                xmlAgreementQuote.FileName+=PrimaryContact.FirstName+"-"+PrimaryContact.LastName+".xml";
                xmlAgreementQuote.SerializeDataToXml(cmNewContract.AgreementQuote);
                xmlAgreementQuote.InsertCustomXmlData(xmlAgreementQuote.FileName, xmlAgreementQuote.XMLNS, fileName, out linkID);
                cc.BindContentControls(xmlAgreementQuote.FileName, fileName, xmlAgreementQuote.XMLNS, xmlAgreementQuote.XMLElementName, linkID, xmlAgreementQuote.xpathContentControlName);
            }
            // 05.17.2021 Not all contracts have quotes
                
            // CHANGE - 04.23.2021 - The following code can be put into a single function with the xmlName, linkID, xmlAgreementQuote as parameters - CHANGE!
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
            xmlTechnicalSupport.FileName="CM-Contract-"+xmlTechnicalSupport.XMLElementName.ToUpper()+"-";
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
        }
    }
}