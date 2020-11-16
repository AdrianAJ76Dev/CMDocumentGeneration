using System;
using System.Text;
using System.IO;
using System.Threading.Tasks;

using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;

namespace CMDocumentGeneration.Models
{
    public static class AzureResources
    {
        /*  08.15.2020
        *   Originally called class "whereItsAt" after Beck song. :-)
        *   I want to store all of the information Azure needs here:
        *   Resource group names
        *   Storage account names
        *   container names
        *   File names
        *   etc.
        */

        private static string cn = string.Empty;
        // Seek out alternatives to hard coding these values
        // AzureContainer names
        private const string CONTAINER_CUSTOM_XML_DATA = "cb-data"; // Container name for my xml files
        private const string CONTAINER_GENERATED_DOCUMENTS = "cb-documents-generated"; // Container name for the final documents generated
        private const string CONTAINER_WORD_TEMPLATES ="cb-template"; // Container name for my Word Templates
        
        // Word Files: Template and testing document
        private const string DOCUMENT_WORD_TEST_NAME = "Test-Doc.docx";
        private const string FILE_CUSTOM_XML_DATA = "cbssl.xml";
        private const string FILE_JSON_DATA = "AutoText.JSON";

        public static MemoryStream GetWordTemplate(string templateName){
        /*  08.15.2020
        *   Put all of the Azure code here and just GET the Word Template
        *   It's ok to be specific now until I learn how to further 
        *   abstract this code
        *   08.30.2020
        *   Passing in the template name and changing the method name to GetWordTemplate
        */
            cn = Environment.GetEnvironmentVariable("AZURE_STORAGE_CONNECTION_STRING");

            BlobContainerClient containerClientTemplate = new BlobContainerClient(cn,CONTAINER_WORD_TEMPLATES);
            BlobClient wordTemplate = containerClientTemplate.GetBlobClient(templateName);
            
            MemoryStream msWordTemplate = new MemoryStream();
            wordTemplate.DownloadTo(msWordTemplate);
            return msWordTemplate;
        }
        public static MemoryStream GetGeneratedDocument(string filename){

            cn = Environment.GetEnvironmentVariable("AZURE_STORAGE_CONNECTION_STRING");

            BlobContainerClient containerClientWordDocument = new BlobContainerClient(cn,CONTAINER_GENERATED_DOCUMENTS);
            BlobClient wordDoc = containerClientWordDocument.GetBlobClient(filename);
            
            MemoryStream msWordDoc = new MemoryStream();
            wordDoc.DownloadTo(msWordDoc);
            return msWordDoc;
        }

        public static MemoryStream GetCustomXmlFile(string filename)
        {
            /*  08.15.2020
            *   Put all of the Azure code here and just GET the Custom XML file
            *   It's ok to be specific now until I learn how to further 
            *   abstract this code
            */
            cn = Environment.GetEnvironmentVariable("AZURE_STORAGE_CONNECTION_STRING");

            BlobContainerClient containerClientTemplate = new BlobContainerClient(cn,CONTAINER_CUSTOM_XML_DATA);
            BlobClient wordCustomXml = containerClientTemplate.GetBlobClient(filename);

            MemoryStream msWordCustomXml = new MemoryStream();
            wordCustomXml.DownloadTo(msWordCustomXml);
            return msWordCustomXml;
        }

        public static MemoryStream GetJSONFile(string filename){
            /*  10.27.2020
            *   Salesforce exports data in JSON BUT it does not export the autoext name
            *   ONLY the Product Name. So Product Name needs to be matched up to 
            *   the correct autotext name
            */

            cn = Environment.GetEnvironmentVariable("AZURE_STORAGE_CONNECTION_STRING");

            BlobContainerClient containerClientData = new BlobContainerClient(cn,CONTAINER_CUSTOM_XML_DATA);
            BlobClient jsonAutoTextName = containerClientData.GetBlobClient(filename);
            
            MemoryStream msAutoTextNames = new MemoryStream();
            jsonAutoTextName.DownloadTo(msAutoTextNames);
            return msAutoTextNames;
        }

        public static void SaveGeneratedDocument(MemoryStream msdoc, string fileName){
            /*  08.15.2020
            *   The Azure code should be responsible for saving the document to Azure
            *   Trying to adhere to separation of concerns in this demo
            *   May separate futher down the road.
            */
            cn = Environment.GetEnvironmentVariable("AZURE_STORAGE_CONNECTION_STRING");

            BlobContainerClient containerClientDoc = new BlobContainerClient(cn,CONTAINER_GENERATED_DOCUMENTS);
            BlobClient wrdDocument = containerClientDoc.GetBlobClient(fileName);

           /*   Learned last week that this code HANGS! if the line below is commented out. 
           *    Probably looking at the END of the memory stream and there's nothing there
           *    to upload!
           */
            msdoc.Seek(0, SeekOrigin.Begin);
            wrdDocument.Upload(msdoc,true);
        }

        public static void SaveCustomXmlFile(MemoryStream msWrdCustomXml, string fileName){
            cn = Environment.GetEnvironmentVariable("AZURE_STORAGE_CONNECTION_STRING");

            BlobContainerClient containerClientDoc = new BlobContainerClient(cn,CONTAINER_CUSTOM_XML_DATA);
            BlobClient xmlFile = containerClientDoc.GetBlobClient(fileName);

           //******************************************************************************
           /*   Learned last week that this code HANGS! if the line below is commented out. 
           *    Probably looking at the END of the memory stream and there's nothing there
           *    to upload!
           */
            msWrdCustomXml.Seek(0, SeekOrigin.Begin);
            //*****************************************************************************

            containerClientDoc.DeleteBlobIfExists(fileName);
            containerClientDoc.UploadBlob(fileName,msWrdCustomXml);
            msWrdCustomXml.Close();
            //xmlFile.Upload(msWrdCustomXml,true);
        }
    }
}
