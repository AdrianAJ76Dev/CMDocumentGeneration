/*  04.17.2021 New version of Azure Resources that uses Azure Managed Identity Code
*   Uses appsettings
*/
using System;
using System.Text;
using System.IO;
using System.Threading.Tasks;

using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Specialized;
using Azure.Storage.Blobs.Models;
using Microsoft.Extensions.Configuration;

using Azure.Identity;

namespace CMDocumentGeneration.Models
{
    static public class AzureResourcesConfig{
        private const string _accountName="stcbfiles";
        private const string _containerWordTemplate="cb-template";
        private const string _containerCustomXMLJSON="cb-data";
        private const string _containerGeneratedDocuments="cb-documents-generated";


        public const string azureResourcesSection = "AzureResourcesConfig";
        static public string AccountName{
            get{return _accountName;}
            }
        static public string ContainerNameWordTemplates{
            get{return _containerWordTemplate;}
            }
        static public string ContainerNameCustomXMLJSONData{
            get{return _containerCustomXMLJSON;}
            }
        static public string ContainerNameGeneratedWordDocuments{
            get{return _containerGeneratedDocuments;}
            }
    }

    static public class AzureResources
    {
        static public string ContractFileName;

        static public MemoryStream GetWordTemplate(string templateName){
            BlobClient wordTemplate = WordResourceInAzure(AzureResourcesConfig.ContainerNameWordTemplates,templateName);
            MemoryStream msWordTemplate = new MemoryStream();
            wordTemplate.DownloadTo(msWordTemplate);
            return msWordTemplate;
        }
        static public MemoryStream GetGeneratedDocument(string filename){
            BlobClient wordDoc = WordResourceInAzure(AzureResourcesConfig.ContainerNameGeneratedWordDocuments,filename);
            MemoryStream msWordDoc = new MemoryStream();
            wordDoc.DownloadTo(msWordDoc);
            return msWordDoc;
        }

        static public void SaveGeneratedDocument(MemoryStream msdoc, string fileName){
            BlobClient wrdDocument = WordResourceInAzure(AzureResourcesConfig.ContainerNameGeneratedWordDocuments,fileName);
           /*   Learned last week that this code HANGS! if the line below is commented out. 
           *    Probably looking at the END of the memory stream and there's nothing there
           *    to upload!
           */
            msdoc.Seek(0, SeekOrigin.Begin);
            wrdDocument.Upload(msdoc,true);
            ContractFileName = fileName;
        }

        static public MemoryStream GetCustomXmlFile(string fileName){
            BlobClient wordCustomXml = WordResourceInAzure(AzureResourcesConfig.ContainerNameCustomXMLJSONData,fileName);
            MemoryStream msWordCustomXml = new MemoryStream();
            wordCustomXml.DownloadTo(msWordCustomXml);
            return msWordCustomXml;
        }
        static public void SaveCustomXmlFile(MemoryStream msWrdCustomXml, string fileName){
            BlobContainerClient containerClientDoc=null;
            BlobClient wrdDocument = WordResourceInAzure(AzureResourcesConfig.ContainerNameCustomXMLJSONData,fileName);

           //******************************************************************************
           /*   Learned last week that this code HANGS! if the line below is commented out. 
           *    Probably looking at the END of the memory stream and there's nothing there
           *    to upload!
           */
            msWrdCustomXml.Seek(0, SeekOrigin.Begin);
            //*****************************************************************************

            containerClientDoc = wrdDocument.GetParentBlobContainerClient();
            containerClientDoc.DeleteBlobIfExists(fileName);
            containerClientDoc.UploadBlob(fileName,msWrdCustomXml);
            msWrdCustomXml.Close();

            //xmlFile.Upload(msWrdCustomXml,true);
        }        

        static public MemoryStream GetJSONFile(string fileName){
            BlobClient jsonAutoTextName = WordResourceInAzure(AzureResourcesConfig.ContainerNameCustomXMLJSONData,fileName);

            MemoryStream msAutoTextNames = new MemoryStream();
            jsonAutoTextName.DownloadTo(msAutoTextNames);
            return msAutoTextNames;
        }

        static private BlobClient WordResourceInAzure(string containerName, string wrdDocResourceName){
            BlobClient wordResource=null;
            string containerEndpoint = string.Format("https://{0}.blob.core.windows.net/{1}",
                                                        AzureResourcesConfig.AccountName,
                                                        containerName);

            // Get a credential and create a client object for the blob container.
            BlobContainerClient containerClient = new BlobContainerClient(new Uri(containerEndpoint), new DefaultAzureCredential());

            //  Retrieves necessary document
            wordResource = containerClient.GetBlobClient(wrdDocResourceName);

            return wordResource;
        }
    }    
}    