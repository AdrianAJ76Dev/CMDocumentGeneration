using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Http;
using CMDocumentGeneration.Models;

namespace CMDocumentGeneration
{
    public class ReturnFile
    {
        private RequestDelegate next;
        public ReturnFile(RequestDelegate nextMiddleWare){
            next = nextMiddleWare;
        }

        public async Task Invoke(HttpContext context){
            await next(context);

            if (AzureResources.ContractFileName != null)
            {
                // 05.03.2021 I need to add the name of the contract to the header here.
                context.Response.Headers.Add("ContractName", AzureResources.ContractFileName);

                // MemoryStream file = AzureResources.GetCustomXmlFile("Love.txt");
                MemoryStream file = AzureResources.GetGeneratedDocument(AzureResources.ContractFileName);
                file.Seek(0,SeekOrigin.Begin);

                context.Response.ContentLength=file.Length;
                await file.CopyToAsync(context.Response.Body);                              
            }
        }
    }
}