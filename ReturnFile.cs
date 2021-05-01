using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;
using System.IO;
using System.Text;
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
            MemoryStream file = AzureResources.GetCustomXmlFile("Love.txt");
            file.Seek(0,SeekOrigin.Begin);
            context.Response.ContentLength=file.Length;
            await file.CopyToAsync(context.Response.Body); 
        }
    }
}