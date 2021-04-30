using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.FileProviders;
using System.Threading.Tasks;
using System.IO;
using CMDocumentGeneration.Models;

namespace CMDocumentGeneration
{
    public class WebAPIResponseMiddleware
    {
        private RequestDelegate next;
        public WebAPIResponseMiddleware(RequestDelegate nextDelegate)
        {
            next = nextDelegate;
        }

        public async Task Invoke(HttpContext context)
        {
            await context.Response.WriteAsync("\nClass Middleware!\nContract Created and HEELLLLOOOOO NURSE!!!!");
            await next(context);
        }
    }

    public class ReturnWordDoc
    {
        private RequestDelegate next;
        public ReturnWordDoc(RequestDelegate nextDelegate)
        {
            next = nextDelegate;
        }

        public async Task Invoke(HttpContext context)
        {
            await next(context);
            /*
            *   MemoryStream newContract = AzureResources.GetGeneratedDocument(fileName);
            *   context.Response.ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            *   This WILL NOT WORK!
            *   context.Response.SendFileAsync(fileName)
           */
           
        }
    }    
}
