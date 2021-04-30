using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;
using System.IO;
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
            string file = @"C:\Users\Adria\Documents\Love.txt";
            context.Response.ContentType="text/plain";
            await context.Response.SendFileAsync(file);
        }
    }
}