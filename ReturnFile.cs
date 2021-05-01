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
            var fileName = @"C:\Users\Adria\Documents\Love.txt";
            using(Stream fileContent = File.OpenRead(fileName)){
                fileContent.Seek(0,SeekOrigin.Begin);
                context.Response.ContentLength=fileContent.Length;
            };
            await context.Response.SendFileAsync(fileName);
        }
    }
}