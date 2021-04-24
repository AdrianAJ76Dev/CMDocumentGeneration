using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;

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
}
