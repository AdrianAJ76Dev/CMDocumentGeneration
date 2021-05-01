using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
namespace CMDocumentGeneration
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        /*
        *   04.29.2021 - Services should be added to save the newly generated contract, for example, to either:
        *   1. Azure
        *   2. Amazon Web Services
        *   3. Google Cloud Services
        *   4. SharePoint
        *   5. Local machine/laptop/PC/Mac
        *   Possible figure out how to make AzureResources into a Service!
        */
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            /*  04.29.2021 - Read appsettings.json section: SettingsCMContract:WordTemplates:Name
            *   This MUST come after routing middleware BECAUSE I'm using routing to retrieve the value
            */
            //  04.30.2021 - THIS WORKS!! IT RETURNS THE CONTENTS OF THE FILE!!!!
            app.UseMiddleware<ReturnFile>();

            app.UseHttpsRedirection();

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
