using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;

namespace CMDocumentGeneration.Services
{
    public interface ICloudStorage{
        Task<MemoryStream> GetWordTemplate(string templateName);
        Task<MemoryStream> GetGeneratedDocument(string filename);
        Task<MemoryStream> GetCustomXmlFile(string fileName);
        Task<MemoryStream> GetJSONFile(string fileName);
        Task SaveGeneratedDocument(MemoryStream msdoc, string fileName);
        Task SaveCustomXmlFile(MemoryStream msWrdCustomXml, string fileName);
    }
}