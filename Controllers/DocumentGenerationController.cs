using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using CMDocumentGeneration.Models;

namespace CMDocumentGeneration.Controllers
{
    [ApiController]
    [Route("DocumentGeneration")]
    public class DocumentGenerationController : ControllerBase
    {
        /*
        [HttpPost("SoleSourceLetter")]
        public IActionResult CMSoleSourceLetter([FromBody] CMSoleSourceLetter soleSourceLetter, CMSoleSourceLetter.Contact primaryContact){
            soleSourceLetter.Generate(primaryContact);
            // Return a Word document
            return new FileStreamResult(AzureResources.GetGeneratedDocument(soleSourceLetter.FileName),"application/vnd.openxmlformats-officedocument.wordprocessingml.document"); 
        }
        */
        
        [HttpPost("Agreement")]
        public void CMContract([FromBody] CMContract cmContract, [FromHeader] string ContractTemplate){
            cmContract.ContractTemplate=ContractTemplate;
            cmContract.Generate(cmContract);
        }
        
        [HttpPost("Agreement\\SpringBoard\\Quote")]
        public void ComplexQuote([FromBody] CMContract.SBComplexQuote cmQuote){
            cmQuote.AddProcessedQuote();
        }
    }
}