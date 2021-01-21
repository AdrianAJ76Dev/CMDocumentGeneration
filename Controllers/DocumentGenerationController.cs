using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
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
        public void CMContract([FromBody] CMContract cmContract){
            cmContract.Generate(cmContract);
            // Return a Word document
            //return new FileStreamResult(AzureResources.GetGeneratedDocument(cmContract.FileName),"application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }
        [HttpPost("Agreement\\SpringBoard\\Quote")]
        public void ComplexQuote([FromBody] CMContract.SBComplexQuote cmQuote){
            cmQuote.AddProcessedQuote();
        }

    }

}