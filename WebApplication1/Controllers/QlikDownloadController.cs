using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Util;
using QlikService.Models;

namespace WebApplication1.Controllers
{
    public class DownloadController : ApiController
    {
        // GET api/download
        public string Get()
        {
            string q = Request.RequestUri.Query;
            String r= QlikService.Models.excelDownload.createExcel("Consumer Sales","testt");
            return r;
          
        }

        // GET api/download/app/excel/object
        [Route("api/download/{app}/{type}/{objectid}")]
        public string Get(string app,string type,string objectid)
        {
            String r = QlikService.Models.excelDownload.createExcel( app, objectid);
            return r;

        }

        /* public string Post([FromBody]dynamic hyperCube)
         {
             Console.WriteLine(hyperCube);

             return "Post Request received as string";
         }

         public string Post([FromBody]string hyperCube)
         {
             Console.WriteLine(hyperCube);

             return "Post Request received as string";
         }*/

        // POST api/values
        public string Post([FromBody] string p)
        {
            //string result = await Request.Content.ReadAsStringAsync();            

            

            return "Post Request received";
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
