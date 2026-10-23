using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

namespace OneDriveASPCoreFileProvider.Controllers
{
    [Route("api/[controller]")]
    public class TestController : Controller
    {
        // GET: api/test
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET: api/test/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST: api/test
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT: api/test/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE: api/test/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
