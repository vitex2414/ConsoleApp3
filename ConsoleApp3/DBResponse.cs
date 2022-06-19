using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp3
{
    class DBResponse
    {
        public string code { get; set; }
        public string result { get; set; }
        public DBResponse(string code, string result)
        {
            this.code = code;
            this.result = result;
        }
        public DBResponse()
        {
        }
    }
}
