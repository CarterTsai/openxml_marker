using System;
using openxml.Services;

namespace openxml
{
    class Program
    {
        static void Main(string[] args)
        {
            var service = new ExcelService();
            service.Create();
        }
    }
}
