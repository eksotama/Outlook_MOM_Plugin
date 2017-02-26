using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace CalScanner
{
    using System;
    using System.IO;

    public class HtmlTemplate
    {
        private string _html;

        public HtmlTemplate()
        {  
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "MOMOutlookAddIn.Resources.template.html";

            
           // var temp = MOMOutlookAddIn.Properties.Resources.template;
           // Stream stream = new MemoryStream(temp);
            //Stream str=new MemoryStream()
            //using (var reader = new StreamReader(new Stream(MOMOutlookAddIn.Properties.Resources.template))
          using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            //using (Stream stream = new MemoryStream(temp))
          using (StreamReader reader = new StreamReader(stream))
          {
              _html = reader.ReadToEnd();
          }
                
        }

        public string Render(object values)
        {
            string output = _html;
            foreach (var p in values.GetType().GetProperties())
                output = output.Replace("[" + p.Name + "]", (p.GetValue(values, null) as string).Trim() ?? string.Empty);
            return output;
        }
    }

   /* public class Program
    {
        void Main()
        {
            var template = new HtmlTemplate(@"C:\template.html");
            var output = template.Render(new
            {
                TITLE = "My Web Page",
                METAKEYWORDS = "Keyword1, Keyword2, Keyword3",
                BODY = "Body content goes here",
                ETC = "etc"
            });
            Console.WriteLine(output);
        }
    }*/
}
