using CsvHelper;
using CsvHelper.Configuration;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GMail_Data_Extractor
{
    class Program
    {

        static void Main(string[] args)
        {

            Console.WriteLine("Provide path to the files folder. Or <enter> to use defaults.");
            var path = Console.ReadLine();

            if (path.Length != 0)
            {
                if (!File.Exists(path))
                {
                    Console.WriteLine("Invalid path provided. Press any key to continue...");
                    Console.Read();
                }
            }
            else
            {
                path = "C:/MailFiles";
            }

            var files = from file in Directory.EnumerateFiles(path, "*.mbox")
                        select file;

            Console.WriteLine("{0} files found to process", files.Count().ToString());
            foreach (var f in files)
            {
                Console.WriteLine("Processing file {0}", f);
                List<MyMail> output = new List<MyMail>();

                // Load every message from a Unix mbox
                using (FileStream stream = new FileStream(f, FileMode.Open))
                {
                    var parser = new MimeParser(stream, MimeFormat.Mbox);

                    while (!parser.IsEndOfStream)
                    {
                        var message = parser.ParseMessage();
                        output.Add(new MyMail(message.From.ToString(),
                            message.Date.ToString(),
                            message.ReplyTo.ToString(),
                            message.Subject,
                            message.To.ToString(),
                            MakeBody(message.TextBody, path + "/", false),
                            MakeBody(message.HtmlBody, path + "/", true)));
                        // do something with the message
                    }
                }

                using (System.IO.TextWriter writer = File.CreateText(f + ".csv"))
                {
                    var csv = new CsvWriter(writer, new CsvConfiguration() { QuoteAllFields = true, HasExcelSeparator = true });
                    csv.WriteRecords(output);

                }

            }




            Console.Read();

        }

        private static string MakeBody(string htmlBody, string path, bool html)
        {
            if (string.IsNullOrEmpty(htmlBody)) return "";
            var guid = Guid.NewGuid() + (html ? ".html" : ".txt");
            File.WriteAllText(path + guid, htmlBody);
            return string.Format("=HIPERVINCULO(\"{0}\")", guid);
        }

        private static string GetText(string label, string line)
        {
            return line.Substring(label.Length);
        }
    }
}
