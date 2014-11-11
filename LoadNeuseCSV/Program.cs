using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace LoadNeuseCSV
{
    internal class SectionCsv
    {
        [CsvColumn(Name = "id", FieldIndex = 1)]
        public int id { get; set; }

        [CsvColumn(Name = "parent", FieldIndex = 2)]
        public int parent { get; set; }

        [CsvColumn(Name = "level", FieldIndex = 3)]
        public int level { get; set; }

        [CsvColumn(Name = "type", FieldIndex = 4)]
        public string type { get; set; }

        [CsvColumn(Name = "code", FieldIndex = 5)]
        public string code { get; set; }

        [CsvColumn(Name = "description", FieldIndex = 6)]
        public string description { get; set; }

        [CsvColumn(Name = "sort", FieldIndex = 7)]
        public string sort { get; set; }
    }

    public class CategoryID
    {
        [CsvColumn(Name = "Category", FieldIndex = 1)]
        public string Code { get; set; }
    }

    internal class Program
    {
        public static string NeuseCategory;
        public static string TemplateSection;


        private static void Main(string[] args)
        {
            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = true
            };

            var outputFileDescription = new CsvFileDescription();

            var cc = new CsvContext();
            var co = new CsvContext();

            IEnumerable<SectionCsv> sections =
                cc.Read<SectionCsv>("C:\\Users\\jlegacy\\Desktop\\BRScategoryDrive.csv", inputFileDescription);

// query data from sections

            IEnumerable<SectionCsv> sectionsByName =
                from p in sections
                where p.level.Equals(3)
                select p;

//keep track of progress//
            var xx = new List<CategoryID>();
            int count = 0;
            foreach (SectionCsv item in sectionsByName)
            {
                var cx = new CsvContext();
                IEnumerable<SectionCsv> allReadyProcessed =
                    cx.Read<SectionCsv>("C:\\Users\\jlegacy\\Desktop\\BRSprocessed.csv", inputFileDescription);

                try
                {
                    IEnumerable<SectionCsv> processedAlready =
                        from j in allReadyProcessed
                        where j.code.Equals(Asc(item.code))
                        select j;
                    if (processedAlready.Any())
                    {
                        continue;
                    }
                }
                catch (Exception e)
                {
                }

                var x = new CategoryID();
                x.Code = Convert.ToString(Asc(item.code));
                xx.Insert(count++, x);

                TemplateSection = Asc(item.code).ToString();
                NeuseCategory = item.id.ToString();
                // executes the appropriate commands to implement the changes to the database
            }

            co.Write(xx, "C:\\Users\\jlegacy\\Desktop\\BRSprocessed.csv", outputFileDescription);

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = @"C:\Users\jlegacy\Documents\Visual Studio 2013\Projects\LoadNeuseProductsCSV\LoadNeuseCSV\bin\Debug\LoadNeuseCSV.exe";
           startInfo.Arguments =  TemplateSection + " " + NeuseCategory;
            Process.Start(startInfo);

        }

        private static int Asc(String ch)
        {
            //Return the character value of the given character
            ch = ch.ToUpper();
            byte[] bytes = Encoding.ASCII.GetBytes(ch);
            string numString = "";
            int j;
            foreach (byte b in bytes)
            {
                if (b >= 48 && b <= 57)
                {
                    numString = numString + (Convert.ToChar(b));
                }
                else
                {
                    j = b - 64;
                    numString = numString + (Convert.ToString(j));
                }
            }

            return Convert.ToInt32(numString);
        }
    }
}