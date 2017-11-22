using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace SFMImporter
{
    public static class Utils
    {
        public static IEnumerable<string> GetFilesWithExtension(string dirname, string extension)
        {
            string[] files = Directory.GetFiles(dirname);

            foreach (string file in files)
            {
                if (file.EndsWith(extension))
                {
                    yield return file;
                }
            }
        }

        public static IEnumerable<ContractorLicense> GetContractorLicenses(IEnumerable<string> files)
        {
            ContractorLicense cl;
            foreach (string file in files)
            {
                cl = new ContractorLicense();
                using (XmlReader xr = XmlReader.Create(file))
                {
                    xr.ReadToFollowing("Document");
                    cl.DocID = xr.GetAttribute("id");
                    cl.Domain = xr.GetAttribute("domain");
                    cl.TypeName = xr.GetAttribute("typeName");
                    cl.Name = xr.GetAttribute("name");
                    cl.IndexQueue = xr.GetAttribute("indexQueue");
                    cl.Index = xr.GetAttribute("index");
                    cl.ConvertFile = xr.GetAttribute("convertFile");
                    cl.ConvertFormat = xr.GetAttribute("converFormat");
                    cl.AutoIndex = xr.GetAttribute("autoIndex");
                    cl.Ocr = xr.GetAttribute("ocr");
                    xr.ReadToFollowing("Field");
                    //Console.WriteLine("Is the element Field? {0}",xr.IsStartElement("Field"));
                    while (xr.IsStartElement("Field"))
                    {
                        string n = xr.GetAttribute("name");

                        switch (n)
                        {
                            case "EMPLOYEE_FIRST_NAME":
                                cl.EmpFirstName = xr.GetAttribute("value");
                                xr.ReadToNextSibling("Field");
                                break;
                            case "EMPLOYEE_LAST_NAME":
                                cl.EmpLastName = xr.GetAttribute("value");
                                xr.ReadToNextSibling("Field");
                                break;
                            case "ICC_Number":
                                cl.IccNumber = xr.GetAttribute("value");
                                xr.ReadToNextSibling("Field");
                                break;
                            case "CONTRACTOR_NAME":
                                cl.ContractorName = xr.GetAttribute("value");
                                xr.ReadToNextSibling("Field");
                                break;
                            case "CONTRACTOR_IL__":
                                cl.ContractorIl = xr.GetAttribute("value");
                                xr.ReadToNextSibling("Field");
                                break;
                            case "DOC_DATE":
                                cl.DocDate = xr.GetAttribute("value");
                                xr.ReadToNextSibling("File");
                                break;
                        }
                    }
                    cl.mimeType = xr.GetAttribute("mimeType");
                    cl.FileName = xr.GetAttribute("name");
                    cl.FilePath = xr.GetAttribute("filePath");
                }
                yield return cl;
            }
        }

        public static string CreateClXlsx(IEnumerable<ContractorLicense> licenses, string filename)
        {
            //Setup the workbook
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            if (xlApp == null)
            {
                return "Excel is not properly installed!!";
            }

            //Write the header row using the public attribute names
            PropertyInfo[] properties = typeof(ContractorLicense).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            int col = properties.Length;
            int r = 1;
            int c = 1;
            foreach (PropertyInfo p in properties)
            {
                xlWorkSheet.Cells[r, c] = p.Name;
                c++;
            }

            r++;
            c = 1;
       
            foreach(ContractorLicense cl in licenses)
            {
                foreach (PropertyInfo p in properties)
                {
                    xlWorkSheet.Cells[r, c] = (p.GetValue(cl));
                    c++;
                }
                c = 1;
                r++;
            }

            xlWorkBook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            return "It worked";
        }
    }
}
