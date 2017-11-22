using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SFMImporter
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static string filename = "c:\\temp\\ContractorLicense.xls";
        static void Main(string[] args)
        {
            log.Info("Starting Extract");
            new Program();
            Console.ReadKey();
        }

        public Program()
        {
            IEnumerable<string> files = Utils.GetFilesWithExtension("z:\\Genes files", ".xml");
            IEnumerable<ContractorLicense> Licenses = Utils.GetContractorLicenses(files);
            Utils.CreateClXlsx(Licenses, filename);
        }
    }
}
