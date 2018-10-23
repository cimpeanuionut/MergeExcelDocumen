using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;

namespace MergeExcelDocument
{
    class Program
    {
        public static string imagePathReport4full = @"\ResultReport4Item.xlsx";
        public static string imagePathReport3full = @"\ResultReport3Item.xlsx";
        public static string finalImagePathReport4full = AssemblyDirectory + imagePathReport4full;
        public static string finalImagePathReport3full = AssemblyDirectory + imagePathReport3full;

        //Path for Report4
        //Project1
        public static string Project1PathReport4 = @"\TAMUExport\Project1\Report4Item.xls";

        //Project2
        public static string Project2PathReport4 = @"\TAMUExport\Project2\Report4Item.xls";

        //Project3
        public static string Project3PathReport4 = @"\TAMUExport\Project3\Report4Item.xls";

        //Project4
        public static string Project4PathReport4 = @"\TAMUExport\Project4\Report4Item.xls";

        //Project5
        public static string Project5PathReport4 = @"\TAMUExport\Project5\Report4Item.xls";

        //Project6
        public static string Project6PathReport4 = @"\TAMUExport\Project6\Report4Item.xls";

        //Project7
        public static string Project7PathReport4 = @"\TAMUExport\Project7\Report4Item.xls";

        //Project8
        public static string Project8PathReport4 = @"\TAMUExport\Project8\Report4Item.xls";

        //Project9   
        public static string Project9PathReport4 = @"\TAMUExport\Project9\Report4Item.xls";

        //Project10
        public static string Project10PathReport4 = @"\TAMUExport\Project10\Report4Item.xls";

        //Project11
        public static string Project11PathReport4 = @"\TAMUExport\Project11\Report4Item.xls";

        //Project12
        public static string Project12PathReport4 = @"\TAMUExport\Project12\Report4Item.xls";

        //Project13
        public static string Project13PathReport4 = @"\TAMUExport\Project13\Report4Item.xls";

        //Project14
        public static string Project14PathReport4 = @"\TAMUExport\Project14\Report4Item.xls";

        //Project15
        public static string Project15PathReport4 = @"\TAMUExport\Project15\Report4Item.xls";

        //Project16
        public static string Project16PathReport4 = @"\TAMUExport\Project16\Report4Item.xls";

        //Project17
        public static string Project17PathReport4 = @"\TAMUExport\Project17\Report4Item.xls";

        //Project18
        public static string Project18PathReport4 = @"\TAMUExport\Project18\Report4Item.xls";

        //Project19
        public static string Project19PathReport4 = @"\TAMUExport\Project19\Report4Item.xls";

        //Project20
        public static string Project20PathReport4 = @"\TAMUExport\Project20\Report4Item.xls";

        //Path for Report3
        //Project1
        public static string Project1PathReport3 = @"\TAMUExport\Project1\Report3Item.xls";

        //Project2
        public static string Project2PathReport3 = @"\TAMUExport\Project2\Report3Item.xls";

        //Project3
        public static string Project3PathReport3 = @"\TAMUExport\Project3\Report3Item.xls";

        //Project4
        public static string Project4PathReport3 = @"\TAMUExport\Project4\Report3Item.xls";

        //Project5
        public static string Project5PathReport3 = @"\TAMUExport\Project5\Report3Item.xls";

        //Project6
        public static string Project6PathReport3 = @"\TAMUExport\Project6\Report3Item.xls";

        //Project7
        public static string Project7PathReport3 = @"\TAMUExport\Project7\Report3Item.xls";

        //Project8
        public static string Project8PathReport3 = @"\TAMUExport\Project8\Report3Item.xls";

        //Project9   
        public static string Project9PathReport3 = @"\TAMUExport\Project9\Report3Item.xls";

        //Project10
        public static string Project10PathReport3 = @"\TAMUExport\Project10\Report3Item.xls";

        //Project11
        public static string Project11PathReport3 = @"\TAMUExport\Project11\Report3Item.xls";

        //Project12
        public static string Project12PathReport3 = @"\TAMUExport\Project12\Report3Item.xls";

        //Project13
        public static string Project13PathReport3 = @"\TAMUExport\Project13\Report3Item.xls";

        //Project14
        public static string Project14PathReport3 = @"\TAMUExport\Project14\Report3Item.xls";

        //Project15
        public static string Project15PathReport3 = @"\TAMUExport\Project15\Report3Item.xls";

        //Project16
        public static string Project16PathReport3 = @"\TAMUExport\Project16\Report3Item.xls";

        //Project17
        public static string Project17PathReport3 = @"\TAMUExport\Project17\Report3Item.xls";

        //Project18
        public static string Project18PathReport3 = @"\TAMUExport\Project18\Report3Item.xls";

        //Project19
        public static string Project19PathReport3 = @"\TAMUExport\Project19\Report3Item.xls";

        //Project20
        public static string Project20PathReport3 = @"\TAMUExport\Project20\Report3Item.xls";

        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return System.IO.Path.GetDirectoryName(path);
            }
        }

        static void Main(string[] args)
        {

            //Merge Report4Items
            MergeExcel.DoMerge(new string[]
             {
            AssemblyDirectory + Project1PathReport4, AssemblyDirectory + Project2PathReport4, AssemblyDirectory + Project3PathReport4, AssemblyDirectory + Project4PathReport4, AssemblyDirectory + Project5PathReport4,
            AssemblyDirectory + Project6PathReport4, AssemblyDirectory + Project7PathReport4, AssemblyDirectory + Project8PathReport4, AssemblyDirectory + Project9PathReport4, AssemblyDirectory + Project10PathReport4,
            AssemblyDirectory + Project11PathReport4, AssemblyDirectory + Project12PathReport4, AssemblyDirectory + Project13PathReport4, AssemblyDirectory + Project14PathReport4, AssemblyDirectory + Project15PathReport4,
            AssemblyDirectory + Project16PathReport4, AssemblyDirectory + Project17PathReport4, AssemblyDirectory + Project18PathReport4, AssemblyDirectory + Project19PathReport4, AssemblyDirectory + Project20PathReport4
            },
            finalImagePathReport4full, "I", 2);

            //Merge Report3Items
            MergeExcel.DoMerge(new string[]
             {
            AssemblyDirectory + Project1PathReport3, AssemblyDirectory + Project2PathReport3, AssemblyDirectory + Project3PathReport3, AssemblyDirectory + Project4PathReport3, AssemblyDirectory + Project5PathReport3,
            AssemblyDirectory + Project6PathReport3, AssemblyDirectory + Project7PathReport3, AssemblyDirectory + Project8PathReport3, AssemblyDirectory + Project9PathReport3, AssemblyDirectory + Project10PathReport3,
            AssemblyDirectory + Project11PathReport3, AssemblyDirectory + Project12PathReport3, AssemblyDirectory + Project13PathReport3, AssemblyDirectory + Project14PathReport3, AssemblyDirectory + Project15PathReport3,
            AssemblyDirectory + Project16PathReport3, AssemblyDirectory + Project17PathReport3, AssemblyDirectory + Project18PathReport3, AssemblyDirectory + Project19PathReport3, AssemblyDirectory + Project20PathReport3
            },
            finalImagePathReport3full, "G", 2);

        }

        }
    }  
 
