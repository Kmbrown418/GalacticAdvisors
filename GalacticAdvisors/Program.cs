using System;
using System.Linq;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace GalacticAdvisors
{
    internal class Program
    {
        static void Main(string[] args)
        {
            XLWorkbook exportFile = new XLWorkbook();
            DataTable dt = BuildPage();
            DataSet ds = new DataSet();

            ds.Tables.Add(dt);
            exportFile.Worksheets.Add(ds);

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string savePath = Path.Combine(desktopPath, "GalacticAdvisors.xlsx");
            exportFile.SaveAs(savePath, false);
        }

        private static DataTable BuildPage()
        {
            //Sets up the columns to export
            DataTable table1 = new DataTable() { TableName = "Data Discovered" };
            table1.Columns.Add("Path");
            table1.Columns.Add("Title");
            table1.Columns.Add("Values");
            table1.Columns.Add("Date");

            //gathers files from the directory and sets the banned file types
            //Wasnt sure which directory to start scaning, so I decided to just do the desktop. 
            //string[] files = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory);
            string[] files = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "*.*", SearchOption.AllDirectories);
            string[] bannedFileTypes = {".png", ".jpg", ".mp3", ".mp4", ".mov"};

            DateTime Starttime = DateTime.Now; //sets the start time, this isnt needed but I figured since the performance is meant to be under 5 minutes that there would be no point to keep going after 5 minutes
            foreach (string file in files)
            {
                //see above comment
                if (DateTime.Now == Starttime.AddMinutes(5))
                {
                    break;
                }

                //gets all file data
                FileInfo fi = new FileInfo(file);
                if (!bannedFileTypes.Contains(fi.Extension.ToLower() ))
                {
                    string FileTitle = file.Split('\\')[file.Split('\\').Count() - 1];
                    string path = fi.Directory.ToString();

                    //searched the file if its an approved file type
                    string FoundData = GetText(path + "\\" + FileTitle);

                    //if data is found, save it. 
                    if (FoundData != "") {
                        table1.Rows.Add(fi.Directory.ToString(), FileTitle, FoundData, DateTime.Now);
                    }
                }
            }
            return table1;
        }

        private static string GetText(string fullpath)
        {
            string line;

            try
            {
                StreamReader sr = new StreamReader(fullpath);
                line = sr.ReadLine();

                string resultString = "";
                while (line != null)
                {
                    string strRegex = "";
                    //Regex for Social Security Number
                    strRegex = @"^(?!666|000|9\d{2})\d{3}-(?!00)\d{2}-(?!0{4})\d{4}";
                    string resultSSN = Regex.Match(line.Trim(), strRegex).Value;

                    //Regex for most common cards
                    strRegex = @"^(^4[0-9]{12}(?:[0-9]{3})?$)|(^(?:5[1-5][0-9]{2}|222[1-9]|22[3-9][0-9]|2[3-6][0-9]{2}|27[01][0-9]|2720)[0-9]{12}$)|(3[47][0-9]{13})|(^3(?:0[0-5]|[68][0-9])[0-9]{11}$)|(^6(?:011|5[0-9]{2})[0-9]{12}$)|(^(?:2131|1800|35\d{3})\d{11}$)";
                    string resultCC = Regex.Match(line.Replace(",", "").Replace("-","").Trim(), strRegex).Value;

                    //combines all found data into a string to store
                    if (resultSSN != "")
                    {
                        resultString += "SSN: " + resultSSN + " |";
                    }
                    if (resultCC != "")
                    {
                        resultString += "CCN: " + resultCC + " |";
                    }

                    line = sr.ReadLine();
                }

                //closes and returns
                sr.Close();
                return resultString.Trim('|');
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
    }
}
