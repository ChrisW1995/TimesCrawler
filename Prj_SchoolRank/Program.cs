using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
// Requires reference to WebDriver.Support.dll
using OpenQA.Selenium.Support.UI;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Threading;

namespace Prj_SchoolRank
{
    class Program
    {
       
        public const string SCHOOL = "SchName";
        public const string RANK = "Rank";
        public const string YEAR = "Year";
        static DataColumn column;
        static DataRow data_row;
        static int rowCount = 0;
        static string region = "";
        static int[] years = new int[2];
        static void Main(string[] args)
        {
            while (true)
            {
                Console.Write("Which region do you wanna get? (1: Asia, 2: World): ");
                region = GetWebCondition(int.Parse(Console.ReadLine()));
                if(region == "")
                {
                    Console.WriteLine("Plz enter following number. (1: Asia, 2: World): ");
                }
                else
                {
                    while (true)
                    {
                        Console.Write("Enter a years range (separate from a spane): ");
                        string[] arr = Console.ReadLine().Split(' ');
                        if(arr.Length != 2)
                            Console.WriteLine("Error, check that whether enter only two numbers.");
                        else
                        {
                            years[0] = int.Parse(arr[0].ToString());
                            years[1] = int.Parse(arr[1].ToString());
                            break;
                        }
                    }
                    break;
                }
                
            }
            

            DataTable table = new DataTable(); 
            IWorkbook workbook = new XSSFWorkbook();


            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = SCHOOL;
            table.Columns.Add(column);

            // Create second column.
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName =  RANK;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.Int32");
            column.ColumnName = YEAR;
            table.Columns.Add(column);

            try
            {
                for (int i = years[0]; i <= years[1]; i++)
                {
                    rowCount = 0;
                    Console.WriteLine($"===================={i}===================== ");
                    using (IWebDriver driver = new FirefoxDriver())
                    {
                        ISheet sheet1 = workbook.CreateSheet(i.ToString());
                        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                        //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(8);
                        driver.Url = $"https://www.timeshighereducation.com/world-university-rankings/{i}/{region}-ranking#!/page/0/length/-1/sort_by/rank/sort_order/asc/cols/stats";
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

                        wait.Until(d => d.PageSource);
                        IList<IWebElement> statsElement = wait.Until(x => x.FindElements(By.TagName("tr")).Skip(1).ToList());
                        //IList<IWebElement> statsElement = driver.FindElements(By.TagName("tr")).Skip(1).ToList(); //stats tab list 
                        IList<IWebElement> table_stats_ths = wait.Until(x => x.FindElements(By.TagName("th")).ToList());
                        IRow column_row = sheet1.CreateRow(rowCount);
                        for (int index = 0; index < table_stats_ths.Count; index++)
                        {
                            column_row.CreateCell(index).SetCellValue(table_stats_ths[index].Text);
                        }
                        rowCount++;
                        foreach (var item in statsElement)
                        {

                            IRow row = sheet1.CreateRow(rowCount);
                            string schoolName = item.FindElement(By.TagName("a")).Text;
                            string rank = item.FindElement(By.ClassName("rank")).Text;
                            var tds = item.FindElements(By.TagName("td")).Skip(2).ToList();
                            Console.Write("\r" + new string(' ', Console.BufferWidth - 10));
                            Console.Write("\r{0}\t{1}", rank, schoolName);
                            row.CreateCell(0).SetCellValue(rank);
                            row.CreateCell(1).SetCellValue(schoolName);
                            for (int index = 0; index < tds.Count; index++)
                            {
                                row.CreateCell(index + 2).SetCellValue(tds[index].Text);
                            }
                            data_row = table.NewRow();
                            data_row[SCHOOL] = schoolName;
                            data_row[RANK] = rank;
                            data_row[YEAR] = i;
                            table.Rows.Add(data_row);
                            rowCount++;
                        }
                        Console.WriteLine();
                        Thread.Sleep(2000);
                        IWebElement tab = driver.FindElement(By.XPath("//*[@id=\"block-system-main\"]/div/div[3]/div/div[1]/div[1]/div/div[1]/ul/li[2]/label"));
                        tab.Click();
                        Thread.Sleep(3000);
                        IList<IWebElement> table_score_ths = wait.Until(x => x.FindElements(By.TagName("th")).Skip(2).ToList());
                        IList<IWebElement> scoreElement = wait.Until(x => x.FindElements(By.TagName("tr")).Skip(1).ToList()); //scores tab list
                                                                                                                              //Get score tab column 
                        for (int index = 0; index < table_score_ths.Count; index++)
                        {
                            column_row.CreateCell(index + table_stats_ths.Count).SetCellValue(table_score_ths[index].Text);
                        }

                        for (int index = 0; index < scoreElement.Count; index++)
                        {
                            Console.Write("\rWriting data.. {0}/{1}", index + 1, scoreElement.Count);
                            var tds = scoreElement[index].FindElements(By.TagName("td")).Skip(2).ToList();
                            for (int index_2 = 0; index_2 < tds.Count(); index_2++)
                            {
                                sheet1.GetRow(index + 1).CreateCell(table_stats_ths.Count + index_2).SetCellValue(tds[index_2].Text);
                            }

                        }
                        sheet1.AutoSizeColumn(1);
                        Console.WriteLine();
                        driver.Quit();
                    }
                    Console.WriteLine("Completed! Next year..");
                }
                Console.WriteLine("Completed.");
                ISheet stati_sheet = workbook.CreateSheet("statistics");
                IRow year_row = stati_sheet.CreateRow(0);
                int count = 1;
                Console.WriteLine("analyzing data..");
                Console.Write($"\t\t\t\t\t");
                for (int i = years[0]; i <= years[1]; i++)
                {
                    Console.Write($"{i}\t");
                    year_row.CreateCell(count).SetCellValue(i);
                    count++;
                }
                Console.WriteLine();
                var query = table.AsEnumerable().GroupBy(r => r[SCHOOL]).Select(x => new { school = x.Key, ranks = x.ToList() });
                rowCount = 1;
                foreach (var item in query)
                {
                    Console.Write($"{item.school}\t");
                    IRow _staticRow = stati_sheet.CreateRow(rowCount++);
                    _staticRow.CreateCell(0).SetCellValue(item.school.ToString());
                    foreach (var rank in item.ranks)
                    {
                        _staticRow.CreateCell(int.Parse(rank[YEAR].ToString()) - years[0] + 1).SetCellValue(rank[RANK].ToString());
                        Console.Write($"{rank[RANK]}\t");
                    }
                    Console.WriteLine();
                }
                stati_sheet.AutoSizeColumn(0);
                FileStream sw = File.Create($"{region} University Rankings_{years[0]}-{years[1]}.xlsx");
                workbook.Write(sw);
                sw.Close();
                Console.WriteLine("Fetching complete! Enter any key to exit.");
            }
            catch (Exception)
            {
                Console.WriteLine("Fetching Data Failed. Please try again or check whether years range is exist.");
                throw;
            }
           
            Console.Read();
        }
        

        public static string GetWebCondition(int num)
        {
            string region = ""; 
            switch (num)
            {
                case 1:
                    region = "regional";
                    break;
                case 2:
                    region = "world";
                    break;
                default:
                    break;

            }
            return region;
        }

    }
}
