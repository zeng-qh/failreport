using FailReport.Controllers;
using MiniExcelLibs;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Linq.Expressions;
using System.Numerics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Transactions;

namespace FailReport.Upload
{

    public class ExcelUploadService : IExcelUploadService
    {

        private string _uploadPath;

        public ExcelUploadService()
        {
            // 使用配置文件或环境变量来设置路径 - 跨平台兼容
            // 在Windows上会使用C:\Logs_Uploads，在Linux上会使用/home/user/Logs_Uploads
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                FailReportController.LogPath = "/home/user/Logs_Uploads";
                FailReportController.FCTPath = "/home/user/Logs_FCT";
                FailReportController.DefaultPathName = "/home/user/Logs_Uploads";
            }
            else
            {
                FailReportController.LogPath = @"C:\Logs_Uploads";
                FailReportController.FCTPath = @"C:\Logs_FCT";
                FailReportController.DefaultPathName = @"C:\Logs_Uploads";
            }
            _uploadPath = FailReportController.LogPath;
            // 若文件夹不存在，则创建
            if (!Directory.Exists(_uploadPath))
            {
                Directory.CreateDirectory(_uploadPath);
            }
            if (!Directory.Exists(FailReportController.FCTPath))
            {
                Directory.CreateDirectory(FailReportController.FCTPath);
            }

        }

        public async Task<FileUploadResult> SaveExcelFileAsync(IFormFile file)
        {
            try
            {
                var fileName = file.FileName;
                var filePath = Path.Combine(@"C:\TE\N51876BTestApp\N51876BTestApp\Resources\", fileName);
                // 文件类型若是.csv 
                if (Path.GetExtension(file.FileName).ToLower() != ".csv")
                {// 生成唯一的文件名，避免覆盖
                    fileName = $"{Guid.NewGuid().ToString().Substring(0, 8)}_{file.FileName}";// {Path.GetExtension(file.FileName)}
                    filePath = Path.Combine(_uploadPath, fileName);
                }

                // 保存文件到服务器
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                return new FileUploadResult
                {
                    Success = true,
                    FilePath = filePath
                };
            }
            catch (Exception ex)
            {
                return new FileUploadResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }
        public async Task<string> ReadExcelData2Txt(string filePath)
        {
            HashSet<string> seen = new HashSet<string>();
            // 替换为Excel文件路径
            // string filePath = @"D:\LOG\1869B1 FCT基础数据_20250804.xlsx";
            string LogsPath = Path.ChangeExtension(filePath, null);
            if (!Directory.Exists(LogsPath))//如果不存在就创建file文件夹
            {
                Directory.CreateDirectory(LogsPath);
            }
            if (!File.Exists(filePath))
            {
                Console.WriteLine("文件不存在！");
                return ("文件不存在！");
            }
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(0); // 获取工作表

                if (sheet.LastRowNum < 0)
                {
                    Console.WriteLine("Excel文件中没有数据！");
                    return ("Excel文件中没有数据！");
                }

                // 获取表头行
                IRow headerRow = sheet.GetRow(0);
                Dictionary<int, string> headerDict = new Dictionary<int, string>();

                // 构建列名到索引的映射
                for (int col = 0; col < headerRow.LastCellNum; col++)
                {
                    ICell? cell = headerRow.GetCell(col);
                    if (cell != null)
                    {
                        //headerDict[col] = cell.ToString();
                        //headerDict[col] = cell.ToString()!; 
                        headerDict[col] = cell.ToString() ?? "";
                    }
                }


                string serialNumber = string.Empty;
                string result = string.Empty;


                // 遍历所有数据行（从第二行开始）
                //for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)
                // 从最后一行开始
                for (int rowNum = sheet.LastRowNum; rowNum >= 1; rowNum--)
                {
                    IRow row = sheet.GetRow(rowNum);
                    if (row == null) continue;

                    serialNumber = GetCellValue(row, headerDict, "SerialNumber");
                    if (!seen.Add(serialNumber)) // Add失败说明已存在（重复）
                    {
                        continue;
                    }
                    result = GetCellValue(row, headerDict, "Result");
                    string testTime = GetCellValue(row, headerDict, "TestTime");
                    string fixture = GetCellValue(row, headerDict, "Fixture");
                    string cavity = GetCellValue(row, headerDict, "Cavity");
                    string operatorName = GetCellValue(row, headerDict, "Operator");
                    string orderNo = GetCellValue(row, headerDict, "OrderNO");
                    string ProdCode = GetCellValue(row, headerDict, "ItemCode");
                    string ProdName = GetCellValue(row, headerDict, "ItemName");
                    string P = @"ProdName:" + ProdName + @"
ProdCode:" + ProdCode + @"
SerialNumber:" + serialNumber + @"
TestUser: " + (operatorName.Replace("Operator/", "")) + @"
TestFixture: " + fixture + @"
Cavity: " + cavity + @"
TestTime: " + testTime + @"
TestResult: " + result + @"
Step	Item	Range	TestValue	Result
";

                    //Console.WriteLine($"行 {rowNum + 1}: SerialNumber = {serialNumber}");
                    //Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~``");
                    string TpTxt = string.Empty;
                    int Index = 1;
                    // 获取测试 Step 
                    foreach (var header in headerDict)
                    {

                        if (header.Value.Contains("Step"))
                        {
                            string stepValue = GetCellValue(row, headerDict, header.Value);
                            //char[] Ch = stepValue.ToArray();
                            //Ch.Dump();                       
                            //TpTxt += $"{Index}\t{(stepValue.Replace('\u200B', '\t'))}\r\n";
                            string[] StepList = stepValue.Split('\u200B');
                            if (StepList.Length < 6)
                            {
                                //Console.WriteLine($"行 {rowNum + 1}: SerialNumber = {serialNumber} 的 Step 数据不完整，跳过该行。");
                                continue;
                            }
                            TpTxt += $"{Index}\t{StepList[0]}\t{StepList[1]}~{StepList[2]}\t{StepList[3]}\t{StepList[5]}\r\n";
                            // Console.WriteLine($"{header.Value}: {stepValue}");
                            Index++;
                        }

                    }
                    string TxtName = LogsPath + "\\" + serialNumber + "___" + DateTime.Now.ToString("yyyyMMddhhmmssffff") + "___" + result + ".txt";
                    //TxtName.Dump();
                    using (StreamWriter writer = new StreamWriter(TxtName))
                    {
                        writer.Write((P + TpTxt));
                        //Thread.Sleep(100);
                    }
                    //(P + TpTxt).Dump();
                    Console.WriteLine("---------------" + TxtName + "生成完成！----------------------");
                }

                //return ("---------------" + "生成完成 " + sheet.LastRowNum + "条数据！----------------------");
                //Console.WriteLine("---------------" + "生成完成 " + sheet.LastRowNum + "条数据！----------------------");
                return LogsPath;
            }
        }

        private static string GetCellValue(IRow row, Dictionary<int, string> headerDict, string columnName)
        {
            var column = headerDict.FirstOrDefault(h => h.Value == columnName);

            if (column.Key >= 0 && column.Key < row.LastCellNum)
            {
                ICell cell = row.GetCell(column.Key);
                if ((!(cell?.DateCellValue is null)) && cell?.DateCellValue.Value.ToString().Length > 0)
                {
                    return cell?.DateCellValue.ToString()??"";
                }
                return cell?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }

        public async Task<string> UnZipFile(string filePath)
        {
            try
            {
                string extractPath = Path.Combine(Path.GetDirectoryName(filePath)!, Path.GetFileNameWithoutExtension(filePath));
                if (!Directory.Exists(extractPath))
                {
                    Directory.CreateDirectory(extractPath);
                }
                using (var archive = System.IO.Compression.ZipFile.OpenRead(filePath))
                {
                    foreach (var entry in archive.Entries)
                    {
                        string destinationPath = Path.Combine(extractPath, entry.FullName);
                        if (string.IsNullOrEmpty(entry.Name)) continue; // Skip directories
                        entry.ExtractToFile(destinationPath, true);
                    }
                }
                await Task.CompletedTask;
                return extractPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解压文件失败: {ex.Message}");
                throw;
            }
        }

        public async Task<string> ReadExcelData2Csv(string filePath, string Type)
        {
            try
            {
                string Dev = Path.GetFileNameWithoutExtension(filePath);
                string guidPart = Regex.Match(Dev, @"^[^_]+").Value + "_";
                //ICT测试数据转为csv
                //FCT测试数据转为csv 
                string _newCsvPath = Path.Combine("C:\\TE", "CsvDataPub", guidPart + Type + "ExcelToCsv_数据.csv");
                if (Type == "FCT")
                {
                    return "处理文件成功>" + await FCTExcelToCsv(filePath, _newCsvPath);
                }
                if (Type == "ICT")
                {
                    return "处理文件成功>" + await ICTExcelToCsv(filePath, _newCsvPath);
                }
                else
                {
                    return "处理文件失败> 转换类型为：" + Type;
                }
            }
            catch (Exception ex)
            {

                return $"处理文件失败> {ex.Message}";
            }
            return $"处理文件失败> 未知的类型 {Type}";
        }


        public async Task<string> FCTExcelToCsv(string filePath, string _newCsvPath)
        {// 配置文件路径
         //string Dev = "10364860FCT基础数据_20251023";
         //string filePath = @"D:\20251015\" + Dev + ".xlsx";

            if (!File.Exists(filePath))
            {
                Console.WriteLine("文件不存在！");
                return "文件不存在！";
            }

            // 读取Excel数据，将第一行作为键
            var rows = MiniExcel.Query(filePath, true).ToList();

            if (rows.Count == 0)
            {
                Console.WriteLine("Excel文件中没有数据！");
                return "Excel文件中没有数据！";
            }

            // 获取表头行
            IDictionary<string, object> headerRow = rows[0] as IDictionary<string, object>;
            IDictionary<string, object> headerRowText = rows[1] as IDictionary<string, object>;

            // 收集所有包含"Step"的测试项目列
            var testColumns = headerRow.Where(r => r.Key.Contains("Step"))
                                      .Select(r => r.Key)
                                      .ToList();
            //await  File.WriteAllText(_newCsvPath, $"", new UTF8Encoding(true));
            await File.WriteAllTextAsync(_newCsvPath, $"", new UTF8Encoding(true));
            // 构建表头测试项目
            StringBuilder csvContent = new StringBuilder();
            StringBuilder csvTestItme = new StringBuilder();
            csvTestItme.Append($"产品条码/测试项目,result,TestTime,Fixture,Cavity,Operator,OrderNO");// 测试项目

            StringBuilder csvMax = new StringBuilder(); //
            csvMax.Append($"上限,,,,,,,,,");// 测试项目
            StringBuilder csvMin = new StringBuilder();
            csvMin.Append($"下限,,,,,,,,,");// 测试项目
            StringBuilder csvUnit = new StringBuilder();
            csvUnit.Append($"单位,,,,");// 测试项目  
            foreach (var column in testColumns)
            { // 获取测试项目值

                string testValue = headerRowText[column]?.ToString() ?? "";

                if (!string.IsNullOrEmpty(testValue))
                {
                    string[] head = testValue.Split('\u200B');
                    csvTestItme.Append($",{QuoteCsvField(head[0])}");
                    csvMax.Append($",{QuoteCsvField(head[1])}");
                    csvMin.Append($",{QuoteCsvField(head[2])}");
                }
            }
            File.AppendAllText(_newCsvPath, $"{csvTestItme.ToString()}\r\n", new UTF8Encoding(true));

            File.AppendAllText(_newCsvPath, $"{csvMin.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{csvMax.ToString()}\r\n", new UTF8Encoding(true));



            //遍历所有Values
            for (int i = 0; i < rows.Count; i++)
            //for (int i = 0; i < 9; i++)
            {
                var row = rows[i] as IDictionary<string, object>;

                // 获取基本信息
                string serialNumber = row["SerialNumber"]?.ToString() ?? "";
                string result = row["Result"]?.ToString() ?? "";
                string testTime = row["TestTime"]?.ToString() ?? "";
                string fixture = row["Fixture"]?.ToString() ?? "";
                string cavity = row["Cavity"]?.ToString() ?? "";
                string operatorName = row["Operator"]?.ToString() ?? "";
                string orderNo = row["OrderNO"]?.ToString() ?? "";
                string itemCode = row["ItemCode"]?.ToString() ?? "";
                string itemName = row["ItemName"]?.ToString() ?? "";

                Console.WriteLine($"处理行 {i + 1}: SerialNumber = {serialNumber}");
                csvContent.Append($"{serialNumber},{result},{testTime},{fixture},{cavity},{operatorName},{orderNo}");// 测试项目 
                //		// 处理每个测试项目，实现列转行
                foreach (var column in testColumns)
                {
                    string TpValStr = row[column]?.ToString() ?? ""; // 使用?可以判断null 但是无法判断空字符串的情况

                    // 获取测试项目值
                    if (!string.IsNullOrEmpty(TpValStr))
                    {
                        string[] TestValue = TpValStr.Split('\u200B');
                        csvContent.Append($",{QuoteCsvField(TestValue[3])}");// 测试项目 
                    }
                }
                csvContent.Append(",\r\n");
            }

            // 写入CSV文件（使用UTF8编码，避免中文乱码）
            //File.WriteAllText(_newCsvPath, csvContent.ToString(), new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, csvContent.ToString(), new UTF8Encoding(true));

            Console.WriteLine($"转换完成，文件已保存至: {_newCsvPath}");
            return _newCsvPath;
        }
        public async Task<string> ICTExcelToCsv(string filePath, string _newCsvPath)
        {

            // 配置文件路径 
            // 37ba7c7e_1_ICT原始数据.xlsx
            //获取文件名称的GUID部分 使用正则表达式 

            if (!File.Exists(filePath))
            {
                Console.WriteLine("文件不存在！");
                return "文件不存在！";
            }

            // 读取Excel数据，将第一行作为键
            var rows = MiniExcel.Query(filePath, true).ToList();

            if (rows.Count == 0)
            {
                Console.WriteLine("Excel文件中没有数据！");
                return "Excel文件中没有数据！";
            }

            // 获取表头行
            IDictionary<string, object> headerRow = rows[0] as IDictionary<string, object>;
            IDictionary<string, object> headerRowText = rows[1] as IDictionary<string, object>;

            // 收集所有包含"Step"的测试项目列
            var testColumns = headerRow.Where(r => r.Key.Contains("Step"))
                                      .Select(r => r.Key)
                                      .ToList();
            //foreach (var column in testColumns)
            //{
            //	column.Dump();
            //}
            //File.WriteAllText(_newCsvPath, $"", new UTF8Encoding(true));
            await File.WriteAllTextAsync(_newCsvPath, $"", new UTF8Encoding(true));
            // 构建表头测试项目
            StringBuilder csvContent = new StringBuilder();
            StringBuilder csvTestItme = new StringBuilder();
            csvTestItme.Append($"产品条码/测试项目");// 测试项目


            StringBuilder csvMin = new StringBuilder();
            csvMin.Append($"下限");// 测试项目
            StringBuilder csvMax = new StringBuilder(); //
            csvMax.Append($"上限");// 测试项目
            StringBuilder dif = new StringBuilder(); //
            dif.Append($"差值");// 测试项目
            StringBuilder csvUnit = new StringBuilder();
            csvUnit.Append($"单位");// 测试项目 
            StringBuilder csvTestTyep = new StringBuilder();
            csvTestTyep.Append($"测试项目类型");// 测试项目 

            StringBuilder R_Min = new StringBuilder();
            R_Min.Append($"最小值");// 测试项目 
            StringBuilder R_Max = new StringBuilder();
            R_Max.Append($"最大值");// 测试项目 
            StringBuilder R_Avg = new StringBuilder();
            R_Avg.Append($"平均值");// 测试项目 

            csvTestItme.Append($",result,Fixture,Cavity,TestTime");
            //csvTestItme.Append($"");// 侧架编号
            //csvTestItme.Append($"");// 穴位
            csvTestTyep.Append($",,,,");
            csvMax.Append($",,,,");
            csvMin.Append($",,,,");
            dif.Append($",,,,");
            R_Min.Append($",,,,");
            R_Max.Append($",,,,");
            R_Avg.Append($",,,,");
            //testColumns.Count().Dump();
            foreach (var column in testColumns)
            { // 获取测试项目值
                string testValue = headerRowText[column]?.ToString() ?? "";

                if (!string.IsNullOrEmpty(testValue))
                {
                    string head = Regex.Replace(testValue, @" +", ",", RegexOptions.Multiline);
                    //testValue.Dump();

                    csvTestItme.Append($",{head.Split(',')[1]}");
                    csvTestTyep.Append($",{head.Split(',')[2]} {head.Split(',')[3]}");
                    //csvMax.Append($",{head.Split(',')[4]}");
                    //csvMin.Append($",{head.Split(',')[5]}");
                    double V = 0;
                    double.TryParse(Regex.Match(head.Split(',')[3], @"^\d+(\.\d+)?").Value, out V);
                    //if (V>0)
                    //{
                    //	
                    //}
                    double Max = 0;
                    double Min = 0;
                    double.TryParse(Regex.Match(head.Split(',')[4], @"^\d+(\.\d+)?").Value, out Max);
                    double.TryParse(Regex.Match(head.Split(',')[5], @"^\d+(\.\d+)?").Value, out Min);
                    double MaxVal = (V * (1 + (Max / 100)));
                    double MinVal = (V * (1 - (Max / 100)));
                    csvMax.Append($",{Math.Round(MaxVal, 4)}");
                    csvMin.Append($",{Math.Round(MinVal, 4)}");
                    dif.Append($",{Math.Round(MaxVal - MinVal, 4)}");
                    R_Min.Append($",=MIN(F9:F{8 + rows.Count})");
                    R_Max.Append($",=Max(F9:F{8 + rows.Count})");
                    R_Avg.Append($",=AVERAGE(F9:F{8 + rows.Count})");


                }
            }
            File.AppendAllText(_newCsvPath, $"{csvTestItme.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{csvTestTyep.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{csvMin.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{csvMax.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{dif.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{R_Min.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{R_Max.ToString()}\r\n", new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, $"{R_Avg.ToString()}\r\n", new UTF8Encoding(true));



            //遍历所有Values
            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i] as IDictionary<string, object>;

                // 获取基本信息
                string serialNumber = row["SerialNumber"]?.ToString() ?? "";
                string result = row["Result"]?.ToString() ?? "";
                string testTime = row["TestTime"]?.ToString() ?? "";
                string fixture = row["Fixture"]?.ToString() ?? "";
                string cavity = row["Cavity"]?.ToString() ?? "";
                string operatorName = row["Operator"]?.ToString() ?? "";
                string orderNo = row["OrderNO"]?.ToString() ?? "";
                string itemCode = row["ItemCode"]?.ToString() ?? "";
                string itemName = row["ItemName"]?.ToString() ?? "";

                Console.WriteLine($"处理行 {i + 1}: SerialNumber = {serialNumber}");
                csvContent.Append($"{serialNumber},{result},{fixture},{cavity},{testTime}");// 测试项目
                                                                                            //		// 处理每个测试项目，实现列转行
                foreach (var column in testColumns)
                {
                    // 获取测试项目值
                    string testValue = row[column]?.ToString() ?? "";

                    if (!string.IsNullOrEmpty(testValue))
                    {
                        string TestValue = Regex.Replace(testValue, @" +", ",", RegexOptions.Multiline);// 空格替换为,

                        // TestValue.Split(",")[6].Dump();
                        //csvContent.Append($",{TestValue.Split(",")[6]}");// 测试项目
                        //csvContent.Append($",{TestValue.Split(",")[1]}");// 测试项目
                        //csvContent.Append($",{column}");// 测试项目
                        double _TestVal = 0;
                        double.TryParse(Regex.Match(TestValue.Split(",")[6], @"^\d+(\.\d+)?").Value, out _TestVal); //只提取数字部分
                        csvContent.Append($",{Math.Round(_TestVal, 2)}");// 测试项目
                    }

                }
                csvContent.Append(",\r\n");
            }

            // 写入CSV文件（使用UTF8编码，避免中文乱码）
            //File.WriteAllText(_newCsvPath, csvContent.ToString(), new UTF8Encoding(true));
            File.AppendAllText(_newCsvPath, csvContent.ToString(), new UTF8Encoding(true));

            Console.WriteLine($"转换完成，文件已保存至: {_newCsvPath}");
            return _newCsvPath;
        }

  

        // 处理CSV字段，添加引号以处理包含逗号或特殊字符的情况
        private static string QuoteCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";

            // 如果字段包含逗号、引号或换行符，需要用引号括起来
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\r") || field.Contains("\n"))
            {
                // 替换引号为两个引号
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

    }
    public interface IExcelUploadService
    {
        /// <summary>
        /// 保存Excel文件到服务器
        /// </summary>
        /// <param name="file">上传的文件</param>
        /// <returns>包含上传结果的对象</returns>
        Task<FileUploadResult> SaveExcelFileAsync(IFormFile file);
        Task<string> ReadExcelData2Txt(string filePath);
        Task<string> UnZipFile(string filePath);


        Task<string> ReadExcelData2Csv(string filePath, string Type);
     
    }

    /// <summary>
    /// 文件上传结果模型
    /// </summary>
    public class FileUploadResult
    {
        public bool Success { get; set; }
        public string FilePath { get; set; } = "";
        public string ErrorMessage { get; set; } = "";
    }



}
