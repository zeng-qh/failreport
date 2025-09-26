using FailReport.Controllers;
using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO.Compression;
using System.Numerics;
using System.Runtime.InteropServices;

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
                var filePath =  Path.Combine(@"C:\TE\N51876BTestApp\N51876BTestApp\Resources\", fileName);
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
                    string itemCode = GetCellValue(row, headerDict, "ItemCode");
                    string itemName = GetCellValue(row, headerDict, "ItemName");
                    string P = @"ProdName
ProdCode:
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

                return ("---------------" + "生成完成 " + sheet.LastRowNum + "条数据！----------------------");
                //Console.WriteLine("---------------" + "生成完成 " + sheet.LastRowNum + "条数据！----------------------");

            }
        }

        private string GetCellValue(IRow row, Dictionary<int, string> headerDict, string columnName)
        {
            var column = headerDict.FirstOrDefault(h => h.Value == columnName);
            if (column.Key >= 0 && column.Key < row.LastCellNum)
            {
                ICell cell = row.GetCell(column.Key);
                return cell?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }

        public Task UnZipFile(string filePath)
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
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解压文件失败: {ex.Message}");
                throw;
            }
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
        Task UnZipFile(string filePath);
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
