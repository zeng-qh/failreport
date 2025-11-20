using FailReport.Upload;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading;
// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace FailReport.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class FailReportController : ControllerBase
    {
        private readonly IExcelUploadService _uploadService;
        const string VerifyNumer = @"^(\-|\+)?\d+(\.\d+)?$";

        private SerialPort? serialPort;

        //Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LOG");
        public static string LogPath;
        public static string FCTPath;
        public static string DefaultPathName;
        // 构造函数注入文件上传服务
        public FailReportController(IExcelUploadService uploadService)
        {
            _uploadService = uploadService;
        }

        /// <summary>
        /// 将发布目录打包为ZIP文件
        /// </summary>
        /// <param name="sourceDir">发布目录路径</param>
        /// <param name="zipFilePath">ZIP文件保存路径</param>
        /// <returns>是否打包成功</returns>
        private bool ZipPublishDirectory(string sourceDir, string zipFilePath)
        {
            try
            {
                // 如果ZIP文件已存在则删除
                if (System.IO.File.Exists(zipFilePath))
                {
                    System.IO.File.Delete(zipFilePath);
                }

                // 创建ZIP文件并添加目录内容
                // 注意：第二个参数为""表示不包含根目录本身，只包含内部文件和子目录
                ZipFile.CreateFromDirectory(sourceDir, zipFilePath, CompressionLevel.Optimal, includeBaseDirectory: false);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"打包失败：{ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 修改LogPath
        /// </summary>
        [HttpGet]
        public string SetLogsPath(string? Path)
        {
            string Mes = string.Empty;
            // 判断path是否为合法路径
            if (!string.IsNullOrEmpty(Path) && Directory.Exists(Path))
            {
                LogPath = Path;
                Mes = "修改成功！";
            }
            return System.Text.Json.JsonSerializer.Serialize(new { LogPath, Mes });
        }


        /// <summary>
        /// 获取全部Fail 
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string GetAllFail(string Path)
        {
            return GetData(Path);
        }

        /// <summary>
        /// 获取Log 列表
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public string GetLogDirectorys()
        {

            string[] folderNames = Directory.GetDirectories(LogPath);
            return System.Text.Json.JsonSerializer.Serialize(folderNames);
        }


        [HttpGet]
        public string OpenFileDirectory(string Path)
        {
            // 打开指定文件夹
            Process.Start("explorer.exe", Path);
            return Path + " 打开成功";
        }

        [HttpGet]
        public string GetFailData()
        {
            string[] AllFail = Directory.GetFileSystemEntries(FCTPath)
                .Where(m => System.Text.RegularExpressions.Regex.IsMatch(m, ("^[^\u4e00-\u9fa5]+$")))
                .ToArray();
            return System.Text.Json.JsonSerializer.Serialize(AllFail);
        }




        [HttpGet]
        public double GetTemp(string? Com, string? SendDBStr)
        {
            Com = Com != null ? Com : "COM61";
            SendDBStr = SendDBStr != null ? SendDBStr : "01 04 04 00 00 01 30 FA";
            string result = string.Empty;

            serialPort = new SerialPort(Com, 9600);
            lock (serialPort)
            {

                serialPort.Open();
                serialPort.DiscardInBuffer();
                Byte[] crcbuf;
                string[] SendDB = SendDBStr.Split(' ');
                List<Byte> bytedata = new List<Byte>();
                foreach (var item in SendDB)
                {
                    bytedata.Add(Byte.Parse(item, System.Globalization.NumberStyles.AllowHexSpecifier));
                }
                crcbuf = new Byte[bytedata.Count];
                crcbuf = bytedata.ToArray();
                serialPort.Write(crcbuf, 0, crcbuf.Count());


                int len;
                byte[] datas;
                Stopwatch sw = new Stopwatch();
                sw.Restart();
                len = serialPort.BytesToRead;
                Thread.Sleep(50);
                while (len < serialPort.BytesToRead)
                {
                    len = serialPort.BytesToRead;
                    Thread.Sleep(10);
                }
                if (len > 0)
                {
                    datas = new byte[len];
                    serialPort.Read(datas, 0, len);
                    for (int i = 0; i < datas.Length; i++)
                    {
                        string str = Convert.ToString(datas[i], 16);
                        if (str.Length == 1)
                        {
                            str = string.Format("0{0}", str);
                        }
                        result += " " + str;
                    }
                }
                serialPort.DiscardInBuffer();
                serialPort.Close();
            }
            return Convert.ToInt32(result.Trim().Substring(9, 5).Replace(" ", ""), 16);
        }

        private string GetData(string PathName)
        {
            string searchString = "Fail";
            string _Pass = "Pass";
            string _Fail = "Fail";
            FailList _failList = new FailList();
            _failList.Data = new List<FailData>();
            _failList.PassReport = new List<PassReport>();
            // 获取指定目录所有的log  csv 的或 txt
            string[] LogAll = Directory.GetFileSystemEntries(PathName);


            string[] FailS = LogAll.Where(m => m.Contains(_Fail)).ToArray();
            string[] PassS = LogAll.Where(m => m.Contains(_Pass)).ToArray();

            if (LogAll.Count() > 0)
            {
                string FileType = Path.GetExtension(LogAll[0].ToString()); //.txt
                try
                {
                    foreach (string FailPath in LogAll)
                    {
                        string TpStr = "";
                        string[] fileLines = System.IO.File.ReadAllLines(FailPath, Encoding.UTF8); // 读取文件内容
                        if (!FailPath.Contains(_Fail) && FailPath.Contains(_Pass))
                        {
                            TpStr = "Pass";
                        }
                        else
                        {
                            TpStr = "Fail";
                        }


                        foreach (string line in fileLines)
                        {
                            //    要查找的字符串
                            if (line.Contains(TpStr) && (!line.Contains("TestResult")))
                            {
                                //Console.WriteLine("Fail:\t" + line + "\t Path: \t" + FailPath);

                                bool? TempHigh = null;
                                if (line.Contains("L<=x<=H"))
                                {
                                    string[] T_IsHigh = line.Replace("\t", "").Split(",");
                                    //48;48;18
                                    if (!T_IsHigh[3].Contains(';'))
                                    {
                                        //decimal Low = decimal.Parse(T_IsHigh[1]);

                                        // 可能是非数组的情况
                                        if (System.Text.RegularExpressions.Regex.IsMatch(T_IsHigh[3], VerifyNumer) &&
                                            System.Text.RegularExpressions.Regex.IsMatch(T_IsHigh[2], VerifyNumer))
                                        {
                                            decimal High = decimal.Parse(T_IsHigh[2]);
                                            decimal Target = decimal.Parse(T_IsHigh[3]);
                                            // 2  > 8
                                            TempHigh = High < Target;
                                        }
                                        else
                                        {
                                            TempHigh = false;
                                        }
                                    }
                                    else
                                    {

                                        //string[] LowStr = T_IsHigh[1].Split(";");
                                        string[] HighStr = T_IsHigh[2].Split(";");
                                        string[] TargetStr = T_IsHigh[3].Split(";");
                                        bool[]? TempHighs = new bool[TargetStr.Count()];

                                        for (int i = 0; i < TargetStr.Count(); i++)
                                        {
                                            string High = HighStr.Length >= TargetStr.Length ? HighStr[i] : HighStr[0];
                                            // 判断是否为数字
                                            if (!System.Text.RegularExpressions.Regex.IsMatch(TargetStr[i], VerifyNumer))
                                            {
                                                TempHighs[i] = false;
                                                continue;
                                            }
                                            ;
                                            decimal TargetValue = decimal.Parse(TargetStr[i]);
                                            TempHighs[i] = decimal.Parse(High) < TargetValue;

                                        }

                                        TempHigh = TempHighs.Where(m => m == true).Count() >= 1;
                                    }
                                }

                                if (TpStr == "Fail")
                                {
                                    _failList.Data.Add(new FailData()
                                    {
                                        FailDate = System.IO.File.GetLastWriteTime(FailPath),
                                        // 这里根据后缀名判断是否要将table 替换为逗号
                                        FailItme = System.IO.Path.GetExtension(FailPath).ToLower() == ".csv" ? line : line.Replace("\t", ","),
                                        FailPath = FailPath,
                                        // AC采集卡IN3, 216, 235, 211.83, V, Fail, 210, L<=x<=H
                                        IsHigh = TempHigh,
                                    });
                                }
                                else
                                {
                                    string TestName = $"{line.Split("\t")[0]}_{line.Split("\t")[1]}";
                                    string TestValues = (line.Split("\t")[3]).Replace(",", "").Trim();
                                    bool IsV = System.Text.RegularExpressions.Regex.IsMatch(TestValues, VerifyNumer);
                                    if (IsV)
                                    {

                                        PassReport? report = _failList.PassReport.FirstOrDefault<PassReport>(M => M.TestName == TestName);

                                        if (report is null)
                                        {

                                            PassReport passReport = new PassReport()
                                            {
                                                TestName = TestName,
                                                PassData = new List<double> { double.Parse(TestValues) }
                                            };

                                            _failList.PassReport.Add(passReport);
                                        }
                                        else
                                        {
                                            report.PassData.Add(double.Parse(TestValues));
                                        }

                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {

                    throw ex;
                }
            }

            // 读取所有包含Fail 的文件
            // 读取Fail 具体项目
            //添加到List


            //failList.Data = Data;

            // _failList.StrMes = "Passed:\t" + LogAll.Count() + "\nFailing:\t"
            //     + FailS.Count() + "\nSuccessRate:\t" + Math.Round((1 - (1.0 * FailS.Count() / LogAll.Count())) * 100, 2) + "%";
            // _failList.FailCount = _failList.Data.Count;

            _failList.StrMes = "Passed:\t" + PassS.Count()
                + "\nFailing:\t" + FailS.Count()
                + "\nSuccessRate:\t" + Math.Round((1 - (1.0 * FailS.Count() / LogAll.Count())) * 100, 2) + "%";
            //_failList.FailCount = _failList.Data.Count;
            _failList.GroupDatas = new List<GroupData>();


            //_failList.GroupDatas = (from P in _failList.Data
            //                        group P by P.FailItme.Split(",")[0] into ped
            //                        select new GroupData
            //                        {
            //                            itmeName = ped.Key,
            //                            itmeCount = ped.Count()
            //                        })
            //                        .OrderBy(m => m.itmeCount)
            //                        .ToList();


            _failList.GroupDatas = _failList.Data
                .GroupBy(m => $"{m.FailItme.Split(',')[0]},{m.FailItme.Split(',')[1]}")
                .Select(p => new GroupData { itmeCount = p.Count(), itmeName = p.Key })
                .OrderByDescending(m => m.itmeCount)
                .ToList();

            // 查询图表需要的数据 


            return System.Text.Json.JsonSerializer.Serialize(_failList);

        }


        /// <summary>
        /// 上传Excel文件到服务器
        /// </summary>
        /// <param name="file">要上传的Excel文件</param>
        /// <returns>上传结果</returns>
        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile file, [FromForm] string? Type)
        {
            try
            {


                // 检查文件是否为空
                if (file == null || file.Length == 0)
                {
                    return BadRequest("请选择要上传的Excel文件");
                }

                // 检查文件类型
                var allowedExtensions = new[] { ".xlsx", ".zip", ".csv" };
                var fileExtension = System.IO.Path.GetExtension(file.FileName).ToLower();
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new
                    {
                        message = "只允许上传Excel文件(.xlsx)或压缩包(.zip)"
                    });
                }

                // 检查文件大小（限制为20MB）
                if (file.Length > 20 * 1024 * 1024)
                {

                    return Ok(new
                    {
                        message = "文件大小超过限制（20MB）",
                        fileName = file.FileName,
                        fileSize = file.Length / 1024 + " KB"
                    });
                }

                // 调用服务上传文件
                var result = await _uploadService.SaveExcelFileAsync(file);

                if (result.Success)
                {
                    if (Type == "FCT" || Type == "ICT")
                    {
                        string res = await _uploadService.ReadExcelData2Csv(result.FilePath, Type);
                        if (res != null)
                        {

                            return Ok(new
                            {
                                message = res.Split(">")[0],
                                filefullPath = res.Split(">")[1].Trim(),
                            });
                        }
                        else
                        {
                            return StatusCode(500, new { message = "文件上传失败", error = "转换CSV文件失败" });
                        }

                    }
                    else
                    {

                        Stopwatch Ps = Stopwatch.StartNew();
                        // 若是Excel文件，则读取数据，zip文件则解压缩 
                        if (result.FilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            await _uploadService.ReadExcelData2Txt(result.FilePath);
                            Debug.WriteLine($"读取Excel数据耗时: {Ps.ElapsedMilliseconds} ms");
                        }
                        else if (result.FilePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) && result.FilePath.IndexOf("ModelConfig.csv") != -1)
                        {
                            // 运行命令行进行发布 


                            // dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish
                            string projectPath = @"C:\TE\N51876BTestApp\"; // 项目文件路径
                            string outputPath = Path.Combine("C:\\TE", "publish"); // 输出目录

                            // 构建dotnet publish命令参数
                            string arguments = $"publish " +
                                              $"-c Release " +
                                              $"-r win-x64 " +
                                              $"--self-contained true " +
                                              $"-p:PublishSingleFile=true " +
                                              $"-p:IncludeNativeLibrariesForSelfExtract=true " +
                                              $"-o \"{outputPath}\"";

                            // 配置进程信息
                            ProcessStartInfo startInfo = new ProcessStartInfo
                            {
                                FileName = "dotnet", // 调用dotnet CLI
                                Arguments = arguments,
                                WorkingDirectory = Path.GetDirectoryName(projectPath), // 项目所在目录
                                RedirectStandardOutput = true, // 捕获输出
                                RedirectStandardError = true,  // 捕获错误
                                UseShellExecute = false,       // 不使用shell
                                CreateNoWindow = true,     // 不显示命令窗口
                                                           // 关键：设置输出编码为 UTF-8
                                StandardOutputEncoding = System.Text.Encoding.UTF8,
                                StandardErrorEncoding = System.Text.Encoding.UTF8
                            };
                            UpdateLog();
                            // 执行命令
                            using (Process process = Process.Start(startInfo))
                            {
                                if (process == null)
                                {
                                    //return "发布失败：无法启动dotnet进程";
                                }

                                // 读取输出和错误信息
                                string output = process.StandardOutput.ReadToEnd();
                                string error = process.StandardError.ReadToEnd();
                                process.WaitForExit(5000);

                                // 返回执行结果
                                if (process.ExitCode == 0)
                                {
                                    //return $"发布成功！输出目录：{outputPath}\n{output}";
                                    // 返回给用户指定的二进制文件
                                    return Ok(new
                                    {
                                        message = output,
                                        filefullPath = result.FilePath,
                                        fileName = System.IO.Path.GetFileName(result.FilePath),
                                        filePath = result.FilePath.Substring(0, result.FilePath.LastIndexOf('\\')),
                                        // 不要文件后缀名
                                        Select = System.IO.Path.GetFileNameWithoutExtension(result.FilePath)
                                    });
                                }
                                else
                                {
                                    //return $"发布失败（代码：{process.ExitCode}）\n错误信息：{error}\n输出：{output}";
                                }
                            }
                        }
                        else if (result.FilePath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                        {
                            // 如果是zip文件，则解压缩
                            await _uploadService.UnZipFile(result.FilePath);
                            Debug.WriteLine($"解压缩文件耗时: {Ps.ElapsedMilliseconds} ms");
                        }
                    }
                    return Ok(new
                    {
                        message = "文件上传成功",
                        filefullPath = result.FilePath,
                        fileName = System.IO.Path.GetFileName(result.FilePath),
                        filePath = result.FilePath.Substring(0, result.FilePath.LastIndexOf('\\')),
                        // 不要文件后缀名
                        Select = System.IO.Path.GetFileNameWithoutExtension(result.FilePath)
                    });
                }
                else
                {
                    return StatusCode(500, new { message = "文件上传失败", error = result.ErrorMessage });
                }
            }
            catch (Exception ex)
            {

                System.IO.File.AppendAllText("C:\\TE\\FailReport_ErrorLog.txt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.Message + "\n");
                return StatusCode(500, new { message = "文件上传过程中发生异常", error = ex.Message });
            }
        }

        // 返回编译后的二进制文件
        [HttpGet]
        public IActionResult DownloadPublishedFile()
        {
            string outputPath = Path.Combine("C:\\TE", "publish"); // 输出目录
            string exeFilePath = Path.Combine(outputPath, "N51876BTestApp.exe");
            byte[] fileBytes = System.IO.File.ReadAllBytes(exeFilePath);
            return File(fileBytes, "application/octet-stream", "N51876BTestApp.exe");
        }

        [HttpGet]
        public IActionResult DownloadCsvDataFile(string FilePath)
        {
            if (FilePath.Contains("TE") && FilePath.EndsWith(".csv", StringComparison.InvariantCultureIgnoreCase))
            {
                if (System.IO.File.Exists(FilePath))
                {
                    //"C:\Logs_Uploads\5f89991f_1_ICT原始数据.xlsx"
                    // 获取文件名称不要路径
                    string fileName = System.IO.Path.GetFileName(FilePath);

                    byte[] fileBytes = System.IO.File.ReadAllBytes(FilePath);

                    return File(fileBytes, "application/octet-stream", fileName);
                }
                else
                {
                    return BadRequest("文件不存在");
                }

            }
            else
            {
                return BadRequest("文件路径不合法");
            }
        }

        private void UpdateLog()
        {
            string filePath = "C:\\TE\\N51876BTestApp\\N51876BTestApp\\Form2.Designer.cs";

            // 读取文件内容
            string oregon = System.IO.File.ReadAllText(filePath);

            // 正则表达式：匹配 Text = "Vx.y.z"; 格式，捕获主版本、次版本、修订号
            string pattern = @"Text\s*=\s*""V(\d+)\.(\d+)\.(\d+)"";";

            // 替换回调：递增修订号（最后一段）
            string updatedContent = Regex.Replace(oregon, pattern, match =>
            {
                // 捕获组1：主版本（如1），组2：次版本（如4），组3：修订号（如1）
                int major = int.Parse(match.Groups[1].Value);
                int minor = int.Parse(match.Groups[2].Value);
                int patch = int.Parse(match.Groups[3].Value);

                // 修订号加1（如1 → 2）
                patch++;

                // 重构版本号字符串
                return $"Text = \"V{major}.{minor}.{patch}\";";
            });

            // 将修改后的内容写回文件
            System.IO.File.WriteAllText(filePath, updatedContent);
            Console.WriteLine("版本号已更新");
        }


        [HttpGet]
        public string SearchCodeLog(string Code = "GCETE-TM-012")
        {
            string res = System.IO.File.ReadAllText(@"C:\TE\TestBoxLogs\" + Code + ".txt", Encoding.UTF8); // 读取文件内容
            return res;
        }
    }


    public class FailData
    {
        public string FailItme { get; set; } = string.Empty;
        public string FailPath { get; set; } = string.Empty;
        public DateTime FailDate { get; set; }
        public bool? IsHigh { get; set; }
    }


    public class FailList
    {
        //   public int FailCount { get; set { if (value < 0) { throw new ArgumentOutOfRangeException("FailCount", "不能设置负数"); } else { FailCount = value; } } }

        public List<FailData> Data { get; set; }

        public string StrMes { get; set; }
        public List<GroupData> GroupDatas { get; set; }
        public List<PassReport> PassReport { get; set; }


    }



    public class GroupData
    {
        public string itmeName { get; set; } = string.Empty;
        public int itmeCount { get; set; }
    }


    public class PassReport
    {/// <summary>
     /// 测试项目名称
     /// </summary>
        public string TestName { get; set; }

        public List<double> PassData { get; set; }
    }
}
