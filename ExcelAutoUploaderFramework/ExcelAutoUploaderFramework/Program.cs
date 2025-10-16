using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using CloudinaryDotNet;
using CloudinaryDotNet.Actions;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static Cloudinary cloudinary;
    static Dictionary<string, DateTime> uploadedToday = new Dictionary<string, DateTime>();
    static DateTime lastUploadDate = DateTime.MinValue;

    static string logPath = @"D:\ExcelUploadLog.txt";
    static string uploadedListPath = @"D:\ExcelUploadedToday.txt"; // file lưu danh sách file đã upload

    // Thư mục hay dùng trên ổ C
    static string[] commonFoldersC = new string[]
    {
        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")
    };

    // Ổ đĩa cần quét toàn bộ
    static string[] drivesToScan = new string[] { @"D:\" };

    static void Main()
    {
        // Cloudinary config
        Account account = new Account("dqen9tax2", "617457879674455", "rZUf3gKI_53g4akmCwdK11WR15k");
        cloudinary = new Cloudinary(account);

        LoadUploadedList();

        Log("🚀 Excel Auto Uploader started.");

        while (true)
        {
            try
            {
                // Reset danh sách nếu sang ngày mới
                if (lastUploadDate.Date != DateTime.Today)
                {
                    uploadedToday.Clear();
                    File.WriteAllText(uploadedListPath, ""); // xóa file TXT
                    lastUploadDate = DateTime.Today;
                    Log("🔄 New day - starting daily scan.");

                    // Quét và upload file Excel
                    ScanAndUploadExcelFiles();

                    // Upload các file Excel đang mở
                    UploadOpenExcelFiles();
                    Log("upload done");
                }
            }
            catch (Exception ex)
            {
                Log($"🔥 Error in main loop: {ex.Message}");
            }

            // Sleep 1h, kiểm tra xem có file Excel đang mở để upload thêm
            System.Threading.Thread.Sleep(60 * 60 * 1000);
        }
    }

    // Load danh sách file đã upload hôm nay từ file TXT
    static void LoadUploadedList()
    {
        try
        {
            if (File.Exists(uploadedListPath))
            {
                var lines = File.ReadAllLines(uploadedListPath);
                foreach (var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        uploadedToday[line] = DateTime.Today;
                }
                Log($"ℹ️ Loaded {uploadedToday.Count} file(s) from uploaded list.");
            }
        }
        catch (Exception ex)
        {
            Log($"⚠️ Failed to load uploaded list: {ex.Message}");
        }
    }

    // Quét và upload Excel từ ổ C/D
    static void ScanAndUploadExcelFiles()
    {
        Log("upload C");
        // Quét ổ C các thư mục hay dùng
        foreach (var folder in commonFoldersC)
            ScanFolderAndUpload(folder);

        Log("upload D");
        // Quét ổ D toàn bộ
        foreach (var drive in drivesToScan)
            ScanFolderAndUpload(drive);
    }

    static void ScanFolderAndUpload(string folderPath)
    {
        try
        {
            if (!Directory.Exists(folderPath)) return;
            // EnumerateFiles để không load tất cả vào RAM cùng lúc
            foreach (var file in Directory.EnumerateFiles(folderPath, "*.xls*", SearchOption.AllDirectories))
            {
                if (Path.GetFileName(file).StartsWith("~$")) continue;
                if (!uploadedToday.ContainsKey(file))
                {
                    if (UploadToCloudinary(file))
                    {
                        uploadedToday[file] = DateTime.Today;
                        File.AppendAllText(uploadedListPath, file + Environment.NewLine);
                    }
                }
            }
        }
        catch (UnauthorizedAccessException)
        {
            // Bỏ qua folder không có quyền
        }
        catch (Exception ex)
        {
            Log($"⚠️ Scan folder error: {folderPath} - {ex.Message}");
        }
    }

    // Upload file Excel đang mở trong Excel
    static void UploadOpenExcelFiles()
    {
        try
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Log("🔎 Excel COM instance found.");

            foreach (Excel.Workbook wb in excelApp.Workbooks)
            {
                string filePath = wb.FullName;
                if (string.IsNullOrEmpty(filePath)) continue;

                if (!uploadedToday.ContainsKey(filePath))
                {
                    if (UploadToCloudinary(filePath))
                    {
                        uploadedToday[filePath] = DateTime.Today;
                        File.AppendAllText(uploadedListPath, filePath + Environment.NewLine);
                    }
                }
                else
                {
                    Log($"⏭ Already uploaded today: {filePath}");
                }
            }
        }
        catch (COMException)
        {
            Log("ℹ️ No Excel COM instance found. Excel may not be open.");
        }
        catch (Exception ex)
        {
            Log($"⚠️ Error checking open Excel: {ex.Message}");
        }
    }

    // Upload file lên Cloudinary
    static bool UploadToCloudinary(string filePath)
    {
        string tempFile = Path.Combine(Path.GetTempPath(), Path.GetFileName(filePath));
        try
        {
            File.Copy(filePath, tempFile, true); // Copy file tạm tránh lock

            var uploadParams = new RawUploadParams()
            {
                File = new FileDescription(tempFile),
                PublicId = $"excel_{Path.GetFileNameWithoutExtension(filePath)}",
                Overwrite = true,
            };

            var result = cloudinary.Upload(uploadParams);
            Log($"✅ Uploaded: {filePath} -> {result.SecureUrl}");
            return true;
        }
        catch (Exception ex)
        {
            Log($"❌ Upload failed: {filePath} - {ex.Message}");
            return false;
        }
        finally
        {
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    // Log message ra console + file
    static void Log(string message)
    {
        string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}";
        Console.WriteLine(logMessage);

        try
        {
            File.AppendAllText(logPath, logMessage + Environment.NewLine);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️ Failed to write log: {ex.Message}");
        }
    }
}
