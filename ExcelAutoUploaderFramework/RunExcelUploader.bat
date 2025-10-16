@echo off
:: =====================================================
:: Batch file để chạy ExcelUploader.exe ẩn console
:: =====================================================

:: Đường dẫn đến file exe, chỉnh theo nơi bạn đặt
set EXE_PATH="D:\ExcelUploader\ExcelUploader.exe"

:: Kiểm tra file exe tồn tại
if not exist %EXE_PATH% (
    echo ❌ File %EXE_PATH% không tồn tại!
    pause
    exit /b
)

:: Chạy exe ẩn console
start "" /min %EXE_PATH%

:: Thoát batch
exit
