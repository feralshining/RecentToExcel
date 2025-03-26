using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using IWshRuntimeLibrary;
using ClosedXML.Excel;

namespace RecentToExcel
{
    public partial class Form1 : Form
    {
        public Form1() => InitializeComponent();
        static string GetShortcutTargetPath(string shortcutPath)
        {
            var shell = new WshShell();
            var shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath); // 바로 가기 로드
            return shortcut.TargetPath; // 대상 경로 반환
        }

        static string GetSaveFilePath()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "엑셀 파일 저장 경로 선택";
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (saveFileDialog.ShowDialog() == DialogResult.OK) return saveFileDialog.FileName;
            }
            return null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string recentFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                               @"Microsoft\Windows\Recent");

            // 최근 항목 가져오기 (.lnk 대상 경로 포함)
            var recentFiles = new DirectoryInfo(recentFolder)
                .GetFiles("*.lnk")
                .Select(file => new
                {
                    FileName = file.Name,
                    TargetPath = GetShortcutTargetPath(file.FullName), // 바로 가기 대상 경로
                LastWriteTime = file.LastWriteTime
                })
                .Where(file => !string.IsNullOrEmpty(file.TargetPath)) // 유효한 대상 경로만 포함
                .ToList();

            string excelFilePath = GetSaveFilePath();
            
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("파일 저장 경로를 선택하지 않았습니다. 프로그램을 종료합니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Recent Files");

                // 첫 줄 필터용 헤더 작성
                worksheet.Cell(1, 1).Value = "File Name";
                worksheet.Cell(1, 2).Value = "Target Path";
                worksheet.Cell(1, 3).Value = "Last Modified Date";

                for (int i = 0; i < recentFiles.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = recentFiles[i].FileName;
                    worksheet.Cell(i + 2, 2).Value = recentFiles[i].TargetPath;
                    worksheet.Cell(i + 2, 3).Value = recentFiles[i].LastWriteTime;
                }

                worksheet.Columns().AdjustToContents(); // 열 크기 자동 조정
                workbook.SaveAs(excelFilePath);
            }
            MessageBox.Show("파일이 성공적으로 지정한 경로에 저장되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
