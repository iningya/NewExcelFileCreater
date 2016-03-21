using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NewExcelFileCreater
{
    /// <summary>
    /// Excelファイル(.xlsx)の新規作成
    /// </summary>
    public class CreateNew
    {
        /// <summary>
        /// 生成するファイル
        /// </summary>
        private string FilePath;

        /// <summary>
        /// コンストラクタで府対象ファイルの設定
        /// </summary>
        /// <param name="filePath"></param>
        private CreateNew(string filePath)
        {
            this.FilePath = filePath;
        }

        /// <summary>
        /// 空のシートがあるファイルを生成。
        /// </summary>
        /// <param name="filePath"></param>
        public static void CreateEmptyFile(string filePath)
        {
            new CreateNew(filePath).Execute();
        }

        /// <summary>
        /// ファイル生成の実行
        /// </summary>
        private void Execute()
        {
            using (var sd = SpreadsheetDocument.Create(this.FilePath, SpreadsheetDocumentType.Workbook))
            {
                var wbp = sd.AddWorkbookPart();
                wbp.Workbook = new Workbook();

                var wsp = wbp.AddNewPart<WorksheetPart>();
                wsp.Worksheet = new Worksheet(new SheetData());

                var sheets = sd.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                //空のシート Sheet1 を追加
                var sheet = new Sheet()
                {
                    Id = sd.WorkbookPart.GetIdOfPart(wsp),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                //保存
                wbp.Workbook.Save();
                //閉じる
                sd.Close();
            }


        }

    }
}
