namespace Apurisk.Application.Excel
{
    public interface IExcelWorkbookGateway
    {
        void EnsureSheets(SheetDefinition[] sheets);
        void ActivateSheet(string sheetName);
        void ShowMessage(string title, string message);
    }
}
