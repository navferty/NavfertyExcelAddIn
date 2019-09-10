namespace NavfertyExcelAddIn.Commons
{
    public interface IDialogService
    {
        void ShowError(string message);
        void ShowInfo(string message);
        bool Ask(string message, string caption);
        void ShowVersion();
    }
}
