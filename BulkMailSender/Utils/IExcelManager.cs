
namespace BulkMailSender.Utils
{
    #region Namespace.
    #endregion

    public interface IExcelManager
    {
        string[] ReadFields(string fileName);

        string[][] Read(string fileName);

        void Insert(string[][] records, int sheetId, string srcFile);

        void Insert(string[][] records, string sheetName, string srcFile);

        void Write(string[][] records, string filePath);

        void Close();
    }
}
