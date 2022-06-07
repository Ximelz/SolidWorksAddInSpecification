using System.Collections.Generic;
using System.IO;

namespace Solidworks_PDM_Specification
{
    /// <summary>
    /// Класс хранящий изначальные настройки программы.
    /// </summary>
    public class Settings
    {
        public Settings()
        {
            ComparsionGlobalVariable = new Dictionary<string, string>();
            Vault = "";
            excelTemplate = Directory.GetCurrentDirectory() + "\\Specification.xltx";
            InitialDictionary();
        }

        public string Vault { get; set; }                                   //Переменная хранящая исходное хранилище
        public Dictionary<string, string> ComparsionGlobalVariable;         //Сопоставление переменных из программы глобальных переменным SolidWorks PDM
        public string excelTemplate { set; get; }

        private void InitialDictionary()
        {
            ComparsionGlobalVariable.Add("Name", "");
            ComparsionGlobalVariable.Add("Designation", "");
            ComparsionGlobalVariable.Add("DrawingPaperSize", "");
            ComparsionGlobalVariable.Add("Section", "");
            ComparsionGlobalVariable.Add("Note", "");
            ComparsionGlobalVariable.Add("Zone", "");
            ComparsionGlobalVariable.Add("Developer", "");
            ComparsionGlobalVariable.Add("Checker", "");
            ComparsionGlobalVariable.Add("NormativeControl", "");
            ComparsionGlobalVariable.Add("Approver", "");
            ComparsionGlobalVariable.Add("Litera", "");
            ComparsionGlobalVariable.Add("InvNumbOrigin", "");
            ComparsionGlobalVariable.Add("ReplacedInvNumb", "");
            ComparsionGlobalVariable.Add("InvNumbDupl", "");
            ComparsionGlobalVariable.Add("ReferenceNumb", "");
            ComparsionGlobalVariable.Add("PrimaryApplication", "");
        }
    }
}
