using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Solidworks_PDM_Specification;

namespace SolidWorksAddIn
{
    [Guid("0492bc13-07b8-49fd-8e01-94c6f11d0fee"), ComVisible(true)]
    public class SolidWorksAddIn : IEdmAddIn5
    {
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            poInfo.mbsAddInName = "Минимальная версия необходимая версия для PDM Professional это 6.4";
            poInfo.mlRequiredVersionMajor = 6;
            poInfo.mlRequiredVersionMinor = 4;
            poInfo.mbsCompany = "МАИ 316 Кафедра";
            poInfo.mlAddInVersion = 1;
            poCmdMgr.AddCmd(1001, "Создать спецификацию автоматически", (int)EdmMenuFlags.EdmMenu_OnlyFiles);
            poCmdMgr.AddCmd(1000, "Создать спецификацию", (int)EdmMenuFlags.EdmMenu_OnlyFiles);
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            IEdmVault5 vault = (IEdmVault5)poCmd.mpoVault;
            IEdmFile5 file;
            file = (IEdmFile5) vault.GetObject(EdmObjectType.EdmObject_File, ppoData[0].mlObjectID1);
            string path = file.GetLocalPath(poCmd.mlCurrentFolderID);
            if (path.Substring(path.Length - 7) != ".SLDASM")
            {
                MessageBox.Show("Выберите сборку");
                return;
            }
            Settings settings;
            string AffectedFileNames = "";
            XML_Convert xml = new XML_Convert();
            xml.Import(out settings, vault.RootFolderPath + "\\Настройки для спецификации\\Настройки\\Settings.xml");
            settings.excelTemplate = vault.RootFolderPath + "\\Настройки для спецификации\\Настройки\\Specification.xltx";
            GetConfigurationForm configureForm = new GetConfigurationForm(vault, path);
            configureForm.ShowDialog();
            DrawingStamp stamp = new DrawingStamp();
            if (!configureForm.flag)
            {
                configureForm.Dispose();
                return;
            }
            configureForm.Dispose();
            
            stamp.Configuration = configureForm.configuration;
             
            if (poCmd.meCmdType == EdmCmdType.EdmCmd_Menu)
            {
                switch (poCmd.mlCmdID)
                {
                    case 1000:
                        SpecificationForm form = new SpecificationForm(path, vault, settings, stamp);
                        form.ShowDialog();
                        form.Dispose();
                        break;
                    case 1001:
                        List<Element> elements = new List<Element>();
                        file = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, ppoData[0].mlObjectID1);
                        ReferenceFiles reference = new ReferenceFiles(path, settings, stamp, vault);
                        elements = reference.GetListElements();
                        SpecificationForm specForm = new SpecificationForm(path, vault, settings, stamp);
                        specForm.exportToExcel(elements);
                        specForm.Dispose();
                        break;
                    default:
                        break;
                }

            }
        }

    }

}

