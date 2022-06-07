using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms.VisualStyles;

namespace Solidworks_PDM_Specification
{
    internal class ReferenceFiles
    {
        public DrawingStamp stamp;
        private Settings settings;
        private string path;
        private List<Element> Elements;
        private IEdmVault5 vault;
        public ReferenceFiles(string path, Settings settings, DrawingStamp stamp, IEdmVault5 vault)
        {
            this.path = path;
            this.settings = settings;
            this.stamp = stamp;
            this.vault = vault;
        }
        
        public List<Element> GetListElements()
        {
            Elements = new List<Element>();
            GetReferencedFiles(null, path, 0, "VKRB", ref Elements, stamp.Configuration);
            stamp.usePdmFlag = true;
            return Elements;
        }

        private void GetReferencedFiles(IEdmReference10 Reference, string FilePath, int Level, string ProjectName,
            ref List<Element> Elements, string configuration)
        {
            try
            {
                if (Reference == null)
                {
                    IEdmFile5 File = null;
                    IEdmFolder5 ParentFolder = null;
                    File = vault.GetFileFromPath(FilePath, out ParentFolder);
                    GetStamp(File);
                    stamp.FilePath = FilePath;
                    Reference = (IEdmReference10) File.GetReferenceTree(ParentFolder.ID);
                    GetReferencedFiles(Reference, "", Level + 1, ProjectName, ref Elements, configuration);
                }
                else
                {
                    IEdmPos5 pos = default(IEdmPos5);
                    pos = Reference.GetFirstChildPosition3(ProjectName, true, true, (int) EdmRefFlags.EdmRef_File,
                        stamp.Configuration, 0);
                    IEdmReference10 @ref = default(IEdmReference10);
                    IEdmFile5 File = null;
                    IEdmFolder5 ParentFolder = null;
                    while ((!pos.IsNull))
                    {
                        @ref = (IEdmReference10) Reference.GetNextChild(pos);
                        File = vault.GetFileFromPath(@ref.FoundPath, out ParentFolder);
                        Elements.Add(GetElementFromVault(File, @ref.RefConfiguration));
                        Elements[Elements.Count - 1].Count = @ref.RefCount;
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GetStamp(IEdmFile5 File)
        {
            IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
            EnumVarObj = (IEdmEnumeratorVariable8)File.GetEnumeratorVariable();
            object currentVar = null;

            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Name"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Name = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Designation"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Designation = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["DrawingPaperSize"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.DrawingPaperSize = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Section"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Section = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Note"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Note = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Zone"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Zone = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Developer"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Developer = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Checker"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Checker = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["NormativeControl"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.NormativeControl = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Approver"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Approver = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Litera"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Litera = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["InvNumbOrigin"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.InvNumbOrigin = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["ReplacedInvNumb"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.ReplacedInvNumb = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["InvNumbDupl"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.InvNumbDupl = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["ReferenceNumb"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.ReferenceNumb = currentVar.ToString().Trim();
            if (currentVar != null)
                EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Name"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Name = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Designation"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Designation = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["DrawingPaperSize"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.DrawingPaperSize = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Section"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Section = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Note"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Note = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Zone"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.Zone = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["PrimaryApplication"], stamp.Configuration, out currentVar);
            if (currentVar != null)
                stamp.PrimaryApplication = currentVar.ToString().Trim();
            
            EnumVarObj.CloseFile(false);
        }

        private Element GetElementFromVault(IEdmFile5 File, string Configuration)
        {
            Element element = new Element();
            IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
            EnumVarObj = (IEdmEnumeratorVariable8)File.GetEnumeratorVariable();
            object currentVar = null;
            #region GetVar
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Name"], Configuration, out currentVar);
            if (currentVar != null)
                element.Name = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Designation"], Configuration, out currentVar);
            if (currentVar != null)
                element.Designation = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["DrawingPaperSize"], Configuration, out currentVar);
            if (currentVar != null)
                element.DrawingPaperSize = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Section"], Configuration, out currentVar);
            if (currentVar != null)
                element.Section = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Note"], Configuration, out currentVar);
            if (currentVar != null)
                element.Note = currentVar.ToString().Trim();
            EnumVarObj.GetVar(settings.ComparsionGlobalVariable["Zone"], Configuration, out currentVar);
            if (currentVar != null)
                element.Zone = currentVar.ToString().Trim();
            #endregion
            EnumVarObj.CloseFile(false);
            //MessageBox.Show(element.Section);
            if (element.Name.Trim() == "" & element.Designation.Trim() == "")
                element.Name = "Имя отсутствует";
            if (element.Section.Trim() == "")
                element.Section = "Прочие изделия";
            return element;
        }

    }
}
