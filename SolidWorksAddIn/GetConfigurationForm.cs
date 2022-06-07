using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPDM.Interop.epdm;

namespace SolidWorksAddIn
{
    public partial class GetConfigurationForm : Form
    {
        private IEdmVault7 vault;
        public bool flag = false;
        public string configuration;
        private string path;
        public GetConfigurationForm(IEdmVault5 vault, string path)
        {
            InitializeComponent();
            this.vault = (IEdmVault7)vault;
            this.path = path;
            GetConfigurationsFile();
        }

        private void GetConfigurationsFile()
        {
            IEdmFile17 file;
            IEdmFolder5 ppoRetParentFolder;
            EdmStrLst5 cfgList = default(EdmStrLst5);
            IEdmPos5 pos = default(IEdmPos5);
            string cfgName = null;
            file = (IEdmFile17)vault.GetFileFromPath(path, out ppoRetParentFolder);
            cfgList = file.GetConfigurations();
            pos = cfgList.GetHeadPosition();
            while (!pos.IsNull)
            {
                cfgName = cfgList.GetNext(pos);
                if (cfgName != "@")
                    configurationComboBox.Items.Add(cfgName);
            }
            configurationComboBox.Text = configurationComboBox.Items[0].ToString();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            configuration = configurationComboBox.Text;
            flag = true;
            Close();
        }

        private void GetConfigurationForm_Load(object sender, EventArgs e)
        {

        }
    }
}
