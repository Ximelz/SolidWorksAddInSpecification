using Aspose.Cells;
using EPDM.Interop.epdm;
using Solidworks_PDM_Specification;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace SolidWorksAddIn
{
    public partial class SpecificationForm : Form
    {
        Dictionary<int, int> ColumnMatching = new Dictionary<int, int>();
        private IEdmVault5 vault;
        private string path;
        private Settings settings;
        private DrawingStamp stamp;
        private List<Element> elements = new List<Element>();
        private string specificationExcelPath;
        private Dictionary<string, int> idSection = new Dictionary<string, int>()
        {
            {"Документация", 0}, {"Комплексы", 0},
            { "Сборочные единицы", 0}, { "Детали", 0},
            { "Стандартные изделия", 0}, { "Прочие изделия", 0},
            { "Материалы", 0}, { "Комплекты", 0}
        };
        private Dictionary<string, List<Element>> nameSections;
        public SpecificationForm(string path, IEdmVault5 vault, Settings settings, DrawingStamp stamp)
        {
            InitializeComponent();
            this.vault = vault;
            this.path = path;
            this.settings = settings;
            this.stamp = stamp;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CreateSpec();
        }

        private void CreateSpec()
        {
            ReferenceFiles reference = new ReferenceFiles(path, settings, stamp, vault);
            elements = reference.GetListElements();
            nameSections = new Dictionary<string, List<Element>>()
            {
                {"Документация", new List<Element>()}, {"Комплексы", new List<Element>()},
                {"Сборочные единицы", new List<Element>()}, {"Детали", new List<Element>()},
                {"Стандартные изделия", new List<Element>()}, {"Прочие изделия", new List<Element>()},
                {"Материалы", new List<Element>()}, {"Комплекты", new List<Element>()}
            };
            int count = 0;
            foreach (Element element in elements)
                if (nameSections.ContainsKey(element.Section))
                {
                    nameSections[element.Section].Add(element);
                    count++;
                }
            AddRows(nameSections, count);
        }

        private void AddRows(Dictionary<string, List<Element>> nameSections, int count)
        {
            int i = 0;
            dataGridView1.Rows.AddCopy(dataGridView1.Rows.Count - 1);
            foreach (KeyValuePair<string, List<Element>> keyValuePair in nameSections)
            {
                idSection[keyValuePair.Key] = dataGridView1.Rows.Count - 2;
                dataGridView1.Rows[dataGridView1.Rows.Count - 2]
                    .SetValues("", "", "", keyValuePair.Key);
                dataGridView1.Rows.AddCopy(dataGridView1.Rows.Count - 2);
                if (keyValuePair.Value.Count > 0)
                {
                    foreach (Element element in keyValuePair.Value)
                        AddNewRow(element);

                }
                i++;
            }
        }

        private void AddNewRow(Element element)
        {
            dataGridView1.Rows[dataGridView1.Rows.Count - 2].SetValues(element.DrawingPaperSize, element.Zone, element.Designation, element.Name, element.Count.ToString(), element.Note);
            dataGridView1.Rows.AddCopy(dataGridView1.Rows.Count - 2);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
        }

        private void UpdateSpecificationButton_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            CreateSpec();
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            int index = dataGridView1.CurrentRow.Index;
            dataGridView1.Rows.InsertCopy(index, index + 1);
        }

        private void exportToExcelButton_Click(object sender, EventArgs e)
        {
            exportToExcel(GetNewListElements());
            Close();
        }

        private void AddElementFromDataGrid(ref List<Element> Elements, int i, string section)
        {
            int currentIndex = Elements.Count;
            Elements.Add(new Element());
            Elements[currentIndex].DrawingPaperSize = dataGridView1.Rows[i].Cells[0].Value.ToString();
            Elements[currentIndex].Zone = dataGridView1.Rows[i].Cells[1].Value.ToString();
            Elements[currentIndex].Designation = dataGridView1.Rows[i].Cells[2].Value.ToString();
            Elements[currentIndex].Name = dataGridView1.Rows[i].Cells[3].Value.ToString();
            if (int.TryParse(dataGridView1.Rows[i].Cells[4].Value.ToString(), out int count))
                Elements[currentIndex].Count = count;
            else
                Elements[currentIndex].Count = 0;
            Elements[currentIndex].Note = dataGridView1.Rows[i].Cells[5].Value.ToString();
            Elements[currentIndex].Section = section;
        }

        private List<Element> GetNewListElements()
        {
            MessageBox.Show(settings.excelTemplate);
            List<Element> Elements = new List<Element>();
            int[] sections = new int[8];
            string[] nameSections = new string[8];
            int i = 0;
            foreach (KeyValuePair<string, int> keyValuePair in idSection)
            {
                sections[i] = keyValuePair.Value;
                nameSections[i] = keyValuePair.Key;
                i++;
            }
            i = 1;
            int rowsCount = dataGridView1.Rows.Count - 2;
            for (int j = 1; j < rowsCount; j++)
                if (i < 9)
                {
                    if (j < sections[i])
                    {
                        AddElementFromDataGrid(ref Elements, j, nameSections[i - 1]);
                    }
                    else
                    {
                        i++;
                    }
                }

            return Elements;
        }

        public void exportToExcel(List<Element> Elements)
        {
            MessageBox.Show(settings.excelTemplate);
            Dictionary<string, List<Element>> nameSections = new Dictionary<string, List<Element>>()
            {
                {"Документация", new List<Element>()}, {"Комплексы", new List<Element>()},
                {"Сборочные единицы", new List<Element>()}, {"Детали", new List<Element>()},
                {"Стандартные изделия", new List<Element>()}, {"Прочие изделия", new List<Element>()},
                {"Материалы", new List<Element>()}, {"Комплекты", new List<Element>()}
            };
            foreach (Element element in Elements)
                if (nameSections.ContainsKey(element.Section))
                    nameSections[element.Section].Add(element);
            using (var workbook = new Workbook(settings.excelTemplate))
            {
                Worksheet ws = workbook.Worksheets["Specification_sheet_1"];
                Cell cell;
                int cellIndex = 5;
                int position = 1;
                int maxCellIndex = 61;
                int Lists = 1;
                SetStamp(ws, Lists);
                foreach (KeyValuePair<string, List<Element>> keyValuePair in nameSections)
                {
                    if (keyValuePair.Value.Count > 0)
                    {
                        cellIndex += 2;
                        cell = ws.Cells["AI" + cellIndex];
                        cell.PutValue(keyValuePair.Key);
                        Style style = cell.GetStyle();
                        style.HorizontalAlignment = TextAlignmentType.Center;
                        style.Font.Underline = FontUnderlineType.Single;
                        style.Font.IsBold = true;
                        cell.SetStyle(style);
                        cellIndex += 4;
                        int listCount = keyValuePair.Value.Count;
                        AddElementsOnSpecificationForm(ref ws, ref cellIndex, maxCellIndex, ref listCount, ref position,
                                                        keyValuePair.Value, ref Lists, workbook);
                    }
                }
                ws = workbook.Worksheets["Specification_sheet_1"];
                cell = ws.Cells["BH67"];
                cell.PutValue(Lists);
                WorksheetCollection wc = workbook.Worksheets;
                wc.RemoveAt(wc.Count - 1);
                saveFileDialog1.Filter = "Xls files(*.xls)|*.xls|All files(*.*)|*.*";
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                    return;
                string saveFileName = saveFileDialog1.FileName;
                string addFileFormat = saveFileName.Substring(saveFileName.Length - 4) == ".xls" ? "" : ".xls";
                specificationExcelPath = saveFileName + addFileFormat;
                try
                {
                    workbook.Save(specificationExcelPath);
                    Process.Start(specificationExcelPath);
                }
                catch
                {
                    MessageBox.Show("Файл открыт в другой программе");
                }
            } 
        }

        private void AddElementsOnSpecificationForm(ref Worksheet worksheet, ref int cellIndex, int maxIndex, ref int listCount, ref int position, List<Element> elements, ref int Lists, Workbook workbook)
        {
            Cell cell;
            int i = 0;
            while (listCount > i)
            {
                if (elements[i].Count / 27 + (elements[i].Count % 27 != 0 ? 1 : 0) > (maxIndex - cellIndex) / 2)
                {
                    WorksheetCollection wc = workbook.Worksheets;
                    Lists++;
                    wc.Add("Specification_sheet_" + (Lists + 1));
                    maxIndex = 67;
                    wc[wc.Count - 1].Copy(wc[wc.Count - 2]);
                    worksheet = workbook.Worksheets["Specification_sheet_" + Lists];
                    cellIndex = 7;
                    SetStamp(worksheet, Lists);
                }

                cell = worksheet.Cells["E" + cellIndex];
                cell.PutValue(elements[i].DrawingPaperSize);

                cell = worksheet.Cells["G" + cellIndex];
                cell.PutValue(elements[i].Zone);

                cell = worksheet.Cells["I" + cellIndex];
                cell.PutValue(position.ToString());

                cell = worksheet.Cells["L" + cellIndex];
                cell.PutValue(elements[i].Designation);

                cell = worksheet.Cells["BD" + cellIndex];
                cell.PutValue(elements[i].Count.ToString());

                cell = worksheet.Cells["BG" + cellIndex];
                cell.PutValue(elements[i].Note);

                cell = worksheet.Cells["AI" + cellIndex];
                if (elements[i].Name.Length > 27)
                    SplitDesignation(worksheet, ref cellIndex, elements[i].Name, cell);
                else
                    cell.PutValue(elements[i].Name);

                position++;
                i++;
                cellIndex += 2;
            }
        }

        private void SplitDesignation(Worksheet worksheet, ref int cellIndex, string designation, Cell cell)
        {
            string[] designationSplit = designation.Split();
            int designationSplitLength = designationSplit.Length;
            designation = designationSplit[0];
            if (designationSplitLength == 1)
            {
                cell.PutValue(designation);
                cell = worksheet.Cells["AI" + cellIndex];
                return;
            }
            for (int i = 1; i < designationSplitLength; i++)
            {
                if (designationSplit[i].Length + designation.Length + 1 < 27)
                {
                    designation += " " + designationSplit[i];
                    if (i + 1 == designationSplitLength)
                    {
                        cell.PutValue(designation);
                        cell = worksheet.Cells["AI" + cellIndex];
                        break;
                    }
                }
                else
                {
                    cell.PutValue(designation);
                    cellIndex += 2;
                    cell = worksheet.Cells["AI" + cellIndex];
                    designation = designationSplit[i];
                    for (int j = i + 1; j < designationSplitLength; j++)
                        designation += " " + designationSplit[j];
                    SplitDesignation(worksheet, ref cellIndex, designation, cell);
                    break;
                }
            }
        }

        private void SetStamp(Worksheet worksheet, int Lists)
        {
            Cell cell;
            cell = worksheet.Cells["C66"];
            cell.PutValue(stamp.InvNumbOrigin);

            cell = worksheet.Cells["C52"];
            cell.PutValue(stamp.ReferenceNumb);

            cell = worksheet.Cells["C37"];
            cell.PutValue(stamp.InvNumbDupl);

            if (Lists == 1)
            {
                cell = worksheet.Cells["K66"];
                cell.PutValue(stamp.Developer);
                cell = worksheet.Cells["W66"];
                if (stamp.DateDeveloper.ToString("dd.MM.yyyy") != "01.01.0001")
                    cell.PutValue(stamp.DateDeveloper.ToString("dd.MM.yyyy"));

                cell = worksheet.Cells["K67"];
                cell.PutValue(stamp.Checker);
                cell = worksheet.Cells["W67"];
                if (stamp.DateChecker.ToString("dd.MM.yyyy") != "01.01.0001")
                    cell.PutValue(stamp.DateChecker.ToString("dd.MM.yyyy"));

                cell = worksheet.Cells["K69"];
                cell.PutValue(stamp.NormativeControl);
                cell = worksheet.Cells["W69"];
                if (stamp.DateNormativeControl.ToString("dd.MM.yyyy") != "01.01.0001")
                    cell.PutValue(stamp.DateNormativeControl.ToString("dd.MM.yyyy"));

                cell = worksheet.Cells["K70"];
                cell.PutValue(stamp.Approver);
                cell = worksheet.Cells["W70"];
                if (stamp.DateApprover.ToString("dd.MM.yyyy") != "01.01.0001")
                    cell.PutValue(stamp.DateApprover.ToString("dd.MM.yyyy"));

                cell = worksheet.Cells["Z66"];
                cell.PutValue(stamp.Name);

                cell = worksheet.Cells["Z63"];
                cell.PutValue(stamp.Designation);

                cell = worksheet.Cells["AY67"];
                cell.PutValue(stamp.Litera);

                cell = worksheet.Cells["AW68"];
                cell.PutValue("АО НИИТМ");
            }
            else
            {
                cell = worksheet.Cells["Z69"];
                cell.PutValue(stamp.Name);

                cell = worksheet.Cells["BK70"];
                cell.PutValue(Lists);
            }
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Xml files(*.xml)|*.xml|All files(*.*)|*.*";
            XML_Convert xml = new XML_Convert();
            DialogResult DialogSaveResult;
            DialogSaveResult = saveFileDialog1.ShowDialog();
            if (!(DialogSaveResult == DialogResult.OK))
                return;
            xml.Export(GetNewListElements(), stamp, saveFileDialog1.FileName);
        }
    }
}
