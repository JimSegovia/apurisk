using System;
using System.Runtime.InteropServices;
using System.Xml;

namespace Apurisk.ExcelAddIn.Excel
{
    public sealed class ExcelWorkbookGateway
    {
        private readonly object _excelApplication;

        public ExcelWorkbookGateway(object excelApplication)
        {
            _excelApplication = excelApplication;
        }

        public bool HasActiveWorkbook
        {
            get
            {
                dynamic excel = _excelApplication;
                return excel.ActiveWorkbook != null;
            }
        }

        public void EnsureSheet(string sheetName, string[] headers)
        {
            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;

            if (workbook == null)
            {
                workbook = excel.Workbooks.Add();
            }

            EnsureSheetInternal(workbook, sheetName, headers);
        }

        public void ActivateSheet(string sheetName)
        {
            dynamic excel = _excelApplication;
            dynamic sheet = FindSheet(excel.ActiveWorkbook, sheetName);
            if (sheet != null)
            {
                sheet.Activate();
            }
        }

        public void ShowMessage(string title, string message)
        {
            dynamic excel = _excelApplication;
            excel.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.GetProperty, null, excel, null);
            System.Windows.Forms.MessageBox.Show(message, title,
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        public string PickRange(string prompt)
        {
            dynamic excel = _excelApplication;
            try
            {
                dynamic picked = excel.GetType().InvokeMember("InputBox",
                    System.Reflection.BindingFlags.InvokeMethod, null, excel,
                    new object[] { prompt, "Apurisk", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8 });

                if (picked == null)
                    return string.Empty;

                string address = picked.Address(true, true, 1, true);
                return address ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public string ReadConfigValue(string keyName)
        {
            try
            {
                string xml = GetConfigXml();
                if (string.IsNullOrEmpty(xml)) return string.Empty;

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xml);

                XmlNode node = doc.SelectSingleNode("//e[@k='" + XmlEscape(keyName) + "']");
                return node != null ? (node.Attributes["v"].Value ?? string.Empty) : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public int GetImpactFieldCount()
        {
            string val = ReadConfigValue("ImpactFieldCount");
            int count;
            if (!int.TryParse(val, out count) || count < 1)
                count = 1;
            return count;
        }

        public System.Collections.Generic.List<RbsRow> ReadRbsFromRanges()
        {
            var result = new System.Collections.Generic.List<RbsRow>();

            string codeAddr = ReadConfigValue("Field.RbsCodeRange");
            string nameAddr = ReadConfigValue("Field.RbsNameRange");

            if (string.IsNullOrEmpty(codeAddr))
                return result;

            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;
            if (workbook == null) return result;

            try
            {
                dynamic codeRange = workbook.Range(codeAddr);
                dynamic nameRange = string.IsNullOrEmpty(nameAddr) ? null : workbook.Range(nameAddr);

                long maxRows = codeRange.Rows.Count;

                for (long row = 1; row <= maxRows; row++)
                {
                    object codeObj = codeRange.Cells[row, 1].Value2;
                    string code = codeObj != null ? codeObj.ToString().Trim() : string.Empty;

                    if (!string.IsNullOrEmpty(code))
                    {
                        string name = string.Empty;
                        if (nameRange != null && row <= nameRange.Rows.Count)
                        {
                            object nameObj = nameRange.Cells[row, 1].Value2;
                            name = nameObj != null ? nameObj.ToString().Trim() : string.Empty;
                        }

                        result.Add(new RbsRow { Code = code, Name = name });
                    }
                }
            }
            catch { }

            return result;
        }

        public System.Collections.Generic.List<RiskRow> ReadRisksFromRanges()
        {
            var result = new System.Collections.Generic.List<RiskRow>();

            string idAddr = ReadConfigValue("Field.RiskIdRange");
            string rbsAddr = ReadConfigValue("Field.RiskRbsCodeRange");
            string descAddr = ReadConfigValue("Field.RiskDescriptionRange");

            if (string.IsNullOrEmpty(idAddr))
                return result;

            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;
            if (workbook == null) return result;

            try
            {
                dynamic idRange = workbook.Range(idAddr);
                dynamic rbsRange = string.IsNullOrEmpty(rbsAddr) ? null : workbook.Range(rbsAddr);
                dynamic descRange = string.IsNullOrEmpty(descAddr) ? null : workbook.Range(descAddr);

                long maxRows = idRange.Rows.Count;

                for (long row = 1; row <= maxRows; row++)
                {
                    object idObj = idRange.Cells[row, 1].Value2;
                    string id = idObj != null ? idObj.ToString().Trim() : string.Empty;

                    if (!string.IsNullOrEmpty(id))
                    {
                        string rbsCode = string.Empty;
                        if (rbsRange != null && row <= rbsRange.Rows.Count)
                        {
                            object rbsObj = rbsRange.Cells[row, 1].Value2;
                            rbsCode = rbsObj != null ? rbsObj.ToString().Trim() : string.Empty;
                        }

                        string desc = string.Empty;
                        if (descRange != null && row <= descRange.Rows.Count)
                        {
                            object descObj = descRange.Cells[row, 1].Value2;
                            desc = descObj != null ? descObj.ToString().Trim() : string.Empty;
                        }

                        result.Add(new RiskRow { Id = id, RbsCode = rbsCode, Description = desc });
                    }
                }
            }
            catch { }

            return result;
        }

        public void SaveAllConfig(Forms.BowTieIntakeForm form)
        {
            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;
            if (workbook == null) return;

            string[] fields =
            {
                "RbsNameRange", "RbsCodeRange", "RiskTableRange", "RiskIdRange", "RiskTopRange",
                "RiskRbsCodeRange", "RiskRbsNameRange", "RiskDescriptionRange", "RiskCauseRange",
                "RiskPotentialEffectRange", "RiskProbabilityRange", "RiskImpactRange", "RiskSeverityRange",
                "RiskMitigationRange", "RiskOwnerRange"
            };

            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("apurisk");
            doc.AppendChild(root);

            foreach (var field in fields)
            {
                string value = form.GetFieldValue(field) ?? string.Empty;
                XmlElement elem = doc.CreateElement("e");
                elem.SetAttribute("k", "Field." + field);
                elem.SetAttribute("v", value);
                root.AppendChild(elem);
            }

            string[] impactKeys = form.GetImpactFieldKeys();
            for (int i = 0; i < impactKeys.Length; i++)
            {
                string value = form.GetFieldValue(impactKeys[i]) ?? string.Empty;
                XmlElement elem = doc.CreateElement("e");
                elem.SetAttribute("k", "Field." + impactKeys[i]);
                elem.SetAttribute("v", value);
                root.AppendChild(elem);
            }

            XmlElement countElem = doc.CreateElement("e");
            countElem.SetAttribute("k", "ImpactFieldCount");
            countElem.SetAttribute("v", form.ImpactCount.ToString());
            root.AppendChild(countElem);

            WriteConfigXml(workbook, doc.OuterXml);
        }

        private string GetConfigXml()
        {
            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;
            if (workbook == null) return string.Empty;

            try
            {
                foreach (dynamic part in workbook.CustomXMLParts)
                {
                    try
                    {
                        string xml = part.XML as string;
                        if (xml != null && xml.Contains("<apurisk"))
                            return xml;
                    }
                    catch { }
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void WriteConfigXml(dynamic workbook, string xml)
        {
            try
            {
                foreach (dynamic part in workbook.CustomXMLParts)
                {
                    try
                    {
                        string partXml = part.XML as string;
                        if (partXml != null && partXml.Contains("<apurisk"))
                        {
                            part.Delete();
                            break;
                        }
                    }
                    catch { }
                }

                workbook.CustomXMLParts.Add(xml);
            }
            catch { }
        }

        private static string XmlEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Replace("&", "&amp;")
                        .Replace("<", "&lt;")
                        .Replace(">", "&gt;")
                        .Replace("\"", "&quot;")
                        .Replace("'", "&apos;");
        }

        private static void EnsureSheetInternal(dynamic workbook, string sheetName, string[] headers)
        {
            dynamic sheet = FindSheet(workbook, sheetName);
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add();
                sheet.Name = sheetName;
            }

            for (int i = 0; i < headers.Length; i++)
            {
                dynamic cell = sheet.Cells[1, i + 1];
                cell.Value2 = headers[i];
                cell.Font.Bold = true;
                Marshal.ReleaseComObject(cell);
            }

            dynamic usedRange = sheet.UsedRange;
            usedRange.Columns.AutoFit();
            Marshal.ReleaseComObject(usedRange);
        }

        private static dynamic FindSheet(dynamic workbook, string sheetName)
        {
            if (workbook == null)
            {
                return null;
            }

            foreach (dynamic sheet in workbook.Worksheets)
            {
                string currentName = sheet.Name as string;
                if (string.Equals(currentName, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }

                Marshal.ReleaseComObject(sheet);
            }

            return null;
        }
    }

    public sealed class RbsRow
    {
        public string Code { get; set; }
        public string Name { get; set; }
    }

    public sealed class RiskRow
    {
        public string Id { get; set; }
        public string RbsCode { get; set; }
        public string Description { get; set; }
    }
}
