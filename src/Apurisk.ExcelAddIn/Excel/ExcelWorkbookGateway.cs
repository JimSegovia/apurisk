using System;
using System.Runtime.InteropServices;
using System.Xml;

namespace Apurisk.ExcelAddIn.Excel
{
    public sealed class ExcelWorkbookGateway
    {
        private readonly object _excelApplication;

        private const string APURISK_XML_NS = "http://apurisk.dev/config";

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
            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;
            if (workbook == null) return string.Empty;

            try
            {
                dynamic part = GetOrCreateConfigPart(workbook);
                if (part == null) return string.Empty;

                string xml = part.XML;
                if (string.IsNullOrEmpty(xml)) return string.Empty;

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xml);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("a", APURISK_XML_NS);

                XmlNode node = doc.SelectSingleNode("//a:e[@k='" + XmlEscape(keyName) + "']", nsmgr);
                return node != null ? (node.Attributes["v"].Value ?? string.Empty) : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public void WriteConfigValue(string keyName, string keyValue)
        {
            dynamic excel = _excelApplication;
            dynamic workbook = excel.ActiveWorkbook;
            if (workbook == null) return;

            try
            {
                dynamic part = GetOrCreateConfigPart(workbook);
                if (part == null) return;

                string xml = part.XML;
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(string.IsNullOrEmpty(xml)
                    ? "<?xml version=\"1.0\"?><c xmlns=\"" + APURISK_XML_NS + "\"/>"
                    : xml);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("a", APURISK_XML_NS);

                XmlNode node = doc.SelectSingleNode("//a:e[@k='" + XmlEscape(keyName) + "']", nsmgr);

                if (node != null)
                {
                    node.Attributes["v"].Value = keyValue;
                }
                else
                {
                    XmlElement elem = doc.CreateElement("e", APURISK_XML_NS);
                    elem.SetAttribute("k", keyName);
                    elem.SetAttribute("v", keyValue);
                    doc.DocumentElement.AppendChild(elem);
                }

                part.LoadXML(doc.OuterXml);
            }
            catch { }
        }

        public int GetImpactFieldCount()
        {
            string val = ReadConfigValue("ImpactFieldCount");
            int count;
            if (!int.TryParse(val, out count) || count < 1)
                count = 1;
            return count;
        }

        public bool RiskIdExists(string riskIdAddress, string riskIdValue)
        {
            dynamic excel = _excelApplication;
            try
            {
                dynamic range = excel.ActiveWorkbook.Range(riskIdAddress);
                foreach (dynamic cell in range.Cells)
                {
                    string cellValue = cell.Value2 != null ? cell.Value2.ToString().Trim() : string.Empty;
                    if (string.Equals(cellValue, riskIdValue, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { }
            return false;
        }

        public void SaveAllConfig(Forms.BowTieIntakeForm form)
        {
            string[] fields =
            {
                "RbsNameRange", "RbsCodeRange", "RiskTableRange", "RiskIdRange", "RiskTopRange",
                "RiskRbsCodeRange", "RiskRbsNameRange", "RiskDescriptionRange", "RiskCauseRange",
                "RiskPotentialEffectRange", "RiskProbabilityRange", "RiskImpactRange", "RiskSeverityRange",
                "RiskMitigationRange", "RiskOwnerRange"
            };

            foreach (var field in fields)
            {
                string value = form.GetFieldValue(field) ?? string.Empty;
                WriteConfigValue("Field." + field, value);
            }

            string[] impactKeys = form.GetImpactFieldKeys();
            for (int i = 0; i < impactKeys.Length; i++)
            {
                string value = form.GetFieldValue(impactKeys[i]) ?? string.Empty;
                WriteConfigValue("Field." + impactKeys[i], value);
            }

            WriteConfigValue("ImpactFieldCount", form.ImpactCount.ToString());
        }

        private dynamic GetOrCreateConfigPart(dynamic workbook)
        {
            try
            {
                foreach (dynamic part in workbook.CustomXMLParts)
                {
                    try
                    {
                        string ns = part.NamespaceURI as string;
                        if (ns == APURISK_XML_NS)
                            return part;
                    }
                    catch { }
                }

                string initXml = "<?xml version=\"1.0\"?><c xmlns=\"" + APURISK_XML_NS + "\"/>";
                return workbook.CustomXMLParts.Add(initXml);
            }
            catch
            {
                return null;
            }
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
}
