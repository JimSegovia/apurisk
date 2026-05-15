using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Apurisk.ExcelAddIn.Excel;

namespace Apurisk.ExcelAddIn.Forms
{
    public sealed class BowTieIntakeForm : Form
    {
        private readonly ExcelWorkbookGateway _gateway;

        private readonly Dictionary<string, TextBox> _fieldBoxes = new Dictionary<string, TextBox>();
        private readonly List<TextBox> _impactBoxes = new List<TextBox>();
        private readonly List<Label> _impactLabels = new List<Label>();

        private string _activeFieldKey;
        private int _impactCount = 1;

        private Button _btnLoadRange;
        private Button _btnAddImpact;
        private Button _btnRemoveImpact;
        private Button _btnAceptar;
        private Button _btnCancelar;
        private Panel _scrollPanel;
        private Panel _impactPanel;

        private static readonly string[] RequiredFields =
        {
            "RbsNameRange", "RbsCodeRange", "RiskTableRange", "RiskIdRange", "RiskTopRange",
            "RiskRbsCodeRange", "RiskDescriptionRange", "RiskCauseRange", "RiskPotentialEffectRange",
            "RiskProbabilityRange", "RiskImpactRange", "RiskSeverityRange", "RiskMitigationRange",
            "RiskOwnerRange"
        };

        private static readonly Dictionary<string, string> FieldLabels = new Dictionary<string, string>
        {
            { "RbsNameRange", "Nombre RBS" },
            { "RbsCodeRange", "Codigo RBS" },
            { "RiskTableRange", "Seleccion automatica" },
            { "RiskIdRange", "ID" },
            { "RiskTopRange", "TOP" },
            { "RiskRbsCodeRange", "Codigo RBS del riesgo" },
            { "RiskRbsNameRange", "Nombre RBS del riesgo" },
            { "RiskDescriptionRange", "Descripcion del riesgo" },
            { "RiskCauseRange", "Causas clave" },
            { "RiskPotentialEffectRange", "Impacto / efecto potencial" },
            { "RiskProbabilityRange", "Probabilidad" },
            { "RiskImpactRange", "Impacto" },
            { "RiskSeverityRange", "Gravedad" },
            { "RiskMitigationRange", "Medidas de mitigacion" },
            { "RiskOwnerRange", "Persona responsable" }
        };

        public BowTieIntakeForm(ExcelWorkbookGateway gateway)
        {
            _gateway = gateway;

            InitializeForm();
            LoadSavedValues();
            SetActiveField("RbsNameRange");
        }

        private void InitializeForm()
        {
            Text = "Apurisk - Ingresar Valores BowTie";
            ClientSize = new Size(640, 620);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular);

            _scrollPanel = new Panel
            {
                Location = new Point(8, 8),
                Size = new Size(624, 550),
                AutoScroll = true,
                BorderStyle = BorderStyle.None
            };
            Controls.Add(_scrollPanel);

            int y = 4;
            int labelWidth = 180;
            int boxWidth = 320;
            int rowHeight = 26;
            int indent = 192;

            foreach (var kvp in FieldLabels)
            {
                var lbl = new Label
                {
                    Text = kvp.Value + ":",
                    Location = new Point(8, y + 4),
                    Size = new Size(labelWidth, 18),
                    TextAlign = ContentAlignment.MiddleRight,
                    Font = Font
                };
                _scrollPanel.Controls.Add(lbl);

                var box = new TextBox
                {
                    Name = "txt" + kvp.Key,
                    Location = new Point(indent, y + 1),
                    Size = new Size(boxWidth, 22),
                    Font = Font,
                    ReadOnly = true,
                    BackColor = SystemColors.Window
                };
                box.Enter += (s, e) => SetActiveField(kvp.Key);
                box.Click += (s, e) => SetActiveField(kvp.Key);
                _scrollPanel.Controls.Add(box);
                _fieldBoxes[kvp.Key] = box;

                y += rowHeight;
            }

            _impactPanel = new Panel
            {
                Location = new Point(0, y),
                Size = new Size(600, 100),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            _scrollPanel.Controls.Add(_impactPanel);

            _btnLoadRange = new Button
            {
                Text = "Seleccionar rango",
                Location = new Point(indent, y),
                Size = new Size(130, 24),
                Font = Font
            };
            _btnLoadRange.Click += BtnLoadRange_Click;
            _scrollPanel.Controls.Add(_btnLoadRange);

            _btnAddImpact = new Button
            {
                Text = "+ Impacto",
                Location = new Point(indent + 136, y),
                Size = new Size(80, 24),
                Font = Font
            };
            _btnAddImpact.Click += BtnAddImpact_Click;
            _scrollPanel.Controls.Add(_btnAddImpact);

            _btnRemoveImpact = new Button
            {
                Text = "- Impacto",
                Location = new Point(indent + 222, y),
                Size = new Size(80, 24),
                Font = Font,
                Enabled = false
            };
            _btnRemoveImpact.Click += BtnRemoveImpact_Click;
            _scrollPanel.Controls.Add(_btnRemoveImpact);

            y += rowHeight + 8;

            _btnAceptar = new Button
            {
                Text = "Aceptar",
                Location = new Point(indent, y),
                Size = new Size(100, 28),
                Font = Font
            };
            _btnAceptar.Click += BtnAceptar_Click;
            _scrollPanel.Controls.Add(_btnAceptar);

            _btnCancelar = new Button
            {
                Text = "Cancelar",
                Location = new Point(indent + 108, y),
                Size = new Size(100, 28),
                Font = Font
            };
            _btnCancelar.Click += BtnCancelar_Click;
            _scrollPanel.Controls.Add(_btnCancelar);

            RenderImpactFields();
        }

        private void SetActiveField(string fieldKey)
        {
            _activeFieldKey = fieldKey;

            foreach (var box in _fieldBoxes.Values)
                box.BackColor = SystemColors.Window;

            foreach (var box in _impactBoxes)
                box.BackColor = SystemColors.Window;

            if (fieldKey == null)
                return;

            if (_fieldBoxes.ContainsKey(fieldKey))
            {
                _fieldBoxes[fieldKey].BackColor = Color.FromArgb(157, 195, 230);
                return;
            }

            if (fieldKey.StartsWith("ImpactCategory"))
            {
                string numStr = fieldKey.Replace("ImpactCategory", "");
                int idx;
                if (int.TryParse(numStr, out idx) && idx >= 1 && idx <= _impactBoxes.Count)
                    _impactBoxes[idx - 1].BackColor = Color.FromArgb(157, 195, 230);
            }
        }

        private void BtnLoadRange_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_activeFieldKey))
            {
                MessageBox.Show("Selecciona primero un cuadro del popup.", "Apurisk",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string label = GetLabelForKey(_activeFieldKey);

            string address = _gateway.PickRange("Selecciona el rango para '" + label + "'.");

            if (string.IsNullOrEmpty(address))
                return;

            SetFieldValue(_activeFieldKey, address);
        }

        private string GetLabelForKey(string fieldKey)
        {
            if (FieldLabels.ContainsKey(fieldKey))
                return FieldLabels[fieldKey];

            if (fieldKey.StartsWith("ImpactCategory"))
            {
                string numStr = fieldKey.Replace("ImpactCategory", "");
                return "Cat. Impacto " + numStr;
            }

            return fieldKey;
        }

        private void SetFieldValue(string fieldKey, string value)
        {
            if (_fieldBoxes.ContainsKey(fieldKey))
            {
                _fieldBoxes[fieldKey].Text = value;
                return;
            }

            if (fieldKey.StartsWith("ImpactCategory"))
            {
                string numStr = fieldKey.Replace("ImpactCategory", "");
                int idx;
                if (int.TryParse(numStr, out idx) && idx >= 1 && idx <= _impactBoxes.Count)
                    _impactBoxes[idx - 1].Text = value;
            }
        }

        private void BtnAddImpact_Click(object sender, EventArgs e)
        {
            _impactCount++;
            RenderImpactFields();
            LoadImpactValues();
            if (_impactBoxes.Count > 0)
            {
                _impactBoxes[_impactBoxes.Count - 1].Focus();
                SetActiveField("ImpactCategory" + _impactCount);
            }
        }

        private void BtnRemoveImpact_Click(object sender, EventArgs e)
        {
            if (_impactCount <= 1)
                return;

            _impactCount--;
            RenderImpactFields();
            LoadImpactValues();
            if (_impactBoxes.Count > 0)
            {
                _impactBoxes[_impactBoxes.Count - 1].Focus();
                SetActiveField("ImpactCategory" + _impactCount);
            }
        }

        private void BtnAceptar_Click(object sender, EventArgs e)
        {
            _gateway.SaveAllConfig(this);
            _gateway.ShowMessage("Apurisk",
                "Los parametros quedaron guardados y se mantendran cuando vuelvas a abrir esta ventana.");

            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        public string GetFieldValue(string fieldKey)
        {
            if (_fieldBoxes.ContainsKey(fieldKey))
                return _fieldBoxes[fieldKey].Text;

            if (fieldKey.StartsWith("ImpactCategory"))
            {
                string numStr = fieldKey.Replace("ImpactCategory", "");
                int idx;
                if (int.TryParse(numStr, out idx) && idx >= 1 && idx <= _impactBoxes.Count)
                    return _impactBoxes[idx - 1].Text;
            }

            return string.Empty;
        }

        public int ImpactCount
        {
            get { return _impactCount; }
        }

        public string[] GetImpactFieldKeys()
        {
            var keys = new string[_impactCount];
            for (int i = 0; i < _impactCount; i++)
                keys[i] = "ImpactCategory" + (i + 1);
            return keys;
        }

        private void RenderImpactFields()
        {
            foreach (var lbl in _impactLabels) _impactPanel.Controls.Remove(lbl);
            foreach (var box in _impactBoxes) _impactPanel.Controls.Remove(box);
            _impactLabels.Clear();
            _impactBoxes.Clear();

            int y = 0;
            int labelWidth = 116;
            int boxWidth = 120;
            int indent = 68;

            for (int i = 1; i <= _impactCount; i++)
            {
                int currentIndex = i;

                var lbl = new Label
                {
                    Text = "Cat. Impacto " + i + ":",
                    Location = new Point(8, y + 4),
                    Size = new Size(labelWidth, 18),
                    TextAlign = ContentAlignment.MiddleRight,
                    Font = Font,
                    BackColor = Color.Transparent
                };
                _impactPanel.Controls.Add(lbl);
                _impactLabels.Add(lbl);

                var box = new TextBox
                {
                    Name = "txtImpactCategory" + i,
                    Location = new Point(indent + labelWidth, y + 1),
                    Size = new Size(boxWidth, 22),
                    Font = Font,
                    ReadOnly = true,
                    BackColor = SystemColors.Window
                };
                box.Enter += (s, ev) => SetActiveField("ImpactCategory" + currentIndex);
                box.Click += (s, ev) => SetActiveField("ImpactCategory" + currentIndex);
                _impactPanel.Controls.Add(box);
                _impactBoxes.Add(box);

                y += 26;
            }

            _impactPanel.Height = y;
            RepositionButtons();
        }

        private void RepositionButtons()
        {
            int y = 4 + (_fieldBoxes.Count * 26);
            int indent = 192;

            _impactPanel.Location = new Point(0, y);
            y += _impactPanel.Height;

            _btnLoadRange.Location = new Point(indent, y + 2);
            _btnAddImpact.Location = new Point(indent + 136, y + 2);
            _btnRemoveImpact.Location = new Point(indent + 222, y + 2);
            _btnRemoveImpact.Enabled = _impactCount > 1;
            y += 38;

            _btnAceptar.Location = new Point(indent, y + 4);
            _btnCancelar.Location = new Point(indent + 108, y + 4);
        }

        private void LoadSavedValues()
        {
            foreach (var key in _fieldBoxes.Keys)
            {
                string saved = _gateway.ReadConfigValue("Field." + key);
                if (!string.IsNullOrEmpty(saved))
                    _fieldBoxes[key].Text = saved;
            }

            _impactCount = _gateway.GetImpactFieldCount();
            if (_impactCount < 1) _impactCount = 1;

            RenderImpactFields();
            LoadImpactValues();
        }

        private void LoadImpactValues()
        {
            for (int i = 0; i < _impactBoxes.Count && i < _impactCount; i++)
            {
                string saved = _gateway.ReadConfigValue("Field.ImpactCategory" + (i + 1));
                if (!string.IsNullOrEmpty(saved))
                    _impactBoxes[i].Text = saved;
            }
        }
    }
}
