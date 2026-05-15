using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Apurisk.ExcelAddIn.Excel;

namespace Apurisk.ExcelAddIn.Forms
{
    public sealed class RbsTreeForm : Form
    {
        private readonly ExcelWorkbookGateway _gateway;

        private TreeNode _root;
        private Panel _canvas;
        private Button _btnExpandAll;
        private Button _btnCollapseAll;

        private const int NODE_WIDTH = 140;
        private const int NODE_HEIGHT = 50;
        private const int H_GAP = 80;
        private const int V_GAP = 24;
        private const int RISK_WIDTH = 200;
        private const int RISK_HEIGHT = 38;
        private const int RISK_GAP = 8;
        private const int MARGIN = 32;

        public RbsTreeForm(ExcelWorkbookGateway gateway)
        {
            _gateway = gateway;

            Text = "Apurisk - Arbol RBS";
            WindowState = FormWindowState.Maximized;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular);

            _btnExpandAll = new Button
            {
                Text = "Expandir todo",
                Location = new Point(12, 10),
                Size = new Size(110, 26),
                Font = new Font("Segoe UI", 8f)
            };
            _btnExpandAll.Click += (s, e) => { if (_root != null) { SetAllExpanded(_root, true); RebuildAndDraw(); } };
            Controls.Add(_btnExpandAll);

            _btnCollapseAll = new Button
            {
                Text = "Colapsar todo",
                Location = new Point(128, 10),
                Size = new Size(110, 26),
                Font = new Font("Segoe UI", 8f)
            };
            _btnCollapseAll.Click += (s, e) => { if (_root != null) { SetAllExpanded(_root, false); _root.Expanded = true; RebuildAndDraw(); } };
            Controls.Add(_btnCollapseAll);

            _canvas = new Panel
            {
                Location = new Point(0, 46),
                Size = new Size(ClientSize.Width, ClientSize.Height - 46),
                BackColor = Color.White,
                AutoScroll = true
            };
            _canvas.Paint += Canvas_Paint;
            _canvas.MouseClick += Canvas_MouseClick;
            _canvas.Resize += (s, e) => _canvas.Invalidate();
            Controls.Add(_canvas);

            Resize += (s, e) =>
            {
                _canvas.Size = new Size(ClientSize.Width, ClientSize.Height - 46);
            };

            BuildTree();
        }

        private void BuildTree()
        {
            var rbsRows = _gateway.ReadRbsFromRanges();
            if (rbsRows.Count == 0)
            {
                MessageBox.Show("No se encontraron datos en los rangos de RBS configurados.\nVerifica que 'Codigo RBS' y 'Nombre RBS' tengan datos en 'Ingresar valores'.\n\nLa ventana se abrira vacia para debug.",
                    "Apurisk - Arbol RBS", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _canvas.Invalidate();
                return;
            }

            var riskRows = _gateway.ReadRisksFromRanges();

            var nodeMap = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in rbsRows)
            {
                var node = new TreeNode { Code = row.Code, Name = row.Name };
                nodeMap[node.Code] = node;
            }

            BuildHierarchy(nodeMap);

            _root = FindRoot(nodeMap);
            if (_root == null && nodeMap.Count > 0)
            {
                var enumerator = nodeMap.Values.GetEnumerator();
                enumerator.MoveNext();
                _root = enumerator.Current;
            }

            if (_root == null)
            {
                MessageBox.Show("No se pudo construir la jerarquia del arbol RBS.", "Apurisk - Arbol RBS",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _canvas.Invalidate();
                return;
            }

            AttachRisks(_root, riskRows);
            RebuildAndDraw();
        }

        private static void BuildHierarchy(Dictionary<string, TreeNode> nodeMap)
        {
            foreach (var kvp in nodeMap)
            {
                string code = kvp.Key;
                TreeNode node = kvp.Value;

                int lastDot = code.LastIndexOf('.');
                if (lastDot > 0)
                {
                    string parentCode = code.Substring(0, lastDot);
                    if (nodeMap.ContainsKey(parentCode))
                    {
                        nodeMap[parentCode].Children.Add(node);
                    }
                }
            }
        }

        private static TreeNode FindRoot(Dictionary<string, TreeNode> nodeMap)
        {
            var childCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in nodeMap)
            {
                foreach (var child in kvp.Value.Children)
                    childCodes.Add(child.Code);
            }

            foreach (var kvp in nodeMap)
            {
                if (!childCodes.Contains(kvp.Key))
                    return kvp.Value;
            }

            return null;
        }

        private static void AttachRisks(TreeNode node, List<RiskRow> allRisks)
        {
            foreach (var risk in allRisks)
            {
                if (string.Equals(risk.RbsCode, node.Code, StringComparison.OrdinalIgnoreCase))
                {
                    node.Risks.Add(risk);
                }
            }

            foreach (var child in node.Children)
                AttachRisks(child, allRisks);
        }

        private void SetAllExpanded(TreeNode node, bool expanded)
        {
            node.Expanded = expanded;
            foreach (var child in node.Children)
                SetAllExpanded(child, expanded);
        }

        private void RebuildAndDraw()
        {
            if (_root == null) return;
            CalculateLayout(_root, 0);
            SetCanvasSize();
            _canvas.Invalidate();
        }

        private double CalculateLayout(TreeNode node, double nextY)
        {
            if (node.Expanded && node.Children.Count > 0)
            {
                foreach (var child in node.Children)
                    nextY = CalculateLayout(child, nextY);

                double firstCenter = node.Children[0].LayoutY;
                double lastCenter = node.Children[node.Children.Count - 1].LayoutY;
                node.LayoutY = (firstCenter + lastCenter) / 2.0;
                return nextY;
            }
            else
            {
                int riskCount = node.Expanded ? Math.Max(1, node.Risks.Count) : 1;
                node.LayoutY = nextY + (double)(riskCount - 1) / 2.0;
                return nextY + riskCount;
            }
        }

        private void SetCanvasSize()
        {
            if (_root == null) return;
            int maxDepth = GetMaxDepth(_root, 0);
            double totalRows = CountRows(_root);

            int width = MARGIN * 2 + maxDepth * (NODE_WIDTH + H_GAP) + RISK_WIDTH + H_GAP;
            int height = MARGIN * 2 + (int)(totalRows * (NODE_HEIGHT + V_GAP));

            _canvas.AutoScrollMinSize = new Size(Math.Max(width, _canvas.ClientSize.Width), Math.Max(height, _canvas.ClientSize.Height));
        }

        private static int GetMaxDepth(TreeNode node, int depth)
        {
            if (!node.Expanded || node.Children.Count == 0)
                return depth + (node.Risks.Count > 0 ? 1 : 0);

            int maxChildDepth = depth;
            foreach (var child in node.Children)
            {
                int childDepth = GetMaxDepth(child, depth + 1);
                if (childDepth > maxChildDepth)
                    maxChildDepth = childDepth;
            }
            return maxChildDepth;
        }

        private static double CountRows(TreeNode node)
        {
            if (!node.Expanded || node.Children.Count == 0)
                return node.Expanded ? Math.Max(1, node.Risks.Count) : 1;

            double total = 0;
            foreach (var child in node.Children)
                total += CountRows(child);
            return total;
        }

        private void Canvas_Paint(object sender, PaintEventArgs e)
        {
            if (_root == null) return;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            int scrollX = _canvas.AutoScrollPosition.X;
            int scrollY = _canvas.AutoScrollPosition.Y;

            GraphicsState state = e.Graphics.Save();
            e.Graphics.TranslateTransform(scrollX, scrollY);
            DrawNode(e.Graphics, _root, 0);
            e.Graphics.Restore(state);
        }

        private void DrawNode(Graphics g, TreeNode node, int depth)
        {
            int x = MARGIN + depth * (NODE_WIDTH + H_GAP);
            int y = MARGIN + (int)(node.LayoutY * (NODE_HEIGHT + V_GAP));

            RectangleF rect = new RectangleF(x, y, NODE_WIDTH, NODE_HEIGHT);
            node.RenderedBounds = rect;

            bool isLeaf = node.Children.Count == 0;
            bool active = node.Expanded;

            Color fillColor = isLeaf ? Color.FromArgb(230, 244, 234) : Color.FromArgb(232, 240, 254);
            Color borderColor = isLeaf ? Color.FromArgb(19, 115, 51) : Color.FromArgb(25, 103, 210);
            if (active) borderColor = isLeaf ? Color.FromArgb(19, 115, 51) : Color.FromArgb(25, 103, 210);

            using (var path = RoundedRect(rect, 6))
            using (var brush = new SolidBrush(fillColor))
            using (var pen = new Pen(borderColor, active ? 2f : 1.5f))
            {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }

            string displayText = node.Code;
            if (!string.IsNullOrEmpty(node.Name))
                displayText += "\n" + node.Name;

            using (var textBrush = new SolidBrush(Color.FromArgb(32, 33, 36)))
            using (var textFont = new Font(Font.FontFamily, 7.5f))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                g.DrawString(displayText, textFont, textBrush, rect, format);
            }

            if (node.Expanded)
            {
                if (node.Children.Count > 0)
                {
                    foreach (var child in node.Children)
                    {
                        DrawConnection(g, rect, child, depth + 1);
                        DrawNode(g, child, depth + 1);
                    }
                }
                else
                {
                    DrawRiskBoxes(g, rect, node);
                }
            }
        }

        private void DrawConnection(Graphics g, RectangleF parentRect, TreeNode child, int childDepth)
        {
            int childX = MARGIN + childDepth * (NODE_WIDTH + H_GAP);
            int childY = MARGIN + (int)(child.LayoutY * (NODE_HEIGHT + V_GAP));

            float startX = parentRect.Right;
            float startY = parentRect.Top + parentRect.Height / 2f;
            float endX = childX;
            float endY = childY + NODE_HEIGHT / 2f;

            float midX = startX + (endX - startX) / 2f;

            using (var pen = new Pen(Color.FromArgb(154, 160, 166), 1.3f))
            {
                g.DrawLine(pen, startX, startY, midX, startY);
                g.DrawLine(pen, midX, startY, midX, endY);
                g.DrawLine(pen, midX, endY, endX, endY);
            }
        }

        private void DrawRiskBoxes(Graphics g, RectangleF parentRect, TreeNode node)
        {
            float riskX = parentRect.Right + 50;
            float baseY = parentRect.Top + parentRect.Height / 2f;
            float totalHeight = node.Risks.Count * (RISK_HEIGHT + RISK_GAP) - RISK_GAP;
            float startY = baseY - totalHeight / 2f;

            for (int i = 0; i < node.Risks.Count; i++)
            {
                float ry = startY + i * (RISK_HEIGHT + RISK_GAP);
                RectangleF riskRect = new RectangleF(riskX, ry, RISK_WIDTH, RISK_HEIGHT);

                using (var path = RoundedRect(riskRect, 4))
                using (var brush = new SolidBrush(Color.FromArgb(254, 247, 224)))
                using (var pen = new Pen(Color.FromArgb(227, 116, 0), 1.2f))
                {
                    g.FillPath(brush, path);
                    g.DrawPath(pen, path);
                }

                string riskText = node.Risks[i].Id;
                if (!string.IsNullOrEmpty(node.Risks[i].Description))
                    riskText += "\n" + node.Risks[i].Description;

                using (var textBrush = new SolidBrush(Color.FromArgb(32, 33, 36)))
                using (var textFont = new Font(Font.FontFamily, 7f))
                {
                    var format = new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center,
                        Trimming = StringTrimming.EllipsisCharacter
                    };
                    g.DrawString(riskText, textFont, textBrush, riskRect, format);
                }

                using (var linePen = new Pen(Color.FromArgb(154, 160, 166), 1f))
                {
                    g.DrawLine(linePen, parentRect.Right, parentRect.Top + parentRect.Height / 2f,
                        parentRect.Right + 25, parentRect.Top + parentRect.Height / 2f);
                    g.DrawLine(linePen, parentRect.Right + 25, parentRect.Top + parentRect.Height / 2f,
                        parentRect.Right + 25, ry + RISK_HEIGHT / 2f);
                    g.DrawLine(linePen, parentRect.Right + 25, ry + RISK_HEIGHT / 2f,
                        riskX, ry + RISK_HEIGHT / 2f);
                }
            }
        }

        private static GraphicsPath RoundedRect(RectangleF rect, float radius)
        {
            var path = new GraphicsPath();
            float r = Math.Min(radius, Math.Min(rect.Width / 2f, rect.Height / 2f));

            path.AddArc(rect.X, rect.Y, r * 2, r * 2, 180, 90);
            path.AddArc(rect.Right - r * 2, rect.Y, r * 2, r * 2, 270, 90);
            path.AddArc(rect.Right - r * 2, rect.Bottom - r * 2, r * 2, r * 2, 0, 90);
            path.AddArc(rect.X, rect.Bottom - r * 2, r * 2, r * 2, 90, 90);
            path.CloseFigure();

            return path;
        }

        private void Canvas_MouseClick(object sender, MouseEventArgs e)
        {
            if (_root == null) return;

            int scrollX = -_canvas.AutoScrollPosition.X;
            int scrollY = -_canvas.AutoScrollPosition.Y;

            int mouseX = e.X - scrollX;
            int mouseY = e.Y - scrollY;

            TreeNode clicked = HitTest(_root, mouseX, mouseY);
            if (clicked != null)
            {
                clicked.Expanded = !clicked.Expanded;
                RebuildAndDraw();
            }
        }

        private static TreeNode HitTest(TreeNode node, int mouseX, int mouseY)
        {
            if (node.RenderedBounds.Contains(mouseX, mouseY))
                return node;

            if (node.Expanded)
            {
                foreach (var child in node.Children)
                {
                    TreeNode hit = HitTest(child, mouseX, mouseY);
                    if (hit != null) return hit;
                }
            }

            return null;
        }
    }

    internal sealed class TreeNode
    {
        public string Code;
        public string Name;
        public bool Expanded;
        public double LayoutY;
        public RectangleF RenderedBounds;
        public List<TreeNode> Children = new List<TreeNode>();
        public List<RiskRow> Risks = new List<RiskRow>();
    }
}
