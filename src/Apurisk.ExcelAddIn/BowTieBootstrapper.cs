using System.Windows.Forms;
using Apurisk.ExcelAddIn.Excel;
using Apurisk.ExcelAddIn.Forms;

namespace Apurisk.ExcelAddIn
{
    internal sealed class BowTieBootstrapper
    {
        private readonly ExcelWorkbookGateway _workbook;

        public BowTieBootstrapper(object excelApplication)
        {
            _workbook = new ExcelWorkbookGateway(excelApplication);
        }

        public void CreateInitialWorkbookBase()
        {
            _workbook.EnsureSheet("Apurisk_RBS", new[] { "CodigoRBS", "Nombre", "PadreRBS", "Nivel", "Descripcion" });
            _workbook.EnsureSheet("Apurisk_RiskMaster_Map", new[] { "CampoApurisk", "ColumnaExcel", "Obligatorio", "Notas" });
            _workbook.EnsureSheet("Apurisk_BowTie_Work", new[] { "RiskID", "RBS", "Elemento", "Tipo", "Valor", "Owner", "Efectividad", "Notas" });
            _workbook.EnsureSheet("Apurisk_Diagram", new[] { "Area reservada para el diagrama BowTie" });
            _workbook.ActivateSheet("Apurisk_RBS");
            MessageBox.Show("Base inicial creada para Apurisk.", "Apurisk", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OpenBowTieIntake()
        {
            if (!_workbook.HasActiveWorkbook)
            {
                MessageBox.Show("No hay un libro activo para trabajar.", "Apurisk",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            using (var form = new BowTieIntakeForm(_workbook))
            {
                form.ShowDialog();
            }
        }

        public void OpenRbsExplorerPlaceholder()
        {
            MessageBox.Show("Aqui abriremos la vista de arbol RBS.", "Apurisk - Analisis BowTie", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OpenBowTiePlaceholder()
        {
            MessageBox.Show("Aqui abriremos la vista BowTie del riesgo seleccionado.", "Apurisk - Analisis BowTie", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ValidatePlaceholder()
        {
            MessageBox.Show("Validacion inicial pendiente.", "Apurisk - Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void InsertValuesPlaceholder()
        {
            MessageBox.Show("Insercion en tabla maestra pendiente.", "Apurisk - Tabla maestra", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
