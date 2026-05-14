using System.Windows.Forms;
using Apurisk.ExcelAddIn.Excel;

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
            _workbook.EnsureSheet("Apurisk_Config", new[] { "Parametro", "Valor", "Notas" });
            _workbook.EnsureSheet("Apurisk_RBS", new[] { "CodigoRBS", "Nombre", "PadreRBS", "Nivel", "Descripcion" });
            _workbook.EnsureSheet("Apurisk_RiskMaster_Map", new[] { "CampoApurisk", "ColumnaExcel", "Obligatorio", "Notas" });
            _workbook.EnsureSheet("Apurisk_BowTie_Work", new[] { "RiskID", "RBS", "Elemento", "Tipo", "Valor", "Owner", "Efectividad", "Notas" });
            _workbook.EnsureSheet("Apurisk_Diagram", new[] { "Area reservada para el diagrama BowTie" });
            _workbook.ActivateSheet("Apurisk_Config");
            MessageBox.Show("Base inicial creada para Apurisk.", "Apurisk", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
