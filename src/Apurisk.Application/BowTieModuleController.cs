using Apurisk.Application.Excel;

namespace Apurisk.Application
{
    public sealed class BowTieModuleController
    {
        private readonly IExcelWorkbookGateway _workbook;

        public BowTieModuleController(IExcelWorkbookGateway workbook)
        {
            _workbook = workbook;
        }

        public void CreateInitialWorkbookBase()
        {
            _workbook.EnsureSheets(new[]
            {
                new SheetDefinition("Apurisk_Config", new[] { "Parametro", "Valor", "Notas" }),
                new SheetDefinition("Apurisk_RBS", new[] { "CodigoRBS", "Nombre", "PadreRBS", "Nivel", "Descripcion" }),
                new SheetDefinition("Apurisk_RiskMaster_Map", new[] { "CampoApurisk", "ColumnaExcel", "Obligatorio", "Notas" }),
                new SheetDefinition("Apurisk_BowTie_Work", new[] { "RiskID", "RBS", "Elemento", "Tipo", "Valor", "Owner", "Efectividad", "Notas" }),
                new SheetDefinition("Apurisk_Diagram", new[] { "Area reservada para el diagrama BowTie" })
            });

            _workbook.ActivateSheet("Apurisk_Config");
            _workbook.ShowMessage("Apurisk", "Base inicial creada. El siguiente paso sera configurar columnas de tabla maestra y catalogo RBS.");
        }

        public void OpenRbsExplorerPlaceholder()
        {
            _workbook.ShowMessage("Apurisk - Analisis BowTie", "Aqui abriremos la vista de arbol RBS: categorias, subcategorias y riesgos asociados.");
        }

        public void OpenBowTiePlaceholder()
        {
            _workbook.ShowMessage("Apurisk - Analisis BowTie", "Aqui abriremos la vista BowTie del riesgo seleccionado.");
        }

        public void ValidatePlaceholder()
        {
            _workbook.ShowMessage("Apurisk - Validacion", "Validacion inicial pendiente: columnas obligatorias, RBS valido y riesgos sin clasificar.");
        }

        public void InsertValuesPlaceholder()
        {
            _workbook.ShowMessage("Apurisk - Tabla maestra", "Insercion pendiente: los valores BowTie se escribiran en la tabla maestra configurada.");
        }
    }
}
