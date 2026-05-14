namespace Apurisk.Core.RiskMaster
{
    public sealed class RiskMasterMapping
    {
        public string RiskIdColumn { get; set; }
        public string RbsCodeColumn { get; set; }
        public string RiskNameColumn { get; set; }
        public string RiskDescriptionColumn { get; set; }
        public string RbsCatalogTableName { get; set; }

        public RiskMasterMapping()
        {
            RiskIdColumn = "ID";
            RbsCodeColumn = "RBS";
            RiskNameColumn = "Riesgo";
            RiskDescriptionColumn = "Descripcion";
            RbsCatalogTableName = "Apurisk_RBS_Catalog";
        }
    }
}
