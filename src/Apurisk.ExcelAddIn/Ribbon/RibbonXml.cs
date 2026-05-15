namespace Apurisk.ExcelAddIn.Ribbon
{
    internal static class RibbonXml
    {
        public static string GetXml()
        {
            return
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabApurisk"" label=""Apurisk"">
        <group id=""grpBowTie"" label=""Analisis BowTie"">
          <button id=""btnApuriskBase"" label=""Crear base"" size=""large"" imageMso=""TableInsert"" onAction=""OnCreateBase""/>
          <button id=""btnApuriskIntake"" label=""Ingresar valores"" size=""large"" imageMso=""DiagramTargetInsertClassic"" onAction=""OnBowTieIntake""/>
          <button id=""btnApuriskRbs"" label=""Arbol RBS"" size=""large"" imageMso=""OrganizationChartInsert"" onAction=""OnOpenRbsExplorer""/>
          <button id=""btnApuriskBowTie"" label=""Analizar"" size=""large"" imageMso=""DiagramExpand"" onAction=""OnOpenBowTie""/>
          <separator id=""sepApuriskBowTie1""/>
          <button id=""btnApuriskValidate"" label=""Validar"" imageMso=""AcceptInvitation"" onAction=""OnValidate""/>
          <button id=""btnApuriskInsert"" label=""Insertar valores"" imageMso=""TableUpdate"" onAction=""OnInsertValues""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
}
