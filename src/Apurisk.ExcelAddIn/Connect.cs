using System;
using System.Runtime.InteropServices;
using Apurisk.ExcelAddIn.Diagnostics;
using Apurisk.ExcelAddIn.Excel;
using Apurisk.ExcelAddIn.Ribbon;
using Extensibility;
using Office = Microsoft.Office.Core;

namespace Apurisk.ExcelAddIn
{
    [ComVisible(true)]
    [Guid("7BD16DC9-26B6-4C37-8E23-A4E80504D9E4")]
    [ProgId("Apurisk.ExcelAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private BowTieBootstrapper _bowTie;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                AddInLog.Write("OnConnection");
                _bowTie = new BowTieBootstrapper(application);
            }
            catch (Exception exception)
            {
                AddInLog.WriteException("OnConnection", exception);
                throw;
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            AddInLog.Write("OnDisconnection");
            _bowTie = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                AddInLog.Write("GetCustomUI | " + ribbonId);
                return RibbonXml.GetXml();
            }
            catch (Exception exception)
            {
                AddInLog.WriteException("GetCustomUI", exception);
                throw;
            }
        }

        public void OnCreateBase(object control)
        {
            AddInLog.Write("OnCreateBase");
            _bowTie.CreateInitialWorkbookBase();
        }

        public void OnOpenRbsExplorer(object control)
        {
            AddInLog.Write("OnOpenRbsExplorer");
            _bowTie.OpenRbsExplorer();
        }

        public void OnOpenBowTie(object control)
        {
            AddInLog.Write("OnOpenBowTie");
            _bowTie.OpenBowTiePlaceholder();
        }

        public void OnValidate(object control)
        {
            AddInLog.Write("OnValidate");
            _bowTie.ValidatePlaceholder();
        }

        public void OnInsertValues(object control)
        {
            AddInLog.Write("OnInsertValues");
            _bowTie.InsertValuesPlaceholder();
        }

        public void OnBowTieIntake(object control)
        {
            AddInLog.Write("OnBowTieIntake");
            _bowTie.OpenBowTieIntake();
        }
    }
}
