using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MechanicalDataTransfer
{
    [Transaction(TransactionMode.Manual)]
    public class MechanicalDataTransferCom : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var form = new MechanicalDataTransferForm(commandData);
            form.ShowDialog();
            return Result.Succeeded;
        }
    }
}