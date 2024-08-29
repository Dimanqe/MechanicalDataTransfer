using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace MechanicalDataTransfer
{
    public partial class MechanicalDataTransferForm : System.Windows.Forms.Form
    {
        private string selectedExcelFilePath = null;
        private ExternalCommandData _commandData;
        public MechanicalDataTransferForm(ExternalCommandData commandData)
        {
            _commandData = commandData;
            InitializeComponent();
        }

        private void selectFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            selectedExcelFilePath=openFileDialog1.FileName;
        }

        private void transferButton_Click(object sender, EventArgs e)
        {
            Document document = _commandData.Application.ActiveUIDocument.Document;

            BuiltInCategory mechanicalEquipmentCategory = BuiltInCategory.OST_MechanicalEquipment;

            FilteredElementCollector mechanicalEquipmentCollector = new FilteredElementCollector(document);
            mechanicalEquipmentCollector.OfCategory(mechanicalEquipmentCategory);
            mechanicalEquipmentCollector.WhereElementIsNotElementType();

            Dictionary<(string SpaceNumber, string RadiatorLength), string> excelData = GetExcelData(selectedExcelFilePath);

            foreach (Element mechanicalEquipmentElement in mechanicalEquipmentCollector)
            {
                BoundingBoxXYZ boundingBox = GetBoundingBox(mechanicalEquipmentElement);
                Space space = FindSpaceContainingBoundingBox(document, boundingBox);

                if (space != null)
                {
                    string spaceNumber = ConvertCustomParameterValue(space.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString());
                    var radiatorLengthInRevit = mechanicalEquipmentElement.get_Parameter(new Guid("ced1a5d6-7b2a-4e2f-95f6-d557ec4eeaa9"));
                    var revitRadiatorLength = radiatorLengthInRevit.AsValueString();

                    foreach (var excelEntry in excelData)
                    {
                        var excelSpaceNumber = ConvertCustomParameterValue(excelEntry.Key.SpaceNumber);
                        var excelRadiatorLength = excelEntry.Key.RadiatorLength;

                        if (spaceNumber == excelSpaceNumber && MatchSpaceAndRadiatorLength(revitRadiatorLength, excelRadiatorLength))
                        {
                            using (Transaction transaction = new Transaction(document, "Set Parameters"))
                            {
                                transaction.Start();
                                Parameter columnHParameter = mechanicalEquipmentElement.LookupParameter("ДСК1_Тепловая мощность");
                                if (columnHParameter != null)
                                {
                                    string columnHValue = excelEntry.Value;       
                                    double powerValue;

                                    if (double.TryParse(columnHValue, out powerValue))
                                    {
                                        columnHValue = columnHValue.Replace(',', '.');
                                        columnHParameter.SetValueString(powerValue.ToString());
                                        mechanicalEquipmentElement.LookupParameter("ДСК1_Тепловая мощность").SetValueString(columnHValue);
                                    }
                                }
                                transaction.Commit();
                            }
                        }
                    }
                }
            }

            TaskDialog.Show("Запись мощности радиаторов", "Выполнено");
        }


        private BoundingBoxXYZ GetBoundingBox(Element element)
        {
            BoundingBoxXYZ boundingBox = element.get_BoundingBox(null);
            return boundingBox;
        }

        private Space FindSpaceContainingBoundingBox(Document document, BoundingBoxXYZ boundingBox)
        {
            FilteredElementCollector spaceCollector = new FilteredElementCollector(document);
            spaceCollector.OfCategory(BuiltInCategory.OST_MEPSpaces);

            foreach (Element spaceElement in spaceCollector)
            {
                Space spaceCandidate = spaceElement as Space;
                if (spaceCandidate != null)
                {
                    BoundingBoxXYZ spaceBoundingBox = spaceCandidate.get_BoundingBox(null);

                    if (spaceBoundingBox != null && AreBoundingBoxesIntersecting(spaceBoundingBox, boundingBox))
                    {
                        return spaceCandidate;
                    }
                }
            }

            return null;
        }

        private bool AreBoundingBoxesIntersecting(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            return (box1.Min.X <= box2.Max.X && box1.Max.X >= box2.Min.X)
                && (box1.Min.Y <= box2.Max.Y && box1.Max.Y >= box2.Min.Y)
                && (box1.Min.Z <= box2.Max.Z && box1.Max.Z >= box2.Min.Z);
        }

        private Dictionary<(string SpaceNumber, string RadiatorLength), string> GetExcelData(string excelFilePath)
        {
            Dictionary<(string SpaceNumber, string RadiatorLength), string> excelData = new Dictionary<(string, string), string>();

            using (FileStream file = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0);

                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        ICell cellSpaceName = row.GetCell(2);
                        string spaceName = cellSpaceName.ToString();

                        ICell cellColumnHValue = row.GetCell(7);
                        string columnHValue = cellColumnHValue.ToString();

                        ICell cellRadiatorLength = row.GetCell(5);
                        string radiatorLength = cellRadiatorLength.ToString();

                        excelData[(spaceName, ConvertRadiatorLengthToMMFromExcelFile(radiatorLength))] = columnHValue;
                    }
                }
            }

            return excelData;
        }

        private bool MatchSpaceAndRadiatorLength(string revitRadiatorLength, string excelRadiatorLength)
        {
            if (string.IsNullOrEmpty(excelRadiatorLength))
            {
                return false;
            }

            return string.Equals(revitRadiatorLength, excelRadiatorLength);
        }

        public static string ConvertRadiatorLengthToMMFromExcelFile(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            string[] inputSplited = input.Split(',');
            if (inputSplited.Length >= 2)
            {
                string[] filteredSecondPart = inputSplited[1].Split(' ');
                if (inputSplited[0] == "0")
                {
                    return string.Join("", filteredSecondPart[0]);
                }
                return string.Join("", inputSplited[0], filteredSecondPart[0]);
            }

            return string.Empty;
        }

        public static string ConvertCustomParameterValue(string originalValue)
        {
            if (string.IsNullOrEmpty(originalValue))
                return string.Empty;

            string[] input = originalValue.Split('.');
            bool flag = false;

            for (int i = 0; i < input.Length; i++)
            {
                if (int.TryParse(input[i], out var result))
                {
                    flag = true;
                    break;            
                }
            }

            if (flag)
            {
                string[] parts = originalValue.Split('.');

                if (parts.Length >= 4 && parts.Length > 1)
                {
                    string resultValue = parts[0] + "/" + string.Join("", parts.Skip(1));
                    return resultValue;
                }
                else if (parts.Length < 4 && parts.Length > 1 && parts[parts.Length - 1].Length > 1)
                {
                    string transformedValue = parts[0] + "/" + parts[parts.Length - 2] + parts[2];
                    return transformedValue;
                }
                else if (parts.Length > 2 && parts.Length < 4 && parts[parts.Length - 1].Length < 2)
                {
                    string transformedValue = parts[0] + "/" + parts[1] + "0" + parts[2];
                    return transformedValue;
                }
                else if (parts.Length < 2)
                {
                    return originalValue;
                }
            }

            return originalValue;
        }
    }

}

