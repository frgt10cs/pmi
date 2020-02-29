using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pmi.Model;

namespace Pmi
{
    class Excel
    {
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
 
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        private static WorksheetPart InsertSheet(WorkbookPart workbookPart, string nameSheet)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(worksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = nameSheet };
            sheets.Append(sheet);
            return worksheetPart;
        }
        
        public static string GetCellValue(Worksheet worksheet, WorkbookPart workbookPart, string nameCell)
        {
            string value = "0";
            Cell theCell = worksheet.Descendants<Cell>().Where(c => c.CellReference == nameCell).FirstOrDefault();
                if (theCell != null && theCell.InnerText.Length > 0)
                {
                    value = theCell.InnerText;
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                if (stringTable != null)
                                {
                                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                }
                                break;
                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            return value;
        }

        private static double GetHour(string data, string name)
        {
            int start = data.IndexOf(name);
            int lenght = 0;
            while (data[start] != ';')
            {
                start++;
            }
            start++;
            while(data[start + lenght] != ')')
            {
                lenght++;
            }
            return double.Parse(data.Substring(start, lenght));
        } 

        public static Employee GetEmployee(string way, string name)
        {
            string sheetName = "Лист1";
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(way, true))
            {
                SharedStringTablePart shareStringPart;
                if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                
                Sheet theSheet = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
                if (theSheet == null)
                {
                    return null;
                }
                WorksheetPart worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(theSheet.Id));
                Employee employee = new Employee();
                int row = 5;
                while(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()) != "Итого по ")
                {
                    //_________________________ВЕСЕННИЙ СЕМЕСТР_______________________
                    string lekEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "O" + row.ToString());
                    string prcEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "R" + row.ToString());
                    string labEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "U" + row.ToString());
                    if (lekEmployee.Contains(name) || prcEmployee.Contains(name) || labEmployee.Contains(name))
                    {
                        double countWeek = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "L" + row.ToString()));

                        Discipline discipline = new Discipline(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()));
                        discipline.Groups = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()).Split('\n').ToList();
                        discipline.CodeOP = "";
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "P" + row.ToString()) != "0" &&
                            lekEmployee.Contains(name))
                        {
                            discipline.Lectures = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "M" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "H" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * countWeek;
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "P" + row.ToString()) != "0" &&
                            prcEmployee.Contains(name))
                        {
                            if (prcEmployee.Contains(';'))
                            {
                                discipline.PracticalWork = GetHour(prcEmployee, name);
                            }
                            else
                            {
                                discipline.PracticalWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "P" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                    double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "I" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * countWeek;
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "S" + row.ToString()) != "0" &&
                            labEmployee.Contains(name))
                        {
                            if (labEmployee.Contains(';'))
                            {
                                discipline.LaboratoryWork = GetHour(labEmployee, name);
                            }
                            else
                            {
                                discipline.LaboratoryWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "S" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                    double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "J" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * countWeek;
                            }
                        }
                        //if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "W" + row.ToString()) != "0" &&
                        //    GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Y" + row.ToString()).Contains(name))
                        //{
                        //    discipline.ConsultationsByTheory = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString())) *
                        //        double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "W" + row.ToString())) * 2;
                        //}
                        //if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Z" + row.ToString()) != "0" &&
                        //    GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AB" + row.ToString()).Contains(name))
                        //{
                        //    discipline.ConsultationsByDiplom = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Z" + row.ToString())) *
                        //        double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString())) * 3.25;
                        //}
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AL" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AN" + row.ToString()).Contains(name))
                        {
                            discipline.Tests = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AL" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * 0.35;
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AO" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AQ" + row.ToString()).Contains(name))
                        {
                            discipline.Exam = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AO" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * 0.35;
                        }
                        employee.AutumnSemester.Disciplines.Add(discipline);
                    }

                    //_________________________ОСЕННИЙ СЕМЕСТР_______________________
                    lekEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AW" + row.ToString());
                    prcEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AZ" + row.ToString());
                    labEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BC" + row.ToString());
                    if (lekEmployee.Contains(name) || prcEmployee.Contains(name) || labEmployee.Contains(name))
                    {
                        double countWeek = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AT" + row.ToString()));

                        Discipline discipline = new Discipline(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()));
                        discipline.Groups = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()).Split('\n').ToList();
                        discipline.CodeOP = "";
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AU" + row.ToString()) != "0" &&
                            lekEmployee.Contains(name))
                        {
                            discipline.Lectures = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AU" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                            double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "H" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * countWeek;
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AX" + row.ToString()) != "0" &&
                            prcEmployee.Contains(name))
                        {
                            if (prcEmployee.Contains(';'))
                            {
                                discipline.PracticalWork = GetHour(prcEmployee, name);
                            }
                            else
                            {
                                discipline.PracticalWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AX" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                    double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "I" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * countWeek;
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BA" + row.ToString()) != "0" &&
                            labEmployee.Contains(name))
                        {
                            if (labEmployee.Contains(';'))
                            {
                                discipline.LaboratoryWork = GetHour(labEmployee, name);
                            }
                            else
                            {
                                discipline.LaboratoryWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BA" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                    double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "J" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * countWeek;
                            }
                        }
                        //if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "W" + row.ToString()) != "0" &&
                        //    GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Y" + row.ToString()).Contains(name))
                        //{
                        //    discipline.ConsultationsByTheory = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString())) *
                        //        double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "W" + row.ToString())) * 2;
                        //}
                        //if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Z" + row.ToString()) != "0" &&
                        //    GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AB" + row.ToString()).Contains(name))
                        //{
                        //    discipline.ConsultationsByDiplom = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Z" + row.ToString())) *
                        //        double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString())) * 3.25;
                        //}
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BT" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BV" + row.ToString()).Contains(name))
                        {
                            discipline.Tests = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BT" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * 0.35;
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BW" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BY" + row.ToString()).Contains(name))
                        {
                            discipline.Exam = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BW" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) *
                                double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "K" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture) * 0.35;
                        }
                        employee.SpringSemester.Disciplines.Add(discipline);
                    }
                    row++;
                }
                return employee;
            }
        }
    }
}
