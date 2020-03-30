using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pmi.Model;
using Pmi.Builders;
using Pmi.Directors;
using Pmi.Service.Implimentation;
using System.Configuration;
using Pmi.Service.Abstraction;

namespace Pmi
{                          
    class Excel
    {                       
        #region shit
        class CellData
        {
            public string Column;
            public uint Row;
            public string Data;
        }

        CacheService<List<ExcelCellFormat>> cacheService;

        public Excel(CacheService<List<ExcelCellFormat>> cacheService)
        {
            this.cacheService = cacheService;
        }

        private string[] Column = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q" };
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
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

        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
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

        private WorksheetPart InsertSheet(WorkbookPart workbookPart, string nameSheet)
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

        public string GetCellValue(Worksheet worksheet, WorkbookPart workbookPart, string nameCell)
        {
            string value = "0";
            Cell theCell = worksheet.Descendants<Cell>().Where(c => c.CellReference == nameCell).FirstOrDefault();
            if (theCell != null && theCell.InnerText.Length > 0)
            {
                value = theCell.CellValue != null ? theCell.CellValue.InnerText : theCell.InnerText;
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

        private double GetHour(string data, string name)
        {
            int start = data.IndexOf(name);
            int lenght = 0;
            while (data[start] != ';')
            {
                start++;
            }
            start++;
            while (data[start + lenght] != ')')
            {
                lenght++;
            }
            return double.Parse(data.Substring(start, lenght));
        }

        public Employee GetEmployee(string path, string lastName, string name, string patronymic, string rank)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, true))
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

                Sheet Sheet1 = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Лист1").FirstOrDefault();
                Sheet Sheet2 = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Лист2").FirstOrDefault();
                Sheet SheetDipl = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Дипл исх данные").FirstOrDefault();
                Sheet SheetPrac = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Рук-ли практики  бак").FirstOrDefault();
                Sheet SheetPracMag = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Рук-ли практики маг ").FirstOrDefault();
                if (Sheet1 == null || Sheet2 == null || SheetDipl == null)
                {
                    return null;
                }
                WorksheetPart worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(Sheet1.Id));
                Employee employee = new Employee();
                employee.LastName = lastName;
                employee.FirstName = name;
                employee.Patronymic = patronymic;
                employee.Rank = rank;

                int row = 5;
                while (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()) != "Итого по ")
                {
                    //_________________________ОСЕННИЙ СЕМЕСТР_______________________
                    string lekEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "O" + row.ToString());
                    string prcEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "R" + row.ToString());
                    string labEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "U" + row.ToString());
                    if (lekEmployee.Contains(lastName) || prcEmployee.Contains(lastName) || labEmployee.Contains(lastName))
                    {
                        Discipline discipline = new Discipline(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()));
                        discipline.Groups = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()).Split('\n').ToList();
                        discipline.CodeOP = "";
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "M" + row.ToString()) != "0" &&
                            lekEmployee.Contains(lastName))
                        {
                            discipline.Lectures = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "N" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                            discipline.ConsultationsByTheory = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "V" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "P" + row.ToString()) != "0" &&
                            prcEmployee.Contains(lastName))
                        {
                            if (prcEmployee.Contains(';'))
                            {
                                discipline.PracticalWork = GetHour(prcEmployee, lastName);
                            }
                            else
                            {
                                discipline.PracticalWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Q" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "S" + row.ToString()) != "0" &&
                            labEmployee.Contains(lastName))
                        {
                            if (labEmployee.Contains(';'))
                            {
                                discipline.LaboratoryWork = GetHour(labEmployee, lastName);
                            }
                            else
                            {
                                discipline.LaboratoryWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "T" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "W" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Y" + row.ToString()).Contains(lastName))
                        {
                            discipline.Coursework = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "X" + row.ToString()));
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AL" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AN" + row.ToString()).Contains(lastName))
                        {
                            discipline.Tests = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AM" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AO" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AQ" + row.ToString()).Contains(lastName))
                        {
                            discipline.Exam = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AP" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        }
                        employee.SpringSemester.Disciplines.Add(discipline);
                    }

                    //_________________________ВЕСЕННИЙ СЕМЕСТР_______________________
                    lekEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AW" + row.ToString());
                    prcEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AZ" + row.ToString());
                    labEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BC" + row.ToString());
                    if (lekEmployee.Contains(lastName) || prcEmployee.Contains(lastName) || labEmployee.Contains(lastName))
                    {
                        Discipline discipline = new Discipline(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()));
                        discipline.Groups = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()).Split('\n').ToList();
                        discipline.CodeOP = "";
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AU" + row.ToString()) != "0" &&
                            lekEmployee.Contains(lastName))
                        {
                            discipline.Lectures = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AV" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                            discipline.ConsultationsByTheory = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BD" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AX" + row.ToString()) != "0" &&
                            prcEmployee.Contains(lastName))
                        {
                            if (prcEmployee.Contains(';'))
                            {
                                discipline.PracticalWork = GetHour(prcEmployee, lastName);
                            }
                            else
                            {
                                discipline.PracticalWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AY" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BA" + row.ToString()) != "0" &&
                            labEmployee.Contains(lastName))
                        {
                            if (labEmployee.Contains(';'))
                            {
                                discipline.LaboratoryWork = GetHour(labEmployee, lastName);
                            }
                            else
                            {
                                discipline.LaboratoryWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BB" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BE" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BG" + row.ToString()).Contains(lastName))
                        {
                            discipline.Coursework = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BF" + row.ToString()));
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BT" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BV" + row.ToString()).Contains(lastName))
                        {
                            discipline.Tests = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BU" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BW" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BY" + row.ToString()).Contains(lastName))
                        {
                            discipline.Exam = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BX" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        }
                        employee.AutumnSemester.Disciplines.Add(discipline);
                    }
                    row++;
                }

                worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(Sheet2.Id));
                row = 4;
                Discipline bachelor = new Discipline("Диплом бакалавры");
                Discipline magister = new Discipline("Диплом магистры");
                Discipline aspirants = new Discipline("Аспиранты");
                Discipline anotherWork = new Discipline("Другие  виды уч. работы");
                Discipline diplPro = new Discipline("по дипл. проект.");
                Discipline gekB = new Discipline("Бакалавры");
                Discipline gekM = new Discipline("Магистры");
                while (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()) != "ГЭК комиссии")
                {
                    if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()).Contains(lastName))
                    {
                        bachelor.Diploms = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "L" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        if (bachelor.Diploms != 0) { employee.AutumnSemester.Disciplines.Add(bachelor); }
                        magister.Diploms = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "O" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        if (magister.Diploms != 0) { employee.AutumnSemester.Disciplines.Add(magister); }
                        diplPro.ConsultationsByDiplom = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BF" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        if (diplPro.ConsultationsByDiplom != 0) { employee.AutumnSemester.Disciplines.Add(diplPro); }
                        anotherWork.AnotherWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BG" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        if (anotherWork.AnotherWork != 0) { employee.AutumnSemester.Disciplines.Add(anotherWork); }
                        aspirants.Aspirants = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "P" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        if (aspirants.Aspirants != 0) { employee.AutumnSemester.Disciplines.Add(aspirants); }
                        gekB.GEK = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "N" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        if (gekB.GEK != 0) { employee.AutumnSemester.Disciplines.Add(gekB); }
                        gekM.GEK = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "M" + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                        if (gekM.GEK != 0) { employee.AutumnSemester.Disciplines.Add(gekM); }
                    }
                    row++;
                }
                worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(SheetDipl.Id));
                row = 2;
                while (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "B" + row.ToString()) != "0")
                {
                    if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "B" + row.ToString()).Contains(lastName))
                    {
                        string group = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString());
                        if (group[1] == '4')
                        {
                            bachelor.Groups.Add(group);
                        }
                        else if (group[1] == '2')
                        {
                            magister.Groups.Add(group);
                        }
                    }
                    row++;
                }
                //________________________________ПРАКТИКА________________________________
                worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(SheetPrac.Id));
                row = 3;
                for (int j = 0; j < 2; j++)
                {
                    while (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "A" + row.ToString()) != "0")
                    {
                        Discipline prac = new Discipline(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "A" + (row - 2).ToString()));
                        bool CanAdd = true;
                        while (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "A" + row.ToString()) != "№  п/п" && GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "A" + row.ToString()) != "0")
                        {
                            if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()).Contains(lastName))
                            {
                                if (CanAdd)
                                {
                                    if (int.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "E" + row.ToString()).Split('.')[1].Substring(0, 2)) < 9)
                                    {
                                        employee.AutumnSemester.Disciplines.Add(prac);
                                    }
                                    else
                                    {
                                        employee.SpringSemester.Disciplines.Add(prac);
                                    }
                                    CanAdd = false;
                                }
                                prac.PracticalWork += double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, (j==1?"K":"L") + row.ToString()), System.Globalization.CultureInfo.InvariantCulture);
                                string code = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "D" + row.ToString());
                                bool found = false;
                                for (int i = 0; i < prac.Groups.Count; i++)
                                {
                                    if (prac.Groups[i].Contains(code))
                                    {
                                        prac.Groups[i] += ", " + GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "B" + row.ToString());
                                        found = true;
                                        break;
                                    }
                                }
                                if (!found)
                                {
                                    prac.Groups.Add(code + " " + GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "B" + row.ToString()));
                                }
                            }
                            row++;
                        }
                        row++;
                    }
                    row = 4;
                    worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(SheetPracMag.Id));
                }
                return employee;
            }
        }
        #endregion

        private Sheet GetSheet(WorkbookPart workbookPart, string nameSheet)
        {
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

            foreach (var ItemSheets in sheets.Elements<Sheet>())
            {
                if (ItemSheets.Name == nameSheet)
                {
                    return ItemSheets;
                }
            }

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            string relationshipId = workbookPart.GetIdOfPart(worksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = nameSheet };
            sheets.Append(sheet);
            return sheet;
        }
        public void CreateRaportInFile(string path, Employee employee)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, true))
            {                
                InitStyles(doc);
                SharedStringTablePart shareStringPart;
                if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                var sheet = GetSheet(doc.WorkbookPart, employee.LastName + " " + employee.FirstName[0] + "." + employee.Patronymic[0] + ".(A4)");
                var worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(sheet.Id));
                //CreateRaport(employee, worksheetPart, shareStringPart);
                //sheet.Remove();
                //doc.WorkbookPart.DeletePart(worksheetPart);
                doc.WorkbookPart.Workbook.Save();
            }
        }       

        /// <summary>
        /// Добавляет стили к документу
        /// </summary>
        /// <param name="document"></param>
        /// <param name="stylesheet"></param>
        private void AppendStylesToDocument(SpreadsheetDocument document, ExcelStylesheet stylesheet)
        {
            var documentStylesheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
            foreach (var font in stylesheet.Fonts)
                documentStylesheet.Fonts.AppendChild(font.CloneNode(true));
            foreach (var fill in stylesheet.Fills)
                documentStylesheet.Fills.AppendChild(fill.CloneNode(true));
            foreach (var border in stylesheet.Borders)
                documentStylesheet.Borders.AppendChild(border.CloneNode(true));            
            foreach (var cellFormat in stylesheet.CellFormats)
                documentStylesheet.CellFormats.AppendChild((CellFormat)cellFormat);
            document.Save();
        }

        /// <summary>
        /// Инициализирует необходимые стили и заносит информацию о них в кэш
        /// </summary>
        /// <param name="document"></param>
        private void InitStyles(SpreadsheetDocument document)
        {
            var workbookpart = document.WorkbookPart;
            var workStylePart = workbookpart.WorkbookStylesPart;
            var styleSheet = workStylePart.Stylesheet;

            // вынести в отдельный метод?
            #region генерация стилей для страниц
            ExcelStylesheetBuilder builder = new ExcelStylesheetBuilder((uint)styleSheet.Fonts.ChildElements.Count,
                (uint)styleSheet.CellFormats.ChildElements.Count);
            ExcelStylesheetDirector director = new ExcelStylesheetDirector() { StylesheetBuilder = builder };

            director.BuildReportStylesheet();
            var reportStylesheet = builder.GetStylesheet();
            #endregion            

            AppendStylesToDocument(document, reportStylesheet);
            cacheService.Cache(reportStylesheet.CellFormats);            
        }

        /// <summary>
        /// Сравнивает два формата ячейки на идентичность по полям
        /// </summary>
        /// <param name="first"></param>
        /// <param name="second"></param>
        /// <returns></returns>
        public bool AreCellFormatEquals(CellFormat first, ExcelCellFormat second)
        {
            return first.FontId == second.FontId && first.Alignment.Horizontal.Value == second.HorizontalAlignment
                && first.Alignment.Vertical.Value == second.VerticalAlignment && first.BorderId == second.BorderId && first.FillId == second.FillId;
        }

        /// <summary>
        /// Проверяет, совпадают ли индексы стилей в документе с индексами стилей в кэше. Сравнивает только первые и последние форматы ячеек.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="stylesheet"></param>
        public bool AreIndexesSame(SpreadsheetDocument document)
        {
            var excelCellFormats =  cacheService.UploadCache();
            int firstId = Convert.ToInt32(excelCellFormats.First().Id);
            int lastId = Convert.ToInt32(excelCellFormats.Last().Id);
            return AreCellFormatEquals(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[firstId] as CellFormat, excelCellFormats.First())
                && AreCellFormatEquals(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[lastId] as CellFormat, excelCellFormats.Last());
        }

        #region shit
        public void CreateRaport(string path, Employee employee)
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = $"{employee.LastName} {employee.FirstName[0]}. {employee.Patronymic[0]}."
            };
            sheets.Append(sheet);
            SharedStringTablePart shareStringPart;
            if (workbookpart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookpart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookpart.AddNewPart<SharedStringTablePart>();
            }
            #region CreateCloumn
            if (worksheetPart.Worksheet.GetFirstChild<Columns>() == null)
            {
                Columns columns = new Columns();
                columns.Append(new Column() { Min = 1, Max = 1, Width = 6.86, CustomWidth = true });
                columns.Append(new Column() { Min = 2, Max = 2, Width = 18, CustomWidth = true });
                columns.Append(new Column() { Min = 3, Max = 3, Width = 7.57, CustomWidth = true });
                columns.Append(new Column() { Min = 4, Max = 4, Width = 4.71, CustomWidth = true });
                columns.Append(new Column() { Min = 5, Max = 5, Width = 5.86, CustomWidth = true });
                columns.Append(new Column() { Min = 6, Max = 6, Width = 3.86, CustomWidth = true });
                columns.Append(new Column() { Min = 7, Max = 7, Width = 7.71, CustomWidth = true });
                columns.Append(new Column() { Min = 8, Max = 8, Width = 7.71, CustomWidth = true });
                columns.Append(new Column() { Min = 9, Max = 9, Width = 5, CustomWidth = true });
                columns.Append(new Column() { Min = 10, Max = 10, Width = 7.43, CustomWidth = true });
                columns.Append(new Column() { Min = 11, Max = 11, Width = 7.57, CustomWidth = true });
                columns.Append(new Column() { Min = 12, Max = 12, Width = 9.14, CustomWidth = true });
                columns.Append(new Column() { Min = 13, Max = 13, Width = 5.71, CustomWidth = true });
                columns.Append(new Column() { Min = 14, Max = 14, Width = 4.43, CustomWidth = true });
                columns.Append(new Column() { Min = 15, Max = 15, Width = 5.71, CustomWidth = true });
                columns.Append(new Column() { Min = 16, Max = 16, Width = 8.43, CustomWidth = true });
                columns.Append(new Column() { Min = 17, Max = 17, Width = 10.43, CustomWidth = true });
                worksheetPart.Worksheet.InsertAt(columns, 0);
            }
            MergeCells mergeCells = new MergeCells();
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() == 0)
            {
                if (worksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheetPart.Worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheetPart.Worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SortState>().First());
                }
                else if (worksheetPart.Worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheetPart.Worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<Scenarios>().First());
                }
                else if (worksheetPart.Worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheetPart.Worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
                }
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:Q1") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:Q2") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("M4:Q4") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("D5:K5") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("M5:Q5") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("M6:Q6") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("C7:L7") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("C8:L8") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("C9:L9") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("G11:H11") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("I11:L11") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A11:B13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("C11:C13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("D11:D13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("E11:E13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("F11:F13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("G12:G13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("H12:H13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("I12:I13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("J12:J13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("K12:K13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("L12:L13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("M11:M13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("N11:N13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("O11:O13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("P11:P13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("Q11:Q13") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A14:B14") });
            }
            #endregion
            #region CreateRow
            CellData[] cells =
            {
                new CellData(){Column = "A", Row = 1, Data = "федеральное государственное бюджетное образовательное учреждение высшего образования "},
                new CellData(){Column = "A", Row = 2, Data = "«Казанский национальный исследовательский технический университет им. А.Н. Туполева-КАИ» (КНИТУ-КАИ)"},
                new CellData(){Column = "M", Row = 4, Data = "УТВЕРЖДАЮ"},
                new CellData(){Column = "D", Row = 5, Data = "ПЛАН УЧЕБНОЙ НАГРУЗКИ"},
                new CellData(){Column = "M", Row = 5, Data = "Зав. кафедрой ПМИ"},
                new CellData(){Column = "M", Row = 6, Data = "Зайдуллин С.С."},
                new CellData(){Column = "O", Row = 7, Data = "подпись, ФИО"},
                new CellData(){Column = "C", Row = 7, Data = $"{employee.Rank}, {employee.FirstName} {employee.LastName} {employee.Patronymic}"},
                new CellData(){Column = "C", Row = 8, Data = "должность, ФИО, ученая степень, ученое звание, доля ставки, штатность"},
                
                new CellData(){Column = "A", Row = 11, Data = "Код ОП,\nиндекс дисциплины,\nнаименование дисциплины"},
                new CellData(){Column = "C", Row = 11, Data = "Группа"},
                new CellData(){Column = "D", Row = 11, Data = "Лекц"},
                new CellData(){Column = "E", Row = 11, Data = "Практ"},
                new CellData(){Column = "F", Row = 11, Data = "Лаб"},
                new CellData(){Column = "G", Row = 11, Data = "Консульт. студ."},
                new CellData(){Column = "I", Row = 11, Data = "Руководство"},
                new CellData(){Column = "M", Row = 11, Data = "ГЭК"},
                new CellData(){Column = "N", Row = 11, Data = "ЗАЧ"},
                new CellData(){Column = "O", Row = 11, Data = "ЭКЗ"},
                new CellData(){Column = "P", Row = 11, Data = "Другие  виды уч. работы"},
                new CellData(){Column = "Q", Row = 11, Data = " ВСЕГО"},
                new CellData(){Column = "G", Row = 12, Data = "по теор. курсу"},
                new CellData(){Column = "H", Row = 12, Data = "по дипл. проект."},
                new CellData(){Column = "I", Row = 12, Data = "асп-ми"},
                new CellData(){Column = "J", Row = 12, Data = "курс. проект. (раб.)"},
                new CellData(){Column = "K", Row = 12, Data = "дипл. проект."},
                new CellData(){Column = "L", Row = 12, Data = "практ."},
            };
            foreach (var data in cells)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(data.Data, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }
            for (uint i = 2; i < 17; i++)
            {
                Cell cell = InsertCellInWorksheet(Column[i], 14, worksheetPart);
                cell.CellValue = new CellValue(i.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                if (i == 2)
                {
                    cell = InsertCellInWorksheet("A", 14, worksheetPart);
                    cell.CellValue = new CellValue("1");
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                }
            }
            #endregion
            mergeCells.Append(new MergeCell() { Reference = new StringValue("A15:Q15") });
            Cell semCell = InsertCellInWorksheet("A", 15, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("О  С  Е  Н  Н  И  Й     С  Е  М  Е  С  Т  Р  ", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            uint row = 16;
            foreach (var discipline in employee.SpringSemester.Disciplines)
            {
                Cell cell = InsertCellInWorksheet("A", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.CodeOP, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("B", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.Name, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("C", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(String.Join(", ", discipline.Groups), shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("D", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Lectures.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("E", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.PracticalWork.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("F", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.LaboratoryWork.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("G", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByTheory.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("H", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByDiplom.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("I", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Aspirants.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("J", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Coursework.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("K", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Diploms.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("L", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Practice.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("M", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.GEK.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("N", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Tests.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("O", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Exam.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("P", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.AnotherWork.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("Q", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.TotalForThisDiscipline().ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                row++;
            }
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:C{row}") });

            #region TotalSpring
            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("Итого за осенний семестр", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            CellData[] totalS =
            {
                new CellData(){Column = "D", Row = row, Data = employee.SpringSemester.TotalForLectures().ToString().Replace(',', '.')},
                new CellData(){Column = "E", Row = row, Data = employee.SpringSemester.TotalForPracticalWork().ToString().Replace(',', '.')},
                new CellData(){Column = "F", Row = row, Data = employee.SpringSemester.TotalForLaboratoryWork().ToString().Replace(',', '.')},
                new CellData(){Column = "G", Row = row, Data = employee.SpringSemester.TotalForConsultationsByTheory().ToString().Replace(',', '.')},
                new CellData(){Column = "H", Row = row, Data = employee.SpringSemester.TotalForConsultationsByDiplom().ToString().Replace(',', '.')},
                new CellData(){Column = "I", Row = row, Data = employee.SpringSemester.TotalForAspirants().ToString().Replace(',', '.')},
                new CellData(){Column = "J", Row = row, Data = employee.SpringSemester.TotalForCoursework().ToString().Replace(',', '.')},
                new CellData(){Column = "K", Row = row, Data = employee.SpringSemester.TotalForDiploms().ToString().Replace(',', '.')},
                new CellData(){Column = "L", Row = row, Data = employee.SpringSemester.TotalForPractice().ToString().Replace(',', '.')},
                new CellData(){Column = "M", Row = row, Data = employee.SpringSemester.TotalForGEK().ToString().Replace(',', '.')},
                new CellData(){Column = "N", Row = row, Data = employee.SpringSemester.TotalForTests().ToString().Replace(',', '.')},
                new CellData(){Column = "O", Row = row, Data = employee.SpringSemester.TotalForExam().ToString().Replace(',', '.')},
                new CellData(){Column = "P", Row = row, Data = employee.SpringSemester.TotalForAnotherWork().ToString().Replace(',', '.')},
                new CellData(){Column = "Q", Row = row, Data = employee.SpringSemester.TotalForSemester().ToString().Replace(',', '.')}
            };
            foreach (var data in totalS)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(data.Data);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            totalS = null;
            #endregion

            row++;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:Q{row}") });
            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("В  Е  С  Е  Н  Н  И  Й     С  Е  М  Е  С  Т  Р  ", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            row++;

            foreach (var discipline in employee.AutumnSemester.Disciplines)
            {
                Cell cell = InsertCellInWorksheet("A", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.CodeOP, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("B", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.Name, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("C", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(String.Join(", ", discipline.Groups), shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell = InsertCellInWorksheet("D", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Lectures.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("E", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.PracticalWork.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("F", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.LaboratoryWork.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("G", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByTheory.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("H", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByDiplom.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("I", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Aspirants.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("J", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Coursework.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("K", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Diploms.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("L", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Practice.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("M", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.GEK.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("N", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Tests.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("O", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Exam.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("P", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.AnotherWork.ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell = InsertCellInWorksheet("Q", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.TotalForThisDiscipline().ToString().Replace(',', '.'));
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                row++;
            }
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:C{row}") });
            #region TotalAutumn
            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("Итого за весенний семестр", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            CellData[] totalA =
            {
                new CellData(){Column = "D", Row = row, Data = employee.AutumnSemester.TotalForLectures().ToString().Replace(',', '.')},
                new CellData(){Column = "E", Row = row, Data = employee.AutumnSemester.TotalForPracticalWork().ToString().Replace(',', '.')},
                new CellData(){Column = "F", Row = row, Data = employee.AutumnSemester.TotalForLaboratoryWork().ToString().Replace(',', '.')},
                new CellData(){Column = "G", Row = row, Data = employee.AutumnSemester.TotalForConsultationsByTheory().ToString().Replace(',', '.')},
                new CellData(){Column = "H", Row = row, Data = employee.AutumnSemester.TotalForConsultationsByDiplom().ToString().Replace(',', '.')},
                new CellData(){Column = "I", Row = row, Data = employee.AutumnSemester.TotalForAspirants().ToString().Replace(',', '.')},
                new CellData(){Column = "J", Row = row, Data = employee.AutumnSemester.TotalForCoursework().ToString().Replace(',', '.')},
                new CellData(){Column = "K", Row = row, Data = employee.AutumnSemester.TotalForDiploms().ToString().Replace(',', '.')},
                new CellData(){Column = "L", Row = row, Data = employee.AutumnSemester.TotalForPractice().ToString().Replace(',', '.')},
                new CellData(){Column = "M", Row = row, Data = employee.AutumnSemester.TotalForGEK().ToString().Replace(',', '.')},
                new CellData(){Column = "N", Row = row, Data = employee.AutumnSemester.TotalForTests().ToString().Replace(',', '.')},
                new CellData(){Column = "O", Row = row, Data = employee.AutumnSemester.TotalForExam().ToString().Replace(',', '.')},
                new CellData(){Column = "P", Row = row, Data = employee.AutumnSemester.TotalForAnotherWork().ToString().Replace(',', '.')},
                new CellData(){Column = "Q", Row = row, Data = employee.AutumnSemester.TotalForSemester().ToString().Replace(',', '.')}
            };
            foreach (var data in totalA)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(data.Data);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            totalA = null;
            #endregion

            row++;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:C{row}") });

            #region Total
            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("ВСЕГО ЗА ГОД", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            CellData[] total =
            {
                new CellData(){Column = "D", Row = row, Data = employee.LecturesForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "E", Row = row, Data = employee.PracticalWorkForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "F", Row = row, Data = employee.LaboratoryWorkForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "G", Row = row, Data = employee.ConsultationsByTheoryForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "H", Row = row, Data = employee.ConsultationsByDiplomForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "I", Row = row, Data = employee.AspirantsForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "J", Row = row, Data = employee.CourseworkForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "K", Row = row, Data = employee.DiplomsForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "L", Row = row, Data = employee.PracticeForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "M", Row = row, Data = employee.GakForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "N", Row = row, Data = employee.TestsForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "O", Row = row, Data = employee.ExamForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "P", Row = row, Data = employee.AnotherWorkForYear().ToString().Replace(',', '.')},
                new CellData(){Column = "Q", Row = row, Data = employee.Year().ToString().Replace(',', '.')}
            };
            foreach (var data in total)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(data.Data);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            total = null;
            #endregion
            row += 2;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"K{row}:O{row}") });
            semCell = InsertCellInWorksheet("K", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            row++;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"K{row}:O{row}") });
            semCell = InsertCellInWorksheet("K", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("подпись преподавателя", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }
        #endregion
    }
}
