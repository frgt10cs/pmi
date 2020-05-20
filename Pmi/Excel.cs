using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using Pmi.Model;
using Pmi.Builders;
using Pmi.Directors;
using Pmi.Service.Abstraction;
using System.Globalization;

namespace Pmi
{
    class Excel
    {
        class CellData
        {
            public string Column;
            public uint Row;
            public string Data;
            public uint StyleIndex;
        }

        public event EventHandler OnProgressChanged;
        public event EventHandler OnStatusChanged;
        private readonly string[] Column = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q" };
        readonly CacheService<List<ExcelCellFormat>> cacheService;

        public Excel(CacheService<List<ExcelCellFormat>> cacheService)
        {
            this.cacheService = cacheService;
        }

        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() {
                    RowIndex = rowIndex
                };
                sheetData.Append(row);
            }

            var cells = row.Elements<Cell>();

            if (cells.Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return cells.First(c => c.CellReference.Value == cellReference);
            }
            else
            {
                var refCell = cells.FirstOrDefault(cell => string.Compare(cell.CellReference.Value, cellReference, true) > 0);

                var newCell = new Cell() {
                    CellReference = cellReference 
                };
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

            int pos = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToList().FindIndex(i => i.InnerText == text);

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            shareStringPart.SharedStringTable.Save();

            return pos;
        }

        private WorksheetPart GetSheet(WorkbookPart workbookPart, string nameSheet)
        {
            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var count = sheets.Elements<Sheet>().Count(i => i.Name.Value.Contains(nameSheet));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var relationshipId = workbookPart.GetIdOfPart(worksheetPart);

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
            var sheet = new Sheet() {
                Id = relationshipId,
                SheetId = sheetId,
                Name = nameSheet + " " + count.ToString()
            };
            sheets.Append(sheet);
            return worksheetPart;
        }

        private string GetCellValue(Worksheet worksheet, WorkbookPart workbookPart, string nameCell)
        {
            string value = "0";
            Cell theCell = worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == nameCell);
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
        private void InitStyles(SpreadsheetDocument document, out List<ExcelCellFormat> cellFormats)
        {
            var workbookpart = document.WorkbookPart;
            var workStylePart = workbookpart.WorkbookStylesPart;
            var styleSheet = workStylePart.Stylesheet;

            // вынести в отдельный метод?
            #region генерация стилей для страниц
            var builder = new ExcelStylesheetBuilder(
                (uint)styleSheet.Fonts.ChildElements.Count,
                (uint)styleSheet.CellFormats.ChildElements.Count,
                (uint)styleSheet.Borders.ChildElements.Count);
            var director = new ExcelStylesheetDirector() {
                StylesheetBuilder = builder 
            };

            director.BuildReportStylesheet();
            var reportStylesheet = builder.GetStylesheet();
            #endregion            

            AppendStylesToDocument(document, reportStylesheet);
            cacheService.Cache(reportStylesheet.CellFormats);
            cellFormats = reportStylesheet.CellFormats;
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
        public bool AreIndexesSame(SpreadsheetDocument document, out List<ExcelCellFormat> cellFormats)
        {
            var excelCellFormats = cacheService.UploadCache();
            int firstId = Convert.ToInt32(excelCellFormats.First().Id);
            int lastId = Convert.ToInt32(excelCellFormats.Last().Id);
            var children = document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements;

            if (children.Count > lastId &&
                AreCellFormatEquals(children[firstId] as CellFormat, excelCellFormats.First()) &&
                AreCellFormatEquals(children[lastId] as CellFormat, excelCellFormats.Last()))
            {
                cellFormats = excelCellFormats;
                return true;
            }
            cellFormats = null;
            return false;
        }

        /// <summary>
        /// Сбор данных о преподавателе
        /// </summary>
        /// <param name="path"> Путь к файлу с данными</param>
        /// <param name="employee"> Преподаватель</param>
        /// <param name="year"> Учебный год</param>
        /// <returns> Преподаватель</returns>
        public Employee GetEmployee(string path, Employee employee)
        {
            OnProgressChanged?.Invoke(0, null);
            OnStatusChanged?.Invoke("Сбор данных: ", null);
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

                var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>();
                Sheet Sheet1 = sheet.FirstOrDefault(s => s.Name == "Лист1");
                Sheet Sheet2 = sheet.FirstOrDefault(s => s.Name == "Лист2");
                Sheet SheetDipl = sheet.FirstOrDefault(s => s.Name == "Дипл исх данные");
                Sheet SheetPrac = sheet.FirstOrDefault(s => s.Name == "Рук-ли практики  бак");
                Sheet SheetPracMag = sheet.FirstOrDefault(s => s.Name == "Рук-ли практики маг ");
                if (Sheet1 == null || Sheet2 == null || SheetDipl == null)
                {
                    return null;
                }
                WorksheetPart worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(Sheet1.Id));
                OnProgressChanged?.Invoke(1, null);
                
                OnStatusChanged?.Invoke("Сбор данных: Дисциплины", null);
                #region Создание дисциплин
                
                int row = 5;
                while (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()) != "Итого по ")
                {
                    //_________________________ОСЕННИЙ СЕМЕСТР_______________________
                    string lekEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "O" + row.ToString());
                    string prcEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "R" + row.ToString());
                    string labEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "U" + row.ToString());
                    if (lekEmployee.Contains(employee.LastName) || prcEmployee.Contains(employee.LastName) || labEmployee.Contains(employee.LastName))
                    {
                        var discipline = new Discipline(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()))
                        {
                            Groups = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()).Split('\n').ToList(),
                            CodeOP = ""
                        };
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "M" + row.ToString()) != "0" &&
                            lekEmployee.Contains(employee.LastName))
                        {
                            discipline.Lectures = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "N" + row.ToString()), CultureInfo.InvariantCulture);
                            discipline.ConsultationsByTheory = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "V" + row.ToString()), CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "P" + row.ToString()) != "0" &&
                            prcEmployee.Contains(employee.LastName))
                        {
                            if (prcEmployee.Contains(';'))
                            {
                                discipline.PracticalWork = GetHour(prcEmployee, employee.LastName);
                            }
                            else
                            {
                                discipline.PracticalWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Q" + row.ToString()), CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "S" + row.ToString()) != "0" &&
                            labEmployee.Contains(employee.LastName))
                        {
                            if (labEmployee.Contains(';'))
                            {
                                discipline.LaboratoryWork = GetHour(labEmployee, employee.LastName);
                            }
                            else
                            {
                                discipline.LaboratoryWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "T" + row.ToString()), CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "W" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "Y" + row.ToString()).Contains(employee.LastName))
                        {
                            discipline.Coursework = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "X" + row.ToString()));
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AL" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AN" + row.ToString()).Contains(employee.LastName))
                        {
                            discipline.Tests = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AM" + row.ToString()), CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AO" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AQ" + row.ToString()).Contains(employee.LastName))
                        {
                            discipline.Exam = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AP" + row.ToString()), CultureInfo.InvariantCulture);
                        }
                        employee.SpringSemester.Disciplines.Add(discipline);
                    }

                    //_________________________ВЕСЕННИЙ СЕМЕСТР_______________________
                    lekEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AW" + row.ToString());
                    prcEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AZ" + row.ToString());
                    labEmployee = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BC" + row.ToString());
                    if (lekEmployee.Contains(employee.LastName) || prcEmployee.Contains(employee.LastName) || labEmployee.Contains(employee.LastName))
                    {
                        Discipline discipline = new Discipline(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()));
                        discipline.Groups = GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "F" + row.ToString()).Split('\n').ToList();
                        discipline.CodeOP = "";
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AU" + row.ToString()) != "0" &&
                            lekEmployee.Contains(employee.LastName))
                        {
                            discipline.Lectures = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AV" + row.ToString()), CultureInfo.InvariantCulture);
                            discipline.ConsultationsByTheory = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BD" + row.ToString()), CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AX" + row.ToString()) != "0" &&
                            prcEmployee.Contains(employee.LastName))
                        {
                            if (prcEmployee.Contains(';'))
                            {
                                discipline.PracticalWork = GetHour(prcEmployee, employee.LastName);
                            }
                            else
                            {
                                discipline.PracticalWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "AY" + row.ToString()), CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BA" + row.ToString()) != "0" &&
                            labEmployee.Contains(employee.LastName))
                        {
                            if (labEmployee.Contains(';'))
                            {
                                discipline.LaboratoryWork = GetHour(labEmployee, employee.LastName);
                            }
                            else
                            {
                                discipline.LaboratoryWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BB" + row.ToString()), CultureInfo.InvariantCulture);
                            }
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BE" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BG" + row.ToString()).Contains(employee.LastName))
                        {
                            discipline.Coursework = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BF" + row.ToString()));
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BT" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BV" + row.ToString()).Contains(employee.LastName))
                        {
                            discipline.Tests = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BU" + row.ToString()), CultureInfo.InvariantCulture);
                        }
                        if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BW" + row.ToString()) != "0" &&
                            GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BY" + row.ToString()).Contains(employee.LastName))
                        {
                            discipline.Exam = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BX" + row.ToString()), CultureInfo.InvariantCulture);
                        }
                        employee.AutumnSemester.Disciplines.Add(discipline);
                    }
                    row++;
                }
                OnProgressChanged?.Invoke(2, null);
                #endregion
                
                OnStatusChanged?.Invoke("Сбор данных: Дополнительные данные", null);
                #region Создание доп. данных
                
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
                    if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()).Contains(employee.LastName))
                    {
                        bachelor.Diploms = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "L" + row.ToString()), CultureInfo.InvariantCulture);
                        if (bachelor.Diploms != 0) { employee.AutumnSemester.Disciplines.Add(bachelor); }
                        magister.Diploms = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "O" + row.ToString()), CultureInfo.InvariantCulture);
                        if (magister.Diploms != 0) { employee.AutumnSemester.Disciplines.Add(magister); }
                        diplPro.ConsultationsByDiplom = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BF" + row.ToString()), CultureInfo.InvariantCulture);
                        if (diplPro.ConsultationsByDiplom != 0) { employee.AutumnSemester.Disciplines.Add(diplPro); }
                        anotherWork.AnotherWork = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "BG" + row.ToString()), CultureInfo.InvariantCulture);
                        if (anotherWork.AnotherWork != 0) { employee.AutumnSemester.Disciplines.Add(anotherWork); }
                        aspirants.Aspirants = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "P" + row.ToString()), CultureInfo.InvariantCulture);
                        if (aspirants.Aspirants != 0) { employee.AutumnSemester.Disciplines.Add(aspirants); }
                        gekB.GEK = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "N" + row.ToString()), CultureInfo.InvariantCulture);
                        if (gekB.GEK != 0) { employee.AutumnSemester.Disciplines.Add(gekB); }
                        gekM.GEK = double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "M" + row.ToString()), CultureInfo.InvariantCulture);
                        if (gekM.GEK != 0) { employee.AutumnSemester.Disciplines.Add(gekM); }
                    }
                    row++;
                }
                OnProgressChanged?.Invoke(3, null);
                worksheetPart = (WorksheetPart)(doc.WorkbookPart.GetPartById(SheetDipl.Id));
                row = 2;
                while (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "B" + row.ToString()) != "0")
                {
                    if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "B" + row.ToString()).Contains(employee.LastName))
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
                OnProgressChanged?.Invoke(4, null);
                #endregion
                
                OnStatusChanged?.Invoke("Сбор данных: Практика", null);
                #region Практика
                
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
                            if (GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, "C" + row.ToString()).Contains(employee.LastName))
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
                                prac.PracticalWork += double.Parse(GetCellValue(worksheetPart.Worksheet, doc.WorkbookPart, (j == 1 ? "K" : "L") + row.ToString()), CultureInfo.InvariantCulture);
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
                OnProgressChanged?.Invoke(5, null);
                #endregion
                return employee;
            }
        }
        
        /// <summary>
        /// Создает структуру отчёта
        /// </summary>
        /// <param name="employee"> Преподаватель</param>
        /// <param name="worksheetPart"> Часть страницы</param>
        /// <param name="shareStringPart"> Таблица строк</param>
        /// <param name="cellFormats"> Стили ячеек</param>
        private void CreateRaport(Employee employee, WorksheetPart worksheetPart, SharedStringTablePart shareStringPart,
            List<ExcelCellFormat> cellFormats)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            
            OnStatusChanged?.Invoke("Создаётся структура документа", null);
            #region Создание ширины столбцов и соединений ячеек
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
            OnProgressChanged?.Invoke(8, null);
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
            OnProgressChanged?.Invoke(9, null);
            #endregion
            #region Создание Шапки документа

            uint Total = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.Total).Id;
            uint ColumnName = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.ColumnName).Id;
            uint ColumnNumber = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.ColumnNumber).Id;
            uint DisciplineCode = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.DisciplineCode).Id;
            uint DisciplineName = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.DisciplineName).Id;
            uint GroupPlan = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.GroupPlan).Id;
            uint ColumnTotal = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.ColumnTotal).Id;
            
            CellData[] cells =
            {
                new CellData(){Column = "A", Row = 1, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.UniveristyInfo).Id, Data = "федеральное государственное бюджетное образовательное учреждение высшего образования "},
                new CellData(){Column = "A", Row = 2, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.UniveristyInfo).Id, Data = "«Казанский национальный исследовательский технический университет им. А.Н. Туполева-КАИ» (КНИТУ-КАИ)"},
                new CellData(){Column = "M", Row = 4, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.Approve).Id, Data = "УТВЕРЖДАЮ"},
                new CellData(){Column = "D", Row = 5, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.Title).Id, Data = "ПЛАН УЧЕБНОЙ НАГРУЗКИ"},
                new CellData(){Column = "M", Row = 5, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.Position).Id, Data = "Зав. кафедрой ПМИ"},
                new CellData(){Column = "M", Row = 6, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.ManagerInfo).Id, Data = "Зайдуллин С.С."},
                new CellData(){Column = "O", Row = 7, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.ManagerInfoMeta).Id, Data = "подпись, ФИО"},
                new CellData(){Column = "C", Row = 7, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.EmployeeInfo).Id, Data = $"{employee.Rank}, {employee.LastName} {employee.FirstName} {employee.Patronymic}, {employee.StudyRank}, {employee.Rate}, {employee.Staffing}"},
                new CellData(){Column = "C", Row = 8, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.EmployeeInfoMeta).Id, Data = "должность, ФИО, ученая степень, ученое звание, доля ставки, штатность"},
                new CellData(){Column = "C", Row = 9, StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.Year).Id, Data = "на  2019 / 2020 учебный год"},

                new CellData(){Column = "A", StyleIndex = ColumnName, Row = 11, Data = "Код ОП,\nиндекс дисциплины,\nнаименование дисциплины"},
                new CellData(){Column = "B", StyleIndex = ColumnName, Row = 11, Data = ""},
                new CellData(){Column = "C", StyleIndex = ColumnName, Row = 11, Data = "Группа"},
                new CellData(){Column = "D", StyleIndex = ColumnName, Row = 11, Data = "Лекц"},
                new CellData(){Column = "E", StyleIndex = ColumnName, Row = 11, Data = "Практ"},
                new CellData(){Column = "F", StyleIndex = ColumnName, Row = 11, Data = "Лаб"},
                new CellData(){Column = "G", StyleIndex = ColumnName, Row = 11, Data = "Консульт. студ."},
                new CellData(){Column = "H", StyleIndex = ColumnName, Row = 11, Data = ""},
                new CellData(){Column = "I", StyleIndex = ColumnName, Row = 11, Data = "Руководство"},
                new CellData(){Column = "J", StyleIndex = ColumnName, Row = 11, Data = ""},
                new CellData(){Column = "K", StyleIndex = ColumnName, Row = 11, Data = ""},
                new CellData(){Column = "L", StyleIndex = ColumnName, Row = 11, Data = ""},
                new CellData(){Column = "M", StyleIndex = ColumnName, Row = 11, Data = "ГЭК"},
                new CellData(){Column = "N", StyleIndex = ColumnName, Row = 11, Data = "ЗАЧ"},
                new CellData(){Column = "O", StyleIndex = ColumnName, Row = 11, Data = "ЭКЗ"},
                new CellData(){Column = "P", StyleIndex = ColumnName, Row = 11, Data = "Другие  виды уч. работы"},
                new CellData(){Column = "Q", StyleIndex = Total, Row = 11, Data = " ВСЕГО"},
                new CellData(){Column = "A", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "B", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "C", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "D", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "E", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "F", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "G", StyleIndex = ColumnName, Row = 12, Data = "по теор. курсу"},
                new CellData(){Column = "H", StyleIndex = ColumnName, Row = 12, Data = "по дипл. проект."},
                new CellData(){Column = "I", StyleIndex = ColumnName, Row = 12, Data = "асп-ми"},
                new CellData(){Column = "J", StyleIndex = ColumnName, Row = 12, Data = "курс. проект. (раб.)"},
                new CellData(){Column = "K", StyleIndex = ColumnName, Row = 12, Data = "дипл. проект."},
                new CellData(){Column = "L", StyleIndex = ColumnName, Row = 12, Data = "практ."},
                new CellData(){Column = "M", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "N", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "O", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "P", StyleIndex = ColumnName, Row = 12, Data = ""},
                new CellData(){Column = "Q", StyleIndex = ColumnName, Row = 12, Data = ""},
            };
            foreach (var data in cells)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(data.Data, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = data.StyleIndex;
            }
            foreach (var column in Column)
            {
                Cell cell = InsertCellInWorksheet(column, 13, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = ColumnName;
            }
            for (uint i = 2; i < 17; i++)
            {
                Cell cell = InsertCellInWorksheet(Column[i], 14, worksheetPart);
                cell.CellValue = new CellValue(i.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = ColumnNumber;
                if (i == 2)
                {
                    cell = InsertCellInWorksheet("A", 14, worksheetPart);
                    cell.CellValue = new CellValue("1");
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    cell.StyleIndex = ColumnNumber;
                    cell = InsertCellInWorksheet("B", 14, worksheetPart);
                    cell.CellValue = new CellValue("");
                    cell.StyleIndex = ColumnNumber;
                }
            }
            OnProgressChanged?.Invoke(10, null);
            #endregion
            
            OnStatusChanged?.Invoke("Заполняются данные по осеннему семестру", null);
            #region Создание осеннего семестра
            
            mergeCells.Append(new MergeCell() { Reference = new StringValue("A15:Q15") });
            Cell semCell = InsertCellInWorksheet("A", 15, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("О  С  Е  Н  Н  И  Й     С  Е  М  Е  С  Т  Р  ", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.SemesterName).Id;

            uint row = 16;
            foreach (var discipline in employee.SpringSemester.Disciplines)
            {
                Cell cell = InsertCellInWorksheet("A", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.CodeOP, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = DisciplineCode;
                cell = InsertCellInWorksheet("B", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.Name, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = DisciplineName;
                cell = InsertCellInWorksheet("C", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(String.Join(", ", discipline.Groups), shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("D", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Lectures.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("E", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.PracticalWork.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("F", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.LaboratoryWork.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("G", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByTheory.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("H", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByDiplom.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("I", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Aspirants.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("J", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Coursework.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("K", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Diploms.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("L", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Practice.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("M", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.GEK.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("N", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Tests.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("O", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Exam.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("P", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.AnotherWork.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("Q", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.TotalForThisDiscipline().ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                row++;
            }
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:C{row}") });
            OnProgressChanged?.Invoke(11, null);
            #endregion
            
            OnStatusChanged?.Invoke("Рассчитываются итоговые значения по осеннему семестру", null);
            #region Итог Осеннего семестра
            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("Итого за осенний семестр", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.SemesterTotalLabel).Id;
            semCell = InsertCellInWorksheet("B", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = ColumnTotal;
            semCell = InsertCellInWorksheet("C", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = ColumnTotal;
            CellData[] totalS =
            {
                new CellData(){Column = "D", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForLectures().ToString()},
                new CellData(){Column = "E", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForPracticalWork().ToString()},
                new CellData(){Column = "F", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForLaboratoryWork().ToString()},
                new CellData(){Column = "G", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForConsultationsByTheory().ToString()},
                new CellData(){Column = "H", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForConsultationsByDiplom().ToString()},
                new CellData(){Column = "I", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForAspirants().ToString()},
                new CellData(){Column = "J", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForCoursework().ToString()},
                new CellData(){Column = "K", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForDiploms().ToString()},
                new CellData(){Column = "L", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForPractice().ToString()},
                new CellData(){Column = "M", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForGEK().ToString()},
                new CellData(){Column = "N", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForTests().ToString()},
                new CellData(){Column = "O", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForExam().ToString()},
                new CellData(){Column = "P", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForAnotherWork().ToString()},
                new CellData(){Column = "Q", Row = row, StyleIndex = ColumnTotal, Data = employee.SpringSemester.TotalForSemester().ToString()}
            };
            foreach (var data in totalS)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(data.Data);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = data.StyleIndex;
            }
            totalS = null;
            OnProgressChanged?.Invoke(12, null);
            #endregion
            
            OnStatusChanged?.Invoke("Заполняются данные по весеннему семестру", null);
            #region Создание весеннего семестра
            
            row++;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:Q{row}") });
            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("В  Е  С  Е  Н  Н  И  Й     С  Е  М  Е  С  Т  Р  ", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.SemesterName).Id;
            row++;

            foreach (var discipline in employee.AutumnSemester.Disciplines)
            {
                Cell cell = InsertCellInWorksheet("A", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.CodeOP, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = DisciplineCode;
                cell = InsertCellInWorksheet("B", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(discipline.Name, shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = DisciplineName;
                cell = InsertCellInWorksheet("C", row, worksheetPart);
                cell.CellValue = new CellValue(InsertSharedStringItem(String.Join(", ", discipline.Groups), shareStringPart).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("D", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Lectures.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("E", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.PracticalWork.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("F", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.LaboratoryWork.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("G", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByTheory.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("H", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.ConsultationsByDiplom.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("I", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Aspirants.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("J", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Coursework.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("K", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Diploms.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("L", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Practice.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("M", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.GEK.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("N", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Tests.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("O", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.Exam.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("P", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.AnotherWork.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                cell = InsertCellInWorksheet("Q", row, worksheetPart);
                cell.CellValue = new CellValue(discipline.TotalForThisDiscipline().ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = GroupPlan;
                row++;
            }
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:C{row}") });
            OnProgressChanged?.Invoke(13, null);
            #endregion
            
            OnStatusChanged?.Invoke("Рассчитываются итоговые значения по весеннему семестру", null);
            #region Итог Осеннего семестра

            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("Итого за весенний семестр", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.SemesterTotalLabel).Id;
            semCell = InsertCellInWorksheet("B", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = ColumnTotal;
            semCell = InsertCellInWorksheet("C", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = ColumnTotal;
            CellData[] totalA =
            {
                new CellData(){Column = "D", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForLectures().ToString()},
                new CellData(){Column = "E", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForPracticalWork().ToString()},
                new CellData(){Column = "F", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForLaboratoryWork().ToString()},
                new CellData(){Column = "G", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForConsultationsByTheory().ToString()},
                new CellData(){Column = "H", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForConsultationsByDiplom().ToString()},
                new CellData(){Column = "I", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForAspirants().ToString()},
                new CellData(){Column = "J", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForCoursework().ToString()},
                new CellData(){Column = "K", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForDiploms().ToString()},
                new CellData(){Column = "L", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForPractice().ToString()},
                new CellData(){Column = "M", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForGEK().ToString()},
                new CellData(){Column = "N", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForTests().ToString()},
                new CellData(){Column = "O", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForExam().ToString()},
                new CellData(){Column = "P", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForAnotherWork().ToString()},
                new CellData(){Column = "Q", Row = row, StyleIndex = ColumnTotal, Data = employee.AutumnSemester.TotalForSemester().ToString()}
            };
            foreach (var data in totalA)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(data.Data);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = data.StyleIndex;
            }
            totalA = null;
            row++;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"A{row}:C{row}") });
            OnProgressChanged?.Invoke(14, null);
            #endregion
            
            OnStatusChanged?.Invoke("Рассчитываются итоговые значения", null);
            #region Итог
            semCell = InsertCellInWorksheet("A", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("ВСЕГО ЗА ГОД", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = Total;
            semCell = InsertCellInWorksheet("B", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = Total;
            semCell = InsertCellInWorksheet("C", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = Total;
            CellData[] total =
            {
                new CellData(){Column = "D", Row = row, StyleIndex = Total, Data = employee.LecturesForYear().ToString()},
                new CellData(){Column = "E", Row = row, StyleIndex = Total, Data = employee.PracticalWorkForYear().ToString()},
                new CellData(){Column = "F", Row = row, StyleIndex = Total, Data = employee.LaboratoryWorkForYear().ToString()},
                new CellData(){Column = "G", Row = row, StyleIndex = Total, Data = employee.ConsultationsByTheoryForYear().ToString()},
                new CellData(){Column = "H", Row = row, StyleIndex = Total, Data = employee.ConsultationsByDiplomForYear().ToString()},
                new CellData(){Column = "I", Row = row, StyleIndex = Total, Data = employee.AspirantsForYear().ToString()},
                new CellData(){Column = "J", Row = row, StyleIndex = Total, Data = employee.CourseworkForYear().ToString()},
                new CellData(){Column = "K", Row = row, StyleIndex = Total, Data = employee.DiplomsForYear().ToString()},
                new CellData(){Column = "L", Row = row, StyleIndex = Total, Data = employee.PracticeForYear().ToString()},
                new CellData(){Column = "M", Row = row, StyleIndex = Total, Data = employee.GakForYear().ToString()},
                new CellData(){Column = "N", Row = row, StyleIndex = Total, Data = employee.TestsForYear().ToString()},
                new CellData(){Column = "O", Row = row, StyleIndex = Total, Data = employee.ExamForYear().ToString()},
                new CellData(){Column = "P", Row = row, StyleIndex = Total, Data = employee.AnotherWorkForYear().ToString()},
                new CellData(){Column = "Q", Row = row, StyleIndex = Total, Data = employee.Year().ToString()}
            };
            foreach (var data in total)
            {
                Cell cell = InsertCellInWorksheet(data.Column, data.Row, worksheetPart);
                cell.CellValue = new CellValue(data.Data);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = data.StyleIndex;
            }
            total = null;
            row += 2;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"K{row}:O{row}") });
            semCell = InsertCellInWorksheet("K", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("_____________________", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            row++;
            mergeCells.Append(new MergeCell() { Reference = new StringValue($"K{row}:O{row}") });
            semCell = InsertCellInWorksheet("K", row, worksheetPart);
            semCell.CellValue = new CellValue(InsertSharedStringItem("подпись преподавателя", shareStringPart).ToString());
            semCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            semCell.StyleIndex = cellFormats.FirstOrDefault(c => c.CellFormatType == ExcelCellFormats.TeacherSignature).Id;
            OnProgressChanged?.Invoke(15, null);
            #endregion
        }

        /// <summary>
        /// Создает отчёт в отдельный файл
        /// </summary>
        /// <param name="path"> Путь к итоговому файлу</param>
        /// <param name="employee"> Преподаватель</param>
        public void CreateRaportSeparate(string path, Employee employee)
        {
            OnStatusChanged?.Invoke("Создаётся документ", null);
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
            OnProgressChanged?.Invoke(6, null);
            
            OnStatusChanged?.Invoke("Загружаются стили", null);
            spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet = new Stylesheet()
            {
                Borders = new Borders(),
                Fonts = new Fonts(),
                Fills = new Fills(),
                CellFormats = new CellFormats()
            };
            ExcelStylesheetBuilder builder = new ExcelStylesheetBuilder(0, 0, 0);
            ExcelStylesheetDirector director = new ExcelStylesheetDirector() { StylesheetBuilder = builder };
            director.BuildReportStylesheet();
            var reportStylesheet = builder.GetStylesheet();
            AppendStylesToDocument(spreadsheetDocument, reportStylesheet);
            OnProgressChanged?.Invoke(7, null);

            CreateRaport(employee, worksheetPart, shareStringPart, reportStylesheet.CellFormats);
            
            OnStatusChanged?.Invoke("Сохранение документа", null);
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
            OnProgressChanged?.Invoke(16, null);
        }

        /// <summary>
        /// Создает отчёт в файле с данными
        /// </summary>
        /// <param name="path"> Путь к файлу</param>
        /// <param name="employee"> Преподаватель</param>
        public void CreateRaportInFile(string path, Employee employee)
        {
            OnStatusChanged?.Invoke("Открывается документ", null);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, true))
            {
                OnProgressChanged?.Invoke(6, null);
                
                OnStatusChanged?.Invoke("Загружается документ", null);
                if (!AreIndexesSame(doc, out List<ExcelCellFormat> cellFormats))
                {
                    InitStyles(doc, out cellFormats);
                }
                SharedStringTablePart shareStringPart;
                if (doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = doc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                var worksheetPart = GetSheet(doc.WorkbookPart, employee.LastName + " " + employee.FirstName[0] + "." + employee.Patronymic[0] + ".");
                OnProgressChanged?.Invoke(7, null);

                CreateRaport(employee, worksheetPart, shareStringPart, cellFormats);
                
                OnStatusChanged?.Invoke("Сохраняется документ", null);
                doc.WorkbookPart.Workbook.Save();
                OnProgressChanged?.Invoke(16, null);
            }
        }
    }
}
