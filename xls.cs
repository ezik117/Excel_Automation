using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;


/*
ИМПОРТИРОВАТЬ БИБЛИОТЕКУ ОФИСА ПО СЛЕДУЮЩЕМУ ПУТИ
C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\14.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll
*/


/// <summary>
/// Класс для автоматизации работы с объектом Excel.Application.
/// Раннее связывание.
/// </summary>
class xls : IDisposable
{
    /// <summary>
    /// Основной объект приложения
    /// </summary>
    private Excel.Application x = null;

    /// <summary>
    /// Рабочая книга
    /// </summary>
    private Excel.Workbook w;

    /// <summary>
    /// Лист
    /// </summary>
    private Excel.Worksheet s;

    /// <summary>
    /// Возвращает лист рабочей книги
    /// </summary>
    public Excel.Worksheet sheet
    {
        get
        {
            return s;
        }
    }

    /// <summary>
    /// Возвращает рабочую книгу
    /// </summary>
    public Excel.Workbook workbook
    {
        get
        {
            return w;
        }
    }

    // может проверятся после каждой операции. 0 - вызов функции не вызывал ошибок в коде. Иначе код ошибки.
    public int lastError;


    /// <summary>
    /// Перечисление всех возможных выравнивателей. В зависимости от функции будет преобразовываться в разные значения Excel
    /// </summary>
    public enum TAlignment
    {
        Top,
        Bottom,
        Left,
        Right,
        Center
    }

    /// <summary>
    /// Для некоторых функций манипулирования стилем задает стиль объекта к которому применяется
    /// </summary>
    public enum TObjectType
    {
        Cell,
        Row,
        Column,
        Selection
    }

    /// <summary>
    /// Для некоторых функций стиля задает положение объекта (например, границы ячейки)
    /// </summary>
    public enum TBorderPosition
    {
        Top,
        Left,
        Right,
        Bottom,
        All
    }

    /// <summary>
    /// Толщина границы
    /// </summary>
    public enum TBorderWeight
    {
        Thick,
        Thin
    }

    /// <summary>
    /// Стили текста
    /// </summary>
    public enum TFontStyle
    {
        bold,
        italic,
        underline,
        normal
    }

    /// <summary>
    /// Формат копирования
    /// </summary>
    public enum TPasteType
    {
        normal,
        keepWidth,
        format
    }

    /// <summary>
    /// Открывает указанный файл и делает активным первый лист
    /// </summary>
    /// <param name="filename">Имя файла</param>
    /// <param name="setVisible">Указывает должен ли открываемый файл быть видим пользователю (false по умолчанию)</param>
    /// <param name="minimized">Указывает должен ли открываемый файл быть минимизирован (false по умолчанию)</param>
    /// <param name="autoupdateLinks">Автоматическое обновление ссылок (false по умолчанию)</param>
    // конструктор.
    public xls(string filename, bool setVisible = false, bool minimized = false, bool autoupdateLinks = false, bool readOnly = false)
    {
        lastError = 0;

        try
        {
            x = new Excel.Application();
            x.Visible = setVisible;
            x.WindowState = (minimized ? Excel.XlWindowState.xlMinimized : Excel.XlWindowState.xlNormal);
            //x.EnableEvents = false;
            //x.ScreenUpdating = false

            w = x.Workbooks.Open(filename, autoupdateLinks, readOnly);
            SetActiveWorksheet(1);
        }
        catch
        {
            lastError = 1;
        }
    }

    /// <summary>
    /// Создает новую книгу и делает активным первый лист
    /// </summary>
    /// <param name="setVisible">Указывает должна ли книга быть видима пользователю (true по умолчанию)</param>
    /// <param name="minimized">Указывает должена ли книга быть минимизирована (false по умолчанию)</param>
    public xls(bool setVisible = true, bool minimized = false)
    {
        lastError = 0;

        try
        {
            x = new Excel.Application();
            x.Visible = setVisible;
            x.WindowState = (minimized ? Excel.XlWindowState.xlMinimized : Excel.XlWindowState.xlNormal);

            w = x.Workbooks.Add();
            SetActiveWorksheet(1);
        }
        catch
        {
            lastError = 1;
        }
    }

    /// <summary>
    /// Устанавливает активный лист по имени
    /// </summary>
    /// <param name="name">Имя листа с учетам регистра</param>
    /// <returns>True, если лист с указанным именем найден и открыт. Иначе False</returns>
    public bool SetActiveWorksheet(string name)
    {
        bool res = true;
        lastError = 0;

        try
        {
            s = ((Excel.Worksheet)w.Sheets[name]);
            s.Activate();
        }
        catch
        {
            res = false;
            lastError = 1;
        }

        return res;
    }

    /// <summary>
    /// Устанавливает активный лист по номеру
    /// </summary>
    /// <param name="index">Номер листа в коллекции листов по порядку слева-напрао, начиная с 0</param>
    /// <returns>True, если лист с указанным именем найден и открыт. Иначе False</returns>
    public bool SetActiveWorksheet(int index)
    {
        bool res = true;
        lastError = 0;

        try
        {
            s = ((Excel.Worksheet)w.Sheets[index]);
            s.Activate();
        }
        catch
        {
            res = false;
            lastError = 1;
        }

        return res;
    }

    /// <summary>
    /// Минимизирует или восстанавливает окно
    /// </summary>
    /// <param name="minimized">Если True-то окно минимизируется, иначе восстанавливается</param>
    public void minimizeExcel(bool minimized)
    {
        x.WindowState = (minimized ? Excel.XlWindowState.xlMinimized : Excel.XlWindowState.xlNormal);
    }

    /// <summary>
    /// Сохраняет текущую книгу с указанным именем
    /// </summary>
    /// <param name="filename">Полный путь с именем файла</param>
    public void saveWorkbookAs(string filename)
    {
        x.DisplayAlerts = false;
        w.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
        x.DisplayAlerts = true;
    }

    /// <summary>
    /// Включает или выключает обработку событий в Excel
    /// </summary>
    /// <param name="enable">True - включить обработку, False - выключить</param>
    public void enableEvents(bool enable)
    {
        x.EnableEvents = enable;
    }

    /// <summary>
    /// Включает или выключает обновление экрана
    /// </summary>
    /// <param name="enable">True - включить обработку, False - выключить</param>
    public void enableScreenUpdating(bool enable)
    {
        x.ScreenUpdating = enable;
    }

    // получает данные из ячейки
    /// <summary>
    /// Возвращает значение ячейки текущего листа
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <returns>Объект значения ячейки в зависимости от ее типа. Для объединенных ячеек возвращает объект MergedArea</returns>
    public object getCellValue(int row, int col)
    {
        lastError = 0;
        var o = s.Cells[row, col] as Excel.Range;

        try
        {
            if (o.MergeCells)
            {
                o = o.MergeArea[1, 1];
            }
        }
        catch
        {
            lastError = 1;
        }

        return (lastError == 1 ? null : o.Value2);
    }

    /// <summary>
    /// Возвращает строковое значение ячейки текущего листа без начальных и конечных пробелов
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <returns>Строка</returns>
    public string getCellValueS(int row, int col)
    {
        lastError = 0;

        try
        {
            var o = s.Cells[row, col];
            if (o == null) return string.Empty;
            o = (o as Excel.Range).Value2;
            if (o == null) return string.Empty;

            return Convert.ToString(o).Trim();
        }
        catch
        {
            lastError = 1;
        }

        return "";
    }

    /// <summary>
    /// Возвращает видимый текст ячейки. Если ячейка скрытая возвращает пустое значение.
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <returns>Текстовое значение ячейки</returns>
    public string getCellValueT(int row, int col)
    {
        return s.Cells[row, col].Text;
    }


    /// <summary>
    /// Возвращает формат ячейки как тектовое значение.
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <returns>Строка формата ячейки</returns>
    public string getCellNumberFormat(int row, int col)
    {
        return s.Cells[row, col].NumberFormat.ToString();
    }

    /// <summary>
    /// Присваивает ячейке значение. Ячейка приобретает тип переданного значения
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <param name="value">Объект любого типа</param>
    public void setCellValue(int row, int col, object value)
    {
        lastError = 0;

        try
        {
            s.Cells[row, col] = value;
        }
        catch
        {
            lastError = 1;
        }
    }

    /// <summary>
    /// Возвращает коллекцию всех листов в книге
    /// </summary>
    /// <param name="capitalized">Если true - то список будет приведен к заглавным буквам, иначе - как есть.</param>
    /// <returns>Массив строк с именами листов</returns>
    public string[] getSheets(bool capitalized = false, bool onlyVisible = false)
    {
        lastError = 0;

        try
        {
            List<string> sheets = new List<string>();
            foreach (Excel.Worksheet ws in w.Worksheets)
            {
                if (onlyVisible && (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible))
                    sheets.Add((capitalized ? ws.Name.ToUpper() : ws.Name));
            }

            return sheets.ToArray();
        }
        catch
        {
            lastError = 1;
        }

        return null;
    }

    /// <summary>
    /// Возвращает имя активного листа
    /// </summary>
    /// <returns>Строка с именем</returns>
    public string getActiveSheetName()
    {
        lastError = 0;

        try
        {
            return s.Name;
        }
        catch
        {
            lastError = 1;
        }

        return "";
    }

    /// <summary>
    /// Устанавливает имя активного листа
    /// </summary>
    /// <param name="name">Новое имя листа</param>
    public void setActiveSheetName(string name)
    {
        s.Name = name;
    }

    /// <summary>
    /// Конвертирует буквенное значение столбца в числовое, например "А" = 1, "АА" = 27
    /// </summary>
    /// <param name="columnName"></param>
    /// <returns>Числовой номер столбца</returns>
    public int CtoI(string columnName)
    {
        int ret = 0;
        lastError = 0;

        try
        {
            columnName = columnName.ToUpper();

            // реверс строки
            char[] tmp = columnName.ToArray();
            Array.Reverse(tmp);

            //перемножаем
            for (int i = 0; i < tmp.Length; i++)
            {
                ret += (int)Math.Pow(26, i) * (1 + tmp[i] - 'A');
            }

            return ret;
        }
        catch
        {
            lastError = 1;
        }

        return -1;
    }

    /// <summary>
    /// Конвертирует числовое значение столбца в буквенное, например 1 = "А", 27 = "АА"
    /// </summary>
    /// <param name="columnNumber">Номер столбца, начинается с 1</param>
    /// <returns>Строковое значение столбца</returns>
    private string CtoA(int columnNumber)
    {
        string columnName = "";

        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }

    /// <summary>
    /// Автоматически расширяет по содержимому все столбцы и/или строки.
    /// </summary>
    /// <param name="fitColumns">Если True, то расширяет столбцы</param>
    /// <param name="fitRows">Если True, то расширяет строки</param>
    public void fitAll(bool fitColumns = true, bool fitRows = true)
    {
        if (fitColumns) s.Columns.AutoFit();
        if (fitRows) s.Rows.AutoFit();
    }


    /// <summary>
    /// Устанавливает цвет ячейки
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <param name="color">Цвет фона</param>
    public void setCellBgColor(int row, int col, Color color)
    {
        s.Cells[row, col].Interior.Color = ColorTranslator.ToOle(color);
    }

    /// <summary>
    /// Устанавливает цвет текста
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <param name="color">Цвет текста</param>
    public void setCellFontColor(int row, int col, Color color)
    {
        s.Cells[row, col].Font.Color = ColorTranslator.ToOle(color);
    }

    /// <summary>
    /// Задает размер шрифта
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <param name="size">Размер шрифта</param>
    public void setCellFontSize(int row, int col, float size)
    {
        s.Cells[row, col].Font.Size = size;
    }

    /// <summary>
    /// Выводит простую границу в выбранный объект
    /// </summary>
    /// <param name="o">Тип объекта для которого применяется действие (выделение, строка, столбец)</param>
    /// <param name="row">Номер строки (применятся к o = Cell, Row)</param>
    /// <param name="col">Номер столбца (применяется к o = Cell, Column)</param>
    public void setCellBorder(TObjectType o, int row = 1, int col = 1)
    {
        Excel.Range r = s.Cells[row, col];

        if (o == TObjectType.Selection)
            r = x.Selection;
        else if (o == TObjectType.Column)
            r = s.Columns[col];
        else if (o == TObjectType.Row)
            r = s.Rows[row];

        r.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
        r.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
        r.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
        r.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
    }

    /// <summary>
    /// Устанавливает цвет для подстроки в ячейке
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <param name="start">Номер первого символа подстроки, начинается с 1</param>
    /// <param name="length">Длина подстроки</param>
    /// <param name="color">Цвет текста</param>
    public void colorizeInnerText(int row, int col, int start, int length, Color color)
    {
        s.Cells[row, col].Characters[start, length].Font.Color = ColorTranslator.ToOle(color);
    }

    /// <summary>
    /// Считывает цвет ячейки
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <returns>Цвет фона</returns>
    public Color getCellBgColor(int row, int col)
    {
        int color = (int)s.Cells[row, col].Interior.Color;
        return ColorTranslator.FromOle(color);
    }

    /// <summary>
    /// Объединяет ячейки.
    /// </summary>
    /// <param name="row1">Строка первой ячейки, начинается с 1</param>
    /// <param name="col1">Столбец первой ячейки, начинается с 1</param>
    /// <param name="row2">Строка второй ячейки, начинается с 1</param>
    /// <param name="col2">Столбец второй ячейки, начинается с 1</param>
    public void mergeCells(int row1, int col1, int row2, int col2)
    {
        selectRange(row1, col1, row2, col2);
        x.Selection.Merge();
    }

    /// <summary>
    /// Копирует ячейки в буфер обмена.
    /// </summary>
    /// <param name="row1">Строка первой ячейки, начинается с 1</param>
    /// <param name="col1">Столбец первой ячейки, начинается с 1</param>
    /// <param name="row2">Строка второй ячейки, начинается с 1</param>
    /// <param name="col2">Столбец второй ячейки, начинается с 1</param>
    public void copyCells(int row1, int col1, int row2, int col2)
    {
        selectRange(row1, col1, row2, col2);
        x.Selection.Copy();
    }

    /// <summary>
    /// Вставляет содержимое буфера обмена в текущий лист
    /// </summary>
    /// <param name="row">Строка в которую копируем, начинается с 1</param>
    /// <param name="col">Столбец в который копируем, начинается с 1</param>
    public void pasteCells(int row, int col, TPasteType t)
    {
        switch (t)
        {
            case TPasteType.normal:
                s.Paste(s.Cells[row, col]);
                break;
            case TPasteType.keepWidth:
                s.Cells[row, col].Select();
                s.Cells[row, col].PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths,
                                                Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                //s.Paste();
                break;
            case TPasteType.format:
                s.Cells[row, col].Select();
                s.Cells[row, col].PasteSpecial(Excel.XlPasteType.xlPasteFormats,
                                                Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                break;
        }
    }

    public Excel.FormatCondition copyConditionalFormat(int row, int col)
    {
        return s.Cells[row, col].FormatConditions;
    }

    public void pasteConditionalFormat(int row, int col, Excel.FormatCondition fc)
    {
        s.Cells[row, col].FormatConditions = fc;
    }

    /// <summary>
    /// Устанавливает вертикальное выравнивание содержимого в ячейке
    /// </summary>
    /// <param name="o">Объект для которого применяется действие (выделение, строка, столбец)</param>
    /// <param name="align">Тип выравнивания (по верху, по низу, по центру)</param>
    /// <param name="row">Номер строки (применятся к o = Cell, Row)</param>
    /// <param name="col">Номер столбца (применяется к o = Cell, Column)</param>
    public void setVerticalAlignment(TObjectType o, TAlignment align, int row = 1, int col = 1)
    {
        Excel.Range r = s.Cells[row, col];

        if (o == TObjectType.Selection)
            r = x.Selection;
        else if (o == TObjectType.Column)
            r = s.Columns[col];
        else if (o == TObjectType.Row)
            r = s.Rows[row];

        if (align == TAlignment.Top)
            r.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
        else if (align == TAlignment.Bottom)
            r.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
        else if (align == TAlignment.Center)
            r.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
    }


    /// <summary>
    /// Считывает ширину столбца в юнитах Excel.
    /// </summary>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <returns></returns>
    public float getColumnWidth(int col)
    {
        return (float)s.Columns[col].ColumnWidth;
    }

    /// <summary>
    /// Устанавливает ширину столбца в юнитах Excel.
    /// </summary>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <param name="width">Ширина столбца в юнитах Excel.</param>
    public void setColumnWidth(int col, float width)
    {
        s.Columns[col].ColumnWidth = width;
    }


    /// <summary>
    /// Установка стиля шрифта (жирный, курсив, подчеркнутый)
    /// </summary>
    /// <param name="row">Строка, начинается с 1</param>
    /// <param name="col">Столбец, начинается с 1</param>
    /// <param name="style">Стиль шрифта. Normal отменяет все шрифты.</param>
    public void setCellFontStyle(int row, int col, TFontStyle style)
    {
        switch (style)
        {
            case TFontStyle.bold:
                s.Cells[row, col].Font.Bold = true;
                break;
            case TFontStyle.italic:
                s.Cells[row, col].Font.Italic = true;
                break;
            case TFontStyle.underline:
                s.Cells[row, col].Font.Underline = true;
                break;
            case TFontStyle.normal:
                s.Cells[row, col].Font.Bold = false;
                s.Cells[row, col].Font.Italic = false;
                s.Cells[row, col].Font.Underline = false;
                break;
        }

    }

    /// <summary>
    /// Устанавливает горизонтальное выравнивание содержимого ячейки 
    /// </summary>
    /// <param name="o">Тип выделения для которого применяется действие (выделение, строка, столбец)</param>
    /// <param name="align">Тип выравнивания (по левому краю, по центру, по правому краю</param>
    /// <param name="row">Номер строки (применятся к o = Cell, Row)</param>
    /// <param name="col">Номер столбца (применяется к o = Cell, Column)</param>
    public void setHorizontalAlignment(TObjectType o, TAlignment align, int row = 1, int col = 1)
    {
        Excel.Range r = s.Cells[row, col];

        if (o == TObjectType.Selection)
            r = x.Selection;
        else if (o == TObjectType.Column)
            r = s.Columns[col];
        else if (o == TObjectType.Row)
            r = s.Rows[row];

        if (align == TAlignment.Left)
            r.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        else if (align == TAlignment.Right)
            r.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
        else if (align == TAlignment.Center)
            r.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
    }

    /// <summary>
    /// Выделяет указанный диапазон ячеек
    /// </summary>
    /// <param name="row1">Строка первой ячейки, начинается с 1</param>
    /// <param name="col1">Столбец первой ячейки, начинается с 1</param>
    /// <param name="row2">Строка второй ячейки, начинается с 1</param>
    /// <param name="col2">Столбец второй ячейки, начинается с 1</param>
    public void selectRange(int row1, int col1, int row2, int col2)
    {
        s.Range[s.Cells[row1, col1], s.Cells[row2, col2]].Select();
    }

    /// <summary>
    /// Группирует ячейки, строки или столбцы.
    /// Если номера столбцов равны нулю, то группируются строки.
    /// Если номера строк равны нулю, то группируются столбцы.
    /// </summary>
    /// <param name="row1">Номер первой группируемой строки, начинается с 1</param>
    /// <param name="col1">Номер первого группируемого столбца, начинается с 1</param>
    /// <param name="row2">Номер второй группируемой строки, начинается с 1</param>
    /// <param name="col2">Номер второго группируемого столбца, начинается с 1</param>
    public void groupRange(int row1, int col1, int row2, int col2)
    {
        if (col1 == 0 && col2 == 0)
        {
            s.Range[$"{row1}:{row2}"].Group(Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }
        else if (row1 == 0 && row2 == 0)
        {
            s.Range[$"{CtoA(col1)}:{CtoA(col2)}"].Group(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        else
        {
            s.Range[s.Cells[row1, col1], s.Cells[row2, col2]].Group(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
    }

    /// <summary>
    /// Раскрывает группы на указанный уровень.
    /// Уровень равный 1, сворачивает данные.
    /// Если rowLevel = 0, то раскрытие выполняется для столбцов.
    /// Если colLevel = 0, то раскрытие выполняется для строк.
    /// </summary>
    /// <param name="rowLevel">Уровень раскрытия строк, начиная с 1</param>
    /// <param name="colLevel">ровень раскрытия столбцов, начиная с 1</param>
    public void groupsShow(int rowLevel, int colLevel=0)
    {
        if (colLevel == 0)
            s.Outline.ShowLevels(RowLevels: rowLevel);
        else if (rowLevel == 0)
            s.Outline.ShowLevels(ColumnLevels: colLevel);
        else
            s.Outline.ShowLevels(rowLevel, colLevel);
    }

    /// <summary>
    /// Выбирает все ячейки на листе
    /// </summary>
    public void selectAll()
    {
        //x.Columns.Select();
        x.Cells.Select();
    }

    /// <summary>
    /// Удаляет ячейку, строку или столбец
    /// </summary>
    /// <param name="o">Выделение для которого применяется действие (выделение, строка или столбец)</param>
    /// <param name="row">Номер строки (применятся к o = Cell, Row)</param>
    /// <param name="col">Номер столбца (применяется к o = Cell, Column)</param>
    public void deleteCells(TObjectType o, int row = 1, int col = 1)
    {
        Excel.Range r = s.Cells[row, col];

        if (o == TObjectType.Selection)
            r = x.Selection;
        else if (o == TObjectType.Column)
            r = s.Columns[col];
        else if (o == TObjectType.Row)
            r = s.Rows[row];

        r.Delete();
    }

    /// <summary>
    /// Устанавливает режим автоматического переноса текста
    /// </summary>
    /// <param name="value">True - перенос разрешен, False-отменен</param>
    /// <param name="o">Тип выделения: выделение, ячейка, строка, столбец</param>
    /// <param name="row">Номер строки (применятся к o = Cell, Row)</param>
    /// <param name="col">Номер столбца (применяется к o = Cell, Column)</param>
    public void setWrapText(bool value, TObjectType o, int row = 1, int col = 1)
    {
        if (o == TObjectType.Selection)
            x.Selection.WrapText = value;
        if (o == TObjectType.Cell)
            s.Cells[row, col].WrapText = value;
        if (o == TObjectType.Column)
            x.Columns[col].WrapText = value;
        if (o == TObjectType.Row)
            x.Rows[row].WrapText = value;
    }

    /// <summary>
    /// Закрытие книги с выходом из Excel
    /// </summary>
    /// <param name="promptSavingChanges">True, если нужно сделать запрос сохранения несохраненных денных, False - закрыть без сохранения</param>
    public void CloseWorkbook(bool promptSavingChanges)
    {
        w.Close(promptSavingChanges, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
    }


    /// <summary>
    /// Закрытие книги без выхода из Excel. Оставляет открытую книгу для дальнейших действий пользователя.
    /// </summary>
    public void QuitWithoutClosing()
    {
        disposed = true;
    }

    // деструктор
    ~xls()
    {
        Dispose(false);
    }

    // реализация disposable
    private bool disposed = false;

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposed)
        {
            if (disposing)
            {
                // здесь указываем все свои ресурсы
                if (x != null) x.Quit();
            }
            // Release unmanaged resources.
            disposed = true;
        }
    }
}

