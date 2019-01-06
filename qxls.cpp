#include "qxls.h"

/**
=========================================================================
                        QXLSAPPLICATION REALIZATION CODE
=========================================================================
**/

QXlsApplication::QXlsApplication(QObject *parent) : QObject(parent),
    m_ExcelApplication(new QAxObject("Excel.Application" ,this)),
    m_WorkbooksObject(m_ExcelApplication->querySubObject("Workbooks"))
{
}

QXlsApplication::~QXlsApplication()
{
    foreach (QXlsWorkbook *WB, m_Workbooks) {
        delete WB;
    }
    m_Workbooks.clear();
    delete m_WorkbooksObject;
    delete m_ExcelApplication;
}

QXlsWorkbook *QXlsApplication::workbook(int index) const
{
    return m_Workbooks.at(index);
}

int QXlsApplication::workbooksCount() const
{
    return m_WorkbooksObject->dynamicCall("Count").toInt();
}

QXlsWorkbook *QXlsApplication::open(const QString &filepath)
{
    QAxObject *WorkbookObj = m_WorkbooksObject->querySubObject( "Open(const QString&)", filepath);
    if(WorkbookObj) {
        QXlsWorkbook* Workbook = new QXlsWorkbook(this, WorkbookObj);
        m_Workbooks.append(Workbook);
        return Workbook;
    }
    return 0;
}

void QXlsApplication::alerts(bool state)
{
    m_ExcelApplication->setProperty("DisplayAlerts", state);
}

void QXlsApplication::exit()
{
    m_ExcelApplication->dynamicCall("Quit()");
}

void QXlsApplication::show()
{
    m_ExcelApplication->dynamicCall("SetVisible(bool)", true);
}

void QXlsApplication::hide()
{
    m_ExcelApplication->dynamicCall("SetVisible(bool)", false);
}

QXlsWorkbook *QXlsApplication::newWorkbook()
{
    QAxObject* newWorkbookObj = m_WorkbooksObject->querySubObject("Add");
    QXlsWorkbook* Workbook = new QXlsWorkbook(this, newWorkbookObj);
    m_Workbooks.append(Workbook);
    return Workbook;
}

QXlsWorkbook *QXlsApplication::activeWorkbook() const
{
    if(workbooksCount() != 0)
    {
        QAxObject* activeRef = m_ExcelApplication->querySubObject("Activeworkbook");
        QString refName = activeRef->property("Name").toString();
        for(int i = 0; i < workbooksCount(); i++)
            if(m_Workbooks.at(i)->name() == refName) return m_Workbooks.at(i);
    }
    return 0;
}

QString QXlsApplication::documentation() const
{
    return m_ExcelApplication->generateDocumentation();
}

/**
=========================================================================
                        END OF QXLSAPPLICATION REALIZATION CODE
=========================================================================
**/

/**
=========================================================================
                        QXLSWORKBOOK REALIZATION CODE
=========================================================================
**/

QXlsWorkbook::QXlsWorkbook(QObject *parent, QAxObject* RefObject) : QObject(parent),
    m_Workbook(RefObject),
    m_WorksheetsObject(m_Workbook->querySubObject("Worksheets"))
{
    m_count = m_WorksheetsObject->dynamicCall("Count").toInt();
    for(int i = 1; i <= m_count; i++) {
        QAxObject* sheetReference = m_WorksheetsObject->querySubObject("item(int)", QVariant(i));
        m_Worksheets.append(new QXlsWorksheet(this, sheetReference));
    }
}

bool QXlsWorkbook::isActive()
{
    return m_active;
}

QString QXlsWorkbook::name() const
{
    return m_Workbook->property("Name").toString();
}

int QXlsWorkbook::worksheetsCount() const
{
    return m_count;
}

QXlsWorksheet* QXlsWorkbook::newWorksheet()
{
    QAxObject* sheetReference = m_WorksheetsObject->querySubObject("Add");
    QXlsWorksheet* sheet = new QXlsWorksheet(this, sheetReference);
    m_Worksheets.append(sheet);
    m_count = m_Worksheets.size();
    return sheet;
}

QXlsWorksheet *QXlsWorkbook::sheet(int index) const
{
    return m_Worksheets.at(index);
}

QXlsWorkbook::~QXlsWorkbook()
{
    foreach(QXlsWorksheet *WS, m_Worksheets) {
        delete WS;
    }
    m_Worksheets.clear();
    delete m_WorksheetsObject;
    delete m_Workbook;
}

void QXlsWorkbook::close()
{
    m_Workbook->dynamicCall("Close()");
}

void QXlsWorkbook::save()
{
    m_Workbook->dynamicCall("Save()");
}

void QXlsWorkbook::saveAs(const QString &filepath)
{
    m_Workbook->dynamicCall("SaveAs(const QString&)", filepath);
}

void QXlsWorkbook::activate()
{
    m_Workbook->dynamicCall("Activate()");
}

QString QXlsWorkbook::documentation() const
{
    return m_Workbook->generateDocumentation();
}

/**
=========================================================================
                        END OF QXLSWORKBOOK REALIZATION CODE
=========================================================================
**/

/**
=========================================================================
                        QXLSWORKSHEET REALIZATION CODE
=========================================================================
**/

QXlsWorksheet::QXlsWorksheet(QObject *parent, QAxObject* RefObject) : QObject(parent)
{
    m_Worksheet = RefObject;
    m_name = RefObject->property("Name").toString();
}

QXlsWorksheet::~QXlsWorksheet()
{
    delete m_Worksheet;
}

QString QXlsWorksheet::name()
{
    return m_name;
}

QString QXlsWorksheet::documentation() const
{
    return m_Worksheet->generateDocumentation();
}

QAxObject *QXlsWorksheet::range(const QXlsRange &_range)
{
    QAxObject* cell1 = cell(_range.ul);
    QAxObject* cell2 = cell(_range.lr);
    QAxObject* range = m_Worksheet->querySubObject("Range(const QVariant&,const QVariant&)", cell1->asVariant(), cell2->asVariant());
    delete cell1;
    delete cell2;
    return range;
}

QAxObject *QXlsWorksheet::range(const QXlsRow &_row)
{
    QAxObject *Range = m_Worksheet->querySubObject( "Range(const QVariant&)",
                                                    QVariant(QString(_row.name + ":" + _row.name)));
    return Range;
}

QAxObject *QXlsWorksheet::range(const QXlsColumn &_column)
{
    QAxObject *Range = m_Worksheet->querySubObject( "Range(const QVariant&)",
                                                    QVariant(QString(_column.name + ":" + _column.name)));
    return Range;
}

QAxObject *QXlsWorksheet::range(const QXlsRow &_row1, const QXlsRow &_row2)
{
    QAxObject *Range = m_Worksheet->querySubObject( "Range(const QVariant&)",
                                                    QVariant(QString(_row1.name + ":" + _row2.name)));
    return Range;
}

QAxObject *QXlsWorksheet::range(const QXlsColumn &_column1, const QXlsColumn &_column2)
{
    QAxObject *Range = m_Worksheet->querySubObject( "Range(const QVariant&)",
                                                    QVariant(QString(_column1.name + ":" + _column2.name)));
    return Range;
}

QAxObject *QXlsWorksheet::cell(const QXlsCell &_cell)
{
    QAxObject* cell = m_Worksheet->querySubObject("Cells(QVariant&,QVariant&)", _cell.row, _cell.col);
    return cell;
}

void QXlsWorksheet::write(const QXlsCell &_cell, const QString &data)
{
    QAxObject* Cell = cell(_cell);
    Cell->setProperty("Value", QVariant(data));
    delete Cell;
}

void QXlsWorksheet::write(int row, int col, const QString &data)
{
    QXlsCell _cell(row, col);
    QAxObject* Cell = cell(_cell);
    Cell->setProperty("Value", QVariant(data));
    delete Cell;
}

void QXlsWorksheet::write(const QXlsRange &_range, QList<QStringList> &data)
{
    QList<QVariant> VariantData;
    foreach (QStringList Entry, data) {
        VariantData << QVariant(Entry);
    }
    QAxObject* Range = range(_range);
    Range->setProperty("Value", QVariant(VariantData));
    delete Range;
}

void QXlsWorksheet::align_vertical(const QXlsRange &_range, const QXlsTextAlignmentV &alignment)
{
    QAxObject* Range = range(_range);
    Range->dynamicCall("VerticalAlignment", alignment.alignment);
    delete Range;
}

void QXlsWorksheet::align_horizontal(const QXlsRange &_range, const QXlsTextAlignmentH &alignment)
{
    QAxObject* Range = range(_range);
    Range->dynamicCall("HorizontalAlignment", alignment.alignment);
    delete Range;
}

void QXlsWorksheet::setBorder(const QXlsRange &_range, const QXlsBorder &border, const QXlsBorderStyle &style)
{
    QAxObject* Range = range(_range);

    QString borderString;
    if(border.border == QXlsBorder::Top) borderString = "Borders(xlEdgeTop)";
    if(border.border == QXlsBorder::Bottom) borderString = "Borders(xlEdgeBottom)";
    if(border.border == QXlsBorder::Left) borderString = "Borders(xlEdgeLeft)";
    if(border.border == QXlsBorder::Right) borderString = "Borders(xlEdgeRight)";

    QAxObject *Border = Range->querySubObject(borderString.toLatin1());

    Border->setProperty("LineStyle", style.style);
    Border->setProperty("Weight", style.weight);

    delete Border;
    delete Range;
}

void QXlsWorksheet::setBorder(const QXlsCell &_cell, const QXlsBorder &border, const QXlsBorderStyle &style)
{
    QAxObject* Cell = cell(_cell);

    QString borderString;

    if(border.border == QXlsBorder::Top) borderString = "Borders(xlEdgeTop)";
    if(border.border == QXlsBorder::Bottom) borderString = "Borders(xlEdgeBottom)";
    if(border.border == QXlsBorder::Left) borderString = "Borders(xlEdgeLeft)";
    if(border.border == QXlsBorder::Right) borderString = "Borders(xlEdgeRight)";


    QAxObject *Border = Cell->querySubObject(borderString.toLatin1());


    Border->setProperty("LineStyle", style.style);
    Border->setProperty("Weight", style.weight);

    delete Border;
    delete Cell;
}

void QXlsWorksheet::select(const QXlsCell &_cell)
{
    QAxObject *Cell = cell(_cell);
    Cell->dynamicCall("Select()");
    delete Cell;
}

void QXlsWorksheet::select(const QXlsRow &_row)
{
    QAxObject *Range = range(_row);
    Range->dynamicCall("Select()");
    delete Range;
}

void QXlsWorksheet::select(const QXlsRow &_row1, const QXlsRow &_row2)
{
    QAxObject *Range = range(_row1, _row2);
    Range->dynamicCall("Select()");
    delete Range;
}

void QXlsWorksheet::select(const QXlsColumn &_column)
{
    QAxObject *Range = range(_column);
    Range->dynamicCall("Select()");
    delete Range;
}

void QXlsWorksheet::select(const QXlsColumn &_column1, const QXlsColumn &_column2)
{
    QAxObject *Range = range(_column1, _column2);
    Range->dynamicCall("Select()");
    delete Range;
}

void QXlsWorksheet::select(const QXlsRange &_range)
{
    QAxObject *Range = range(_range);
    Range->dynamicCall("Select()");
    delete Range;
}

void QXlsWorksheet::copy(const QXlsRange &_range)
{
    QAxObject *Range = range(_range);
    Range->dynamicCall("Copy()");
    delete Range;
}

void QXlsWorksheet::copy(const QXlsCell &_cell)
{
    QAxObject *Cell = cell(_cell);
    Cell->dynamicCall("Copy()");
    delete Cell;
}

void QXlsWorksheet::copy(const QXlsRow &_row)
{
    QAxObject *Range = range(_row);
    Range->dynamicCall("Copy()");
    delete Range;
}

void QXlsWorksheet::copy(const QXlsRow &_row1, const QXlsRow &_row2)
{
    QAxObject *Range = range(_row1, _row2);
    Range->dynamicCall("Copy()");
    delete Range;
}

void QXlsWorksheet::copy(const QXlsColumn &_column)
{
    QAxObject *Range = range(_column);
    Range->dynamicCall("Copy()");
    delete Range;
}

void QXlsWorksheet::copy(const QXlsColumn &_column1, const QXlsColumn &_column2)
{
    QAxObject *Range = range(_column1, _column2);
    Range->dynamicCall("Copy()");
    delete Range;
}

void QXlsWorksheet::paste(const QXlsRange &_range)
{
    QAxObject *Range = range(_range);
    Range->dynamicCall("Select()");
    m_Worksheet->dynamicCall("Paste()");
    delete Range;
}

void QXlsWorksheet::paste(const QXlsCell &_cell)
{
    QAxObject *Cell = cell(_cell);
    Cell->dynamicCall("Select()");
    m_Worksheet->dynamicCall("Paste()");
    delete Cell;
}

void QXlsWorksheet::merge(const QXlsRange &_range)
{
    QAxObject *Range = range(_range);
    Range->dynamicCall("MergeCells", true);
    delete Range;
}

void QXlsWorksheet::setBackgroundColor(const QXlsCell &_cell, const QColor &color)
{
    QAxObject* Cell = cell(_cell);
    QAxObject* Interior = Cell->querySubObject("Interior");
    Interior->setProperty("Color", QColor(color));
    delete Interior;
    delete Cell;
}

void QXlsWorksheet::setBackgroundColor(const QXlsRange &_range, const QColor &color)
{
    QAxObject* Range = range(_range);
    QAxObject* Interior = Range->querySubObject("Interior");
    Interior->setProperty("Color", QColor(color));
    delete Interior;
    delete Range;
}

void QXlsWorksheet::setFont(const QXlsCell &_cell, const QXlsFont &_font)
{
    QAxObject* Cell = cell(_cell);
    QAxObject *Font = Cell->querySubObject("Font");
    Font->setProperty("Name", QFont(_font.font));
    Font->setProperty("Size", QVariant(_font.size));
    Font->setProperty("Color", QColor(_font.color));
    delete Font;
    delete Cell;
}

void QXlsWorksheet::setFont(const QXlsRange &_range, const QXlsFont &_font)
{
    QAxObject* Range = range(_range);
    QAxObject *Font = Range->querySubObject("Font");
    Font->setProperty("Name", QFont(_font.font));
    Font->setProperty("Size", QVariant(_font.size));
    Font->setProperty("Color", QColor(_font.color));
    delete Font;
    delete Range;
}

void QXlsWorksheet::setWidth(const QXlsColumn &_column, int width)
{
    QAxObject *Range = range(_column);
    QAxObject *Width = Range->querySubObject("Columns");
    Width->setProperty("ColumnWidth",width);
    delete Width;
    delete Range;
}

void QXlsWorksheet::setWidth(const QXlsRange &_range, int width)
{
    QAxObject *Range = range(_range);
    QAxObject *Width = Range->querySubObject("Columns");
    Width->setProperty("ColumnWidth",width);
    delete Width;
    delete Range;
}

void QXlsWorksheet::setHeight(const QXlsRow &_row, int height)
{
    QAxObject *Range = range(_row);
    QAxObject *Height = Range->querySubObject("Rows");
    Height->setProperty("RowHeight",height);
    delete Height;
    delete Range;
}

void QXlsWorksheet::setHeight(const QXlsRange &_range, int height)
{
    QAxObject *Range = range(_range);
    QAxObject *Height = Range->querySubObject("Rows");
    Height->setProperty("RowHeight",height);
    delete Height;
    delete Range;
}

QString QXlsWorksheet::read(const QXlsCell &_cell)
{
    QAxObject *Cell = cell(_cell);
    QVariant Value = Cell->property("Value");
    delete Cell;
    return Value.toString();
}

QVariant QXlsWorksheet::property(const QXlsCell &_cell, const QString &_property)
{
    QAxObject *Cell = cell(_cell);
    QVariant property = Cell->property(_property.toLatin1());
    delete Cell;
    return property;
}

/**
=========================================================================
                        END OF QXLSWORKSHEET REALIZATION CODE
=========================================================================
**/

QXlsRow::QXlsRow(QString _name) : name(_name)
{
}

QXlsColumn::QXlsColumn(QString _name) : name(_name)
{
}

QXlsRange::QXlsRange(QXlsCell _ul, QXlsCell _lr) : ul(_ul), lr(_lr)
{
}

QXlsCell::QXlsCell(int _row, int _col) : row(_row), col(_col)
{
}

QXlsBorderStyle::QXlsBorderStyle(QXlsBorderStyle::BorderStyle _style, BorderWeight _weight) :
    style(_style), weight(_weight)
{
}

QXlsBorder::QXlsBorder(QXlsBorder::Border _border) :
    border(_border)
{
}

QXlsTextAlignmentV::QXlsTextAlignmentV(QXlsTextAlignmentV::Alignment _alignment) : alignment(_alignment)
{
}

QXlsTextAlignmentH::QXlsTextAlignmentH(QXlsTextAlignmentH::Alignment _alignment) : alignment(_alignment)
{
}

QXlsFont::QXlsFont(QColor _color, QFont _font, int _size) : color(_color), font(_font), size(_size)
{
}
