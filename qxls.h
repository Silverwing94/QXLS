/*
QXLS - a small wrapper library around ActiveQt QAxObjects,
representing MS Excel app and its subobjects,
used to interact and control MS Excel from Qt
Copyright (C) 2018 Konstantin "Silverwing" Yurkov kotyurkov@yandex.ru

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

#ifndef QXLS_H
#define QXLS_H

#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>


#include <QStringList>
#include <QColor>
#include <QFont>

class QXlsWorkbook;
class QXlsWorksheet;
class QXlsRow;
class QXlsColumn;
class QXlsCell;
class QXlsRange;
class QXlsBorder;
class QXlsBorderStyle;
class QXlsTextAlignmentH;
class QXlsTextAlignmentV;
class QXlsFont;

class QXlsApplication : public QObject
{
    Q_OBJECT

public:

    explicit QXlsApplication(QObject *parent = 0);
    ~QXlsApplication();
    QXlsWorkbook* workbook(int) const;
    QXlsWorkbook* newWorkbook();
    QXlsWorkbook* activeWorkbook() const;
    QString documentation() const;
    int workbooksCount() const;
    QXlsWorkbook* open(const QString &filepath);
    void alerts(bool state);
    void exit();

private:

    QAxObject* m_ExcelApplication;
    QAxObject* m_WorkbooksObject;
    QList<QXlsWorkbook*> m_Workbooks;
    QXlsWorkbook* m_ActiveWorkbook;

public slots:
    void show();
    void hide();
};

class QXlsWorkbook : public QObject
{
    Q_OBJECT

public:

    explicit QXlsWorkbook(QObject *parent = 0, QAxObject* RefObject = 0);
    bool isActive();
    QString name() const;
    int worksheetsCount() const;
    QXlsWorksheet* newWorksheet();
    QXlsWorksheet* sheet(int) const;
    QString documentation() const;
    ~QXlsWorkbook();

private:

    QAxObject* m_Workbook;
    QAxObject* m_WorksheetsObject;
    int m_count;
    QList<QXlsWorksheet*> m_Worksheets;
    bool m_active;
    int m_index;

public slots:

    void close();
    void save();
    void saveAs(const QString &filepath);
    void activate();

};

class QXlsWorksheet : public QObject
{
    Q_OBJECT

public:

    explicit QXlsWorksheet(QObject *parent = 0,  QAxObject* RefObject = 0);
    ~QXlsWorksheet();
     QString name();
     QString documentation() const;

     void write(const QXlsCell &_cell, const QString &data);
     void write(int row, int col, const QString &data);
     void write(const QXlsRange &_range, QList<QStringList> &data);

     void align_vertical(const QXlsRange &_range, const QXlsTextAlignmentV &alignment);
     void align_horizontal(const QXlsRange &_range, const QXlsTextAlignmentH &alignment);

     void setBorder(const QXlsRange &_range, const QXlsBorder &border, const QXlsBorderStyle &style);
     void setBorder(const QXlsCell &_cell, const QXlsBorder &border, const QXlsBorderStyle &style);

     void select(const QXlsCell &_cell);
     void select(const QXlsRow &_row);
     void select(const QXlsRow &_row1, const QXlsRow &_row2);
     void select(const QXlsColumn &_column);
     void select(const QXlsColumn &_column1, const QXlsColumn &_column2);
     void select(const QXlsRange &_range);

     void copy(const QXlsCell &_cell);
     void copy(const QXlsRow &_row);
     void copy(const QXlsRow &_row1, const QXlsRow &_row2);
     void copy(const QXlsColumn &_column);
     void copy(const QXlsColumn &_column1, const QXlsColumn &_column2);
     void copy(const QXlsRange &_range);

     void paste(const QXlsRange &_range);
     void paste(const QXlsCell &_cell);

     void merge(const QXlsRange &_range);

     void setBackgroundColor(const QXlsCell &_cell, const QColor &color);
     void setBackgroundColor(const QXlsRange &_range, const QColor &color);

     void setFont(const QXlsCell &_cell, const QXlsFont &_font);
     void setFont(const QXlsRange &_range, const QXlsFont &_font);

     void setWidth(const QXlsColumn &_column, int width);
     void setWidth(const QXlsRange &_range, int width);
     void setHeight(const QXlsRow &_row, int height);
     void setHeight(const QXlsRange &_range, int height);

     QString read(const QXlsCell &_cell);
     QVariant property(const QXlsCell &_cell, const QString &_property);

private:

    QAxObject* m_Worksheet;
    QString m_name;
    QAxObject* range(const QXlsRange &_range);
    QAxObject* range(const QXlsRow &_row);
    QAxObject* range(const QXlsColumn &_column);
    QAxObject* range(const QXlsRow &_row1, const QXlsRow &_row2);
    QAxObject* range(const QXlsColumn &_column1, const QXlsColumn &_column2);
    QAxObject* cell(const QXlsCell &_cell);
};

class QXlsRow
{
public:
    explicit QXlsRow(QString _name);
    const QString name;
};

class QXlsColumn
{
public:
    explicit QXlsColumn(QString _name);
    const QString name;
};

class QXlsCell
{
public:
    QXlsCell(int _row, int _col);
    const int row;
    const int col;
};

class QXlsRange
{
public:
    QXlsRange(QXlsCell _ul, QXlsCell _lr);
    const QXlsCell ul;
    const QXlsCell lr;
};

class QXlsTextAlignmentH {
public:
    enum Alignment {Left = -4131, Center = -4108, Right = -4152};
    QXlsTextAlignmentH(Alignment _alignment = QXlsTextAlignmentH::Center);
    const Alignment alignment;
};

class QXlsTextAlignmentV {
public:
    enum Alignment {Top = -4160, Center = -4108, Bottom = -4107};
    QXlsTextAlignmentV(Alignment _alignment = QXlsTextAlignmentV::Center);
    const Alignment alignment;
};

class QXlsBorderStyle {
public:
    enum BorderStyle {Solid = 1, Dotted = -4118, Dashed = -4115, DashDot = 4};
    enum BorderWeight {Hairline = 1, Thin = 2, Medium = -4138, Thick = 4};
    QXlsBorderStyle(BorderStyle _style = QXlsBorderStyle::Solid, BorderWeight _weight = QXlsBorderStyle::Thin);
    const BorderStyle style;
    const BorderWeight weight;
};

class QXlsBorder {
public:
    enum Border {Top = 1, Bottom = 2, Right = 3, Left = 4, All = 5};
    QXlsBorder(Border _border = QXlsBorder::All);
    const Border border;
};

class QXlsFont {
public:
    QXlsFont(QColor _color = QColor(0,0,0), QFont _font = QFont("Arial"), int _size = 10);
    QColor color;
    QFont font;
    int size;
};

#endif // QXLS_H
