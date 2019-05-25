#include <QCoreApplication>
#include <QVariant>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

void write()
{
    // [1]  Writing excel file(*.xlsx)
    QXlsx::Document xlsxW;
    xlsxW.write("A1", "Hello Qt!"); // write "Hello Qt!" to cell(A,1). it's shared string.
    xlsxW.saveAs("Test.xlsx"); // save the document as 'Test.xlsx'
}

void read()
{
    // [2] Reading excel file(*.xlsx)
    Document xlsxR("Test.xlsx");
    if (xlsxR.load()) // load excel file
    {
        int row = 1; int col = 1;
        Cell* cell = xlsxR.cellAt(row, col); // get cell pointer.
        if ( cell != nullptr )
        {
            QVariant var = cell->readValue(); // read cell value (number(double), QDateTime, QString ...)
            qDebug() << var; // display value. it is 'Hello Qt!'.
        }
    }
}

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    write();
    read();

    return 0;
}
