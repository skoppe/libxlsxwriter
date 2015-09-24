# Fork

This fork is used to generate bindings to the [D](https://www.dlang.org) programming language.

D can call C-code so only the `include/xlsxwriter.h` (and includes) have been translated. I used the nice htod tool from Digital Mars.

Nothing much has changed except that I needed to removed references to `include/xlsxwriter/third_party/queue.h` since it was using  C macro's and they don't translate well to D.

This fork breaks the normal build, because the headers files have been modified without modifications to the source files. For that reason I decided to remove the travis.yml file.

I only use this to generate the xlsxwriter.d file and use that in other projects (more on that later...)

# libxlsxwriter


A C library for creating Excel XLSX files.


![demo image](http://libxlsxwriter.github.io/demo.png)


## The libxlsxwriter library

Libxlsxwriter is a C library that can be used to write text, numbers, formulas and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file.

It supports features such as:

- 100% compatible Excel XLSX files
- Full Excel formatting
- Merged cells
- Autofilters
- Defined names
- Memory optimisation mode for writing large files
- Source code available on [GitHub](https://github.com/jmcnamara/libxlsxwriter)
- FreeBSD license
- ANSI C
- Works with GCC 4.4, 4.6, 4.7, 4.8, 4.9, Clang, ICC and TCC.
- Works on Linux, FreeBSD, OS X and iOS.
- The only dependency is on `zlib`


Here is an example that was used to create the spreadsheet shown above:


```C
#include "xlsxwriter.h"

int main() {

    /* Create a new workbook and add a worksheet. */
    lxw_workbook  *workbook  = new_workbook("demo.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    /* Add a format. */
    lxw_format *format = workbook_add_format(workbook);

    /* Set the bold property for the format */
    format_set_bold(format);

    /* Widen the first column to make the text clearer. */
    worksheet_set_column(worksheet, 0, 0, 20, NULL, NULL);

    /* Write some simple text. */
    worksheet_write_string(worksheet, 0, 0, "Hello", NULL);

    /* Text with formatting. */
    worksheet_write_string(worksheet, 1, 0, "World", format);

    /* Writer some numbers. */
    worksheet_write_number(worksheet, 2, 0, 123,     NULL);
    worksheet_write_number(worksheet, 3, 0, 123.456, NULL);

    workbook_close(workbook);

    return 0;
}

```



See the [full documentation](http://libxlsxwriter.github.io) for the getting started guide, a tutorial, the main API documentation and examples.
