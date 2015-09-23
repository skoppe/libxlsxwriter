/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page workbook_page The Workbook object
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * See @ref workbook.h for full details of the functionality.
 *
 * @file workbook.h
 *
 * @brief Functions related to creating an Excel xlsx workbook.
 *
 * The Workbook is the main object exposed by the libxlsxwriter library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * @code
 *     #include "xlsxwriter.h"
 *
 *     int main() {
 *
 *         lxw_workbook  *workbook  = new_workbook("filename.xlsx");
 *         lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
 *
 *         worksheet_write_string(worksheet, 0, 0, "Hello Excel", NULL);
 *
 *         return workbook_close(workbook);
 *     }
 * @endcode
 *
 * @image html workbook01.png
 *
 */
#ifndef __LXW_WORKBOOK_H__
#define __LXW_WORKBOOK_H__

#include <stdint.h>
#include <stdio.h>

#include "worksheet.h"
#include "shared_strings.h"
#include "hash_table.h"
#include "common.h"

/* Define the queue.h structs for the workbook lists. */
struct lxw_worksheets {
    struct lxw_worksheet *stqh_first;/* first element */
    struct lxw_worksheet **stqh_last;/* addr of last next element */
};
struct lxw_defined_names {
    struct lxw_defined_name *tqh_first; /* first element */
    struct lxw_defined_name **tqh_last; /* addr of last next element */
};

#define LXW_DEFINED_NAME_LENGTH 128

/* Struct to represent a defined name. */
typedef struct lxw_defined_name {
    int16_t index;
    uint8_t hidden;
    char name[LXW_DEFINED_NAME_LENGTH];
    char app_name[LXW_DEFINED_NAME_LENGTH];
    char formula[LXW_DEFINED_NAME_LENGTH];
    char normalised_name[LXW_DEFINED_NAME_LENGTH];
    char normalised_sheetname[LXW_DEFINED_NAME_LENGTH];

    /* List pointers for queue.h. */
    struct {
        struct lxw_defined_name *tqe_next;  /* next element */
        struct lxw_defined_name **tqe_prev; /* address of previous next element */
    } list_pointers;
} lxw_defined_name;

/**
 * @brief Errors conditions encountered when closing the Workbook and writing
 * the Excel file to disk.
 */
enum lxw_close_error {
    /** No error */
    LXW_CLOSE_ERROR_NONE,
    /** Error encountered when creating file zip container */
    LXW_CLOSE_ERROR_ZIP
        /* TODO. Need to add/document more. */
};

/**
 * @brief Workbook options.
 *
 * Optional parameters when creating a new Workbook object via
 * new_workbook_opt().
 *
 * Currently only the `constant_memory` property is supported:
 *
 * * `constant_memory`
 */
typedef struct lxw_workbook_options {
    /** Optimise the workbook to use constant memory for worksheets */
    uint8_t constant_memory;
} lxw_workbook_options;

/**
 * @brief Struct to represent an Excel workbook.
 *
 * The members of the lxw_workbook struct aren't modified directly. Instead
 * the workbook properties are set by calling the functions shown in
 * workbook.h.
 */
typedef struct lxw_workbook {

    FILE *file;
    struct lxw_worksheets *worksheets;
    struct lxw_formats *formats;
    struct lxw_defined_names *defined_names;
    lxw_sst *sst;
    lxw_doc_properties *properties;
    const char *filename;
    lxw_workbook_options options;

    uint16_t num_sheets;
    uint16_t first_sheet;
    uint32_t active_sheet;
    uint16_t num_xf_formats;
    uint16_t num_format_count;

    uint16_t font_count;
    uint16_t border_count;
    uint16_t fill_count;
    uint8_t optimize;

    lxw_hash_table *used_xf_formats;

} lxw_workbook;


/* *INDENT-OFF* */
#ifdef __cplusplus
extern "C" {
#endif
/* *INDENT-ON* */

/**
 * @brief Create a new workbook object.
 *
 * @param filename The name of the new Excel file to create.
 *
 * @return A lxw_workbook instance.
 *
 * The `%new_workbook()` constructor is used to create a new Excel workbook
 * with a given filename:
 *
 * @code
 *     lxw_workbook *workbook  = new_workbook("filename.xlsx");
 * @endcode
 *
 * When specifying a filename it is recommended that you use an `.xlsx`
 * extension or Excel will generate a warning when opening the file.
 *
 */
lxw_workbook *new_workbook(const char *filename);

/**
 * @brief Create a new workbook object, and set the workbook options.
 *
 * @param filename The name of the new Excel file to create.
 * @param options  Workbook options.
 *
 * @return A lxw_workbook instance.
 *
 * This method is the same as the `new_workbook()` constructor but allows
 * additional options to be set.
 *
 * @code
 *    lxw_workbook_options options = {.constant_memory = 1};
 *
 *    lxw_workbook  *workbook  = new_workbook_opt("filename.xlsx", &options);
 * @endcode
 *
 * Note, in this mode a row of data is written and then discarded when a cell
 * in a new row is added via one of the worksheet `worksheet_write_*()`
 * methods.  Therefore, once this mode is active, data should be written in
 * sequential row order.
 *
 * See @ref working_with_memory for more details.
 *
 */
lxw_workbook *new_workbook_opt(const char *filename,
                               lxw_workbook_options *options);

/**
 * @brief Add a new worksheet to a workbook:
 *
 * @param workbook  Pointer to a lxw_workbook instance.
 * @param sheetname Optional worksheet name, defaults to Sheet1, etc.
 *
 * @return A lxw_worksheet instance.
 *
 * The `%workbook_add_worksheet()` method adds a new worksheet to a workbook:
 *
 * At least one worksheet should be added to a new workbook: The @ref
 * worksheet.h "Worksheet" object is used to write data and configure a
 * worksheet in the workbook.
 *
 * The `sheetname` parameter is optional. If it is `NULL` the default
 * Excel convention will be followed, i.e. Sheet1, Sheet2, etc.:
 *
 * @code
 *     worksheet = workbook_add_worksheet(workbook, NULL  );     // Sheet1
 *     worksheet = workbook_add_worksheet(workbook, "Foglio2");  // Foglio2
 *     worksheet = workbook_add_worksheet(workbook, "Data");     // Data
 *     worksheet = workbook_add_worksheet(workbook, NULL  );     // Sheet4
 *
 * @endcode
 *
 * @image html workbook02.png
 *
 * The worksheet name must be a valid Excel worksheet name, i.e. it must be
 * less than 32 character and it cannot contain any of the characters:
 *
 *     / \ [ ] : * ?
 *
 * In addition, you cannot use the same, case insensitive, `sheetname` for more
 * than one worksheet.
 *
 */
lxw_worksheet *workbook_add_worksheet(lxw_workbook *workbook,
                                      const char *sheetname);

/**
 * @brief Create a new @ref format.h "Format" object to formats cells in
 *        worksheets.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 *
 * @return A lxw_format instance.
 *
 * The `workbook_add_format()` function can be used to create new @ref
 * format.h "Format" objects which are used to apply formatting to a cell.
 *
 * @code
 *    // Create the Format.
 *    lxw_format *format = workbook_add_format(workbook);
 *
 *    // Set some of the format properties.
 *    format_set_bold(format);
 *    format_set_font_color(format, LXW_COLOR_RED);
 *
 *    // Use the format to change the text format in a cell.
 *    worksheet_write_string(worksheet, 0, 0, "Hello", format);
 * @endcode
 *
 * See @ref format.h "the Format object" and @ref working_with_formats
 * sections for more details about Format properties and how to set them.
 *
 */
lxw_format *workbook_add_format(lxw_workbook *workbook);

/**
 * @brief Close the Workbook object and write the XLSX file.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 *
 * @return A #lxw_close_error.
 *
 * The `%workbook_close()` function closes a Workbook object, writes the Excel
 * file to disk, frees any memory allocated internally to the Workbook and
 * frees the object itself.
 *
 * @code
 *     workbook_close(workbook);
 * @endcode
 *
 * The `%workbook_close()` function returns any #lxw_close_error error codes
 * encountered when creating the Excel file. The error code can be returned
 * from the program main or the calling function:
 *
 * @code
 *     return workbook_close(workbook);
 * @endcode
 *
 */
uint8_t workbook_close(lxw_workbook *workbook);

/**
 * @brief Create a defined name in the workbook to use as a variable.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The defined name.
 * @param formula  The cell or range that the defined name refers to.
 *
 * @return 0 for success, non-zero on error.
 *
 * This method is used to defined a name that can be used to represent a
 * value, a single cell or a range of cells in a workbook: These defined names
 * can then be used in formulas:
 *
 * @code
 *     workbook_define_name(workbook, "Exchange_rate", "=0.96");
 *     worksheet_write_formula(worksheet, 2, 1, "=Exchange_rate", NULL);
 *
 * @endcode
 *
 * @image html defined_name.png
 *
 * As in Excel a name defined like this is "global" to the workbook and can be
 * referred to from any worksheet:
 *
 * @code
 *     // Global workbook name.
 *     workbook_define_name(workbook, "Sales", "=Sheet1!$G$1:$H$10");
 * @endcode
 *
 * It is also possible to define a local/worksheet name by prefixing it with
 * the sheet name using the syntax `'sheetname!definedname'`:
 *
 * @code
 *     // Local worksheet name.
 *     workbook_define_name(workbook, "Sheet2!Sales", "=Sheet2!$G$1:$G$10");
 * @endcode
 *
 * If the sheet name contains spaces or special characters you must follow the
 * Excel convention and enclose it in single quotes:
 *
 * @code
 *     workbook_define_name(workbook, "'New Data'!Sales", "=Sheet2!$G$1:$G$10");
 * @endcode
 *
 * The rules for names in Excel are explained in the
 * [Microsoft Office
documentation](http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx).
 *
 */
uint8_t workbook_define_name(lxw_workbook *workbook, const char *name,
                             const char *formula);

void _free_workbook(lxw_workbook *workbook);
void _workbook_assemble_xml_file(lxw_workbook *workbook);
void _set_default_xf_indices(lxw_workbook *workbook);

/* Declarations required for unit testing. */
#ifdef TESTING

STATIC void _workbook_xml_declaration(lxw_workbook *self);
STATIC void _workbook_xml_declaration(lxw_workbook *self);
STATIC void _write_workbook(lxw_workbook *self);
STATIC void _write_file_version(lxw_workbook *self);
STATIC void _write_workbook_pr(lxw_workbook *self);
STATIC void _write_book_views(lxw_workbook *self);
STATIC void _write_workbook_view(lxw_workbook *self);
STATIC void _write_sheet(lxw_workbook *self,
                         const char *name, uint32_t sheet_id, uint8_t hidden);
STATIC void _write_sheets(lxw_workbook *self);
STATIC void _write_calc_pr(lxw_workbook *self);

STATIC void _write_defined_name(lxw_workbook *self,
                                lxw_defined_name *define_name);
STATIC void _write_defined_names(lxw_workbook *self);

STATIC uint8_t _store_defined_name(lxw_workbook *self, const char *name,
                                   const char *app_name, const char *formula,
                                   int16_t index, uint8_t hidden);

#endif /* TESTING */

/* *INDENT-OFF* */
#ifdef __cplusplus
}
#endif
/* *INDENT-ON* */

#endif /* __LXW_WORKBOOK_H__ */
