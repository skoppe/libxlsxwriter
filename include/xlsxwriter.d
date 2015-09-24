/* Converted to D from xlsxwriter.h by htod */
module dexcel.c.xlsxwriter;
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @file xlsxwriter.h
 *
 * xlsxwriter - A library for creating Excel XLSX files.
 *
 */
//C     #ifndef __LXW_XLSXWRITER_H__
//C     #define __LXW_XLSXWRITER_H__

//C     #include "xlsxwriter/workbook.h"
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
//C     #ifndef __LXW_WORKBOOK_H__
//C     #define __LXW_WORKBOOK_H__

//C     #include <stdint.h>
import std.stdint;
//C     #include <stdio.h>
import std.c.stdio;

//C     #include "worksheet.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page worksheet_page The Worksheet object
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * See @ref worksheet.h for full details of the functionality.
 *
 * @file worksheet.h
 *
 * @brief Functions related to adding data and formatting to a worksheet.
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * A Worksheet object isn't created directly. Instead a worksheet is
 * created by calling the workbook_add_worksheet() function from a
 * Workbook object:
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
 */
//C     #ifndef __LXW_WORKSHEET_H__
//C     #define __LXW_WORKSHEET_H__

//C     #include <stdio.h>
//C     #include <stdlib.h>
//import std.c.stdlib;
//C     #include <stdint.h>
//C     #include <string.h>
//import std.c.string;

//C     #include "shared_strings.h"
/*
 * libxlsxwriter
 * 
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * shared_strings - A libxlsxwriter library for creating Excel XLSX
 *                  sst files.
 *
 */
//C     #ifndef __LXW_SST_H__
//C     #define __LXW_SST_H__

//C     #include <string.h>
//C     #include <stdint.h>

//C     #include "common.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * common - Common functions and defines for the libxlsxwriter library.
 *
 */
//C     #ifndef __LXW_COMMON_H__
//C     #define __LXW_COMMON_H__

//C     #include <time.h>
import std.c.time;

//C     #ifndef TESTING
//C     #define STATIC static
//C     #else
//alias static STATIC;
//C     #define STATIC
//C     #endif

//C     #define LXW_SHEETNAME_MAX  32
//C     #define LXW_SHEETNAME_LEN  65
const LXW_SHEETNAME_MAX = 32;

const LXW_SHEETNAME_LEN = 65;
//C     enum lxw_boolean {
//C         LXW_FALSE,
//C         LXW_TRUE
//C     };
enum lxw_boolean
{
    LXW_FALSE,
    LXW_TRUE,
}

//C     #define LXW_IGNORE 1

const LXW_IGNORE = 1;
//C     #define ERROR(message)                              fprintf(stderr, "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

//C     #define MEM_ERROR()                                 ERROR("Memory allocation failed.")

//C     #define GOTO_LABEL_ON_MEM_ERROR(pointer, label)     if (!pointer) {                                     MEM_ERROR();                                    goto label;                                 }

//C     #define RETURN_ON_MEM_ERROR(pointer, error)         if (!pointer) {                                     MEM_ERROR();                                    return error;                               }

//C     #define LXW_WARN(message)                           fprintf(stderr, "[WARN]: " message "\n")

/* Define the queue.h structs for the formats list. */
//C     struct lxw_formats {
//C         struct lxw_format *stqh_first;/* first element */
//C         struct lxw_format **stqh_last;/* addr of last next element */
//C     };
struct lxw_formats
{
    lxw_format *stqh_first;
    lxw_format **stqh_last;
}

/* Define the queue.h structs for the generic data structs. */
//C     struct lxw_tuples {
//C         struct lxw_tuple *stqh_first;/* first element */
//C         struct lxw_tuple **stqh_last;/* addr of last next element */
//C     };
struct lxw_tuples
{
    lxw_tuple *stqh_first;
    lxw_tuple **stqh_last;
}

//C     typedef struct lxw_tuple {
//C         char *key;
//C         char *value;

//C         struct {
//C             struct lxw_tuple *stqe_next; /* next element */
//C         } list_pointers;
struct _N1
{
    lxw_tuple *stqe_next;
}
//C     } lxw_tuple;
struct lxw_tuple
{
    char *key;
    char *value;
    _N1 list_pointers;
}
extern (C):

//C     typedef struct lxw_doc_properties {
//C         char *title;
//C         char *subject;
//C         char *author;
//C         char *manager;
//C         char *company;
//C         char *category;
//C         char *keywords;
//C         char *comments;
//C         char *status;
//C         time_t created;
//C     } lxw_doc_properties;
struct lxw_doc_properties
{
    char *title;
    char *subject;
    char *author;
    char *manager;
    char *company;
    char *category;
    char *keywords;
    char *comments;
    char *status;
    time_t created;
}


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_COMMON_H__ */

//C     #define NUM_SST_BUCKETS 8
/* STAILQ_HEAD() declaration. */
const NUM_SST_BUCKETS = 8;
//C     struct sst_order_list {
//C         struct sst_element *stqh_first;
//C         struct sst_element **stqh_last;
//C     };
struct sst_order_list
{
    sst_element *stqh_first;
    sst_element **stqh_last;
}

/* SLIST_HEAD() declaration. */
//C     struct sst_bucket_list {
//C         struct sst_element *slh_first;
//C     };
struct sst_bucket_list
{
    sst_element *slh_first;
}

/*
 * Elements of the SST table. They contain pointers to allow them to
 * be stored in lists in the the hash table buckets and also pointers to
 * track the insertion order in a separate list.
 */
//C     struct sst_element {
//C         size_t index;
//C         char *string;

//C         struct {
//C             struct sst_element *stqe_next; /* next element */
//C         } sst_order_pointers;
struct _N2
{
    sst_element *stqe_next;
}
//C         struct {
//C             struct sst_element *sle_next;  /* next element */
//C         } sst_list_pointers;
struct _N3
{
    sst_element *sle_next;
}
//C     };
struct sst_element
{
    size_t index;
    char *string;
    _N2 sst_order_pointers;
    _N3 sst_list_pointers;
}

/*
 * Struct to represent a sst.
 */
//C     typedef struct lxw_sst {
//C         FILE *file;

//C         size_t num_buckets;
//C         size_t used_buckets;
//C         size_t string_count;
//C         size_t unique_count;

//C         struct sst_order_list *order_list;
//C         struct sst_bucket_list **buckets;

//C     } lxw_sst;
struct lxw_sst
{
    FILE *file;
    size_t num_buckets;
    size_t used_buckets;
    size_t string_count;
    size_t unique_count;
    sst_order_list *order_list;
    sst_bucket_list **buckets;
}

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     lxw_sst *_new_sst();
lxw_sst * _new_sst();
//C     void _free_sst(lxw_sst *sst);
void  _free_sst(lxw_sst *sst);
//C     int32_t _get_sst_index(lxw_sst *sst, const char *string);
int32_t  _get_sst_index(lxw_sst *sst, char *string);
//C     void _sst_assemble_xml_file(lxw_sst *self);
void  _sst_assemble_xml_file(lxw_sst *self);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     STATIC void _sst_xml_declaration(lxw_sst *self);

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_SST_H__ */
//C     #include "common.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * common - Common functions and defines for the libxlsxwriter library.
 *
 */
//C     #ifndef __LXW_COMMON_H__
//C     #define __LXW_COMMON_H__

//C     #include <time.h>

//C     #ifndef TESTING
//C     #define STATIC static
//C     #else
//C     #define STATIC
//C     #endif

//C     #define LXW_SHEETNAME_MAX  32
//C     #define LXW_SHEETNAME_LEN  65

//C     enum lxw_boolean {
//C         LXW_FALSE,
//C         LXW_TRUE
//C     };

//C     #define LXW_IGNORE 1

//C     #define ERROR(message)                              fprintf(stderr, "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

//C     #define MEM_ERROR()                                 ERROR("Memory allocation failed.")

//C     #define GOTO_LABEL_ON_MEM_ERROR(pointer, label)     if (!pointer) {                                     MEM_ERROR();                                    goto label;                                 }

//C     #define RETURN_ON_MEM_ERROR(pointer, error)         if (!pointer) {                                     MEM_ERROR();                                    return error;                               }

//C     #define LXW_WARN(message)                           fprintf(stderr, "[WARN]: " message "\n")

/* Define the queue.h structs for the formats list. */
//C     struct lxw_formats {
//C         struct lxw_format *stqh_first;/* first element */
//C         struct lxw_format **stqh_last;/* addr of last next element */
//C     };

/* Define the queue.h structs for the generic data structs. */
//C     struct lxw_tuples {
//C         struct lxw_tuple *stqh_first;/* first element */
//C         struct lxw_tuple **stqh_last;/* addr of last next element */
//C     };

//C     typedef struct lxw_tuple {
//C         char *key;
//C         char *value;

//C         struct {
//C             struct lxw_tuple *stqe_next; /* next element */
//C         } list_pointers;
//C     } lxw_tuple;

//C     typedef struct lxw_doc_properties {
//C         char *title;
//C         char *subject;
//C         char *author;
//C         char *manager;
//C         char *company;
//C         char *category;
//C         char *keywords;
//C         char *comments;
//C         char *status;
//C         time_t created;
//C     } lxw_doc_properties;


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_COMMON_H__ */
//C     #include "format.h"
/*
 * libxlsxwriter
 * 
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page format_page The Format object
 *
 * The Format object represents an the formatting properties that can be
 * applied to a cell including: fonts, colors, patterns,
 * borders, alignment and number formatting.
 *
 * See @ref format.h for full details of the functionality.
 * 
 * @file format.h
 *
 * @brief Functions and properties for adding formatting to cells in Excel.
 *
 * This section describes the functions and properties that are available for
 * formatting cells in Excel.
 *
 * The properties of a cell that can be formatted include: fonts, colors,
 * patterns, borders, alignment and number formatting.
 *
 * @image html formats_intro.png
 *
 * Formats in `libxlswriter` are accessed via the lxw_format
 * struct. Throughout this document these will be referred to simply as
 * *Formats*.
 *
 * Formats are created by calling the workbook_add_format() method as
 * follows:
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 * @endcode
 *
 * The members of the lxw_format struct aren't modified directly. Instead the
 * format properties are set by calling the function shown in this section.
 * For example:
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
 *
 * @endcode
 *
 * The full range of formatting options that can be applied using
 * `libxlswriter` are shown below.
 *
 */
//C     #ifndef __LXW_FORMAT_H__
//C     #define __LXW_FORMAT_H__

//C     #include <stdint.h>
//C     #include <string.h>
//C     #include "hash_table.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * hash_table - Hash table functions for libxlsxwriter.
 *
 */

//C     #ifndef __LXW_HASH_TABLE_H__
//C     #define __LXW_HASH_TABLE_H__

//C     #include "common.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * common - Common functions and defines for the libxlsxwriter library.
 *
 */
//C     #ifndef __LXW_COMMON_H__
//C     #define __LXW_COMMON_H__

//C     #include <time.h>

//C     #ifndef TESTING
//C     #define STATIC static
//C     #else
//C     #define STATIC
//C     #endif

//C     #define LXW_SHEETNAME_MAX  32
//C     #define LXW_SHEETNAME_LEN  65

//C     enum lxw_boolean {
//C         LXW_FALSE,
//C         LXW_TRUE
//C     };

//C     #define LXW_IGNORE 1

//C     #define ERROR(message)                              fprintf(stderr, "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

//C     #define MEM_ERROR()                                 ERROR("Memory allocation failed.")

//C     #define GOTO_LABEL_ON_MEM_ERROR(pointer, label)     if (!pointer) {                                     MEM_ERROR();                                    goto label;                                 }

//C     #define RETURN_ON_MEM_ERROR(pointer, error)         if (!pointer) {                                     MEM_ERROR();                                    return error;                               }

//C     #define LXW_WARN(message)                           fprintf(stderr, "[WARN]: " message "\n")

/* Define the queue.h structs for the formats list. */
//C     struct lxw_formats {
//C         struct lxw_format *stqh_first;/* first element */
//C         struct lxw_format **stqh_last;/* addr of last next element */
//C     };

/* Define the queue.h structs for the generic data structs. */
//C     struct lxw_tuples {
//C         struct lxw_tuple *stqh_first;/* first element */
//C         struct lxw_tuple **stqh_last;/* addr of last next element */
//C     };

//C     typedef struct lxw_tuple {
//C         char *key;
//C         char *value;

//C         struct {
//C             struct lxw_tuple *stqe_next; /* next element */
//C         } list_pointers;
//C     } lxw_tuple;

//C     typedef struct lxw_doc_properties {
//C         char *title;
//C         char *subject;
//C         char *author;
//C         char *manager;
//C         char *company;
//C         char *category;
//C         char *keywords;
//C         char *comments;
//C         char *status;
//C         time_t created;
//C     } lxw_doc_properties;


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_COMMON_H__ */

/* Macro to loop over hash table elements in insertion orfder. */
//C     #define LXW_FOREACH_ORDERED(elem, hash_table)     STAILQ_FOREACH((elem), (hash_table)->order_list, lxw_hash_order_pointers)

/* List declarations. */
//C     struct lxw_hash_order_list {
//C         struct lxw_hash_element *stqh_first;/* first element */
//C         struct lxw_hash_element **stqh_last;/* addr of last next element */
//C     };
struct lxw_hash_order_list
{
    lxw_hash_element *stqh_first;
    lxw_hash_element **stqh_last;
}
//C     struct lxw_hash_bucket_list {
//C         struct lxw_hash_element *slh_first; /* first element */
//C     };
struct lxw_hash_bucket_list
{
    lxw_hash_element *slh_first;
}

/* LXW_HASH hash table struct. */
//C     typedef struct lxw_hash_table {
//C         size_t num_buckets;
//C         size_t used_buckets;
//C         size_t unique_count;
//C         uint8_t free_key;
//C         uint8_t free_value;

//C         struct lxw_hash_order_list *order_list;
//C         struct lxw_hash_bucket_list **buckets;
//C     } lxw_hash_table;
struct lxw_hash_table
{
    size_t num_buckets;
    size_t used_buckets;
    size_t unique_count;
    uint8_t free_key;
    uint8_t free_value;
    lxw_hash_order_list *order_list;
    lxw_hash_bucket_list **buckets;
}

/*
 * LXW_HASH table element struct.
 *
 * The hash elements contain pointers to allow them to be stored in
 * lists in the the hash table buckets and also pointers to track the
 * insertion order in a separate list.
 */
//C     typedef struct lxw_hash_element {
//C         void *key;
//C         void *value;

//C         struct {
//C             struct lxw_hash_element *stqe_next; /* next element */
//C         } lxw_hash_order_pointers;
struct _N4
{
    lxw_hash_element *stqe_next;
}
//C         struct {
//C             struct lxw_hash_element *sle_next;  /* next element */
//C         } lxw_hash_list_pointers;
struct _N5
{
    lxw_hash_element *sle_next;
}
//C     } lxw_hash_element;
struct lxw_hash_element
{
    void *key;
    void *value;
    _N4 lxw_hash_order_pointers;
    _N5 lxw_hash_list_pointers;
}


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     lxw_hash_element *_hash_key_exists(lxw_hash_table *lxw_hash, void *key,
//C                                        size_t key_len);
lxw_hash_element * _hash_key_exists(lxw_hash_table *lxw_hash, void *key, size_t key_len);
//C     lxw_hash_element *_insert_hash_element(lxw_hash_table *lxw_hash, void *key,
//C                                            void *value, size_t key_len);
lxw_hash_element * _insert_hash_element(lxw_hash_table *lxw_hash, void *key, void *value, size_t key_len);
//C     lxw_hash_table *_new_lxw_hash(size_t num_buckets, uint8_t free_key,
//C                                   uint8_t free_value);
lxw_hash_table * _new_lxw_hash(size_t num_buckets, uint8_t free_key, uint8_t free_value);
//C     void _free_lxw_hash(lxw_hash_table *lxw_hash);
void  _free_lxw_hash(lxw_hash_table *lxw_hash);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_HASH_TABLE_H__ */

//C     #include "common.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * common - Common functions and defines for the libxlsxwriter library.
 *
 */
//C     #ifndef __LXW_COMMON_H__
//C     #define __LXW_COMMON_H__

//C     #include <time.h>

//C     #ifndef TESTING
//C     #define STATIC static
//C     #else
//C     #define STATIC
//C     #endif

//C     #define LXW_SHEETNAME_MAX  32
//C     #define LXW_SHEETNAME_LEN  65

//C     enum lxw_boolean {
//C         LXW_FALSE,
//C         LXW_TRUE
//C     };

//C     #define LXW_IGNORE 1

//C     #define ERROR(message)                              fprintf(stderr, "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

//C     #define MEM_ERROR()                                 ERROR("Memory allocation failed.")

//C     #define GOTO_LABEL_ON_MEM_ERROR(pointer, label)     if (!pointer) {                                     MEM_ERROR();                                    goto label;                                 }

//C     #define RETURN_ON_MEM_ERROR(pointer, error)         if (!pointer) {                                     MEM_ERROR();                                    return error;                               }

//C     #define LXW_WARN(message)                           fprintf(stderr, "[WARN]: " message "\n")

/* Define the queue.h structs for the formats list. */
//C     struct lxw_formats {
//C         struct lxw_format *stqh_first;/* first element */
//C         struct lxw_format **stqh_last;/* addr of last next element */
//C     };

/* Define the queue.h structs for the generic data structs. */
//C     struct lxw_tuples {
//C         struct lxw_tuple *stqh_first;/* first element */
//C         struct lxw_tuple **stqh_last;/* addr of last next element */
//C     };

//C     typedef struct lxw_tuple {
//C         char *key;
//C         char *value;

//C         struct {
//C             struct lxw_tuple *stqe_next; /* next element */
//C         } list_pointers;
//C     } lxw_tuple;

//C     typedef struct lxw_doc_properties {
//C         char *title;
//C         char *subject;
//C         char *author;
//C         char *manager;
//C         char *company;
//C         char *category;
//C         char *keywords;
//C         char *comments;
//C         char *status;
//C         time_t created;
//C     } lxw_doc_properties;


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_COMMON_H__ */

/**
 * @brief The type for RGB colors in libxlswriter.
 *
 * The type for RGB colors in libxlswriter. The valid range is `0x000000`
 * (black) to `0xFFFFFF` (white). See @ref working_with_colors.
 */
//C     typedef int32_t lxw_color_t;
alias int32_t lxw_color_t;

//C     #define LXW_FORMAT_FIELD_LEN            128
//C     #define LXW_DEFAULT_FONT_NAME           "Calibri"
const LXW_FORMAT_FIELD_LEN = 128;
//C     #define LXW_DEFAULT_FONT_FAMILY         2
//C     #define LXW_DEFAULT_FONT_THEME          1
const LXW_DEFAULT_FONT_FAMILY = 2;
//C     #define LXW_PROPERTY_UNSET              -1
const LXW_DEFAULT_FONT_THEME = 1;
//C     #define LXW_COLOR_UNSET                 -1
const LXW_PROPERTY_UNSET = -1;
//C     #define LXW_COLOR_MASK                  0xFFFFFF
const LXW_COLOR_UNSET = -1;
//C     #define LXW_MIN_FONT_SIZE               1
const LXW_COLOR_MASK = 0xFFFFFF;
//C     #define LXW_MAX_FONT_SIZE               409
const LXW_MIN_FONT_SIZE = 1;

const LXW_MAX_FONT_SIZE = 409;
//C     #define LXW_FORMAT_FIELD_COPY(dst, src)                 do{                                                     strncpy(dst, src, LXW_FORMAT_FIELD_LEN -1);         dst[LXW_FORMAT_FIELD_LEN - 1] = '\0';           } while (0)

/** Format underline values for format_set_underline(). */
//C     enum lxw_format_underlines {
    /** Single underline */
//C         LXW_UNDERLINE_SINGLE = 1,

    /** Double underline */
//C         LXW_UNDERLINE_DOUBLE,

    /** Single accounting underline */
//C         LXW_UNDERLINE_SINGLE_ACCOUNTING,

    /** Double accounting underline */
//C         LXW_UNDERLINE_DOUBLE_ACCOUNTING
//C     };
enum lxw_format_underlines
{
    LXW_UNDERLINE_SINGLE = 1,
    LXW_UNDERLINE_DOUBLE,
    LXW_UNDERLINE_SINGLE_ACCOUNTING,
    LXW_UNDERLINE_DOUBLE_ACCOUNTING,
}

/** Superscript and subscript values for format_set_font_script(). */
//C     enum lxw_format_scripts {

    /** Superscript font */
//C         LXW_FONT_SUPERSCRIPT = 1,

    /** Subscript font */
//C         LXW_FONT_SUBSCRIPT
//C     };
enum lxw_format_scripts
{
    LXW_FONT_SUPERSCRIPT = 1,
    LXW_FONT_SUBSCRIPT,
}

/** Alignment values for format_set_align(). */
//C     enum lxw_format_alignments {
    /** No alignment. Cell will use Excel's default for the data type */
//C         LXW_ALIGN_NONE = 0,

    /** Left horizontal alignment */
//C         LXW_ALIGN_LEFT,

    /** Center horizontal alignment */
//C         LXW_ALIGN_CENTER,

    /** Right horizontal alignment */
//C         LXW_ALIGN_RIGHT,

    /** Cell fill horizontal alignment */
//C         LXW_ALIGN_FILL,

    /** Justify horizontal alignment */
//C         LXW_ALIGN_JUSTIFY,

    /** Center Across horizontal alignment */
//C         LXW_ALIGN_CENTER_ACROSS,

    /** Left horizontal alignment */
//C         LXW_ALIGN_DISTRIBUTED,

    /** Top vertical alignment */
//C         LXW_ALIGN_VERTICAL_TOP,

    /** Bottom vertical alignment */
//C         LXW_ALIGN_VERTICAL_BOTTOM,

    /** Center vertical alignment */
//C         LXW_ALIGN_VERTICAL_CENTER,

    /** Justify vertical alignment */
//C         LXW_ALIGN_VERTICAL_JUSTIFY,

    /** Distributed vertical alignment */
//C         LXW_ALIGN_VERTICAL_DISTRIBUTED
//C     };
enum lxw_format_alignments
{
    LXW_ALIGN_NONE,
    LXW_ALIGN_LEFT,
    LXW_ALIGN_CENTER,
    LXW_ALIGN_RIGHT,
    LXW_ALIGN_FILL,
    LXW_ALIGN_JUSTIFY,
    LXW_ALIGN_CENTER_ACROSS,
    LXW_ALIGN_DISTRIBUTED,
    LXW_ALIGN_VERTICAL_TOP,
    LXW_ALIGN_VERTICAL_BOTTOM,
    LXW_ALIGN_VERTICAL_CENTER,
    LXW_ALIGN_VERTICAL_JUSTIFY,
    LXW_ALIGN_VERTICAL_DISTRIBUTED,
}

//C     enum lxw_format_diagonal_types {
//C         LXW_DIAGONAL_BORDER_UP = 1,
//C         LXW_DIAGONAL_BORDER_DOWN,
//C         LXW_DIAGONAL_BORDER_UP_DOWN
//C     };
enum lxw_format_diagonal_types
{
    LXW_DIAGONAL_BORDER_UP = 1,
    LXW_DIAGONAL_BORDER_DOWN,
    LXW_DIAGONAL_BORDER_UP_DOWN,
}

/** Predefined values for common colors. */
//C     enum lxw_defined_colors {
    /** Black */
//C         LXW_COLOR_BLACK = 0x000000,

    /** Blue */
//C         LXW_COLOR_BLUE = 0x0000FF,

    /** Brown */
//C         LXW_COLOR_BROWN = 0x800000,

    /** Cyan */
//C         LXW_COLOR_CYAN = 0x00FFFF,

    /** Gray */
//C         LXW_COLOR_GRAY = 0x808080,

    /** Green */
//C         LXW_COLOR_GREEN = 0x008000,

    /** Lime */
//C         LXW_COLOR_LIME = 0x00FF00,

    /** Magenta */
//C         LXW_COLOR_MAGENTA = 0xFF00FF,

    /** Navy */
//C         LXW_COLOR_NAVY = 0x000080,

    /** Orange */
//C         LXW_COLOR_ORANGE = 0xFF6600,

    /** Pink */
//C         LXW_COLOR_PINK = 0xFF00FF,

    /** Purple */
//C         LXW_COLOR_PURPLE = 0x800080,

    /** Red */
//C         LXW_COLOR_RED = 0xFF0000,

    /** Silver */
//C         LXW_COLOR_SILVER = 0xC0C0C0,

    /** White */
//C         LXW_COLOR_WHITE = 0xFFFFFF,

    /** Yellow */
//C         LXW_COLOR_YELLOW = 0xFFFF00
//C     };
enum lxw_defined_colors
{
    LXW_COLOR_BLACK,
    LXW_COLOR_BLUE = 255,
    LXW_COLOR_BROWN = 8388608,
    LXW_COLOR_CYAN = 65535,
    LXW_COLOR_GRAY = 8421504,
    LXW_COLOR_GREEN = 32768,
    LXW_COLOR_LIME = 65280,
    LXW_COLOR_MAGENTA = 16711935,
    LXW_COLOR_NAVY = 128,
    LXW_COLOR_ORANGE = 16737792,
    LXW_COLOR_PINK = 16711935,
    LXW_COLOR_PURPLE = 8388736,
    LXW_COLOR_RED = 16711680,
    LXW_COLOR_SILVER = 12632256,
    LXW_COLOR_WHITE = 16777215,
    LXW_COLOR_YELLOW = 16776960,
}

/** Pattern value for use with format_set_pattern(). */
//C     enum lxw_format_patterns {
    /** Empty pattern */
//C         LXW_PATTERN_NONE = 0,

    /** Solid pattern */
//C         LXW_PATTERN_SOLID,

    /** Medium gray pattern */
//C         LXW_PATTERN_MEDIUM_GRAY,

    /** Dark gray pattern */
//C         LXW_PATTERN_DARK_GRAY,

    /** Light gray pattern */
//C         LXW_PATTERN_LIGHT_GRAY,

    /** Dark horizontal line pattern */
//C         LXW_PATTERN_DARK_HORIZONTAL,

    /** Dark vertical line pattern */
//C         LXW_PATTERN_DARK_VERTICAL,

    /** Dark diagonal stripe pattern */
//C         LXW_PATTERN_DARK_DOWN,

    /** Reverse dark diagonal stripe pattern */
//C         LXW_PATTERN_DARK_UP,

    /** Dark grid pattern */
//C         LXW_PATTERN_DARK_GRID,

    /** Dark trellis pattern */
//C         LXW_PATTERN_DARK_TRELLIS,

    /** Light horizontal Line pattern */
//C         LXW_PATTERN_LIGHT_HORIZONTAL,

    /** Light vertical line pattern */
//C         LXW_PATTERN_LIGHT_VERTICAL,

    /** Light diagonal stripe pattern */
//C         LXW_PATTERN_LIGHT_DOWN,

    /** Reverse light diagonal stripe pattern */
//C         LXW_PATTERN_LIGHT_UP,

    /** Light grid pattern */
//C         LXW_PATTERN_LIGHT_GRID,

    /** Light trellis pattern */
//C         LXW_PATTERN_LIGHT_TRELLIS,

    /** 12.5% gray pattern */
//C         LXW_PATTERN_GRAY_125,

    /** 6.25% gray pattern */
//C         LXW_PATTERN_GRAY_0625
//C     };
enum lxw_format_patterns
{
    LXW_PATTERN_NONE,
    LXW_PATTERN_SOLID,
    LXW_PATTERN_MEDIUM_GRAY,
    LXW_PATTERN_DARK_GRAY,
    LXW_PATTERN_LIGHT_GRAY,
    LXW_PATTERN_DARK_HORIZONTAL,
    LXW_PATTERN_DARK_VERTICAL,
    LXW_PATTERN_DARK_DOWN,
    LXW_PATTERN_DARK_UP,
    LXW_PATTERN_DARK_GRID,
    LXW_PATTERN_DARK_TRELLIS,
    LXW_PATTERN_LIGHT_HORIZONTAL,
    LXW_PATTERN_LIGHT_VERTICAL,
    LXW_PATTERN_LIGHT_DOWN,
    LXW_PATTERN_LIGHT_UP,
    LXW_PATTERN_LIGHT_GRID,
    LXW_PATTERN_LIGHT_TRELLIS,
    LXW_PATTERN_GRAY_125,
    LXW_PATTERN_GRAY_0625,
}

/** Cell border styles for use with format_set_border(). */
//C     enum lxw_format_borders {
    /** No border */
//C         LXW_BORDER_NONE,

    /** Thin border style */
//C         LXW_BORDER_THIN,

    /** Medium border style */
//C         LXW_BORDER_MEDIUM,

    /** Dashed border style */
//C         LXW_BORDER_DASHED,

    /** Dotted border style */
//C         LXW_BORDER_DOTTED,

    /** Thick border style */
//C         LXW_BORDER_THICK,

    /** Double border style */
//C         LXW_BORDER_DOUBLE,

    /** Hair border style */
//C         LXW_BORDER_HAIR,

    /** Medium dashed border style */
//C         LXW_BORDER_MEDIUM_DASHED,

    /** Dash-dot border style */
//C         LXW_BORDER_DASH_DOT,

    /** Medium dash-dot border style */
//C         LXW_BORDER_MEDIUM_DASH_DOT,

    /** Dash-dot-dot border style */
//C         LXW_BORDER_DASH_DOT_DOT,

    /** Medium dash-dot-dot border style */
//C         LXW_BORDER_MEDIUM_DASH_DOT_DOT,

    /** Slant dash-dot border style */
//C         LXW_BORDER_SLANT_DASH_DOT
//C     };
enum lxw_format_borders
{
    LXW_BORDER_NONE,
    LXW_BORDER_THIN,
    LXW_BORDER_MEDIUM,
    LXW_BORDER_DASHED,
    LXW_BORDER_DOTTED,
    LXW_BORDER_THICK,
    LXW_BORDER_DOUBLE,
    LXW_BORDER_HAIR,
    LXW_BORDER_MEDIUM_DASHED,
    LXW_BORDER_DASH_DOT,
    LXW_BORDER_MEDIUM_DASH_DOT,
    LXW_BORDER_DASH_DOT_DOT,
    LXW_BORDER_MEDIUM_DASH_DOT_DOT,
    LXW_BORDER_SLANT_DASH_DOT,
}

/**
 * @brief Struct to represent the formatting properties of an Excel format.
 *
 * Formats in `libxlswriter` are accessed via this struct.
 *
 * The members of the lxw_format struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in format.h.
 *
 * For example:
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
 *
 * @endcode
 *
 */
//C     typedef struct lxw_format {

//C         FILE *file;

//C         lxw_hash_table *xf_format_indices;
//C         uint16_t *num_xf_formats;

//C         int32_t xf_index;
//C         int32_t dxf_index;

//C         char num_format[LXW_FORMAT_FIELD_LEN];
//C         char font_name[LXW_FORMAT_FIELD_LEN];
//C         char font_scheme[LXW_FORMAT_FIELD_LEN];
//C         uint8_t num_format_index;
//C         uint16_t font_index;
//C         uint8_t has_font;
//C         uint8_t has_dxf_font;
//C         uint16_t font_size;
//C         uint8_t bold;
//C         uint8_t italic;
//C         lxw_color_t font_color;
//C         uint8_t underline;
//C         uint8_t font_strikeout;
//C         uint8_t font_outline;
//C         uint8_t font_shadow;
//C         uint8_t font_script;
//C         uint8_t font_family;
//C         uint8_t font_charset;
//C         uint8_t font_condense;
//C         uint8_t font_extend;
//C         uint8_t theme;
//C         uint8_t hyperlink;

//C         uint8_t hidden;
//C         uint8_t locked;

//C         uint8_t text_h_align;
//C         uint8_t text_wrap;
//C         uint8_t text_v_align;
//C         uint8_t text_justlast;
//C         int16_t rotation;

//C         lxw_color_t fg_color;
//C         lxw_color_t bg_color;
//C         uint8_t pattern;
//C         uint8_t has_fill;
//C         uint8_t has_dxf_fill;
//C         int32_t fill_index;
//C         int32_t fill_count;

//C         int32_t border_index;
//C         uint8_t has_border;
//C         uint8_t has_dxf_border;
//C         int32_t border_count;

//C         uint8_t bottom;
//C         uint8_t diag_border;
//C         uint8_t diag_type;
//C         uint8_t left;
//C         uint8_t right;
//C         uint8_t top;
//C         lxw_color_t bottom_color;
//C         lxw_color_t diag_color;
//C         lxw_color_t left_color;
//C         lxw_color_t right_color;
//C         lxw_color_t top_color;

//C         uint8_t indent;
//C         uint8_t shrink;
//C         uint8_t merge_range;
//C         uint8_t reading_order;
//C         uint8_t just_distrib;
//C         uint8_t color_indexed;
//C         uint8_t font_only;

//C         struct {
//C             struct lxw_format *stqe_next; /* next element */
//C         } list_pointers;
struct _N6
{
    lxw_format *stqe_next;
}
//C     } lxw_format;
struct lxw_format
{
    FILE *file;
    lxw_hash_table *xf_format_indices;
    uint16_t *num_xf_formats;
    int32_t xf_index;
    int32_t dxf_index;
    char [128]num_format;
    char [128]font_name;
    char [128]font_scheme;
    uint8_t num_format_index;
    uint16_t font_index;
    uint8_t has_font;
    uint8_t has_dxf_font;
    uint16_t font_size;
    uint8_t bold;
    uint8_t italic;
    lxw_color_t font_color;
    uint8_t underline;
    uint8_t font_strikeout;
    uint8_t font_outline;
    uint8_t font_shadow;
    uint8_t font_script;
    uint8_t font_family;
    uint8_t font_charset;
    uint8_t font_condense;
    uint8_t font_extend;
    uint8_t theme;
    uint8_t hyperlink;
    uint8_t hidden;
    uint8_t locked;
    uint8_t text_h_align;
    uint8_t text_wrap;
    uint8_t text_v_align;
    uint8_t text_justlast;
    int16_t rotation;
    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t pattern;
    uint8_t has_fill;
    uint8_t has_dxf_fill;
    int32_t fill_index;
    int32_t fill_count;
    int32_t border_index;
    uint8_t has_border;
    uint8_t has_dxf_border;
    int32_t border_count;
    uint8_t bottom;
    uint8_t diag_border;
    uint8_t diag_type;
    uint8_t left;
    uint8_t right;
    uint8_t top;
    lxw_color_t bottom_color;
    lxw_color_t diag_color;
    lxw_color_t left_color;
    lxw_color_t right_color;
    lxw_color_t top_color;
    uint8_t indent;
    uint8_t shrink;
    uint8_t merge_range;
    uint8_t reading_order;
    uint8_t just_distrib;
    uint8_t color_indexed;
    uint8_t font_only;
    _N6 list_pointers;
}

/*
 * Struct to represent the font component of a format.
 */
//C     typedef struct lxw_font {

//C         char font_name[LXW_FORMAT_FIELD_LEN];
//C         uint16_t font_size;
//C         uint8_t bold;
//C         uint8_t italic;
//C         lxw_color_t font_color;
//C         uint8_t underline;
//C         uint8_t font_strikeout;
//C         uint8_t font_outline;
//C         uint8_t font_shadow;
//C         uint8_t font_script;
//C         uint8_t font_family;
//C         uint8_t font_charset;
//C         uint8_t font_condense;
//C         uint8_t font_extend;
//C     } lxw_font;
struct lxw_font
{
    char [128]font_name;
    uint16_t font_size;
    uint8_t bold;
    uint8_t italic;
    lxw_color_t font_color;
    uint8_t underline;
    uint8_t font_strikeout;
    uint8_t font_outline;
    uint8_t font_shadow;
    uint8_t font_script;
    uint8_t font_family;
    uint8_t font_charset;
    uint8_t font_condense;
    uint8_t font_extend;
}

/*
 * Struct to represent the border component of a format.
 */
//C     typedef struct lxw_border {

//C         uint8_t bottom;
//C         uint8_t diag_border;
//C         uint8_t diag_type;
//C         uint8_t left;
//C         uint8_t right;
//C         uint8_t top;

//C         lxw_color_t bottom_color;
//C         lxw_color_t diag_color;
//C         lxw_color_t left_color;
//C         lxw_color_t right_color;
//C         lxw_color_t top_color;

//C     } lxw_border;
struct lxw_border
{
    uint8_t bottom;
    uint8_t diag_border;
    uint8_t diag_type;
    uint8_t left;
    uint8_t right;
    uint8_t top;
    lxw_color_t bottom_color;
    lxw_color_t diag_color;
    lxw_color_t left_color;
    lxw_color_t right_color;
    lxw_color_t top_color;
}

/*
 * Struct to represent the fill component of a format.
 */
//C     typedef struct lxw_fill {

//C         lxw_color_t fg_color;
//C         lxw_color_t bg_color;
//C         uint8_t pattern;

//C     } lxw_fill;
struct lxw_fill
{
    lxw_color_t fg_color;
    lxw_color_t bg_color;
    uint8_t pattern;
}


/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     lxw_format *_new_format();
lxw_format * _new_format();
//C     void _free_format(lxw_format *format);
void  _free_format(lxw_format *format);
//C     int32_t _get_xf_index(lxw_format *format);
int32_t  _get_xf_index(lxw_format *format);
//C     lxw_font *_get_font_key(lxw_format *format);
lxw_font * _get_font_key(lxw_format *format);
//C     lxw_border *_get_border_key(lxw_format *format);
lxw_border * _get_border_key(lxw_format *format);
//C     lxw_fill *_get_fill_key(lxw_format *format);
lxw_fill * _get_fill_key(lxw_format *format);

/**
 * @brief Set the font used in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param font_name Cell font name.
 *
 * Specify the font used used in the cell format:
 *
 * @code
 *     format_set_font_name(format, "Avenir Black Oblique");
 * @endcode
 *
 * @image html format_set_font_name.png
 *
 * Excel can only display fonts that are installed on the system that it is
 * running on. Therefore it is generally best to use the fonts that come as
 * standard with Excel such as Calibri, Times New Roman and Courier New.
 *
 * The default font in Excel 2007, and later, is Calibri.
 */
//C     void format_set_font_name(lxw_format *format, const char *font_name);
void  format_set_font_name(lxw_format *format, char *font_name);

/**
 * @brief Set the size of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param size   The cell font size.
 *
 * Set the font size of the cell format:
 *
 * @code
 *     format_set_font_size(format, 30);
 * @endcode
 *
 * Excel adjusts the height of a row to accommodate the largest font
 * size in the row. You can also explicitly specify the height of a
 * row using the worksheet_set_row() function.
 */
//C     void format_set_font_size(lxw_format *format, uint16_t size);
void  format_set_font_size(lxw_format *format, uint16_t size);

/**
 * @brief Set the color of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell font color.
 *
 *
 * Set the font color:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_font_color(format, "red");
 *
 *     worksheet_write_string(worksheet, 0, 0, "wheelbarrow", format);
 * @endcode
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 * @note
 * The format_set_font_color() method is used to set the font color in a
 * cell. To set the color of a cell background use the format_set_bg_color()
 * and format_set_pattern() methods.
 */
//C     void format_set_font_color(lxw_format *format, lxw_color_t color);
void  format_set_font_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Turn on bold for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the bold property of the font:
 *
 * @code
 *     format_set_bold(format);
 * @endcode
 */
//C     void format_set_bold(lxw_format *format);
void  format_set_bold(lxw_format *format);

/**
 * @brief Turn on italic for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the italic property of the font:
 *
 * @code
 *     format_set_italic(format);
 * @endcode
 */
//C     void format_set_italic(lxw_format *format);
void  format_set_italic(lxw_format *format);

/**
 * @brief Turn on underline for the format:
 *
 * @param format Pointer to a Format instance.
 * @param style Underline style.
 *
 * Set the underline property of the format:
 *
 * @code
 *     format_set_underline(format, LXW_UNDERLINE_SINGLE);
 * @endcode
 *
 * The available underline styles are:
 *
 * - #LXW_UNDERLINE_SINGLE
 * - #LXW_UNDERLINE_DOUBLE
 * - #LXW_UNDERLINE_SINGLE_ACCOUNTING
 * - #LXW_UNDERLINE_DOUBLE_ACCOUNTING
 *
 */
//C     void format_set_underline(lxw_format *format, uint8_t style);
void  format_set_underline(lxw_format *format, uint8_t style);

/**
 * @brief Set the strikeout property of the font.
 *
 * @param format Pointer to a Format instance.
 */
//C     void format_set_font_strikeout(lxw_format *format);
void  format_set_font_strikeout(lxw_format *format);

/**
 * @brief Set the superscript/subscript property of the font.
 *
 * @param format Pointer to a Format instance.
 * @param style  Superscript or subscript style.
 *
 * Set the superscript o subscript property of the font.
 *
 * The available script styles are:
 *
 * - #LXW_FONT_SUPERSCRIPT
 * - #LXW_FONT_SUBSCRIPT
 */
//C     void format_set_font_script(lxw_format *format, uint8_t style);
void  format_set_font_script(lxw_format *format, uint8_t style);

/**
 * @brief Set the number format for a cell.
 *
 * @param format      Pointer to a Format instance.
 * @param num_format The cell number format string.
 *
 * This method is used to define the numerical format of a number in
 * Excel. It controls whether a number is displayed as an integer, a
 * floating point number, a date, a currency value or some other user
 * defined format.
 *
 * The numerical format of a cell can be specified by using a format
 * string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_num_format(format, "d mmm yyyy");
 * @endcode
 *
 * Format strings can control any aspect of number formatting allowed by Excel:
 *
 * @dontinclude format_num_format.c
 * @skipline set_num_format
 * @until 1209
 * 
 * @image html format_set_num_format.png
 *
 * The number system used for dates is described in @ref working_with_dates.
 *
 * For more information on number formats in Excel refer to the
 * [Microsoft documentation on cell formats](http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx).
 */
//C     void format_set_num_format(lxw_format *format, const char *num_format);
void  format_set_num_format(lxw_format *format, char *num_format);

/**
 * @brief Set the Excel built-in number format for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param index  The built-in number format index for the cell.
 *
 * This function is similar to format_set_num_format() except that it takes an
 * index to a limited number of Excel's built-in number formats instead of a
 * user defined format string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_num_format(format, 0x0F);     // d-mmm-yy
 * @endcode
 *
 * @note
 *
 * Unless you need to specifically access one of Excel's built-in number
 * formats the format_set_num_format() function above is a better
 * solution. The format_set_num_format_index() function is mainly included for
 * backward compatibility and completeness.
 *
 * The Excel built-in number formats as shown in the table below:
 *
 *   | Index | Index | Format String                                        |
 *   | ----- | ----- | ---------------------------------------------------- |
 *   | 0     | 0x00  | `General`                                            |
 *   | 1     | 0x01  | `0`                                                  |
 *   | 2     | 0x02  | `0.00`                                               |
 *   | 3     | 0x03  | `#,##0`                                              |
 *   | 4     | 0x04  | `#,##0.00`                                           |
 *   | 5     | 0x05  | `($#,##0_);($#,##0)`                                 |
 *   | 6     | 0x06  | `($#,##0_);[Red]($#,##0)`                            |
 *   | 7     | 0x07  | `($#,##0.00_);($#,##0.00)`                           |
 *   | 8     | 0x08  | `($#,##0.00_);[Red]($#,##0.00)`                      |
 *   | 9     | 0x09  | `0%`                                                 |
 *   | 10    | 0x0a  | `0.00%`                                              |
 *   | 11    | 0x0b  | `0.00E+00`                                           |
 *   | 12    | 0x0c  | `# ?/?`                                              |
 *   | 13    | 0x0d  | `# ??/??`                                            |
 *   | 14    | 0x0e  | `m/d/yy`                                             |
 *   | 15    | 0x0f  | `d-mmm-yy`                                           |
 *   | 16    | 0x10  | `d-mmm`                                              |
 *   | 17    | 0x11  | `mmm-yy`                                             |
 *   | 18    | 0x12  | `h:mm AM/PM`                                         |
 *   | 19    | 0x13  | `h:mm:ss AM/PM`                                      |
 *   | 20    | 0x14  | `h:mm`                                               |
 *   | 21    | 0x15  | `h:mm:ss`                                            |
 *   | 22    | 0x16  | `m/d/yy h:mm`                                        |
 *   | ...   | ...   | ...                                                  |
 *   | 37    | 0x25  | `(#,##0_);(#,##0)`                                   |
 *   | 38    | 0x26  | `(#,##0_);[Red](#,##0)`                              |
 *   | 39    | 0x27  | `(#,##0.00_);(#,##0.00)`                             |
 *   | 40    | 0x28  | `(#,##0.00_);[Red](#,##0.00)`                        |
 *   | 41    | 0x29  | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`            |
 *   | 42    | 0x2a  | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`         |
 *   | 43    | 0x2b  | `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`    |
 *   | 44    | 0x2c  | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` |
 *   | 45    | 0x2d  | `mm:ss`                                              |
 *   | 46    | 0x2e  | `[h]:mm:ss`                                          |
 *   | 47    | 0x2f  | `mm:ss.0`                                            |
 *   | 48    | 0x30  | `##0.0E+0`                                           |
 *   | 49    | 0x31  | `@`                                                  |
 *
 *  @note
 *
 *  -  Numeric formats 23 to 36 are not documented by Microsoft and may differ
 *     in international versions. The listed date and currency formats may also
 *     vary depending on system settings.
 *
 *  - The dollar sign in the above format appears as the defined local currency
 *    symbol.
 *
 *  - These formats can also be set via format_set_num_format().
 */
//C     void format_set_num_format_index(lxw_format *format, uint8_t index);
void  format_set_num_format_index(lxw_format *format, uint8_t index);

/**
 * @brief Set the cell unlocked state.
 *
 * @param format Pointer to a Format instance.
 *
 * This property can be used to allow modification of a cell in a protected
 * worksheet. In Excel, cell locking is turned on by default for all
 * cells. However, it only has an effect if the worksheet has been protected
 * using the worksheet worksheet_protect() method:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_unlocked(format);
 *
 *     // Enable worksheet protection.
 *     worksheet_protect(worksheet);
 *
 *     // This cell cannot be edited.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", NULL);
 *
 *     // This cell can be edited.
 *     worksheet_write_formula(worksheet, 1, 0, "=1+2", format);
 * @endcode
 */
//C     void format_set_unlocked(lxw_format *format);
void  format_set_unlocked(lxw_format *format);

/**
 * @brief Hide formulas in a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This property is used to hide a formula while still displaying its result. This
 * is generally used to hide complex calculations from end users who are only
 * interested in the result. It only has an effect if the worksheet has been
 * protected using the worksheet write_protect() method:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_hidden(format);
 *
 *     // Enable worksheet protection.
 *     worksheet_protect(worksheet);
 *
 *     // The formula in this cell isn't visible.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", format);
 * @endcode
 */
//C     void format_set_hidden(lxw_format *format);
void  format_set_hidden(lxw_format *format);

/**
 * @brief Set the alignment for data in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param alignment The horizontal and or vertical alignment direction.
 *
 * This method is used to set the horizontal and vertical text alignment within a
 * cell. The following are the available horizontal alignments:
 *
 * - #LXW_ALIGN_LEFT
 * - #LXW_ALIGN_CENTER
 * - #LXW_ALIGN_RIGHT
 * - #LXW_ALIGN_FILL
 * - #LXW_ALIGN_JUSTIFY
 * - #LXW_ALIGN_CENTER_ACROSS
 * - #LXW_ALIGN_DISTRIBUTED
 *
 * The following are the available vertical alignments:
 *
 * - #LXW_ALIGN_VERTICAL_TOP
 * - #LXW_ALIGN_VERTICAL_BOTTOM
 * - #LXW_ALIGN_VERTICAL_CENTER
 * - #LXW_ALIGN_VERTICAL_JUSTIFY
 * - #LXW_ALIGN_VERTICAL_DISTRIBUTED
 *
 * As in Excel, vertical and horizontal alignments can be combined:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_align(format, LXW_ALIGN_CENTER);
 *     format_set_align(format, LXW_ALIGN_VERTICAL_CENTER);
 *
 *     worksheet_set_row(0, 30);
 *     worksheet_write_string(worksheet, 0, 0, "Some Text", format);
 * @endcode
 *
 * Text can be aligned across two or more adjacent cells using the
 * center_across property. However, for genuine merged cells it is better to
 * use the worksheet_merge_range() worksheet method.
 *
 * The vertical justify option can be used to provide automatic text wrapping
 * in a cell. The height of the cell will be adjusted to accommodate the
 * wrapped text. To specify where the text wraps use the
 * format_set_text_wrap() method.
 */
//C     void format_set_align(lxw_format *format, uint8_t alignment);
void  format_set_align(lxw_format *format, uint8_t alignment);

/**
 * @brief Wrap text in a cell.
 *
 * Turn text wrapping on for text in a cell.
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_text_wrap(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Some long text to wrap in a cell", format);
 * @endcode
 *
 * If you wish to control where the text is wrapped you can add newline characters
 * to the string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_text_wrap(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "It's\na bum\nwrap", format);
 * @endcode
 *
 * Excel will adjust the height of the row to accommodate the wrapped text. A
 * similar effect can be obtained without newlines using the
 * format_set_align() function with #LXW_ALIGN_VERTICAL_JUSTIFY.
 */
//C     void format_set_text_wrap(lxw_format *format);
void  format_set_text_wrap(lxw_format *format);

/**
 * @brief Set the rotation of the text in a cell.
 *
 * @param format Pointer to a Format instance.
 * @param angle  Rotation angle in the range -90 to 90 and 270.
 *
 * Set the rotation of the text in a cell. The rotation can be any angle in the
 * range -90 to 90 degrees:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_rotation(format, 30);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is rotated", format);
 * @endcode
 *
 * The angle 270 is also supported. This indicates text where the letters run from
 * top to bottom.
 */
//C     void format_set_rotation(lxw_format *format, int16_t angle);
void  format_set_rotation(lxw_format *format, int16_t angle);

/**
 * @brief Set the cell text indentation level.
 *
 * @param format Pointer to a Format instance.
 * @param level  Indentation level.
 *
 * This method can be used to indent text in a cell. The argument, which should be
 * an integer, is taken as the level of indentation:
 *
 * @code
 *     format1 = workbook_add_format(workbook);
 *     format2 = workbook_add_format(workbook);
 *
 *     format_set_indent(format1, 1);
 *     format_set_indent(format2, 2);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is indented 1 level",  format1);
 *     worksheet_write_string(worksheet, 1, 0, "This text is indented 2 levels", format2);
 * @endcode
 *
 * @image html text_indent.png
 *
 * @note
 * Indentation is a horizontal alignment property. It will override any other
 * horizontal properties but it can be used in conjunction with vertical
 * properties.
 */
//C     void format_set_indent(lxw_format *format, uint8_t level);
void  format_set_indent(lxw_format *format, uint8_t level);

/**
 * @brief Turn on the text "shrink to fit" for a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This method can be used to shrink text so that it fits in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_shrink(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Honey, I shrunk the text!", format);
 * @endcode
 */
//C     void format_set_shrink(lxw_format *format);
void  format_set_shrink(lxw_format *format);

/**
 * @brief Set the background fill pattern for a cell
 *
 * @param format Pointer to a Format instance.
 * @param index  Pattern index.
 *
 * Set the background pattern for a cell.
 *
 * The most common pattern is a solid fill of the background color:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID);
 *     format_set_bg_color(format, LXW_COLOR_YELLOW);
 * @endcode
 *
 * The available fill patterns are:
 *
 *    Fill Type                     | Define
 *    ----------------------------- | -----------------------------
 *    Solid                         | #LXW_PATTERN_SOLID
 *    Medium gray                   | #LXW_PATTERN_MEDIUM_GRAY
 *    Dark gray                     | #LXW_PATTERN_DARK_GRAY
 *    Light gray                    | #LXW_PATTERN_LIGHT_GRAY
 *    Dark horizontal line          | #LXW_PATTERN_DARK_HORIZONTAL
 *    Dark vertical line            | #LXW_PATTERN_DARK_VERTICAL
 *    Dark diagonal stripe          | #LXW_PATTERN_DARK_DOWN
 *    Reverse dark diagonal stripe  | #LXW_PATTERN_DARK_UP
 *    Dark grid                     | #LXW_PATTERN_DARK_GRID
 *    Dark trellis                  | #LXW_PATTERN_DARK_TRELLIS
 *    Light horizontal line         | #LXW_PATTERN_LIGHT_HORIZONTAL
 *    Light vertical line           | #LXW_PATTERN_LIGHT_VERTICAL
 *    Light diagonal stripe         | #LXW_PATTERN_LIGHT_DOWN
 *    Reverse light diagonal stripe | #LXW_PATTERN_LIGHT_UP
 *    Light grid                    | #LXW_PATTERN_LIGHT_GRID
 *    Light trellis                 | #LXW_PATTERN_LIGHT_TRELLIS
 *    12.5% gray                    | #LXW_PATTERN_GRAY_125
 *    6.25% gray                    | #LXW_PATTERN_GRAY_0625
 *
 */
//C     void format_set_pattern(lxw_format *format, uint8_t index);
void  format_set_pattern(lxw_format *format, uint8_t index);

/**
 * @brief Set the pattern background color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern background color.
 *
 * The format_set_bg_color() method can be used to set the background color of
 * a pattern. Patterns are defined via the format_set_pattern() method. If a
 * pattern hasn't been defined then a solid fill pattern is used as the
 * default.
 *
 * Here is an example of how to set up a solid fill in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID);
 *     format_set_bg_color(format, LXW_COLOR_GREEN);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Ray", format);
 * @endcode
 *
 * @image html formats_set_bg_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
//C     void format_set_bg_color(lxw_format *format, lxw_color_t color);
void  format_set_bg_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the pattern foreground color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern foreground  color.
 *
 * The format_set_fg_color() method can be used to set the foreground color of
 * a pattern.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
//C     void format_set_fg_color(lxw_format *format, lxw_color_t color);
void  format_set_fg_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the cell border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell border style:
 *
 * @code
 *     format_set_border(format, LXW_BORDER_THIN);
 * @endcode 
 *
 * Individual border elements can be configured using the following functions with
 * the same parameters:
 *
 * - format_set_bottom()
 * - format_set_top()
 * - format_set_left()
 * - format_set_right()
 *
 * A cell border is comprised of a border on the bottom, top, left and right.
 * These can be set to the same value using format_set_border() or
 * individually using the relevant method calls shown above.
 *
 * The following border styles are available:
 *
 * - #LXW_BORDER_THIN
 * - #LXW_BORDER_MEDIUM
 * - #LXW_BORDER_DASHED
 * - #LXW_BORDER_DOTTED
 * - #LXW_BORDER_THICK
 * - #LXW_BORDER_DOUBLE
 * - #LXW_BORDER_HAIR
 * - #LXW_BORDER_MEDIUM_DASHED
 * - #LXW_BORDER_DASH_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT
 * - #LXW_BORDER_DASH_DOT_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT_DOT
 * - #LXW_BORDER_SLANT_DASH_DOT
 *
 *  The most commonly used style is the `thin` style.
 */
//C     void format_set_border(lxw_format *format, uint8_t style);
void  format_set_border(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell bottom border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell bottom border style. See format_set_border() for details on the
 * border styles.
 */
//C     void format_set_bottom(lxw_format *format, uint8_t style);
void  format_set_bottom(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell top border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell top border style. See format_set_border() for details on the border
 * styles.
 */
//C     void format_set_top(lxw_format *format, uint8_t style);
void  format_set_top(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell left border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell left border style. See format_set_border() for details on the
 * border styles.
 */
//C     void format_set_left(lxw_format *format, uint8_t style);
void  format_set_left(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell right border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell right border style. See format_set_border() for details on the
 * border styles.
 */
//C     void format_set_right(lxw_format *format, uint8_t style);
void  format_set_right(lxw_format *format, uint8_t style);

/**
 * @brief Set the color of the cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * Individual border elements can be configured using the following methods with
 * the same parameters:
 *
 * - format_set_bottom_color()
 * - format_set_top_color()
 * - format_set_left_color()
 * - format_set_right_color()
 *
 * Set the color of the cell borders. A cell border is comprised of a border
 * on the bottom, top, left and right. These can be set to the same color
 * using format_set_border_color() or individually using the relevant method
 * calls shown above.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
//C     void format_set_border_color(lxw_format *format, lxw_color_t color);
void  format_set_border_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the bottom cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_bottom_color(lxw_format *format, lxw_color_t color);
void  format_set_bottom_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the top cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_top_color(lxw_format *format, lxw_color_t color);
void  format_set_top_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the left cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_left_color(lxw_format *format, lxw_color_t color);
void  format_set_left_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the right cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_right_color(lxw_format *format, lxw_color_t color);
void  format_set_right_color(lxw_format *format, lxw_color_t color);

//C     void format_set_diag_type(lxw_format *format, uint8_t value);
void  format_set_diag_type(lxw_format *format, uint8_t value);
//C     void format_set_diag_color(lxw_format *format, lxw_color_t color);
void  format_set_diag_color(lxw_format *format, lxw_color_t color);
//C     void format_set_diag_border(lxw_format *format, uint8_t value);
void  format_set_diag_border(lxw_format *format, uint8_t value);
//C     void format_set_font_outline(lxw_format *format);
void  format_set_font_outline(lxw_format *format);
//C     void format_set_font_shadow(lxw_format *format);
void  format_set_font_shadow(lxw_format *format);
//C     void format_set_font_family(lxw_format *format, uint8_t value);
void  format_set_font_family(lxw_format *format, uint8_t value);
//C     void format_set_font_charset(lxw_format *format, uint8_t value);
void  format_set_font_charset(lxw_format *format, uint8_t value);
//C     void format_set_font_scheme(lxw_format *format, const char *font_scheme);
void  format_set_font_scheme(lxw_format *format, char *font_scheme);
//C     void format_set_font_condense(lxw_format *format);
void  format_set_font_condense(lxw_format *format);
//C     void format_set_font_extend(lxw_format *format);
void  format_set_font_extend(lxw_format *format);
//C     void format_set_reading_order(lxw_format *format, uint8_t value);
void  format_set_reading_order(lxw_format *format, uint8_t value);
//C     void format_set_theme(lxw_format *format, uint8_t value);
void  format_set_theme(lxw_format *format, uint8_t value);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_FORMAT_H__ */
//C     #include "utility.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @file utility.h
 *
 * @brief Utility functions for libxlsxwriter.
 *
 * <!-- Copyright 2014-2015, John McNamara, jmcnamara@cpan.org -->
 *
 */

//C     #ifndef __LXW_UTILITY_H__
//C     #define __LXW_UTILITY_H__

//C     #include <stdint.h>
//C     #include "common.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * common - Common functions and defines for the libxlsxwriter library.
 *
 */
//C     #ifndef __LXW_COMMON_H__
//C     #define __LXW_COMMON_H__

//C     #include <time.h>

//C     #ifndef TESTING
//C     #define STATIC static
//C     #else
//C     #define STATIC
//C     #endif

//C     #define LXW_SHEETNAME_MAX  32
//C     #define LXW_SHEETNAME_LEN  65

//C     enum lxw_boolean {
//C         LXW_FALSE,
//C         LXW_TRUE
//C     };

//C     #define LXW_IGNORE 1

//C     #define ERROR(message)                              fprintf(stderr, "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

//C     #define MEM_ERROR()                                 ERROR("Memory allocation failed.")

//C     #define GOTO_LABEL_ON_MEM_ERROR(pointer, label)     if (!pointer) {                                     MEM_ERROR();                                    goto label;                                 }

//C     #define RETURN_ON_MEM_ERROR(pointer, error)         if (!pointer) {                                     MEM_ERROR();                                    return error;                               }

//C     #define LXW_WARN(message)                           fprintf(stderr, "[WARN]: " message "\n")

/* Define the queue.h structs for the formats list. */
//C     struct lxw_formats {
//C         struct lxw_format *stqh_first;/* first element */
//C         struct lxw_format **stqh_last;/* addr of last next element */
//C     };

/* Define the queue.h structs for the generic data structs. */
//C     struct lxw_tuples {
//C         struct lxw_tuple *stqh_first;/* first element */
//C         struct lxw_tuple **stqh_last;/* addr of last next element */
//C     };

//C     typedef struct lxw_tuple {
//C         char *key;
//C         char *value;

//C         struct {
//C             struct lxw_tuple *stqe_next; /* next element */
//C         } list_pointers;
//C     } lxw_tuple;

//C     typedef struct lxw_doc_properties {
//C         char *title;
//C         char *subject;
//C         char *author;
//C         char *manager;
//C         char *company;
//C         char *category;
//C         char *keywords;
//C         char *comments;
//C         char *status;
//C         time_t created;
//C     } lxw_doc_properties;


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_COMMON_H__ */

/* Max col: $XFD\0 */
//C     #define MAX_COL_NAME_LENGTH   5

const MAX_COL_NAME_LENGTH = 5;
/* Max cell: $XFWD$1048576\0 */
//C     #define MAX_CELL_NAME_LENGTH  14

const MAX_CELL_NAME_LENGTH = 14;
/* Max range: $XFWD$1048576:$XFWD$1048576\0 */
//C     #define MAX_CELL_RANGE_LENGTH (MAX_CELL_NAME_LENGTH * 2)

//C     #define EPOCH_1900            0
//C     #define EPOCH_1904            1
const EPOCH_1900 = 0;

const EPOCH_1904 = 1;
/**
 * @brief Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *      worksheet_write_string(worksheet, CELL("A1"), "Foo", NULL);
 *
 *      //Same as:
 *      worksheet_write_string(worksheet, 0, 0,       "Foo", NULL);
 * @endcode
 *
 * @note
 *
 * This macro shouldn't be used in performance critical situations since it
 * expands to two function calls.
 */
//C     #define CELL(cell)     lxw_get_row(cell), lxw_get_col(cell)

/**
 * @brief Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *     worksheet_set_column(worksheet, COLS("B:D"), 20, NULL, NULL);
 *
 *     // Same as:
 *     worksheet_set_column(worksheet, 1, 3,        20, NULL, NULL);
 * @endcode
 *
 */
//C     #define COLS(cols)     lxw_get_col(cols), lxw_get_col_2(cols)

/**
 * @brief Convert an Excel `A1:B2` range into a `(first_row, first_col,
 *        last_row, last_col)` sequence.
 *
 * Convert an Excel `A1:B2` range into a `(first_row, first_col, last_row,
 * last_col)` sequence.
 *
 * This is a little syntactic shortcut to help with worksheet layout.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 */
//C     #define RANGE(range)     lxw_get_row(range), lxw_get_col(range), lxw_get_row_2(range), lxw_get_col_2(range)

/** @brief Struct to represent a date and time in Excel.
 *
 * Struct to represent a date and time in Excel. See @ref working_with_dates.
 */
//C     typedef struct lxw_datetime {

    /** Year     : 1900 - 9999 */
//C         int year;
    /** Month    : 1 - 12 */
//C         int month;
    /** Day      : 1 - 31 */
//C         int day;
    /** Hour     : 0 - 23 */
//C         int hour;
    /** Minute   : 0 - 59 */
//C         int min;
    /** Seconds  : 0 - 59.999 */
//C         double sec;

//C     } lxw_datetime;
struct lxw_datetime
{
    int year;
    int month;
    int day;
    int hour;
    int min;
    double sec;
}

/* Create a quoted version of the worksheet name */
//C     char *lxw_quote_sheetname(char *str);
char * lxw_quote_sheetname(char *str);

 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     void lxw_col_to_name(char *col_name, int col_num, uint8_t absolute);
void  lxw_col_to_name(char *col_name, int col_num, uint8_t absolute);

//C     void lxw_rowcol_to_cell(char *cell_name, int row, int col);
void  lxw_rowcol_to_cell(char *cell_name, int row, int col);

//C     void lxw_rowcol_to_cell_abs(char *cell_name,
//C                                 int row,
//C                                 int col, uint8_t abs_row, uint8_t abs_col);
void  lxw_rowcol_to_cell_abs(char *cell_name, int row, int col, uint8_t abs_row, uint8_t abs_col);

//C     void lxw_range(char *range,
//C                    int first_row, int first_col, int last_row, int last_col);
void  lxw_range(char *range, int first_row, int first_col, int last_row, int last_col);

//C     void lxw_range_abs(char *range,
//C                        int first_row, int first_col, int last_row, int last_col);
void  lxw_range_abs(char *range, int first_row, int first_col, int last_row, int last_col);

//C     uint32_t lxw_get_row(const char *row_str);
uint32_t  lxw_get_row(char *row_str);
//C     uint16_t lxw_get_col(const char *col_str);
uint16_t  lxw_get_col(char *col_str);
//C     uint32_t lxw_get_row_2(const char *row_str);
uint32_t  lxw_get_row_2(char *row_str);
//C     uint16_t lxw_get_col_2(const char *col_str);
uint16_t  lxw_get_col_2(char *col_str);

//C     double _datetime_to_excel_date(lxw_datetime *datetime, uint8_t date_1904);
double  _datetime_to_excel_date(lxw_datetime *datetime, uint8_t date_1904);

//C     char *lxw_strdup(const char *str);
char * lxw_strdup(char *str);

//C     void lxw_str_tolower(char *str);
void  lxw_str_tolower(char *str);

//C     FILE *lxw_tmpfile(void);
FILE * lxw_tmpfile();

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_UTILITY_H__ */

//C     #define LXW_ROW_MAX 1048576
//C     #define LXW_COL_MAX 16384
const LXW_ROW_MAX = 1048576;
//C     #define LXW_COL_META_MAX 128
const LXW_COL_MAX = 16384;
//C     #define LXW_HEADER_FOOTER_MAX 255
const LXW_COL_META_MAX = 128;

const LXW_HEADER_FOOTER_MAX = 255;
/* The Excel 2007 specification says that the maximum number of page
 * breaks is 1026. However, in practice it is actually 1023. */
//C     #define LXW_BREAKS_MAX 1023

const LXW_BREAKS_MAX = 1023;
/** Default column width in Excel */
//C     #define LXW_DEF_COL_WIDTH 8.43

const LXW_DEF_COL_WIDTH = 8.43;
/** Default row height in Excel */
//C     #define LXW_DEF_ROW_HEIGHT 15

const LXW_DEF_ROW_HEIGHT = 15;
/** Error codes from `worksheet_write*()` functions. */
//C     enum lxw_write_error {
    /** No error. */
//C         LXW_WRITE_ERROR_NONE = 0,
    /** Row or column index out of range. */
//C         LXW_RANGE_ERROR,
    /** String exceeds Excel's LXW_STRING_LENGTH_ERROR limit. */
//C         LXW_STRING_LENGTH_ERROR,
    /** Error finding string index. */
//C         LXW_STRING_HASH_ERROR
//C     };
enum lxw_write_error
{
    LXW_WRITE_ERROR_NONE,
    LXW_RANGE_ERROR,
    LXW_STRING_LENGTH_ERROR,
    LXW_STRING_HASH_ERROR,
}

/** Gridline options using in `worksheet_gridlines()`. */
//C     enum lxw_gridlines {
    /** Hide screen and print gridlines. */
//C         LXW_HIDE_ALL_GRIDLINES = 0,
    /** Show screen gridlines. */
//C         LXW_SHOW_SCREEN_GRIDLINES,
    /** Show print gridlines. */
//C         LXW_SHOW_PRINT_GRIDLINES,
    /** Show screen and print gridlines. */
//C         LXW_SHOW_ALL_GRIDLINES
//C     };
enum lxw_gridlines
{
    LXW_HIDE_ALL_GRIDLINES,
    LXW_SHOW_SCREEN_GRIDLINES,
    LXW_SHOW_PRINT_GRIDLINES,
    LXW_SHOW_ALL_GRIDLINES,
}

/** Data type to represent a row value.
 *
 * The maximum row in Excel is 1,048,576.
 */
//C     typedef uint32_t lxw_row_t;
alias uint32_t lxw_row_t;

/** Data type to represent a column value.
 *
 * The maximum column in Excel is 16,384.
 */
//C     typedef uint16_t lxw_col_t;
alias uint16_t lxw_col_t;

//C     enum cell_types {
//C         NUMBER_CELL = 1,
//C         STRING_CELL,
//C         INLINE_STRING_CELL,
//C         FORMULA_CELL,
//C         ARRAY_FORMULA_CELL,
//C         BLANK_CELL,
//C         HYPERLINK_URL,
//C         HYPERLINK_INTERNAL,
//C         HYPERLINK_EXTERNAL
//C     };
enum cell_types
{
    NUMBER_CELL = 1,
    STRING_CELL,
    INLINE_STRING_CELL,
    FORMULA_CELL,
    ARRAY_FORMULA_CELL,
    BLANK_CELL,
    HYPERLINK_URL,
    HYPERLINK_INTERNAL,
    HYPERLINK_EXTERNAL,
}

/* Define the queue.h TAILQ structs for the list head types. */
//C     struct lxw_table_cells {
//C         struct lxw_cell *tqh_first; /* first element */
//C         struct lxw_cell **tqh_last; /* addr of last next element */
//C     };
struct lxw_table_cells
{
    lxw_cell *tqh_first;
    lxw_cell **tqh_last;
}
//C     struct lxw_table_rows {
//C         struct lxw_row *tqh_first; /* first element */
//C         struct lxw_row **tqh_last; /* addr of last next element */
//C     };
struct lxw_table_rows
{
    lxw_row *tqh_first;
    lxw_row **tqh_last;
}
//C     struct lxw_merged_ranges {
//C         struct lxw_merged_range *stqh_first;/* first element */
//C         struct lxw_merged_range **stqh_last;/* addr of last next element */
//C     };
struct lxw_merged_ranges
{
    lxw_merged_range *stqh_first;
    lxw_merged_range **stqh_last;
}

/**
 * @brief Options for rows and columns.
 *
 * Options struct for the worksheet_set_column() and worksheet_set_row()
 * functions.
 *
 * It has the following members but currently only the `hidden` property is
 * supported:
 *
 * * `hidden`
 * * `level`
 * * `collapsed`
 */
//C     typedef struct lxw_row_col_options {
    /** Hide the row/column */
//C         uint8_t hidden;
//C         uint8_t level;
//C         uint8_t collapsed;
//C     } lxw_row_col_options;
struct lxw_row_col_options
{
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
}

//C     typedef struct lxw_col_options {
//C         lxw_col_t firstcol;
//C         lxw_col_t lastcol;
//C         double width;
//C         lxw_format *format;
//C         uint8_t hidden;
//C         uint8_t level;
//C         uint8_t collapsed;
//C     } lxw_col_options;
struct lxw_col_options
{
    lxw_col_t firstcol;
    lxw_col_t lastcol;
    double width;
    lxw_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
}

//C     typedef struct lxw_merged_range {
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;

//C         struct {
//C             struct lxw_merged_range *stqe_next; /* next element */
//C         } list_pointers;
struct _N7
{
    lxw_merged_range *stqe_next;
}
//C     } lxw_merged_range;
struct lxw_merged_range
{
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
    _N7 list_pointers;
}

//C     typedef struct lxw_repeat_rows {
//C         uint8_t in_use;
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C     } lxw_repeat_rows;
struct lxw_repeat_rows
{
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
}

//C     typedef struct lxw_repeat_cols {
//C         uint8_t in_use;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;
//C     } lxw_repeat_cols;
struct lxw_repeat_cols
{
    uint8_t in_use;
    lxw_col_t first_col;
    lxw_col_t last_col;
}

//C     typedef struct lxw_print_area {
//C         uint8_t in_use;
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;
//C     } lxw_print_area;
struct lxw_print_area
{
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
}

//C     typedef struct lxw_autofilter {
//C         uint8_t in_use;
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;
//C     } lxw_autofilter;
struct lxw_autofilter
{
    uint8_t in_use;
    lxw_row_t first_row;
    lxw_row_t last_row;
    lxw_col_t first_col;
    lxw_col_t last_col;
}

/**
 * @brief Header and footer options.
 *
 * Optional parameters used in the worksheet_set_header_opt() and
 * worksheet_set_footer_opt() functions.
 *
 */
//C     typedef struct lxw_header_footer_options {
    /** Header or footer margin in inches. Excel default is 0.3. */
//C         double margin;
//C     } lxw_header_footer_options;
struct lxw_header_footer_options
{
    double margin;
}

/**
 * @brief Struct to represent an Excel worksheet.
 *
 * The members of the lxw_worksheet struct aren't modified directly. Instead
 * the worksheet properties are set by calling the functions shown in
 * worksheet.h.
 */
//C     typedef struct lxw_worksheet {

//C         FILE *file;
//C         FILE *optimize_tmpfile;
//C         struct lxw_table_rows *table;
//C         struct lxw_table_rows *hyperlinks;
//C         struct lxw_cell **array;
//C         struct lxw_merged_ranges *merged_ranges;

//C         lxw_row_t dim_rowmin;
//C         lxw_row_t dim_rowmax;
//C         lxw_col_t dim_colmin;
//C         lxw_col_t dim_colmax;

//C         lxw_sst *sst;
//C         char *name;
//C         char *quoted_name;

//C         uint32_t index;
//C         uint8_t active;
//C         uint8_t selected;
//C         uint8_t hidden;
//C         uint32_t *active_sheet;

//C         lxw_col_options **col_options;
//C         uint16_t col_options_max;

//C         double *col_sizes;
//C         uint16_t col_sizes_max;

//C         lxw_format **col_formats;
//C         uint16_t col_formats_max;

//C         uint8_t col_size_changed;
//C         uint8_t optimize;
//C         struct lxw_row *optimize_row;

//C         uint16_t fit_height;
//C         uint16_t fit_width;
//C         uint16_t horizontal_dpi;
//C         uint16_t hlink_count;
//C         uint16_t page_start;
//C         uint16_t print_scale;
//C         uint16_t rel_count;
//C         uint16_t vertical_dpi;
//C         uint8_t filter_on;
//C         uint8_t fit_page;
//C         uint8_t hcenter;
//C         uint8_t orientation;
//C         uint8_t outline_changed;
//C         uint8_t page_order;
//C         uint8_t page_setup_changed;
//C         uint8_t page_view;
//C         uint8_t paper_size;
//C         uint8_t print_gridlines;
//C         uint8_t print_headers;
//C         uint8_t print_options_changed;
//C         uint8_t screen_gridlines;
//C         uint8_t tab_color;
//C         uint8_t vba_codename;
//C         uint8_t vcenter;

//C         double margin_left;
//C         double margin_right;
//C         double margin_top;
//C         double margin_bottom;
//C         double margin_header;
//C         double margin_footer;

//C         uint8_t header_footer_changed;
//C         char header[LXW_HEADER_FOOTER_MAX];
//C         char footer[LXW_HEADER_FOOTER_MAX];

//C         struct lxw_repeat_rows repeat_rows;
//C         struct lxw_repeat_cols repeat_cols;
//C         struct lxw_print_area print_area;
//C         struct lxw_autofilter autofilter;

//C         uint16_t merged_range_count;

//C         lxw_row_t *hbreaks;
//C         lxw_col_t *vbreaks;

//C         struct lxw_rel_tuples *external_hyperlinks;

//C         struct {
//C             struct lxw_worksheet *stqe_next; /* next element */
//C         } list_pointers;
struct _N8
{
    lxw_worksheet *stqe_next;
}

struct lxw_rel_tuples {
    lxw_rel_tuple *stqh_first;
    lxw_rel_tuple **stqh_last;
}

struct _N13
{
    lxw_rel_tuple *stqe_next; /* next element */
}
struct lxw_rel_tuple {

    char *type;
    char *target;
    char *target_mode;

    _N13 list_pointers;

};

//C     } lxw_worksheet;
struct lxw_worksheet
{
    FILE *file;
    FILE *optimize_tmpfile;
    lxw_table_rows *table;
    lxw_table_rows *hyperlinks;
    lxw_cell **array;
    lxw_merged_ranges *merged_ranges;
    lxw_row_t dim_rowmin;
    lxw_row_t dim_rowmax;
    lxw_col_t dim_colmin;
    lxw_col_t dim_colmax;
    lxw_sst *sst;
    char *name;
    char *quoted_name;
    uint32_t index;
    uint8_t active;
    uint8_t selected;
    uint8_t hidden;
    uint32_t *active_sheet;
    lxw_col_options **col_options;
    uint16_t col_options_max;
    double *col_sizes;
    uint16_t col_sizes_max;
    lxw_format **col_formats;
    uint16_t col_formats_max;
    uint8_t col_size_changed;
    uint8_t optimize;
    lxw_row *optimize_row;
    uint16_t fit_height;
    uint16_t fit_width;
    uint16_t horizontal_dpi;
    uint16_t hlink_count;
    uint16_t page_start;
    uint16_t print_scale;
    uint16_t rel_count;
    uint16_t vertical_dpi;
    uint8_t filter_on;
    uint8_t fit_page;
    uint8_t hcenter;
    uint8_t orientation;
    uint8_t outline_changed;
    uint8_t page_order;
    uint8_t page_setup_changed;
    uint8_t page_view;
    uint8_t paper_size;
    uint8_t print_gridlines;
    uint8_t print_headers;
    uint8_t print_options_changed;
    uint8_t screen_gridlines;
    uint8_t tab_color;
    uint8_t vba_codename;
    uint8_t vcenter;
    double margin_left;
    double margin_right;
    double margin_top;
    double margin_bottom;
    double margin_header;
    double margin_footer;
    uint8_t header_footer_changed;
    char [255]header;
    char [255]footer;
    lxw_repeat_rows repeat_rows;
    lxw_repeat_cols repeat_cols;
    lxw_print_area print_area;
    lxw_autofilter autofilter;
    uint16_t merged_range_count;
    lxw_row_t *hbreaks;
    lxw_col_t *vbreaks;
    lxw_rel_tuples *external_hyperlinks;
    _N8 list_pointers;
}

/*
 * Worksheet initialisation data.
 */
//C     typedef struct lxw_worksheet_init_data {
//C         uint32_t index;
//C         uint8_t hidden;
//C         uint8_t optimize;
//C         uint32_t *active_sheet;
//C         lxw_sst *sst;
//C         char *name;
//C         char *quoted_name;

//C     } lxw_worksheet_init_data;
struct lxw_worksheet_init_data
{
    uint32_t index;
    uint8_t hidden;
    uint8_t optimize;
    uint32_t *active_sheet;
    lxw_sst *sst;
    char *name;
    char *quoted_name;
}

/* Struct to represent a worksheet row. */
//C     typedef struct lxw_row {
//C         lxw_row_t row_num;
//C         double height;
//C         lxw_format *format;
//C         uint8_t hidden;
//C         uint8_t level;
//C         uint8_t collapsed;
//C         uint8_t row_changed;
//C         uint8_t data_changed;
//C         struct lxw_table_cells *cells;

    /* List pointers for queue.h. */
//C         struct {
//C             struct lxw_row *tqe_next;  /* next element */
//C             struct lxw_row **tqe_prev; /* address of previous next element */
//C         } list_pointers;
struct _N9
{
    lxw_row *tqe_next;
    lxw_row **tqe_prev;
}
//C     } lxw_row;
struct lxw_row
{
    lxw_row_t row_num;
    double height;
    lxw_format *format;
    uint8_t hidden;
    uint8_t level;
    uint8_t collapsed;
    uint8_t row_changed;
    uint8_t data_changed;
    lxw_table_cells *cells;
    _N9 list_pointers;
}

/* Struct to represent a worksheet cell. */
//C     typedef struct lxw_cell {
//C         lxw_row_t row_num;
//C         lxw_col_t col_num;
//C         enum cell_types type;
//C         lxw_format *format;

//C         union {
//C             double number;
//C             int32_t string_id;
//C             char *string;
//C         } u;
union _N10
{
    double number;
    int32_t string_id;
    char *string;
}

//C         double formula_result;
//C         char *user_data1;
//C         char *user_data2;

    /* List pointers for queue.h. */
//C         struct {
//C             struct lxw_cell *tqe_next;  /* next element */
//C             struct lxw_cell **tqe_prev; /* address of previous next element */
//C         } list_pointers;
struct _N11
{
    lxw_cell *tqe_next;
    lxw_cell **tqe_prev;
}
//C     } lxw_cell;
struct lxw_cell
{
    lxw_row_t row_num;
    lxw_col_t col_num;
    cell_types type;
    lxw_format *format;
    _N10 u;
    double formula_result;
    char *user_data1;
    char *user_data2;
    _N11 list_pointers;
}

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/**
 * @brief Write a number to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param number    The number to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `worksheet_write_number()` function writes numeric types to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_number(worksheet, 0, 0, 123456, NULL);
 *     worksheet_write_number(worksheet, 1, 0, 2.3451, NULL);
 * @endcode
 *
 * @image html write_number01.png
 *
 * The native data type for all numbers in Excel is a IEEE-754 64-bit
 * double-precision floating point, which is also the default type used by
 * `%worksheet_write_number`.
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 *     format_set_num_format(format, "$#,##0.00");
 *
 *     worksheet_write_number(worksheet, 0, 0, 1234.567, format);
 * @endcode
 *
 * @image html write_number02.png
 *
 */
//C     int8_t worksheet_write_number(lxw_worksheet *worksheet,
//C                                   lxw_row_t row,
//C                                   lxw_col_t col, double number,
//C                                   lxw_format *format);
int8_t  worksheet_write_number(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, double number, lxw_format *format);
/**
 * @brief Write a string to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param string    String to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `%worksheet_write_string()` function writes a string to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is English!", NULL);
 * @endcode
 *
 * @image html write_string01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object:
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 *     format_set_bold(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is Bold!", format);
 * @endcode
 *
 * @image html write_string02.png
 *
 * Unicode strings are supported in UTF-8 encoding. This generally requires
 * that your source file is UTF-8 encoded or that the data has been read from
 * a UTF-8 source:
 *
 * @code
 *    worksheet_write_string(worksheet, 0, 0, "Это фраза на русском!", NULL);
 * @endcode
 *
 * @image html write_string03.png
 *
 */
//C     int8_t worksheet_write_string(lxw_worksheet *worksheet,
//C                                   lxw_row_t row,
//C                                   lxw_col_t col, const char *string,
//C                                   lxw_format *format);
int8_t  worksheet_write_string(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, char *string, lxw_format *format);
/**
 * @brief Write a formula to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `%worksheet_write_formula()` function writes a formula or function to
 * the cell specified by `row` and `column`:
 *
 * @code
 *  worksheet_write_formula(worksheet, 0, 0, "=B3 + 6",                    NULL);
 *  worksheet_write_formula(worksheet, 1, 0, "=SIN(PI()/4)",               NULL);
 *  worksheet_write_formula(worksheet, 2, 0, "=SUM(A1:A2)",                NULL);
 *  worksheet_write_formula(worksheet, 3, 0, "=IF(A3>1,\"Yes\", \"No\")",  NULL);
 *  worksheet_write_formula(worksheet, 4, 0, "=AVERAGE(1, 2, 3, 4)",       NULL);
 *  worksheet_write_formula(worksheet, 5, 0, "=DATEVALUE(\"1-Jan-2013\")", NULL);
 * @endcode
 *
 * @image html write_formula01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores a
 * default value of `0`. The correct formula result is displayed in Excel, as
 * shown in the example above, since it recalculates the formulas when it loads
 * the file. For cases where this is an issue see the
 * `worksheet_write_formula_num()` function and the discussion in that section.
 *
 * Formulas must be written with the US style separator/range operator which
 * is a comma (not semi-colon). Therefore a formula with multiple values
 * should be written as follows:
 *
 * @code
 *     // OK.
 *     worksheet_write_formula(worksheet, 0, 0, "=SUM(1, 2, 3)", NULL);
 *
 *     // NO. Error on load.
 *     worksheet_write_formula(worksheet, 1, 0, "=SUM(1; 2; 3)", NULL);
 * @endcode
 *
 */
//C     int8_t worksheet_write_formula(lxw_worksheet *worksheet,
//C                                    lxw_row_t row,
//C                                    lxw_col_t col, const char *formula,
//C                                    lxw_format *format);
int8_t  worksheet_write_formula(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, char *formula, lxw_format *format);
/**
 * @brief Write an array formula to a worksheet cell.
 *
 * @param worksheet
 * @param first_row   The first row of the range. (All zero indexed.)
 * @param first_col   The first column of the range.
 * @param last_row    The last row of the range.
 * @param last_col    The last col of the range.
 * @param formula     Array formula to write to cell.
 * @param format      A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
  * The `%worksheet_write_array_formula()` function writes an array formula to
 * a cell range. In Excel an array formula is a formula that performs a
 * calculation on a set of values.
 *
 * In Excel an array formula is indicated by a pair of braces around the
 * formula: `{=SUM(A1:B1*A2:B2)}`.
 *
 * Array formulas can return a single value or a range or values. For array
 * formulas that return a range of values you must specify the range that the
 * return values will be written to. This is why this function has `first_`
 * and `last_` row/column parameters. The RANGE() macro can also be used to
 * specify the range:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 4, 0, 6, 0,     "{=TREND(C5:C7,B5:B7)}", NULL);
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_write_array_formula(worksheet, RANGE("A5:A7"), "{=TREND(C5:C7,B5:B7)}", NULL);
 * @endcode
 *
 * If the array formula returns a single value then the `first_` and `last_`
 * parameters should be the same:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 1, 0, 1, 0,     "{=SUM(B1:C1*B2:C2)}", NULL);
 *     worksheet_write_array_formula(worksheet, RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}", NULL);
 * @endcode
 *
 */
//C     int8_t worksheet_write_array_formula(lxw_worksheet *worksheet,
//C                                          lxw_row_t first_row,
//C                                          lxw_col_t first_col,
//C                                          lxw_row_t last_row,
//C                                          lxw_col_t last_col,
//C                                          const char *formula, lxw_format *format);
int8_t  worksheet_write_array_formula(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, char *formula, lxw_format *format);

//C     int8_t worksheet_write_array_formula_num(lxw_worksheet *worksheet,
//C                                              lxw_row_t first_row,
//C                                              lxw_col_t first_col,
//C                                              lxw_row_t last_row,
//C                                              lxw_col_t last_col,
//C                                              const char *formula,
//C                                              lxw_format *format, double result);
int8_t  worksheet_write_array_formula_num(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, char *formula, lxw_format *format, double result);

/**
 * @brief Write a date or time to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param datetime  The datetime to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `worksheet_write_datetime()` function can be used to write a date or
 * time to the cell specified by `row` and `column`:
 *
 * @dontinclude dates_and_times02.c
 * @skip include
 * @until num_format
 * @skip Feb
 * @until }
 *
 * The `format` parameter should be used to apply formatting to the cell using
 * a @ref format.h "Format" object as shown above. Without a date format the
 * datetime will appear as a number only.
 *
 * See @ref working_with_dates for more information about handling dates and
 * times in libxlsxwriter.
 */
//C     int8_t worksheet_write_datetime(lxw_worksheet *worksheet,
//C                                     lxw_row_t row,
//C                                     lxw_col_t col, lxw_datetime *datetime,
//C                                     lxw_format *format);
int8_t  worksheet_write_datetime(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_datetime *datetime, lxw_format *format);

//C     int8_t worksheet_write_url_opt(lxw_worksheet *worksheet,
//C                                    lxw_row_t row_num,
//C                                    lxw_col_t col_num, const char *url,
//C                                    lxw_format *format, const char *string,
//C                                    const char *tooltip);
int8_t  worksheet_write_url_opt(lxw_worksheet *worksheet, lxw_row_t row_num, lxw_col_t col_num, char *url, lxw_format *format, char *string, char *tooltip);
/**
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param url       The url to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 *
 * The `%worksheet_write_url()` function is used to write a URL/hyperlink to a
 * worksheet cell specified by `row` and `column`.
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "http://libxlsxwriter.github.io", url_format);
 * @endcode
 *
 * @image html hyperlinks_short.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a @ref
 * format.h "Format" object. The typical worksheet format for a hyperlink is a
 * blue underline:
 *
 * @code
 *    lxw_format *url_format   = workbook_add_format(workbook);
 *
 *    format_set_underline (url_format, LXW_UNDERLINE_SINGLE);
 *    format_set_font_color(url_format, LXW_COLOR_BLUE);
 *
 * @endcode
 *
 * The usual web style URI's are supported: `%http://`, `%https://`, `%ftp://`
 * and `mailto:` :
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "ftp://www.python.org/",    url_format);
 *     worksheet_write_url(worksheet, 1, 0, "http://www.python.org/",   url_format);
 *     worksheet_write_url(worksheet, 2, 0, "https://www.python.org/",  url_format);
 *     worksheet_write_url(worksheet, 3, 0, "mailto:jmcnamaracpan.org", url_format);
 *
 * @endcode
 *
 * An Excel hyperlink is comprised of two elements: the displayed string and
 * the non-displayed link. By default the displayed string is the same as the
 * link. However, it is possible to overwrite it with any other
 * `libxlsxwriter` type using the appropriate `worksheet_write_*()`
 * function. The most common case is to overwrite the displayed link text with
 * another string:
 *
 * @code
 *  // Write a hyperlink but overwrite the displayed string.
 *  worksheet_write_url   (worksheet, 2, 0, "http://libxlsxwriter.github.io", url_format);
 *  worksheet_write_string(worksheet, 2, 0, "Read the documentation.",        url_format);
 *
 * @endcode
 *
 * @image html hyperlinks_short2.png
 *
 * Two local URIs are supported: `internal:` and `external:`. These are used
 * for hyperlinks to internal worksheet references or external workbook and
 * worksheet references:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:Sheet2!A1",                url_format);
 *     worksheet_write_url(worksheet, 1, 0, "internal:Sheet2!B2",                url_format);
 *     worksheet_write_url(worksheet, 2, 0, "internal:Sheet2!A1:B2",             url_format);
 *     worksheet_write_url(worksheet, 3, 0, "internal:'Sales Data'!A1",          url_format);
 *     worksheet_write_url(worksheet, 4, 0, "external:c:\\temp\\foo.xlsx",       url_format);
 *     worksheet_write_url(worksheet, 5, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
 *     worksheet_write_url(worksheet, 6, 0, "external:..\\foo.xlsx",             url_format);
 *     worksheet_write_url(worksheet, 7, 0, "external:..\\foo.xlsx#Sheet2!A1",   url_format);
 *     worksheet_write_url(worksheet, 8, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
 *
 * @endcode
 *
 * Worksheet references are typically of the form `Sheet1!A1`. You can also
 * link to a worksheet range using the standard Excel notation:
 * `Sheet1!A1:B2`.
 *
 * In external links the workbook and worksheet name must be separated by the
 * `#` character:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
 * @endcode
 *
 * You can also link to a named range in the target worksheet: For example say
 * you have a named range called `my_name` in the workbook `c:\temp\foo.xlsx`
 * you could link to it as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\temp\\foo.xlsx#my_name", url_format);
 *
 * @endcode
 *
 * Excel requires that worksheet names containing spaces or non alphanumeric
 * characters are single quoted as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:'Sales Data'!A1", url_format);
 * @endcode
 *
 * Links to network files are also supported. Network files normally begin
 * with two back slashes as follows `\\NETWORK\etc`. In order to represent
 * this in a C string literal the backslashes should be escaped:
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
 * @endcode
 *
 *
 * Alternatively, you can use Windows style forward slashes. These are
 * translated internally to backslashes:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:/temp/foo.xlsx",     url_format);
 *     worksheet_write_url(worksheet, 1, 0, "external://NET/share/foo.xlsx", url_format);
 *
 * @endcode
 *
 *
 * **Note:**
 *
 *    libxlsxwriter will escape the following characters in URLs as required
 *    by Excel: `\s " < > \ [ ]  ^ { }` unless the URL already contains `%%xx`
 *    style escapes. In which case it is assumed that the URL was escaped
 *    correctly by the user and will by passed directly to Excel.
 *
 */
//C     int8_t worksheet_write_url(lxw_worksheet *worksheet,
//C                                lxw_row_t row,
//C                                lxw_col_t col, const char *url,
//C                                lxw_format *format);
int8_t  worksheet_write_url(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, char *url, lxw_format *format);

/**
 * @brief Write a formatted blank worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * Write a blank cell specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_blank(worksheet, 1, 1, border_format);
 * @endcode
 *
 * This function is used to add formatting to a cell which doesn't contain a
 * string or number value.
 *
 * Excel differentiates between an "Empty" cell and a "Blank" cell. An Empty
 * cell is a cell which doesn't contain data or formatting whilst a Blank cell
 * doesn't contain data but does contain formatting. Excel stores Blank cells
 * but ignores Empty cells.
 *
 * As such, if you write an empty cell without formatting it is ignored.
 *
 */
//C     int8_t worksheet_write_blank(lxw_worksheet *worksheet,
//C                                  lxw_row_t row, lxw_col_t col,
//C                                  lxw_format *format);
int8_t  worksheet_write_blank(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, lxw_format *format);

/**
 * @brief Write a formula to a worksheet cell with a user defined result.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 * @param result    A user defined result for a formula.
 *
 * @return A #lxw_write_error code.
 *
 * The `%worksheet_write_formula_num()` function writes a formula or Excel
 * function to the cell specified by `row` and `column` with a user defined
 * result:
 *
 * @code
 *     // Required as a workaround only.
 *     worksheet_write_formula_num(worksheet, 0, 0, "=1 + 2", NULL, 3);
 * @endcode
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores
 * the value `0` as the formula result. It then sets a global flag in the XLSX
 * file to say that all formulas and functions should be recalculated when the
 * file is opened.
 *
 * This is the method recommended in the Excel documentation and in general it
 * works fine with spreadsheet applications.
 *
 * However, applications that don't have a facility to calculate formulas,
 * such as Excel Viewer, or some mobile applications will only display the `0`
 * results.
 *
 * If required, the `%worksheet_write_formula_num()` function can be used to
 * specify a formula and its result.
 *
 * This function is rarely required and is only provided for compatibility
 * with some third party applications. For most applications the
 * worksheet_write_formula() function is the recommended way of writing
 * formulas.
 *
 */
//C     int8_t worksheet_write_formula_num(lxw_worksheet *worksheet,
//C                                        lxw_row_t row,
//C                                        lxw_col_t col,
//C                                        const char *formula,
//C                                        lxw_format *format, double result);
int8_t  worksheet_write_formula_num(lxw_worksheet *worksheet, lxw_row_t row, lxw_col_t col, char *formula, lxw_format *format, double result);

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height.
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * The `%worksheet_set_row()` function is used to change the default
 * properties of a row. The most common use for this function is to change the
 * height of a row:
 *
 * @code
 *     // Set the height of Row 1 to 20.
 *     worksheet_set_row(worksheet, 0, 20, NULL, NULL);
 * @endcode
 *
 * The other common use for `%worksheet_set_row()` is to set the a @ref
 * format.h "Format" for all cells in the row:
 *
 * @code
 *     lxw_format *bold = workbook_add_format(workbook);
 *     format_set_bold(bold);
 *
 *     // Set the header row to bold.
 *     worksheet_set_row(worksheet, 0, 15, bold, NULL);
 * @endcode
 *
 * If you wish to set the format of a row without changing the height you can
 * pass the default row height of #LXW_DEF_ROW_HEIGHT = 15:
 *
 * @code
 *     worksheet_set_row(worksheet, 0, LXW_DEF_ROW_HEIGHT, format, NULL);
 *     worksheet_set_row(worksheet, 0, 15, format, NULL); // Same as above.
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the row that don't
 * have a format. As with Excel the row format is overridden by an explicit
 * cell format. For example:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1, NULL);
 *
 *     // Cell A1 in Row 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell B1 in Row 1 keeps format2.
 *     worksheet_write_string(worksheet, 0, 1, "Hello", format2);
 * @endcode
 *
 * The `options` parameter is a #lxw_row_col_options struct. It has the
 * following members but currently only the `hidden` property is supported:
 *
 * - `hidden`
 * - `level`
 * - `collapsed`
 *
 * The `"hidden"` option is used to hide a row. This can be used, for
 * example, to hide intermediary steps in a complicated calculation:
 *
 * @code
 *     lxw_row_col_options options = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     // Hide the fourth row.
 *     worksheet_set_row(worksheet, 3, 20, NULL, &options);
 * @endcode
 *
 */
//C     int8_t worksheet_set_row(lxw_worksheet *worksheet,
//C                              lxw_row_t row,
//C                              double height,
//C                              lxw_format *format, lxw_row_col_options *options);
int8_t  worksheet_set_row(lxw_worksheet *worksheet, lxw_row_t row, double height, lxw_format *format, lxw_row_col_options *options);

/**
 * @brief Set the properties for one or more columns of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param width     The width of the column(s).
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * The `%worksheet_set_column()` function can be used to change the default
 * properties of a single column or a range of columns:
 *
 * @code
 *     // Width of columns B:D set to 30.
 *     worksheet_set_column(worksheet, 1, 3, 30, NULL, NULL);
 *
 * @endcode
 *
 * If `%worksheet_set_column()` is applied to a single column the value of
 * `first_col` and `last_col` should be the same:
 *
 * @code
 *     // Width of column B set to 30.
 *     worksheet_set_column(worksheet, 1, 1, 30, NULL, NULL);
 *
 * @endcode
 *
 * It is also possible, and generally clearer, to specify a column range using
 * the form of `COLS()` macro:
 *
 * @code
 *     worksheet_set_column(worksheet, 4, 4, 20, NULL, NULL);
 *     worksheet_set_column(worksheet, 5, 8, 30, NULL, NULL);
 *
 *     // Same as the examples above but clearer.
 *     worksheet_set_column(worksheet, COLS("E:E"), 20, NULL, NULL);
 *     worksheet_set_column(worksheet, COLS("F:H"), 30, NULL, NULL);
 *
 * @endcode
 *
 * The width corresponds to the column width value that is specified in
 * Excel. It is approximately equal to the length of a string in the default
 * font of Calibri 11. Unfortunately, there is no way to specify "AutoFit" for
 * a column in the Excel file format. This feature is only available at
 * runtime from within Excel. It is possible to simulate "AutoFit" by tracking
 * the width of the data in the column as your write it.
 *
 * As usual the @ref format.h `format` parameter is optional. If you wish to
 * set the format without changing the width you can pass default col width of
 * #LXW_DEF_COL_WIDTH = 8.43:
 *
 * @code
 *     lxw_format *bold = workbook_add_format(workbook);
 *     format_set_bold(bold);
 *
 *     // Set the first column to bold.
 *     worksheet_set_column(worksheet, 0, 0, LXW_DEF_COL_HEIGHT, bold, NULL);
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the column that
 * don't have a format. For example:
 *
 * @code
 *     // Column 1 has format1.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format1, NULL);
 *
 *     // Cell A1 in column 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell A2 in column 1 keeps format2.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", format2);
 * @endcode
 *
 * As in Excel a row format takes precedence over a default column format:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1, NULL);
 *
 *     // Col 1 has format2.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format2, NULL);
 *
 *     // Cell A1 defaults to format1, the row format.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *    // Cell A2 keeps format2, the column format.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", NULL);
 * @endcode
 *
 * The `options` parameter is a #lxw_row_col_options struct. It has the
 * following members but currently only the `hidden` property is supported:
 *
 * - `hidden`
 * - `level`
 * - `collapsed`
 *
 * The `"hidden"` option is used to hide a column. This can be used, for
 * example, to hide intermediary steps in a complicated calculation:
 *
 * @code
 *     lxw_row_col_options options = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, NULL, &options);
 * @endcode
 *
 */
//C     int8_t worksheet_set_column(lxw_worksheet *worksheet, lxw_col_t first_col,
//C                                 lxw_col_t last_col, double width,
//C                                 lxw_format *format, lxw_row_col_options *options);
int8_t  worksheet_set_column(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col, double width, lxw_format *format, lxw_row_col_options *options);

/**
 * @brief Merge a range of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 * @param string    String to write to the merged range.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return 0 for success, non-zero on error.
 *
 * The `%worksheet_merge_range()` function allows cells to be merged together
 * so that they act as a single area.
 *
 * Excel generally merges and centers cells at same time. To get similar
 * behaviour with libxlsxwriter you need to apply a @ref format.h "Format"
 * object with the appropriate alignment:
 *
 * @code
 *     lxw_format *merge_format = workbook_add_format(workbook);
 *     format_set_align(merge_format, LXW_ALIGN_CENTER);
 *
 *     worksheet_merge_range(worksheet, 1, 1, 1, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * It is possible to apply other formatting to the merged cells as well:
 *
 * @code
 *    format_set_align   (merge_format, LXW_ALIGN_CENTER);
 *    format_set_align   (merge_format, LXW_ALIGN_VERTICAL_CENTER);
 *    format_set_border  (merge_format, LXW_BORDER_DOUBLE);
 *    format_set_bold    (merge_format);
 *    format_set_bg_color(merge_format, 0xD7E4BC);
 *
 *    worksheet_merge_range(worksheet, 2, 1, 3, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * @image html merge.png
 *
 * The `%worksheet_merge_range()` function writes a `char*` string using
 * `worksheet_write_string()`. In order to write other data types, such as a
 * number or a formula, you can overwrite the first cell with a call to one of
 * the other write functions. The same Format should be used as was used in
 * the merged range.
 *
 * @code
 *    // First write a range with a blank string.
 *    worksheet_merge_range (worksheet, 1, 1, 1, 3, "", format);
 *
 *    // Then overwrite the first cell with a number.
 *    worksheet_write_number(worksheet, 1, 1, 123, format);
 * @endcode
 */
//C     uint8_t worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                   lxw_col_t first_col, lxw_row_t last_row,
//C                                   lxw_col_t last_col, const char *string,
//C                                   lxw_format *format);
uint8_t  worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col, char *string, lxw_format *format);

/**
 * @brief Set the autofilter area in the worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return 0 for success, non-zero on error.
 *
 * The `%worksheet_autofilter()` method allows an autofilter to be added to a
 * worksheet.
 *
 * An autofilter is a way of adding drop down lists to the headers of a 2D
 * range of worksheet data. This allows users to filter the data based on
 * simple criteria so that some data is shown and some is hidden.
 *
 * @image html autofilter.png
 *
 * To add an autofilter to a worksheet:
 *
 * @code
 *     worksheet_autofilter(worksheet, 0, 0, 50, 3);
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_autofilter(worksheet, RANGE("A1:D51"));
 * @endcode
 *
 * Note: it isn't currently possible to apply filter conditions to the
 * autofilter.
 */
//C     uint8_t worksheet_autofilter(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                  lxw_col_t first_col, lxw_row_t last_row,
//C                                  lxw_col_t last_col);
uint8_t  worksheet_autofilter(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);

 /**
  * @brief Make a worksheet the active, i.e., visible worksheet.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_activate()` function is used to specify which worksheet is
  * initially visible in a multi-sheet workbook:
  *
  * @code
  *     lxw_worksheet *worksheet1 = workbook_add_worksheet(workbook, NULL);
  *     lxw_worksheet *worksheet2 = workbook_add_worksheet(workbook, NULL);
  *     lxw_worksheet *worksheet3 = workbook_add_worksheet(workbook, NULL);
  *
  *     worksheet_activate(worksheet3);
  * @endcode
  *
  * @image html worksheet_activate.png
  *
  * More than one worksheet can be selected via the `worksheet_select()`
  * function, see below, however only one worksheet can be active.
  *
  * The default active worksheet is the first worksheet.
  *
  */
//C     void worksheet_activate(lxw_worksheet *worksheet);
void  worksheet_activate(lxw_worksheet *worksheet);

 /**
  * @brief Set a worksheet tab as selected.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_select()` function is used to indicate that a worksheet is
  * selected in a multi-sheet workbook:
  *
  * @code
  *     worksheet_activate(worksheet1);
  *     worksheet_select(worksheet2);
  *     worksheet_select(worksheet3);
  *
  * @endcode
  *
  * A selected worksheet has its tab highlighted. Selecting worksheets is a
  * way of grouping them together so that, for example, several worksheets
  * could be printed in one go. A worksheet that has been activated via the
  * `worksheet_activate()` function will also appear as selected.
  *
  */
//C     void worksheet_select(lxw_worksheet *worksheet);
void  worksheet_select(lxw_worksheet *worksheet);

/**
 * @brief Set the page orientation as landscape.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to landscape:
 *
 * @code
 *     worksheet_set_landscape(worksheet);
 * @endcode
 */
//C     void worksheet_set_landscape(lxw_worksheet *worksheet);
void  worksheet_set_landscape(lxw_worksheet *worksheet);

/**
 * @brief Set the page orientation as portrait.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to portrait. The default worksheet orientation is portrait, so this
 * function isn't generally required:
 *
 * @code
 *     worksheet_set_portrait(worksheet);
 * @endcode
 */
//C     void worksheet_set_portrait(lxw_worksheet *worksheet);
void  worksheet_set_portrait(lxw_worksheet *worksheet);

/**
 * @brief Set the page layout to page view mode.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to display the worksheet in "Page View/Layout" mode:
 *
 * @code
 *     worksheet_set_page_view(worksheet);
 * @endcode
 */
//C     void worksheet_set_page_view(lxw_worksheet *worksheet);
void  worksheet_set_page_view(lxw_worksheet *worksheet);

/**
 * @brief Set the paper type for printing.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param paper_type The Excel paper format type.
 *
 * This function is used to set the paper format for the printed output of a
 * worksheet. The following paper styles are available:
 *
 *
 *   Index    | Paper format            | Paper size
 *   :------- | :---------------------- | :-------------------
 *   0        | Printer default         | Printer default
 *   1        | Letter                  | 8 1/2 x 11 in
 *   2        | Letter Small            | 8 1/2 x 11 in
 *   3        | Tabloid                 | 11 x 17 in
 *   4        | Ledger                  | 17 x 11 in
 *   5        | Legal                   | 8 1/2 x 14 in
 *   6        | Statement               | 5 1/2 x 8 1/2 in
 *   7        | Executive               | 7 1/4 x 10 1/2 in
 *   8        | A3                      | 297 x 420 mm
 *   9        | A4                      | 210 x 297 mm
 *   10       | A4 Small                | 210 x 297 mm
 *   11       | A5                      | 148 x 210 mm
 *   12       | B4                      | 250 x 354 mm
 *   13       | B5                      | 182 x 257 mm
 *   14       | Folio                   | 8 1/2 x 13 in
 *   15       | Quarto                  | 215 x 275 mm
 *   16       | ---                     | 10x14 in
 *   17       | ---                     | 11x17 in
 *   18       | Note                    | 8 1/2 x 11 in
 *   19       | Envelope 9              | 3 7/8 x 8 7/8
 *   20       | Envelope 10             | 4 1/8 x 9 1/2
 *   21       | Envelope 11             | 4 1/2 x 10 3/8
 *   22       | Envelope 12             | 4 3/4 x 11
 *   23       | Envelope 14             | 5 x 11 1/2
 *   24       | C size sheet            | ---
 *   25       | D size sheet            | ---
 *   26       | E size sheet            | ---
 *   27       | Envelope DL             | 110 x 220 mm
 *   28       | Envelope C3             | 324 x 458 mm
 *   29       | Envelope C4             | 229 x 324 mm
 *   30       | Envelope C5             | 162 x 229 mm
 *   31       | Envelope C6             | 114 x 162 mm
 *   32       | Envelope C65            | 114 x 229 mm
 *   33       | Envelope B4             | 250 x 353 mm
 *   34       | Envelope B5             | 176 x 250 mm
 *   35       | Envelope B6             | 176 x 125 mm
 *   36       | Envelope                | 110 x 230 mm
 *   37       | Monarch                 | 3.875 x 7.5 in
 *   38       | Envelope                | 3 5/8 x 6 1/2 in
 *   39       | Fanfold                 | 14 7/8 x 11 in
 *   40       | German Std Fanfold      | 8 1/2 x 12 in
 *   41       | German Legal Fanfold    | 8 1/2 x 13 in
 *
 * Note, it is likely that not all of these paper types will be available to
 * the end user since it will depend on the paper formats that the user's
 * printer supports. Therefore, it is best to stick to standard paper types:
 *
 * @code
 *     worksheet_set_paper(worksheet1, 1);  // US Letter
 *     worksheet_set_paper(worksheet2, 9);  // A4
 * @endcode
 *
 * If you do not specify a paper type the worksheet will print using the
 * printer's default paper style.
 */
//C     void worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);
void  worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);

/**
 * @brief Set the worksheet margins for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param left    Left margin in inches.   Excel default is 0.7.
 * @param right   Right margin in inches.  Excel default is 0.7.
 * @param top     Top margin in inches.    Excel default is 0.75.
 * @param bottom  Bottom margin in inches. Excel default is 0.75.
 *
 * The `%worksheet_set_margins()` function is used to set the margins of the
 * worksheet when it is printed. The units are in inches. Specifying `-1` for
 * any parameter will give the default Excel value as shown above.
 *
 * @code
 *    worksheet_set_margins(worksheet, 1.3, 1.2, -1, -1);
 * @endcode
 *
 */
//C     void worksheet_set_margins(lxw_worksheet *worksheet, double left,
//C                                double right, double top, double bottom);
void  worksheet_set_margins(lxw_worksheet *worksheet, double left, double right, double top, double bottom);

/**
 * @brief Set the printed page header caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 *
 * @return 0 for success, non-zero on error.
 *
 * Headers and footers are generated using a string which is a combination of
 * plain text and control characters.
 *
 * The available control character are:
 *
 *
 *   | Control         | Category      | Description           |
 *   | --------------- | ------------- | --------------------- |
 *   | `&L`            | Justification | Left                  |
 *   | `&C`            |               | Center                |
 *   | `&R`            |               | Right                 |
 *   | `&P`            | Information   | Page number           |
 *   | `&N`            |               | Total number of pages |
 *   | `&D`            |               | Date                  |
 *   | `&T`            |               | Time                  |
 *   | `&F`            |               | File name             |
 *   | `&A`            |               | Worksheet name        |
 *   | `&Z`            |               | Workbook path         |
 *   | `&fontsize`     | Font          | Font size             |
 *   | `&"font,style"` |               | Font name and style   |
 *   | `&U`            |               | Single underline      |
 *   | `&E`            |               | Double underline      |
 *   | `&S`            |               | Strikethrough         |
 *   | `&X`            |               | Superscript           |
 *   | `&Y`            |               | Subscript             |
 *
 *
 * Text in headers and footers can be justified (aligned) to the left, center
 * and right by prefixing the text with the control characters `&L`, `&C` and
 * `&R`.
 *
 * For example (with ASCII art representation of the results):
 *
 * @code
 *     worksheet_set_header(worksheet, "&LHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Hello                                                         |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&CHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&RHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                                                         Hello |
 *     |                                                               |
 *
 *
 * @endcode
 *
 * For simple text, if you do not specify any justification the text will be
 * centred. However, you must prefix the text with `&C` if you specify a font
 * name or any other formatting:
 *
 * @code
 *     worksheet_set_header(worksheet, "Hello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 * @endcode
 *
 * You can have text in each of the justification regions:
 *
 * @code
 *     worksheet_set_header(worksheet, "&LCiao&CBello&RCielo");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Ciao                     Bello                          Cielo |
 *     |                                                               |
 *
 * @endcode
 *
 * The information control characters act as variables that Excel will update
 * as the workbook or worksheet changes. Times and dates are in the users
 * default format:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CPage &P of &N");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                        Page 1 of 6                            |
 *     |                                                               |
 *
 *     worksheet_set_header(worksheet, "&CUpdated at &T");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                    Updated at 12:30 PM                        |
 *     |                                                               |
 *
 * @endcode
 *
 * You can specify the font size of a section of the text by prefixing it with
 * the control character `&n` where `n` is the font size:
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&30Hello Big");
 *     worksheet_set_header(worksheet2, "&C&10Hello Small");
 *
 * @endcode
 *
 * You can specify the font of a section of the text by prefixing it with the
 * control sequence `&"font,style"` where `fontname` is a font name such as
 * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
 * "Courier New" or "Times New Roman" and `style` is one of the standard
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&\"Courier New,Italic\"Hello");
 *     worksheet_set_header(worksheet2, "&C&\"Courier New,Bold Italic\"Hello");
 *     worksheet_set_header(worksheet3, "&C&\"Times New Roman,Regular\"Hello");
 *
 * @endcode
 *
 * It is possible to combine all of these features together to create
 * sophisticated headers and footers. As an aid to setting up complicated
 * headers and footers you can record a page set-up as a macro in Excel and
 * look at the format strings that VBA produces. Remember however that VBA
 * uses two double quotes `""` to indicate a single double quote. For the last
 * example above the equivalent VBA code looks like this:
 *
 * @code
 *     .LeftHeader = ""
 *     .CenterHeader = "&""Times New Roman,Regular""Hello"
 *     .RightHeader = ""
 *
 * @endcode
 *
 * Alternatively you can inspect the header and footer strings in an Excel
 * file by unzipping it and grepping the XML sub-files. The following shows
 * how to do that using libxml's xmllint to format the XML for clarity:
 *
 * @code
 *
 *    $ unzip myfile.xlsm -d myfile
 *    $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer"
 *
 *      <headerFooter scaleWithDoc="0">
 *        <oddHeader>&amp;L&amp;P</oddHeader>
 *      </headerFooter>
 *
 * @endcode
 *
 * Note that in this case you need to unescape the Html. In the above example
 * the header string would be `&L&P`.
 *
 * To include a single literal ampersand `&` in a header or footer you should
 * use a double ampersand `&&`:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CCuriouser && Curiouser - Attorneys at Law");
 * @endcode
 *
 * Note, the header or footer string must be less than 255 characters. Strings
 * longer than this will not be written.
 *
 */
//C     uint8_t worksheet_set_header(lxw_worksheet *worksheet, char *string);
uint8_t  worksheet_set_header(lxw_worksheet *worksheet, char *string);

/**
 * @brief Set the printed page footer caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 *
 * @return 0 for success, non-zero on error.
 *
 * The syntax of this function is the same as worksheet_set_header().
 *
 */
//C     uint8_t worksheet_set_footer(lxw_worksheet *worksheet, char *string);
uint8_t  worksheet_set_footer(lxw_worksheet *worksheet, char *string);

/**
 * @brief Set the printed page header caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 * @param options   Header options.
 *
 * @return 0 for success, non-zero on error.
 *
 * The syntax of this function is the same as worksheet_set_header() with an
 * additional parameter to specify options for the header.
 *
 * Currently, the only available option is the header margin:
 *
 * @code
 *
 *    lxw_header_footer_options header_options = { 0.2 };
 *
 *    worksheet_set_header_opt(worksheet, "Some text", &header_options);
 *
 * @endcode
 *
 */
//C     uint8_t worksheet_set_header_opt(lxw_worksheet *worksheet, char *string,
//C                                      lxw_header_footer_options *options);
uint8_t  worksheet_set_header_opt(lxw_worksheet *worksheet, char *string, lxw_header_footer_options *options);

/**
 * @brief Set the printed page footer caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 * @param options   Footer options.
 *
 * @return 0 for success, non-zero on error.
 *
 * The syntax of this function is the same as worksheet_set_header_opt().
 *
 */
//C     uint8_t worksheet_set_footer_opt(lxw_worksheet *worksheet, char *string,
//C                                      lxw_header_footer_options *options);
uint8_t  worksheet_set_footer_opt(lxw_worksheet *worksheet, char *string, lxw_header_footer_options *options);

/**
 * @brief Set the horizontal page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * The `%worksheet_set_h_pagebreaks()` function adds horizontal page breaks to
 * a worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Horizontal page breaks act between rows.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxw_row_t and the last element of the array must be 0:
 *
 * @code
 *    lxw_row_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxw_row_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet1, breaks1);
 *    worksheet_set_h_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between rows 20 and 21 you must specify the break at
 * row 21. However in zero index notation this is actually row 20:
 *
 * @code
 *    // Break between row 20 and 21.
 *    lxw_row_t breaks[] = {20, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 horizontal page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
//C     void worksheet_set_h_pagebreaks(lxw_worksheet *worksheet, lxw_row_t breaks[]);
void  worksheet_set_h_pagebreaks(lxw_worksheet *worksheet, lxw_row_t *breaks);

/**
 * @brief Set the vertical page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * The `%worksheet_set_v_pagebreaks()` function adds vertical page breaks to a
 * worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Vertical page breaks act between columns.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxw_col_t and the last element of the array must be 0:
 *
 * @code
 *    lxw_col_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxw_col_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet1, breaks1);
 *    worksheet_set_v_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between columns 20 and 21 you must specify the break
 * at column 21. However in zero index notation this is actually column 20:
 *
 * @code
 *    // Break between column 20 and 21.
 *    lxw_col_t breaks[] = {20, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 vertical page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
//C     void worksheet_set_v_pagebreaks(lxw_worksheet *worksheet, lxw_col_t breaks[]);
void  worksheet_set_v_pagebreaks(lxw_worksheet *worksheet, lxw_col_t *breaks);

/**
 * @brief Set the order in which pages are printed.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_print_across()` function is used to change the default
 * print direction. This is referred to by Excel as the sheet "page order":
 *
 * @code
 *     worksheet_print_across(worksheet);
 * @endcode
 *
 * The default page order is shown below for a worksheet that extends over 4
 * pages. The order is called "down then across":
 *
 *     [1] [3]
 *     [2] [4]
 *
 * However, by using the `print_across` function the print order will be
 * changed to "across then down":
 *
 *     [1] [2]
 *     [3] [4]
 *
 */
//C     void worksheet_print_across(lxw_worksheet *worksheet);
void  worksheet_print_across(lxw_worksheet *worksheet);

/**
 * @brief Set the option to display or hide gridlines on the screen and
 *        the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param option    Gridline option.
 *
 * Display or hide screen and print gridlines using one of the values of
 * @ref lxw_gridlines.
 *
 * @code
 *    worksheet_gridlines(worksheet1, LXW_HIDE_ALL_GRIDLINES);
 *
 *    worksheet_gridlines(worksheet2, LXW_SHOW_PRINT_GRIDLINES);
 * @endcode
 *
 * The Excel default is that the screen gridlines are on  and the printed
 * worksheet is off.
 *
 */
//C     void worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);
void  worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);

/**
 * @brief Center the printed page horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data horizontally between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_horizontally(worksheet);
 * @endcode
 *
 */
//C     void worksheet_center_horizontally(lxw_worksheet *worksheet);
void  worksheet_center_horizontally(lxw_worksheet *worksheet);

/**
 * @brief Center the printed page vertically.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data vertically between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_vertically(worksheet);
 * @endcode
 *
 */
//C     void worksheet_center_vertically(lxw_worksheet *worksheet);
void  worksheet_center_vertically(lxw_worksheet *worksheet);

/**
 * @brief Set the option to print the row and column headers on the printed
 *        page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * When printing a worksheet from Excel the row and column headers (the row
 * numbers on the left and the column letters at the top) aren't printed by
 * default.
 *
 * This function sets the printer option to print these headers:
 *
 * @code
 *    worksheet_print_row_col_headers(worksheet);
 * @endcode
 *
 */
//C     void worksheet_print_row_col_headers(lxw_worksheet *worksheet);
void  worksheet_print_row_col_headers(lxw_worksheet *worksheet);

/**
 * @brief Set the number of rows to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row First row of repeat range.
 * @param last_row  Last row of repeat range.
 *
 * For large Excel documents it is often desirable to have the first row or
 * rows of the worksheet print out at the top of each page.
 *
 * This can be achieved by using this function. The parameters `first_row`
 * and `last_row` are zero based:
 *
 * @code
 *     worksheet_repeat_rows(worksheet, 0, 0); // Repeat the first row.
 *     worksheet_repeat_rows(worksheet, 0, 1); // Repeat the first two rows.
 * @endcode
 *
 * @return 0 for success, non-zero on error.
 */
//C     uint8_t worksheet_repeat_rows(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                   lxw_row_t last_row);
uint8_t  worksheet_repeat_rows(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_row_t last_row);

/**
 * @brief Set the number of columns to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col First column of repeat range.
 * @param last_col  Last column of repeat range.
 *
 * For large Excel documents it is often desirable to have the first column or
 * columns of the worksheet print out at the left of each page.
 *
 * This can be achieved by using this function. The parameters `first_col`
 * and `last_col` are zero based:
 *
 * @code
 *     worksheet_repeat_columns(worksheet, 0, 0); // Repeat the first col.
 *     worksheet_repeat_columns(worksheet, 0, 1); // Repeat the first two cols.
 * @endcode
 *
 * @return 0 for success, non-zero on error.
 */
//C     uint8_t worksheet_repeat_columns(lxw_worksheet *worksheet,
//C                                      lxw_col_t first_col, lxw_col_t last_col);
uint8_t  worksheet_repeat_columns(lxw_worksheet *worksheet, lxw_col_t first_col, lxw_col_t last_col);

/**
 * @brief Set the print area for a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * This function is used to specify the area of the worksheet that will be
 * printed. The RANGE() macro is often convenient for this.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 *
 * In order to set a row or column range you must specify the entire range:
 *
 * @code
 *     worksheet_print_area(worksheet, RANGE("A1:H1048576")); // Same as A:H.
 * @endcode
 *
 * @return 0 for success, non-zero on error.
 */
//C     uint8_t worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                  lxw_col_t first_col, lxw_row_t last_row,
//C                                  lxw_col_t last_col);
uint8_t  worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row, lxw_col_t first_col, lxw_row_t last_row, lxw_col_t last_col);
/**
 * @brief Fit the printed area to a specific number of pages both vertically
 *        and horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param width     Number of pages horizontally.
 * @param height    Number of pages vertically.
 *
 * The `%worksheet_fit_to_pages()` function is used to fit the printed area to
 * a specific number of pages both vertically and horizontally. If the printed
 * area exceeds the specified number of pages it will be scaled down to
 * fit. This ensures that the printed area will always appear on the specified
 * number of pages even if the page size or margins change:
 *
 * @code
 *     worksheet_fit_to_pages(worksheet1, 1, 1); // Fit to 1x1 pages.
 *     worksheet_fit_to_pages(worksheet2, 2, 1); // Fit to 2x1 pages.
 *     worksheet_fit_to_pages(worksheet3, 1, 2); // Fit to 1x2 pages.
 * @endcode
 *
 * The print area can be defined using the `worksheet_print_area()` function
 * as described above.
 *
 * A common requirement is to fit the printed output to `n` pages wide but
 * have the height be as long as necessary. To achieve this set the `height`
 * to zero:
 *
 * @code
 *     // 1 page wide and as long as necessary.
 *     worksheet_fit_to_pages(worksheet, 1, 0);
 * @endcode
 *
 * **Note**:
 *
 * - Although it is valid to use both `%worksheet_fit_to_pages()` and
 *   `worksheet_set_print_scale()` on the same worksheet Excel only allows one
 *   of these options to be active at a time. The last function call made will
 *   set the active option.
 *
 * - The `%worksheet_fit_to_pages()` function will override any manual page
 *   breaks that are defined in the worksheet.
 *
 * - When using `%worksheet_fit_to_pages()` it may also be required to set the
 *   printer paper size using `worksheet_set_paper()` or else Excel will
 *   default to "US Letter".
 *
 */
//C     void worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width,
//C                                 uint16_t height);
void  worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width, uint16_t height);

/**
 * @brief Set the start page number when printing.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param start_page Starting page number.
 *
 * The `%worksheet_set_start_page()` function is used to set the number of
 * the starting page when the worksheet is printed out:
 *
 * @code
 *     // Start print from page 2.
 *     worksheet_set_start_page(worksheet, 2);
 * @endcode
 */
//C     void worksheet_set_start_page(lxw_worksheet *worksheet, uint16_t start_page);
void  worksheet_set_start_page(lxw_worksheet *worksheet, uint16_t start_page);

/**
 * @brief Set the scale factor for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param scale     Print scale of worksheet to be printed.
 *
 * This function sets the scale factor of the printed page. The Scale factor
 * must be in the range `10 <= scale <= 400`:
 *
 * @code
 *     worksheet_set_print_scale(worksheet1, 75);
 *     worksheet_set_print_scale(worksheet2, 400);
 * @endcode
 *
 * The default scale factor is 100. Note, `%worksheet_set_print_scale()` does
 * not affect the scale of the visible page in Excel. For that you should use
 * `worksheet_set_zoom()`.
 *
 * Note that although it is valid to use both `worksheet_fit_to_pages()` and
 * `%worksheet_set_print_scale()` on the same worksheet Excel only allows one
 * of these options to be active at a time. The last function call made will
 * set the active option.
 *
 */
//C     void worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);
void  worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);

//C     lxw_worksheet *_new_worksheet(lxw_worksheet_init_data *init_data);
lxw_worksheet * _new_worksheet(lxw_worksheet_init_data *init_data);
//C     void _free_worksheet(lxw_worksheet *worksheet);
void  _free_worksheet(lxw_worksheet *worksheet);
//C     void _worksheet_assemble_xml_file(lxw_worksheet *worksheet);
void  _worksheet_assemble_xml_file(lxw_worksheet *worksheet);
//C     void _worksheet_write_single_row(lxw_worksheet *worksheet);
void  _worksheet_write_single_row(lxw_worksheet *worksheet);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     STATIC void _worksheet_xml_declaration(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_worksheet(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_dimension(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_view(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_views(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_format_pr(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_data(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_page_margins(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_page_setup(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_col_info(lxw_worksheet *worksheet,
//C                                           lxw_col_options *options);
//C     STATIC void _write_row(lxw_worksheet *worksheet, lxw_row *row, char *spans);
//C     STATIC lxw_row *_get_row_list(struct lxw_table_rows *table,
//C                                   lxw_row_t row_num);

//C     STATIC void _worksheet_write_merge_cell(lxw_worksheet *worksheet,
//C                                             lxw_merged_range *merged_range);
//C     STATIC void _worksheet_write_merge_cells(lxw_worksheet *worksheet);

//C     STATIC void _worksheet_write_odd_header(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_odd_footer(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_header_footer(lxw_worksheet *worksheet);

//C     STATIC void _worksheet_write_print_options(lxw_worksheet *worksheet);
//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_WORKSHEET_H__ */
//C     #include "shared_strings.h"
/*
 * libxlsxwriter
 * 
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * shared_strings - A libxlsxwriter library for creating Excel XLSX
 *                  sst files.
 *
 */
//C     #ifndef __LXW_SST_H__
//C     #define __LXW_SST_H__

//C     #include <string.h>
//C     #include <stdint.h>

//C     #include "common.h"

//C     #define NUM_SST_BUCKETS 8
/* STAILQ_HEAD() declaration. */
//C     struct sst_order_list {
//C         struct sst_element *stqh_first;
//C         struct sst_element **stqh_last;
//C     };

/* SLIST_HEAD() declaration. */
//C     struct sst_bucket_list {
//C         struct sst_element *slh_first;
//C     };

/*
 * Elements of the SST table. They contain pointers to allow them to
 * be stored in lists in the the hash table buckets and also pointers to
 * track the insertion order in a separate list.
 */
//C     struct sst_element {
//C         size_t index;
//C         char *string;

//C         struct {
//C             struct sst_element *stqe_next; /* next element */
//C         } sst_order_pointers;
//C         struct {
//C             struct sst_element *sle_next;  /* next element */
//C         } sst_list_pointers;
//C     };

/*
 * Struct to represent a sst.
 */
//C     typedef struct lxw_sst {
//C         FILE *file;

//C         size_t num_buckets;
//C         size_t used_buckets;
//C         size_t string_count;
//C         size_t unique_count;

//C         struct sst_order_list *order_list;
//C         struct sst_bucket_list **buckets;

//C     } lxw_sst;

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     lxw_sst *_new_sst();
//C     void _free_sst(lxw_sst *sst);
//C     int32_t _get_sst_index(lxw_sst *sst, const char *string);
//C     void _sst_assemble_xml_file(lxw_sst *self);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     STATIC void _sst_xml_declaration(lxw_sst *self);

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_SST_H__ */
//C     #include "hash_table.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * hash_table - Hash table functions for libxlsxwriter.
 *
 */

//C     #ifndef __LXW_HASH_TABLE_H__
//C     #define __LXW_HASH_TABLE_H__

//C     #include "common.h"

/* Macro to loop over hash table elements in insertion orfder. */
//C     #define LXW_FOREACH_ORDERED(elem, hash_table)     STAILQ_FOREACH((elem), (hash_table)->order_list, lxw_hash_order_pointers)

/* List declarations. */
//C     struct lxw_hash_order_list {
//C         struct lxw_hash_element *stqh_first;/* first element */
//C         struct lxw_hash_element **stqh_last;/* addr of last next element */
//C     };
//C     struct lxw_hash_bucket_list {
//C         struct lxw_hash_element *slh_first; /* first element */
//C     };

/* LXW_HASH hash table struct. */
//C     typedef struct lxw_hash_table {
//C         size_t num_buckets;
//C         size_t used_buckets;
//C         size_t unique_count;
//C         uint8_t free_key;
//C         uint8_t free_value;

//C         struct lxw_hash_order_list *order_list;
//C         struct lxw_hash_bucket_list **buckets;
//C     } lxw_hash_table;

/*
 * LXW_HASH table element struct.
 *
 * The hash elements contain pointers to allow them to be stored in
 * lists in the the hash table buckets and also pointers to track the
 * insertion order in a separate list.
 */
//C     typedef struct lxw_hash_element {
//C         void *key;
//C         void *value;

//C         struct {
//C             struct lxw_hash_element *stqe_next; /* next element */
//C         } lxw_hash_order_pointers;
//C         struct {
//C             struct lxw_hash_element *sle_next;  /* next element */
//C         } lxw_hash_list_pointers;
//C     } lxw_hash_element;


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     lxw_hash_element *_hash_key_exists(lxw_hash_table *lxw_hash, void *key,
//C                                        size_t key_len);
//C     lxw_hash_element *_insert_hash_element(lxw_hash_table *lxw_hash, void *key,
//C                                            void *value, size_t key_len);
//C     lxw_hash_table *_new_lxw_hash(size_t num_buckets, uint8_t free_key,
//C                                   uint8_t free_value);
//C     void _free_lxw_hash(lxw_hash_table *lxw_hash);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_HASH_TABLE_H__ */
//C     #include "common.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * common - Common functions and defines for the libxlsxwriter library.
 *
 */
//C     #ifndef __LXW_COMMON_H__
//C     #define __LXW_COMMON_H__

//C     #include <time.h>

//C     #ifndef TESTING
//C     #define STATIC static
//C     #else
//C     #define STATIC
//C     #endif

//C     #define LXW_SHEETNAME_MAX  32
//C     #define LXW_SHEETNAME_LEN  65

//C     enum lxw_boolean {
//C         LXW_FALSE,
//C         LXW_TRUE
//C     };

//C     #define LXW_IGNORE 1

//C     #define ERROR(message)                              fprintf(stderr, "[ERROR][%s:%d]: " message "\n", __FILE__, __LINE__)

//C     #define MEM_ERROR()                                 ERROR("Memory allocation failed.")

//C     #define GOTO_LABEL_ON_MEM_ERROR(pointer, label)     if (!pointer) {                                     MEM_ERROR();                                    goto label;                                 }

//C     #define RETURN_ON_MEM_ERROR(pointer, error)         if (!pointer) {                                     MEM_ERROR();                                    return error;                               }

//C     #define LXW_WARN(message)                           fprintf(stderr, "[WARN]: " message "\n")

/* Define the queue.h structs for the formats list. */
//C     struct lxw_formats {
//C         struct lxw_format *stqh_first;/* first element */
//C         struct lxw_format **stqh_last;/* addr of last next element */
//C     };

/* Define the queue.h structs for the generic data structs. */
//C     struct lxw_tuples {
//C         struct lxw_tuple *stqh_first;/* first element */
//C         struct lxw_tuple **stqh_last;/* addr of last next element */
//C     };

//C     typedef struct lxw_tuple {
//C         char *key;
//C         char *value;

//C         struct {
//C             struct lxw_tuple *stqe_next; /* next element */
//C         } list_pointers;
//C     } lxw_tuple;

//C     typedef struct lxw_doc_properties {
//C         char *title;
//C         char *subject;
//C         char *author;
//C         char *manager;
//C         char *company;
//C         char *category;
//C         char *keywords;
//C         char *comments;
//C         char *status;
//C         time_t created;
//C     } lxw_doc_properties;


 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_COMMON_H__ */

/* Define the queue.h structs for the workbook lists. */
//C     struct lxw_worksheets {
//C         struct lxw_worksheet *stqh_first;/* first element */
//C         struct lxw_worksheet **stqh_last;/* addr of last next element */
//C     };
struct lxw_worksheets
{
    lxw_worksheet *stqh_first;
    lxw_worksheet **stqh_last;
}
//C     struct lxw_defined_names {
//C         struct lxw_defined_name *tqh_first; /* first element */
//C         struct lxw_defined_name **tqh_last; /* addr of last next element */
//C     };
struct lxw_defined_names
{
    lxw_defined_name *tqh_first;
    lxw_defined_name **tqh_last;
}

//C     #define LXW_DEFINED_NAME_LENGTH 128

const LXW_DEFINED_NAME_LENGTH = 128;
/* Struct to represent a defined name. */
//C     typedef struct lxw_defined_name {
//C         int16_t index;
//C         uint8_t hidden;
//C         char name[LXW_DEFINED_NAME_LENGTH];
//C         char app_name[LXW_DEFINED_NAME_LENGTH];
//C         char formula[LXW_DEFINED_NAME_LENGTH];
//C         char normalised_name[LXW_DEFINED_NAME_LENGTH];
//C         char normalised_sheetname[LXW_DEFINED_NAME_LENGTH];

    /* List pointers for queue.h. */
//C         struct {
//C             struct lxw_defined_name *tqe_next;  /* next element */
//C             struct lxw_defined_name **tqe_prev; /* address of previous next element */
//C         } list_pointers;
struct _N12
{
    lxw_defined_name *tqe_next;
    lxw_defined_name **tqe_prev;
}
//C     } lxw_defined_name;
struct lxw_defined_name
{
    int16_t index;
    uint8_t hidden;
    char [128]name;
    char [128]app_name;
    char [128]formula;
    char [128]normalised_name;
    char [128]normalised_sheetname;
    _N12 list_pointers;
}

/**
 * @brief Errors conditions encountered when closing the Workbook and writing
 * the Excel file to disk.
 */
//C     enum lxw_close_error {
    /** No error */
//C         LXW_CLOSE_ERROR_NONE,
    /** Error encountered when creating file zip container */
//C         LXW_CLOSE_ERROR_ZIP
        /* TODO. Need to add/document more. */
//C     };
enum lxw_close_error
{
    LXW_CLOSE_ERROR_NONE,
    LXW_CLOSE_ERROR_ZIP,
}

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
//C     typedef struct lxw_workbook_options {
    /** Optimise the workbook to use constant memory for worksheets */
//C         uint8_t constant_memory;
//C     } lxw_workbook_options;
struct lxw_workbook_options
{
    uint8_t constant_memory;
}

/**
 * @brief Struct to represent an Excel workbook.
 *
 * The members of the lxw_workbook struct aren't modified directly. Instead
 * the workbook properties are set by calling the functions shown in
 * workbook.h.
 */
//C     typedef struct lxw_workbook {

//C         FILE *file;
//C         struct lxw_worksheets *worksheets;
//C         struct lxw_formats *formats;
//C         struct lxw_defined_names *defined_names;
//C         lxw_sst *sst;
//C         lxw_doc_properties *properties;
//C         const char *filename;
//C         lxw_workbook_options options;

//C         uint16_t num_sheets;
//C         uint16_t first_sheet;
//C         uint32_t active_sheet;
//C         uint16_t num_xf_formats;
//C         uint16_t num_format_count;

//C         uint16_t font_count;
//C         uint16_t border_count;
//C         uint16_t fill_count;
//C         uint8_t optimize;

//C         lxw_hash_table *used_xf_formats;

//C     } lxw_workbook;
struct lxw_workbook
{
    FILE *file;
    lxw_worksheets *worksheets;
    lxw_formats *formats;
    lxw_defined_names *defined_names;
    lxw_sst *sst;
    lxw_doc_properties *properties;
    char *filename;
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
}


/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
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
//C     lxw_workbook *new_workbook(const char *filename);
lxw_workbook * new_workbook(char *filename);

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
//C     lxw_workbook *new_workbook_opt(const char *filename,
//C                                    lxw_workbook_options *options);
lxw_workbook * new_workbook_opt(char *filename, lxw_workbook_options *options);

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
//C     lxw_worksheet *workbook_add_worksheet(lxw_workbook *workbook,
//C                                           const char *sheetname);
lxw_worksheet * workbook_add_worksheet(lxw_workbook *workbook, char *sheetname);

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
//C     lxw_format *workbook_add_format(lxw_workbook *workbook);
lxw_format * workbook_add_format(lxw_workbook *workbook);

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
//C     uint8_t workbook_close(lxw_workbook *workbook);
uint8_t  workbook_close(lxw_workbook *workbook);

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
//C     uint8_t workbook_define_name(lxw_workbook *workbook, const char *name,
//C                                  const char *formula);
uint8_t  workbook_define_name(lxw_workbook *workbook, char *name, char *formula);

//C     void _free_workbook(lxw_workbook *workbook);
void  _free_workbook(lxw_workbook *workbook);
//C     void _workbook_assemble_xml_file(lxw_workbook *workbook);
void  _workbook_assemble_xml_file(lxw_workbook *workbook);
//C     void _set_default_xf_indices(lxw_workbook *workbook);
void  _set_default_xf_indices(lxw_workbook *workbook);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     STATIC void _workbook_xml_declaration(lxw_workbook *self);
//C     STATIC void _workbook_xml_declaration(lxw_workbook *self);
//C     STATIC void _write_workbook(lxw_workbook *self);
//C     STATIC void _write_file_version(lxw_workbook *self);
//C     STATIC void _write_workbook_pr(lxw_workbook *self);
//C     STATIC void _write_book_views(lxw_workbook *self);
//C     STATIC void _write_workbook_view(lxw_workbook *self);
//C     STATIC void _write_sheet(lxw_workbook *self,
//C                              const char *name, uint32_t sheet_id, uint8_t hidden);
//C     STATIC void _write_sheets(lxw_workbook *self);
//C     STATIC void _write_calc_pr(lxw_workbook *self);

//C     STATIC void _write_defined_name(lxw_workbook *self,
//C                                     lxw_defined_name *define_name);
//C     STATIC void _write_defined_names(lxw_workbook *self);

//C     STATIC uint8_t _store_defined_name(lxw_workbook *self, const char *name,
//C                                        const char *app_name, const char *formula,
//C                                        int16_t index, uint8_t hidden);

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_WORKBOOK_H__ */
//C     #include "xlsxwriter/worksheet.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page worksheet_page The Worksheet object
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * See @ref worksheet.h for full details of the functionality.
 *
 * @file worksheet.h
 *
 * @brief Functions related to adding data and formatting to a worksheet.
 *
 * The Worksheet object represents an Excel worksheet. It handles
 * operations such as writing data to cells or formatting worksheet
 * layout.
 *
 * A Worksheet object isn't created directly. Instead a worksheet is
 * created by calling the workbook_add_worksheet() function from a
 * Workbook object:
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
 */
//C     #ifndef __LXW_WORKSHEET_H__
//C     #define __LXW_WORKSHEET_H__

//C     #include <stdio.h>
//C     #include <stdlib.h>
//C     #include <stdint.h>
//C     #include <string.h>

//C     #include "shared_strings.h"
//C     #include "common.h"
//C     #include "format.h"
//C     #include "utility.h"

//C     #define LXW_ROW_MAX 1048576
//C     #define LXW_COL_MAX 16384
//C     #define LXW_COL_META_MAX 128
//C     #define LXW_HEADER_FOOTER_MAX 255

/* The Excel 2007 specification says that the maximum number of page
 * breaks is 1026. However, in practice it is actually 1023. */
//C     #define LXW_BREAKS_MAX 1023

/** Default column width in Excel */
//C     #define LXW_DEF_COL_WIDTH 8.43

/** Default row height in Excel */
//C     #define LXW_DEF_ROW_HEIGHT 15

/** Error codes from `worksheet_write*()` functions. */
//C     enum lxw_write_error {
    /** No error. */
//C         LXW_WRITE_ERROR_NONE = 0,
    /** Row or column index out of range. */
//C         LXW_RANGE_ERROR,
    /** String exceeds Excel's LXW_STRING_LENGTH_ERROR limit. */
//C         LXW_STRING_LENGTH_ERROR,
    /** Error finding string index. */
//C         LXW_STRING_HASH_ERROR
//C     };

/** Gridline options using in `worksheet_gridlines()`. */
//C     enum lxw_gridlines {
    /** Hide screen and print gridlines. */
//C         LXW_HIDE_ALL_GRIDLINES = 0,
    /** Show screen gridlines. */
//C         LXW_SHOW_SCREEN_GRIDLINES,
    /** Show print gridlines. */
//C         LXW_SHOW_PRINT_GRIDLINES,
    /** Show screen and print gridlines. */
//C         LXW_SHOW_ALL_GRIDLINES
//C     };

/** Data type to represent a row value.
 *
 * The maximum row in Excel is 1,048,576.
 */
//C     typedef uint32_t lxw_row_t;

/** Data type to represent a column value.
 *
 * The maximum column in Excel is 16,384.
 */
//C     typedef uint16_t lxw_col_t;

//C     enum cell_types {
//C         NUMBER_CELL = 1,
//C         STRING_CELL,
//C         INLINE_STRING_CELL,
//C         FORMULA_CELL,
//C         ARRAY_FORMULA_CELL,
//C         BLANK_CELL,
//C         HYPERLINK_URL,
//C         HYPERLINK_INTERNAL,
//C         HYPERLINK_EXTERNAL
//C     };

/* Define the queue.h TAILQ structs for the list head types. */
//C     struct lxw_table_cells {
//C         struct lxw_cell *tqh_first; /* first element */
//C         struct lxw_cell **tqh_last; /* addr of last next element */
//C     };
//C     struct lxw_table_rows {
//C         struct lxw_row *tqh_first; /* first element */
//C         struct lxw_row **tqh_last; /* addr of last next element */
//C     };
//C     struct lxw_merged_ranges {
//C         struct lxw_merged_range *stqh_first;/* first element */
//C         struct lxw_merged_range **stqh_last;/* addr of last next element */
//C     };

/**
 * @brief Options for rows and columns.
 *
 * Options struct for the worksheet_set_column() and worksheet_set_row()
 * functions.
 *
 * It has the following members but currently only the `hidden` property is
 * supported:
 *
 * * `hidden`
 * * `level`
 * * `collapsed`
 */
//C     typedef struct lxw_row_col_options {
    /** Hide the row/column */
//C         uint8_t hidden;
//C         uint8_t level;
//C         uint8_t collapsed;
//C     } lxw_row_col_options;

//C     typedef struct lxw_col_options {
//C         lxw_col_t firstcol;
//C         lxw_col_t lastcol;
//C         double width;
//C         lxw_format *format;
//C         uint8_t hidden;
//C         uint8_t level;
//C         uint8_t collapsed;
//C     } lxw_col_options;

//C     typedef struct lxw_merged_range {
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;

//C         struct {
//C             struct lxw_merged_range *stqe_next; /* next element */
//C         } list_pointers;
//C     } lxw_merged_range;

//C     typedef struct lxw_repeat_rows {
//C         uint8_t in_use;
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C     } lxw_repeat_rows;

//C     typedef struct lxw_repeat_cols {
//C         uint8_t in_use;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;
//C     } lxw_repeat_cols;

//C     typedef struct lxw_print_area {
//C         uint8_t in_use;
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;
//C     } lxw_print_area;

//C     typedef struct lxw_autofilter {
//C         uint8_t in_use;
//C         lxw_row_t first_row;
//C         lxw_row_t last_row;
//C         lxw_col_t first_col;
//C         lxw_col_t last_col;
//C     } lxw_autofilter;

/**
 * @brief Header and footer options.
 *
 * Optional parameters used in the worksheet_set_header_opt() and
 * worksheet_set_footer_opt() functions.
 *
 */
//C     typedef struct lxw_header_footer_options {
    /** Header or footer margin in inches. Excel default is 0.3. */
//C         double margin;
//C     } lxw_header_footer_options;

/**
 * @brief Struct to represent an Excel worksheet.
 *
 * The members of the lxw_worksheet struct aren't modified directly. Instead
 * the worksheet properties are set by calling the functions shown in
 * worksheet.h.
 */
//C     typedef struct lxw_worksheet {

//C         FILE *file;
//C         FILE *optimize_tmpfile;
//C         struct lxw_table_rows *table;
//C         struct lxw_table_rows *hyperlinks;
//C         struct lxw_cell **array;
//C         struct lxw_merged_ranges *merged_ranges;

//C         lxw_row_t dim_rowmin;
//C         lxw_row_t dim_rowmax;
//C         lxw_col_t dim_colmin;
//C         lxw_col_t dim_colmax;

//C         lxw_sst *sst;
//C         char *name;
//C         char *quoted_name;

//C         uint32_t index;
//C         uint8_t active;
//C         uint8_t selected;
//C         uint8_t hidden;
//C         uint32_t *active_sheet;

//C         lxw_col_options **col_options;
//C         uint16_t col_options_max;

//C         double *col_sizes;
//C         uint16_t col_sizes_max;

//C         lxw_format **col_formats;
//C         uint16_t col_formats_max;

//C         uint8_t col_size_changed;
//C         uint8_t optimize;
//C         struct lxw_row *optimize_row;

//C         uint16_t fit_height;
//C         uint16_t fit_width;
//C         uint16_t horizontal_dpi;
//C         uint16_t hlink_count;
//C         uint16_t page_start;
//C         uint16_t print_scale;
//C         uint16_t rel_count;
//C         uint16_t vertical_dpi;
//C         uint8_t filter_on;
//C         uint8_t fit_page;
//C         uint8_t hcenter;
//C         uint8_t orientation;
//C         uint8_t outline_changed;
//C         uint8_t page_order;
//C         uint8_t page_setup_changed;
//C         uint8_t page_view;
//C         uint8_t paper_size;
//C         uint8_t print_gridlines;
//C         uint8_t print_headers;
//C         uint8_t print_options_changed;
//C         uint8_t screen_gridlines;
//C         uint8_t tab_color;
//C         uint8_t vba_codename;
//C         uint8_t vcenter;

//C         double margin_left;
//C         double margin_right;
//C         double margin_top;
//C         double margin_bottom;
//C         double margin_header;
//C         double margin_footer;

//C         uint8_t header_footer_changed;
//C         char header[LXW_HEADER_FOOTER_MAX];
//C         char footer[LXW_HEADER_FOOTER_MAX];

//C         struct lxw_repeat_rows repeat_rows;
//C         struct lxw_repeat_cols repeat_cols;
//C         struct lxw_print_area print_area;
//C         struct lxw_autofilter autofilter;

//C         uint16_t merged_range_count;

//C         lxw_row_t *hbreaks;
//C         lxw_col_t *vbreaks;

//C         struct lxw_rel_tuples *external_hyperlinks;

//C         struct {
//C             struct lxw_worksheet *stqe_next; /* next element */
//C         } list_pointers;

//C     } lxw_worksheet;

/*
 * Worksheet initialisation data.
 */
//C     typedef struct lxw_worksheet_init_data {
//C         uint32_t index;
//C         uint8_t hidden;
//C         uint8_t optimize;
//C         uint32_t *active_sheet;
//C         lxw_sst *sst;
//C         char *name;
//C         char *quoted_name;

//C     } lxw_worksheet_init_data;

/* Struct to represent a worksheet row. */
//C     typedef struct lxw_row {
//C         lxw_row_t row_num;
//C         double height;
//C         lxw_format *format;
//C         uint8_t hidden;
//C         uint8_t level;
//C         uint8_t collapsed;
//C         uint8_t row_changed;
//C         uint8_t data_changed;
//C         struct lxw_table_cells *cells;

    /* List pointers for queue.h. */
//C         struct {
//C             struct lxw_row *tqe_next;  /* next element */
//C             struct lxw_row **tqe_prev; /* address of previous next element */
//C         } list_pointers;
//C     } lxw_row;

/* Struct to represent a worksheet cell. */
//C     typedef struct lxw_cell {
//C         lxw_row_t row_num;
//C         lxw_col_t col_num;
//C         enum cell_types type;
//C         lxw_format *format;

//C         union {
//C             double number;
//C             int32_t string_id;
//C             char *string;
//C         } u;

//C         double formula_result;
//C         char *user_data1;
//C         char *user_data2;

    /* List pointers for queue.h. */
//C         struct {
//C             struct lxw_cell *tqe_next;  /* next element */
//C             struct lxw_cell **tqe_prev; /* address of previous next element */
//C         } list_pointers;
//C     } lxw_cell;

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

/**
 * @brief Write a number to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param number    The number to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `worksheet_write_number()` function writes numeric types to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_number(worksheet, 0, 0, 123456, NULL);
 *     worksheet_write_number(worksheet, 1, 0, 2.3451, NULL);
 * @endcode
 *
 * @image html write_number01.png
 *
 * The native data type for all numbers in Excel is a IEEE-754 64-bit
 * double-precision floating point, which is also the default type used by
 * `%worksheet_write_number`.
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 *     format_set_num_format(format, "$#,##0.00");
 *
 *     worksheet_write_number(worksheet, 0, 0, 1234.567, format);
 * @endcode
 *
 * @image html write_number02.png
 *
 */
//C     int8_t worksheet_write_number(lxw_worksheet *worksheet,
//C                                   lxw_row_t row,
//C                                   lxw_col_t col, double number,
//C                                   lxw_format *format);
/**
 * @brief Write a string to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param string    String to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `%worksheet_write_string()` function writes a string to the cell
 * specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is English!", NULL);
 * @endcode
 *
 * @image html write_string01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object:
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 *     format_set_bold(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This phrase is Bold!", format);
 * @endcode
 *
 * @image html write_string02.png
 *
 * Unicode strings are supported in UTF-8 encoding. This generally requires
 * that your source file is UTF-8 encoded or that the data has been read from
 * a UTF-8 source:
 *
 * @code
 *    worksheet_write_string(worksheet, 0, 0, "Это фраза на русском!", NULL);
 * @endcode
 *
 * @image html write_string03.png
 *
 */
//C     int8_t worksheet_write_string(lxw_worksheet *worksheet,
//C                                   lxw_row_t row,
//C                                   lxw_col_t col, const char *string,
//C                                   lxw_format *format);
/**
 * @brief Write a formula to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `%worksheet_write_formula()` function writes a formula or function to
 * the cell specified by `row` and `column`:
 *
 * @code
 *  worksheet_write_formula(worksheet, 0, 0, "=B3 + 6",                    NULL);
 *  worksheet_write_formula(worksheet, 1, 0, "=SIN(PI()/4)",               NULL);
 *  worksheet_write_formula(worksheet, 2, 0, "=SUM(A1:A2)",                NULL);
 *  worksheet_write_formula(worksheet, 3, 0, "=IF(A3>1,\"Yes\", \"No\")",  NULL);
 *  worksheet_write_formula(worksheet, 4, 0, "=AVERAGE(1, 2, 3, 4)",       NULL);
 *  worksheet_write_formula(worksheet, 5, 0, "=DATEVALUE(\"1-Jan-2013\")", NULL);
 * @endcode
 *
 * @image html write_formula01.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a
 * @ref format.h "Format" object.
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores a
 * default value of `0`. The correct formula result is displayed in Excel, as
 * shown in the example above, since it recalculates the formulas when it loads
 * the file. For cases where this is an issue see the
 * `worksheet_write_formula_num()` function and the discussion in that section.
 *
 * Formulas must be written with the US style separator/range operator which
 * is a comma (not semi-colon). Therefore a formula with multiple values
 * should be written as follows:
 *
 * @code
 *     // OK.
 *     worksheet_write_formula(worksheet, 0, 0, "=SUM(1, 2, 3)", NULL);
 *
 *     // NO. Error on load.
 *     worksheet_write_formula(worksheet, 1, 0, "=SUM(1; 2; 3)", NULL);
 * @endcode
 *
 */
//C     int8_t worksheet_write_formula(lxw_worksheet *worksheet,
//C                                    lxw_row_t row,
//C                                    lxw_col_t col, const char *formula,
//C                                    lxw_format *format);
/**
 * @brief Write an array formula to a worksheet cell.
 *
 * @param worksheet
 * @param first_row   The first row of the range. (All zero indexed.)
 * @param first_col   The first column of the range.
 * @param last_row    The last row of the range.
 * @param last_col    The last col of the range.
 * @param formula     Array formula to write to cell.
 * @param format      A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
  * The `%worksheet_write_array_formula()` function writes an array formula to
 * a cell range. In Excel an array formula is a formula that performs a
 * calculation on a set of values.
 *
 * In Excel an array formula is indicated by a pair of braces around the
 * formula: `{=SUM(A1:B1*A2:B2)}`.
 *
 * Array formulas can return a single value or a range or values. For array
 * formulas that return a range of values you must specify the range that the
 * return values will be written to. This is why this function has `first_`
 * and `last_` row/column parameters. The RANGE() macro can also be used to
 * specify the range:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 4, 0, 6, 0,     "{=TREND(C5:C7,B5:B7)}", NULL);
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_write_array_formula(worksheet, RANGE("A5:A7"), "{=TREND(C5:C7,B5:B7)}", NULL);
 * @endcode
 *
 * If the array formula returns a single value then the `first_` and `last_`
 * parameters should be the same:
 *
 * @code
 *     worksheet_write_array_formula(worksheet, 1, 0, 1, 0,     "{=SUM(B1:C1*B2:C2)}", NULL);
 *     worksheet_write_array_formula(worksheet, RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}", NULL);
 * @endcode
 *
 */
//C     int8_t worksheet_write_array_formula(lxw_worksheet *worksheet,
//C                                          lxw_row_t first_row,
//C                                          lxw_col_t first_col,
//C                                          lxw_row_t last_row,
//C                                          lxw_col_t last_col,
//C                                          const char *formula, lxw_format *format);

//C     int8_t worksheet_write_array_formula_num(lxw_worksheet *worksheet,
//C                                              lxw_row_t first_row,
//C                                              lxw_col_t first_col,
//C                                              lxw_row_t last_row,
//C                                              lxw_col_t last_col,
//C                                              const char *formula,
//C                                              lxw_format *format, double result);

/**
 * @brief Write a date or time to a worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param datetime  The datetime to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * The `worksheet_write_datetime()` function can be used to write a date or
 * time to the cell specified by `row` and `column`:
 *
 * @dontinclude dates_and_times02.c
 * @skip include
 * @until num_format
 * @skip Feb
 * @until }
 *
 * The `format` parameter should be used to apply formatting to the cell using
 * a @ref format.h "Format" object as shown above. Without a date format the
 * datetime will appear as a number only.
 *
 * See @ref working_with_dates for more information about handling dates and
 * times in libxlsxwriter.
 */
//C     int8_t worksheet_write_datetime(lxw_worksheet *worksheet,
//C                                     lxw_row_t row,
//C                                     lxw_col_t col, lxw_datetime *datetime,
//C                                     lxw_format *format);

//C     int8_t worksheet_write_url_opt(lxw_worksheet *worksheet,
//C                                    lxw_row_t row_num,
//C                                    lxw_col_t col_num, const char *url,
//C                                    lxw_format *format, const char *string,
//C                                    const char *tooltip);
/**
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param url       The url to write to the cell.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 *
 * The `%worksheet_write_url()` function is used to write a URL/hyperlink to a
 * worksheet cell specified by `row` and `column`.
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "http://libxlsxwriter.github.io", url_format);
 * @endcode
 *
 * @image html hyperlinks_short.png
 *
 * The `format` parameter is used to apply formatting to the cell. This
 * parameter can be `NULL` to indicate no formatting or it can be a @ref
 * format.h "Format" object. The typical worksheet format for a hyperlink is a
 * blue underline:
 *
 * @code
 *    lxw_format *url_format   = workbook_add_format(workbook);
 *
 *    format_set_underline (url_format, LXW_UNDERLINE_SINGLE);
 *    format_set_font_color(url_format, LXW_COLOR_BLUE);
 *
 * @endcode
 *
 * The usual web style URI's are supported: `%http://`, `%https://`, `%ftp://`
 * and `mailto:` :
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "ftp://www.python.org/",    url_format);
 *     worksheet_write_url(worksheet, 1, 0, "http://www.python.org/",   url_format);
 *     worksheet_write_url(worksheet, 2, 0, "https://www.python.org/",  url_format);
 *     worksheet_write_url(worksheet, 3, 0, "mailto:jmcnamaracpan.org", url_format);
 *
 * @endcode
 *
 * An Excel hyperlink is comprised of two elements: the displayed string and
 * the non-displayed link. By default the displayed string is the same as the
 * link. However, it is possible to overwrite it with any other
 * `libxlsxwriter` type using the appropriate `worksheet_write_*()`
 * function. The most common case is to overwrite the displayed link text with
 * another string:
 *
 * @code
 *  // Write a hyperlink but overwrite the displayed string.
 *  worksheet_write_url   (worksheet, 2, 0, "http://libxlsxwriter.github.io", url_format);
 *  worksheet_write_string(worksheet, 2, 0, "Read the documentation.",        url_format);
 *
 * @endcode
 *
 * @image html hyperlinks_short2.png
 *
 * Two local URIs are supported: `internal:` and `external:`. These are used
 * for hyperlinks to internal worksheet references or external workbook and
 * worksheet references:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:Sheet2!A1",                url_format);
 *     worksheet_write_url(worksheet, 1, 0, "internal:Sheet2!B2",                url_format);
 *     worksheet_write_url(worksheet, 2, 0, "internal:Sheet2!A1:B2",             url_format);
 *     worksheet_write_url(worksheet, 3, 0, "internal:'Sales Data'!A1",          url_format);
 *     worksheet_write_url(worksheet, 4, 0, "external:c:\\temp\\foo.xlsx",       url_format);
 *     worksheet_write_url(worksheet, 5, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
 *     worksheet_write_url(worksheet, 6, 0, "external:..\\foo.xlsx",             url_format);
 *     worksheet_write_url(worksheet, 7, 0, "external:..\\foo.xlsx#Sheet2!A1",   url_format);
 *     worksheet_write_url(worksheet, 8, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
 *
 * @endcode
 *
 * Worksheet references are typically of the form `Sheet1!A1`. You can also
 * link to a worksheet range using the standard Excel notation:
 * `Sheet1!A1:B2`.
 *
 * In external links the workbook and worksheet name must be separated by the
 * `#` character:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\foo.xlsx#Sheet2!A1",   url_format);
 * @endcode
 *
 * You can also link to a named range in the target worksheet: For example say
 * you have a named range called `my_name` in the workbook `c:\temp\foo.xlsx`
 * you could link to it as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:\\temp\\foo.xlsx#my_name", url_format);
 *
 * @endcode
 *
 * Excel requires that worksheet names containing spaces or non alphanumeric
 * characters are single quoted as follows:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "internal:'Sales Data'!A1", url_format);
 * @endcode
 *
 * Links to network files are also supported. Network files normally begin
 * with two back slashes as follows `\\NETWORK\etc`. In order to represent
 * this in a C string literal the backslashes should be escaped:
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:\\\\NET\\share\\foo.xlsx", url_format);
 * @endcode
 *
 *
 * Alternatively, you can use Windows style forward slashes. These are
 * translated internally to backslashes:
 *
 * @code
 *     worksheet_write_url(worksheet, 0, 0, "external:c:/temp/foo.xlsx",     url_format);
 *     worksheet_write_url(worksheet, 1, 0, "external://NET/share/foo.xlsx", url_format);
 *
 * @endcode
 *
 *
 * **Note:**
 *
 *    libxlsxwriter will escape the following characters in URLs as required
 *    by Excel: `\s " < > \ [ ]  ^ { }` unless the URL already contains `%%xx`
 *    style escapes. In which case it is assumed that the URL was escaped
 *    correctly by the user and will by passed directly to Excel.
 *
 */
//C     int8_t worksheet_write_url(lxw_worksheet *worksheet,
//C                                lxw_row_t row,
//C                                lxw_col_t col, const char *url,
//C                                lxw_format *format);

/**
 * @brief Write a formatted blank worksheet cell.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return A #lxw_write_error code.
 *
 * Write a blank cell specified by `row` and `column`:
 *
 * @code
 *     worksheet_write_blank(worksheet, 1, 1, border_format);
 * @endcode
 *
 * This function is used to add formatting to a cell which doesn't contain a
 * string or number value.
 *
 * Excel differentiates between an "Empty" cell and a "Blank" cell. An Empty
 * cell is a cell which doesn't contain data or formatting whilst a Blank cell
 * doesn't contain data but does contain formatting. Excel stores Blank cells
 * but ignores Empty cells.
 *
 * As such, if you write an empty cell without formatting it is ignored.
 *
 */
//C     int8_t worksheet_write_blank(lxw_worksheet *worksheet,
//C                                  lxw_row_t row, lxw_col_t col,
//C                                  lxw_format *format);

/**
 * @brief Write a formula to a worksheet cell with a user defined result.
 *
 * @param worksheet Pointer to the lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param col       The zero indexed column number.
 * @param formula   Formula string to write to cell.
 * @param format    A pointer to a Format instance or NULL.
 * @param result    A user defined result for a formula.
 *
 * @return A #lxw_write_error code.
 *
 * The `%worksheet_write_formula_num()` function writes a formula or Excel
 * function to the cell specified by `row` and `column` with a user defined
 * result:
 *
 * @code
 *     // Required as a workaround only.
 *     worksheet_write_formula_num(worksheet, 0, 0, "=1 + 2", NULL, 3);
 * @endcode
 *
 * Libxlsxwriter doesn't calculate the value of a formula and instead stores
 * the value `0` as the formula result. It then sets a global flag in the XLSX
 * file to say that all formulas and functions should be recalculated when the
 * file is opened.
 *
 * This is the method recommended in the Excel documentation and in general it
 * works fine with spreadsheet applications.
 *
 * However, applications that don't have a facility to calculate formulas,
 * such as Excel Viewer, or some mobile applications will only display the `0`
 * results.
 *
 * If required, the `%worksheet_write_formula_num()` function can be used to
 * specify a formula and its result.
 *
 * This function is rarely required and is only provided for compatibility
 * with some third party applications. For most applications the
 * worksheet_write_formula() function is the recommended way of writing
 * formulas.
 *
 */
//C     int8_t worksheet_write_formula_num(lxw_worksheet *worksheet,
//C                                        lxw_row_t row,
//C                                        lxw_col_t col,
//C                                        const char *formula,
//C                                        lxw_format *format, double result);

/**
 * @brief Set the properties for a row of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param row       The zero indexed row number.
 * @param height    The row height.
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * The `%worksheet_set_row()` function is used to change the default
 * properties of a row. The most common use for this function is to change the
 * height of a row:
 *
 * @code
 *     // Set the height of Row 1 to 20.
 *     worksheet_set_row(worksheet, 0, 20, NULL, NULL);
 * @endcode
 *
 * The other common use for `%worksheet_set_row()` is to set the a @ref
 * format.h "Format" for all cells in the row:
 *
 * @code
 *     lxw_format *bold = workbook_add_format(workbook);
 *     format_set_bold(bold);
 *
 *     // Set the header row to bold.
 *     worksheet_set_row(worksheet, 0, 15, bold, NULL);
 * @endcode
 *
 * If you wish to set the format of a row without changing the height you can
 * pass the default row height of #LXW_DEF_ROW_HEIGHT = 15:
 *
 * @code
 *     worksheet_set_row(worksheet, 0, LXW_DEF_ROW_HEIGHT, format, NULL);
 *     worksheet_set_row(worksheet, 0, 15, format, NULL); // Same as above.
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the row that don't
 * have a format. As with Excel the row format is overridden by an explicit
 * cell format. For example:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1, NULL);
 *
 *     // Cell A1 in Row 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell B1 in Row 1 keeps format2.
 *     worksheet_write_string(worksheet, 0, 1, "Hello", format2);
 * @endcode
 *
 * The `options` parameter is a #lxw_row_col_options struct. It has the
 * following members but currently only the `hidden` property is supported:
 *
 * - `hidden`
 * - `level`
 * - `collapsed`
 *
 * The `"hidden"` option is used to hide a row. This can be used, for
 * example, to hide intermediary steps in a complicated calculation:
 *
 * @code
 *     lxw_row_col_options options = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     // Hide the fourth row.
 *     worksheet_set_row(worksheet, 3, 20, NULL, &options);
 * @endcode
 *
 */
//C     int8_t worksheet_set_row(lxw_worksheet *worksheet,
//C                              lxw_row_t row,
//C                              double height,
//C                              lxw_format *format, lxw_row_col_options *options);

/**
 * @brief Set the properties for one or more columns of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col The zero indexed first column.
 * @param last_col  The zero indexed last column.
 * @param width     The width of the column(s).
 * @param format    A pointer to a Format instance or NULL.
 * @param options   Optional row parameters: hidden, level, collapsed.
 *
 * The `%worksheet_set_column()` function can be used to change the default
 * properties of a single column or a range of columns:
 *
 * @code
 *     // Width of columns B:D set to 30.
 *     worksheet_set_column(worksheet, 1, 3, 30, NULL, NULL);
 *
 * @endcode
 *
 * If `%worksheet_set_column()` is applied to a single column the value of
 * `first_col` and `last_col` should be the same:
 *
 * @code
 *     // Width of column B set to 30.
 *     worksheet_set_column(worksheet, 1, 1, 30, NULL, NULL);
 *
 * @endcode
 *
 * It is also possible, and generally clearer, to specify a column range using
 * the form of `COLS()` macro:
 *
 * @code
 *     worksheet_set_column(worksheet, 4, 4, 20, NULL, NULL);
 *     worksheet_set_column(worksheet, 5, 8, 30, NULL, NULL);
 *
 *     // Same as the examples above but clearer.
 *     worksheet_set_column(worksheet, COLS("E:E"), 20, NULL, NULL);
 *     worksheet_set_column(worksheet, COLS("F:H"), 30, NULL, NULL);
 *
 * @endcode
 *
 * The width corresponds to the column width value that is specified in
 * Excel. It is approximately equal to the length of a string in the default
 * font of Calibri 11. Unfortunately, there is no way to specify "AutoFit" for
 * a column in the Excel file format. This feature is only available at
 * runtime from within Excel. It is possible to simulate "AutoFit" by tracking
 * the width of the data in the column as your write it.
 *
 * As usual the @ref format.h `format` parameter is optional. If you wish to
 * set the format without changing the width you can pass default col width of
 * #LXW_DEF_COL_WIDTH = 8.43:
 *
 * @code
 *     lxw_format *bold = workbook_add_format(workbook);
 *     format_set_bold(bold);
 *
 *     // Set the first column to bold.
 *     worksheet_set_column(worksheet, 0, 0, LXW_DEF_COL_HEIGHT, bold, NULL);
 * @endcode
 *
 * The `format` parameter will be applied to any cells in the column that
 * don't have a format. For example:
 *
 * @code
 *     // Column 1 has format1.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format1, NULL);
 *
 *     // Cell A1 in column 1 defaults to format1.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *     // Cell A2 in column 1 keeps format2.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", format2);
 * @endcode
 *
 * As in Excel a row format takes precedence over a default column format:
 *
 * @code
 *     // Row 1 has format1.
 *     worksheet_set_row(worksheet, 0, 15, format1, NULL);
 *
 *     // Col 1 has format2.
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, format2, NULL);
 *
 *     // Cell A1 defaults to format1, the row format.
 *     worksheet_write_string(worksheet, 0, 0, "Hello", NULL);
 *
 *    // Cell A2 keeps format2, the column format.
 *     worksheet_write_string(worksheet, 1, 0, "Hello", NULL);
 * @endcode
 *
 * The `options` parameter is a #lxw_row_col_options struct. It has the
 * following members but currently only the `hidden` property is supported:
 *
 * - `hidden`
 * - `level`
 * - `collapsed`
 *
 * The `"hidden"` option is used to hide a column. This can be used, for
 * example, to hide intermediary steps in a complicated calculation:
 *
 * @code
 *     lxw_row_col_options options = {.hidden = 1, .level = 0, .collapsed = 0};
 *
 *     worksheet_set_column(worksheet, COLS("A:A"), 8.43, NULL, &options);
 * @endcode
 *
 */
//C     int8_t worksheet_set_column(lxw_worksheet *worksheet, lxw_col_t first_col,
//C                                 lxw_col_t last_col, double width,
//C                                 lxw_format *format, lxw_row_col_options *options);

/**
 * @brief Merge a range of cells.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 * @param string    String to write to the merged range.
 * @param format    A pointer to a Format instance or NULL.
 *
 * @return 0 for success, non-zero on error.
 *
 * The `%worksheet_merge_range()` function allows cells to be merged together
 * so that they act as a single area.
 *
 * Excel generally merges and centers cells at same time. To get similar
 * behaviour with libxlsxwriter you need to apply a @ref format.h "Format"
 * object with the appropriate alignment:
 *
 * @code
 *     lxw_format *merge_format = workbook_add_format(workbook);
 *     format_set_align(merge_format, LXW_ALIGN_CENTER);
 *
 *     worksheet_merge_range(worksheet, 1, 1, 1, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * It is possible to apply other formatting to the merged cells as well:
 *
 * @code
 *    format_set_align   (merge_format, LXW_ALIGN_CENTER);
 *    format_set_align   (merge_format, LXW_ALIGN_VERTICAL_CENTER);
 *    format_set_border  (merge_format, LXW_BORDER_DOUBLE);
 *    format_set_bold    (merge_format);
 *    format_set_bg_color(merge_format, 0xD7E4BC);
 *
 *    worksheet_merge_range(worksheet, 2, 1, 3, 3, "Merged Range", merge_format);
 *
 * @endcode
 *
 * @image html merge.png
 *
 * The `%worksheet_merge_range()` function writes a `char*` string using
 * `worksheet_write_string()`. In order to write other data types, such as a
 * number or a formula, you can overwrite the first cell with a call to one of
 * the other write functions. The same Format should be used as was used in
 * the merged range.
 *
 * @code
 *    // First write a range with a blank string.
 *    worksheet_merge_range (worksheet, 1, 1, 1, 3, "", format);
 *
 *    // Then overwrite the first cell with a number.
 *    worksheet_write_number(worksheet, 1, 1, 123, format);
 * @endcode
 */
//C     uint8_t worksheet_merge_range(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                   lxw_col_t first_col, lxw_row_t last_row,
//C                                   lxw_col_t last_col, const char *string,
//C                                   lxw_format *format);

/**
 * @brief Set the autofilter area in the worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * @return 0 for success, non-zero on error.
 *
 * The `%worksheet_autofilter()` method allows an autofilter to be added to a
 * worksheet.
 *
 * An autofilter is a way of adding drop down lists to the headers of a 2D
 * range of worksheet data. This allows users to filter the data based on
 * simple criteria so that some data is shown and some is hidden.
 *
 * @image html autofilter.png
 *
 * To add an autofilter to a worksheet:
 *
 * @code
 *     worksheet_autofilter(worksheet, 0, 0, 50, 3);
 *
 *     // Same as above using the RANGE() macro.
 *     worksheet_autofilter(worksheet, RANGE("A1:D51"));
 * @endcode
 *
 * Note: it isn't currently possible to apply filter conditions to the
 * autofilter.
 */
//C     uint8_t worksheet_autofilter(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                  lxw_col_t first_col, lxw_row_t last_row,
//C                                  lxw_col_t last_col);

 /**
  * @brief Make a worksheet the active, i.e., visible worksheet.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_activate()` function is used to specify which worksheet is
  * initially visible in a multi-sheet workbook:
  *
  * @code
  *     lxw_worksheet *worksheet1 = workbook_add_worksheet(workbook, NULL);
  *     lxw_worksheet *worksheet2 = workbook_add_worksheet(workbook, NULL);
  *     lxw_worksheet *worksheet3 = workbook_add_worksheet(workbook, NULL);
  *
  *     worksheet_activate(worksheet3);
  * @endcode
  *
  * @image html worksheet_activate.png
  *
  * More than one worksheet can be selected via the `worksheet_select()`
  * function, see below, however only one worksheet can be active.
  *
  * The default active worksheet is the first worksheet.
  *
  */
//C     void worksheet_activate(lxw_worksheet *worksheet);

 /**
  * @brief Set a worksheet tab as selected.
  *
  * @param worksheet Pointer to a lxw_worksheet instance to be updated.
  *
  * The `%worksheet_select()` function is used to indicate that a worksheet is
  * selected in a multi-sheet workbook:
  *
  * @code
  *     worksheet_activate(worksheet1);
  *     worksheet_select(worksheet2);
  *     worksheet_select(worksheet3);
  *
  * @endcode
  *
  * A selected worksheet has its tab highlighted. Selecting worksheets is a
  * way of grouping them together so that, for example, several worksheets
  * could be printed in one go. A worksheet that has been activated via the
  * `worksheet_activate()` function will also appear as selected.
  *
  */
//C     void worksheet_select(lxw_worksheet *worksheet);

/**
 * @brief Set the page orientation as landscape.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to landscape:
 *
 * @code
 *     worksheet_set_landscape(worksheet);
 * @endcode
 */
//C     void worksheet_set_landscape(lxw_worksheet *worksheet);

/**
 * @brief Set the page orientation as portrait.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to set the orientation of a worksheet's printed page
 * to portrait. The default worksheet orientation is portrait, so this
 * function isn't generally required:
 *
 * @code
 *     worksheet_set_portrait(worksheet);
 * @endcode
 */
//C     void worksheet_set_portrait(lxw_worksheet *worksheet);

/**
 * @brief Set the page layout to page view mode.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * This function is used to display the worksheet in "Page View/Layout" mode:
 *
 * @code
 *     worksheet_set_page_view(worksheet);
 * @endcode
 */
//C     void worksheet_set_page_view(lxw_worksheet *worksheet);

/**
 * @brief Set the paper type for printing.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param paper_type The Excel paper format type.
 *
 * This function is used to set the paper format for the printed output of a
 * worksheet. The following paper styles are available:
 *
 *
 *   Index    | Paper format            | Paper size
 *   :------- | :---------------------- | :-------------------
 *   0        | Printer default         | Printer default
 *   1        | Letter                  | 8 1/2 x 11 in
 *   2        | Letter Small            | 8 1/2 x 11 in
 *   3        | Tabloid                 | 11 x 17 in
 *   4        | Ledger                  | 17 x 11 in
 *   5        | Legal                   | 8 1/2 x 14 in
 *   6        | Statement               | 5 1/2 x 8 1/2 in
 *   7        | Executive               | 7 1/4 x 10 1/2 in
 *   8        | A3                      | 297 x 420 mm
 *   9        | A4                      | 210 x 297 mm
 *   10       | A4 Small                | 210 x 297 mm
 *   11       | A5                      | 148 x 210 mm
 *   12       | B4                      | 250 x 354 mm
 *   13       | B5                      | 182 x 257 mm
 *   14       | Folio                   | 8 1/2 x 13 in
 *   15       | Quarto                  | 215 x 275 mm
 *   16       | ---                     | 10x14 in
 *   17       | ---                     | 11x17 in
 *   18       | Note                    | 8 1/2 x 11 in
 *   19       | Envelope 9              | 3 7/8 x 8 7/8
 *   20       | Envelope 10             | 4 1/8 x 9 1/2
 *   21       | Envelope 11             | 4 1/2 x 10 3/8
 *   22       | Envelope 12             | 4 3/4 x 11
 *   23       | Envelope 14             | 5 x 11 1/2
 *   24       | C size sheet            | ---
 *   25       | D size sheet            | ---
 *   26       | E size sheet            | ---
 *   27       | Envelope DL             | 110 x 220 mm
 *   28       | Envelope C3             | 324 x 458 mm
 *   29       | Envelope C4             | 229 x 324 mm
 *   30       | Envelope C5             | 162 x 229 mm
 *   31       | Envelope C6             | 114 x 162 mm
 *   32       | Envelope C65            | 114 x 229 mm
 *   33       | Envelope B4             | 250 x 353 mm
 *   34       | Envelope B5             | 176 x 250 mm
 *   35       | Envelope B6             | 176 x 125 mm
 *   36       | Envelope                | 110 x 230 mm
 *   37       | Monarch                 | 3.875 x 7.5 in
 *   38       | Envelope                | 3 5/8 x 6 1/2 in
 *   39       | Fanfold                 | 14 7/8 x 11 in
 *   40       | German Std Fanfold      | 8 1/2 x 12 in
 *   41       | German Legal Fanfold    | 8 1/2 x 13 in
 *
 * Note, it is likely that not all of these paper types will be available to
 * the end user since it will depend on the paper formats that the user's
 * printer supports. Therefore, it is best to stick to standard paper types:
 *
 * @code
 *     worksheet_set_paper(worksheet1, 1);  // US Letter
 *     worksheet_set_paper(worksheet2, 9);  // A4
 * @endcode
 *
 * If you do not specify a paper type the worksheet will print using the
 * printer's default paper style.
 */
//C     void worksheet_set_paper(lxw_worksheet *worksheet, uint8_t paper_type);

/**
 * @brief Set the worksheet margins for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param left    Left margin in inches.   Excel default is 0.7.
 * @param right   Right margin in inches.  Excel default is 0.7.
 * @param top     Top margin in inches.    Excel default is 0.75.
 * @param bottom  Bottom margin in inches. Excel default is 0.75.
 *
 * The `%worksheet_set_margins()` function is used to set the margins of the
 * worksheet when it is printed. The units are in inches. Specifying `-1` for
 * any parameter will give the default Excel value as shown above.
 *
 * @code
 *    worksheet_set_margins(worksheet, 1.3, 1.2, -1, -1);
 * @endcode
 *
 */
//C     void worksheet_set_margins(lxw_worksheet *worksheet, double left,
//C                                double right, double top, double bottom);

/**
 * @brief Set the printed page header caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 *
 * @return 0 for success, non-zero on error.
 *
 * Headers and footers are generated using a string which is a combination of
 * plain text and control characters.
 *
 * The available control character are:
 *
 *
 *   | Control         | Category      | Description           |
 *   | --------------- | ------------- | --------------------- |
 *   | `&L`            | Justification | Left                  |
 *   | `&C`            |               | Center                |
 *   | `&R`            |               | Right                 |
 *   | `&P`            | Information   | Page number           |
 *   | `&N`            |               | Total number of pages |
 *   | `&D`            |               | Date                  |
 *   | `&T`            |               | Time                  |
 *   | `&F`            |               | File name             |
 *   | `&A`            |               | Worksheet name        |
 *   | `&Z`            |               | Workbook path         |
 *   | `&fontsize`     | Font          | Font size             |
 *   | `&"font,style"` |               | Font name and style   |
 *   | `&U`            |               | Single underline      |
 *   | `&E`            |               | Double underline      |
 *   | `&S`            |               | Strikethrough         |
 *   | `&X`            |               | Superscript           |
 *   | `&Y`            |               | Subscript             |
 *
 *
 * Text in headers and footers can be justified (aligned) to the left, center
 * and right by prefixing the text with the control characters `&L`, `&C` and
 * `&R`.
 *
 * For example (with ASCII art representation of the results):
 *
 * @code
 *     worksheet_set_header(worksheet, "&LHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Hello                                                         |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&CHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 *
 *     worksheet_set_header(worksheet, "&RHello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                                                         Hello |
 *     |                                                               |
 *
 *
 * @endcode
 *
 * For simple text, if you do not specify any justification the text will be
 * centred. However, you must prefix the text with `&C` if you specify a font
 * name or any other formatting:
 *
 * @code
 *     worksheet_set_header(worksheet, "Hello");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                          Hello                                |
 *     |                                                               |
 *
 * @endcode
 *
 * You can have text in each of the justification regions:
 *
 * @code
 *     worksheet_set_header(worksheet, "&LCiao&CBello&RCielo");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     | Ciao                     Bello                          Cielo |
 *     |                                                               |
 *
 * @endcode
 *
 * The information control characters act as variables that Excel will update
 * as the workbook or worksheet changes. Times and dates are in the users
 * default format:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CPage &P of &N");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                        Page 1 of 6                            |
 *     |                                                               |
 *
 *     worksheet_set_header(worksheet, "&CUpdated at &T");
 *
 *      ---------------------------------------------------------------
 *     |                                                               |
 *     |                    Updated at 12:30 PM                        |
 *     |                                                               |
 *
 * @endcode
 *
 * You can specify the font size of a section of the text by prefixing it with
 * the control character `&n` where `n` is the font size:
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&30Hello Big");
 *     worksheet_set_header(worksheet2, "&C&10Hello Small");
 *
 * @endcode
 *
 * You can specify the font of a section of the text by prefixing it with the
 * control sequence `&"font,style"` where `fontname` is a font name such as
 * Windows font descriptions: "Regular", "Italic", "Bold" or "Bold Italic":
 * "Courier New" or "Times New Roman" and `style` is one of the standard
 *
 * @code
 *     worksheet_set_header(worksheet1, "&C&\"Courier New,Italic\"Hello");
 *     worksheet_set_header(worksheet2, "&C&\"Courier New,Bold Italic\"Hello");
 *     worksheet_set_header(worksheet3, "&C&\"Times New Roman,Regular\"Hello");
 *
 * @endcode
 *
 * It is possible to combine all of these features together to create
 * sophisticated headers and footers. As an aid to setting up complicated
 * headers and footers you can record a page set-up as a macro in Excel and
 * look at the format strings that VBA produces. Remember however that VBA
 * uses two double quotes `""` to indicate a single double quote. For the last
 * example above the equivalent VBA code looks like this:
 *
 * @code
 *     .LeftHeader = ""
 *     .CenterHeader = "&""Times New Roman,Regular""Hello"
 *     .RightHeader = ""
 *
 * @endcode
 *
 * Alternatively you can inspect the header and footer strings in an Excel
 * file by unzipping it and grepping the XML sub-files. The following shows
 * how to do that using libxml's xmllint to format the XML for clarity:
 *
 * @code
 *
 *    $ unzip myfile.xlsm -d myfile
 *    $ xmllint --format `find myfile -name "*.xml" | xargs` | egrep "Header|Footer"
 *
 *      <headerFooter scaleWithDoc="0">
 *        <oddHeader>&amp;L&amp;P</oddHeader>
 *      </headerFooter>
 *
 * @endcode
 *
 * Note that in this case you need to unescape the Html. In the above example
 * the header string would be `&L&P`.
 *
 * To include a single literal ampersand `&` in a header or footer you should
 * use a double ampersand `&&`:
 *
 * @code
 *     worksheet_set_header(worksheet, "&CCuriouser && Curiouser - Attorneys at Law");
 * @endcode
 *
 * Note, the header or footer string must be less than 255 characters. Strings
 * longer than this will not be written.
 *
 */
//C     uint8_t worksheet_set_header(lxw_worksheet *worksheet, char *string);

/**
 * @brief Set the printed page footer caption.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 *
 * @return 0 for success, non-zero on error.
 *
 * The syntax of this function is the same as worksheet_set_header().
 *
 */
//C     uint8_t worksheet_set_footer(lxw_worksheet *worksheet, char *string);

/**
 * @brief Set the printed page header caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The header string.
 * @param options   Header options.
 *
 * @return 0 for success, non-zero on error.
 *
 * The syntax of this function is the same as worksheet_set_header() with an
 * additional parameter to specify options for the header.
 *
 * Currently, the only available option is the header margin:
 *
 * @code
 *
 *    lxw_header_footer_options header_options = { 0.2 };
 *
 *    worksheet_set_header_opt(worksheet, "Some text", &header_options);
 *
 * @endcode
 *
 */
//C     uint8_t worksheet_set_header_opt(lxw_worksheet *worksheet, char *string,
//C                                      lxw_header_footer_options *options);

/**
 * @brief Set the printed page footer caption with additional options.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param string    The footer string.
 * @param options   Footer options.
 *
 * @return 0 for success, non-zero on error.
 *
 * The syntax of this function is the same as worksheet_set_header_opt().
 *
 */
//C     uint8_t worksheet_set_footer_opt(lxw_worksheet *worksheet, char *string,
//C                                      lxw_header_footer_options *options);

/**
 * @brief Set the horizontal page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * The `%worksheet_set_h_pagebreaks()` function adds horizontal page breaks to
 * a worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Horizontal page breaks act between rows.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxw_row_t and the last element of the array must be 0:
 *
 * @code
 *    lxw_row_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxw_row_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet1, breaks1);
 *    worksheet_set_h_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between rows 20 and 21 you must specify the break at
 * row 21. However in zero index notation this is actually row 20:
 *
 * @code
 *    // Break between row 20 and 21.
 *    lxw_row_t breaks[] = {20, 0};
 *
 *    worksheet_set_h_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 horizontal page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
//C     void worksheet_set_h_pagebreaks(lxw_worksheet *worksheet, lxw_row_t breaks[]);

/**
 * @brief Set the vertical page breaks on a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param breaks    Array of page breaks.
 *
 * The `%worksheet_set_v_pagebreaks()` function adds vertical page breaks to a
 * worksheet. A page break causes all the data that follows it to be printed
 * on the next page. Vertical page breaks act between columns.
 *
 * The function takes an array of one or more page breaks. The type of the
 * array data is @ref lxw_col_t and the last element of the array must be 0:
 *
 * @code
 *    lxw_col_t breaks1[] = {20, 0}; // 1 page break. Zero indicates the end.
 *    lxw_col_t breaks2[] = {20, 40, 60, 80, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet1, breaks1);
 *    worksheet_set_v_pagebreaks(worksheet2, breaks2);
 * @endcode
 *
 * To create a page break between columns 20 and 21 you must specify the break
 * at column 21. However in zero index notation this is actually column 20:
 *
 * @code
 *    // Break between column 20 and 21.
 *    lxw_col_t breaks[] = {20, 0};
 *
 *    worksheet_set_v_pagebreaks(worksheet, breaks);
 * @endcode
 *
 * There is an Excel limitation of 1023 vertical page breaks per worksheet.
 *
 * Note: If you specify the "fit to page" option via the
 * `worksheet_fit_to_pages()` function it will override all manual page
 * breaks.
 *
 */
//C     void worksheet_set_v_pagebreaks(lxw_worksheet *worksheet, lxw_col_t breaks[]);

/**
 * @brief Set the order in which pages are printed.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * The `%worksheet_print_across()` function is used to change the default
 * print direction. This is referred to by Excel as the sheet "page order":
 *
 * @code
 *     worksheet_print_across(worksheet);
 * @endcode
 *
 * The default page order is shown below for a worksheet that extends over 4
 * pages. The order is called "down then across":
 *
 *     [1] [3]
 *     [2] [4]
 *
 * However, by using the `print_across` function the print order will be
 * changed to "across then down":
 *
 *     [1] [2]
 *     [3] [4]
 *
 */
//C     void worksheet_print_across(lxw_worksheet *worksheet);

/**
 * @brief Set the option to display or hide gridlines on the screen and
 *        the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param option    Gridline option.
 *
 * Display or hide screen and print gridlines using one of the values of
 * @ref lxw_gridlines.
 *
 * @code
 *    worksheet_gridlines(worksheet1, LXW_HIDE_ALL_GRIDLINES);
 *
 *    worksheet_gridlines(worksheet2, LXW_SHOW_PRINT_GRIDLINES);
 * @endcode
 *
 * The Excel default is that the screen gridlines are on  and the printed
 * worksheet is off.
 *
 */
//C     void worksheet_gridlines(lxw_worksheet *worksheet, uint8_t option);

/**
 * @brief Center the printed page horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data horizontally between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_horizontally(worksheet);
 * @endcode
 *
 */
//C     void worksheet_center_horizontally(lxw_worksheet *worksheet);

/**
 * @brief Center the printed page vertically.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * Center the worksheet data vertically between the margins on the printed
 * page:
 *
 * @code
 *     worksheet_center_vertically(worksheet);
 * @endcode
 *
 */
//C     void worksheet_center_vertically(lxw_worksheet *worksheet);

/**
 * @brief Set the option to print the row and column headers on the printed
 *        page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 *
 * When printing a worksheet from Excel the row and column headers (the row
 * numbers on the left and the column letters at the top) aren't printed by
 * default.
 *
 * This function sets the printer option to print these headers:
 *
 * @code
 *    worksheet_print_row_col_headers(worksheet);
 * @endcode
 *
 */
//C     void worksheet_print_row_col_headers(lxw_worksheet *worksheet);

/**
 * @brief Set the number of rows to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row First row of repeat range.
 * @param last_row  Last row of repeat range.
 *
 * For large Excel documents it is often desirable to have the first row or
 * rows of the worksheet print out at the top of each page.
 *
 * This can be achieved by using this function. The parameters `first_row`
 * and `last_row` are zero based:
 *
 * @code
 *     worksheet_repeat_rows(worksheet, 0, 0); // Repeat the first row.
 *     worksheet_repeat_rows(worksheet, 0, 1); // Repeat the first two rows.
 * @endcode
 *
 * @return 0 for success, non-zero on error.
 */
//C     uint8_t worksheet_repeat_rows(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                   lxw_row_t last_row);

/**
 * @brief Set the number of columns to repeat at the top of each printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_col First column of repeat range.
 * @param last_col  Last column of repeat range.
 *
 * For large Excel documents it is often desirable to have the first column or
 * columns of the worksheet print out at the left of each page.
 *
 * This can be achieved by using this function. The parameters `first_col`
 * and `last_col` are zero based:
 *
 * @code
 *     worksheet_repeat_columns(worksheet, 0, 0); // Repeat the first col.
 *     worksheet_repeat_columns(worksheet, 0, 1); // Repeat the first two cols.
 * @endcode
 *
 * @return 0 for success, non-zero on error.
 */
//C     uint8_t worksheet_repeat_columns(lxw_worksheet *worksheet,
//C                                      lxw_col_t first_col, lxw_col_t last_col);

/**
 * @brief Set the print area for a worksheet.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param first_row The first row of the range. (All zero indexed.)
 * @param first_col The first column of the range.
 * @param last_row  The last row of the range.
 * @param last_col  The last col of the range.
 *
 * This function is used to specify the area of the worksheet that will be
 * printed. The RANGE() macro is often convenient for this.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 *
 * In order to set a row or column range you must specify the entire range:
 *
 * @code
 *     worksheet_print_area(worksheet, RANGE("A1:H1048576")); // Same as A:H.
 * @endcode
 *
 * @return 0 for success, non-zero on error.
 */
//C     uint8_t worksheet_print_area(lxw_worksheet *worksheet, lxw_row_t first_row,
//C                                  lxw_col_t first_col, lxw_row_t last_row,
//C                                  lxw_col_t last_col);
/**
 * @brief Fit the printed area to a specific number of pages both vertically
 *        and horizontally.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param width     Number of pages horizontally.
 * @param height    Number of pages vertically.
 *
 * The `%worksheet_fit_to_pages()` function is used to fit the printed area to
 * a specific number of pages both vertically and horizontally. If the printed
 * area exceeds the specified number of pages it will be scaled down to
 * fit. This ensures that the printed area will always appear on the specified
 * number of pages even if the page size or margins change:
 *
 * @code
 *     worksheet_fit_to_pages(worksheet1, 1, 1); // Fit to 1x1 pages.
 *     worksheet_fit_to_pages(worksheet2, 2, 1); // Fit to 2x1 pages.
 *     worksheet_fit_to_pages(worksheet3, 1, 2); // Fit to 1x2 pages.
 * @endcode
 *
 * The print area can be defined using the `worksheet_print_area()` function
 * as described above.
 *
 * A common requirement is to fit the printed output to `n` pages wide but
 * have the height be as long as necessary. To achieve this set the `height`
 * to zero:
 *
 * @code
 *     // 1 page wide and as long as necessary.
 *     worksheet_fit_to_pages(worksheet, 1, 0);
 * @endcode
 *
 * **Note**:
 *
 * - Although it is valid to use both `%worksheet_fit_to_pages()` and
 *   `worksheet_set_print_scale()` on the same worksheet Excel only allows one
 *   of these options to be active at a time. The last function call made will
 *   set the active option.
 *
 * - The `%worksheet_fit_to_pages()` function will override any manual page
 *   breaks that are defined in the worksheet.
 *
 * - When using `%worksheet_fit_to_pages()` it may also be required to set the
 *   printer paper size using `worksheet_set_paper()` or else Excel will
 *   default to "US Letter".
 *
 */
//C     void worksheet_fit_to_pages(lxw_worksheet *worksheet, uint16_t width,
//C                                 uint16_t height);

/**
 * @brief Set the start page number when printing.
 *
 * @param worksheet  Pointer to a lxw_worksheet instance to be updated.
 * @param start_page Starting page number.
 *
 * The `%worksheet_set_start_page()` function is used to set the number of
 * the starting page when the worksheet is printed out:
 *
 * @code
 *     // Start print from page 2.
 *     worksheet_set_start_page(worksheet, 2);
 * @endcode
 */
//C     void worksheet_set_start_page(lxw_worksheet *worksheet, uint16_t start_page);

/**
 * @brief Set the scale factor for the printed page.
 *
 * @param worksheet Pointer to a lxw_worksheet instance to be updated.
 * @param scale     Print scale of worksheet to be printed.
 *
 * This function sets the scale factor of the printed page. The Scale factor
 * must be in the range `10 <= scale <= 400`:
 *
 * @code
 *     worksheet_set_print_scale(worksheet1, 75);
 *     worksheet_set_print_scale(worksheet2, 400);
 * @endcode
 *
 * The default scale factor is 100. Note, `%worksheet_set_print_scale()` does
 * not affect the scale of the visible page in Excel. For that you should use
 * `worksheet_set_zoom()`.
 *
 * Note that although it is valid to use both `worksheet_fit_to_pages()` and
 * `%worksheet_set_print_scale()` on the same worksheet Excel only allows one
 * of these options to be active at a time. The last function call made will
 * set the active option.
 *
 */
//C     void worksheet_set_print_scale(lxw_worksheet *worksheet, uint16_t scale);

//C     lxw_worksheet *_new_worksheet(lxw_worksheet_init_data *init_data);
//C     void _free_worksheet(lxw_worksheet *worksheet);
//C     void _worksheet_assemble_xml_file(lxw_worksheet *worksheet);
//C     void _worksheet_write_single_row(lxw_worksheet *worksheet);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     STATIC void _worksheet_xml_declaration(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_worksheet(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_dimension(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_view(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_views(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_format_pr(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_sheet_data(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_page_margins(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_page_setup(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_col_info(lxw_worksheet *worksheet,
//C                                           lxw_col_options *options);
//C     STATIC void _write_row(lxw_worksheet *worksheet, lxw_row *row, char *spans);
//C     STATIC lxw_row *_get_row_list(struct lxw_table_rows *table,
//C                                   lxw_row_t row_num);

//C     STATIC void _worksheet_write_merge_cell(lxw_worksheet *worksheet,
//C                                             lxw_merged_range *merged_range);
//C     STATIC void _worksheet_write_merge_cells(lxw_worksheet *worksheet);

//C     STATIC void _worksheet_write_odd_header(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_odd_footer(lxw_worksheet *worksheet);
//C     STATIC void _worksheet_write_header_footer(lxw_worksheet *worksheet);

//C     STATIC void _worksheet_write_print_options(lxw_worksheet *worksheet);
//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_WORKSHEET_H__ */
//C     #include "xlsxwriter/format.h"
/*
 * libxlsxwriter
 * 
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @page format_page The Format object
 *
 * The Format object represents an the formatting properties that can be
 * applied to a cell including: fonts, colors, patterns,
 * borders, alignment and number formatting.
 *
 * See @ref format.h for full details of the functionality.
 * 
 * @file format.h
 *
 * @brief Functions and properties for adding formatting to cells in Excel.
 *
 * This section describes the functions and properties that are available for
 * formatting cells in Excel.
 *
 * The properties of a cell that can be formatted include: fonts, colors,
 * patterns, borders, alignment and number formatting.
 *
 * @image html formats_intro.png
 *
 * Formats in `libxlswriter` are accessed via the lxw_format
 * struct. Throughout this document these will be referred to simply as
 * *Formats*.
 *
 * Formats are created by calling the workbook_add_format() method as
 * follows:
 *
 * @code
 *     lxw_format *format = workbook_add_format(workbook);
 * @endcode
 *
 * The members of the lxw_format struct aren't modified directly. Instead the
 * format properties are set by calling the function shown in this section.
 * For example:
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
 *
 * @endcode
 *
 * The full range of formatting options that can be applied using
 * `libxlswriter` are shown below.
 *
 */
//C     #ifndef __LXW_FORMAT_H__
//C     #define __LXW_FORMAT_H__

//C     #include <stdint.h>
//C     #include <string.h>
//C     #include "hash_table.h"

//C     #include "common.h"

/**
 * @brief The type for RGB colors in libxlswriter.
 *
 * The type for RGB colors in libxlswriter. The valid range is `0x000000`
 * (black) to `0xFFFFFF` (white). See @ref working_with_colors.
 */
//C     typedef int32_t lxw_color_t;

//C     #define LXW_FORMAT_FIELD_LEN            128
//C     #define LXW_DEFAULT_FONT_NAME           "Calibri"
//C     #define LXW_DEFAULT_FONT_FAMILY         2
//C     #define LXW_DEFAULT_FONT_THEME          1
//C     #define LXW_PROPERTY_UNSET              -1
//C     #define LXW_COLOR_UNSET                 -1
//C     #define LXW_COLOR_MASK                  0xFFFFFF
//C     #define LXW_MIN_FONT_SIZE               1
//C     #define LXW_MAX_FONT_SIZE               409

//C     #define LXW_FORMAT_FIELD_COPY(dst, src)                 do{                                                     strncpy(dst, src, LXW_FORMAT_FIELD_LEN -1);         dst[LXW_FORMAT_FIELD_LEN - 1] = '\0';           } while (0)

/** Format underline values for format_set_underline(). */
//C     enum lxw_format_underlines {
    /** Single underline */
//C         LXW_UNDERLINE_SINGLE = 1,

    /** Double underline */
//C         LXW_UNDERLINE_DOUBLE,

    /** Single accounting underline */
//C         LXW_UNDERLINE_SINGLE_ACCOUNTING,

    /** Double accounting underline */
//C         LXW_UNDERLINE_DOUBLE_ACCOUNTING
//C     };

/** Superscript and subscript values for format_set_font_script(). */
//C     enum lxw_format_scripts {

    /** Superscript font */
//C         LXW_FONT_SUPERSCRIPT = 1,

    /** Subscript font */
//C         LXW_FONT_SUBSCRIPT
//C     };

/** Alignment values for format_set_align(). */
//C     enum lxw_format_alignments {
    /** No alignment. Cell will use Excel's default for the data type */
//C         LXW_ALIGN_NONE = 0,

    /** Left horizontal alignment */
//C         LXW_ALIGN_LEFT,

    /** Center horizontal alignment */
//C         LXW_ALIGN_CENTER,

    /** Right horizontal alignment */
//C         LXW_ALIGN_RIGHT,

    /** Cell fill horizontal alignment */
//C         LXW_ALIGN_FILL,

    /** Justify horizontal alignment */
//C         LXW_ALIGN_JUSTIFY,

    /** Center Across horizontal alignment */
//C         LXW_ALIGN_CENTER_ACROSS,

    /** Left horizontal alignment */
//C         LXW_ALIGN_DISTRIBUTED,

    /** Top vertical alignment */
//C         LXW_ALIGN_VERTICAL_TOP,

    /** Bottom vertical alignment */
//C         LXW_ALIGN_VERTICAL_BOTTOM,

    /** Center vertical alignment */
//C         LXW_ALIGN_VERTICAL_CENTER,

    /** Justify vertical alignment */
//C         LXW_ALIGN_VERTICAL_JUSTIFY,

    /** Distributed vertical alignment */
//C         LXW_ALIGN_VERTICAL_DISTRIBUTED
//C     };

//C     enum lxw_format_diagonal_types {
//C         LXW_DIAGONAL_BORDER_UP = 1,
//C         LXW_DIAGONAL_BORDER_DOWN,
//C         LXW_DIAGONAL_BORDER_UP_DOWN
//C     };

/** Predefined values for common colors. */
//C     enum lxw_defined_colors {
    /** Black */
//C         LXW_COLOR_BLACK = 0x000000,

    /** Blue */
//C         LXW_COLOR_BLUE = 0x0000FF,

    /** Brown */
//C         LXW_COLOR_BROWN = 0x800000,

    /** Cyan */
//C         LXW_COLOR_CYAN = 0x00FFFF,

    /** Gray */
//C         LXW_COLOR_GRAY = 0x808080,

    /** Green */
//C         LXW_COLOR_GREEN = 0x008000,

    /** Lime */
//C         LXW_COLOR_LIME = 0x00FF00,

    /** Magenta */
//C         LXW_COLOR_MAGENTA = 0xFF00FF,

    /** Navy */
//C         LXW_COLOR_NAVY = 0x000080,

    /** Orange */
//C         LXW_COLOR_ORANGE = 0xFF6600,

    /** Pink */
//C         LXW_COLOR_PINK = 0xFF00FF,

    /** Purple */
//C         LXW_COLOR_PURPLE = 0x800080,

    /** Red */
//C         LXW_COLOR_RED = 0xFF0000,

    /** Silver */
//C         LXW_COLOR_SILVER = 0xC0C0C0,

    /** White */
//C         LXW_COLOR_WHITE = 0xFFFFFF,

    /** Yellow */
//C         LXW_COLOR_YELLOW = 0xFFFF00
//C     };

/** Pattern value for use with format_set_pattern(). */
//C     enum lxw_format_patterns {
    /** Empty pattern */
//C         LXW_PATTERN_NONE = 0,

    /** Solid pattern */
//C         LXW_PATTERN_SOLID,

    /** Medium gray pattern */
//C         LXW_PATTERN_MEDIUM_GRAY,

    /** Dark gray pattern */
//C         LXW_PATTERN_DARK_GRAY,

    /** Light gray pattern */
//C         LXW_PATTERN_LIGHT_GRAY,

    /** Dark horizontal line pattern */
//C         LXW_PATTERN_DARK_HORIZONTAL,

    /** Dark vertical line pattern */
//C         LXW_PATTERN_DARK_VERTICAL,

    /** Dark diagonal stripe pattern */
//C         LXW_PATTERN_DARK_DOWN,

    /** Reverse dark diagonal stripe pattern */
//C         LXW_PATTERN_DARK_UP,

    /** Dark grid pattern */
//C         LXW_PATTERN_DARK_GRID,

    /** Dark trellis pattern */
//C         LXW_PATTERN_DARK_TRELLIS,

    /** Light horizontal Line pattern */
//C         LXW_PATTERN_LIGHT_HORIZONTAL,

    /** Light vertical line pattern */
//C         LXW_PATTERN_LIGHT_VERTICAL,

    /** Light diagonal stripe pattern */
//C         LXW_PATTERN_LIGHT_DOWN,

    /** Reverse light diagonal stripe pattern */
//C         LXW_PATTERN_LIGHT_UP,

    /** Light grid pattern */
//C         LXW_PATTERN_LIGHT_GRID,

    /** Light trellis pattern */
//C         LXW_PATTERN_LIGHT_TRELLIS,

    /** 12.5% gray pattern */
//C         LXW_PATTERN_GRAY_125,

    /** 6.25% gray pattern */
//C         LXW_PATTERN_GRAY_0625
//C     };

/** Cell border styles for use with format_set_border(). */
//C     enum lxw_format_borders {
    /** No border */
//C         LXW_BORDER_NONE,

    /** Thin border style */
//C         LXW_BORDER_THIN,

    /** Medium border style */
//C         LXW_BORDER_MEDIUM,

    /** Dashed border style */
//C         LXW_BORDER_DASHED,

    /** Dotted border style */
//C         LXW_BORDER_DOTTED,

    /** Thick border style */
//C         LXW_BORDER_THICK,

    /** Double border style */
//C         LXW_BORDER_DOUBLE,

    /** Hair border style */
//C         LXW_BORDER_HAIR,

    /** Medium dashed border style */
//C         LXW_BORDER_MEDIUM_DASHED,

    /** Dash-dot border style */
//C         LXW_BORDER_DASH_DOT,

    /** Medium dash-dot border style */
//C         LXW_BORDER_MEDIUM_DASH_DOT,

    /** Dash-dot-dot border style */
//C         LXW_BORDER_DASH_DOT_DOT,

    /** Medium dash-dot-dot border style */
//C         LXW_BORDER_MEDIUM_DASH_DOT_DOT,

    /** Slant dash-dot border style */
//C         LXW_BORDER_SLANT_DASH_DOT
//C     };

/**
 * @brief Struct to represent the formatting properties of an Excel format.
 *
 * Formats in `libxlswriter` are accessed via this struct.
 *
 * The members of the lxw_format struct aren't modified directly. Instead the
 * format properties are set by calling the functions shown in format.h.
 *
 * For example:
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
 *
 * @endcode
 *
 */
//C     typedef struct lxw_format {

//C         FILE *file;

//C         lxw_hash_table *xf_format_indices;
//C         uint16_t *num_xf_formats;

//C         int32_t xf_index;
//C         int32_t dxf_index;

//C         char num_format[LXW_FORMAT_FIELD_LEN];
//C         char font_name[LXW_FORMAT_FIELD_LEN];
//C         char font_scheme[LXW_FORMAT_FIELD_LEN];
//C         uint8_t num_format_index;
//C         uint16_t font_index;
//C         uint8_t has_font;
//C         uint8_t has_dxf_font;
//C         uint16_t font_size;
//C         uint8_t bold;
//C         uint8_t italic;
//C         lxw_color_t font_color;
//C         uint8_t underline;
//C         uint8_t font_strikeout;
//C         uint8_t font_outline;
//C         uint8_t font_shadow;
//C         uint8_t font_script;
//C         uint8_t font_family;
//C         uint8_t font_charset;
//C         uint8_t font_condense;
//C         uint8_t font_extend;
//C         uint8_t theme;
//C         uint8_t hyperlink;

//C         uint8_t hidden;
//C         uint8_t locked;

//C         uint8_t text_h_align;
//C         uint8_t text_wrap;
//C         uint8_t text_v_align;
//C         uint8_t text_justlast;
//C         int16_t rotation;

//C         lxw_color_t fg_color;
//C         lxw_color_t bg_color;
//C         uint8_t pattern;
//C         uint8_t has_fill;
//C         uint8_t has_dxf_fill;
//C         int32_t fill_index;
//C         int32_t fill_count;

//C         int32_t border_index;
//C         uint8_t has_border;
//C         uint8_t has_dxf_border;
//C         int32_t border_count;

//C         uint8_t bottom;
//C         uint8_t diag_border;
//C         uint8_t diag_type;
//C         uint8_t left;
//C         uint8_t right;
//C         uint8_t top;
//C         lxw_color_t bottom_color;
//C         lxw_color_t diag_color;
//C         lxw_color_t left_color;
//C         lxw_color_t right_color;
//C         lxw_color_t top_color;

//C         uint8_t indent;
//C         uint8_t shrink;
//C         uint8_t merge_range;
//C         uint8_t reading_order;
//C         uint8_t just_distrib;
//C         uint8_t color_indexed;
//C         uint8_t font_only;

//C         struct {
//C             struct lxw_format *stqe_next; /* next element */
//C         } list_pointers;
//C     } lxw_format;

/*
 * Struct to represent the font component of a format.
 */
//C     typedef struct lxw_font {

//C         char font_name[LXW_FORMAT_FIELD_LEN];
//C         uint16_t font_size;
//C         uint8_t bold;
//C         uint8_t italic;
//C         lxw_color_t font_color;
//C         uint8_t underline;
//C         uint8_t font_strikeout;
//C         uint8_t font_outline;
//C         uint8_t font_shadow;
//C         uint8_t font_script;
//C         uint8_t font_family;
//C         uint8_t font_charset;
//C         uint8_t font_condense;
//C         uint8_t font_extend;
//C     } lxw_font;

/*
 * Struct to represent the border component of a format.
 */
//C     typedef struct lxw_border {

//C         uint8_t bottom;
//C         uint8_t diag_border;
//C         uint8_t diag_type;
//C         uint8_t left;
//C         uint8_t right;
//C         uint8_t top;

//C         lxw_color_t bottom_color;
//C         lxw_color_t diag_color;
//C         lxw_color_t left_color;
//C         lxw_color_t right_color;
//C         lxw_color_t top_color;

//C     } lxw_border;

/*
 * Struct to represent the fill component of a format.
 */
//C     typedef struct lxw_fill {

//C         lxw_color_t fg_color;
//C         lxw_color_t bg_color;
//C         uint8_t pattern;

//C     } lxw_fill;


/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     lxw_format *_new_format();
//C     void _free_format(lxw_format *format);
//C     int32_t _get_xf_index(lxw_format *format);
//C     lxw_font *_get_font_key(lxw_format *format);
//C     lxw_border *_get_border_key(lxw_format *format);
//C     lxw_fill *_get_fill_key(lxw_format *format);

/**
 * @brief Set the font used in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param font_name Cell font name.
 *
 * Specify the font used used in the cell format:
 *
 * @code
 *     format_set_font_name(format, "Avenir Black Oblique");
 * @endcode
 *
 * @image html format_set_font_name.png
 *
 * Excel can only display fonts that are installed on the system that it is
 * running on. Therefore it is generally best to use the fonts that come as
 * standard with Excel such as Calibri, Times New Roman and Courier New.
 *
 * The default font in Excel 2007, and later, is Calibri.
 */
//C     void format_set_font_name(lxw_format *format, const char *font_name);

/**
 * @brief Set the size of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param size   The cell font size.
 *
 * Set the font size of the cell format:
 *
 * @code
 *     format_set_font_size(format, 30);
 * @endcode
 *
 * Excel adjusts the height of a row to accommodate the largest font
 * size in the row. You can also explicitly specify the height of a
 * row using the worksheet_set_row() function.
 */
//C     void format_set_font_size(lxw_format *format, uint16_t size);

/**
 * @brief Set the color of the font used in the cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell font color.
 *
 *
 * Set the font color:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_font_color(format, "red");
 *
 *     worksheet_write_string(worksheet, 0, 0, "wheelbarrow", format);
 * @endcode
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 * @note
 * The format_set_font_color() method is used to set the font color in a
 * cell. To set the color of a cell background use the format_set_bg_color()
 * and format_set_pattern() methods.
 */
//C     void format_set_font_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Turn on bold for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the bold property of the font:
 *
 * @code
 *     format_set_bold(format);
 * @endcode
 */
//C     void format_set_bold(lxw_format *format);

/**
 * @brief Turn on italic for the format font.
 *
 * @param format Pointer to a Format instance.
 *
 * Set the italic property of the font:
 *
 * @code
 *     format_set_italic(format);
 * @endcode
 */
//C     void format_set_italic(lxw_format *format);

/**
 * @brief Turn on underline for the format:
 *
 * @param format Pointer to a Format instance.
 * @param style Underline style.
 *
 * Set the underline property of the format:
 *
 * @code
 *     format_set_underline(format, LXW_UNDERLINE_SINGLE);
 * @endcode
 *
 * The available underline styles are:
 *
 * - #LXW_UNDERLINE_SINGLE
 * - #LXW_UNDERLINE_DOUBLE
 * - #LXW_UNDERLINE_SINGLE_ACCOUNTING
 * - #LXW_UNDERLINE_DOUBLE_ACCOUNTING
 *
 */
//C     void format_set_underline(lxw_format *format, uint8_t style);

/**
 * @brief Set the strikeout property of the font.
 *
 * @param format Pointer to a Format instance.
 */
//C     void format_set_font_strikeout(lxw_format *format);

/**
 * @brief Set the superscript/subscript property of the font.
 *
 * @param format Pointer to a Format instance.
 * @param style  Superscript or subscript style.
 *
 * Set the superscript o subscript property of the font.
 *
 * The available script styles are:
 *
 * - #LXW_FONT_SUPERSCRIPT
 * - #LXW_FONT_SUBSCRIPT
 */
//C     void format_set_font_script(lxw_format *format, uint8_t style);

/**
 * @brief Set the number format for a cell.
 *
 * @param format      Pointer to a Format instance.
 * @param num_format The cell number format string.
 *
 * This method is used to define the numerical format of a number in
 * Excel. It controls whether a number is displayed as an integer, a
 * floating point number, a date, a currency value or some other user
 * defined format.
 *
 * The numerical format of a cell can be specified by using a format
 * string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_num_format(format, "d mmm yyyy");
 * @endcode
 *
 * Format strings can control any aspect of number formatting allowed by Excel:
 *
 * @dontinclude format_num_format.c
 * @skipline set_num_format
 * @until 1209
 * 
 * @image html format_set_num_format.png
 *
 * The number system used for dates is described in @ref working_with_dates.
 *
 * For more information on number formats in Excel refer to the
 * [Microsoft documentation on cell formats](http://office.microsoft.com/en-gb/assistance/HP051995001033.aspx).
 */
//C     void format_set_num_format(lxw_format *format, const char *num_format);

/**
 * @brief Set the Excel built-in number format for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param index  The built-in number format index for the cell.
 *
 * This function is similar to format_set_num_format() except that it takes an
 * index to a limited number of Excel's built-in number formats instead of a
 * user defined format string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_num_format(format, 0x0F);     // d-mmm-yy
 * @endcode
 *
 * @note
 *
 * Unless you need to specifically access one of Excel's built-in number
 * formats the format_set_num_format() function above is a better
 * solution. The format_set_num_format_index() function is mainly included for
 * backward compatibility and completeness.
 *
 * The Excel built-in number formats as shown in the table below:
 *
 *   | Index | Index | Format String                                        |
 *   | ----- | ----- | ---------------------------------------------------- |
 *   | 0     | 0x00  | `General`                                            |
 *   | 1     | 0x01  | `0`                                                  |
 *   | 2     | 0x02  | `0.00`                                               |
 *   | 3     | 0x03  | `#,##0`                                              |
 *   | 4     | 0x04  | `#,##0.00`                                           |
 *   | 5     | 0x05  | `($#,##0_);($#,##0)`                                 |
 *   | 6     | 0x06  | `($#,##0_);[Red]($#,##0)`                            |
 *   | 7     | 0x07  | `($#,##0.00_);($#,##0.00)`                           |
 *   | 8     | 0x08  | `($#,##0.00_);[Red]($#,##0.00)`                      |
 *   | 9     | 0x09  | `0%`                                                 |
 *   | 10    | 0x0a  | `0.00%`                                              |
 *   | 11    | 0x0b  | `0.00E+00`                                           |
 *   | 12    | 0x0c  | `# ?/?`                                              |
 *   | 13    | 0x0d  | `# ??/??`                                            |
 *   | 14    | 0x0e  | `m/d/yy`                                             |
 *   | 15    | 0x0f  | `d-mmm-yy`                                           |
 *   | 16    | 0x10  | `d-mmm`                                              |
 *   | 17    | 0x11  | `mmm-yy`                                             |
 *   | 18    | 0x12  | `h:mm AM/PM`                                         |
 *   | 19    | 0x13  | `h:mm:ss AM/PM`                                      |
 *   | 20    | 0x14  | `h:mm`                                               |
 *   | 21    | 0x15  | `h:mm:ss`                                            |
 *   | 22    | 0x16  | `m/d/yy h:mm`                                        |
 *   | ...   | ...   | ...                                                  |
 *   | 37    | 0x25  | `(#,##0_);(#,##0)`                                   |
 *   | 38    | 0x26  | `(#,##0_);[Red](#,##0)`                              |
 *   | 39    | 0x27  | `(#,##0.00_);(#,##0.00)`                             |
 *   | 40    | 0x28  | `(#,##0.00_);[Red](#,##0.00)`                        |
 *   | 41    | 0x29  | `_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)`            |
 *   | 42    | 0x2a  | `_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)`         |
 *   | 43    | 0x2b  | `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`    |
 *   | 44    | 0x2c  | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` |
 *   | 45    | 0x2d  | `mm:ss`                                              |
 *   | 46    | 0x2e  | `[h]:mm:ss`                                          |
 *   | 47    | 0x2f  | `mm:ss.0`                                            |
 *   | 48    | 0x30  | `##0.0E+0`                                           |
 *   | 49    | 0x31  | `@`                                                  |
 *
 *  @note
 *
 *  -  Numeric formats 23 to 36 are not documented by Microsoft and may differ
 *     in international versions. The listed date and currency formats may also
 *     vary depending on system settings.
 *
 *  - The dollar sign in the above format appears as the defined local currency
 *    symbol.
 *
 *  - These formats can also be set via format_set_num_format().
 */
//C     void format_set_num_format_index(lxw_format *format, uint8_t index);

/**
 * @brief Set the cell unlocked state.
 *
 * @param format Pointer to a Format instance.
 *
 * This property can be used to allow modification of a cell in a protected
 * worksheet. In Excel, cell locking is turned on by default for all
 * cells. However, it only has an effect if the worksheet has been protected
 * using the worksheet worksheet_protect() method:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_unlocked(format);
 *
 *     // Enable worksheet protection.
 *     worksheet_protect(worksheet);
 *
 *     // This cell cannot be edited.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", NULL);
 *
 *     // This cell can be edited.
 *     worksheet_write_formula(worksheet, 1, 0, "=1+2", format);
 * @endcode
 */
//C     void format_set_unlocked(lxw_format *format);

/**
 * @brief Hide formulas in a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This property is used to hide a formula while still displaying its result. This
 * is generally used to hide complex calculations from end users who are only
 * interested in the result. It only has an effect if the worksheet has been
 * protected using the worksheet write_protect() method:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_hidden(format);
 *
 *     // Enable worksheet protection.
 *     worksheet_protect(worksheet);
 *
 *     // The formula in this cell isn't visible.
 *     worksheet_write_formula(worksheet, 0, 0, "=1+2", format);
 * @endcode
 */
//C     void format_set_hidden(lxw_format *format);

/**
 * @brief Set the alignment for data in the cell.
 *
 * @param format    Pointer to a Format instance.
 * @param alignment The horizontal and or vertical alignment direction.
 *
 * This method is used to set the horizontal and vertical text alignment within a
 * cell. The following are the available horizontal alignments:
 *
 * - #LXW_ALIGN_LEFT
 * - #LXW_ALIGN_CENTER
 * - #LXW_ALIGN_RIGHT
 * - #LXW_ALIGN_FILL
 * - #LXW_ALIGN_JUSTIFY
 * - #LXW_ALIGN_CENTER_ACROSS
 * - #LXW_ALIGN_DISTRIBUTED
 *
 * The following are the available vertical alignments:
 *
 * - #LXW_ALIGN_VERTICAL_TOP
 * - #LXW_ALIGN_VERTICAL_BOTTOM
 * - #LXW_ALIGN_VERTICAL_CENTER
 * - #LXW_ALIGN_VERTICAL_JUSTIFY
 * - #LXW_ALIGN_VERTICAL_DISTRIBUTED
 *
 * As in Excel, vertical and horizontal alignments can be combined:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_align(format, LXW_ALIGN_CENTER);
 *     format_set_align(format, LXW_ALIGN_VERTICAL_CENTER);
 *
 *     worksheet_set_row(0, 30);
 *     worksheet_write_string(worksheet, 0, 0, "Some Text", format);
 * @endcode
 *
 * Text can be aligned across two or more adjacent cells using the
 * center_across property. However, for genuine merged cells it is better to
 * use the worksheet_merge_range() worksheet method.
 *
 * The vertical justify option can be used to provide automatic text wrapping
 * in a cell. The height of the cell will be adjusted to accommodate the
 * wrapped text. To specify where the text wraps use the
 * format_set_text_wrap() method.
 */
//C     void format_set_align(lxw_format *format, uint8_t alignment);

/**
 * @brief Wrap text in a cell.
 *
 * Turn text wrapping on for text in a cell.
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_text_wrap(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Some long text to wrap in a cell", format);
 * @endcode
 *
 * If you wish to control where the text is wrapped you can add newline characters
 * to the string:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_text_wrap(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "It's\na bum\nwrap", format);
 * @endcode
 *
 * Excel will adjust the height of the row to accommodate the wrapped text. A
 * similar effect can be obtained without newlines using the
 * format_set_align() function with #LXW_ALIGN_VERTICAL_JUSTIFY.
 */
//C     void format_set_text_wrap(lxw_format *format);

/**
 * @brief Set the rotation of the text in a cell.
 *
 * @param format Pointer to a Format instance.
 * @param angle  Rotation angle in the range -90 to 90 and 270.
 *
 * Set the rotation of the text in a cell. The rotation can be any angle in the
 * range -90 to 90 degrees:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_rotation(format, 30);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is rotated", format);
 * @endcode
 *
 * The angle 270 is also supported. This indicates text where the letters run from
 * top to bottom.
 */
//C     void format_set_rotation(lxw_format *format, int16_t angle);

/**
 * @brief Set the cell text indentation level.
 *
 * @param format Pointer to a Format instance.
 * @param level  Indentation level.
 *
 * This method can be used to indent text in a cell. The argument, which should be
 * an integer, is taken as the level of indentation:
 *
 * @code
 *     format1 = workbook_add_format(workbook);
 *     format2 = workbook_add_format(workbook);
 *
 *     format_set_indent(format1, 1);
 *     format_set_indent(format2, 2);
 *
 *     worksheet_write_string(worksheet, 0, 0, "This text is indented 1 level",  format1);
 *     worksheet_write_string(worksheet, 1, 0, "This text is indented 2 levels", format2);
 * @endcode
 *
 * @image html text_indent.png
 *
 * @note
 * Indentation is a horizontal alignment property. It will override any other
 * horizontal properties but it can be used in conjunction with vertical
 * properties.
 */
//C     void format_set_indent(lxw_format *format, uint8_t level);

/**
 * @brief Turn on the text "shrink to fit" for a cell.
 *
 * @param format Pointer to a Format instance.
 *
 * This method can be used to shrink text so that it fits in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *     format_set_shrink(format);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Honey, I shrunk the text!", format);
 * @endcode
 */
//C     void format_set_shrink(lxw_format *format);

/**
 * @brief Set the background fill pattern for a cell
 *
 * @param format Pointer to a Format instance.
 * @param index  Pattern index.
 *
 * Set the background pattern for a cell.
 *
 * The most common pattern is a solid fill of the background color:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID);
 *     format_set_bg_color(format, LXW_COLOR_YELLOW);
 * @endcode
 *
 * The available fill patterns are:
 *
 *    Fill Type                     | Define
 *    ----------------------------- | -----------------------------
 *    Solid                         | #LXW_PATTERN_SOLID
 *    Medium gray                   | #LXW_PATTERN_MEDIUM_GRAY
 *    Dark gray                     | #LXW_PATTERN_DARK_GRAY
 *    Light gray                    | #LXW_PATTERN_LIGHT_GRAY
 *    Dark horizontal line          | #LXW_PATTERN_DARK_HORIZONTAL
 *    Dark vertical line            | #LXW_PATTERN_DARK_VERTICAL
 *    Dark diagonal stripe          | #LXW_PATTERN_DARK_DOWN
 *    Reverse dark diagonal stripe  | #LXW_PATTERN_DARK_UP
 *    Dark grid                     | #LXW_PATTERN_DARK_GRID
 *    Dark trellis                  | #LXW_PATTERN_DARK_TRELLIS
 *    Light horizontal line         | #LXW_PATTERN_LIGHT_HORIZONTAL
 *    Light vertical line           | #LXW_PATTERN_LIGHT_VERTICAL
 *    Light diagonal stripe         | #LXW_PATTERN_LIGHT_DOWN
 *    Reverse light diagonal stripe | #LXW_PATTERN_LIGHT_UP
 *    Light grid                    | #LXW_PATTERN_LIGHT_GRID
 *    Light trellis                 | #LXW_PATTERN_LIGHT_TRELLIS
 *    12.5% gray                    | #LXW_PATTERN_GRAY_125
 *    6.25% gray                    | #LXW_PATTERN_GRAY_0625
 *
 */
//C     void format_set_pattern(lxw_format *format, uint8_t index);

/**
 * @brief Set the pattern background color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern background color.
 *
 * The format_set_bg_color() method can be used to set the background color of
 * a pattern. Patterns are defined via the format_set_pattern() method. If a
 * pattern hasn't been defined then a solid fill pattern is used as the
 * default.
 *
 * Here is an example of how to set up a solid fill in a cell:
 *
 * @code
 *     format = workbook_add_format(workbook);
 *
 *     format_set_pattern (format, LXW_PATTERN_SOLID);
 *     format_set_bg_color(format, LXW_COLOR_GREEN);
 *
 *     worksheet_write_string(worksheet, 0, 0, "Ray", format);
 * @endcode
 *
 * @image html formats_set_bg_color.png
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
//C     void format_set_bg_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the pattern foreground color for a cell.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell pattern foreground  color.
 *
 * The format_set_fg_color() method can be used to set the foreground color of
 * a pattern.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 *
 */
//C     void format_set_fg_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the cell border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell border style:
 *
 * @code
 *     format_set_border(format, LXW_BORDER_THIN);
 * @endcode 
 *
 * Individual border elements can be configured using the following functions with
 * the same parameters:
 *
 * - format_set_bottom()
 * - format_set_top()
 * - format_set_left()
 * - format_set_right()
 *
 * A cell border is comprised of a border on the bottom, top, left and right.
 * These can be set to the same value using format_set_border() or
 * individually using the relevant method calls shown above.
 *
 * The following border styles are available:
 *
 * - #LXW_BORDER_THIN
 * - #LXW_BORDER_MEDIUM
 * - #LXW_BORDER_DASHED
 * - #LXW_BORDER_DOTTED
 * - #LXW_BORDER_THICK
 * - #LXW_BORDER_DOUBLE
 * - #LXW_BORDER_HAIR
 * - #LXW_BORDER_MEDIUM_DASHED
 * - #LXW_BORDER_DASH_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT
 * - #LXW_BORDER_DASH_DOT_DOT
 * - #LXW_BORDER_MEDIUM_DASH_DOT_DOT
 * - #LXW_BORDER_SLANT_DASH_DOT
 *
 *  The most commonly used style is the `thin` style.
 */
//C     void format_set_border(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell bottom border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell bottom border style. See format_set_border() for details on the
 * border styles.
 */
//C     void format_set_bottom(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell top border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell top border style. See format_set_border() for details on the border
 * styles.
 */
//C     void format_set_top(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell left border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell left border style. See format_set_border() for details on the
 * border styles.
 */
//C     void format_set_left(lxw_format *format, uint8_t style);

/**
 * @brief Set the cell right border style.
 *
 * @param format Pointer to a Format instance.
 * @param style  Border style index.
 *
 * Set the cell right border style. See format_set_border() for details on the
 * border styles.
 */
//C     void format_set_right(lxw_format *format, uint8_t style);

/**
 * @brief Set the color of the cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * Individual border elements can be configured using the following methods with
 * the same parameters:
 *
 * - format_set_bottom_color()
 * - format_set_top_color()
 * - format_set_left_color()
 * - format_set_right_color()
 *
 * Set the color of the cell borders. A cell border is comprised of a border
 * on the bottom, top, left and right. These can be set to the same color
 * using format_set_border_color() or individually using the relevant method
 * calls shown above.
 *
 * The color should be an RGB integer value, see @ref working_with_colors.
 */
//C     void format_set_border_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the bottom cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_bottom_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the top cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_top_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the left cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_left_color(lxw_format *format, lxw_color_t color);

/**
 * @brief Set the color of the right cell border.
 *
 * @param format Pointer to a Format instance.
 * @param color  The cell border color.
 *
 * See format_set_border_color() for details on the border colors.
 */
//C     void format_set_right_color(lxw_format *format, lxw_color_t color);

//C     void format_set_diag_type(lxw_format *format, uint8_t value);
//C     void format_set_diag_color(lxw_format *format, lxw_color_t color);
//C     void format_set_diag_border(lxw_format *format, uint8_t value);
//C     void format_set_font_outline(lxw_format *format);
//C     void format_set_font_shadow(lxw_format *format);
//C     void format_set_font_family(lxw_format *format, uint8_t value);
//C     void format_set_font_charset(lxw_format *format, uint8_t value);
//C     void format_set_font_scheme(lxw_format *format, const char *font_scheme);
//C     void format_set_font_condense(lxw_format *format);
//C     void format_set_font_extend(lxw_format *format);
//C     void format_set_reading_order(lxw_format *format, uint8_t value);
//C     void format_set_theme(lxw_format *format, uint8_t value);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif /* TESTING */

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_FORMAT_H__ */
//C     #include "xlsxwriter/utility.h"
/*
 * libxlsxwriter
 *
 * Copyright 2014-2015, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */

/**
 * @file utility.h
 *
 * @brief Utility functions for libxlsxwriter.
 *
 * <!-- Copyright 2014-2015, John McNamara, jmcnamara@cpan.org -->
 *
 */

//C     #ifndef __LXW_UTILITY_H__
//C     #define __LXW_UTILITY_H__

//C     #include <stdint.h>
//C     #include "common.h"

/* Max col: $XFD\0 */
//C     #define MAX_COL_NAME_LENGTH   5

/* Max cell: $XFWD$1048576\0 */
//C     #define MAX_CELL_NAME_LENGTH  14

/* Max range: $XFWD$1048576:$XFWD$1048576\0 */
//C     #define MAX_CELL_RANGE_LENGTH (MAX_CELL_NAME_LENGTH * 2)

//C     #define EPOCH_1900            0
//C     #define EPOCH_1904            1

/**
 * @brief Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *      worksheet_write_string(worksheet, CELL("A1"), "Foo", NULL);
 *
 *      //Same as:
 *      worksheet_write_string(worksheet, 0, 0,       "Foo", NULL);
 * @endcode
 *
 * @note
 *
 * This macro shouldn't be used in performance critical situations since it
 * expands to two function calls.
 */
//C     #define CELL(cell)     lxw_get_row(cell), lxw_get_col(cell)

/**
 * @brief Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *     worksheet_set_column(worksheet, COLS("B:D"), 20, NULL, NULL);
 *
 *     // Same as:
 *     worksheet_set_column(worksheet, 1, 3,        20, NULL, NULL);
 * @endcode
 *
 */
//C     #define COLS(cols)     lxw_get_col(cols), lxw_get_col_2(cols)

/**
 * @brief Convert an Excel `A1:B2` range into a `(first_row, first_col,
 *        last_row, last_col)` sequence.
 *
 * Convert an Excel `A1:B2` range into a `(first_row, first_col, last_row,
 * last_col)` sequence.
 *
 * This is a little syntactic shortcut to help with worksheet layout.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 */
//C     #define RANGE(range)     lxw_get_row(range), lxw_get_col(range), lxw_get_row_2(range), lxw_get_col_2(range)

/** @brief Struct to represent a date and time in Excel.
 *
 * Struct to represent a date and time in Excel. See @ref working_with_dates.
 */
//C     typedef struct lxw_datetime {

    /** Year     : 1900 - 9999 */
//C         int year;
    /** Month    : 1 - 12 */
//C         int month;
    /** Day      : 1 - 31 */
//C         int day;
    /** Hour     : 0 - 23 */
//C         int hour;
    /** Minute   : 0 - 59 */
//C         int min;
    /** Seconds  : 0 - 59.999 */
//C         double sec;

//C     } lxw_datetime;

/* Create a quoted version of the worksheet name */
//C     char *lxw_quote_sheetname(char *str);

 /* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     extern "C" {
//C     #endif
/* *INDENT-ON* */

//C     void lxw_col_to_name(char *col_name, int col_num, uint8_t absolute);

//C     void lxw_rowcol_to_cell(char *cell_name, int row, int col);

//C     void lxw_rowcol_to_cell_abs(char *cell_name,
//C                                 int row,
//C                                 int col, uint8_t abs_row, uint8_t abs_col);

//C     void lxw_range(char *range,
//C                    int first_row, int first_col, int last_row, int last_col);

//C     void lxw_range_abs(char *range,
//C                        int first_row, int first_col, int last_row, int last_col);

//C     uint32_t lxw_get_row(const char *row_str);
//C     uint16_t lxw_get_col(const char *col_str);
//C     uint32_t lxw_get_row_2(const char *row_str);
//C     uint16_t lxw_get_col_2(const char *col_str);

//C     double _datetime_to_excel_date(lxw_datetime *datetime, uint8_t date_1904);

//C     char *lxw_strdup(const char *str);

//C     void lxw_str_tolower(char *str);

//C     FILE *lxw_tmpfile(void);

/* Declarations required for unit testing. */
//C     #ifdef TESTING

//C     #endif

/* *INDENT-OFF* */
//C     #ifdef __cplusplus
//C     }
//C     #endif
/* *INDENT-ON* */

//C     #endif /* __LXW_UTILITY_H__ */

//C     #define LXW_VERSION "0.1.5"

//C     #endif /* __LXW_XLSXWRITER_H__ */
