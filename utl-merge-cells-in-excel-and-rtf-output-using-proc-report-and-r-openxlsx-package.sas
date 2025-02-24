%let pgm=utl-merge-cells-in-excel-and-rtf-output-using-proc-report-and-r-openxlsx-package;

%stop_submission;

Merge cells in excel and rtf output using proc report and r openxl package

sas ods rtf
https://tinyurl.com/24scydv6
https://github.com/rogerjdeangelis/utl-merge-cells-in-excel-and-rtf-output-using-proc-report-and-r-openxlsx-package/blob/main/want.rtf

sas ods excel
https://tinyurl.com/y626hrh8
https://github.com/rogerjdeangelis/utl-merge-cells-in-excel-and-rtf-output-using-proc-report-and-r-openxlsx-package/blob/main/want.xlsx

r openxlsx package
https://tinyurl.com/2taampp2
https://github.com/rogerjdeangelis/utl-merge-cells-in-excel-and-rtf-output-using-proc-report-and-r-openxlsx-package/blob/main/wantxl.xlsx

github
https://tinyurl.com/yc7y67kw
https://github.com/rogerjdeangelis/utl-merge-cells-in-excel-and-rtf-output-using-proc-report-and-r-openxlsx-package


     SOLUTIONS

         1 r openxlsx
         2 sas ods excel
         3 sas odsrtf
         4 ~200 excel related repos

/*********************************************************************************************************************************/
/*                                         |                                            |                                        */
/*                INPUT                    |            PROCESS                         |           OUTPUT                       */
/*                =====                    |            =======                         |           ======                       */
/*                                         |                                            |                                        */
/* d:/xls/wantxl.xlsx                      | 1 R OPENXLSX MERGE CHANGE (INPLACE)        |  d:/xls/wantxl.xlsx (in place edit)    */
/*                                         | ===================================        |                                        */
/* --------------------+                   |                                            |  --------------------+        MERGED   */
/* | A1| fx |MODALITY  |                   | %utl_rbeginx;                              |  | A1| fx |MODALITY  |           V     */
/* ------------------------------------+   | parmcards4;                                |  ------------------------------------+ */
/* [_] |   A      |  B   |  C   |  D   |   | library(openxlsx)                          |  [_] |   A      |  B   |  C   |  D   | */
/* ------------------------------------|   | wb<-loadWorkbook("d:/xls/wantxl.xlsx")     |  ------------------------------------| */
/*  1  |MODALITY  |CANCER|BENIGN|CHANGE|   | mergeCells(wb, "have",cols=4,rows=2:3)     |   1  |MODALITY  |CANCER|BENIGN|CHANGE| */
/*  -- |----------+------+------+------|   | saveWorkbook(wb                            |   -- |----------+------+------+------| */
/*  2  |Only scan | 140  | 300  | 1.6  |   |   ,"d:/xls/wantxl.xlsx"                    |   2  |Only scan | 140  | 300  |      | */
/*  -- |----------+------+------+------|   |   ,overwrite = TRUE)                       |   -- |----------+------+------+      | */
/*  3  |With LCS  | 160  | 320  | 1.6  |   | ;;;;                                       |   3  |With LCS  | 160  | 320  | 1.6  | */
/*  -- |----------+------+------+------|   | %utl_rendx;                                |   -- |----------+------+------+------| */
/* [HAVE]                                  |                                            |  [HAVE]                                */
/*                                         |                                            |                                        */
/*                                         |-------------------------------------------------------------------------------------*/
/*                                         |                                            |                                        */
/*                                         | 2 SAS ODS EXCEL                            |                                        */
/* SD1.HAVE                                | ===============                            |                                        */
/*                                         |                                            |                                        */
/* MODALITY  CANCER    BENIGN    CHANGE    | ods excel file='d:\xls\want.xlsx';         |  d:/xls/wantxl.xlsx                    */
/*                                         | proc report data=sd1.have spanrows;        |                                        */
/* Only scan   140       300       1.6     | column modality Cancer Benign change;      |  --------------------+        MERGED   */
/* With LCS    160       320       1.6     | define Cancer/display style={cellwidth=7%} |  | A1| fx |MODALITY  |           V     */
/*                                         |  "Cancerous number" CENTER;                |  ------------------------------------+ */
/*                                         | define Benign/display style={cellwidth=6%} |  [_] |   A      |  B   |  C   |  D   | */
/* options validvarname=upcase;            |  "Benign number" CENTER;                   |  ------------------------------------| */
/* libname sd1 "d:/sd1";                   | define change / order                      |   1  |MODALITY  |CANCER|BENIGN|CHANGE| */
/* data sd1.have;                          |   style={cellwidth=7% vjust=m}             |   -- |----------+------+------+------| */
/*  infile cards dlm='|';                  |   "AUC diff. (%)" CENTER;                  |   2  |Only scan | 140  | 300  |      | */
/*  input modality :$20.                   | run;                                       |   -- |----------+------+------+ 1.6  | */
/*    Cancer Benign Change;                | ods excel close;                           |   3  |With LCS  | 160  | 320  |      | */
/* cards4;                                 |                                            |   -- |----------+------+------+------| */
/* Only scan|140|300|1.6                   |                                            |  [HAVE]                                */
/* With LCS|160|320|1.6                    |                                            |                                        */
/* ;;;;                                    |                                            |                                        */
/* run;quit;                               |                                            |                                        */
/*                                         |-------------------------------------------------------------------------------------*/
/*                                         |                                            |                                        */
/*                                         | 3 SAS ODS RTF                              |                                        */
/* * CREATE d:/xls/wantxl.xlsx sheet have; | =============                              |                                        */
/*                                         |                                            |  d:\rtf\want.rtf                       */
/* %utlfkil(d:/xls/wantxl.xlsx);           | ods rtf file='d:\rtf\want.rtf';            |                                        */
/* %utl_rbeginx;                           | proc report data=sd1.have spanrows;        |  +-------------------------------+     */
/* parmcards4;                             | column modality Cancer Benign change;      |  |MODALITY  |CANCER|BENIGN|CHANGE|     */
/* library(openxlsx)                       | define Cancer/display style={cellwidth=7%} |  |----------+------+------+------|     */
/* library(haven)                          |  "Cancerous number" CENTER;                |  |Only scan | 140  | 300  |      |     */
/* have<-read_sas(                         | define Benign/display style={cellwidth=6%} |  |----------+------+------+ 1.6  |     */
/*   "d:/sd1/have.sas7bdat")               |  "Benign number" CENTER;                   |  |With LCS  | 160  | 320  |      |     */
/* wb <- createWorkbook()                  | define change  / order                     |  +----------+------+------+------+     */
/* addWorksheet(wb, "have")                |   style={cellwidth=7% vjust=m}             |                                        */
/* writeData(wb, sheet="have",x=have)      |   "AUC diff. (%)" CENTER;                  |                                        */
/* saveWorkbook(                           | run;                                       |                                        */
/*     wb                                  | ods rtf close;                             |                                        */
/*    ,"d:/xls/wantxl.xlsx"                |                                            |                                        */
/*    ,overwrite=TRUE)                     |                                            |                                        */
/* ;;;;                                    |                                            |                                        */
/* %utl_rendx;                             |                                            |                                        */
/*                                         |                                            |                                        */
/*********************************************************************************************************************************/

REPO
--------------------------------------------------------------------------------------------------------------------------------------------------
https://github.com/rogerjdeangelis/excel-how-do-I-remove-troublesome-characters-before-importing
https://github.com/rogerjdeangelis/ods_excel_does_not_always_honor_start_at--bug
https://github.com/rogerjdeangelis/utl-Delete-all-files-in-a-directory-with-a-specified-extension-ie-delete-excel-files
https://github.com/rogerjdeangelis/utl-Import-excel-sheet-as-character-fixing-truncation-mixed-type-columns-and-appending-issues
https://github.com/rogerjdeangelis/utl-Import-the-datepart-of-an-excel-datetime-formatted-columns
https://github.com/rogerjdeangelis/utl-Password-protect-an-EXCEL-file-in-sas-without-X-command
https://github.com/rogerjdeangelis/utl-add-a-tab-to-excel-that-autmatically-impot-a-sas-datasets-www.colectica-com
https://github.com/rogerjdeangelis/utl-add-monthly-worksheets-to-an-existing-yearly-excel-workbook
https://github.com/rogerjdeangelis/utl-adding-a-password-to-an-existing-excel-workbook
https://github.com/rogerjdeangelis/utl-adding-a-second-ods-excel-created-sheet-to-a-closed-ods-excel-workbook
https://github.com/rogerjdeangelis/utl-adding-a-sheet-to-an-existing-open-and-on-screen-or-saved-excel-workbook
https://github.com/rogerjdeangelis/utl-appending-records-to-an-existing-excel-sheet
https://github.com/rogerjdeangelis/utl-apply-excel-styling-across-multiple-spreadsheets-using-openxlsx-in-r
https://github.com/rogerjdeangelis/utl-applying-meta-data-and-importing-data-from-an-excel-named-range
https://github.com/rogerjdeangelis/utl-avoid-storing-numbers-as-text-when-exporting-mixed-type-to-excel
https://github.com/rogerjdeangelis/utl-calculate-percentage-by-group-in-wps-r-python-excel-sql-no-sql
https://github.com/rogerjdeangelis/utl-casting-and-reformatting-excel-data-before-importing
https://github.com/rogerjdeangelis/utl-clear-named-and-unnamed-cell-ranges-in-excel
https://github.com/rogerjdeangelis/utl-combine-text-in-an-excel-column-down-multiple-rows-by-group
https://github.com/rogerjdeangelis/utl-concatenating-thirty-seven-excel-tabs-while-correcting-column-types-and-using-longest-length
https://github.com/rogerjdeangelis/utl-convert-excel-to-csv-by-dropping-down-to-r-or-python
https://github.com/rogerjdeangelis/utl-copy-all-sas-datasets-in-work-library-to-tabs-in-one-excel-workbook
https://github.com/rogerjdeangelis/utl-create-a-pdf-excel-html-proc-report-with-greek-letters
https://github.com/rogerjdeangelis/utl-create-graphs-in-excel-using-excel-chart-templates
https://github.com/rogerjdeangelis/utl-creating-a-two-by-two-grid-of-reports-in-excel
https://github.com/rogerjdeangelis/utl-creating-multiple-odbc-tables-in-a-one-excel-sheet
https://github.com/rogerjdeangelis/utl-do-not-add-data-transformations-to-create-csv-files-from-excel-or-any-other-data-structure
https://github.com/rogerjdeangelis/utl-does-the-excel-named-range-table-exist
https://github.com/rogerjdeangelis/utl-does-the-excel-sheet-exist
https://github.com/rogerjdeangelis/utl-drop-down-to-powershell-and-programatically-create-an-odbc-data-source-for-excel-wps-r-rodbc
https://github.com/rogerjdeangelis/utl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates
https://github.com/rogerjdeangelis/utl-excel-changing-cell-contents-inside-proc-report
https://github.com/rogerjdeangelis/utl-excel-fixing-bad-formatting-using-passthru
https://github.com/rogerjdeangelis/utl-excel-grid-of-four-reports-in-one-sheet
https://github.com/rogerjdeangelis/utl-excel-hiding-columns-height-and-weight-in-sheet-class
https://github.com/rogerjdeangelis/utl-excel-hiperlinks-click-on-your-favorite-baseball-player-and-a-google-search-will-pop-up
https://github.com/rogerjdeangelis/utl-excel-import-individual-cells
https://github.com/rogerjdeangelis/utl-excel-import-number-strings-with-spaces
https://github.com/rogerjdeangelis/utl-excel-report-with-two-side-by-side-graphs-below_python
https://github.com/rogerjdeangelis/utl-excel-use-the-name-of-the-last-variable-in-the-pdv-for-sheet-name
https://github.com/rogerjdeangelis/utl-excel-using-proc-report-workarea-columns-to-operate--on-arbitrary-row
https://github.com/rogerjdeangelis/utl-extract-sheet-names-from-multiple-excel-versions-using-r
https://github.com/rogerjdeangelis/utl-extracting-hyperlinks-from-an-excel-sheet-python
https://github.com/rogerjdeangelis/utl-find-out-which-excel-columns-are-dates-and-assign-date-type
https://github.com/rogerjdeangelis/utl-fix-excel-columns-with-mutiple-datatypes-on-the-excel-side-using-ms-sql-and-passthru
https://github.com/rogerjdeangelis/utl-fix-excel-date-fields-on-the-excel-side-using-ms-sql-and-passthru
https://github.com/rogerjdeangelis/utl-force-excel-to-read-all-the-columns-as-numeric-or-character
https://github.com/rogerjdeangelis/utl-formatting-ai-seacrh-output-in-pdf-rtf-and-excel-format-perplexity-chatGPT-results
https://github.com/rogerjdeangelis/utl-get-the-color-of-a-cell-in-excel-xlsx
https://github.com/rogerjdeangelis/utl-highlight-existing-cells-in-excel-sheet2-that-correspond-to-cells-in-sheet1-with-specified-value
https://github.com/rogerjdeangelis/utl-highlite-sas-dataset-and-view-the-table-in-excel-without-sas-access
https://github.com/rogerjdeangelis/utl-how-to-check-whether-a-student-is-in-the-Excel-sheet-class
https://github.com/rogerjdeangelis/utl-how-to-load-excel-sheets-with-sheet-names-with-31-characters
https://github.com/rogerjdeangelis/utl-import-a-messy-excel-file
https://github.com/rogerjdeangelis/utl-import-all-excel-columns-as-character
https://github.com/rogerjdeangelis/utl-import-all-excel-columns-as-character-three-solutions
https://github.com/rogerjdeangelis/utl-import-all-excel-dates-as-character-strings-and-convert-back-to-SAS-dates
https://github.com/rogerjdeangelis/utl-import-all-excel-workskkets-and-named-ranges--in-all-workbooks-in-a-directory
https://github.com/rogerjdeangelis/utl-import-csv-file-to-excel-with-leading-zeros
https://github.com/rogerjdeangelis/utl-import-excel-when-column-names-are-excel-dates
https://github.com/rogerjdeangelis/utl-import-excel-workbooks-in-all-folders-and-subfolders
https://github.com/rogerjdeangelis/utl-importing-excel-datetime-values-in-xlsx-and-xlsx-workbooks
https://github.com/rogerjdeangelis/utl-importing-excel-string_of_32-thousand-characters-SAS-XLConnect
https://github.com/rogerjdeangelis/utl-importing-excel-when-sheetname-has-spaces
https://github.com/rogerjdeangelis/utl-importing-inconsistently-formatted-excel-dates-numeric-and-charater-in-the-same-column
https://github.com/rogerjdeangelis/utl-importing-multiple-excel-worksheets-without-access-to-pc-files
https://github.com/rogerjdeangelis/utl-in-palce-updates-to-an-existing-shared-excel-workbook
https://github.com/rogerjdeangelis/utl-join-a-sas-table-with-an-excel-table-when-column-names-that-are-dates
https://github.com/rogerjdeangelis/utl-keep-orginal-SAS-table-but-mask-excel-output
https://github.com/rogerjdeangelis/utl-keeping-leading-and-trailing-zeros-in-character-fields-with-ods-excel-output
https://github.com/rogerjdeangelis/utl-layout-ods-excel-reports-in-a-grid
https://github.com/rogerjdeangelis/utl-load-and-extract-ms-excel-document-properties-metadata
https://github.com/rogerjdeangelis/utl-manipulate-excel-directly-using-passthru-microsoft-sql-wps-r-rodbc
https://github.com/rogerjdeangelis/utl-no-need-for-sql-or-sort-merge-use-a-elegant-hash-excel-vlookup
https://github.com/rogerjdeangelis/utl-ods-excel-color-code-every-other-column-in-a-specified-row
https://github.com/rogerjdeangelis/utl-ods-excel-hilite-diagonal-cells
https://github.com/rogerjdeangelis/utl-ods-excel-update-excel-sheet-in-place-python
https://github.com/rogerjdeangelis/utl-ods-export-sas-table-to-excel-with-rotated-column-headers
https://github.com/rogerjdeangelis/utl-pivot-excel-columns-and-output-a-database-table
https://github.com/rogerjdeangelis/utl-pivot-transpose-an-excel-sheet-with-columns-that-are-excel-dates
https://github.com/rogerjdeangelis/utl-posting-your-problem-with-an-ascii-image-that-looks-just-like-an-excel-shee
https://github.com/rogerjdeangelis/utl-preserving-excel-formatting-when-writing-to-an-existing-worksheet
https://github.com/rogerjdeangelis/utl-programatically-downlaod-an-excel-file-from-the-web
https://github.com/rogerjdeangelis/utl-programatically-execute-excel-vba-macro-using-sas-python
https://github.com/rogerjdeangelis/utl-programatically-search-all-cells-in-an-excel-sheet-for-an-arbitrary-string-python-openxl
https://github.com/rogerjdeangelis/utl-remove-sheet-from-excel-workbook
https://github.com/rogerjdeangelis/utl-remove-sheet-from-existing-excel-worksheet-unix-and-windows-R
https://github.com/rogerjdeangelis/utl-rename-excel-columns-to-common-names-before-creating-sas-tables
https://github.com/rogerjdeangelis/utl-renaming-duplicate-excel-column-names-before-importing
https://github.com/rogerjdeangelis/utl-safe-way-import-excel-time-value
https://github.com/rogerjdeangelis/utl-safely-sending-dates-or-datetimes-back-and-forth-to-excel
https://github.com/rogerjdeangelis/utl-sas-ods-bidirectional-hyperlinked-table-of-contents-in-ods-pdf-html-and-excel
https://github.com/rogerjdeangelis/utl-sas-ods-excel-to-create-excel-report-and-separate-png-graph-finally-r-for-layout-in-excel
https://github.com/rogerjdeangelis/utl-sas-to-and-from-sqllite-excel-ms-access-spss-stata-using-r-packages-without-sas
https://github.com/rogerjdeangelis/utl-select-the-diagonal-values-from-a-dataset-in-excel-r-wps-python
https://github.com/rogerjdeangelis/utl-select-the-top-ten-rows-from-excel-table-without-importing-to-sas
https://github.com/rogerjdeangelis/utl-select-type-and-length-using-odbc-excel-passthru-query
https://github.com/rogerjdeangelis/utl-send-all-tables-in-a-sas-library-to-excel
https://github.com/rogerjdeangelis/utl-sending-a-formula-to-excel-to-reference-a-cell-in-another-sheet
https://github.com/rogerjdeangelis/utl-seven-algorithms-to-convert-a-sas-dataset-to-an-excel-workbook
https://github.com/rogerjdeangelis/utl-side-by-side-proc-report-output-in-pdf-html-and-excel
https://github.com/rogerjdeangelis/utl-side-by-side-reports-within-arbitrary-positions-in-one-excel-sheet-wps-r
https://github.com/rogerjdeangelis/utl-side-by-side-sas-tables-in-one-excel-sheet
https://github.com/rogerjdeangelis/utl-simple-r-code-to-covert-excel-to-sas-and-sas-to-excel
https://github.com/rogerjdeangelis/utl-simple-three-letter-commands-to-format-perplexity-AI-results-for-word-pdf-text-and-excel
https://github.com/rogerjdeangelis/utl-single-click-and-eight-excel-tabs-are-converted-to-csv-files
https://github.com/rogerjdeangelis/utl-skilled-nursing-cost-reports-2011-2019-in-excel
https://github.com/rogerjdeangelis/utl-subset-a-database-table-based-on-a-list-of-names-in-excel
https://github.com/rogerjdeangelis/utl-substituting-name-and-label-to-column-headings-in-excel
https://github.com/rogerjdeangelis/utl-tables-to-specific-excel-cells
https://github.com/rogerjdeangelis/utl-update-an-excel-workbook-in-place
https://github.com/rogerjdeangelis/utl-update-an-existing-excel-named-range-R-python-sas
https://github.com/rogerjdeangelis/utl-update-existing-excel-sheet-in-place-using-r-dcom-client
https://github.com/rogerjdeangelis/utl-using-column-position-instead-of-excel-column-names-due-to-misspellings-sas-r-python
https://github.com/rogerjdeangelis/utl-using-excel-to-get-a-usefull-proc-tabulate-output-table
https://github.com/rogerjdeangelis/utl-using-only-r-openxlsx-to-add-excel-formulas-to-an-existing-sheet
https://github.com/rogerjdeangelis/utl-using-proc-odstext-to-add-documentation-tabs-to-your-excel-workbook
https://github.com/rogerjdeangelis/utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc
https://github.com/rogerjdeangelis/utl-very-fast-summation-of-ll6-columns-in-excel-without-importing-to-sheet
https://github.com/rogerjdeangelis/utl-within-sql-join-a-a-text-or-excel-file-to-a-sas-or-foreign-table-without-proc-import
https://github.com/rogerjdeangelis/utl-wps-create-a-pie-chart-in-excel-using-wps-proc-gchart
https://github.com/rogerjdeangelis/utl_1130pm_batch_SAS_job_that_imports_an_excel_if_modified_that_day
https://github.com/rogerjdeangelis/utl_adding_SAS_graphics_at_an_arbitrary_position_into_existing_excel_sheets
https://github.com/rogerjdeangelis/utl_convert_a_sas_dataset_to_excel_without_sas_or_excel_only_need_R
https://github.com/rogerjdeangelis/utl_create_many_excel_workbooks_for_selected_cities_in_sashelp_zipcode_with_logging
https://github.com/rogerjdeangelis/utl_creating_www_hyperlinks_in_ods_excel
https://github.com/rogerjdeangelis/utl_excel-copying-a-worlbook-with-many-named-ranges-into-sas-tables
https://github.com/rogerjdeangelis/utl_excel_Import_and_transpose_range_A9-Y97_using_only_one_procedure
https://github.com/rogerjdeangelis/utl_excel_add_formula_inplace
https://github.com/rogerjdeangelis/utl_excel_add_formulas
https://github.com/rogerjdeangelis/utl_excel_add_sheet
https://github.com/rogerjdeangelis/utl_excel_add_to_sheet
https://github.com/rogerjdeangelis/utl_excel_combining_sheets_without_common_names_types_lengths
https://github.com/rogerjdeangelis/utl_excel_create_a_sheet_for_each_table_with_variable_name_position_and_label
https://github.com/rogerjdeangelis/utl_excel_create_sql_insert_and_value_statements_to_update_databases
https://github.com/rogerjdeangelis/utl_excel_determine_type_length
https://github.com/rogerjdeangelis/utl_excel_experimenting-with-the-new-ods-excel-destination
https://github.com/rogerjdeangelis/utl_excel_exporting_data_with_leading_zeros
https://github.com/rogerjdeangelis/utl_excel_fix_type_length_on_import
https://github.com/rogerjdeangelis/utl_excel_highlight_individual_cells_based_on_indicator_variables
https://github.com/rogerjdeangelis/utl_excel_import_all_columns_as_character_and_preserve_long_variable_names
https://github.com/rogerjdeangelis/utl_excel_import_data_from_a_xlsx_file_where_first_2_rows_are_header
https://github.com/rogerjdeangelis/utl_excel_import_entire_directory
https://github.com/rogerjdeangelis/utl_excel_import_long_colnames
https://github.com/rogerjdeangelis/utl_excel_import_only_female_students
https://github.com/rogerjdeangelis/utl_excel_import_sas_functions_fail_on_cells_with_mutiple_line_breaks
https://github.com/rogerjdeangelis/utl_excel_import_sub_rectangle
https://github.com/rogerjdeangelis/utl_excel_import_two_excel_ranges_within_one_sheet
https://github.com/rogerjdeangelis/utl_excel_import_xlsm_to_sas_dataset
https://github.com/rogerjdeangelis/utl_excel_importing_unicode_and_other_special_characters_without_changing_sas_encoding
https://github.com/rogerjdeangelis/utl_excel_merge_two-sheets
https://github.com/rogerjdeangelis/utl_excel_reading_a_single_cell
https://github.com/rogerjdeangelis/utl_excel_sas_wps_r_import_xlsx_without_sas_access_to_pc_files
https://github.com/rogerjdeangelis/utl_excel_update_inplace
https://github.com/rogerjdeangelis/utl_excel_update_rectangle
https://github.com/rogerjdeangelis/utl_excel_update_xlsm_workbook_using_SAS_dataset
https://github.com/rogerjdeangelis/utl_excel_updating-named-ranged-cells
https://github.com/rogerjdeangelis/utl_excel_using_a_cell_value_for_the_name_of_sas_dataset
https://github.com/rogerjdeangelis/utl_excel_using_byval_sex_and_sheet_interval_bygroups_to_create_multiple_worksheets
https://github.com/rogerjdeangelis/utl_fix_excel_column_names_before_import
https://github.com/rogerjdeangelis/utl_how_to_get_data_from_excel_file_into_wps_sas_procedure
https://github.com/rogerjdeangelis/utl_import_all_excel_workbooks_created_in_the_previous_seven_days
https://github.com/rogerjdeangelis/utl_import_data_from_excel_sheet_with_headers_and_footers_without_specifying_range_option
https://github.com/rogerjdeangelis/utl_import_excel_column_names_that_contain_a_dollar_sign_and_rename_without
https://github.com/rogerjdeangelis/utl_import_excel_unicode
https://github.com/rogerjdeangelis/utl_importing_three_excel_tables_that_are_in_one_sheet
https://github.com/rogerjdeangelis/utl_joining_and_updating_excel_sheets_without_importing_data
https://github.com/rogerjdeangelis/utl_maintaining_all_significant_digits_when_importing_excel_sheet
https://github.com/rogerjdeangelis/utl_maintaining_numeric_significance_when_exporting_and_importing_excel_workbooks
https://github.com/rogerjdeangelis/utl_ods_excel_conditionaly_higlight_individua1_cells
https://github.com/rogerjdeangelis/utl_ods_excel_create_a_table_of_contents_with_links_to_and_from_each_sheet
https://github.com/rogerjdeangelis/utl_ods_excel_font_size_and_justification_proc_report_titles_formatting
https://github.com/rogerjdeangelis/utl_ods_excel_merging_cells_after_column_header_and_before_column_names
https://github.com/rogerjdeangelis/utl_passthru_to_excel_to_fix_column_names
https://github.com/rogerjdeangelis/utl_proc_import_columns_as_character_from_excel_linux_or_windows
https://github.com/rogerjdeangelis/utl_programatically_execute_excel_macro_using_wps_proc_python
https://github.com/rogerjdeangelis/utl_put_excel_sheetnames_into_sas_macro_variable
https://github.com/rogerjdeangelis/utl_renaming_duplicate_excel_columns_to_avoid_name_collisions_when_importing
https://github.com/rogerjdeangelis/utl_sas_v5_transport_file_to_excel
https://github.com/rogerjdeangelis/utl_side_by_side_excel_reports
https://github.com/rogerjdeangelis/utl_stacking-strings-in-one-excel-cell-using-ods-excel-newline-carriage-return-line-feed-tags
https://github.com/rogerjdeangelis/utl_table_of_contents_with_excel_links_to_sheets


/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
