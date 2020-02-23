/**
 * BasicWorkbook.h
 * 
 * Declarations and typedefs for BasicWorkbook, a class that enables
 * the creation of multi sheet Office Open XML workbook files using
 * only a few methods. Cells containing numeric values, formulas, and
 * strings are supported.
 * 
 * Written in 2019 by Ben Tesch.
 *
 * To the extent possible under law, the author has dedicated all copyright
 * and related and neighboring rights to this software to the public domain
 * worldwide. This software is distributed without any warranty.
 * The text of the CC0 Public Domain Dedication should be reproduced at the
 * end of this file. If not, see http://creativecommons.org/publicdomain/zero/1.0/
 */

#ifndef BASIC_WORKBOOK_H_
#define BASIC_WORKBOOK_H_

#include <cinttypes>
#include <string>
#include <utility>
#include <set>
#include <vector>
#include "IttyZip.h"

namespace BasicWorkbook
{
  /**
   * These are the default row and column maximum indices
   * for a worksheet according to the ECMA376 standard.
   * A popular office software suite has the same limits.
   */
  const uint32_t MAX_ROW = 1048576u;
  const uint32_t MAX_COL = 16384u;

  /**
   * These minimum and maximum column widths (in characters)
   * are the same limits as those in a popular office
   * software suite.
   */
  const double MIN_COL_WIDTH = 0.0;
  const double MAX_COL_WIDTH = 255.0;

  /**
   * This type holds a cell reference as a pair of uint32_t numbers.
   */
  typedef struct
  {
    uint32_t row;
    uint32_t col;
  } integerref_t;

  /**
   * BasicWorkbook only supports three cell types:
   * 0: a cell containing a numeric value
   * 1: a cell containing a formula
   * 2: a cell containing a string
   */
  enum class CellType : uint8_t
  {
    NUMBER = 0u,
    FORMULA = 1u,
    STRING = 2u,
    EMPTY = 3u
  };

  /**
   * Number formats supported for NUMBER and FORMULA type cells.
   * General is the Office Open XML General cell format type
   * and also the default if another format is not specified.
   * TEXT is used for string cells
   * FIX is for fixed point
   * SCI is for scientific notation
   * PCT is for percentage; a 0.1 cell value results in 10%
   * The number after FIX, SCI, or PCT represents the number
   * of displayed places to the right of the decimal point.
   */
  enum class NumberFormat : uint8_t
  {
    GENERAL = 0u,
    TEXT    = 49u,
    FIX0    = 100u,
    FIX1    = 101u,
    FIX2    = 102u,
    FIX3    = 103u,
    FIX4    = 104u,
    FIX5    = 105u,
    FIX6    = 106u,
    FIX7    = 107u,
    FIX8    = 108u,
    FIX9    = 109u,
    FIX10   = 110u,
    FIX11   = 111u,
    FIX12   = 112u,
    FIX13   = 113u,
    FIX14   = 114u,
    FIX15   = 115u,
    FIX16   = 116u,
    SCI0    = 117u,
    SCI1    = 118u,
    SCI2    = 119u,
    SCI3    = 120u,
    SCI4    = 121u,
    SCI5    = 122u,
    SCI6    = 123u,
    SCI7    = 124u,
    SCI8    = 125u,
    SCI9    = 126u,
    SCI10   = 127u,
    SCI11   = 128u,
    SCI12   = 129u,
    SCI13   = 130u,
    SCI14   = 131u,
    SCI15   = 132u,
    SCI16   = 133u,
    PCT0    = 134u,
    PCT1    = 135u,
    PCT2    = 136u,
    PCT3    = 137u,
    PCT4    = 138u,
    PCT5    = 139u,
    PCT6    = 140u,
    PCT7    = 141u,
    PCT8    = 142u,
    PCT9    = 143u,
    PCT10   = 144u,
    PCT11   = 145u,
    PCT12   = 146u,
    PCT13   = 147u,
    PCT14   = 148u,
    PCT15   = 149u,
    PCT16   = 150u
  };

  /**
   * Possible kinds of horizontal alignment of the value
   * in the cell.
   */
  enum class HorizontalAlignment : uint8_t
  {
    GENERAL = 0u,
    LEFT    = 1u,
    CENTER  = 2u,
    RIGHT   = 3u
  };

  /**
   * Possible kinds of vertical alignment of the value
   * in the cell.
   */
  enum class VerticalAlignment : uint8_t
  {
    BOTTOM = 0u,
    CENTER = 1u,
    TOP    = 2u
  };

  /**
   * This struct holds the style information for
   * a cell.
   */
  typedef struct
  {
    NumberFormat num_format;
    HorizontalAlignment horiz_align;
    VerticalAlignment vert_align;
    bool wrap_text;
    bool bold;
  } cell_style_t;

  /**
   * Define the default cell styles.
   */
  const cell_style_t generic_style = {NumberFormat::GENERAL, HorizontalAlignment::GENERAL, VerticalAlignment::BOTTOM, false, false};
  const cell_style_t generic_string_style = {NumberFormat::TEXT, HorizontalAlignment::GENERAL, VerticalAlignment::BOTTOM, false, false};

  /**
   * This is the representation of a single cell in
   * BasicWorkbook. Both formula type cells and
   * string type cells store their value in the
   * same string.
   */
  typedef struct
  {
    integerref_t integerref;
    CellType type;
    size_t style_index;
    double num_val;
    std::string str_fml_val;
  } cell_t;

  /**
   * This type holds the starting (upper left) reference and
   * the ending (lower right) reference for a merged cell.
   */
  typedef struct
  {
    integerref_t start_ref;
    integerref_t end_ref;
  } merged_cell_t;

  struct cell_sort_compare
  {
    bool operator() (const cell_t &a, const cell_t &b) const noexcept;
  };

  struct merged_cell_sort_compare
  {
    bool operator() (const merged_cell_t &a, const merged_cell_t &b) const noexcept;
  };

  struct column_widths_sort_compare
  {
    bool operator() (const std::pair<uint32_t, double> &a, const std::pair<uint32_t, double> &b) const noexcept;
  };

  uint32_t column_to_integer(const std::string &column) noexcept(false);
  std::string integer_to_column(uint32_t decimal) noexcept(false);
  integerref_t mixedref_to_integerref(const std::string &mixedref) noexcept(false);
  std::string integerref_to_mixedref(const uint32_t row, const uint32_t col) noexcept(false);
  std::string integerref_to_mixedref(const integerref_t &integerref) noexcept(false);
  bool case_insensitive_same(const std::string &a, const std::string &b) noexcept;

  class Workbook;

  class Sheet
  {
  public:
    void add_number_cell(const uint32_t row, const uint32_t col, const double number, const cell_style_t &cell_style = generic_style) noexcept(false);
    void add_number_cell(const integerref_t &integerref, const double number, const cell_style_t &cell_style = generic_style) noexcept(false);
    void add_number_cell(const std::string &mixedref, const double number, const cell_style_t &cell_style = generic_style) noexcept(false);
    //void add_merged_number_cell(const uint32_t start_row, const uint32_t start_col, const uint32_t end_row, const uint32_t end_col, const double number, const cell_style_t &cell_style = generic_style) noexcept(false);
    //void add_merged_number_cell(const integerref_t &start_ref, const integerref_t &end_ref, const double number, const cell_style_t &cell_style = generic_style) noexcept(false);
    //void add_merged_number_cell(const std::string &start_ref, const std::string &end_ref, const double number, const cell_style_t &cell_style = generic_style) noexcept(false);
    void add_formula_cell(const uint32_t row, const uint32_t col, const std::string &formula, const cell_style_t &cell_style = generic_style) noexcept(false);
    void add_formula_cell(const integerref_t &integerref, const std::string &formula, const cell_style_t &cell_style = generic_style) noexcept(false);
    void add_formula_cell(const std::string &mixedref, const std::string &formula, const cell_style_t &cell_style = generic_style) noexcept(false);
    //void add_merged_formula_cell(const uint32_t start_row, const uint32_t start_col, const uint32_t end_row, const uint32_t end_col, const std::string &formula, const cell_style_t &cell_style = generic_style) noexcept(false);
    //void add_merged_formula_cell(const integerref_t &start_ref, const integerref_t &end_ref, const std::string &formula, const cell_style_t &cell_style = generic_style) noexcept(false);
    //void add_merged_formula_cell(const std::string &start_ref, const std::string &end_ref, const std::string &formula, const cell_style_t &cell_style = generic_style) noexcept(false);
    void add_string_cell(const uint32_t row, const uint32_t col, const std::string &value, const cell_style_t &cell_style = generic_string_style) noexcept(false);
    void add_string_cell(const integerref_t &integerref, const std::string &value, const cell_style_t &cell_style = generic_string_style) noexcept(false);
    void add_string_cell(const std::string &mixedref, const std::string &value, const cell_style_t &cell_style = generic_string_style) noexcept(false);
    void add_merged_string_cell(const uint32_t start_row, const uint32_t start_col, const uint32_t end_row, const uint32_t end_col, const std::string &value, const cell_style_t &cell_style = generic_string_style) noexcept(false);
    void add_merged_string_cell(const integerref_t &start_ref, const integerref_t &end_ref, const std::string &value, const cell_style_t &cell_style = generic_string_style) noexcept(false);
    void add_merged_string_cell(const std::string &start_ref, const std::string &end_ref, const std::string &value, const cell_style_t &cell_style = generic_string_style) noexcept(false);
    void set_column_width(const uint32_t col, const double width) noexcept(false);
    void set_column_width(const std::string &column, const double width) noexcept(false);
    std::string get_name(void) const noexcept;

  private:
    Sheet(const std::string &name_, const std::string &filename_, const uint32_t sheetId_, const std::string &relId_, Workbook &workbook_) noexcept(false);
    void add_empty_cell(const integerref_t &integerref, const cell_style_t &cell_style) noexcept(false);
    std::string generate_file(void) const noexcept;

    /**
     * Reference to the enclosing workbook.
     * Solely used to call workbook.addStyle().
     */
    Workbook &workbook;

    /**
     * The name of the Sheet as displayed on the tab used to view
     * the Sheet in a popular office software suite.
     * Must not be empty. This is the only data member besides the
     * cells that is provided by the user of BasicWorkbook (through
     * Workbook::addSheet).
     */
    std::string name;

    /**
     * The filename of the Sheet .xml file inside the Workbook
     * ZIP archive. Managed automatically by the Workbook class.
     */
    std::string filename;

    /**
     * The numeric ID of this Sheet in the Workbook. Managed
     * automatically by the Workbook class.
     */
    uint32_t sheetId;

    /**
     * The relationship ID of this Sheet in the Workbook. Managed
     * automatically by the Workbook class.
     */
    std::string relId;

    /**
     * This set stores the indices of columns with one or more
     * cells so that the width of all non-empty columns can
     * be set to best fit if not otherwise specified.
     */
    std::set<uint32_t> used_columns;

    /**
     * This set holds any custom column widths.
     * The column index is the first element of the pair;
     * the width is the second element of the pair.
     */
    std::set<std::pair<uint32_t,double>, column_widths_sort_compare> column_widths;

    /**
     * This Sheet's cells are stored in a set so that they are
     * maintained in sorted order for easier Sheet .xml file
     * generation later. This also helps detect duplicate additions
     * of a cell at the same reference.
     *
     * Cells are added through the various overloaded methods
     * Sheet::add_number_cell()
     * Sheet::add_formula_cell()
     * Sheet::add_string_cell()
     */
    std::set<cell_t, cell_sort_compare> cells;

    /**
     * Merged cell references are stored in this set because these
     * are needed to generate the Sheet .xml file.
     * Duplicate / overlapping merged cells is handled implicitly
     * by ordinary duplicate cell detection.
     */
    std::set<merged_cell_t, merged_cell_sort_compare> merged_cells;

    friend class Workbook;
  };

  class Workbook
  {
  public:
    Workbook(void) noexcept;
    Sheet& addSheet(const std::string &name) noexcept(false);
    size_t addStyle(const cell_style_t &cell_style) noexcept;
    void publish(const std::string &filename) noexcept(false);

  private:
    /**
     * All of this Workbook's sheets are stored in this vector.
     * This is not a set (which would have faster duplicate name
     * search) because the sheets are stored in the order entered,
     * and this might not be a sorted order.
     */
    std::vector<Sheet> sheets;

    /**
     * The various cell styles actually in use are stored in
     * this array so that the styles.xml file only needs to define
     * styles that are really used in this workbook.
     */
    std::vector<cell_style_t> cell_styles;

    /**
     * All Office Open XML files are stored in ZIP archives as the
     * container format. This IttyZip archive manages the conversion
     * from XML files represented as std::strings into an actual
     * ZIP archive file on disk.
     */
    IttyZip::IttyZip archive;
  };
}

#endif /* #ifndef BASIC_WORKBOOK_H_ */

/*
Creative Commons Legal Code

CC0 1.0 Universal

    CREATIVE COMMONS CORPORATION IS NOT A LAW FIRM AND DOES NOT PROVIDE
    LEGAL SERVICES. DISTRIBUTION OF THIS DOCUMENT DOES NOT CREATE AN
    ATTORNEY-CLIENT RELATIONSHIP. CREATIVE COMMONS PROVIDES THIS
    INFORMATION ON AN "AS-IS" BASIS. CREATIVE COMMONS MAKES NO WARRANTIES
    REGARDING THE USE OF THIS DOCUMENT OR THE INFORMATION OR WORKS
    PROVIDED HEREUNDER, AND DISCLAIMS LIABILITY FOR DAMAGES RESULTING FROM
    THE USE OF THIS DOCUMENT OR THE INFORMATION OR WORKS PROVIDED
    HEREUNDER.

Statement of Purpose

The laws of most jurisdictions throughout the world automatically confer
exclusive Copyright and Related Rights (defined below) upon the creator
and subsequent owner(s) (each and all, an "owner") of an original work of
authorship and/or a database (each, a "Work").

Certain owners wish to permanently relinquish those rights to a Work for
the purpose of contributing to a commons of creative, cultural and
scientific works ("Commons") that the public can reliably and without fear
of later claims of infringement build upon, modify, incorporate in other
works, reuse and redistribute as freely as possible in any form whatsoever
and for any purposes, including without limitation commercial purposes.
These owners may contribute to the Commons to promote the ideal of a free
culture and the further production of creative, cultural and scientific
works, or to gain reputation or greater distribution for their Work in
part through the use and efforts of others.

For these and/or other purposes and motivations, and without any
expectation of additional consideration or compensation, the person
associating CC0 with a Work (the "Affirmer"), to the extent that he or she
is an owner of Copyright and Related Rights in the Work, voluntarily
elects to apply CC0 to the Work and publicly distribute the Work under its
terms, with knowledge of his or her Copyright and Related Rights in the
Work and the meaning and intended legal effect of CC0 on those rights.

1. Copyright and Related Rights. A Work made available under CC0 may be
protected by copyright and related or neighboring rights ("Copyright and
Related Rights"). Copyright and Related Rights include, but are not
limited to, the following:

  i. the right to reproduce, adapt, distribute, perform, display,
     communicate, and translate a Work;
 ii. moral rights retained by the original author(s) and/or performer(s);
iii. publicity and privacy rights pertaining to a person's image or
     likeness depicted in a Work;
 iv. rights protecting against unfair competition in regards to a Work,
     subject to the limitations in paragraph 4(a), below;
  v. rights protecting the extraction, dissemination, use and reuse of data
     in a Work;
 vi. database rights (such as those arising under Directive 96/9/EC of the
     European Parliament and of the Council of 11 March 1996 on the legal
     protection of databases, and under any national implementation
     thereof, including any amended or successor version of such
     directive); and
vii. other similar, equivalent or corresponding rights throughout the
     world based on applicable law or treaty, and any national
     implementations thereof.

2. Waiver. To the greatest extent permitted by, but not in contravention
of, applicable law, Affirmer hereby overtly, fully, permanently,
irrevocably and unconditionally waives, abandons, and surrenders all of
Affirmer's Copyright and Related Rights and associated claims and causes
of action, whether now known or unknown (including existing as well as
future claims and causes of action), in the Work (i) in all territories
worldwide, (ii) for the maximum duration provided by applicable law or
treaty (including future time extensions), (iii) in any current or future
medium and for any number of copies, and (iv) for any purpose whatsoever,
including without limitation commercial, advertising or promotional
purposes (the "Waiver"). Affirmer makes the Waiver for the benefit of each
member of the public at large and to the detriment of Affirmer's heirs and
successors, fully intending that such Waiver shall not be subject to
revocation, rescission, cancellation, termination, or any other legal or
equitable action to disrupt the quiet enjoyment of the Work by the public
as contemplated by Affirmer's express Statement of Purpose.

3. Public License Fallback. Should any part of the Waiver for any reason
be judged legally invalid or ineffective under applicable law, then the
Waiver shall be preserved to the maximum extent permitted taking into
account Affirmer's express Statement of Purpose. In addition, to the
extent the Waiver is so judged Affirmer hereby grants to each affected
person a royalty-free, non transferable, non sublicensable, non exclusive,
irrevocable and unconditional license to exercise Affirmer's Copyright and
Related Rights in the Work (i) in all territories worldwide, (ii) for the
maximum duration provided by applicable law or treaty (including future
time extensions), (iii) in any current or future medium and for any number
of copies, and (iv) for any purpose whatsoever, including without
limitation commercial, advertising or promotional purposes (the
"License"). The License shall be deemed effective as of the date CC0 was
applied by Affirmer to the Work. Should any part of the License for any
reason be judged legally invalid or ineffective under applicable law, such
partial invalidity or ineffectiveness shall not invalidate the remainder
of the License, and in such case Affirmer hereby affirms that he or she
will not (i) exercise any of his or her remaining Copyright and Related
Rights in the Work or (ii) assert any associated claims and causes of
action with respect to the Work, in either case contrary to Affirmer's
express Statement of Purpose.

4. Limitations and Disclaimers.

 a. No trademark or patent rights held by Affirmer are waived, abandoned,
    surrendered, licensed or otherwise affected by this document.
 b. Affirmer offers the Work as-is and makes no representations or
    warranties of any kind concerning the Work, express, implied,
    statutory or otherwise, including without limitation warranties of
    title, merchantability, fitness for a particular purpose, non
    infringement, or the absence of latent or other defects, accuracy, or
    the present or absence of errors, whether or not discoverable, all to
    the greatest extent permissible under applicable law.
 c. Affirmer disclaims responsibility for clearing rights of other persons
    that may apply to the Work or any use thereof, including without
    limitation any person's Copyright and Related Rights in the Work.
    Further, Affirmer disclaims responsibility for obtaining any necessary
    consents, permissions or other rights required for any use of the
    Work.
 d. Affirmer understands and acknowledges that Creative Commons is not a
    party to this document and has no duty or obligation with respect to
    this CC0 or use of the Work.
*/
