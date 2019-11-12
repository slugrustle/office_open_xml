/**
 * BasicWorkbook.h
 * 
 * TODO: proper file header.
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
    STRING = 2u
  };

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
    double num_val;
    std::string str_fml_val;
  } cell_t;

  struct cell_sort_compare
  {
    bool operator() (const cell_t &a, const cell_t &b) const noexcept;
  };

  uint32_t column_to_integer(const std::string &column) noexcept(false);
  std::string integer_to_column(uint32_t decimal) noexcept(false);
  integerref_t mixedref_to_integerref(const std::string &mixedref) noexcept(false);
  std::string integerref_to_mixedref(const uint32_t row, const uint32_t col) noexcept(false);
  std::string integerref_to_mixedref(const integerref_t &integerref) noexcept(false);
  bool case_insensitive_same(const std::string &a, const std::string &b) noexcept;

  class Sheet
  {
  public:
    void add_number_cell(const uint32_t row, const uint32_t col, const double number) noexcept(false);
    void add_number_cell(const integerref_t &integerref, const double number) noexcept(false);
    void add_number_cell(const std::string &mixedref, const double number) noexcept(false);
    void add_formula_cell(const uint32_t row, const uint32_t col, const std::string &formula) noexcept(false);
    void add_formula_cell(const integerref_t &integerref, const std::string &formula) noexcept(false);
    void add_formula_cell(const std::string &mixedref, const std::string &formula) noexcept(false);
    void add_string_cell(const uint32_t row, const uint32_t col, const std::string &value) noexcept(false);
    void add_string_cell(const integerref_t &integerref, const std::string &value) noexcept(false);
    void add_string_cell(const std::string &mixedref, const std::string &value) noexcept(false);
    std::string get_name(void) const noexcept;

  private:
    Sheet(const std::string &name_, const std::string &filename_, const uint32_t sheetId_, const std::string &relId_) noexcept(false);
    std::string generate_file(void) const noexcept;

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

    friend class Workbook;
  };

  class Workbook
  {
  public:
    Workbook(const std::string filename) noexcept(false);
    Sheet& addSheet(const std::string &name) noexcept(false);
    void publish(void) noexcept(false);

  private:
    /**
     * All of this Workbook's sheets are stored in this vector.
     * Compared to a set, duplicate name search is O(n) instead
     * of O(log n), but I don't anticipate very many sheets compared
     * to cells, and the case-insensitive duplicate name comparison
     * is easier to write in a linear search over a vector.
     * Basically, I'm punting on this one.
     */
    std::vector<Sheet> sheets;

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
