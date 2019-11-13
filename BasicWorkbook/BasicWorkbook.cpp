/**
 * BasicWorkbook.cpp
 * 
 * Definitions for BasicWorkbook, a class that enables the creation of
 * multi sheet Office Open XML workbook files using only a few methods.
 * Cells containing numeric values, formulas, and strings are supported.
 * 
 * Written in 2019 by Ben Tesch.
 *
 * To the extent possible under law, the author has dedicated all copyright
 * and related and neighboring rights to this software to the public domain
 * worldwide. This software is distributed without any warranty.
 * The text of the CC0 Public Domain Dedication should be reproduced at the
 * end of this file. If not, see http://creativecommons.org/publicdomain/zero/1.0/
 */

#include <exception>
#include <stdexcept>
#include <cctype>
#include <limits>
#include <chrono>
#include <ctime>
#include "BasicWorkbook.h"

namespace BasicWorkbook
{
  /**
   * The sorting criterion here makes storing all the Sheet's
   * cells in a set practical. It sorts first by row, and then
   * within row by column.
   *
   * Since sets are always kept in sorted order, and Sheets
   * must be written in order of row, one only need to iterate
   * the set to get the cells in the desired order.
   */
  bool cell_sort_compare::operator() (const cell_t &a, const cell_t &b) const noexcept
  {
    if (a.integerref.row < b.integerref.row ||
        (a.integerref.row == b.integerref.row && 
         a.integerref.col < b.integerref.col)) return true;
    return false;
  }

  /**
   * Convert from a column index expressed as a string
   * A, B, ..., Z, AA, AB, ..., ZZ, AAA, AAB, ...
   * to the equivalent integer index, where A = 1, B = 2,
   * Z = 26, AA = 27, and so on.
   */
  uint32_t column_to_integer(const std::string &column) noexcept(false)
  {
    if (column.empty()) throw std::invalid_argument("column_to_integer() received empty input string.");

    uint32_t integer = 0u;
    for (size_t jChar = 0u; jChar < column.size()-1u; jChar++)
    {
      char this_char = std::toupper(column.at(jChar));
      if (!std::isupper(this_char)) throw std::invalid_argument("column_to_integer() received non-alphabetic input character.");
      integer = 26u * (integer + static_cast<uint32_t>(this_char) - 64u);
    }

    char last_char = std::toupper(column.at(column.size()-1u));
    if (!std::isupper(last_char)) throw std::invalid_argument("column_to_integer() received non-alphabetic input character.");
    integer += static_cast<uint32_t>(last_char) - 64u;

    if (integer > MAX_COL) throw std::out_of_range("column_to_integer() received a too large column index.");

    return integer;
  }

  /**
   * Convert a column index expressed as an integer (>= 1) to
   * the equivalent alphabetic index, where A = 1, B = 2,
   * Z = 26, AA = 27, and so on.
   */
  std::string integer_to_column(uint32_t integer) noexcept(false)
  {
    if (integer == 0u) throw std::invalid_argument("integer_to_column() received an input of 0.");
    if (integer > MAX_COL) throw std::out_of_range("integer_to_column() received a too large column index.");
    
    std::string column;

    while (integer > 0u)
    {
      uint32_t remainder = integer % 26u;
      integer--;
      integer /= 26u;
      if (remainder == 0u) column.insert(0u, 1u, 'Z');
      else column.insert(0u, 1u, static_cast<char>(remainder + 64u));
    }

    return column;
  }

  /**
   * Convert a cell reference from alphabetic column / integer row format---
   * A1, C2, DH59, etc.---to integerref format, where both row and column
   * are expressed as integers >= 1.
   */
  integerref_t mixedref_to_integerref(const std::string &mixedref) noexcept(false)
  {
    bool found_first_alpha = false;
    bool found_first_decimal = false;
    size_t row_start = std::string::npos;
    
    for (size_t jChar = 0u; jChar < mixedref.size(); jChar++)
    {
      char this_char = mixedref.at(jChar);
      if (std::isalpha(this_char))
      {
        found_first_alpha = true;
        if (found_first_decimal) throw std::invalid_argument("mixedref_to_integerref() received an invalid cell reference.");
      }
      else if (std::isdigit(this_char))
      {
        if (!found_first_alpha) throw std::invalid_argument("mixedref_to_integerref() received an invalid cell reference.");
        if (!found_first_decimal) row_start = jChar;
        found_first_decimal = true;
      }
      else
      {
        throw std::invalid_argument("mixedref_to_integerref() received an invalid cell reference.");
      }
    }

    if (!found_first_alpha || !found_first_decimal) throw std::invalid_argument("mixedref_to_integerref() received an invalid cell reference.");

    std::string column_str = mixedref.substr(0u, row_start);
    integerref_t integerref;
    integerref.col = column_to_integer(column_str);

    size_t after_int = 0u;
    int64_t row_decode = 0ll;
    try
    {
      row_decode = std::stoll(mixedref.substr(row_start), &after_int, 10);
    }
    catch (const std::invalid_argument &e)
    {
      e;
      throw std::invalid_argument("mixedref_to_integerref() received an invalid cell reference.");
    }
    catch (const std::out_of_range &e)
    {
      e;
      throw std::invalid_argument("mixedref_to_integerref() received an invalid cell reference.");
    }

    if (after_int != mixedref.size() - row_start || 
        row_decode < 1ll || 
        row_decode > static_cast<int64_t>(MAX_ROW))
    {
      throw std::invalid_argument("mixedref_to_integerref() received an invalid cell reference.");
    }

    integerref.row = static_cast<uint32_t>(row_decode);
    return integerref;
  }

  /**
   * Convert a cell reference from integerref format, where both row and column
   * are expressed as integers >= 1, to mixedref format which has alphabetic
   * column representation and integer row representation: B8, D22, AH11, etc.
   *
   * The integerref_t struct is a little inconvenient for the caller, so this
   * overload is provided to make things easier.
   */
  std::string integerref_to_mixedref(const uint32_t row, const uint32_t col) noexcept(false)
  {
    integerref_t integerref;
    integerref.row = row;
    integerref.col = col;
    return integerref_to_mixedref(integerref);
  }

  /**
   * Convert a cell reference from integerref format, where both row and column
   * are expressed as integers >= 1, to mixedref format which has alphabetic
   * column representation and integer row representation: B8, D22, AH11, etc.
   */
  std::string integerref_to_mixedref(const integerref_t &integerref) noexcept(false)
  {
    if (integerref.col < 1u || 
        integerref.col > MAX_COL ||
        integerref.row < 1u ||
        integerref.row > MAX_ROW)
    {
      throw std::invalid_argument("integerref_to_mixedref() received an invalid cell reference.");
    }

    return integer_to_column(integerref.col) + std::to_string(integerref.row);
  }

  /**
  * Returns true if a and b are the same modulo uppercase/lowercase.
  * Returns false otherwise.
  */
  bool case_insensitive_same(const std::string &a, const std::string &b) noexcept
  {
    if (a.size() != b.size()) return false;

    for (size_t jChar = 0u; jChar < a.size(); jChar++)
    {
      if (std::tolower(a.at(jChar)) != std::tolower(b.at(jChar))) return false;
    }

    return true;
  }

  /**
   * Add a cell with a numeric value to this Sheet at the specified row & column.
   * integerref_t is a little inconvenient for the caller, so this interface is
   * provided to make things easier.
   */
  void Sheet::add_number_cell(const uint32_t row, const uint32_t col, const double number) noexcept(false)
  {
    integerref_t integerref;
    integerref.row = row;
    integerref.col = col;
    this->add_number_cell(integerref, number);
  }

  /**
   * Add a cell with a numeric value to this Sheet at the specified row & column.
   */
  void Sheet::add_number_cell(const integerref_t &integerref, const double number) noexcept(false)
  {
    if (integerref.col < 1u || 
        integerref.col > MAX_COL ||
        integerref.row < 1u ||
        integerref.row > MAX_ROW)
    {
      throw std::invalid_argument("add_number_cell() received an invalid cell reference.");
    }

    cell_t cell = {0};
    cell.integerref = integerref;
    cell.type = CellType::NUMBER;
    cell.num_val = number;
    
    std::pair<std::set<cell_t,cell_sort_compare>::iterator, bool> insret = cells.insert(std::move(cell));
    if (!insret.second) throw std::exception("add_number_cell() encountered duplicate insertion of a cell at the same reference.");
  }

  /**
   * Add a cell with a numeric value to this Sheet at the specified row & column.
   * The caller may prefer to use mixedref format, especially if they are working
   * with formulas. This interface is provided to support that.
   */
  void Sheet::add_number_cell(const std::string &mixedref, const double number) noexcept(false)
  {
    integerref_t integerref = mixedref_to_integerref(mixedref);
    this->add_number_cell(integerref, number);
  }

  /**
   * Add a cell with a formula to this Sheet at the specified row & column.
   * integerref_t is a little inconvenient for the caller, so this interface is
   * provided to make things easier.
   */
  void Sheet::add_formula_cell(const uint32_t row, const uint32_t col, const std::string &formula) noexcept(false)
  {
    integerref_t integerref;
    integerref.row = row;
    integerref.col = col;
    this->add_formula_cell(integerref, formula);
  }

  /**
   * Add a cell with a formula to this Sheet at the specified row & column.
   */
  void Sheet::add_formula_cell(const integerref_t &integerref, const std::string &formula) noexcept(false)
  {
    if (integerref.col < 1u || 
        integerref.col > MAX_COL ||
        integerref.row < 1u ||
        integerref.row > MAX_ROW)
    {
      throw std::invalid_argument("add_formula_cell() received an invalid cell reference.");
    }

    cell_t cell = {0};
    cell.integerref = integerref;
    cell.type = CellType::FORMULA;
    cell.str_fml_val = formula;
    cell.num_val = std::numeric_limits<double>::quiet_NaN();

    std::pair<std::set<cell_t,cell_sort_compare>::iterator, bool> insret = cells.insert(std::move(cell));
    if (!insret.second) throw std::exception("add_formula_cell() encountered duplicate insertion of a cell at the same reference.");
  }

  /**
   * Add a cell with a formula to this Sheet at the specified row & column.
   * The caller may prefer to use mixedref format, especially if they are working
   * with formulas. This interface is provided to support that.
   */
  void Sheet::add_formula_cell(const std::string &mixedref, const std::string &formula) noexcept(false)
  {
    integerref_t integerref = mixedref_to_integerref(mixedref);
    this->add_formula_cell(integerref, formula);
  }

  /**
   * Add a cell with a string value to this Sheet at the specified row & column.
   * integerref_t is a little inconvenient for the caller, so this interface is
   * provided to make things easier.
   */
  void Sheet::add_string_cell(const uint32_t row, const uint32_t col, const std::string &value) noexcept(false)
  {
    integerref_t integerref;
    integerref.row = row;
    integerref.col = col;
    this->add_string_cell(integerref, value);
  }

  /**
   * Add a cell with a string value to this Sheet at the specified row & column.
   */
  void Sheet::add_string_cell(const integerref_t &integerref, const std::string &value) noexcept(false)
  {
    if (integerref.col < 1u || 
        integerref.col > MAX_COL ||
        integerref.row < 1u ||
        integerref.row > MAX_ROW)
    {
      throw std::invalid_argument("add_string_cell() received an invalid cell reference.");
    }

    cell_t cell = {0};
    cell.integerref = integerref;
    cell.type = CellType::STRING;
    cell.str_fml_val = value;
    cell.num_val = std::numeric_limits<double>::quiet_NaN();

    std::pair<std::set<cell_t,cell_sort_compare>::iterator, bool> insret = cells.insert(std::move(cell));
    if (!insret.second) throw std::exception("add_string_cell() encountered duplicate insertion of a cell at the same reference.");
  }


  /**
   * Add a cell with a string value to this Sheet at the specified row & column.
   * The caller may prefer to use mixedref format, especially if they are working
   * with formulas. This interface is provided to support that.
   */
  void Sheet::add_string_cell(const std::string &mixedref, const std::string &value) noexcept(false)
  {
    integerref_t integerref = mixedref_to_integerref(mixedref);
    this->add_string_cell(integerref, value);
  }

  /**
   * Retrieve the name of this Sheet; this is the name displayed
   * on the Sheet's tab in a popular office software suite.
   */
  std::string Sheet::get_name(void) const noexcept
  {
    return name;
  }

  /**
   * Private Sheet constructor called by Workbook::addSheet().
   * The thought is that the user of this program shouldn't have to manage
   * filenames, sheetIds, and relIds within the workbook. They still get to
   * name the Sheet, since name is displayed on the Sheet's tab in a
   * popular office software suite.
   */
  Sheet::Sheet(const std::string &name_, const std::string &filename_, const uint32_t sheetId_, const std::string &relId_) noexcept(false) :
    name(name_), filename(filename_), sheetId(sheetId_), relId(relId_)
  {
    /* Nothing. */
  }

  /**
   * Produces a string holding the contents of this Sheet's xml
   * file inside the actual workbook ZIP archive.
   */
  std::string Sheet::generate_file(void) const noexcept
  {
    std::string file;
    file += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    file += u8"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">";
    file += u8"<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>";
    file += u8"<sheetFormatPr defaultRowHeight=\"17\"/>";
    
    if (cells.empty())
    {
      file += u8"<sheetData/>";
    }
    else
    {
      file += u8"<sheetData>";
      uint32_t this_row = 0u;

      for (std::set<cell_t,cell_sort_compare>::const_iterator cell_itr = cells.cbegin();
           cell_itr != cells.cend();
           cell_itr++)
      {
        const cell_t &this_cell = *cell_itr;
        
        if (this_cell.integerref.row > this_row)
        {
          if (this_row > 0u) file += u8"</row>";
          this_row = this_cell.integerref.row;
          file += u8"<row r=\"" + std::to_string(this_row) + "\">";
        }
        
        std::string mixedref = integerref_to_mixedref(this_cell.integerref);

        if (this_cell.type == CellType::NUMBER)
        {
          file += u8"<c r=\"" + mixedref + "\">";
          file += u8"<v>" + std::to_string(this_cell.num_val) + "</v></c>";
        }
        else if (this_cell.type == CellType::FORMULA)
        {
          file += u8"<c r=\"" + mixedref + "\">";
          file += u8"<f>" + this_cell.str_fml_val + "</f></c>";
        }
        else
        {
          file += u8"<c r=\"" + mixedref + "\" t=\"inlineStr\">";
          file += u8"<is><t>" + this_cell.str_fml_val + "</t></is></c>";
        }
      }
      file += u8"</row></sheetData>";
    }
    
    file += u8"</worksheet>";
    return file;
  }

  /**
   * Workbook constructor. filename is the name of the file that
   * this workbook will be written to, such as Book1.xlsx.
   */
  Workbook::Workbook(const std::string filename) noexcept(false)
  {
    archive.open(filename);
  }

  /**
   * Adds a new blank Sheet to this Workbook and returns a 
   * reference to it. The reference is meant to be used by
   * the caller to add cells to the Sheet.
   *
   * name is the name of the Sheet that appears in the tab
   * that is used to view the sheet in a popular office software
   * suite.
   */
  Sheet& Workbook::addSheet(const std::string &name) noexcept(false)
  {
    if (name.empty()) throw std::invalid_argument("addsheet() received an empty name for a new sheet.");

    for (size_t jSheet = 0u; jSheet < sheets.size(); jSheet++)
    {
      if (case_insensitive_same(name, sheets.at(jSheet).get_name())) throw std::exception("addSheet() received a new sheet with the same name as an existing sheet.");
    }

    uint32_t sheetId = sheets.size() + 1u;
    std::string sheet_filename = "xl/worksheets/sheet" + std::to_string(sheetId) + ".xml";
    std::string sheet_relId = "rId" + std::to_string(sheetId + 1u);
    sheets.push_back(std::move(Sheet(name, sheet_filename, sheetId, sheet_relId)));
    return sheets.back();
  }

  void Workbook::publish(void) noexcept(false)
  {
    if (sheets.empty()) throw std::exception("publish() called, but Workbook has no Sheets.");

    {
      std::string content_types;
      content_types += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
      content_types += u8"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">";
      content_types += u8"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>";
      content_types += u8"<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
      content_types += u8"<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>";
    
      for (size_t jSheet = 0u; jSheet < sheets.size(); jSheet++)
        content_types += u8"<Override PartName=\"/" + sheets.at(jSheet).filename + "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";

      content_types += u8"<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
      content_types += u8"<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>";
      content_types += u8"<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>";
      content_types += u8"</Types>";
      archive.addFile("[Content_Types].xml", content_types);
    }

    {
      std::string rels;
      rels += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
      rels += u8"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
      rels += u8"<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>";
      rels += u8"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
      rels += u8"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>";
      rels += u8"</Relationships>";
      archive.addFile("_rels/.rels", rels);
    }

    {
      std::string app;
      app += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
      app += u8"<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">";
      app += u8"<Application>BasicWorkbook</Application>";
      app += u8"<AppVersion>1.0</AppVersion>";
      app += u8"<DocSecurity>0</DocSecurity>";
      app += u8"<ScaleCrop>false</ScaleCrop>";
      app += u8"<HeadingPairs>";
      app += u8"<vt:vector size=\"2\" baseType=\"variant\">";
      app += u8"<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>";
      app += u8"<vt:variant><vt:i4>" + std::to_string(sheets.size()) + "</vt:i4></vt:variant>";
      app += u8"</vt:vector>";
      app += u8"</HeadingPairs>";
      app += u8"<TitlesOfParts>";
      app += u8"<vt:vector size=\"" + std::to_string(sheets.size()) + "\" baseType=\"lpstr\">";

      for (size_t jSheet = 0u; jSheet < sheets.size(); jSheet++)
        app += u8"<vt:lpstr>" + sheets.at(jSheet).name + "</vt:lpstr>";

      app += u8"</vt:vector>";
      app += u8"</TitlesOfParts>";
      app += u8"<LinksUpToDate>false</LinksUpToDate>";
      app += u8"<SharedDoc>false</SharedDoc>";
      app += u8"<HyperlinksChanged>false</HyperlinksChanged>";
      app += u8"</Properties>";
      archive.addFile("docProps/app.xml", app);
    }

    {
      std::string core;
      core += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
      core += u8"<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">";
      core += u8"<dc:creator/>";
      core += u8"<cp:lastModifiedBy/>";

      std::time_t timepoint = std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
      std::tm timestruct = IttyZip::gmtime_locked(timepoint);
      const size_t TIME_BUF_SIZE = 22u;
      char timestamp[TIME_BUF_SIZE];
      size_t retval = strftime(timestamp, TIME_BUF_SIZE, "%Y-%m-%dT%H:%M:%SZ", &timestruct);
      if (retval == 0u) throw std::exception("Could not assemble timestamp string for core.xml in publish().");
      std::string time_string(timestamp);

      core += u8"<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + time_string + "</dcterms:created>";
      core += u8"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + time_string + "</dcterms:modified>";
      core += u8"</cp:coreProperties>";
      archive.addFile("docProps/core.xml", core);
    }

    {
      std::string rels;
      rels += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
      rels += u8"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
      rels += u8"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";

      for (size_t jSheet = 0u; jSheet < sheets.size(); jSheet++)
      {
        rels += u8"<Relationship Id=\"";
        rels += sheets.at(jSheet).relId;
        rels += u8"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"";
        rels += sheets.at(jSheet).filename.substr(3);
        rels += "\"/>";
      }

      rels += u8"</Relationships>";
      archive.addFile("xl/_rels/workbook.xml.rels", rels);
    }

    {
      std::string styles;
      styles += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
      styles += u8"<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">";
      styles += u8"<fonts count=\"1\"><font>";
      styles += u8"<sz val=\"12\"/>";
      styles += u8"<color rgb=\"FF000000\"/>";
      styles += u8"<name val=\"Calibri\"/>";
      styles += u8"<family val=\"2\"/>";
      styles += u8"<scheme val=\"minor\"/>";
      styles += u8"</font></fonts>";
      styles += u8"<fills count=\"1\"><fill>";
      styles += u8"<patternFill patternType=\"none\"/>";
      styles += u8"</fill></fills>";
      styles += u8"<borders count=\"1\"><border>";
      styles += u8"<left/><right/><top/><bottom/><diagonal/>";
      styles += u8"</border></borders>";
      styles += u8"<cellStyleXfs count=\"1\">";
      styles += u8"<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>";
      styles += u8"</cellStyleXfs>";
      styles += u8"<cellXfs count=\"1\">";
      styles += u8"<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>";
      styles += u8"</cellXfs>";
      styles += u8"<cellStyles count=\"1\">";
      styles += u8"<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>";
      styles += u8"</cellStyles>";
      styles += u8"<dxfs count=\"0\"/>";
      styles += u8"<tableStyles count=\"0\"/>";
      styles += u8"</styleSheet>";
      archive.addFile("xl/styles.xml", styles);
    }

    {
      std::string workbook;
      workbook += u8"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
      workbook += u8"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">";
      workbook += u8"<sheets>";
      
      for (size_t jSheet = 0u; jSheet < sheets.size(); jSheet++)
      {
        workbook += u8"<sheet name=\"";
        workbook += sheets.at(jSheet).name;
        workbook += u8"\" sheetId=\"";
        workbook += std::to_string(sheets.at(jSheet).sheetId);
        workbook += u8"\" r:id=\"";
        workbook += sheets.at(jSheet).relId;
        workbook += u8"\"/>";
      }

      workbook += u8"</sheets>";
      workbook += u8"<calcPr fullPrecision=\"1\"/>";
      workbook += u8"</workbook>";
      archive.addFile("xl/workbook.xml", workbook);
    }

    for (size_t jSheet = 0u; jSheet < sheets.size(); jSheet++)
      archive.addFile(sheets.at(jSheet).filename, sheets.at(jSheet).generate_file());

    archive.finalize();
  }
}

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
