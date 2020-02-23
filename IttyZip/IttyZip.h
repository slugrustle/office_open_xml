/**
 * IttyZip.h
 * 
 * Declarations and typedefs for IttyZip, a class that generates
 * ZIP archive files from C++ strings. It stays lightweight by
 * foregoing compression.
 * 
 * Written in 2019 by Ben Tesch.
 * Originally distributed at https://github.com/slugrustle/office_open_xml
 *
 * To the extent possible under law, the author has dedicated all copyright
 * and related and neighboring rights to this software to the public domain
 * worldwide. This software is distributed without any warranty.
 * The text of the CC0 Public Domain Dedication should be reproduced at the
 * end of this file. If not, see http://creativecommons.org/publicdomain/zero/1.0/
 */

#ifndef ITTY_ZIP_H_
#define ITTY_ZIP_H_

#include <string>
#include <cinttypes>
#include <fstream>
#include <utility>
#include <exception>
#include <set>
#include <ctime>

namespace IttyZip
{
  /**
   * Messages for the "what()" in exceptions thrown by IttyZip
   */
  const char CANNOT_OPEN_MESG[]      = "IttyZip cannot open the output file for writing.";
  const char DOUBLE_OPEN_MESG[]      = "IttyZip::open() was called with an output file already open.";
  const char NOT_OPENED_MESG[]       = "IttyZip::addFile() or finalize() called either before the output file has been opened or after it has been closed.";
  const char UNEXPECTED_CLOSE_MESG[] = "IttyZip exception: The output file closed unexpectedly.";
  const char OUTPUT_FAIL_MESG[]      = "IttyZip exception: The output stream failed.";
  const char EMPTY_FINALIZE_MESG[]   = "IttyZip::finalize() was called on an empty IttyZip object.";
  const char DUPLICATE_FILE_MESG[]   = "IttyZip::addFile() was called twice with the same filename.";

  /**
   * Struct to hold a standard DOS format time + date stamp.
   * Note that this format is still around in 2019.
   */
  typedef struct
  {
    uint16_t time;
    uint16_t date;
  } dostimedate_t;

  /**
   * Struct to hold the local file header of a file
   * in a ZIP archive.
   */
  typedef struct
  {
    uint32_t signature;
    uint16_t extract_version;
    uint16_t general_bit_flag;
    uint16_t compression_method;
    dostimedate_t file_mod_timedate;
    uint32_t crc32;
    uint32_t size_compressed;
    uint32_t size_uncompressed;
    uint16_t filename_length;
    uint16_t extra_field_length;
    std::string filename;
  } localheader_t;

  /**
   * Struct to hold the central directory file header
   * of a file in a ZIP archive.
   */
  typedef struct
  {
    uint32_t signature;
    uint16_t version_made_by;
    uint16_t extract_version;
    uint16_t general_bit_flag;
    uint16_t compression_method;
    dostimedate_t file_mod_timedate;
    uint32_t crc32;
    uint32_t size_compressed;
    uint32_t size_uncompressed;
    uint16_t filename_length;
    uint16_t extra_field_length;
    uint16_t comment_length;
    uint16_t disk_number_start;
    uint16_t internal_attributes;
    uint32_t external_attributes;
    uint32_t local_header_offset;
    std::string filename;
  } dirheader_t;

  /**
   * Struct to hold a ZIP archive end of central directory
   * record.
   */
  typedef struct
  {
    uint32_t signature;
    uint16_t disk_number;
    uint16_t dir_start_disk_number;
    uint16_t this_disk_entries;
    uint16_t total_entries;
    uint32_t central_dir_size;
    uint32_t central_dir_offset;
    uint16_t comment_length;
  } endrecord_t;

  std::tm localtime_locked(const std::time_t &timepoint) noexcept;
  std::tm gmtime_locked(const std::time_t &timepoint) noexcept;
  dostimedate_t dosTimeDate(void) noexcept;
  uint32_t crc32(const std::string &data) noexcept;
  void uint16_to_buffer(const uint16_t in, char *out) noexcept;
  void uint32_to_buffer(const uint32_t in, char *out) noexcept;

  class IttyZip 
  {
  public:
    IttyZip(void) noexcept;
    IttyZip(const std::string &outputFilename) noexcept(false);
    void open(const std::string &outputFilename) noexcept(false);
    void addFile(const std::string &filename, const std::string &contents) noexcept(false);
    void finalize(void) noexcept(false);

  private:
    std::pair<localheader_t, dirheader_t> generateHeaders(const std::string &filename, const uint32_t file_size, const uint32_t file_crc32) const noexcept;
    uint32_t writeLocalheader(const localheader_t &localheader) noexcept(false);
    void storeDirheader(const dirheader_t &dirheader) noexcept;
    endrecord_t generateEndRecord(void) const noexcept;
    void writeEndRecord(const endrecord_t &end_record) noexcept(false);

    /**
     * The number of files already stored in this IttyZip archive.
     */
    uint16_t num_files;

    /**
     * An ofstream for writing to the output ZIP file.
     */
    std::ofstream out_file;

    /**
     * True if an output file has been opened and finalize()
     * has not yet been called. addFile() may be called when
     * opened is true. finalize() may be called if at least
     * one file has been added to this IttyZip object and
     * opened is true.
     */
    bool opened;

    /**
     * Offset in bytes from the start of the output file
     * at which the next local header (or at which the central
     * directory) will be written.
     */
    uint32_t next_offset;

    /**
     * Temporary storage for the ZIP archive central directory,
     * since this part of the file is written at the very end.
     */
    std::string central_directory;

    /**
     * Set containing the full filenames of all files previously
     * added to the IttyZip archive. Purely used to check for
     * duplicate files.
     */
    std::set<std::string> filenames;
  };
}

#endif /* #ifndef ITTY_ZIP_H_ */

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
