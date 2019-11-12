/**
 * IttyZip.cpp
 * 
 * Definitions for IttyZip, a class that generates ZIP archive
 * files from C++ strings. It stays lightweight by foregoing 
 * compression.
 * 
 * Written in 2019 by Ben Tesch.
 *
 * To the extent possible under law, the author has dedicated all copyright
 * and related and neighboring rights to this software to the public domain
 * worldwide. This software is distributed without any warranty.
 * The text of the CC0 Public Domain Dedication should be reproduced at the
 * end of this file. If not, see http://creativecommons.org/publicdomain/zero/1.0/
 */

#include "IttyZip.h"
#include <chrono>
#include <ctime>
#include <algorithm>
#include <thread>
#include <mutex>

namespace IttyZip
{

  /* mutex used by localtime_locked and gmtime_locked */
  static std::mutex time_mutex;

  /**
   * localtime is not thread safe, and neither localtime_s nor localtime_r
   * is portable. Recommend using localtime_locked everywhere (or an
   * equivalent replacement here) if IttyZip is incorporated into a
   * multithreaded program.
   */
  std::tm localtime_locked(const std::time_t &timepoint) noexcept
  {
    std::lock_guard<std::mutex> localtime_lock(time_mutex);
    std::tm timestruct = *std::localtime(&timepoint);
    return timestruct;
  }

  /**
   * gmtime is not thread safe, and neither gmtime_s nor gmtime_r
   * is portable. Recommend using gmtime_locked everywhere (or an
   * equivalent replacement here) if IttyZip is incorporated into a
   * multithreaded program.
   */
  std::tm gmtime_locked(const std::time_t &timepoint) noexcept
  {
    std::lock_guard<std::mutex> gmtime_lock(time_mutex);
    std::tm timestruct = *std::gmtime(&timepoint);
    return timestruct;
  }

  /**
   * Creates a DOS type time + date stamp from
   * the present system time + date.
   * Calls localtime_locked.
   */
  dostimedate_t dosTimeDate(void) noexcept
  {
    std::time_t timepoint = std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
    std::tm timestruct = localtime_locked(timepoint);
    dostimedate_t output;

    output.time = 0x001Fu & (std::min(timestruct.tm_sec, 59) / 2);
    output.time |= 0x07E0u & (timestruct.tm_min << 5);
    output.time |= 0xF800u & (timestruct.tm_hour << 11);

    output.date = 0x001Fu & timestruct.tm_mday;
    output.date |= 0x01E0u & ((timestruct.tm_mon + 1) << 5);
    if (timestruct.tm_year >= 80 && timestruct.tm_year - 80 < 128)
    {
      output.date |= 0xFE00u & ((timestruct.tm_year - 80) << 9);
    }

    return output;
  }

  /**
   * Calculates the CRC-32 checksum variant used by ZIP on 
   * the input string data.
   */
  uint32_t crc32(const std::string &data) noexcept
  {
    static const uint32_t crc32_table[256] = {0x00000000u, 0x77073096u, 0xEE0E612Cu, 0x990951BAu, 0x076DC419u, 0x706AF48Fu, 0xE963A535u, 0x9E6495A3u, 0x0EDB8832u, 0x79DCB8A4u, 0xE0D5E91Eu, 
      0x97D2D988u, 0x09B64C2Bu, 0x7EB17CBDu, 0xE7B82D07u, 0x90BF1D91u, 0x1DB71064u, 0x6AB020F2u, 0xF3B97148u, 0x84BE41DEu, 0x1ADAD47Du, 0x6DDDE4EBu, 0xF4D4B551u, 0x83D385C7u, 0x136C9856u, 0x646BA8C0u, 
      0xFD62F97Au, 0x8A65C9ECu, 0x14015C4Fu, 0x63066CD9u, 0xFA0F3D63u, 0x8D080DF5u, 0x3B6E20C8u, 0x4C69105Eu, 0xD56041E4u, 0xA2677172u, 0x3C03E4D1u, 0x4B04D447u, 0xD20D85FDu, 0xA50AB56Bu, 0x35B5A8FAu, 
      0x42B2986Cu, 0xDBBBC9D6u, 0xACBCF940u, 0x32D86CE3u, 0x45DF5C75u, 0xDCD60DCFu, 0xABD13D59u, 0x26D930ACu, 0x51DE003Au, 0xC8D75180u, 0xBFD06116u, 0x21B4F4B5u, 0x56B3C423u, 0xCFBA9599u, 0xB8BDA50Fu, 
      0x2802B89Eu, 0x5F058808u, 0xC60CD9B2u, 0xB10BE924u, 0x2F6F7C87u, 0x58684C11u, 0xC1611DABu, 0xB6662D3Du, 0x76DC4190u, 0x01DB7106u, 0x98D220BCu, 0xEFD5102Au, 0x71B18589u, 0x06B6B51Fu, 0x9FBFE4A5u, 
      0xE8B8D433u, 0x7807C9A2u, 0x0F00F934u, 0x9609A88Eu, 0xE10E9818u, 0x7F6A0DBBu, 0x086D3D2Du, 0x91646C97u, 0xE6635C01u, 0x6B6B51F4u, 0x1C6C6162u, 0x856530D8u, 0xF262004Eu, 0x6C0695EDu, 0x1B01A57Bu, 
      0x8208F4C1u, 0xF50FC457u, 0x65B0D9C6u, 0x12B7E950u, 0x8BBEB8EAu, 0xFCB9887Cu, 0x62DD1DDFu, 0x15DA2D49u, 0x8CD37CF3u, 0xFBD44C65u, 0x4DB26158u, 0x3AB551CEu, 0xA3BC0074u, 0xD4BB30E2u, 0x4ADFA541u, 
      0x3DD895D7u, 0xA4D1C46Du, 0xD3D6F4FBu, 0x4369E96Au, 0x346ED9FCu, 0xAD678846u, 0xDA60B8D0u, 0x44042D73u, 0x33031DE5u, 0xAA0A4C5Fu, 0xDD0D7CC9u, 0x5005713Cu, 0x270241AAu, 0xBE0B1010u, 0xC90C2086u, 
      0x5768B525u, 0x206F85B3u, 0xB966D409u, 0xCE61E49Fu, 0x5EDEF90Eu, 0x29D9C998u, 0xB0D09822u, 0xC7D7A8B4u, 0x59B33D17u, 0x2EB40D81u, 0xB7BD5C3Bu, 0xC0BA6CADu, 0xEDB88320u, 0x9ABFB3B6u, 0x03B6E20Cu, 
      0x74B1D29Au, 0xEAD54739u, 0x9DD277AFu, 0x04DB2615u, 0x73DC1683u, 0xE3630B12u, 0x94643B84u, 0x0D6D6A3Eu, 0x7A6A5AA8u, 0xE40ECF0Bu, 0x9309FF9Du, 0x0A00AE27u, 0x7D079EB1u, 0xF00F9344u, 0x8708A3D2u, 
      0x1E01F268u, 0x6906C2FEu, 0xF762575Du, 0x806567CBu, 0x196C3671u, 0x6E6B06E7u, 0xFED41B76u, 0x89D32BE0u, 0x10DA7A5Au, 0x67DD4ACCu, 0xF9B9DF6Fu, 0x8EBEEFF9u, 0x17B7BE43u, 0x60B08ED5u, 0xD6D6A3E8u, 
      0xA1D1937Eu, 0x38D8C2C4u, 0x4FDFF252u, 0xD1BB67F1u, 0xA6BC5767u, 0x3FB506DDu, 0x48B2364Bu, 0xD80D2BDAu, 0xAF0A1B4Cu, 0x36034AF6u, 0x41047A60u, 0xDF60EFC3u, 0xA867DF55u, 0x316E8EEFu, 0x4669BE79u, 
      0xCB61B38Cu, 0xBC66831Au, 0x256FD2A0u, 0x5268E236u, 0xCC0C7795u, 0xBB0B4703u, 0x220216B9u, 0x5505262Fu, 0xC5BA3BBEu, 0xB2BD0B28u, 0x2BB45A92u, 0x5CB36A04u, 0xC2D7FFA7u, 0xB5D0CF31u, 0x2CD99E8Bu, 
      0x5BDEAE1Du, 0x9B64C2B0u, 0xEC63F226u, 0x756AA39Cu, 0x026D930Au, 0x9C0906A9u, 0xEB0E363Fu, 0x72076785u, 0x05005713u, 0x95BF4A82u, 0xE2B87A14u, 0x7BB12BAEu, 0x0CB61B38u, 0x92D28E9Bu, 0xE5D5BE0Du, 
      0x7CDCEFB7u, 0x0BDBDF21u, 0x86D3D2D4u, 0xF1D4E242u, 0x68DDB3F8u, 0x1FDA836Eu, 0x81BE16CDu, 0xF6B9265Bu, 0x6FB077E1u, 0x18B74777u, 0x88085AE6u, 0xFF0F6A70u, 0x66063BCAu, 0x11010B5Cu, 0x8F659EFFu, 
      0xF862AE69u, 0x616BFFD3u, 0x166CCF45u, 0xA00AE278u, 0xD70DD2EEu, 0x4E048354u, 0x3903B3C2u, 0xA7672661u, 0xD06016F7u, 0x4969474Du, 0x3E6E77DBu, 0xAED16A4Au, 0xD9D65ADCu, 0x40DF0B66u, 0x37D83BF0u, 
      0xA9BCAE53u, 0xDEBB9EC5u, 0x47B2CF7Fu, 0x30B5FFE9u, 0xBDBDF21Cu, 0xCABAC28Au, 0x53B39330u, 0x24B4A3A6u, 0xBAD03605u, 0xCDD70693u, 0x54DE5729u, 0x23D967BFu, 0xB3667A2Eu, 0xC4614AB8u, 0x5D681B02u, 
      0x2A6F2B94u, 0xB40BBE37u, 0xC30C8EA1u, 0x5A05DF1Bu, 0x2D02EF8Du};

    uint32_t crc_reg = 0xFFFFFFFFu;
    for (size_t jChar = 0u; jChar < data.size(); jChar++)
    {
      uint8_t table_indx = static_cast<uint8_t>(0x000000FFu & crc_reg) ^ static_cast<uint8_t>(data.at(jChar));
      crc_reg >>= 8;
      crc_reg ^= crc32_table[table_indx];
    }
    crc_reg ^= 0xFFFFFFFFu;
    return crc_reg;
  }

  /**
   * Helper routine that stores a uint16_t value in
   * an output buffer >= 2 bytes long in little endian
   * byte order.
   */
  void uint16_to_buffer(const uint16_t in, char *out) noexcept
  {
    out[0] = static_cast<char>(0x00FFu & in);
    out[1] = static_cast<char>(0x00FFu & (in >> 8));
  }

  /**
   * Helper routine that stores a uint32_t value in
   * an output buffer >= 4 bytes long in little endian
   * byte order.
   */
  void uint32_to_buffer(const uint32_t in, char *out) noexcept
  {
    out[0] = static_cast<char>(0x000000FFu & in);
    out[1] = static_cast<char>(0x000000FFu & (in >> 8));
    out[2] = static_cast<char>(0x000000FFu & (in >> 16));
    out[3] = static_cast<char>(0x000000FFu & (in >> 24));
  }

  /**
   * Default constructor.
   * Use the IttyZip::open() method to specify the output file
   * if the IttyZip object is constructed with this constructor.
   */
  IttyZip::IttyZip(void) noexcept : num_files(0u), opened(false), next_offset(0u) { }

  /**
   * Constructor that takes an output file filename
   * and attempts to open said output file.
   */
  IttyZip::IttyZip(const std::string &outputFilename) noexcept(false) : num_files(0u), next_offset(0u)
  {
    opened = false;
    out_file.open(outputFilename, std::ios::binary | std::ios::out | std::ios::trunc);
    if (!out_file.is_open())
    {
      throw std::exception(CANNOT_OPEN_MESG);
    }
    else
    {
      opened = true;
    }
  }

  /**
   * open() attempts to open the output file specified
   * by outputFilename.
   *
   * open() is only meant to be called after finalize() or on
   * a default constructed IttyZip object.
   */
  void IttyZip::open(const std::string &outputFilename) noexcept(false)
  {
    if (opened || out_file.is_open())
    {
      throw std::exception(DOUBLE_OPEN_MESG);
    }
    else
    {
      opened = false;
      num_files = 0u;
      next_offset = 0u;
      central_directory.clear();
      out_file.open(outputFilename, std::ios::binary | std::ios::out | std::ios::trunc);
      if (!out_file.is_open())
      {
        throw std::exception(CANNOT_OPEN_MESG);
      }
      else
      {
        opened = true;
      }
    }
  }

  /**
   * addFile() adds a new file to the IttyZip archive.
   * filename specifies the full path of the file in the archive.
   * contents stores the contents of the file.
   *
   * Since addFile() writes the contents to the output archive
   * immediately, addFile() may only be called on an IttyZip
   * object that has an open output file.
   */
  void IttyZip::addFile(const std::string &filename, const std::string &contents) noexcept(false)
  {
    if (!opened)
    {
      throw std::exception(NOT_OPENED_MESG);
    }
    else if (!out_file.is_open())
    {
      throw std::exception(UNEXPECTED_CLOSE_MESG);
    }
    else if (out_file.fail())
    {
      throw std::exception(OUTPUT_FAIL_MESG);
    }
    else
    {
      uint32_t file_crc32 = crc32(contents);
      std::pair<localheader_t, dirheader_t> file_headers = generateHeaders(filename, static_cast<uint32_t>(contents.size()), file_crc32);
      std::pair<std::set<std::string>::iterator, bool> ins_ret = filenames.insert(file_headers.first.filename);
      if (!ins_ret.second)
      {
        throw std::exception(DUPLICATE_FILE_MESG);
      }
      else
      {
        storeDirheader(file_headers.second);
        next_offset += writeLocalheader(file_headers.first);
        out_file.write(contents.c_str(), contents.length());
        next_offset += static_cast<uint32_t>(contents.size());
        num_files++;
      }
    }
  }

  /**
   * finalize() writes the central directory and the end of
   * central directory record to the output ZIP file and then
   * closes the output file.
   *
   * At least one file must have been added to this IttyZip
   * object before finalize() is called. This in turn requires
   * that the output file has already been opened.
   */
  void IttyZip::finalize(void) noexcept(false)
  {
    if (num_files == 0u || next_offset == 0u || central_directory.empty())
    {
      throw std::exception(EMPTY_FINALIZE_MESG);
    }
    else if (!opened)
    {
      throw std::exception(NOT_OPENED_MESG);
    }
    else if (!out_file.is_open())
    {
      throw std::exception(UNEXPECTED_CLOSE_MESG);
    }
    else if (out_file.fail())
    {
      throw std::exception(OUTPUT_FAIL_MESG);
    }
    else
    {
      out_file.write(central_directory.c_str(), central_directory.size());
      endrecord_t end_record = generateEndRecord();
      writeEndRecord(end_record);
      out_file.close();
      opened = false;
      next_offset = 0u;
      central_directory.clear();
      num_files = 0u;
    }
  }

  /**
   * Generates the local file header and the central directory
   * file header for the file with name filename, size file_size
   * (in bytes) and file contents CRC-32 checksum file_crc32.
   */
  std::pair<localheader_t, dirheader_t> IttyZip::generateHeaders(const std::string &filename, const uint32_t file_size, const uint32_t file_crc32) const noexcept
  {
    std::pair<localheader_t, dirheader_t> output;
    /* The signatures are defined by the ZIP specification. */
    output.first.signature = 0x04034b50u;
    output.second.signature = 0x02014b50u;
    /**
     * 000A is the least demanding version specifier.
     * 00 means the file external attributes are compatible
     * with MS-DOS.
     * 0A means ZIP version 1.0 could extract this file. This
     * is the default and appropriate, as we are not using
     * special features.
     */
    output.first.extract_version = 0x000Au;
    output.second.version_made_by = output.first.extract_version;
    output.second.extract_version = output.first.extract_version;
    /* Not using anything special in the general bit flag field. */
    output.first.general_bit_flag = 0u;
    output.second.general_bit_flag = output.first.general_bit_flag;
    /* "Compression" method is store: no compression. */
    output.first.compression_method = 0u;
    output.second.compression_method = output.first.compression_method;
    /* Simply use the current time and date for file modification. */
    output.first.file_mod_timedate = dosTimeDate();
    output.second.file_mod_timedate = output.first.file_mod_timedate;
    output.first.crc32 = file_crc32;
    output.second.crc32 = output.first.crc32;
    /* No compression, so compressed and uncompressed file sizes are the same. */
    output.first.size_compressed = file_size;
    output.second.size_compressed = output.first.size_compressed;
    output.first.size_uncompressed = output.first.size_compressed;
    output.second.size_uncompressed = output.first.size_compressed;
    output.first.filename_length = static_cast<uint16_t>(filename.size());
    output.second.filename_length = output.first.filename_length;
    /* Nothing special in the extra field; no file comment. */
    output.first.extra_field_length = 0u;
    output.second.extra_field_length = 0u;
    output.second.comment_length = 0u;
    /* Archive is contiguous; everything is on disk 0. */
    output.second.disk_number_start = 0u;
    /* No special file attributes. */
    output.second.internal_attributes = 0u;
    output.second.external_attributes = 0u;
    output.second.local_header_offset = next_offset;
    /* Using substr in case filename is longer than 65535 bytes. */
    output.first.filename = filename.substr(0u, output.first.filename_length);
    output.second.filename = output.first.filename;
    return output;
  }

  /**
   * Writes the local file header localheader to the output file
   * and returns the number of bytes written.
   */
  uint32_t IttyZip::writeLocalheader(const localheader_t &localheader) noexcept(false)
  {
    if (!opened)
    {
      throw std::exception(NOT_OPENED_MESG);
    }
    else if (!out_file.is_open())
    {
      throw std::exception(UNEXPECTED_CLOSE_MESG);
    }
    else if (out_file.fail())
    {
      throw std::exception(OUTPUT_FAIL_MESG);
    }
    else
    {
      char write_buffer[4];
      uint32_to_buffer(localheader.signature, write_buffer);
      out_file.write(write_buffer, 4);
      uint16_to_buffer(localheader.extract_version, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(localheader.general_bit_flag, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(localheader.compression_method, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(localheader.file_mod_timedate.time, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(localheader.file_mod_timedate.date, write_buffer);
      out_file.write(write_buffer, 2);
      uint32_to_buffer(localheader.crc32, write_buffer);
      out_file.write(write_buffer, 4);
      uint32_to_buffer(localheader.size_compressed, write_buffer);
      out_file.write(write_buffer, 4);
      uint32_to_buffer(localheader.size_uncompressed, write_buffer);
      out_file.write(write_buffer, 4);
      uint16_to_buffer(localheader.filename_length, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(localheader.extra_field_length, write_buffer);
      out_file.write(write_buffer, 2);
      out_file.write(localheader.filename.c_str(), localheader.filename_length);
      return 30u + static_cast<uint32_t>(localheader.filename_length);
    }

    return 0u;
  }

  /**
   * Appends the central directory file header dirheader
   * to the data member central_directory.
   */
  void IttyZip::storeDirheader(const dirheader_t &dirheader) noexcept
  {
    char store_buffer[4];
    uint32_to_buffer(dirheader.signature, store_buffer);
    central_directory.append(store_buffer, 4);
    uint16_to_buffer(dirheader.version_made_by, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.extract_version, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.general_bit_flag, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.compression_method, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.file_mod_timedate.time, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.file_mod_timedate.date, store_buffer);
    central_directory.append(store_buffer, 2);
    uint32_to_buffer(dirheader.crc32, store_buffer);
    central_directory.append(store_buffer, 4);
    uint32_to_buffer(dirheader.size_compressed, store_buffer);
    central_directory.append(store_buffer, 4);
    uint32_to_buffer(dirheader.size_uncompressed, store_buffer);
    central_directory.append(store_buffer, 4);
    uint16_to_buffer(dirheader.filename_length, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.extra_field_length, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.comment_length, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.disk_number_start, store_buffer);
    central_directory.append(store_buffer, 2);
    uint16_to_buffer(dirheader.internal_attributes, store_buffer);
    central_directory.append(store_buffer, 2);
    uint32_to_buffer(dirheader.external_attributes, store_buffer);
    central_directory.append(store_buffer, 4);
    uint32_to_buffer(dirheader.local_header_offset, store_buffer);
    central_directory.append(store_buffer, 4);
    central_directory.append(dirheader.filename, 0u, dirheader.filename_length);
  }

  /**
   * Generates the end of central directory record.
   * Purely a subroutine of finalize().
   */
  endrecord_t IttyZip::generateEndRecord(void) const noexcept
  {
    endrecord_t output;
    /* The signatures are defined by the ZIP specification. */
    output.signature = 0x06054b50u;
    /* Everything is on disk 0 in this archive. */
    output.disk_number = 0u;
    output.dir_start_disk_number = 0u;
    output.this_disk_entries = num_files;
    output.total_entries = output.this_disk_entries;
    output.central_dir_size = static_cast<uint32_t>(central_directory.size());
    output.central_dir_offset = next_offset;
    /* No file comment. */
    output.comment_length = 0u;
    return output;
  }

  /**
   * Writes the end of central directory record to
   * the output file. Purely a subroutine of finalize().
   */
  void IttyZip::writeEndRecord(const endrecord_t &end_record) noexcept(false)
  {
    if (!opened)
    {
      throw std::exception(NOT_OPENED_MESG);
    }
    else if (!out_file.is_open())
    {
      throw std::exception(UNEXPECTED_CLOSE_MESG);
    }
    else if (out_file.fail())
    {
      throw std::exception(OUTPUT_FAIL_MESG);
    }
    else
    {
      char write_buffer[4];
      uint32_to_buffer(end_record.signature, write_buffer);
      out_file.write(write_buffer, 4);
      uint16_to_buffer(end_record.disk_number, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(end_record.dir_start_disk_number, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(end_record.this_disk_entries, write_buffer);
      out_file.write(write_buffer, 2);
      uint16_to_buffer(end_record.total_entries, write_buffer);
      out_file.write(write_buffer, 2);
      uint32_to_buffer(end_record.central_dir_size, write_buffer);
      out_file.write(write_buffer, 4);
      uint32_to_buffer(end_record.central_dir_offset, write_buffer);
      out_file.write(write_buffer, 4);
      uint16_to_buffer(end_record.comment_length, write_buffer);
      out_file.write(write_buffer, 2);
    }
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
