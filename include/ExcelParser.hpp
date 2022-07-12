/**
 * @file    ExcelParser.hpp
 * @author  James Horner
 * @brief   This file contains the declaration of the ExcelParser Singleton for accessing Excel files in C++.
 * @details The ExcelParser is responsible for opening, parsing, and storing the data from multiple Excel files. 
 *          The class is a singleton and includes mutexes to make it thread-safe.
 * @date    2022-07-07
 *
 * @copyright Copyright (c) 2022
 */
#ifndef ExcelParser_HPP
#define ExcelParser_HPP

#include <algorithm>
#include <iostream>
#include <locale>
#include <map>
#include <mutex>
#include <string>
#include <sstream>

#include <boost/property_tree/ptree.hpp>
#include <boost/property_tree/xml_parser.hpp>

#include <zip.h>

#define XML_ATTR "<xmlattr>"

namespace excel_parser
{
    /**
     * @brief Enumeration of the different value types a cell can take.
     */
    typedef enum
    {
        NUMBER,
        STRING
    } CellType;

    /**
     * @brief Structural representation of the type and contents of a cell.
     */
    typedef struct
    {
        CellType type;
        std::string value;
    } cell_t;

    /**
     * @brief   Type definition representing a row of cells in a sheet.
     * @note    The string entry denotes the Excel column index starting at 'A'.
     */
    using row = std::map<std::string, cell_t>;

    /**
     * @brief   Type definition representing a sheet in an Excel file.
     */
    using sheet = std::map<int, row>;

    /**
     * @brief   Type definition representing a map of XML attribute names to attribute values. 
     */
    using xml_attributes = std::map<std::string, std::string>;

    /**
     * @brief   Class ExcelParser is a Singleton that controls access to the contents of Excel files.
     * @details The singleton instance is responsible for opening, parsing, storing, and supplying
     *          access to the Excel sheet contents. The class uses mutexes to ensure mutual exclusion 
     *          between client threads.
     */
    class ExcelParser
    {

    private:
        /// Instance of the ExcelParser Singleton.
        static ExcelParser *instance;
        /// Mutex to control access to the class.
        static std::mutex io_mutex;
        /// Map of file names to the map of shared strings in the file
        static std::map<std::string, std::map<int, std::string>> shared_strings_map;
        /// Map of file names to the map of sheets in the file
        static std::map<std::string, std::map<std::string, sheet>> sheets_map;

    protected:
        /**
         * @brief   Constructor for the ExcelParser class only to be used by the getInstance method.
         */
        ExcelParser() {}

        /**
         * @brief           Method readSharedStrings reads the shared strings file in the Excel archive then uses its
         *                  content to populate the shared strings map.
         * @param file_name string name of the Excel file that is being read.
         * @param book      pointer to the libzip handle for the Excel file.
         */
        static void readSharedStrings(std::string file_name, zip *book);

        /**
         * @brief       Method readWorkbookToTrees reads the workbook file in the Excel archive then uses its contents
         *              to read the sheets of the Excel file into boost property trees.
         * @param book  pointer to the libzip handle for the Excel file.
         * @return      std::map<std::string, boost::property_tree::ptree> map of sheet names to property trees of XML.
         */
        static std::map<std::string, boost::property_tree::ptree> readWorkbookToTrees(zip *book);

        /**
         * @brief                       Method parseSheetTrees populates the sheets map using the map of strings to
         *                              property trees of XML by parsing the XML into sheets of rows of cells.
         * @param file_name             string name of the Excel file that is being read.
         * @param name_sheet_tree_map   map of sheet names to property trees of XML.
         */
        static void parseSheetTrees(std::string file_name, std::map<std::string, boost::property_tree::ptree> name_sheet_tree_map);

        /**
         * @brief               Method parseSheet parses an individual sheet of XML into a sheet object that is returned.
         * @param sheet_tree    a property tree of XML from an Excel sheet.
         * @return              excel_parser::sheet  sheet object representation of the XML sheet.
         */
        static excel_parser::sheet parseSheet(boost::property_tree::ptree sheet_tree);

        /**
         * @brief           Method readFileFromArchive reads an individual file from the Excel archive into a property
         *                  tree of XML.
         * @param book      pointer to the libzip handle for the Excel file.
         * @param file_name string name of the file to be read from the archive.
         * @return          boost::property_tree::ptree  a boost property tree of XML.
         * @note            The file name should not contain any path to the file as libzip will search for any files
         *                  whose name matches.
         */
        static boost::property_tree::ptree readFileFromArchive(zip *book, std::string file_name);

        /**
         * @brief       Method getAttributes retrieves the XML attributes for the top level node of a property tree of XML.
         * @param tree  a boost property tree of XML.
         * @return      std::map<std::string, std::string>   map of string attribute names to string attribute values.
         */
        static std::map<std::string, std::string> getAttributes(boost::property_tree::ptree tree);

    public:
        /**
         * @brief   Deleted cloning constructor so only one controller can exist.
         */
        ExcelParser(ExcelParser &other) = delete;

        /**
         * @brief   Deleted assignment operator so only one controller can exist.
         */
        void operator=(const ExcelParser &) = delete;

        /**
         * @brief   Method getInstance retrieves the instance of the ExcelParser singleton.
         * @return  ExcelParser* pointer to the instance of the class.
         */
        static ExcelParser *getInstance();

        /**
         * @brief           Method openExcelFile opens an Excel file and parses its contents into internal data structures.
         * @param file_name string name of the file to be opened.
         */
        static void openExcelFile(std::string file_name);

        /**
         * @brief           Method closeExcelFile closes and discards the data of an Excel file.
         * @param file_name string name of the file to be opened.
         */
        static void closeExcelFile(std::string file_name);

        /**
         * @brief               Method getSheet returns the sheet object with the given name from the specified file.
         * @param file_name     string name of the file which the sheet is in.
         * @param sheet_name    string name of the sheet of which to get the associated object.
         * @return              sheet object with the data contained within the sheet.
         */
        static sheet getSheet(std::string file_name, std::string sheet_name);

        /**
         * @brief                       Method getSharedString retrieves the Shared String with the given index in the
         *                              specified file.
         * @param file_name             string name of the file which the shared string is in.
         * @param shared_string_index   index of the shared string in the file.
         * @return                      std::string value of the string at the given index.
         * @note                        The index of the shared string is simply the int value of a cell_t structure when
         *                              the CellType is STRING.
         */
        static std::string getSharedString(std::string file_name, int shared_string_index);

        /**
         * @brief           Method getSheetNames retrieves the names of all the sheets in a given Excel file.
         * @param file_name Name of the file from which to retrive all the sheet names.
         * @return          std::vector<std::string> of names of sheets.
         */
        static std::vector<std::string> getSheetNames(std::string file_name);
    };
}

#endif /* ExcelParser_HPP */
