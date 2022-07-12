#include "ExcelParser.hpp"

using namespace excel_parser;

/********************************************************************************************************************
 * PROTECTED MEMBERS ************************************************************************************************
 ********************************************************************************************************************/
ExcelParser *ExcelParser::instance = nullptr;

std::mutex ExcelParser::io_mutex;

std::map<std::string, std::map<int, std::string>> ExcelParser::shared_strings_map;

std::map<std::string, std::map<std::string, sheet>> ExcelParser::sheets_map;

/********************************************************************************************************************
 * PUBLIC METHODS ***************************************************************************************************
 ********************************************************************************************************************/
ExcelParser *ExcelParser::getInstance()
{
	std::lock_guard<std::mutex> lock(io_mutex);

	if (instance == nullptr)
	{
		instance = new ExcelParser();
	}
	return instance;
}

void ExcelParser::openExcelFile(std::string file_name)
{
	std::lock_guard<std::mutex> lock(io_mutex);
	if (sheets_map.find(file_name) == sheets_map.end())
	{
		int err = 0;
		zip *book = zip_open(file_name.c_str(), 0, &err);
		if (err)
		{
			std::string error_message = "[Excel Parser] (ERROR) Error opening spreadsheet archive: " + std::to_string(err);
			throw std::runtime_error(error_message);
		}

		readSharedStrings(file_name, book);
		parseSheetTrees(file_name, readWorkbookToTrees(book));

		zip_close(book);
	}
}

void ExcelParser::closeExcelFile(std::string file_name)
{
	std::lock_guard<std::mutex> lock(io_mutex);
	if (sheets_map.find(file_name) != sheets_map.end())
	{
		sheets_map.erase(file_name);
	}
	if (shared_strings_map.find(file_name) != shared_strings_map.end())
	{
		shared_strings_map.erase(file_name);
	}
}

sheet ExcelParser::getSheet(std::string file_name, std::string sheet_name)
{
	std::lock_guard<std::mutex> lock(io_mutex);
	if (sheets_map.find(file_name) == sheets_map.end())
	{
		std::string error_message = "[Excel Parser] (ERROR) Error finding spreadsheet with name: " + file_name;
		throw std::runtime_error(error_message);
	}
	else if (sheets_map.at(file_name).find(sheet_name) == sheets_map.at(file_name).end())
	{
		std::string error_message = "[Excel Parser] (ERROR) Error finding sheet with name \"" + sheet_name + "\" in file " + file_name;
		throw std::runtime_error(error_message);
	}
	else
	{
		return sheets_map.at(file_name).at(sheet_name);
	}
}

std::string ExcelParser::getSharedString(std::string file_name, int shared_string_index)
{
	std::lock_guard<std::mutex> lock(io_mutex);
	if (shared_strings_map.find(file_name) == shared_strings_map.end())
	{
		std::string error_message = "[Excel Parser] (ERROR) Error finding spreadsheet with name: " + file_name;
		throw std::runtime_error(error_message);
	}
	else if (shared_strings_map.at(file_name).find(shared_string_index) == shared_strings_map.at(file_name).end())
	{
		std::string error_message = "[Excel Parser] (ERROR) Error finding shared string with index " + std::to_string(shared_string_index) + " in file " + file_name;
		throw std::runtime_error(error_message);
	}
	else
	{
		return shared_strings_map.at(file_name).at(shared_string_index);
	}
}

std::vector<std::string> ExcelParser::getSheetNames(std::string file_name)
{
	std::lock_guard<std::mutex> lock(io_mutex);
	if (sheets_map.find(file_name) == sheets_map.end())
	{
		std::string error_message = "[Excel Parser] (ERROR) Error finding spreadsheet with name: " + file_name;
		throw std::runtime_error(error_message);
	}
	else
	{
		std::vector<std::string> sheet_names;
		for (auto &s : sheets_map.at(file_name))
		{
			sheet_names.push_back(s.first);
		}
		return sheet_names;
	}
}

/********************************************************************************************************************
 * PROTECTED METHODS ************************************************************************************************
 ********************************************************************************************************************/
void ExcelParser::readSharedStrings(std::string file_name, zip *book)
{
	int index = -1;
	boost::property_tree::ptree strings_tree = readFileFromArchive(book, std::string("sharedStrings.xml"));
	try
	{
		boost::property_tree::ptree string_trees = strings_tree.get_child("sst");
		for (auto &string_values : string_trees)
		{
			if (string_values.first.compare(XML_ATTR) != 0)
			{
				try
				{
					++index;
					if (string_values.second.find("r") != string_values.second.not_found())
					{
						std::string s;
						// boost::property_tree::ptree rows = string_values.second.get_child("si");
						for (auto &r : string_values.second)
						{
							s = s + r.second.get_child("t").data();
						}
						shared_strings_map[file_name].emplace(std::pair<int, std::string>(index, s));
					}
					else
					{
						boost::property_tree::ptree string_tree = string_values.second.get_child("t");
						shared_strings_map[file_name].emplace(std::pair<int, std::string>(index, string_tree.data()));
					}
				}
				catch (boost::property_tree::ptree_error ptree_error)
				{
					std::cout << "[Excel Parser] (ERROR) Error accessing the shared string at index " << index << " in the property tree: " << ptree_error.what() << std::endl;
				}
			}
		}
	}
	catch (boost::property_tree::ptree_error ptree_error)
	{
		std::cout << "[Excel Parser] (ERROR) Error accessing the shared strings property tree: " << ptree_error.what() << std::endl;
	}
}

std::map<std::string, boost::property_tree::ptree> ExcelParser::readWorkbookToTrees(zip *book)
{
	boost::property_tree::ptree workbook_tree;
	std::map<int, std::string> id_name_map;
	std::map<std::string, boost::property_tree::ptree> name_sheet_tree_map;
	try
	{
		// Search for the file of given file_name and parse the XML workbook into a property tree.
		workbook_tree = readFileFromArchive(book, "workbook.xml");
	}
	catch (std::runtime_error runtime_error)
	{
		std::cout << "[Excel Parser] (ERROR) Reading the workbook: " << runtime_error.what() << std::endl;
		exit(1);
	}
	try
	{
		boost::property_tree::ptree sheets_tree = workbook_tree.get_child("workbook.sheets");
		for (auto &sheets_values : sheets_tree)
		{
			boost::property_tree::ptree sheet_attributes = sheets_values.second.get_child(XML_ATTR);
			std::string file_name = sheet_attributes.get_child("name").data();
			int id = stoi(sheet_attributes.get_child("r:id").data().erase(0, 3));
			id_name_map.emplace(std::pair<int, std::string>(id, file_name));
		}
	}
	catch (boost::property_tree::ptree_error ptree_error)
	{
		std::cout << "[Excel Parser] (ERROR) Error accessing the workbook property tree: " << ptree_error.what() << std::endl;
	}

	for (std::map<int, std::string>::iterator it = id_name_map.begin(); it != id_name_map.end(); ++it)
	{
		std::string sheet_name = "sheet" + std::to_string(it->first) + ".xml";
		boost::property_tree::ptree sheet_tree;
		try
		{
			sheet_tree = readFileFromArchive(book, sheet_name);
			name_sheet_tree_map.emplace(std::pair<std::string, boost::property_tree::ptree>(it->second, sheet_tree));
		}
		catch (std::runtime_error runtime_error)
		{
			std::cout << "[Excel Parser] (ERROR) Reading " << it->second << " sheet: " << runtime_error.what() << std::endl;
		}
	}
	return name_sheet_tree_map;
}

void ExcelParser::parseSheetTrees(std::string file_name, std::map<std::string, boost::property_tree::ptree> name_sheet_tree_map)
{
	for (std::map<std::string, boost::property_tree::ptree>::iterator it = name_sheet_tree_map.begin(); it != name_sheet_tree_map.end(); ++it)
	{
		sheets_map[file_name].emplace(std::pair<std::string, sheet>(it->first, parseSheet(it->second)));
	}
}

sheet ExcelParser::parseSheet(boost::property_tree::ptree sheet_tree)
{
	sheet s = sheet();
	boost::property_tree::ptree rows_tree = sheet_tree.get_child("worksheet.sheetData");
	for (auto &row_tree : rows_tree)
	{
		if (row_tree.first.compare("row") == 0)
		{
			int row_id;
			row r = row();
			try
			{
				row_id = stoi(getAttributes(row_tree.second).at("r"));
				for (auto &column_tree : row_tree.second)
				{
					if (column_tree.first.compare("c") == 0)
					{
						try
						{
							cell_t c;
							c.value = column_tree.second.get_child("v").data();
							xml_attributes cell_attributes = getAttributes(column_tree.second);
							std::string cell_name = cell_attributes.at("r");
							cell_name.erase(std::remove_if(cell_name.begin(), cell_name.end(), [](char c)
														   { return !std::isalpha(c); }),
											cell_name.end());
							CellType cell_type;
							try
							{
								std::string type_string = cell_attributes.at("t");
								cell_type = STRING;
							}
							catch (std::out_of_range out_of_range)
							{
								cell_type = NUMBER;
							}
							c.type = cell_type;
							r.emplace(std::pair<std::string, cell_t>(cell_name, c));
						}
						catch (boost::property_tree::ptree_error ptree_error)
						{
						}
					}
				}
				s.emplace(std::pair<int, row>(row_id, r));
			}
			catch (std::runtime_error runtime_error)
			{
				std::cout << "[Excel Parser] (ERROR) Error retrieving XML attributes in row: " << row_tree.second.data() << runtime_error.what() << std::endl;
			}
			catch (std::out_of_range out_of_range)
			{
				std::cout << "[Excel Parser] (ERROR) Error parsing sheet could not value with key: " << out_of_range.what() << std::endl;
			}
		}
	}
	return s;
}

boost::property_tree::ptree ExcelParser::readFileFromArchive(zip *book, std::string file_name)
{
	// Search for the file of given file_name
	zip_uint64_t location = zip_name_locate(book, file_name.c_str(), ZIP_FL_NODIR);
	if (location < 0)
	{
		std::string error_message = "[Excel Parser] (ERROR) Error cannot find file in provided archive with file_name: " + file_name;
		throw new std::runtime_error(error_message);
	}

	// Get the information on the workbook file.
	struct zip_stat workbook_stat;
	zip_stat_init(&workbook_stat);
	zip_stat_index(book, location, 0, &workbook_stat);
	if (!(workbook_stat.valid & (ZIP_STAT_NAME | ZIP_STAT_SIZE)))
	{
		std::string error_message = "[Excel Parser] (ERROR) Error retrieving metadata for file " + file_name + ": " + std::to_string(workbook_stat.valid);
		throw new std::runtime_error(error_message);
	}

	// Alloc memory for its uncompressed contents
	char *contents = new char[workbook_stat.size];

	// Read the compressed file
	zip_file *f = zip_fopen_index(book, location, 0);
	if (zip_fread(f, contents, workbook_stat.size) == -1)
	{
		std::string error_message = "[Excel Parser] (ERROR) Error reading file " + file_name + ".";
		throw new std::runtime_error(error_message);
	}

	// And close the file and archive
	zip_fclose(f);

	// Parse the XML workbook into a property tree.
	boost::property_tree::ptree tree;
	std::stringstream stream = std::stringstream(std::string(contents, workbook_stat.size));
	try
	{
		boost::property_tree::xml_parser::read_xml(stream, tree);
		return tree;
	}
	catch (boost::property_tree::xml_parser_error parse_error)
	{
		std::string error_message = "[Excel Parser] (ERROR) Error parsing information in " + file_name + " at line " + std::to_string(parse_error.line()) + ": " + parse_error.what();
		throw new std::runtime_error(error_message);
	}
	catch (boost::property_tree::ptree_error ptree_error)
	{
		std::string error_message = "[Excel Parser] (ERROR) Error accessing " + file_name + " property tree: " + ptree_error.what();
		throw new std::runtime_error(error_message);
	}
}

xml_attributes ExcelParser::getAttributes(boost::property_tree::ptree tree)
{
	xml_attributes attributes_map;
	try
	{
		boost::property_tree::ptree attributes_tree = tree.get_child(XML_ATTR);
		for (auto &attribute_value : attributes_tree)
		{
			attributes_map.emplace(std::pair<std::string, std::string>(attribute_value.first, attribute_value.second.data()));
		}
	}
	catch (boost::property_tree::ptree_error ptree_error)
	{
		std::string error_message = std::string("[Excel Parser] (ERROR) Error finding attributes of property tree: ") + ptree_error.what();
		throw new std::runtime_error(error_message);
	}
	return attributes_map;
}
