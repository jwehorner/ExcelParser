#include <iostream>

#include "DirectoryConfig.hpp"

#include "ExcelParser.hpp"

using namespace std;
using namespace excel_parser;

int test_openExcelFile();
int test_getSheet();
int test_getSharedString();
int test_getSheetNames();

int main()
{
	int passed = test_openExcelFile();
	cout << "Get openExcelFile passed " << passed << "/1 tests." << endl;
	passed = test_getSheet();
	cout << "Get getSheet passed " << passed << "/2 tests." << endl;
	passed = test_getSharedString();
	cout << "Get getSharedString passed " << passed << "/2 tests." << endl;
	passed = test_getSheetNames();
	cout << "Get getSheetNames passed " << passed << "/1 tests." << endl;
}

int test_openExcelFile()
{
	int test_passes = 0;
	ExcelParser *parser = ExcelParser::getInstance();
	string test_name = string(PROJECT_DIRECTORY) + string("/input/TestBook.xlsx");
	try
	{
		parser->openExcelFile(test_name);
		sheet s = parser->getSheet(test_name, "sheet");
		++test_passes;
	}
	catch (runtime_error e)
	{
		cout << e.what() << endl;
	}
	return test_passes;
}
int test_getSheet()
{
	int test_passes = 0;
	ExcelParser *parser = ExcelParser::getInstance();
	string test_name = string(PROJECT_DIRECTORY) + string("/input/TestBook.xlsx");
	try
	{
		parser->openExcelFile(test_name);
		sheet sheet1 = parser->getSheet(test_name, "sheet");
		if (sheet1.size() > 0)
		{
			++test_passes;
		}
		sheet sheet2 = parser->getSheet(test_name, "2sheetOrNot2sheet");
		if (sheet2.size() > 0)
		{
			++test_passes;
		}
	}
	catch (runtime_error e)
	{
		cout << e.what() << endl;
	}
	return test_passes;
}
int test_getSharedString()
{
	int test_passes = 0;
	ExcelParser *parser = ExcelParser::getInstance();
	string test_name = string(PROJECT_DIRECTORY) + string("/input/TestBook.xlsx");
	try
	{
		parser->openExcelFile(test_name);
		sheet sheet = parser->getSheet(test_name, "sheet");
		if (parser->getSharedString(test_name, stoi(sheet.at(1).at("A").value)).compare("TestColum") == 0)
		{
			++test_passes;
		}
		if (parser->getSharedString(test_name, stoi(sheet.at(2).at("A").value)).compare("row 1") == 0)
		{
			++test_passes;
		}
	}
	catch (runtime_error e)
	{
		cout << e.what() << endl;
	}
	return test_passes;
}
int test_getSheetNames()
{
	int test_passes = 0;
	ExcelParser *parser = ExcelParser::getInstance();
	string test_name = string(PROJECT_DIRECTORY) + string("/input/TestBook.xlsx");
	try
	{
		parser->openExcelFile(test_name);
		vector<string> names = parser->getSheetNames(test_name);
		if (names.size() == 2)
		{
			++test_passes;
		}
	}
	catch (runtime_error e)
	{
		cout << e.what() << endl;
	}
	return test_passes;
}
