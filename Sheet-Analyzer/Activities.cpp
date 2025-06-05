#include "Activities.h"
#include <string>
#include <list>
#include <vector>
#include <map>
#include <map>
#include <memory> 
#include "Utilities.h" 
#include <iostream>
#include <xlnt/cell/cell.hpp>
#include <xlnt/xlnt.hpp>  
 

Activities::Activities()
{  
        activitiesList = {
        { 0, "Exit" },
        { 1, "Add Design Table" },
        { 2, "Add N/A to Empty Cells" },
        { 3, "Get Excel Content Stats" }};
}

map<int, string> Activities::getActivityById(const int& activityId)
{
    if(activitiesList.size() ==  0)
    {
        return {};
	}
    auto it = activitiesList.find(activityId);
    if (it != activitiesList.end())
    {
        return { *it };
    }
    return {};
}

map<int, string> Activities::getActivitiesList()
{
    if(activitiesList.size() == 0)
    {
        return {};
    } 
	return activitiesList;
}

void Activities::getExcelContentStats(const string& filePath)
{
    Utilities utilities; // Create an instance of the Utilities class
    // Read the file using the utilities module
    utilities.ReadFile(filePath); // Initialize the utilities module

    auto res = utilities.getExtractedData();

    auto dimension = utilities.getDimensions();
    std::cout << "Rows: " << dimension.first << ", Columns: " << dimension.second << "\n";


    auto modes = utilities.getColumnModes();

    for (size_t i = 0; i < modes.size(); ++i) {
        std::cout << "Col " << i << " Most ocurring value: " << modes[i] << "\n";
    }
}

void Activities::addNAToEmptyCells(const string& filePath)
{
    Utilities utilities; // Create an instance of the Utilities class
    // Read the file using the utilities module
	//utilities.ExtractedData.clear(); // Clear any previous data
    utilities.ReadFile(filePath); // Initialize the utilities module
    std::vector<ExcelRow>*  res = utilities.getExtractedData();
    // Iterate through the extracted data and replace empty cells with "N/A"
    for (ExcelRow& row : *res) {
        for (auto& cell : row.cells) {
            if (cell.empty()) {
                cell = "N/A"; // Replace empty cell with "N/A"
            }
        }
    }
    // Save the modified data back to the file or handle it as needed
	utilities.UpdateExtractedData(*res);
}

void Activities::adddTableDesign(const string& filePath,string &color)
{
	// saving extracted data from utilities module to a new file adding a desiged table to the excel file with a specific color int 
    Utilities utilities; // Create an instance of the Utilities class
    // Read the file using the utilities module 
    utilities.ReadFile(filePath); // Initialize the utilities module
    std::vector<ExcelRow>*  res = utilities.getExtractedData();
	// Check if the extracted data is empty
    if (res->empty()) {
        std::cout << "No data extracted from the file.\n";
        return;
	}

    // Create a workbook and add a worksheet
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.active_sheet();
    ws.title("StyledData");

    // Determine fill color
    xlnt::fill fill_style;
    if (color == "red") {
        fill_style = xlnt::fill::solid(xlnt::color::red());
    }
    else if (color == "blue") {
        fill_style = xlnt::fill::solid(xlnt::color::blue());
    }
    else if (color == "green") {
        fill_style = xlnt::fill::solid(xlnt::color::green());
    }
    else {
        fill_style = xlnt::fill::solid(xlnt::color::white()); // Default/fallback color
    }

    // Write cell data and apply fill style
    for (std::size_t row_idx = 0; row_idx < res->size(); ++row_idx) {
        const ExcelRow& row = (*res)[row_idx];
        for (std::size_t col_idx = 0; col_idx < row.cells.size(); ++col_idx) {
            std::string value = row.cells[col_idx];
            auto cell = ws.cell(xlnt::cell_reference(static_cast<int>(col_idx + 1), static_cast<int>(row_idx + 1)));
            cell.value(value);
            cell.fill(fill_style);
        }
    }

    // Save the styled Excel file
	string outputFilePath = "styled_data"; // Specify the output file path  
	// add random suffix to the output file name to avoid overwriting
	outputFilePath += std::to_string(rand() % 1000) + ".xlsx"; // Random suffix for uniqueness
    wb.save(outputFilePath);
}
 

 
