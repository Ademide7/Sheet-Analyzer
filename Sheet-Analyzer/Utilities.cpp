#include "Utilities.h"  
#include <iostream>  
#include <xlnt/xlnt.hpp>  
#include <vector>  
#include <string>    
#include <memory> // Ensure this header is included for std::make_unique
#include <unordered_map>

Utilities::Utilities(){ 
	ExtractedData = make_unique<std::vector<ExcelRow>>();
}

void Utilities::setExtractedData(const std::vector<ExcelRow>& data)
{
	ExtractedData = make_unique<std::vector<ExcelRow>>(data);
}

 vector<ExcelRow>* Utilities::getExtractedData()  {
	return ExtractedData.get();
}

void Utilities::ReadFile(const string& filePath)
{
	try
	{
		// read excel file  
		xlnt::workbook wb;

		// Convert std::string to std::filesystem::path or const char* as required by xlnt::workbook::load
		wb.load(filePath);

		xlnt::worksheet ws = wb.active_sheet();

		std::vector<ExcelRow> table;

		for (auto row : ws.rows(false)) {
			ExcelRow er;
			for (auto cell : row) {
				if (cell.has_value()) {
					er.cells.push_back(cell.to_string());
				}
				else {
					er.cells.push_back("");
				}
			}
			table.push_back(er);
		}

		setExtractedData(table);
	}

	catch (const xlnt::exception& ex) {
		std::cerr << "xlnt exception: " << ex.what() << std::endl;
	}

}

std::pair<size_t, size_t> Utilities::getDimensions() const {
	if (!ExtractedData || ExtractedData->empty())
		return { 0, 0 };

	size_t rows = ExtractedData->size();
	size_t cols = ExtractedData->at(0).cells.size();
	return { rows, cols };
}

std::vector<std::string> Utilities::getColumnModes() const {
	std::vector<std::string> modes;
	if (!ExtractedData || ExtractedData->empty()) return modes;

	size_t cols = ExtractedData->at(0).cells.size();
	modes.resize(cols);

	for (size_t col = 0; col < cols; ++col) {
		std::unordered_map<std::string, int> freq;
		int maxCount = 0;
		std::string mode;

		for (const auto& row : *ExtractedData) {
			if (col < row.cells.size()) {
				const auto& val = row.cells[col];
				freq[val]++;
				if (freq[val] > maxCount) {
					maxCount = freq[val];
					mode = val;
				}
			}
		}

		modes[col] = mode;
	}

	return modes;
}

void Utilities::UpdateExtractedData(vector<ExcelRow> &res)
{
	//Overwrite the file with updated data
	xlnt::workbook wb;
	xlnt::worksheet ws = wb.create_sheet();

	// Write data from 'res' into the worksheet
	for (std::size_t row_idx = 0; row_idx < res.size(); ++row_idx) {
		const ExcelRow& row = res[row_idx];
		for (std::size_t col_idx = 0; col_idx < row.cells.size(); ++col_idx) {
			std::string value = row.cells[col_idx];
			if (value.empty()) value = "N/A"; // Handle empty cells

			// Excel cells are 1-indexed
			ws.cell(xlnt::cell_reference(static_cast<int>(col_idx + 1), static_cast<int>(row_idx + 1))).value(value);
		}
	}

	// Save the workbook
	wb.save("updated_data.xlsx");
}


