#pragma once
#include <string>
#include <vector>
#include <memory> // Ensure this header is included for std::make_unique

using namespace std;

struct ExcelRow {
	std::vector<std::string> cells;
};

class Utilities
{
   public:
	 unique_ptr<std::vector<ExcelRow>> ExtractedData;

  public: 
	  Utilities(); // constructor
 
	 void ReadFile(const string& filePath);
	 void setExtractedData(const std::vector<ExcelRow>& data);
	 vector<ExcelRow>* getExtractedData() ;
	 std::pair<size_t, size_t> getDimensions() const;
	 std::vector<std::string> getColumnModes() const;
	 void UpdateExtractedData(vector<ExcelRow> &res);
	  
};

