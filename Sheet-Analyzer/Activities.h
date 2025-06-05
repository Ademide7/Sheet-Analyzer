#pragma once
#include <string>
#include <list>
#include <vector>
#include <map>
#include <memory> 

using namespace std;
class Activities
{
public:
	map<int, string> activitiesList;

public:
	Activities();
	map<int, string> getActivityById(const int& activityId);
	map<int, string> getActivitiesList();


	//Get Excel Content Stats
	void getExcelContentStats(const string& filePath);
	void addNAToEmptyCells(const string & filePath);
	void adddTableDesign(const string& filePath, string& color);


};

