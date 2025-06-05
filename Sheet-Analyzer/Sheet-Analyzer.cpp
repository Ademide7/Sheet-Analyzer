// Sheet-Analyzer.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include "Utilities.h"
#include <iostream>
#include "Utilities.h"
#include "Activities.h"


 
int main()
{
    std::cout << "::::::::::: WELCOME TO Sheet-Analyzer Console :::::::::::";
	std::cout << std::endl;
	std::cout << "This is a console application to read and analyze Excel files." << std::endl;
	std::cout << "Please enter your file path.(e.g ==> C:\\Users\\LENOVO\\Downloads\\random_mixed_data.xlsx)" << std::endl;
	std::cout << "Enter file path: ";
    std::string filePath;
	std::cin >> filePath;   // Get the file path from the user
	cout << "File read successfully!" << endl;
	std::cout << "WARNING DO NOT CLOSE THIS CONSOLE!!!!!!!!!" << std::endl;
	std::cout << "THIS MIGHT CAUSE LOSS OF FILE OR DATA , USE THE EXIT OPTION BELOW" << std::endl;

	// create looping option of operations 
	int choice = -1;
	while (choice != 0)
	{
		std::cout << "You can now perform operations on the file." << std::endl; 
		std::cout << "Please select an operation below." << std::endl;
		Activities activities;
		for(auto& operation : activities.getActivitiesList()) {
			std::cout << operation.first << ". " << operation.second << std::endl;
		}

		std::cout << "Enter your choice (0 to exit): ";
		std::cin >> choice;
	    std::string colorChoice;
		switch (choice)
		{
		   case 1:  
			   std::cout << "1. selected :: Add Design Table" << std::endl;
			   //choose colors for the design table
			   std::cout << "Please choose a color for the design table (e.g., red, blue, green): ";
			   std::cin >> colorChoice;
			   // Set the color for the design table 
			   activities.adddTableDesign(filePath,colorChoice);
			   break;
		   case 2:  
			   std::cout << "2. selected :: Add N/A to Empty Cells" << std::endl;
			   activities.addNAToEmptyCells(filePath);
			   break;
		   case 3:  
			   std::cout << "3. selected :: Get Excel Content Stats" << std::endl;
			   activities.getExcelContentStats(filePath);
			   break;
		   case 0: 
			   std::cout << "0. selected :: Exit" << std::endl;
			   break;
		  default:
			break;
		}
	} 
	std::cout << "Thank You for using my Console App :) .  Dont forget to give me a star" << std::endl;
	return 0;
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
