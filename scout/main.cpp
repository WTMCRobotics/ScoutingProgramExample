#include <SFML/Graphics.hpp>
#include <string>
#include <iostream>

#import "C:\\Program Files (x86)\\Common Files\\microsoft shared\\OFFICE12\\MSO.DLL" \
    rename( "RGB", "MSORGB" )

using namespace Office;

#import "C:\\Program Files (x86)\\Common Files\\microsoft shared\\VBA\\VBA6\\VBE6EXT.OLB"

using namespace VBIDE;

#import "C:\\Program Files (x86)\\Microsoft Office\\OFFICE12\\EXCEL.EXE" \
    rename( "DialogBox", "ExcelDialogBox" ) \
    rename( "RGB", "ExcelRGB" ) \
    rename( "CopyFile", "ExcelCopyFile" ) \
    rename( "ReplaceText", "ExcelReplaceText" ) \
    exclude( "IFont", "IPicture" ) no_dual_interfaces

const int WIN_WIDTH = 999;
const int WIN_HEIGHT = 700;

const int BORDER_WIDTH = 2;

const int HEAD_HEIGHT = 50;
const int BOX_HEIGHT = 325;
const int BOX_WIDTH = 333;

int main()
{
	sf::RenderWindow window(sf::VideoMode(WIN_WIDTH, WIN_HEIGHT), "Scouting Program", sf::Style::Close | sf::Style::Titlebar);

	sf::RectangleShape header(sf::Vector2f(WIN_WIDTH, HEAD_HEIGHT));
	header.setPosition(sf::Vector2f(0, BORDER_WIDTH));
	header.setFillColor(sf::Color::Color(248, 248, 248));
	header.setOutlineColor(sf::Color::Black);
	header.setOutlineThickness(BORDER_WIDTH);

	sf::RectangleShape r1Box(sf::Vector2f(BOX_WIDTH, BOX_HEIGHT));
	r1Box.setPosition(sf::Vector2f(0, HEAD_HEIGHT + (BORDER_WIDTH * 2)));
	r1Box.setFillColor(sf::Color::Color(248, 248, 248));
	r1Box.setOutlineColor(sf::Color::Black);
	r1Box.setOutlineThickness(BORDER_WIDTH);

	sf::RectangleShape r2Box(sf::Vector2f(BOX_WIDTH, BOX_HEIGHT));
	r2Box.setPosition(sf::Vector2f(BOX_WIDTH + BORDER_WIDTH, HEAD_HEIGHT + (BORDER_WIDTH * 2)));
	r2Box.setFillColor(sf::Color::Color(248, 248, 248));
	r2Box.setOutlineColor(sf::Color::Black);
	r2Box.setOutlineThickness(BORDER_WIDTH);

	sf::RectangleShape r3Box(sf::Vector2f(BOX_WIDTH, BOX_HEIGHT));
	r3Box.setPosition(sf::Vector2f((BOX_WIDTH + BORDER_WIDTH) * 2, HEAD_HEIGHT + (BORDER_WIDTH * 2)));
	r3Box.setFillColor(sf::Color::Color(248, 248, 248));
	r3Box.setOutlineColor(sf::Color::Black);
	r3Box.setOutlineThickness(BORDER_WIDTH);

	sf::RectangleShape b1Box(sf::Vector2f(BOX_WIDTH, BOX_HEIGHT));
	b1Box.setPosition(sf::Vector2f(0, HEAD_HEIGHT + BOX_WIDTH - BORDER_WIDTH));
	b1Box.setFillColor(sf::Color::Color(248, 248, 248));
	b1Box.setOutlineColor(sf::Color::Black);
	b1Box.setOutlineThickness(BORDER_WIDTH);

	sf::RectangleShape b2Box(sf::Vector2f(BOX_WIDTH, BOX_HEIGHT));
	b2Box.setPosition(sf::Vector2f(BOX_WIDTH + BORDER_WIDTH, HEAD_HEIGHT + BOX_WIDTH - BORDER_WIDTH));
	b2Box.setFillColor(sf::Color::Color(248, 248, 248));
	b2Box.setOutlineColor(sf::Color::Black);
	b2Box.setOutlineThickness(BORDER_WIDTH);

	sf::RectangleShape b3Box(sf::Vector2f(BOX_WIDTH, BOX_HEIGHT));
	b3Box.setPosition(sf::Vector2f((BOX_WIDTH + BORDER_WIDTH) * 2, HEAD_HEIGHT + BOX_WIDTH - BORDER_WIDTH));
	b3Box.setFillColor(sf::Color::Color(248, 248, 248));
	b3Box.setOutlineColor(sf::Color::Black);
	b3Box.setOutlineThickness(BORDER_WIDTH);

	for (int i = 0; i < 8; i++)
	{
		std::cout << "Joystick " << i << " : ";
		if (sf::Joystick::isConnected(i))
			std::cout << "true\n";
		else
			std::cout << "false\n";
	}

	// Excel Application Object
	Excel::_ApplicationPtr pXL;
	// Create Instance of an Excel Application
	pXL.CreateInstance(L"Excel.Application");
	// Do not open Excel for us to view
	pXL->PutVisible(false);
	// This will get rid of an alert that asks you if you want to overwrite when you save
	pXL->PutDisplayAlerts(false);
	// Open an Excel Workbook (file)
	Excel::_WorkbookPtr pBook = pXL->Workbooks->Open(L"C:\\Users\\Test\\Documents\\Visual Studio 2013\\Projects\\scout\\scout.xlsx");

	// Create objects for all the sheets
	Excel::_WorksheetPtr pSheetRed1 = pXL->Worksheets->GetItem("6101");
	Excel::_WorksheetPtr pSheetMatches = pXL->Worksheets->GetItem("matches");

	// Create objects for the cells in the two sheets
	Excel::RangePtr pRed1Cells = pSheetRed1->Cells;
	Excel::RangePtr pMatchesCells = pSheetMatches->Cells;	

	// Puts the value 5 in Row 3 Column 5 (of the "6101" sheet)
	pRed1Cells->PutItem(3, 5, 5);
	// Puts the value 6101 in Row 3 Column 3 (of the "matches" sheet)
	pMatchesCells->PutItem(3, 3, 6101);

	// Save both sheets
	pSheetRed1->SaveAs(L"C:\\Users\\Test\\Documents\\Visual Studio 2013\\Projects\\scout\\scout.xlsx");
	pSheetMatches->SaveAs(L"C:\\Users\\Test\\Documents\\Visual Studio 2013\\Projects\\scout\\scout.xlsx");

	// Need to use _bstr_t to convert to the right type
	// Prints what is in Row 1 Column 1
	std::cout << _bstr_t(pRed1Cells->GetItem(1, 1)) << std::endl;
	// Prints what is in Row 1 Column 3
	std::cout << _bstr_t(pMatchesCells->GetItem(1, 3)) << std::endl;

	// Need to use static_cast<int> to convert to integers for addition
	// Prints the sum of the value on the "6101" sheet in Row 3 Column 5 with the value on the "matches" sheet Row 3 Column 3
	std::cout << static_cast<int>(pRed1Cells->GetItem(3, 5)) + static_cast<int>(pMatchesCells->GetItem(3, 3)) << std::endl;

	// Important to close the workbook and the application
	pBook->Close();
	pXL->Quit();

	while (window.isOpen())
	{
		sf::Event event;
		while (window.pollEvent(event))
		{
			if (event.type == sf::Event::Closed)
				window.close();
		}

		window.clear(sf::Color::Blue);
		window.draw(header);
		window.draw(r1Box);
		window.draw(r2Box);
		window.draw(r3Box);
		window.draw(b1Box);
		window.draw(b2Box);
		window.draw(b3Box);
		window.display();
	}

	return 0;
}