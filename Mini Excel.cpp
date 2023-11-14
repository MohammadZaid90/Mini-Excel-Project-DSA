#include <iostream>
#include<conio.h>
#include<Windows.h>
#include<fstream>
#include"CustomMiniExcel.h"
using namespace std;

void Header();
void gotoxy(int x, int y);
void logo();

int main()
{
    HANDLE console_color;
    console_color = GetStdHandle(STD_OUTPUT_HANDLE);


    logo();
    Sleep(2000);
    system("cls");

    SetConsoleTextAttribute(console_color, 4);

    char ope;
    cout << "Press (Y/y) For load Data & Press (N/n) for new Sheet : ";
    cin >> ope;
    CustomExcel<int> CustomExcel(ope);
    bool flag = true;
    CustomExcel.setCurrentToHeadCustom();

    while (true)
    {
        if (flag) {
            system("cls");
            Header();
            cout << "Printed Form / Edit Here\n\n";
            CustomExcel.printCustomSheet();
            flag = false;
            Sleep(50);
        }
        if (GetAsyncKeyState(VK_LEFT))
        {
            CustomExcel.moveCurrentLeftCustom();
            flag = true;
        }
        else if (GetAsyncKeyState(VK_RIGHT))
        {
            CustomExcel.moveCurrentRightCustom();
            flag = true;
        }
        else if (GetAsyncKeyState(VK_UP))
        {
            CustomExcel.moveCurrentUpCustom();
            flag = true;
        }
        else if (GetAsyncKeyState(VK_DOWN))
        {
            CustomExcel.moveCurrentDownCustom();
            flag = true;
        }
        else if (GetAsyncKeyState('A'))
        {
            CustomExcel.insertCustomColumnToLeft();
            flag = true;
        }
        else if (GetAsyncKeyState('S'))
        {
            CustomExcel.insertCustomRowBelow();
            flag = true;
        }
        else if (GetAsyncKeyState('D'))
        {
            CustomExcel.insertCustomColumnToRight();
            flag = true;
        }
        else if (GetAsyncKeyState('W'))
        {
            CustomExcel.insertCustomRowAbove();
            flag = true;
        }
        else if (GetAsyncKeyState('T'))
        {
            CustomExcel.deleteCustomColumn();
            flag = true;
        }
        else if (GetAsyncKeyState('R'))
        {
            CustomExcel.deleteCustomRow();
            flag = true;
        }
        else if (GetAsyncKeyState('I'))
        {
            int input;
            cout << "Enter value ** ";
            cin >> input;
            CustomExcel.changeCellValue(input);
            flag = true;
        }
        else if (GetAsyncKeyState('O'))
        {
            CustomExcel.storeDataCustom();
        }
    }
}

void Header()
{
    HANDLE console_color;
    console_color = GetStdHandle(STD_OUTPUT_HANDLE);
    SetConsoleTextAttribute(console_color, 3);
    cout << "                             ****************************************************" << endl;
    cout << "                             ---------------- MINI EXCEL PROJECT ----------------" << endl;
    cout << "                             ****************************************************" << endl;
    cout << "                                     WELCOME TO OUR FIRST DSA MINI PROJECT       \n\n" << endl;
}

void logo()
{
    HANDLE console_color;
    console_color = GetStdHandle(STD_OUTPUT_HANDLE);

    int x, y;
    system("cls");
    x = 9;
    y = 9;
    SetConsoleTextAttribute(console_color, 2);
    gotoxy(x, y);
    cout << "          ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";
    SetConsoleTextAttribute(console_color, 3);
    gotoxy(x, y + 1);
    cout << "          ||                                                                          ||";
    gotoxy(x, y + 2);
    cout << "          ||                                                                          ||";
    gotoxy(x, y + 3);
    cout << "          ||                                                                          ||";
    SetConsoleTextAttribute(console_color, 6);
    gotoxy(x, y + 4);
    cout << "          ||                             MINI EXCEL                                   ||";
    SetConsoleTextAttribute(console_color, 3);
    gotoxy(x, y + 5);
    cout << "          ||                                                                          ||";
    gotoxy(x, y + 6);
    cout << "          ||                                                                          ||";
    SetConsoleTextAttribute(console_color, 2);
    gotoxy(x, y + 7);
    cout << "          ||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||";


}

void gotoxy(int x, int y)
{
    COORD coordinates;
    coordinates.X = x;
    coordinates.Y = y;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
}
