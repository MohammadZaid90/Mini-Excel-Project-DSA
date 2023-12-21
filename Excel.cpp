#include <iostream>
#include <cctype>
#include <string>
#include <Windows.h>
#include <conio.h>
#include <cmath>
#include <iomanip>
#include <algorithm>
#include <cstring>
#include <vector>
#include <fstream>
#include <sstream>
#include <cstdlib>  
#ifdef _WIN32       
#endif

using namespace std;

template <typename T>
class MiniExcel
{
     struct Cell
    {
    public:
        T data;
        Cell *up;
        Cell *down;
        Cell *left;
        Cell *right;
        Cell(T val)
        {
            data = val;
            up = nullptr;
            down = nullptr;
            left = nullptr;
            right = nullptr;
        }
    };

    //----------------Iterator-----------------

    struct Iterator
    {
    public:
        Cell *iterator;

        Iterator(Cell *PresentCell)
        {
            iterator = PresentCell;
        }

        //------increment for moving right--------

        Iterator &operator++()  
        {
            iterator = iterator->right;
            return *this;
        }

        //------deccrement for moving left--------

        Iterator &operator--()
        {
            iterator = iterator->left;
            return *this; 
        }

        //------increment for moving down--------

        Iterator &operator++(int)
        {
            iterator = iterator->down;
            return *this;
        }

        //------decrement for moving up--------

        Iterator &operator--(int)
        {
            iterator = iterator->up;
            return *this;
        }

        //------returning data--------

        T &operator*()
        {
            return iterator->data;
        }

        bool operator==(const Iterator &other) const
        {
            return iterator == other.iterator;
        }

        bool operator!=(const Iterator &other) const{
            return iterator != other.iterator;
        }
    };

    Iterator begin()
    {
        Cell *temp = present;
        while(temp->left)
        {
            temp = temp->left;
        }
        while(temp->up)
        {
            temp = temp->up;
        }
        return Iterator(temp);
    }

    Iterator Stop()
    {
        return Iterator(nullptr);
    }

public:

    Cell *present;
    vector<string> CopyCutData;
    bool isColumn;
    Cell *initialCell;

    //---------------Constructor----------------

    MiniExcel()
    {
        present = new Cell("1");  //-------1st element------
        initialCell = present;
        present->left = nullptr;

        //----------------------1st row--------------------

        for(int i = 0 ; i < 4 ; i++)
        {
            Cell *FirstRow = new Cell("1");
            present->right = FirstRow;
            FirstRow->left = present;
            FirstRow->up = nullptr;
            present = FirstRow;
        }
        present = initialCell;

        //----------------------2nd row--------------------

        Cell *secondRowStart = new Cell("2");
        present->down = secondRowStart;
        secondRowStart->up = present;

        for (int i = 0; i < 4; i++)
        {
            Cell *SecondRow = new Cell("2");
            secondRowStart->right = SecondRow;
            SecondRow->left = secondRowStart;
            SecondRow->up = present;
            secondRowStart = SecondRow;
        }
        present = initialCell;

        //----------------------3rd row--------------------

        Cell *thirdRowStart = new Cell("3");
        present->down->down = thirdRowStart;
        thirdRowStart->up = present->down;

        for (int i = 0; i < 4; i++)
        {
            Cell *ThirdRow = new Cell("3");
            thirdRowStart->right = ThirdRow;
            ThirdRow->left = thirdRowStart;
            ThirdRow->up = present;
            thirdRowStart = ThirdRow;
        }
        present = initialCell;

        //----------------------4th row--------------------

        Cell *forthRowStart = new Cell("4");
        present->down->down->down = forthRowStart;
        forthRowStart->up = present->down->down;

        for (int i = 0; i < 4; i++)
        {
            Cell *ForthRow = new Cell("4");
            forthRowStart->right = ForthRow;
            ForthRow->left = forthRowStart;
            ForthRow->up = present;
            forthRowStart = ForthRow;
        }
        present = initialCell;

        //----------------------5th row--------------------

        Cell *fifthRowStart = new Cell("5");
        present->down->down->down->down = fifthRowStart;
        fifthRowStart->up = present->down->down->down;

        for (int i = 0; i < 4; i++)
        {
            Cell *FifthRow = new Cell("5");
            fifthRowStart->right = FifthRow;
            FifthRow->left = fifthRowStart;
            FifthRow->up = present;
            fifthRowStart = FifthRow;
            fifthRowStart->down = nullptr;
        }
        present = initialCell;



        Cell *presentPtr = present;       // Pointer to the current cell
        Cell *nextPtr = present->down; // Pointer to the cell in the next row

        for (int i = 1; i <= 16; i++) // Making up and down connections between 1 to 4 columns
        {
            presentPtr = presentPtr->right;
            nextPtr = nextPtr->right;

            presentPtr->down = nextPtr;
            nextPtr->up = presentPtr;

            if (i % 4 == 0)
            {
                while (presentPtr->left && nextPtr->left)
                {
                    presentPtr = presentPtr->left;
                    nextPtr = nextPtr->left;
                }
                presentPtr = presentPtr->down;
                nextPtr = nextPtr->down;
            }
        }

        present = initialCell;
    }

    //---------------insert row above----------------

    void insertRowAbove()
    {
        Cell *temp = present;
        Cell *Row = new Cell("0");
        while(temp->left)
        {
            temp = temp->left;
        }

        if(temp->up == nullptr)
        {
            Row->down = temp;
            temp->up = Row;
            Row->up = nullptr;
            while(temp->right)
            {
                Cell *rowEle = new Cell("0");
                Row->right = rowEle;
                rowEle->left = Row;
                temp = temp->right;
                temp->up = rowEle;
                rowEle->down = temp;
                Row = Row->right;
                rowEle->up = nullptr;
            }
        }
        else{
            Row->down = temp;
            Row->up = temp->up;
            temp->up->down = Row;
            temp->up = Row;

            while(temp->right)
            {
                Cell *rowEle = new Cell("0");
                rowEle->left = Row;
                Row->right = rowEle;
                temp = temp->right;
                rowEle->up = temp->up;
                temp->up->down = rowEle;
                rowEle->down = temp;
                temp->up = rowEle;
            }
        }
    }

    //---------------insert row below----------------
    void insertRowBelow()
    {
        Cell *temp = present;
        Cell *Row = new Cell("0");

        while(temp->left)
        {
            temp = temp->left;
        }

        if(temp->down == nullptr)
        {
            Row->up = temp;
            temp->down = Row;
            Row->down = nullptr;

            while(temp->right)
            {
                Cell *newRow = new Cell("0");
                newRow->left = Row;
                Row->right = newRow;
                temp = temp->right;
                temp->down = newRow;
                newRow->up = temp;
                Row = newRow;
                newRow->down = nullptr;
            }
        }
        else{
            Row->up = temp;
            Row->down = temp->down;
            temp->down->up = Row;
            temp->down = Row;
            while(temp->right)
            {
                Cell *newRow = new Cell("0");
                newRow->left = Row;
                Row->right = newRow;
                temp = temp->right;
                newRow->up = temp;
                newRow->down = temp->down;
                temp->down->up = newRow;
                Row = Row->right;
            }
        }
    }

    //---------------insert column left----------------

    void InsertColumnLeft()
    {
        Cell *presentCell = present;
        Cell *newCol = new Cell("0");

        while(presentCell->up)
        {
            presentCell = presentCell->up;
        }

        if(presentCell->left == nullptr)
        {
            newCol->right = presentCell;
            presentCell->left = newCol;
            newCol->left = nullptr;

            while(presentCell->down)
            {
                Cell *newColEle = new Cell("0");
                newCol->down = newColEle;
                newColEle->up = newCol;
                presentCell = presentCell->down;
                newColEle->right = presentCell;
                presentCell->left = newColEle;
                newColEle->left = nullptr;
                newCol = newCol->down;
            }
        }
        else{
            newCol->right = presentCell;
            newCol->left = presentCell->left;
            presentCell->left->right = newCol;
            presentCell->left = newCol;

            while(presentCell->down)
            {
                Cell *newColEle = new Cell("0");
                newCol->down = newColEle;
                newColEle->up = newCol;
                presentCell = presentCell->down;
                newColEle->right = presentCell;
                newColEle->left = presentCell->left;
                presentCell->left->right = newColEle;
                presentCell->left = newColEle;
                newCol = newCol->down;
            }
        }
    }

    //---------------insert column right----------------

    void InsertColumnRight()
    {
        Cell *temp = present;
        Cell *Col = new Cell("0");

        while(temp->up)
        {
            temp = temp->up;
        }

        if(temp->right == nullptr)
        {
            Col->left = temp;
            temp->right = Col;
            Col->right = nullptr;

            while(temp->down)
            {
                Cell *newColEle = new Cell("0");
                Col->down = newColEle;
                newColEle->up = Col;
                temp = temp->down;
                newColEle->left = temp;
                temp->right = newColEle;
                newColEle->right = nullptr;
                Col = Col->down;
            }
        }
        else{
            Col->left = temp;
            Col->right = temp->right;
            temp->right->left = Col;
            temp->right = Col;

            while(temp->down)
            {
                Cell *colEle = new Cell("0");
                Col->down = colEle;
                colEle->up = Col;
                temp = temp->down;
                colEle->left = temp;
                colEle->right = temp->right;
                temp->right->left = colEle;
                temp->right = colEle;
                Col = Col->down; 
            }  
        }
    }

    //---------------Clean row----------------

    void CleanRow()
    {
        Cell *temp = present;
        while(temp->right)
        {
            temp = temp->right;
        }
        while(temp)
        {
            temp->data = "*";
            temp = temp->left;
        }
    }

    //---------------Clean column----------------

    void CleanColumn()
    {
        Cell *temp = present;
        while(temp->up)
        {
            temp = temp->up;
        }
        while(temp)
        {
            temp->data = "*";
            temp = temp->down;
        }
    }

    //----------------------reset color-----------------------

    void SetColor(int textColor, int bgColor)
    {
        HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(hConsole, textColor | (bgColor << 4));
    }
    void ResetColor()
    {
        HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(hConsole, 15); // Assuming default color is white text on black background
    }
    void SetConsoleColor(int color)
    {
    #ifdef _WIN32
        SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), color);
    #else
        cout << "\033[1;" << color << "m"; // ANSI escape code for setting text color
    #endif
    }

    void ResetConsoleColor()
    {
    #ifdef _WIN32
        SetConsoleColor(7); // Reset to default color on Windows
    #else
        cout << "\033[0m"; // ANSI escape code for resetting text color on Unix-based systems
    #endif
    }

    //----------------------Display Sheet-----------------------

    void DisplaySheet()
    {
        SetConsoleColor(11);
        ResetConsoleColor();

        Cell *top = present;

        while(top->up)
        {
            top = top->up;
        }

        while(top->left)
        {
            top = top->left;
        }

        Cell *RowBegin = top;

        while(RowBegin)
        {
            SetConsoleColor(5);
            cout<<"|";
            ResetConsoleColor();

            Cell *ColBegin = RowBegin;

            while(ColBegin)
            {
                if(ColBegin == present)
                {
                    SetColor(23,45);
                    cout<<" "<<setw(5) << ColBegin->data <<" ";
                    ResetColor();

                    SetConsoleColor(5);
                    cout<<"|";
                    ResetConsoleColor();
                }
                else{
                    SetConsoleColor(14);
                    cout<<"|"<<setw(5)<<ColBegin->data<<" ";
                    SetConsoleColor(5);
                    cout<<"|";
                    ResetConsoleColor();
                }

                ColBegin = ColBegin->right;
            }
            cout<<endl;

            SetConsoleColor(5);
            ResetConsoleColor();

            RowBegin = RowBegin->down;
        }
    }

    //----------------------Delete Row----------------------

    void DelRow()
    {
        Cell *topCellRow = present;
        Cell *temp;

        while(topCellRow->left)
        {
            topCellRow = topCellRow->left;
        }
        temp = topCellRow;

        if(topCellRow->up == nullptr && topCellRow->down == nullptr)
        {
            return;
        }
        else if(topCellRow->up == nullptr)
        {
            present = present->down;
            while(topCellRow)
            {
                topCellRow->down->up = nullptr;
                topCellRow = topCellRow->right;
                temp = topCellRow;
            }
        }
        else if(present->down == nullptr)
        {
            present = present->up;
            while(topCellRow)
            {
                topCellRow->up->down = nullptr;
                topCellRow = topCellRow->right;
                temp = topCellRow;
            }
        }
        else{
            present = present->up;
            while(topCellRow)
            {
                temp->up->down = temp->down;
                temp->down->up = temp->up;
                topCellRow = topCellRow->right;
                temp = topCellRow;
            }
        }
    } 

    //----------------------Delete Column----------------------

    void DelCol()
    {
        Cell *topCellCol = present;
        Cell *temp = nullptr;

        while(topCellCol->up)
        {
            topCellCol = topCellCol->up;
        }
        temp = topCellCol;

        if(topCellCol->left == nullptr && topCellCol->right == nullptr)
        {
            return;
        }
        else if(topCellCol->left == nullptr)
        {
            present = present->right;
            while(topCellCol)
            {
                temp->right->left = nullptr;
                topCellCol = topCellCol->down;
                temp = topCellCol;
            }
        }
        else if(present->right == nullptr)
        {
            present = present->left;
            while(topCellCol)
            {
                temp->left->right = nullptr;
                topCellCol = topCellCol->down;
                temp = topCellCol;
            }
        }
        else{
            present = present->left;
            while(topCellCol)
            {
                temp->left->right = temp->right;
                temp->right->left = temp->left;
                topCellCol = topCellCol->down;
                temp = topCellCol;
            }
        }
    }

    //--------------------Delete by left Shift--------------------

    void DelByLeftShift()
    {
        Cell *temp = this->present;

        if(temp->right != nullptr)
        {
            Cell *next = temp->right;

            while(next != nullptr)
            {
                temp->data = next->data;
                temp = temp->right;
                next = next->right;
            }

            temp->data = "'''";
        }
    }

    //--------------------insert by right Shift--------------------

    void InsertbyRightShift()
    {
        Cell *temp = present;
        while(present->right)
        {
            present = present->right;
        }
       InsertColumnRight(); 

       while(present != temp)
       {
        present->right->data = present->data;
        present = present->left;
       }

       present->data = "'''";
    }

    //--------------------Delete by up Shift--------------------

    void DelByUpShift()
    {
        Cell *temp = this->present;
        if(temp->down != nullptr)
        {
            Cell *next = temp->down;

            while(next != nullptr)
            {
                temp->data = next->data;
                temp = temp->down;
                next = next->down;
            }
            temp->data = "'''";
        }
    }

    //--------------------inser by down Shift--------------------

    void insertByDownShift()
    {
        Cell *temp = this->present;
        while(present->down)
        {
            present = present->down;
        }
        insertRowBelow();
        while(present != temp)
        {
            present->down->data = present->data;
            present = present->up;
        }
    }

    //--------------------Check whether digit or not--------------------

    bool StringToInt(string &str)
    {
        for(char c : str)
        {
            if(!isdigit(c))
            {
                return false;
            }
        }
        return true;
    }

    //--------------------Range sum--------------------

    void RangeSum(Cell *s, Cell *e, Cell *l)
    {
        int sum = 0;

        bool flag = checkTravrseLocation(s, e);

        if (flag)
        {
            while (s != e->right)
            {
                s = s->right;
                sum += stoi(s->data);
            }
        }
        else
        {
            while (s != e->down)
            {

                sum += stoi(s->data);
                s = s->down;
            }
        }
        present = l;
        l->data = to_string(sum);
    }

    //--------------------Range average--------------------

    void RangeAverage(Cell *s, Cell *e, Cell *l)
    {
        double sum = 0;
        int count = 0;

        bool flg = checkTravrseLocation(s,e);

        if(flg)
        {
            while(s != e->right)
            {
                sum += stoi(s->data);
                count++;
                s = s->right;
            }
        }
        else{
            while(s != e->down)
            {
                sum += stoi(s->data);
                count++;
                s = s->down;
            }
        }

        sum = sum / count;
        present = l;
        l->data = to_string(sum);
    }

    //--------------------Move left--------------------

    void moveLeft()
    {
        if(present->left)
        {
            present = present->left;
        }
    }

    //--------------------Move Right--------------------

    void moveRight()
    {
        if(present->right)
        {
            present = present->right;
        }
    }

    //--------------------Move Up--------------------

    void moveUp()
    {
        if(present->up)
        {
            present = present->up;
        }
    }

    //--------------------Move Down--------------------

    void movedown()
    {
        if(present->down)
        {
            present = present->down;
        }
    }

    //--------------------Move Down--------------------

    bool checkTravrseLocation(Cell *start, Cell *stop)
    {
        while(start != nullptr)
        {
            if(start == stop)
            {
                return true;
            }
            start = start->down;
        }
        return false;
    }

    //--------------------Cut Data member--------------------

    void CutData(Cell *SCell,Cell *ECell)
    {
        CopyCutData.clear();
        bool moveToRight = checkTravrseLocation(SCell,ECell);
        isColumn = moveToRight;

        if(moveToRight)
        {
            while(SCell != ECell->right)
            {
                CopyCutData.push_back(SCell->data);
                SCell->data = "?";
                SCell = SCell->right;
            }
        }
        else{
            while(SCell != ECell->down)
            {
                CopyCutData.push_back(SCell->data);
                SCell->data = "?";
                SCell = SCell->down;
            }
        }
    }

    //--------------------Copy Data member--------------------

    void CopyData(Cell *SCell,Cell *ECell)
    {
        CopyCutData.clear();
        bool moveToRight = checkTravrseLocation(SCell,ECell);
        isColumn = moveToRight;

        if(moveToRight)
        {
            while(SCell != ECell->right)
            {
                CopyCutData.push_back(SCell->data);
                SCell = SCell->right;
            }
        }
        else{
            while(SCell != ECell->down)
            {
                CopyCutData.push_back(SCell->data);
                SCell = SCell->down;
            }
        }
    }

    //--------------------Paste DataMember--------------------

    void PasteData()
    {
        if(isColumn)
        {
            for(int i = 0 ; i < CopyCutData.size(); i++)
            {
                present->data = CopyCutData[i];
                if(present->right == nullptr && i != CopyCutData.size() -1 )
                {
                    Cell *newcell =  new Cell("0");
                    newcell->left = present;
                    present->right = newcell;
                }
                else if(present->right == nullptr && i == CopyCutData.size() - 1)
                {
                    present = present;
                }
                else{
                    present = present->right;
                }
            }
        }
        else{
            for(int i = 0 ; i < CopyCutData.size() ; i++)
            {
                present->data = CopyCutData[i];
                if(present->down == nullptr)
                {
                    Cell *newcell = new Cell("0");
                    newcell->up = present;
                    present->down = newcell;
                }
                else if(present->down == nullptr && i == CopyCutData.size() -1)
                {
                    present = present;
                }
                else{
                    present = present->down;
                }
            }
        }
    }

    //--------------------Input Data Member--------------------

    int InputCell(string val)
    {
        int Cel;
        cout<<val;
        while(_kbhit())
        {
            getch();
        }
        cin>>Cel;
        return Cel - 1;
    }

    //--------------------Change value--------------------

    void ChangeValue()
    {
        cout<<"Enter data: ";
        while(_kbhit())
        {
            getch();
        }
        cin >> present->data;
    }

    //--------------------Header--------------------

    void printHeader()
    {
            //color theme blue
            HANDLE console_color;
            console_color = GetStdHandle(STD_OUTPUT_HANDLE);
            SetConsoleTextAttribute(console_color, 3);

            cout << "                             ****************************************************" << endl;
            cout << "                             ---------------- MINI EXCEL PROJECT ----------------" << endl;
            cout << "                             ****************************************************" << endl;
            cout << "                                     WELCOME TO OUR FIRST DSA MINI PROJECT       " << endl;
            cout << "                                       Code By Mohammad Zaid & Furqan            \n\n" <<endl;       
    }

    //--------------------Menu--------------------

    void Menu()
    {
        //Color theme red
        system("cls");
        HANDLE console_color;
        console_color = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(console_color, 4);

        cout << "                                             ------------------------------" << endl;
        cout << "                                             |    Excel Simulation Menu   |" << endl;
        cout << "                                             ------------------------------" << endl<<endl;
        cout << "A. Arrow Up        * Move Up" << endl;
        cout << "B. Arrow Down      * Move Down" << endl;
        cout << "C. Arrow Left      * Move Left" << endl;
        cout << "D. Arrow Right     * Move Right" << endl;
        cout << "E. Ctrl + Up       * Insert Row To Above" << endl;
        cout << "F. Ctrl + Down     * Insert Row To Below" << endl;
        cout << "G. Ctrl + Right    * Insert Column To Right" << endl;
        cout << "H. Ctrl + L        * Insert Column To Left" << endl;
        cout << "I. Shift + R       * Insert Cell By Right Shift" << endl;
        cout << "J. Shift + D       * Insert Cell By Down Shift" << endl;
        cout << "K. Delete + L      * Delete Cell By Left Shift" << endl;
        cout << "L. Delete + U      * Delete Cell By Up Shift" << endl;
        cout << "M. Delete + C      * Delete Column" << endl;
        cout << "N. Delete + R      * Delete Row" << endl;
        cout << "O. Tab + C         * Clear Column" << endl;
        cout << "P. Tab + R         * Clear Row" << endl;
        cout << "Q. Shift + S       * Get Range Sum" << endl;
        cout << "R. Shift + A       * Get Range Average" << endl;
        cout << "S. Ctrl + W        * Copy Range" << endl;
        cout << "T. Ctrl + X        * Cut Range" << endl;
        cout << "U. Ctrl + V        * Paste Range" << endl;
        cout << "V. Ctrl + S        * Save Sheet" << endl;
        cout << "W. Insert          * Change Cell Data" << endl;
        cout << "X. Esc             * Exit" << endl;
    }

    //--------------------Get Cell--------------------

     Cell *GetCell(int row, int column)
    {
        Cell *temp = present;

        for (int i = 0; i < row && temp->down; i++)
        {

            temp = temp->down;
        }

        for (int i = 0; i < column && temp->right; i++)
        {
            temp = temp->right;
        }

        return temp;
    }

    //--------------------Load Data--------------------

    void LoadData()
    {
        ifstream file("Excel.txt", ios::in);
        if (!file.is_open())
        {
            cerr << "No Data Load Yet!!!!" << endl;
            return;
        }

        Cell *temCell = nullptr;
        Cell *temRow = nullptr;

        string line;
        while (getline(file, line))
        {
            vector<string> celValues;
            splitString(line, celValues, ',');

            while (!celValues.empty())
            {
                Cell *celElement = new Cell(celValues.front()); 
                celValues.erase(celValues.begin());             

                if (temCell == nullptr && temRow == nullptr)
                {
                    temCell = celElement;
                    temRow = temCell;
                    present = temCell;
                    initialCell = temCell;
                }
                else if (temCell == nullptr && temRow != nullptr)
                {
                    temCell = celElement;
                    temRow->down = celElement;
                    celElement->up = temRow;
                    temRow = temRow->down;
                }
                else
                {
                    temCell->right = celElement;
                    celElement->left = temCell;
                    temCell = temCell->right;
                }
            }
            temCell = nullptr;
        }

        Cell *conec1 = initialCell;
        Cell *conec2 = initialCell->down;

        Cell *temConec1 = conec1;
        Cell *temConec2 = conec2;

        while (conec2 != nullptr)
        {
            temConec1 = conec1;
            temConec2 = conec2;
            while (temConec2 != nullptr)
            {
                temConec1->down = temConec2;
                temConec2->up = temConec1;
                temConec1 = temConec1->right;
                temConec2 = temConec2->right;
            }
            conec1 = conec2;
            conec2 = conec2->down;
        }

        file.close();
    }

    //--------------------Save File--------------------

    void SaveFile()
    {
        fstream file;
        file.open("Excel.txt" , ios::out);

        for(Iterator it = begin(); it!= Stop(); it++)
        {
            bool firstElement = true;
            for(Iterator t = it; t != Stop(); ++t)
            {
                if(!firstElement)
                {
                    file << ",";
                }
                else{
                    firstElement = false;
                }
                file << *t;
            }
            file << endl;
        }
        file.close();
    }

    //--------------------Splitting --------------------

    void splitString(const string &inputString, vector<string> &outputVector, char delimiter)
    {
        string token;
        for (char c : inputString)
        {
            if (c == delimiter)
            {
                outputVector.push_back(token);
                token.clear();
            }
            else
            {
                token += c;
            }
        }
        if (!token.empty())
        {
            outputVector.push_back(token);
        }
    }
    //--------------------comma separation from file --------------------

    T parsing(const string &line, int num)
    {
        int count = 1;
        string temp;
        for (char c : line)
        {
            if (c == ',')
            {
                count++;
            }
            else if (count == num)
            {
                temp += c;
            }
        }
        try
        {
            return stoi(temp);
        }
        catch (const invalid_argument &e)
        {
            cerr << "Invalid argument: " << e.what() << endl;
            // Handling the error, possibly by returning a special value or throwing a different exception.
            return T(); // Returning a default value for T
        }
    }

    //--------------------x and y position on console --------------------

    void gotoxy(int x, int y)
    {
        COORD coordinates;
        coordinates.X = x;
        coordinates.Y = y;
        SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
    }

    //--------------------Front Display--------------------

    void logoFront()
    {
        HANDLE console_color;
        console_color = GetStdHandle(STD_OUTPUT_HANDLE);

        int x, y;
        system("cls");
        x = 9;
        y = 9;
        SetConsoleTextAttribute(console_color, 2);
        gotoxy(x, y);
        cout << "           |#####################################^#########################################|";
        SetConsoleTextAttribute(console_color, 3);
        gotoxy(x, y + 1);
        cout << "          /*/                                                                             /*/";
        gotoxy(x, y + 2);
        cout << "          /*/                                                                             /*/";
        gotoxy(x, y + 3);
        cout << "          /*/                                                                             /*/";
        SetConsoleTextAttribute(console_color, 6);
        gotoxy(x, y + 4);
        cout << "          /$/                                MINI EXCEL                                   /$/";
        SetConsoleTextAttribute(console_color, 3);
        gotoxy(x, y + 5);
        cout << "          /*/                                                                             /*/";
        gotoxy(x, y + 6); 
        cout << "          /*/                                                                             /*/";
        gotoxy(x, y + 7);
        cout << "          /*/                                                                             /*/";
        SetConsoleTextAttribute(console_color, 2);
        gotoxy(x, y + 8);
        cout << "           |#####################################v#########################################|";


    }

    //--------------------Display Option--------------------

    string option()
    {
        string op = "";
        printHeader();

        HANDLE console_color;
        console_color = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(console_color, 8);

        cout<<"1. Mini Excel Sheet.... "<<endl;
        cout<<"2. Mini Excel Menu.... "<<endl;
        cout<<"3. Exit...."<<endl;
        cout<<"Enter any Option.....! ";
        cin>>op;
        return op;
    }
};

int main()
{
    MiniExcel<string> miniexcel;
    miniexcel.logoFront();
    Sleep(1000);
    system("cls");
    miniexcel.LoadData();

    string op = "";

    do
    {
        system("cls");
        op = miniexcel.option();

        if(op == "1")
        {
            system("cls");
            miniexcel.printHeader();
            miniexcel.DisplaySheet();

            bool running = true;

            while(running)
            {

                if (GetAsyncKeyState(VK_UP) && 1)
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.moveUp();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_DOWN))
                {

                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.movedown();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_LEFT))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.moveLeft();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_RIGHT))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.moveRight();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState(VK_UP))
                {

                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.insertRowAbove();
                    miniexcel.DisplaySheet();
                }
        
                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState(VK_DOWN))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.insertRowBelow();
                    miniexcel.DisplaySheet();
                }
    
                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState(VK_RIGHT))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.InsertColumnRight();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState(VK_LEFT))
                {

                    system("cls");
                    miniexcel.InsertColumnLeft();
                    miniexcel.printHeader();
                    miniexcel.DisplaySheet();
                }
            
                if (GetAsyncKeyState(VK_SHIFT) & GetAsyncKeyState('R'))
                {

                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.InsertbyRightShift();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_SHIFT) & GetAsyncKeyState('D'))
                {

                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.insertByDownShift();
                    miniexcel.DisplaySheet();
                }
            
                if (GetAsyncKeyState(VK_DELETE) & GetAsyncKeyState('L'))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.DelByLeftShift();
                    miniexcel.DisplaySheet();
                }
            
                if (GetAsyncKeyState(VK_DELETE) & GetAsyncKeyState('U'))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.DelByUpShift();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_DELETE) & GetAsyncKeyState('C'))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.DelCol();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_DELETE) & GetAsyncKeyState('R'))
                {

                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.DelRow();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_TAB) & GetAsyncKeyState('C'))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.CleanColumn();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_TAB) & GetAsyncKeyState('R'))
                {

                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.CleanRow();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_SHIFT) & GetAsyncKeyState('S'))
                {
                    int SRow = miniexcel.InputCell("Enter begining row: ");
                    int SCol = miniexcel.InputCell("Enter beginning column: ");

                    int ERow = miniexcel.InputCell("Enter ending  row: ");
                    int ECol = miniexcel.InputCell("Enter ending  column: ");

                    int destinationRows = miniexcel.InputCell("Enter destination row: ");
                    int destinationColum = miniexcel.InputCell("Enter destination column: ");

                    miniexcel.RangeSum(miniexcel.GetCell(SRow, SCol), miniexcel.GetCell(ERow, ECol), miniexcel.GetCell(destinationRows, destinationColum));
 
                    system("cls");

                    miniexcel.printHeader();
                    miniexcel.DisplaySheet();
                }

                
                if (GetAsyncKeyState(VK_SHIFT) & GetAsyncKeyState('A'))
                {

                    int SRow = miniexcel.InputCell("Enter begining row: ");
                    int SCol = miniexcel.InputCell("Enter beginning column: ");

                    int ERow = miniexcel.InputCell("Enter ending  row: ");
                    int ECol = miniexcel.InputCell("Enter ending  column: ");

                    int destinationRows = miniexcel.InputCell("Enter destination row: ");
                    int destinationColum = miniexcel.InputCell("Enter destination column: ");

                    miniexcel.RangeAverage(miniexcel.GetCell(SRow, SCol), miniexcel.GetCell(ERow, ECol), miniexcel.GetCell(destinationRows, destinationColum));
 
                    system("cls");

                    miniexcel.printHeader();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState('W')) // we use ctrl + c for copy cammand but ctrl + c terminate the program so we use ctrl + w
                {
                    
                    int SRow = miniexcel.InputCell("Enter begining row: ");
                    int SCol = miniexcel.InputCell("Enter beginning column: ");

                    int ERow = miniexcel.InputCell("Enter ending  row: ");
                    int ECol = miniexcel.InputCell("Enter ending  column: ");

                    miniexcel.CopyData(miniexcel.GetCell(SRow, SCol), miniexcel.GetCell(ERow, ECol));
 
                    system("cls");

                    miniexcel.printHeader();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState('X'))
                {

                    int SRow = miniexcel.InputCell("Enter begining row: ");
                    int SCol = miniexcel.InputCell("Enter beginning column: ");

                    int ERow = miniexcel.InputCell("Enter ending  row: ");
                    int ECol = miniexcel.InputCell("Enter ending  column: ");

                    miniexcel.CutData(miniexcel.GetCell(SRow, SCol), miniexcel.GetCell(ERow, ECol));
 
                    system("cls");

                    miniexcel.printHeader();
                    miniexcel.DisplaySheet();
                }
                
                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState('V'))
                {
                    miniexcel.PasteData();
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_CONTROL) & GetAsyncKeyState('S'))
                {

                    miniexcel.SaveFile();
                }

                if (GetAsyncKeyState(VK_INSERT))
                {
                    system("cls");
                    miniexcel.printHeader();
                    miniexcel.ChangeValue();
                    miniexcel.DisplaySheet();
                }

                if (GetAsyncKeyState(VK_ESCAPE))
                {
                    break;
                }


                Sleep(200);
            }
        }
        else if(op == "2")
        {
            system("cls");
            miniexcel.Menu();
            getch();
        }
        else if(op == "3")
        {
            system("cls");

            HANDLE console_color;
            console_color = GetStdHandle(STD_OUTPUT_HANDLE);
            SetConsoleTextAttribute(console_color, 14);

            miniexcel.gotoxy(37,9+4);
            cout<<"Thanks For using Our excel Simulatore!!!"<<endl;
            Sleep(1500);
            system("cls");
        }
        else
        {
            cout<<"\n                   Enter Please Valid Option !!!";
            getch();
        }
    }
    while(op != "3");
    return 0;
}