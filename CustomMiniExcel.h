#pragma once
#include <iostream>
#include <fstream>
#include <vector>
#include <string>
using namespace std;

//Making of Excel Class
template<typename DataType>
class CustomExcel {
private:
    //Making of Cell Class
    class CustomCell {
    public:
        DataType value;
        CustomCell* next;
        CustomCell* prev;
        CustomCell* up;
        CustomCell* down;

        CustomCell(DataType value) {
            this->value = value;
            this->next = nullptr;
            this->prev = nullptr;
            this->up = nullptr;
            this->down = nullptr;
        }
    };

    CustomCell* sheetHead;
    size_t numRows;
    size_t numCols;
    CustomCell* temp;
    string currentIndex;

public:
    CustomCell* currentCellPtr;
    class CustomIterator {
    public:
        CustomCell* temp;
        CustomIterator(CustomCell* temp) {
            this->temp = temp;
        }

        bool operator != (CustomCell* check) {
            return check != temp;
        }

        void operator ++ () {
            temp = temp->next;
        }

        void operator -- () {
            temp = temp->prev;
        }

        void operator -- (int) {
            temp = temp->up;
        }

        void operator ++ (int) {
            temp = temp->down;
        }
    };

    CustomExcel(char operation) {
        if (operation == 'Y' || operation == 'y') {
            loadDataCustom();
            return;
        }

        CustomCell* hd = new CustomCell(1);
        currentCellPtr = temp = sheetHead = hd;
        numRows = numCols = 5;
        currentIndex = "1,1";

        for (int a = 1; a <= numRows; a++) {
            for (int b = 1; b < numCols; b++) {
                insertCustomRight(b + 1);
                temp = temp->next;
            }
            if (a != numRows) {
                temp = sheetHead;
                for (int k = 1; k < a; k++) {
                    temp = temp->down;
                }
                insertCustomDown(a + 1);
                temp = temp->down;
            }
        }
    }
    // print the sheet 
    void printCustomSheet() {
        CustomCell* newtemp = currentCellPtr;
        CustomCell* temp = sheetHead;
        for (int i = 0; i < numRows; i++) {
            for (int j = 0; j < numCols; j++) {
                if (temp == newtemp) {
                    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
                    int k = 12;
                    SetConsoleTextAttribute(hConsole, k);
                    cout << temp->value << "  ";
                    k = 7;
                    SetConsoleTextAttribute(hConsole, k);
                }
                else
                    cout << temp->value << "  ";
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 0; k <= i; k++) {
                temp = temp->down;
            }
            cout << endl;
        }
        cout << endl;
    }

    //insert at right postiion
    void insertCustomRight(DataType value) {
        CustomCell* newCell = new CustomCell(value);

        newCell->prev = temp;


        if (temp->next == nullptr)
            newCell->next = nullptr;
        else {
            newCell->next = temp->next;
            temp->next->prev = newCell;
        }
        if (temp->up == nullptr)
            newCell->up = nullptr;
        else {
            newCell->up = temp->up->next;
            temp->up->next->down = newCell;
        }
        if (temp->down == nullptr || temp->down->next == nullptr)
            newCell->down = nullptr;
        else {
            newCell->down = temp->down->next;
            temp->down->next->up = newCell;
        }
        temp->next = newCell;
    }

    // insert at left positioin
    void insertCustomLeft(DataType value) {
        CustomCell* newCell = new CustomCell(value);
        newCell->next = temp;
        if (temp->prev == nullptr)
            newCell->prev = nullptr;
        else {
            newCell->prev = temp->prev;
            temp->prev->next = newCell;
        }
        if (temp->up == nullptr)
            newCell->up = nullptr;
        else {
            newCell->up = temp->up->prev;
            temp->up->prev->down = newCell;
        }
        if (temp->down == nullptr || temp->down->prev == nullptr)
            newCell->down = nullptr;
        else {
            newCell->down = temp->down->prev;
            temp->down->prev->up = newCell;
        }
        temp->prev = newCell;
    }

    //insert down
    void insertCustomDown(DataType value) {
        CustomCell* newCell = new CustomCell(value);
        newCell->up = temp;
        if (temp->down == nullptr)
            newCell->down = nullptr;
        else {
            newCell->down = temp->down;
            temp->down->up = newCell;
        }
        if (temp->next == nullptr || temp->next->down == nullptr)
            newCell->next = nullptr;
        else {
            newCell->next = temp->next->down;
            temp->next->down->prev = newCell;
        }
        if (temp->prev == nullptr)
            newCell->prev = nullptr;
        else {
            newCell->prev = temp->prev->down;
            temp->prev->down->next = newCell;
        }
        temp->down = newCell;
    }

    //insert above
    void insertCustomAbove(DataType value) {
        CustomCell* newCell = new CustomCell(value);
        newCell->down = temp;
        if (temp->up == nullptr)
            newCell->up = nullptr;
        else {
            newCell->up = temp->up;
            temp->up->down = newCell;
        }
        if (temp->next == nullptr || temp->next->up == nullptr)
            newCell->next = nullptr;
        else {
            newCell->next = temp->next->up;
            temp->next->up->prev = newCell;
        }
        if (temp->prev == nullptr)
            newCell->prev = nullptr;
        else {
            newCell->prev = temp->prev->up;
            temp->prev->up->next = newCell;
        }
        temp->up = newCell;
    }

    //change cellvalue causes a change before storing the value
    void changeCellValue(DataType value) {
        currentCellPtr->value = value;
    }

    //clear the required columnn
    void clearCustomRow(int row) {
        for (int i = 1; i <= numCols; i++) {
            changeCellValue(row, i, 0);
        }
    }

    //clear the required row
    void clearCustomColumn(int col) {
        for (int i = 1; i <= numRows; i++) {
            changeCellValue(i, col, 0);
        }
    }

    //insert a new column to right postion
    void insertCustomColumnToRight() {
        temp = sheetHead;
        size_t col = currentIndex[2] - '0';
        for (int i = 0; i < col - 1; i++) {
            temp = temp->next;
        }
        for (int i = 0; i < numRows; i++) {
            insertCustomRight(0);
            temp = temp->down;
        }
        numCols++;
    }

    //insert a new column to left postion
    void insertCustomColumnToLeft() {
        temp = sheetHead;
        size_t col = currentIndex[2] - '0';
        for (int i = 0; i < col - 1; i++) {
            temp = temp->next;
        }
        for (int i = 0; i < numRows; i++) {
            insertCustomLeft(0);
            temp = temp->down;
        }
        numCols++;
    }

    //insert a new row below the current
    void insertCustomRowBelow() {
        temp = sheetHead;
        size_t row = currentIndex[0] - '0';
        for (int i = 0; i < row - 1; i++) {
            temp = temp->down;
        }
        for (int i = 0; i < numCols; i++) {
            insertCustomDown(0);
            temp = temp->next;
        }
        numRows++;
    }

    //insert a new row above the current row
    void insertCustomRowAbove() {
        temp = sheetHead;
        size_t row = currentIndex[0] - '0';
        for (int i = 0; i < row - 1; i++) {
            temp = temp->down;
        }
        for (int i = 0; i < numCols; i++) {
            insertCustomAbove(0);
            temp = temp->next;
        }
        numRows++;
    }

    //delete the selected row
    void deleteCustomRow() {
        temp = sheetHead;
        size_t row = currentIndex[0] - '0';
        for (int i = 0; i < row - 2; i++) {
            temp = temp->down;
        }
        if (row == 1) {
            sheetHead = sheetHead->down;
        }
        else if (row == numRows) {
            for (int i = 0; i < numCols; i++) {
                temp->down = nullptr;
                temp = temp->next;
            }
        }
        else {
            for (int i = 0; i < numCols; i++) {
                temp->down->down->up = temp;
                temp->down = temp->down->down;

                temp = temp->next;
            }
        }
        numRows--;
    }

    //delete the selected column
    void deleteCustomColumn() {
        temp = sheetHead;
        size_t col = currentIndex[2] - '0';
        for (int i = 0; i < col - 2; i++) {
            temp = temp->next;
        }
        if (col == 1) {
            sheetHead = sheetHead->next;
        }
        else if (col == numCols) {
            for (int i = 0; i < numRows; i++) {
                temp->next = nullptr;
                temp = temp->down;
            }
        }
        else {
            for (int i = 0; i < numRows; i++) {
                temp->next->next->prev = temp;
                temp->next = temp->next->next;

                temp = temp->down;
            }
        }
        numCols--;
    }

    //return the sum of required range
    int getCustomRangeSum(string startingIndex, string endingIndex) {
        int sum = 0;
        bool flag = false;
        size_t Rstart, Cstart, Rend, Cend;
        Rstart = startingIndex[0] - '0';
        Cstart = startingIndex[2] - '0';
        Rend = endingIndex[0] - '0';
        Cend = endingIndex[2] - '0';
        temp = sheetHead;
        for (int i = 1; i <= numRows; i++) {
            for (int j = 1; j <= numCols; j++) {
                if (i == Rend && Cend == j)
                    return sum;
                else if (i == Rstart && Cstart == j)
                    flag = true;
                if (flag)
                    sum += temp->value;
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 1; k <= i; k++) {
                temp = temp->down;
            }
        }
    }

    // total sum

    int Sum(string startingIndex, string endingIndex) {
        int sum = 0;
        size_t Rstart, Cstart, Rend, Cend;
        Rstart = startingIndex[0] - '0';
        Cstart = startingIndex[2] - '0';
        Rend = endingIndex[0] - '0';
        Cend = endingIndex[2] - '0';
        setCurrentCellCustom(startingIndex);
        for (int i = Rstart; i <= Rend; i++) {
            for (int j = 1; j <= numCols; j++) {
                if (i == Rend && j == Cend) return sum;
                sum += temp->value;
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 1; k <= i; k++) {
                temp = temp->down;
            }
        }
        return sum;
    }

    //calculate the required average
    float getCustomRangeAverage(string startingIndex, string endingIndex) {
        float sum = 0, count = 0;
        bool flag = false;
        size_t Rstart, Cstart, Rend, Cend;
        Rstart = startingIndex[0] - '0';
        Cstart = startingIndex[2] - '0';
        Rend = endingIndex[0] - '0';
        Cend = endingIndex[2] - '0';
        temp = sheetHead;
        for (int i = 1; i <= numRows; i++) {
            for (int j = 1; j <= numCols; j++) {
                if (i == Rend && Cend == j)
                    return sum / count;
                else if (i == Rstart && Cstart == j)
                    flag = true;
                if (flag) {
                    sum += temp->value;
                    count++;
                }
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 1; k <= i; k++) {
                temp = temp->down;
            }
        }
        return sum / count;
    }

    //give the minimum number
    int getCustomRangeMinimum(string startingIndex, string endingIndex) {
        int min = INT16_MAX;
        bool flag = false;
        size_t Rstart, Cstart, Rend, Cend;
        Rstart = startingIndex[0] - '0';
        Cstart = startingIndex[2] - '0';
        Rend = endingIndex[0] - '0';
        Cend = endingIndex[2] - '0';
        temp = sheetHead;
        for (int i = 1; i <= numRows; i++) {
            for (int j = 1; j <= numCols; j++) {
                if (i == Rend && Cend == j)
                    return min;
                else if (i == Rstart && Cstart == j)
                    flag = true;
                if (flag && min > temp->value) {
                    min = temp->value;
                }
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 1; k <= i; k++) {
                temp = temp->down;
            }
        }
        return min;
    }

    //calculate the maximum number
    int getCustomRangeMaximum(string startingIndex, string endingIndex) {
        int max = 0;
        bool flag = false;
        size_t Rstart, Cstart, Rend, Cend;
        Rstart = startingIndex[0] - '0';
        Cstart = startingIndex[2] - '0';
        Rend = endingIndex[0] - '0';
        Cend = endingIndex[2] - '0';
        temp = sheetHead;
        for (int i = 1; i <= numRows; i++) {
            for (int j = 1; j <= numCols; j++) {
                if (i == Rend && Cend == j)
                    return max;
                else if (i == Rstart && Cstart == j)
                    flag = true;
                if (flag && max < temp->value) {
                    max = temp->value;
                }
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 1; k <= i; k++) {
                temp = temp->down;
            }
        }
        return max;
    }

    //copies the number
    void copyCustom(string startingIndex, string endingIndex) {
        bool flag = false;
        vector<DataType> vec;
        size_t Rstart, Cstart, Rend, Cend;
        Rstart = startingIndex[0] - '0';
        Cstart = startingIndex[2] - '0';
        Rend = endingIndex[0] - '0';
        Cend = endingIndex[2] - '0';
        temp = sheetHead;
        for (int i = 1; i <= numRows; i++) {
            for (int j = 1; j <= numCols; j++) {
                if (i == Rend && Cend == j)
                    return;
                else if (i == Rstart && Cstart == j)
                    flag = true;
                if (flag) {
                    vec.push_back(temp->value);
                }
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 1; k <= i; k++) {
                temp = temp->down;
            }
        }
    }

    //cut the required number
    void cutCustom(string startingIndex, string endingIndex) {
        bool flag = false;
        vector<DataType> vec;
        size_t Rstart, Cstart, Rend, Cend;
        Rstart = startingIndex[0] - '0';
        Cstart = startingIndex[2] - '0';
        Rend = endingIndex[0] - '0';
        Cend = endingIndex[2] - '0';
        temp = sheetHead;
        for (int i = 1; i <= numRows; i++) {
            for (int j = 1; j <= numCols; j++) {
                if (i == Rend && Cend == j)
                    return;
                else if (i == Rstart && Cstart == j)
                    flag = true;
                if (flag) {
                    vec.push_back(temp->value);
                    temp->value = 0;
                }
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 1; k <= i; k++) {
                temp = temp->down;
            }
        }
    }

    //set the value for the current cell
    void setCurrentCellCustom(string index) {
        size_t rows, cols;
        rows = index[0] - '0';
        cols = index[2] - '0';
        int i = 1, j = 1;
        temp = sheetHead;
        while (i < rows) {
            temp = temp->down;
            i++;
        }
        while (j < cols) {
            temp = temp->next;
            j++;
        }
    }

    //insert at the cell by right shift
    void insertCellByRightShiftCustom(string index) {
        setCurrentCellCustom(index);
        insertCustomLeft(9);
        temp = sheetHead;
        for (int i = 0; i < numCols - 1; i++) {
            temp = temp->next;
        }
        for (int i = 0; i < numRows; i++) {
            if (temp->next == nullptr)
                insertCustomRight(0);
            temp = temp->down;
        }
        numCols++;
    }

    //insert the cell by down shift
    void insertCellByDownShiftCustom(string index) {
        setCurrentCellCustom(index);
        insertCustomAbove(9);
        temp = sheetHead;
        for (int i = 0; i < numRows - 1; i++) {
            temp = temp->down;
        }
        for (int i = 0; i < numCols; i++) {
            if (temp->down == nullptr)
                insertCustomDown(0);
            temp = temp->next;
        }
        numRows++;
    }

    //for moving up
    void moveCurrentUpCustom() {
        if (currentCellPtr->up != nullptr) {
            currentCellPtr = currentCellPtr->up;
            int Ri = currentIndex[0] - '0';
            int Ci = currentIndex[2] - '0';
            Ri--;
            currentIndex = to_string(Ri) + ',' + to_string(Ci);
        }
    }

    //for moving down
    void moveCurrentDownCustom() {
        if (currentCellPtr->down != nullptr) {
            currentCellPtr = currentCellPtr->down;
            int Ri = currentIndex[0] - '0';
            int Ci = currentIndex[2] - '0';
            Ri++;
            currentIndex = to_string(Ri) + ',' + to_string(Ci);
        }
        else
            insertCustomRowBelow();
    }

    //for moving left
    void moveCurrentLeftCustom() {
        if (currentCellPtr->prev != nullptr) {
            currentCellPtr = currentCellPtr->prev;
            int Ri = currentIndex[0] - '0';
            int Ci = currentIndex[2] - '0';
            Ci--;
            currentIndex = to_string(Ri) + ',' + to_string(Ci);
        }
    }

    //for moving right
    void moveCurrentRightCustom() {
        if (currentCellPtr->next != nullptr) {
            currentCellPtr = currentCellPtr->next;
            int Ri = currentIndex[0] - '0';
            int Ci = currentIndex[2] - '0';
            Ci++;
            currentIndex = to_string(Ri) + ',' + to_string(Ci);
        }
        else
            insertCustomColumnToRight();
    }

    //makes the current as a head in the sheet
    void setCurrentToHeadCustom() {
        currentCellPtr = sheetHead;
    }

    //stores the data in the file i.e File Handling
    void storeDataCustom() {
        temp = sheetHead;
        fstream files;
        files.open("custom_data.txt", ios::out);
        files << numRows << "," << numCols << endl;
        for (int i = 0; i < numRows; i++) {
            for (int j = 0; j < numCols; j++) {
                files << temp->value << ",";
                temp = temp->next;
            }
            temp = sheetHead;
            for (int k = 0; k <= i; k++) {
                temp = temp->down;
            }
            files << endl;
        }
        files.close();
    }

    //load from the file
    void loadDataCustom() {
        string data;
        fstream files;
        files.open("custom_data.txt", ios::in);
        files >> data;
        size_t rs = stoi(getFieldCustom(data, 1));
        size_t cs = stoi(getFieldCustom(data, 2));
        numRows = rs;
        numCols = cs;
        files >> data;
        CustomCell* hd = new CustomCell(stoi(getFieldCustom(data, 1)));
        currentCellPtr = temp = sheetHead = hd;
        currentIndex = "1,1";
        int i = 1;
        while (!files.eof()) {
            if (data == "") break;
            cout << data << endl;
            for (int j = 1; j < cs; j++) {
                insertCustomRight(stoi(getFieldCustom(data, j + 1)));
                temp = temp->next;
            }
            data = "";
            files >> data;
            if (i != rs) {
                temp = sheetHead;
                for (int k = 1; k < i; k++) {
                    temp = temp->down;
                }
                cout << "DATA : " << getFieldCustom(data, 1);
                insertCustomDown(stoi(getFieldCustom(data, 1)));
                temp = temp->down;
            }
            i++;
        }
    }

    //for getting
    string getFieldCustom(string record, int field) {
        int commaCount = 1;
        string item = "";
        for (int x = 0; x < record.length(); x++) {
            if (record[x] == ',') {
                commaCount = commaCount + 1;
            }
            else if (commaCount == field) {
                item = item + record[x];
            }
        }
        return item;
    }
};
