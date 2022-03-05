# Excel-UDF-limitation-evasion
An add-in for MS Excel that allows you to insert tables (matrices) onto a sheet through user-defined functions.
Specifically, it provides a Sub that places a two-dimensional array on the sheet. All you have to do after installing the add-in is call this Sub with the appropriate parameters.
## Demonstration 
![demonstration](https://github.com/1azar/Excel-UDF-limitation-evasion/blob/main/demo.gif)
## Add-in implementation details
To place a two-dimensional matrix on sheet, you need to call Sub "set_mtx_on_cell" from your function with the following arguments:
 - ByRef array() as Variant - target array
 - ByRef cell_name as String - cell address of the upper left corner for the inserting matrix
 - ByRef book_name as String - the name of the book where the matrix will be inserted
 - ByRef sheet_name as String - the name of the sheet where the matrix will be inserted

Next, this Sub will store the received arguments as global variables in the "Mtx_module" module.
BOOK.cls implements the "App_AfterCalculate" method, which calls after the calculation of each function in all open books and this method has UDF no limitations.
"App_AfterCalculate" checks previously defined global variables, if they are empty, then nothing happens, otherwise it places the matrix on the sheet in the corresponding book.
At the end of the Sub, the data of these global variables will be erased.

## Installation
 - Download [Mtx_add-in.xlam](https://github.com/1azar/Excel-UDF-limitation-evasion/blob/main/Mtx_add-in.xlam).
 - Optionaly recommended to place the add-in in the `C:\..\..\AppData\Roaming\Microsoft\AddIns\` folder.
 - Open an Empty book in Excel then go to `File->Options->Add-in->Go->Browse` and select downloaded [Mtx_add-in.xlam](https://github.com/1azar/Excel-UDF-limitation-evasion/blob/main/Mtx_add-in.xlam), check the box next to the selected add-in then click OK. Enable macros if required.
 - write `=generate_matrix(3;3)` in amy cell to make sure the Add-in is working.

## Usage
There are 2 main ways to use the add-on:
1) Implementation of your custom functions inside the add-in project (for example, the demonstrative function =generate_matrix is implemented in this way). Then to insert a 2d array on a sheet, you just need to call Sub `Call set_mtx_on_cell(my_arr, Application.Caller.Address, ActiveWorkbook.Name, ActiveSheet.Name)` from any module in the project.

![demonstration](https://github.com/1azar/Excel-UDF-limitation-evasion/blob/main/%D0%A1%D0%BA%D1%80%D0%B8%D0%BD%D1%88%D0%BE%D1%82%2005-03-2022%20150740.jpg)

2) Establish a connection with the add-in and work as with a regular module. To do this, in the VBA editor insert and open your module, then go to the `Tool->Referencess..->Browse` and select the downloaded add-in (Specify "Microsoft Excel Files" in the search filter). Press ok. Now you can call add-in Sub from your book project.

![demonstration](https://github.com/1azar/Excel-UDF-limitation-evasion/blob/main/dem1.png)
![demonstration](https://github.com/1azar/Excel-UDF-limitation-evasion/blob/main/%D0%A1%D0%BA%D1%80%D0%B8%D0%BD%D1%88%D0%BE%D1%82%2005-03-2022%20152837.jpg)
  

