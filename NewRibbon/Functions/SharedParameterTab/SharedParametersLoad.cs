﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WF = System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
//usingi do excela
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
//usingi do excela
namespace SharedParametersLoad

{

    [Transaction(TransactionMode.Manual)]

    [Regeneration(RegenerationOption.Manual)]

    public class SharedParam: IExternalCommand

    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)

        {

            //Get application and documnet objects

            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            Application app = doc.Application;
            try
            {
                //DO KLASY readSP
                //Open Shared Parameter file and collect SharedParam ExteralDefinition
                List<ExternalDefinition> spFileExternalDef = new List<ExternalDefinition>();
                DefinitionFile spFile = doc.Application.OpenSharedParameterFile();
                DefinitionGroups gr = spFile.Groups;

                    foreach (DefinitionGroup dG in gr)
                    {
                            foreach (ExternalDefinition eD in dG.Definitions)
                            {
                                spFileExternalDef.Add(eD);
                            }
                    }
                //@DO KLASY
                //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                //TODO:DO KLASY readExcel
                //Load ExcelFile and Categories to assign
                string spFileName = spFile.Filename;
                string excelFileName = "";
                TaskDialog.Show("Info", "Wybierz plik excel z kategoriami przypisaymi do parametrów");

                WF.OpenFileDialog openFileDialog1 = new WF.OpenFileDialog();

                openFileDialog1.InitialDirectory = "c:\\";
                openFileDialog1.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|.xlsx;*.xlsm;*.xls";
                openFileDialog1.RestoreDirectory = true;
                
                if (openFileDialog1.ShowDialog() == WF.DialogResult.OK)
                {
                    excelFileName = openFileDialog1.FileName;
                }

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                List<List<string>> categoryListFromExcel = new List<List<string>>();

                //wartości odejmowane od iteratorów wynikają z konstrukcji przykłądowego excela -> łatwiej opóścić od razu niepotrzebne kolumny i wiersze
                for (int i = 8; i <= rowCount; i++)
                {
                    List<string> sublist = new List<string>();
                    for (int j = 11; j <= colCount; j++)
                    {
                        sublist.Add(xlRange.Cells[i, j].Value);
                        sublist.RemoveAll(item => item == null);
                    }
                    categoryListFromExcel.Add(sublist);  
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                //@DO KLASY readExcel
                //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                //TODO:DO KLASY category list
                Categories categories = doc.Settings.Categories;

                SortedList<string, Category> allCategories = new SortedList<string, Category>();

                foreach (Category c in categories)
                {
                    allCategories.Add(c.Name, c);
                }

                List<List<Category>> categoryList = new List<List<Category>>();
                foreach (List<string> sub in categoryListFromExcel)
                {
                    List<Category> sublist = new List<Category>();
                    foreach (string cat in sub)
                    {
                        sublist.Add(allCategories[cat]);
                    }
                    categoryList.Add(sublist);
                }

                //@DO KLASY category list
                //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                //TODO:(3) generalnie to trzeba dodać jeszcze iterator do wyboru pomiędzy parametrem Type i Instance i ewentualnie grupę do któej ma zostać dodany parametr
                //TODO:(3) trzeba zobaczyc co jest z ... Array group = Enum.GetValues(typeof(BuiltInParameterGroup));
                Transaction trans = new Transaction(doc, "Dodanie SP");

                trans.Start();
                
                    List<CategorySet> catSetList = new List<CategorySet>();
                    BindingMap bindMap = doc.ParameterBindings;
                    for (int i = 0; i < spFileExternalDef.Count; i++)
                    {
                        CategorySet catSet = app.Create.NewCategorySet();
                        foreach (Category n in categoryList[i])
                        {
                            catSet.Insert(n);
                        }
                        InstanceBinding bind = app.Create.NewInstanceBinding(catSet);

                    Array group = Enum.GetValues(typeof(BuiltInParameterGroup));

                    var allSomeEnumValues = (BuiltInParameterGroup[])Enum.GetValues(typeof(BuiltInParameterGroup));
                    BuiltInParameterGroup hgs = allSomeEnumValues[97];


                    bindMap.Insert(spFileExternalDef[i],bind);
                    }
                
                trans.Commit();

                TaskDialog.Show("Info", "Dodanie parametrów współdzielonych zostało zakończone");
                return Result.Succeeded;
            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // If user decided to cancel the operation return Result.Canceled
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                // If something went wrong return Result.Failed
                message = ex.Message;
                return Result.Failed;
            }
        }     
    }

}


