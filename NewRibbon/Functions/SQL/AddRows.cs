using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Data.SqlClient;
using QC = Microsoft.Data.SqlClient;

namespace AddRows
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AddRowsSQL:IExternalCommand
    {


        public static void GetUserDetails(List<string> lista, Element e)
        {
            String par = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
            if (par == null)
            {
                lista.Add("Brak danych w Revit");
            }
            else
            {
                lista.Add(par);
            }
        }


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            Document doc = uiapp.ActiveUIDocument.Document;

            //create filter and cxollect elements
            ElementFilter catFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
            
            var elems = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .WherePasses(catFilter)
                .ToElements();

            string textToSearch = "AECTEST";
            Parameter parameter = elems[0].get_Parameter(new Guid("8160aed7-fe71-4da4-aea4-ad283b76a76a"));

            ParameterValueProvider pvp
              = new ParameterValueProvider(parameter.Id);

            FilterStringRuleEvaluator fnrvStr
                = new FilterStringEquals();

            FilterRule fRule
              = new FilterStringRule(pvp, fnrvStr, textToSearch, false);

            ElementParameterFilter filter
              = new ElementParameterFilter(fRule);

            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            ElementParameterFilter equalFilter
              = new ElementParameterFilter(fRule);

            IList<Element> filterByParam
              = collector.WherePasses(equalFilter)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .ToElements();






            //define lists to store parameters and feed them
            List<int> listaID = new List<int>();

            List<string> listaKomentarz = new List<string>();
            List<string> listaZnak = new List<string>();
            List<string> listaGuid = new List<string>();

            
            foreach (var e in filterByParam)
            {
                int parID = e.get_Parameter(BuiltInParameter.ID_PARAM).AsInteger();
                listaID.Add(parID);


                String parKomentarz = e.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
                if (parKomentarz == null)
                {
                    listaKomentarz.Add("Brak danych w Revit");
                }
                else
                {
                    listaKomentarz.Add(parKomentarz);
                }
                //BuiltInParameter.ALL_MODEL_MARK

                String parZnak = e.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString();
                if (parZnak==null)
                {
                    listaZnak.Add("Brak danych w Revit");
                }
                else
                {
                    listaZnak.Add(parZnak);
                }
                
                String parGuid = e.get_Parameter(new Guid("61063061-01d9-4ecd-9876-f7ad906ef142")).AsString();
                if (parGuid == null)
                {
                    listaGuid.Add("Brak danych w Revit");
                }
                else
                {
                    listaGuid.Add(parGuid);
                }
            }


            //define the connection string of azure database.

            var cnString = 
                "Server=tcp:revit-to-sql-svr-name.database.windows.net,1433;" +
                "Initial Catalog=RevitToSQL_DBName_01;" +
                "Persist Security Info=False;" +
                "User ID=kopeclogin;" +
                "Password=1qazXSW@;" +
                "MultipleActiveResultSets=False;" +
                "Encrypt=True;" +
                "TrustServerCertificate=False;" +
                "Connection Timeout=30;";

            //define the insert sql command, here I insert data into the GenericModel table in azure db.
            //@@@@@@@@@@@@@@@@@@
                    // define datatable (dt) it could be helpfull in revittoexcel project
                  var dt = new DataTable();
                  dt.Columns.Add("ElementsId");
                  dt.Columns.Add("Komentarz");
                  dt.Columns.Add("Znak");
                  dt.Columns.Add("ElementGUID");

                int n = 0;
    
                for (int i = 0; i < filterByParam.Count; i++)
                {
                    dt.Rows.Add(n, listaKomentarz[i], listaZnak[i], listaGuid[i]);
                    n++;
                }
/*
            using (var sqlBulk = new SqlBulkCopy(cnString))
            {
                sqlBulk.DestinationTableName = "GenericModels";
                sqlBulk.WriteToServer(dt);
            }
*/
            
                using (var connection = new SqlConnection(cnString))
                {
                    connection.Open();

                    var transaction = connection.BeginTransaction();

                        using (var sqlBulk = new SqlBulkCopy(connection, SqlBulkCopyOptions.KeepIdentity, transaction))
                        {
                            //sqlBulk.BatchSize = 5000;
                            sqlBulk.DestinationTableName = "GenericModels";
                            sqlBulk.WriteToServer(dt);
                        }
                    transaction.Commit();
                }

            
            TaskDialog.Show("Element Info", elems.Count.ToString());
            return Result.Succeeded;

        }
    }
}

