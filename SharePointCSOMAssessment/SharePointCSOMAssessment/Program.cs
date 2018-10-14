using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Worksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;
using Sheets = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using System.Runtime.InteropServices;
using Field = Microsoft.SharePoint.Client.Field;

namespace SharePointCSOMAssessment
{
    class Program
    {
        public static System.Data.DataTable dataTable;

        static void Main(string[] args)
        {
            string userName = "dharanendra.sheetal@acuvate.com";
            Console.WriteLine("Enter your password.");
            SecureString sec = new SecureString();

            SecureString password = GetPassword();

            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/SharePointDemo"))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(userName, password);

                try
                {
                    printExcelFileDetails(clientContext);

                    UploadFilesAndData(clientContext, dataTable);

                    UploadFileToSP(clientContext);

                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught : " + e.Message);
                }
            }
            Console.ReadKey();
        }

        //Get Excel file Details
        public static void printExcelFileDetails(ClientContext clientContext)
        {
            List emplist = clientContext.Web.Lists.GetByTitle("UserDocuments");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><RowLimit></RowLimit></View>";

            ListItemCollection empcoll = emplist.GetItems(camlQuery);
            clientContext.Load(empcoll);
            clientContext.ExecuteQuery();

            Microsoft.SharePoint.Client.File excelFile = empcoll[0].File;
            clientContext.Load(excelFile);
            clientContext.ExecuteQuery();
            var filepath = empcoll[0].File.ServerRelativeUrl;
            var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, filepath);

            //Console.WriteLine("file path :" + filepath);
            //Console.WriteLine("file info :" + fileInfo);
            //Console.WriteLine("file name :" + excelFile.Name);

            /***MyMyMy*****/
            Microsoft.SharePoint.Client.File createfileinvs = empcoll[0].File;
            if (createfileinvs != null)
            {
                try
                {
                    FileInformation fileInfor = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, createfileinvs.ServerRelativeUrl);

                    var fileName = Path.Combine(@"D:\", (string)empcoll[0].File.Name);

                    if (System.IO.File.Exists(fileName))
                    {
                        System.IO.File.Delete(fileName);
                    }

                    using (var fileStream = System.IO.File.Create(fileName))
                    {
                        fileInfo.Stream.CopyTo(fileStream);
                        fileInfo.Stream.Close();
                        fileStream.Dispose();
                    }
                }
                catch (Exception exc)
                {
                    Console.WriteLine("Exception exc : " + exc.Message);
                }

            }

            bool isError = true;
            string strErrorMsg = string.Empty;
            try
            {
                dataTable = new System.Data.DataTable("EmployeeExcelDataTable");

                ClientResult<System.IO.Stream> data = excelFile.OpenBinaryStream();
                clientContext.Load(excelFile);
                clientContext.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(clientContext, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
                                        Console.WriteLine("Cell data :" + GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i)));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);
                        }
                    }
                }
                isError = false;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception exx " + e);
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
        }

        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
                throw;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
            return value;
        }

        /***********************Uploading Data and Files in Specific List Library*******************/
        public static void UploadFilesAndData(ClientContext clientContext, System.Data.DataTable dataTable)
        {
            Application app1 = new Application();
            if (System.IO.File.Exists(@"D:\FileUploadData.xlsx"))
            {
                Console.WriteLine("Exists file");
            }

            Workbook work1 = (Microsoft.Office.Interop.Excel.Workbook)(app1.Workbooks._Open(@"D:\FileUploadData.xlsx", System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value));

            int numberOfWorkbooks = app1.Workbooks.Count;
            Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)work1.Worksheets[1];
            int numberOfSheets = work1.Worksheets.Count;

            try
            {

                if (dataTable.Rows.Count > 0)
                {
                    Console.WriteLine("-------------------Uploading file--------------------");

                    List l = clientContext.Web.Lists.GetByTitle("FileUpload");
                    clientContext.Load(l);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("List name " + l.Title + " desc :" + l.Description);

                    for (int count = 0; count < dataTable.Rows.Count; count++)
                    {
                        try
                        {
                            if (count > -1)
                            {
                                string filepath = dataTable.Rows[count]["FilePath"].ToString();
                                string status = dataTable.Rows[count]["Status"].ToString();
                                string createdBy = dataTable.Rows[count]["Created By"].ToString();
                                string department = dataTable.Rows[count]["Dept"].ToString();
                                string uploadStatus = dataTable.Rows[count]["Upload Status"].ToString();
                                string reason = dataTable.Rows[count]["Reason"].ToString();
                                long sizeoffile = new System.IO.FileInfo(filepath.Replace(@"\\", @"\")).Length;
                                var fs = new FileStream("@" + filepath, FileMode.Open);

                                if (sizeoffile > 100 && sizeoffile < 2097150)
                                {
                                    FileCreationInformation file = new FileCreationInformation();
                                    file.Content = System.IO.File.ReadAllBytes(filepath.Replace(@"\\", @"\"));
                                    file.Overwrite = true;
                                    file.Url = Path.Combine("FileUpload/", Path.GetFileName(filepath.Replace(@"\\", @"\")));
                                    Microsoft.SharePoint.Client.File uploadfile = l.RootFolder.Files.Add(file);

                                    clientContext.Load(uploadfile);
                                    clientContext.ExecuteQuery();

                                    ListItem li = uploadfile.ListItemAllFields;

                                    Field choice = l.Fields.GetByInternalNameOrTitle("Status");

                                    FieldChoice fldChoice = clientContext.CastTo<FieldChoice>(choice);

                                    string[] statusArray = status.Split(',');
                                    string statusPut = string.Empty;
                                    for (int statusCount = 0; statusCount < statusArray.Length; statusCount++)
                                    {
                                        if (fldChoice.Choices.Contains(statusArray[statusCount]))
                                        {
                                            if (statusCount == statusArray.Length - 1)
                                            {
                                                statusPut += statusArray[statusCount];
                                            }
                                            else
                                            {
                                                statusPut += statusArray[statusCount] + ";";
                                            }
                                        }
                                    }

                                    li["CreatedBy"] = createdBy;
                                    li["SizeOfFile"] = sizeoffile;
                                    li["FileType"] = Path.GetExtension(filepath.Replace(@"\\", @"\"));
                                    li["Status"] = statusPut;
                                    li["Dept"] = "2";

                                    li.Update();
                                    clientContext.ExecuteQuery();
                                    sheet1.Cells[count + 2, 5] = "Success";
                                    sheet1.Cells[count + 2, 6] = "N/A";
                                }
                                else
                                {
                                    Console.WriteLine("File : " + Path.GetFileName(filepath.Replace(@"\\", @"\")) + " could not be uploaded since file size is not in range");
                                    sheet1.Cells[count + 2, 5] = "Error";
                                    sheet1.Cells[count + 2, 6] = "File Size Exceeds Specified Range";
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Exception : "+ex);
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                Console.WriteLine("Exception :" + ee.Message);
            }

            work1.Save();

            work1.Close(true, @"D:\FileUploadData.xlsx", System.Reflection.Missing.Value);
            app1.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app1);

            Marshal.ReleaseComObject(sheet1);
            Marshal.ReleaseComObject(work1);
            Marshal.ReleaseComObject(app1);
        }

        /******************Upload Status File******************/
        public static void UploadFileToSP(ClientContext clientContext)
        {
            try
            {
                Console.WriteLine("---------------Uploading file to Share Point--------------");
                var newfile = @"D:\FileUploadData.xlsx";

                FileCreationInformation file = new FileCreationInformation();
                file.Content = System.IO.File.ReadAllBytes(newfile);
                file.Overwrite = true;
                file.Url = Path.Combine("UserDocuments/", Path.GetFileName(newfile));

                List l = clientContext.Web.Lists.GetByTitle("DemoLibrary");
                var f = l.RootFolder.Files.Add(file);

                clientContext.ExecuteQuery();
            }
            catch (Exception)
            {
                throw;
            }
        }

        //Password Secure String
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString  
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}
