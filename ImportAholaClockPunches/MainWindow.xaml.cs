using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using DataValidationDLL;
using DateSearchDLL;
using EmployeeDateEntryDLL;
using EmployeePunchedHoursDLL;
using NewEventLogDLL;
using NewEmployeeDLL;

namespace ImportAholaClockPunches
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        EmployeeDateEntryClass TheEmployeeDateEntryClass = new EmployeeDateEntryClass();
        EmployeePunchedHoursClass TheEmployeePunchedHoursClass = new EmployeePunchedHoursClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();

        //setting up the data
        ImportAholaPunchesDataSet TheImportAholaPunchesDataSet = new ImportAholaPunchesDataSet();
        FindEmployeeByPayIDDataSet TheFindEmployeeByPayIDDataSet = new FindEmployeeByPayIDDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void expCloseProgram_Expanded(object sender, RoutedEventArgs e)
        {
            expCloseProgram.IsExpanded = false;
            TheMessagesClass.CloseTheProgram();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ResetControls();
        }

        private void Window_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ResetControls();
        }
        private void ResetControls()
        {
            TheImportAholaPunchesDataSet.importaholapunches.Rows.Clear();

            dgrPunches.ItemsSource = TheImportAholaPunchesDataSet.importaholapunches;
        }

        private void expImportExcel_Expanded(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strValueForValidation;
            int intEmployeeID = 0;
            int intPayID = 0;
            DateTime datCreatedDateTime = DateTime.Now;
            DateTime datPunchedDateTime = DateTime.Now;
            DateTime datActualDateTime = DateTime.Now;
            string strPayGroup = "";
            string strPunchMode = "";
            string strPunchType = "";
            string strPunchSouce = "";
            string strPunchIPAddress = "";
            DateTime datLastUpdate = DateTime.Now;
            bool blnFailedValidation = false;
            int intRecordsReturned;
            double douProcessDate;
            string strFirstName = "";
            string strLastName = "";

            try
            {
                expImportExcel.IsExpanded = false;
                TheImportAholaPunchesDataSet.importaholapunches.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 2; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strValueForValidation = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2).ToUpper();

                    blnFailedValidation = TheDataValidationClass.VerifyIntegerData(strValueForValidation);

                    if(blnFailedValidation == true)
                    {
                        throw new Exception();
                    }
                    else
                    {
                        intPayID = Convert.ToInt32(strValueForValidation);

                        TheFindEmployeeByPayIDDataSet = TheEmployeeClass.FindEmployeeByPayID(intPayID);

                        intRecordsReturned = TheFindEmployeeByPayIDDataSet.FindEmployeeByPayID.Rows.Count;

                        if(intRecordsReturned < 1)
                        {
                            throw new Exception();
                        }
                        else
                        {
                            intEmployeeID = TheFindEmployeeByPayIDDataSet.FindEmployeeByPayID[0].EmployeeID;
                            strFirstName = TheFindEmployeeByPayIDDataSet.FindEmployeeByPayID[0].FirstName;
                            strLastName = TheFindEmployeeByPayIDDataSet.FindEmployeeByPayID[0].LastName;
                        }
                    }

                    strValueForValidation = Convert.ToString((range.Cells[intCounter, 8] as Excel.Range).Value2).ToUpper();

                    blnFailedValidation = TheDataValidationClass.VerifyDoubleData(strValueForValidation);

                    if(blnFailedValidation == true)
                    {
                        throw new Exception();
                    }
                    else
                    {
                        douProcessDate = Convert.ToDouble(strValueForValidation);

                        datActualDateTime = DateTime.FromOADate(douProcessDate);
                    }

                    strValueForValidation = Convert.ToString((range.Cells[intCounter, 9] as Excel.Range).Value2).ToUpper();

                    blnFailedValidation = TheDataValidationClass.VerifyDoubleData(strValueForValidation);

                    if (blnFailedValidation == true)
                    {
                        throw new Exception();
                    }
                    else
                    {
                        douProcessDate = Convert.ToDouble(strValueForValidation);

                        datPunchedDateTime = DateTime.FromOADate(douProcessDate);
                    }

                    strValueForValidation = Convert.ToString((range.Cells[intCounter, 10] as Excel.Range).Value2).ToUpper();

                    blnFailedValidation = TheDataValidationClass.VerifyDateData(strValueForValidation);

                    if (blnFailedValidation == true)
                    {
                        throw new Exception();
                    }
                    else
                    {
                        datCreatedDateTime = Convert.ToDateTime(strValueForValidation);
                    }

                    strPayGroup = Convert.ToString((range.Cells[intCounter, 14] as Excel.Range).Value2).ToUpper();
                    strPunchMode = Convert.ToString((range.Cells[intCounter, 15] as Excel.Range).Value2).ToUpper();
                    strPunchType = Convert.ToString((range.Cells[intCounter, 16] as Excel.Range).Value2).ToUpper();
                    strPunchSouce = Convert.ToString((range.Cells[intCounter, 17] as Excel.Range).Value2).ToUpper();
                    strPunchIPAddress = Convert.ToString((range.Cells[intCounter, 20] as Excel.Range).Value2).ToUpper();

                    strValueForValidation = Convert.ToString((range.Cells[intCounter, 28] as Excel.Range).Value2).ToUpper();

                    blnFailedValidation = TheDataValidationClass.VerifyDateData(strValueForValidation);

                    if (blnFailedValidation == true)
                    {
                        throw new Exception();
                    }
                    else
                    {
                        datLastUpdate = Convert.ToDateTime(strValueForValidation);
                    }

                    ImportAholaPunchesDataSet.importaholapunchesRow NewPunchRow = TheImportAholaPunchesDataSet.importaholapunches.NewimportaholapunchesRow();

                    NewPunchRow.ActualDateTime = datActualDateTime;
                    NewPunchRow.CreatedDateTime = datCreatedDateTime;
                    NewPunchRow.EmployeeID = intEmployeeID;
                    NewPunchRow.FirstName = strFirstName;
                    NewPunchRow.LastName = strLastName;
                    NewPunchRow.LastUpdate = datLastUpdate;
                    NewPunchRow.PayGroup = strPayGroup;
                    NewPunchRow.PayID = intPayID;
                    NewPunchRow.PunchDateTime = datPunchedDateTime;
                    NewPunchRow.PunchIPAddress = strPunchIPAddress;
                    NewPunchRow.PunchMode = strPunchMode;
                    NewPunchRow.PunchSource = strPunchSouce;
                    NewPunchRow.PunchType = strPunchType;
                    
                    TheImportAholaPunchesDataSet.importaholapunches.Rows.Add(NewPunchRow);

                }

                dgrPunches.ItemsSource = TheImportAholaPunchesDataSet.importaholapunches;
                PleaseWait.Close();

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Ahola Clock Punches // Main Window // Import Excel  " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
