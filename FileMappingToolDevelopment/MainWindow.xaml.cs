using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace FileMappingToolDevelopment
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<string> folderPath = new List<string>();
        List<MapTargetOrderFileNameList> TargetFileOrderNameInfoList = new List<MapTargetOrderFileNameList>();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();          
            openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;
            

            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                textBox_CSVFileName.Text = openFileDialog.FileName;
            }
           
        }

        private void btn_Mapping_Click(object sender, RoutedEventArgs e)
        {
            folderPath = new List<string>(Directory.GetFileSystemEntries(textBox_ClusterFolder.Text, "*", SearchOption.AllDirectories));
            string filePath = textBox_CSVFileName.Text.ToString();
            string targetFilePath = textBox_TargetLog.Text.ToString();
            TargetFileOrderNameInfoList = ReadTargetLogFile(targetFilePath);
            ReadExcelFile(filePath, TargetFileOrderNameInfoList);
        }

        private List<MapTargetOrderFileNameList> ReadTargetLogFile(string logFilePath)
        {
            string readLine;

            List<MapTargetOrderFileNameList> listTargetData = new List<MapTargetOrderFileNameList>();
            int countRow = File.ReadAllLines(logFilePath).Length;
            try
            {
                using (var reader = new System.IO.StreamReader(logFilePath))
                {

                     reader.ReadLine();   //skip first header line
                  //  int inputLinesIndex = 0;

                    while ((readLine = reader.ReadLine()) != null)
                    {
                        string[] getEachLine = readLine.Split(',');
                        string getTargetOrderName = getEachLine[0].ToString();
                        string getTargetFileName = getEachLine[getEachLine.Length-1].ToString();

                        string getNameOnly = Path.GetFileNameWithoutExtension(getTargetFileName);

                        listTargetData.Add(new MapTargetOrderFileNameList() { OrderNumber = getTargetOrderName, FileName = getTargetFileName });
                       
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Occured");
            }

            return listTargetData;
      
        }


        private void ReadExcelFile(string excelFilePath, List<MapTargetOrderFileNameList> targetLogInfoList)
        {        
            string getNameOnly = "";
            //  List<MapOrderFileNameList> FileOrderNameInfoList = new List<MapOrderFileNameList>();
            //int countRow = File.ReadAllLines(excelFilePath).Length;

            string fileName = Path.GetFileNameWithoutExtension(excelFilePath);
            string debugPath = System.AppDomain.CurrentDomain.BaseDirectory;

            string outputFileName = debugPath+ @"Output\"+ fileName + "_new.csv";  //@"C:\Users\Prarthana_Bataju\Desktop\DataClusteringDevelopment\FileMappingToolDevelopment\data\aa.csv"; // Use a sensible file name.
          
            try
            {

                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + excelFilePath + "'; Extended Properties= 'Excel 12.0 XML;HDR=No;IMEX=1'";
                string queryString = "SELECT * FROM [Sheet1$]";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(queryString, connection);

                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    int inputLinesIndex = 0;

                    while (reader.Read())
                    {
                        string getOrderName = reader[14].ToString();
                        string BBColumn = reader[53].ToString();

                        //if (BBColumn != "" && BBColumn.Contains("699-0504"))
                        //{
                        //    string aa = BBColumn.Replace(" ", "");
                            

                        //}

                        MapTargetOrderFileNameList targetOrderNumber = targetLogInfoList.Find(x => x.OrderNumber.Equals(getOrderName));
                        if (targetOrderNumber != null)
                        {
                            getNameOnly = targetOrderNumber.FileName;
                        }

                        int countColumns = reader.FieldCount;
                        //string[] colValues = new string[countColumns+3];
                        StringBuilder sb = new StringBuilder();

                        for (int i=0; i<countColumns;i++)
                        {
                            sb.Append(reader[i].ToString().Replace("\n", String.Empty).Replace(',', ' '));
                          //  sb.Append(reader[i].ToString().Replace(',', ' ')); 

                            sb.Append(",");                            
                        }

                        if (inputLinesIndex == 0)
                        {
                            sb.Append("CADFileName" + "," + "FolderName" + "," + "FileName");
                        }
                        else
                        {
                            if (getNameOnly != "")
                            {
                                List<MapFolderFileNameList> clusterMapResult = SearchInClusterFolder(getNameOnly);

                                if (clusterMapResult.Count != 0)
                                {
                                    sb.Append(getNameOnly + "," + clusterMapResult[0].FolderName.ToString() + "," + clusterMapResult[0].FileName.ToString());
                                    //readLine = readLine + "," + getNameOnly + "," + clusterMapResult[0].FolderName.ToString() + "," + clusterMapResult[0].FileName.ToString();
                                }
                            }
                        }

                        if (sb != null)
                        {
                            using (StreamWriter writer = new StreamWriter(new FileStream(outputFileName, FileMode.Append, FileAccess.Write), Encoding.UTF8))
                            {
                                writer.WriteLine(sb.ToString());
                            }
                                                      
                        }

                        inputLinesIndex++;
                    }

                    reader.Close();
                }


                //using (var reader = new System.IO.StreamReader(excelFilePath))
                //{

                //    reader.ReadLine();   //skip first header line
                //    int inputLinesIndex = 0;

                //    while ((readLine = reader.ReadLine()) != null)
                //    {

                //        string[] getEachLine = readLine.Split(',');
                //        string getOrderName = getEachLine[14].ToString(); //getEachLine[getEachLine.Length -1].ToString();
                //                                                          // string getFileName = getEachLine[10].ToString();

                //        //string testname = "MVNCGT-0084-JLP174KPYKC000002_orig";

                //        MapTargetOrderFileNameList targetOderNumber = targetLogInfoList.Find(x => x.OrderNumber.Equals(getOrderName));
                //        string getNameOnly = targetOderNumber.FileName;

                //        if (inputLinesIndex == 0)
                //        {
                //            string addHeader = "CADFileName" + "," + "FolderName" + "," + "FileName";
                //            readLine = readLine + "," + addHeader;
                //        }
                //        else
                //        {
                //            List<MapFolderFileNameList> clusterMapResult = SearchInClusterFolder(getNameOnly);

                //            if (clusterMapResult.Count != 0)
                //                readLine = readLine + "," + getNameOnly + "," + clusterMapResult[0].FolderName.ToString() + "," + clusterMapResult[0].FileName.ToString();
                //        }

                //        string[] outputFileLines = File.ReadAllLines(outputFileName);
                //        if (inputLinesIndex < outputFileLines.Length)
                //        {
                //            outputFileLines[inputLinesIndex] = readLine;
                //            File.WriteAllLines(outputFileName, outputFileLines);
                //        }
                //        inputLinesIndex++;
                //    }

                //}
            }
            catch (Exception ex) {
                MessageBox.Show("Error Occured");
            }

            
            MessageBox.Show("Completed");
        }
   
        private List<MapFolderFileNameList> SearchInClusterFolder(string comparefilename)
        {
            string name = Path.GetFileNameWithoutExtension(comparefilename);
            List<MapFolderFileNameList> FolderFileNameInfoList = new List<MapFolderFileNameList>();

            foreach (string fileVal in folderPath.FindAll(element => element.Contains(name)))
            {
                string clusterVal = fileVal.Replace(textBox_ClusterFolder.Text,"");
                string[] splitVal = clusterVal.Split('\\');
                
                FolderFileNameInfoList.Add(new MapFolderFileNameList() { FolderName = splitVal[1], FileName = splitVal[2] });
            }
           
            //try
            //{
            //    string imagePath = textBox_ClusterFolder.Text.ToString(); //@"C:\Users\Prarthana_Bataju\Desktop\DataClusteringDevelopment\FileMappingToolDevelopment\data\clustereddata_90303\K350";               
            //    bool check = true;
                
            //    foreach (string folderName in folderPath)
            //    {
            //        if (check == false)
            //            break;

            //        string[] filelist = Directory.GetFiles(folderName);
            //        foreach (string imageName in filelist)
            //        {
            //            string imageResult = Path.GetFileNameWithoutExtension(imageName);
            //            if (imageResult == comparefilename)
            //            {
            //                string getFolderName = Path.GetFileNameWithoutExtension(folderName);
            //                FolderFileNameInfoList.Add(new MapFolderFileNameList() { FolderName = getFolderName, FileName = imageResult });
            //                check = false;
            //                break;
            //            }
            //        }
            //    }

                
            //}
            //catch { }

            return FolderFileNameInfoList;
        }
     

        private void btn_BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                    textBox_ClusterFolder.Text = dialog.SelectedPath.ToString();              
            }
        }

     
        private void btn_TargetLog_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
           // openFileDialog.Filter = "log files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                textBox_TargetLog.Text = openFileDialog.FileName;
            }
        }
    }

    public class MapTargetOrderFileNameList
    {
        public string OrderNumber { get; set; }
        public string FileName { get; set; }

    }

    public class MapFolderFileNameList
    {
        public string FolderName { get; set; }
        public string FileName { get; set; }
    }
}
