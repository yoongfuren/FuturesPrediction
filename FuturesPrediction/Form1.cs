using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;


namespace FuturesPrediction
{
    public partial class Form1 : Form
    {

        private string strConn = "server=.\\SQLExpress;database=evadb;User ID=sa;Password=1234;Trusted_Connection=True;";

        public Form1()
        {
            InitializeComponent();
            
        }

       

        private void Timer1_Tick(object Sender, EventArgs e)
        {
           
            label1.Text = DateTime.Now.ToString();

            if (!CheckInternet()) {
                StopTimer();
                MessageBox.Show("無網路了~");
            }
            GetExcelData();
            //InsertToDB(new List<string>());
        }

        private void GetExcelData()
        {

            /*步驟1：設定Excel的屬性、路徑*/

            //設定讀取的Excel屬性
            string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" +

            //路徑(檔案讀取路徑)
            "Data Source=C:\\test\\test.xlsx;" +

            //選擇Excel版本
            //Excel 12.0 針對Excel 2010、2007版本(OLEDB.12.0)
            //Excel 8.0 針對Excel 97-2003版本(OLEDB.4.0)
            //Excel 5.0 針對Excel 97(OLEDB.4.0)
            "Extended Properties='Excel 12.0;" +

            //開頭是否為資料
            //若指定值為 Yes，代表 Excel 檔中的工作表第一列是欄位名稱，oleDB直接從第二列讀取
            //若指定值為 No，代表 Excel 檔中的工作表第一列就是資料了，沒有欄位名稱，oleDB直接從第一列讀取
            "HDR=NO;" +

            //IMEX=0 為「匯出模式」，能對檔案進行寫入的動作。
            //IMEX=1 為「匯入模式」，能對檔案進行讀取的動作。
            //IMEX=2 為「連結模式」，能對檔案進行讀取與寫入的動作。
            "IMEX=2'";



            /*步驟2：依照Excel的屬性及路徑開啟檔案*/

            //Excel路徑及相關資訊匯入
            OleDbConnection GetXLS = new OleDbConnection(strCon);

            //打開檔案
            GetXLS.Open();



            /*步驟3：搜尋此Excel的所有工作表，找到特定工作表進行讀檔，並將其資料存入List*/

            //搜尋xls的工作表(工作表名稱需要加$字串)
            DataTable Table = GetXLS.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            //查詢此Excel所有的工作表名稱
            string SelectSheetName = "";
            foreach (DataRow row in Table.Rows)
            {
                //抓取Xls各個Sheet的名稱(+'$')-有的名稱需要加名稱''，有的不用
                SelectSheetName = (string)row["TABLE_NAME"];

                //工作表名稱有特殊字元、空格，需加'工作表名稱$'，ex：'Sheet_A$'
                //工作表名稱沒有特殊字元、空格，需加工作表名稱$，ex：SheetA$
                //所有工作表名稱為Sheet1，讀取此工作表的內容
                if (SelectSheetName == "test$")
                {
                    //select 工作表名稱
                    OleDbCommand cmSheetA = new OleDbCommand(" SELECT * FROM [test$] ", GetXLS);
                    OleDbDataReader drSheetA = cmSheetA.ExecuteReader();

                    //讀取工作表SheetA資料
                    List<List<String>> ListSheetA = new List<List<String>>();
                    int cnt = 0;
                    while (drSheetA.Read())
                    {
                        List<String> list = new List<String>();
                        //工作表SheetA的資料存入List
                        list.Add(drSheetA[0].ToString());
                        list.Add(drSheetA[1].ToString());

                        ListSheetA.Add(list);
                        //cnt++;
                    }

                    /*步驟4：關閉檔案*/

                    //結束關閉讀檔(必要，不關會有error)
                    drSheetA.Close();
                    GetXLS.Close();

                    InsertToDB(ListSheetA);
                    
                }

            }

        }

        private void InsertToDB(List<List<String>> ListSheetA)
        {

            String strSQL = @"";

            foreach (List <String> List in ListSheetA) {
                strSQL += @"insert into table1(value1,value2) values('"+ List[0] +"','"+ List[1] +"');";
            }

            //建立連接
            SqlConnection myConn = new SqlConnection(strConn);


            //打開連接
            myConn.Open();


            //建立SQL命令對象
            SqlCommand myCommand = new SqlCommand(strSQL, myConn);


            //得到Data結果集
            SqlDataReader myDataReader = myCommand.ExecuteReader();

        }

        /**檢查網路連線*/
        private Boolean CheckInternet()
        {

            System.Net.NetworkInformation.Ping ping = new System.Net.NetworkInformation.Ping();

            if (ping.Send("www.google.com").Status == System.Net.NetworkInformation.IPStatus.Success)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void StopTimer() {
            button1.Text = "Start";
            timer1.Enabled = false;
            timer1.Stop();
        }

        private void StartTimer() {
            button1.Text = "Stop";
            timer1.Interval = 3000;
            timer1.Tick += new EventHandler(Timer1_Tick);
            timer1.Enabled = true;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (button1.Text == "Stop")
            {
                StopTimer();
            }
            else
            {
                StartTimer();
            }
        }
    }

    
}
