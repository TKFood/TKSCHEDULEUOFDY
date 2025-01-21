using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Xml;
using System.Xml.Linq;
using TKITDLL;
using System.Text.RegularExpressions;



namespace TKSCHEDULEUOFDY
{
    public partial class FrmSCHEDULEUOFDY : Form
    {

        int TIMEOUT_LIMITS = 240;

        public FrmSCHEDULEUOFDY()
        {
            InitializeComponent();

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
            timer1.Start();
        }

        #region FROM
        /// <summary>
        /// 每分鐘檢查1次，並每分鐘執行1次
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            //每天鐘執行1次

            try
            {

            }
            catch { }

            try
            {
                //ERP-PURTCPURTD採購單簽核
                //TKUOF.TRIGGER.PURTCPURTD.EndFormTrigger
                UPDATE_PURTC_PURTD();
            }
            catch { }

            try
            {
                //ERP-PURTEPURTF採購變更單簽核
                //TKUOF.TRIGGER.PURTEPURTF.EndFormTrigger
                UPDATE_PURTE_PURTF();
            }
            catch { }

            try
            {
                //DY採購單>轉入UOF簽核
                NEWPURTCPURTD();
            }
            catch { }

            try
            {
                //DY採購變更單>轉入UOF簽核
                NEWPURTEPURTF();
            }
            catch { }

        }


        #endregion

        #region FUNCTION
        public void NEWPURTCPURTD()
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp22"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                DataSet ds1 = new DataSet();
                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT TC001,TC002,UDF01
                                    FROM [DY].dbo.PURTC
                                    WHERE TC014='N' AND (UDF01 IN ('Y','y') )
                                    ORDER BY TC001,TC002
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    foreach (DataRow dr in ds1.Tables["ds1"].Rows)
                    {
                        ADD_PURTCPURTD_TB_WKF_EXTERNAL_TASK(dr["TC001"].ToString().Trim(), dr["TC002"].ToString().Trim());
                    }
                }
                else
                {

                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            UPDATEPURTCUDF01();
        }

        public void ADD_PURTCPURTD_TB_WKF_EXTERNAL_TASK(string TC001, string TC002)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();

            //找出ERP的單據資料
            DataTable DT = SEARCHPURTCPURTD(TC001, TC002);
            //用建單人找出建單人+部門的UOF資訊
            DataTable DTUPFDEP = SEARCHUOFDEP(DT.Rows[0]["TC011"].ToString());
          
     

            string account = DT.Rows[0]["TC011"].ToString();
            string groupId = DT.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
            string fillerName = DT.Rows[0]["MV002"].ToString();
            string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = "DY-"+DT.Rows[0]["TC001"].ToString().Trim() + DT.Rows[0]["TC002"].ToString().Trim();

            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //正式的id
            string PURTCID = SEARCHFORM_UOF_VERSION_ID("PUR40.採購單-大潁");

            if (!string.IsNullOrEmpty(PURTCID))
            {
                Form.SetAttribute("formVersionId", PURTCID);
            }


            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點 FormFieldValue
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //ID 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "ID");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC001	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC001");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC001"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC002	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC003	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC003"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC004	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC004");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC004"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC004NAME	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC004NAME");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC004NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC010	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC010");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC010"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC005	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC005");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC005"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC006	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC006");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC006"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC027	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC027");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC027"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC008	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC008");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC008"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC028	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC028");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC028"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC009	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC009");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC009"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC018	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC018");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC018"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC018NAME	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC018NAME");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC018NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //	TC011
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC011");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC011"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC011NAME	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC011NAME");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC011NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC037	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC037");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC037"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC038	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC038");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC038"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TC021	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TC021");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TC021"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //PURTD
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "PURTD");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點 DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            //DataGrid 加入至 TB 節點底下
            XmlNode PURTD = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='PURTD']");
            PURTD.AppendChild(DataGrid);


            foreach (DataRow od in DT.Rows)
            {
                // 新增 Row
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", (rowscounts).ToString());

                //Row	TD003
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD003");
                Cell.SetAttribute("fieldValue", od["TD003"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB005
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB005");
                Cell.SetAttribute("fieldValue", od["TB005"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD004
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD004");
                Cell.SetAttribute("fieldValue", od["TD004"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD005
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD005");
                Cell.SetAttribute("fieldValue", od["TD005"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD006
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD006");
                Cell.SetAttribute("fieldValue", od["TD006"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD007
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD007");
                Cell.SetAttribute("fieldValue", od["TD007"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD008
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD008");
                Cell.SetAttribute("fieldValue", od["TD008"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD009
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD009");
                Cell.SetAttribute("fieldValue", od["TD009"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD010
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD010");
                Cell.SetAttribute("fieldValue", od["TD010"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD011
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD011");
                Cell.SetAttribute("fieldValue", od["TD011"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD012
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD012");
                Cell.SetAttribute("fieldValue", od["TD012"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD015
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD015");
                Cell.SetAttribute("fieldValue", od["TD015"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD019
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD019");
                Cell.SetAttribute("fieldValue", od["TD019"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD026
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD026");
                Cell.SetAttribute("fieldValue", od["TD026"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD027
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD027");
                Cell.SetAttribute("fieldValue", od["TD027"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD028
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD028");
                Cell.SetAttribute("fieldValue", od["TD028"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TD014
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD014");
                Cell.SetAttribute("fieldValue", od["TD014"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);


                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='PURTD']/DataGrid");
                DataGridS.AppendChild(Row);

            }


            ////用ADDTACK，直接啟動起單
            //ADDTACK(Form);

            //ADD TO DB
            ////string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            //sqlConn = new SqlConnection(connectionString);

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [UOF].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{0}')
                                        ",  EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }



        }
        public void UPDATEPURTCUDF01()
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"
                                    UPDATE  [DY].dbo.PURTC  
                                    SET UDF01 = 'UOF'
                                    WHERE TC014 = 'N' AND (UDF01 IN ('Y','y') )
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public DataTable SEARCHPURTCPURTD(string TC001, string TC002)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //庫存數量看LA009 IN ('20004','20006','20008','20019','20020'

                sbSql.AppendFormat(@"  
                                   SELECT *
                                    ,USER_GUID,NAME
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    ,SUMLA011
                                    ,MA002 AS TC004NAME
                                    ,(CASE WHEN TC018='1' THEN '1.應稅內含'  WHEN TC018='2' THEN '2.應稅外加'  WHEN TC018='3' THEN '3.零稅率' WHEN TC018='4' THEN '4.免稅' WHEN TC018='9' THEN '9.不計稅' END) AS TC018NAME
                                    ,NAME AS TC011NAME
                                    FROM 
                                    (
                                        SELECT 
                                        [PURTC].[COMPANY]
                                        ,[PURTC].[CREATOR]
                                        ,[PURTC].[USR_GROUP]
                                        ,[PURTC].[CREATE_DATE]
                                        ,[PURTC].[MODIFIER]
                                        ,[PURTC].[MODI_DATE]
                                        ,[PURTC].[FLAG]
                                        ,[PURTC].[CREATE_TIME]
                                        ,[PURTC].[MODI_TIME]
                                        ,[PURTC].[TRANS_TYPE]
                                        ,[PURTC].[TRANS_NAME]
                                        ,[PURTC].[sync_date]
                                        ,[PURTC].[sync_time]
                                        ,[PURTC].[sync_mark]
                                        ,[PURTC].[sync_count]
                                        ,[PURTC].[DataUser]
                                        ,[PURTC].[DataGroup]
                                        ,[PURTC].[TC001]
                                        ,[PURTC].[TC002]
                                        ,[PURTC].[TC003]
                                        ,[PURTC].[TC004]
                                        ,[PURTC].[TC005]
                                        ,[PURTC].[TC006]
                                        ,[PURTC].[TC007]
                                        ,[PURTC].[TC008]
                                        ,[PURTC].[TC009]
                                        ,[PURTC].[TC010]
                                        ,[PURTC].[TC011]
                                        ,[PURTC].[TC012]
                                        ,[PURTC].[TC013]
                                        ,[PURTC].[TC014]
                                        ,[PURTC].[TC015]
                                        ,[PURTC].[TC016]
                                        ,[PURTC].[TC017]
                                        ,[PURTC].[TC018]
                                        ,[PURTC].[TC019]
                                        ,[PURTC].[TC020]
                                        ,[PURTC].[TC021]
                                        ,[PURTC].[TC022]
                                        ,[PURTC].[TC023]
                                        ,[PURTC].[TC024]
                                        ,[PURTC].[TC025]
                                        ,[PURTC].[TC026]
                                        ,[PURTC].[TC027]
                                        ,[PURTC].[TC028]
                                        ,[PURTC].[TC029]
                                        ,[PURTC].[TC030]
                                        ,[PURTC].[TC031]
                                        ,[PURTC].[TC032]
                                        ,[PURTC].[TC033]
                                        ,[PURTC].[TC034]
                                        ,[PURTC].[TC035]
                                        ,[PURTC].[TC036]
                                        ,[PURTC].[TC037]
                                        ,[PURTC].[TC038]
                                        ,[PURTC].[TC039]
                                        ,[PURTC].[TC040]
                                        ,[PURTC].[TC041]
                                        ,[PURTC].[TC042]
                                        ,[PURTC].[TC043]
                                        ,[PURTC].[TC044]
                                        ,[PURTC].[TC045]
                                        ,[PURTC].[TC046]
                                        ,[PURTC].[TC047]
                                        ,[PURTC].[TC048]
                                        ,[PURTC].[TC049]
                                        ,[PURTC].[TC050]
                                        ,[PURTC].[TC051]
                                        ,[PURTC].[TC052]
                                        ,[PURTC].[TC053]
                                        ,[PURTC].[TC054]
                                        ,[PURTC].[TC055]
                                        ,[PURTC].[TC056]
                                        ,[PURTC].[TC057]
                                        ,[PURTC].[TC058]
                                        ,[PURTC].[TC059]
                                        ,[PURTC].[TC060]
                                        ,[PURTC].[TC061]
                                        ,[PURTC].[TC062]
                                        ,[PURTC].[TC063]
                                        ,[PURTC].[TC064]
                                        ,[PURTC].[TC065]
                                        ,[PURTC].[TC066]
                                        ,[PURTC].[TC067]
                                        ,[PURTC].[TC068]
                                        ,[PURTC].[TC069]
                                        ,[PURTC].[TC070]
                                        ,[PURTC].[TC071]
                                        ,[PURTC].[TC072]
                                        ,[PURTC].[TC073]
                                        ,[PURTC].[TC074]
                                        ,[PURTC].[TC075]
                                        ,[PURTC].[TC076]
                                        ,[PURTC].[TC077]
                                        ,[PURTC].[TC078]
                                        ,[PURTC].[TC079]
                                        ,[PURTC].[TC080]
                                        ,[PURTC].[UDF01] AS PURTCUDF01
                                        ,[PURTC].[UDF02] AS PURTCUDF02
                                        ,[PURTC].[UDF03] AS PURTCUDF03
                                        ,[PURTC].[UDF04] AS PURTCUDF04
                                        ,[PURTC].[UDF05] AS PURTCUDF05
                                        ,[PURTC].[UDF06] AS PURTCUDF06
                                        ,[PURTC].[UDF07] AS PURTCUDF07
                                        ,[PURTC].[UDF08] AS PURTCUDF08
                                        ,[PURTC].[UDF09] AS PURTCUDF09
                                        ,[PURTC].[UDF10] AS PURTCUDF10
                                        ,[PURTD].[TD001]
                                        ,[PURTD].[TD002]
                                        ,[PURTD].[TD003]
                                        ,[PURTD].[TD004]
                                        ,[PURTD].[TD005]
                                        ,[PURTD].[TD006]
                                        ,[PURTD].[TD007]
                                        ,[PURTD].[TD008]
                                        ,[PURTD].[TD009]
                                        ,[PURTD].[TD010]
                                        ,[PURTD].[TD011]
                                        ,[PURTD].[TD012]
                                        ,[PURTD].[TD013]
                                        ,[PURTD].[TD014]
                                        ,[PURTD].[TD015]
                                        ,[PURTD].[TD016]
                                        ,[PURTD].[TD017]
                                        ,[PURTD].[TD018]
                                        ,[PURTD].[TD019]
                                        ,[PURTD].[TD020]
                                        ,[PURTD].[TD021]
                                        ,[PURTD].[TD022]
                                        ,[PURTD].[TD023]
                                        ,[PURTD].[TD024]
                                        ,[PURTD].[TD025]
                                        ,[PURTD].[TD026]
                                        ,[PURTD].[TD027]
                                        ,[PURTD].[TD028]
                                        ,[PURTD].[TD029]
                                        ,[PURTD].[TD030]
                                        ,[PURTD].[TD031]
                                        ,[PURTD].[TD032]
                                        ,[PURTD].[TD033]
                                        ,[PURTD].[TD034]
                                        ,[PURTD].[TD035]
                                        ,[PURTD].[TD036]
                                        ,[PURTD].[TD037]
                                        ,[PURTD].[TD038]
                                        ,[PURTD].[TD039]
                                        ,[PURTD].[TD040]
                                        ,[PURTD].[TD041]
                                        ,[PURTD].[TD042]
                                        ,[PURTD].[TD043]
                                        ,[PURTD].[TD044]
                                        ,[PURTD].[TD045]
                                        ,[PURTD].[TD046]
                                        ,[PURTD].[TD047]
                                        ,[PURTD].[TD048]
                                        ,[PURTD].[TD049]
                                        ,[PURTD].[TD050]
                                        ,[PURTD].[TD051]
                                        ,[PURTD].[TD052]
                                        ,[PURTD].[TD053]
                                        ,[PURTD].[TD054]
                                        ,[PURTD].[TD055]
                                        ,[PURTD].[TD056]
                                        ,[PURTD].[TD057]
                                        ,[PURTD].[TD058]
                                        ,[PURTD].[TD059]
                                        ,[PURTD].[TD060]
                                        ,[PURTD].[TD061]
                                        ,[PURTD].[TD062]
                                        ,[PURTD].[TD063]
                                        ,[PURTD].[TD064]
                                        ,[PURTD].[TD065]
                                        ,[PURTD].[TD066]
                                        ,[PURTD].[TD067]
                                        ,[PURTD].[TD068]
                                        ,[PURTD].[TD069]
                                        ,[PURTD].[TD070]
                                        ,[PURTD].[TD071]
                                        ,[PURTD].[TD072]
                                        ,[PURTD].[TD073]
                                        ,[PURTD].[TD074]
                                        ,[PURTD].[TD075]
                                        ,[PURTD].[TD076]
                                        ,[PURTD].[TD077]
                                        ,[PURTD].[TD078]
                                        ,[PURTD].[TD079]
                                        ,[PURTD].[TD080]
                                        ,[PURTD].[TD081]
                                        ,[PURTD].[TD082]
                                        ,[PURTD].[TD083]
                                        ,[PURTD].[TD084]
                                        ,[PURTD].[TD085]
                                        ,[PURTD].[TD086]
                                        ,[PURTD].[TD087]
                                        ,[PURTD].[TD088]
                                        ,[PURTD].[TD089]
                                        ,[PURTD].[TD090]
                                        ,[PURTD].[TD091]
                                        ,[PURTD].[TD092]
                                        ,[PURTD].[TD093]
                                        ,[PURTD].[TD094]
                                        ,[PURTD].[TD095]
                                        ,[PURTD].[UDF01]  AS PURTDUDF01
                                        ,[PURTD].[UDF02]  AS PURTDUDF02
                                        ,[PURTD].[UDF03]  AS PURTDUDF03
                                        ,[PURTD].[UDF04]  AS PURTDUDF04
                                        ,[PURTD].[UDF05]  AS PURTDUDF05
                                        ,[PURTD].[UDF06]  AS PURTDUDF06
                                        ,[PURTD].[UDF07]  AS PURTDUDF07
                                        ,[PURTD].[UDF08]  AS PURTDUDF08
                                        ,[PURTD].[UDF09]  AS PURTDUDF09
                                        ,[PURTD].[UDF10]  AS PURTDUDF10
                                        ,[TB_EB_USER].USER_GUID,NAME
                                        ,(SELECT TOP 1 MV002 FROM [DY].dbo.CMSMV WHERE MV001=TC011) AS 'MV002'
                                        ,(SELECT TOP 1 MA002 FROM [DY].dbo.PURMA WHERE MA001=TC004) AS 'MA002'
                                        ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [DY].dbo.INVLA WITH(NOLOCK) WHERE LA001=TD004 AND LA009 IN ('20004','20006','20008','20019','20020')) AS SUMLA011
                                        ,(SELECT TOP 1 CONVERT(NVARCHAR,TB005)+',需求日:'+CONVERT(NVARCHAR,TB011)+',數量:'+CONVERT(NVARCHAR,TB009)+' '+CONVERT(NVARCHAR,TB007) FROM  [DY].dbo.PURTB WHERE TB001=[PURTD].TD026 AND TB002=[PURTD].TD027 AND TB003=[PURTD].TD028) AS TB005
                                        FROM [DY].dbo.PURTD,[DY].dbo.PURTC
                                        LEFT JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= TC011 COLLATE Chinese_Taiwan_Stroke_BIN
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND TC001='{0}' AND TC002='{1}'
                                    ) AS TEMP
                              
                                    ", TC001, TC002);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataTable SEARCHUOFDEP(string ACCOUNT)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [GROUP_NAME] AS 'DEPNAME'
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]+','+[GROUP_NAME]+',False' AS 'DEPNO'
                                    ,[TB_EB_USER].[USER_GUID]
                                    ,[ACCOUNT]
                                    ,[NAME]
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]
                                    ,[TITLE_ID]     
                                    ,[GROUP_NAME]
                                    ,[GROUP_CODE]
                                    ,[TB_EB_EMPL_DEP].ORDERS
                                    FROM [192.168.1.223].[UOF].[dbo].[TB_EB_USER],[192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP],[192.168.1.223].[UOF].[dbo].[TB_EB_GROUP]
                                    WHERE [TB_EB_USER].[USER_GUID]=[TB_EB_EMPL_DEP].[USER_GUID]
                                    AND [TB_EB_EMPL_DEP].[GROUP_ID]=[TB_EB_GROUP].[GROUP_ID]
                                    AND ISNULL([TB_EB_GROUP].[GROUP_CODE],'')<>''
                                    AND [ACCOUNT]='{0}'
                                    ORDER BY [TB_EB_EMPL_DEP].ORDERS

                                    ", ACCOUNT);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public string SEARCHFORM_UOF_VERSION_ID(string FORM_NAME)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                   SELECT TOP 1 RTRIM(LTRIM(TB_WKF_FORM_VERSION.FORM_VERSION_ID)) FORM_VERSION_ID,TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION,TB_WKF_FORM_VERSION.ISSUE_CTL
                                    ,TB_WKF_FORM.FORM_NAME
                                    FROM [UOF].dbo.TB_WKF_FORM_VERSION,[UOF].dbo.TB_WKF_FORM
                                    WHERE 1=1
                                    AND TB_WKF_FORM_VERSION.FORM_ID=TB_WKF_FORM.FORM_ID
                                    AND TB_WKF_FORM_VERSION.ISSUE_CTL=1
                                    AND FORM_NAME='{0}'
                                    ORDER BY TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION DESC

                                    ", FORM_NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["FORM_VERSION_ID"].ToString();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void NEWPURTEPURTF()
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp22"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                DataSet ds1 = new DataSet();
                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT TE001,TE002,TE003,UDF01
                                    FROM [DY].dbo.PURTE
                                    WHERE TE017='N' AND (UDF01 IN ('Y','y') )
                                    ORDER BY TE001,TE002,TE003
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    foreach (DataRow dr in ds1.Tables["ds1"].Rows)
                    {
                        ADD_PURTEPURTF_TB_WKF_EXTERNAL_TASK(dr["TE001"].ToString().Trim(), dr["TE002"].ToString().Trim(), dr["TE003"].ToString().Trim());
                    }


                    //ADDTB_WKF_EXTERNAL_TASK("A311", "20210415007");
                }
                else
                {

                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            UPDATEPURTEUDF01();
        }

        public void ADD_PURTEPURTF_TB_WKF_EXTERNAL_TASK(string TE001, string TE002, string TE003)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();


            DataTable DT = SEARCHPURTEPURTF(TE001, TE002, TE003);
            DataTable DTUPFDEP = SEARCHUOFDEP(DT.Rows[0]["TE037"].ToString());

            string account = DT.Rows[0]["TE037"].ToString();
            string groupId = DT.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
            string fillerName = DT.Rows[0]["MV002"].ToString();
            string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();

            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO = DTUPFDEP.Rows[0]["DEPNO"].ToString();

            string EXTERNAL_FORM_NBR = "DY-" + DT.Rows[0]["TE001"].ToString().Trim() + DT.Rows[0]["TE002"].ToString().Trim() + DT.Rows[0]["TE003"].ToString().Trim();

            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //正式的id
            string PURTEID = SEARCHFORM_UOF_VERSION_ID("PUR50.採購變更單-大潁");

            if (!string.IsNullOrEmpty(PURTEID))
            {
                Form.SetAttribute("formVersionId", PURTEID);
            }


            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點 FormFieldValue
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //ID 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "ID");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE001	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE001");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE001"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE002	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE003	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE003"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE004
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE004");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE004"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE006
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE006");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE006"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE005
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE005");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE005"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE005NAME
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE005NAME");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE005NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE007
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE007");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE007"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE008
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE008");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE008"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE009
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE009");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE009"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE010
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE010");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE010"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE023
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE023");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE023"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE011
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE011");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE011"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE012
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE012");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE012"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE015
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE015");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE015"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE018
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE018");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE018"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE018NAME
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE018NAME");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE018NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE019
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE019");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE019"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE020
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE020");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE020"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE022
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE022");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE022"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE024
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE024");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE024"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE027
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE027");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE027"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE037
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE037");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE037"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE037NAME
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE037NAME");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE037NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE043
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE043");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE043"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE045
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE045");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE045"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TE046
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TE046");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TE046"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);







            //建立節點FieldItem
            //PURTF
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "PURTF");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點 DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            //DataGrid 加入至 TB 節點底下
            XmlNode PURTD = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='PURTF']");
            PURTD.AppendChild(DataGrid);


            foreach (DataRow od in DT.Rows)
            {
                // 新增 Row
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", (rowscounts).ToString());

                //Row	TF004
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF004");
                Cell.SetAttribute("fieldValue", od["TF004"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF005
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF005");
                Cell.SetAttribute("fieldValue", od["TF005"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF006
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF006");
                Cell.SetAttribute("fieldValue", od["TF006"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF007
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF007");
                Cell.SetAttribute("fieldValue", od["TF007"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF008
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF008");
                Cell.SetAttribute("fieldValue", od["TF008"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF009
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF009");
                Cell.SetAttribute("fieldValue", od["TF009"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF010
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF010");
                Cell.SetAttribute("fieldValue", od["TF010"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF011
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF011");
                Cell.SetAttribute("fieldValue", od["TF011"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF012
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF012");
                Cell.SetAttribute("fieldValue", od["TF012"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF013
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF013");
                Cell.SetAttribute("fieldValue", od["TF013"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF014
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF014");
                Cell.SetAttribute("fieldValue", od["TF014"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF015
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF015");
                Cell.SetAttribute("fieldValue", od["TF015"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF017
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF017");
                Cell.SetAttribute("fieldValue", od["TF017"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF018
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF018");
                Cell.SetAttribute("fieldValue", od["TF018"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF021
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF021");
                Cell.SetAttribute("fieldValue", od["TF021"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF022
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF022");
                Cell.SetAttribute("fieldValue", od["TF022"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TF030
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TF030");
                Cell.SetAttribute("fieldValue", od["TF030"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);


                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='PURTF']/DataGrid");
                DataGridS.AppendChild(Row);

            }

            ////用ADDTACK，直接啟動起單
            //ADDTACK(Form);

            //ADD TO DB
            ////string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            //sqlConn = new SqlConnection(connectionString);

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [UOF].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{0}')
                                        ", EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }
        }
        public void UPDATEPURTEUDF01()
        {

            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            int result;

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"
                                    UPDATE  [DY].dbo.PURTE  
                                    SET UDF01 = 'UOF'
                                    WHERE TE017 = 'N' AND (UDF01 IN ('Y','y') )
                                    ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public DataTable SEARCHPURTEPURTF(string TE001, string TE002, string TE003)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //庫存數量看LA009 IN ('20004','20006','20008','20019','20020'

                sbSql.AppendFormat(@"  
                                   SELECT *
                                    ,USER_GUID,NAME
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    ,SUMLA011
                                    ,MA002 AS TE005NAME
                                    ,(CASE WHEN TE018='1' THEN '1.應稅內含'  WHEN TE018='2' THEN '2.應稅外加'  WHEN TE018='3' THEN '3.零稅率' WHEN TE018='4' THEN '4.免稅' WHEN TE018='9' THEN '9.不計稅' END) AS TE018NAME
                                    ,NAME AS TE037NAME
                                    FROM 
                                    (
                                    SELECT 
                                    [PURTE].[COMPANY]
                                    ,[PURTE].[CREATOR]
                                    ,[PURTE].[USR_GROUP]
                                    ,[PURTE].[CREATE_DATE]
                                    ,[PURTE].[MODIFIER]
                                    ,[PURTE].[MODI_DATE]
                                    ,[PURTE].[FLAG]
                                    ,[PURTE].[CREATE_TIME]
                                    ,[PURTE].[MODI_TIME]
                                    ,[PURTE].[TRANS_TYPE]
                                    ,[PURTE].[TRANS_NAME]
                                    ,[PURTE].[sync_date]
                                    ,[PURTE].[sync_time]
                                    ,[PURTE].[sync_mark]
                                    ,[PURTE].[sync_count]
                                    ,[PURTE].[DataUser]
                                    ,[PURTE].[DataGroup]
                                    ,[PURTE].[TE001]
                                    ,[PURTE].[TE002]
                                    ,[PURTE].[TE003]
                                    ,[PURTE].[TE004]
                                    ,[PURTE].[TE005]
                                    ,[PURTE].[TE006]
                                    ,[PURTE].[TE007]
                                    ,[PURTE].[TE008]
                                    ,[PURTE].[TE009]
                                    ,[PURTE].[TE010]
                                    ,[PURTE].[TE011]
                                    ,[PURTE].[TE012]
                                    ,[PURTE].[TE013]
                                    ,[PURTE].[TE014]
                                    ,[PURTE].[TE015]
                                    ,[PURTE].[TE016]
                                    ,[PURTE].[TE017]
                                    ,[PURTE].[TE018]
                                    ,[PURTE].[TE019]
                                    ,[PURTE].[TE020]
                                    ,[PURTE].[TE021]
                                    ,[PURTE].[TE022]
                                    ,[PURTE].[TE023]
                                    ,[PURTE].[TE024]
                                    ,[PURTE].[TE025]
                                    ,[PURTE].[TE026]
                                    ,[PURTE].[TE027]
                                    ,[PURTE].[TE028]
                                    ,[PURTE].[TE029]
                                    ,[PURTE].[TE030]
                                    ,[PURTE].[TE031]
                                    ,[PURTE].[TE032]
                                    ,[PURTE].[TE033]
                                    ,[PURTE].[TE034]
                                    ,[PURTE].[TE035]
                                    ,[PURTE].[TE036]
                                    ,[PURTE].[TE037]
                                    ,[PURTE].[TE038]
                                    ,[PURTE].[TE039]
                                    ,[PURTE].[TE040]
                                    ,[PURTE].[TE041]
                                    ,[PURTE].[TE042]
                                    ,[PURTE].[TE043]
                                    ,[PURTE].[TE045]
                                    ,[PURTE].[TE046]
                                    ,[PURTE].[TE047]
                                    ,[PURTE].[TE048]
                                    ,[PURTE].[TE103]
                                    ,[PURTE].[TE107]
                                    ,[PURTE].[TE108]
                                    ,[PURTE].[TE109]
                                    ,[PURTE].[TE110]
                                    ,[PURTE].[TE113]
                                    ,[PURTE].[TE114]
                                    ,[PURTE].[TE115]
                                    ,[PURTE].[TE118]
                                    ,[PURTE].[TE119]
                                    ,[PURTE].[TE120]
                                    ,[PURTE].[TE121]
                                    ,[PURTE].[TE122]
                                    ,[PURTE].[TE123]
                                    ,[PURTE].[TE124]
                                    ,[PURTE].[TE125]
                                    ,[PURTE].[TE134]
                                    ,[PURTE].[TE135]
                                    ,[PURTE].[TE136]
                                    ,[PURTE].[TE137]
                                    ,[PURTE].[TE138]
                                    ,[PURTE].[TE139]
                                    ,[PURTE].[TE140]
                                    ,[PURTE].[TE141]
                                    ,[PURTE].[TE142]
                                    ,[PURTE].[TE143]
                                    ,[PURTE].[TE144]
                                    ,[PURTE].[TE145]
                                    ,[PURTE].[TE146]
                                    ,[PURTE].[TE147]
                                    ,[PURTE].[TE148]
                                    ,[PURTE].[TE149]
                                    ,[PURTE].[TE150]
                                    ,[PURTE].[TE151]
                                    ,[PURTE].[TE152]
                                    ,[PURTE].[TE153]
                                    ,[PURTE].[TE154]
                                    ,[PURTE].[TE155]
                                    ,[PURTE].[TE156]
                                    ,[PURTE].[TE157]
                                    ,[PURTE].[TE158]
                                    ,[PURTE].[TE159]
                                    ,[PURTE].[TE160]
                                    ,[PURTE].[TE161]
                                    ,[PURTE].[TE162]
                                    ,[PURTE].[UDF01]  AS 'PURTFUDE01'
                                    ,[PURTE].[UDF02]  AS 'PURTFUDE02'
                                    ,[PURTE].[UDF03]  AS 'PURTFUDE03'
                                    ,[PURTE].[UDF04]  AS 'PURTFUDE04'
                                    ,[PURTE].[UDF05]  AS 'PURTFUDE05'
                                    ,[PURTE].[UDF06]  AS 'PURTFUDE06'
                                    ,[PURTE].[UDF07]  AS 'PURTFUDE07'
                                    ,[PURTE].[UDF08]  AS 'PURTFUDE08'
                                    ,[PURTE].[UDF09]  AS 'PURTFUDE09'
                                    ,[PURTE].[UDF10]  AS 'PURTFUDE10'
                                    ,[PURTF].[TF001]
                                    ,[PURTF].[TF002]
                                    ,[PURTF].[TF003]
                                    ,[PURTF].[TF004]
                                    ,[PURTF].[TF005]
                                    ,[PURTF].[TF006]
                                    ,[PURTF].[TF007]
                                    ,[PURTF].[TF008]
                                    ,[PURTF].[TF009]
                                    ,[PURTF].[TF010]
                                    ,[PURTF].[TF011]
                                    ,[PURTF].[TF012]
                                    ,[PURTF].[TF013]
                                    ,[PURTF].[TF014]
                                    ,[PURTF].[TF015]
                                    ,[PURTF].[TF016]
                                    ,[PURTF].[TF017]
                                    ,[PURTF].[TF018]
                                    ,[PURTF].[TF019]
                                    ,[PURTF].[TF020]
                                    ,[PURTF].[TF021]
                                    ,[PURTF].[TF022]
                                    ,[PURTF].[TF023]
                                    ,[PURTF].[TF024]
                                    ,[PURTF].[TF025]
                                    ,[PURTF].[TF026]
                                    ,[PURTF].[TF027]
                                    ,[PURTF].[TF028]
                                    ,[PURTF].[TF029]
                                    ,[PURTF].[TF030]
                                    ,[PURTF].[TF031]
                                    ,[PURTF].[TF032]
                                    ,[PURTF].[TF033]
                                    ,[PURTF].[TF034]
                                    ,[PURTF].[TF035]
                                    ,[PURTF].[TF036]
                                    ,[PURTF].[TF037]
                                    ,[PURTF].[TF038]
                                    ,[PURTF].[TF039]
                                    ,[PURTF].[TF040]
                                    ,[PURTF].[TF041]
                                    ,[PURTF].[TF104]
                                    ,[PURTF].[TF105]
                                    ,[PURTF].[TF106]
                                    ,[PURTF].[TF107]
                                    ,[PURTF].[TF108]
                                    ,[PURTF].[TF109]
                                    ,[PURTF].[TF110]
                                    ,[PURTF].[TF111]
                                    ,[PURTF].[TF112]
                                    ,[PURTF].[TF113]
                                    ,[PURTF].[TF114]
                                    ,[PURTF].[TF118]
                                    ,[PURTF].[TF119]
                                    ,[PURTF].[TF120]
                                    ,[PURTF].[TF121]
                                    ,[PURTF].[TF122]
                                    ,[PURTF].[TF123]
                                    ,[PURTF].[TF124]
                                    ,[PURTF].[TF125]
                                    ,[PURTF].[TF126]
                                    ,[PURTF].[TF127]
                                    ,[PURTF].[TF128]
                                    ,[PURTF].[TF129]
                                    ,[PURTF].[TF130]
                                    ,[PURTF].[TF131]
                                    ,[PURTF].[TF132]
                                    ,[PURTF].[TF133]
                                    ,[PURTF].[TF134]
                                    ,[PURTF].[TF135]
                                    ,[PURTF].[TF136]
                                    ,[PURTF].[TF137]
                                    ,[PURTF].[TF138]
                                    ,[PURTF].[TF139]
                                    ,[PURTF].[TF140]
                                    ,[PURTF].[TF141]
                                    ,[PURTF].[TF142]
                                    ,[PURTF].[TF143]
                                    ,[PURTF].[TF144]
                                    ,[PURTF].[TF145]
                                    ,[PURTF].[TF146]
                                    ,[PURTF].[TF147]
                                    ,[PURTF].[TF148]
                                    ,[PURTF].[TF149]
                                    ,[PURTF].[TF150]
                                    ,[PURTF].[TF151]
                                    ,[PURTF].[TF152]
                                    ,[PURTF].[TF153]
                                    ,[PURTF].[TF154]
                                    ,[PURTF].[TF155]
                                    ,[PURTF].[TF156]
                                    ,[PURTF].[TF157]
                                    ,[PURTF].[TF158]
                                    ,[PURTF].[TF159]
                                    ,[PURTF].[TF160]
                                    ,[PURTF].[TF161]
                                    ,[PURTF].[TF162]
                                    ,[PURTF].[TF163]
                                    ,[PURTF].[TF164]
                                    ,[PURTF].[TF165]
                                    ,[PURTF].[TF166]
                                    ,[PURTF].[TF167]
                                    ,[PURTF].[TF168]
                                    ,[PURTF].[TF169]
                                    ,[PURTF].[TF170]
                                    ,[PURTF].[TF171]
                                    ,[PURTF].[TF172]
                                    ,[PURTF].[TF173]
                                    ,[PURTF].[UDF01] AS 'PURTFUDF01'
                                    ,[PURTF].[UDF02] AS 'PURTFUDF02'
                                    ,[PURTF].[UDF03] AS 'PURTFUDF03'
                                    ,[PURTF].[UDF04] AS 'PURTFUDF04'
                                    ,[PURTF].[UDF05] AS 'PURTFUDF05'
                                    ,[PURTF].[UDF06] AS 'PURTFUDF06'
                                    ,[PURTF].[UDF07] AS 'PURTFUDF07'
                                    ,[PURTF].[UDF08] AS 'PURTFUDF08'
                                    ,[PURTF].[UDF09] AS 'PURTFUDF09'
                                    ,[PURTF].[UDF10] AS 'PURTFUDF10'
                                    ,[TB_EB_USER].USER_GUID,NAME
                                    ,(SELECT TOP 1 MV002 FROM [DY].dbo.CMSMV WHERE MV001=TE037) AS 'MV002'
                                    ,(SELECT TOP 1 MA002 FROM [DY].dbo.PURMA WHERE MA001=TE005) AS 'MA002'
                                    ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [DY].dbo.INVLA WITH(NOLOCK) WHERE LA001=TF005 AND LA009 IN ('20004','20006','20008','20019','20020')) AS SUMLA011
                                    FROM [DY].dbo.PURTF,[DY].dbo.PURTE
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= TE037 COLLATE Chinese_Taiwan_Stroke_BIN
                                    WHERE TE001=TF001 AND TE002=TF002 AND TE003=TF003
                                    AND TE001='{0}' AND TE002='{1}' AND TE003='{2}'
                                    ) AS TEMP
                              
                                    ", TE001, TE002, TE003);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void UPDATE_PURTC_PURTD()
        {
            string DOC_NBR = "";
            string ACCOUNT = "";
            string MODIFIER = null;

            string FORMID;
            string TC001;
            string TC002;

            string ISCLOSE;

            DataTable DT = FIND_UOF_PURTC_PORTD();

            if (DT != null && DT.Rows.Count >= 1)
            {
                foreach (DataRow DR in DT.Rows)
                {
                    TC001 = DR["TC001"].ToString().Trim();
                    TC002 = DR["TC002"].ToString().Trim();

                    DOC_NBR = DR["DOC_NBR"].ToString().Trim();
                    ACCOUNT = DR["ACCOUNT"].ToString().Trim();
                    MODIFIER = DR["ACCOUNT"].ToString().Trim();
                    FORMID = DR["DOC_NBR"].ToString().Trim();

                    UPDATE_PURTC_PORTD_EXE(TC001, TC002, FORMID, MODIFIER);
                }
            }
        }

        public DataTable FIND_UOF_PURTC_PORTD()
        {

            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    WITH TEMP AS (
                                    SELECT 
                                        [FORM_NAME],
                                        [DOC_NBR],
	                                    [CURRENT_DOC].value('(/Form/FormFieldValue/FieldItem[@fieldId=""TC001""]/@fieldValue)[1]', 'NVARCHAR(100)') AS TC001,
                                        [CURRENT_DOC].value('(/Form/FormFieldValue/FieldItem[@fieldId=""TC002""]/@fieldValue)[1]', 'NVARCHAR(100)') AS TC002,
                                        TASK_ID,
                                        TASK_STATUS,
                                        TASK_RESULT
                                        FROM [UOF].[dbo].TB_WKF_TASK
                                        LEFT JOIN [UOF].[dbo].[TB_WKF_FORM_VERSION] ON[TB_WKF_FORM_VERSION].FORM_VERSION_ID = TB_WKF_TASK.FORM_VERSION_ID
                                        LEFT JOIN [UOF].[dbo].[TB_WKF_FORM] ON[TB_WKF_FORM].FORM_ID = [TB_WKF_FORM_VERSION].FORM_ID
                                        WHERE[FORM_NAME] = 'PUR40.採購單-大潁'
                                        AND TASK_STATUS = '2'
                                        AND TASK_RESULT = '0'

                                    )
                                    SELECT TEMP.*,
                                    (
                                        SELECT TOP 1 [TB_EB_USER].ACCOUNT
                                        FROM [UOF].[dbo].TB_WKF_TASK_NODE
                                        LEFT JOIN [UOF].[dbo].[TB_EB_USER]
                                            ON[TB_EB_USER].USER_GUID = [TB_WKF_TASK_NODE].ACTUAL_SIGNER

                                        WHERE 1=1
                                        AND ISNULL([TB_WKF_TASK_NODE].ACTUAL_SIGNER,'')<>''
	                                    AND [TB_WKF_TASK_NODE].TASK_ID = TEMP.TASK_ID
                                        ORDER BY FINISH_TIME DESC
                                    ) AS ACCOUNT
                                    FROM TEMP
                                    WHERE 1=1
                                    AND REPLACE(TC001+TC002,',','')  IN
                                    (
                                        SELECT REPLACE(TC001+TC002,' ' ,'')
                                        FROM [192.168.1.105].[DY].dbo.PURTC
                                        WHERE TC014 IN ('N')
                                    )                            

                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATE_PURTC_PORTD_EXE(string TC001, string TC002, string FORMID, string MODIFIER)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);


            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder queryString = new StringBuilder();

            queryString.AppendFormat(@"   
                                       UPDATE [DY].dbo.PURTC SET TC014='Y' WHERE TC001=@TC001 AND TC002=@TC002 
                                       UPDATE [DY].dbo.PURTD SET TD018='Y' WHERE TD001=@TC001 AND TD002=@TC002 

                                       UPDATE [DY].dbo.PURTC SET UDF02=@UDF02 WHERE TC001=@TC001 AND TC002=@TC002 

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(sqlConn.ConnectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;
                    //command.Parameters.Add("@TA014", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value = FORMID;


                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }
        }

        public void UPDATE_PURTE_PURTF()
        {
            string DOC_NBR = "";
            string ACCOUNT = "";
            string MODIFIER = null;

            string FORMID;
            string TE001;
            string TE002;
            string TE003;

            string ISCLOSE;

            DataTable DT = FIND_UOF_PURTE_PORTF();

            if (DT != null && DT.Rows.Count >= 1)
            {
                foreach (DataRow DR in DT.Rows)
                {
                    TE001 = DR["TE001"].ToString().Trim();
                    TE002 = DR["TE002"].ToString().Trim();
                    TE003 = DR["TE003"].ToString().Trim();

                    DOC_NBR = DR["DOC_NBR"].ToString().Trim();
                    ACCOUNT = DR["NOWACCOUNT"].ToString().Trim();
                    MODIFIER = DR["NOWACCOUNT"].ToString().Trim();
                    FORMID = DR["DOC_NBR"].ToString().Trim();

                    UPDATE_PURTE_PORTF_EXE(TE001, TE002, TE003, FORMID, MODIFIER);
                }
            }
        }

        public DataTable FIND_UOF_PURTE_PORTF()
        {

            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"                                      
                                    WITH TEMP AS (
                                    SELECT 
                                        [FORM_NAME],
                                        [DOC_NBR],
	                                    [CURRENT_DOC].value('(/Form/FormFieldValue/FieldItem[@fieldId=""TE001""]/@fieldValue)[1]', 'NVARCHAR(100)') AS TE001,
                                        [CURRENT_DOC].value('(/Form/FormFieldValue/FieldItem[@fieldId=""TE002""]/@fieldValue)[1]', 'NVARCHAR(100)') AS TE002,
                                        [CURRENT_DOC].value('(/Form/FormFieldValue/FieldItem[@fieldId=""TE003""]/@fieldValue)[1]', 'NVARCHAR(100)') AS TE003,
                                        TASK_ID,
                                        TASK_STATUS,
                                        TASK_RESULT
                                        FROM[UOF].[dbo].TB_WKF_TASK
                                        LEFT JOIN[UOF].[dbo].[TB_WKF_FORM_VERSION] ON[TB_WKF_FORM_VERSION].FORM_VERSION_ID = TB_WKF_TASK.FORM_VERSION_ID
                                        LEFT JOIN[UOF].[dbo].[TB_WKF_FORM] ON[TB_WKF_FORM].FORM_ID = [TB_WKF_FORM_VERSION].FORM_ID
                                        WHERE[FORM_NAME] = 'PUR50.採購變更單-大潁'
                                        AND TASK_STATUS = '2'
                                        AND TASK_RESULT = '0'

                                    )
                                    SELECT TEMP.*,
                                    (
                                        SELECT TOP 1[TB_EB_USER].ACCOUNT
                                        FROM[UOF].[dbo].TB_WKF_TASK_NODE
                                        LEFT JOIN[UOF].[dbo].[TB_EB_USER]
                                            ON[TB_EB_USER].USER_GUID = [TB_WKF_TASK_NODE].ACTUAL_SIGNER

                                    WHERE 1=1
                                        AND ISNULL([TB_WKF_TASK_NODE].ACTUAL_SIGNER,'')<>''
	                                    AND[TB_WKF_TASK_NODE].TASK_ID = TEMP.TASK_ID
                                       ORDER BY FINISH_TIME DESC
                                    ) AS NOWACCOUNT
                                    FROM TEMP
                                    WHERE 1=1
                                    AND REPLACE(TE001+TE002+TE003,',','')  IN
                                    (
                                        SELECT REPLACE(TE001+TE002+TE003,' ' ,'')
                                        FROM[192.168.1.105].[DY].dbo.PURTE
                                    WHERE TE017 IN('N')
                                    )                            
                    

                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATE_PURTE_PORTF_EXE(string TE001, string TE002, string TE003, string FORMID, string MODIFIER)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            string TC001 = TE001;
            string TC002 = TE002;
            string TD001 = TE001;
            string TD002 = TE002;
            string TF001 = TE001;
            string TF002 = TE002;
            string TF003 = TE003;

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);


            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder queryString = new StringBuilder();

            queryString.AppendFormat(@"   
                                       --INSERT PURTD

                                        INSERT INTO [DY].dbo.PURTD
                                        (
                                        COMPANY,CREATOR,USR_GROUP,CREATE_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,DataGroup
                                        ,TD001
                                        ,TD002
                                        ,TD003
                                        ,TD004
                                        ,TD005
                                        ,TD006
                                        ,TD007
                                        ,TD008
                                        ,TD009
                                        ,TD010
                                        ,TD011
                                        ,TD012
                                        ,TD014
                                        ,TD015
                                        ,TD016
                                        ,TD017
                                        ,TD018
                                        ,TD019
                                        ,TD020
                                        ,TD022
                                        ,TD025
                                        )

                                        SELECT 
                                        COMPANY,CREATOR,USR_GROUP,CREATE_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,DataGroup
                                        ,TF001
                                        ,TF002
                                        ,TF104
                                        ,TF005
                                        ,TF006
                                        ,TF007
                                        ,TF008
                                        ,TF009
                                        ,TF010
                                        ,TF011
                                        ,TF012
                                        ,TF013
                                        ,TF030
                                        ,TF015
                                        ,'N'
                                        ,TF022
                                        ,'Y'
                                        ,TF018
                                        ,TF019
                                        ,TF020
                                        ,TF021
                                        FROM [DY].dbo.PURTF
                                        WHERE TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003
                                        AND TF001+TF002+TF104 NOT IN (SELECT TD001+TD002+TD003  FROM [DY].dbo.PURTD WHERE TD001=@TD001 AND TD002=@TD002)

                                        --UPDATE PURTD

                                        UPDATE [DY].dbo.PURTD
                                        SET 
                                        TD004=TF005
                                        ,TD005=TF006
                                        ,TD006=TF007
                                        ,TD007=TF008
                                        ,TD008=TF009
                                        ,TD009=TF010
                                        ,TD010=TF011
                                        ,TD011=TF012
                                        ,TD012=TF013
                                        ,TD014=TF030
                                        ,TD015=TF015
                                        ,TD017=TF022
                                        ,TD019=TF018
                                        ,TD020=TF019
                                        ,TD022=TF020
                                        ,TD025=TF021
                                        FROM [DY].dbo.PURTF
                                        WHERE TD001=@TD001 AND TD002=@TD002 AND TD003=TF104
                                        AND TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003


                                        --UPDATE PURTC

                                        UPDATE [DY].dbo.PURTC
                                        SET 
                                        TC004=TE005
                                        ,TC005=TE007
                                        ,TC006=TE008
                                        ,TC007=TE009
                                        ,TC008=TE010
                                        ,TC015=TE013
                                        ,TC016=TE014
                                        ,TC017=TE015
                                        ,TC018=TE018
                                        ,TC021=TE019
                                        ,TC022=TE020
                                        ,TC026=TE022
                                        ,TC027=TE023
                                        ,TC028=TE024
                                        ,TC009=TE027
                                        ,TC035=TE029
                                        ,TC011=TE037
                                        ,TC047=TE039
                                        ,TC048=TE040
                                        ,TC050=TE041
                                        ,TC036=TE043
                                        ,TC037=TE045
                                        ,TC038=TE046
                                        ,TC039=TE047
                                        ,TC040=TE048
                                        FROM [DY].dbo.PURTE
                                        WHERE TC001=@TC001 AND TC002=@TC002
                                        AND TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                        --更新PURTC的未稅、稅額、總金額、數量
                                        UPDATE [DY].dbo.PURTC
                                        SET TC019=(CASE WHEN TC018='1' THEN (SELECT ISNULL(ISNULL(ROUND(SUM(TD011)/(1+TC026),0),0),0) FROM [DY].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='2' THEN (SELECT ISNULL(SUM(TD011),0) FROM [DY].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='3' THEN (SELECT ISNULL(SUM(TD011),0) FROM [DY].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='4' THEN (SELECT ISNULL(SUM(TD011),0) FROM [DY].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='9' THEN (SELECT ISNULL(SUM(TD011),0) FROM [DY].dbo.PURTD WHERE TD001+TD002=TC001+TC002)  
                                                                            END
                                                                            END
                                                                            END 
                                                                            END
                                                                            END)
                                        ,TC020=(CASE WHEN TC018='1' THEN (SELECT (ISNULL(SUM(TD011),0)-ISNULL(ROUND(SUM(TD011)/(1+TC026),0),0)) FROM [DY].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='2' THEN (SELECT ISNULL(ROUND(SUM(TD011)*TC026,0),0) FROM [DY].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='3' THEN 0 
                                                                            ELSE CASE WHEN TC018='4' THEN 0
                                                                            ELSE CASE WHEN TC018='9' THEN 0 
                                                                            END
                                                                            END
                                                                            END 
                                                                            END
                                                                            END)
                                        ,TC023=(SELECT ISNULL(SUM(TD008),0) FROM [DY].dbo.PURTD WHERE TD001=TC001 AND TD002=TC002)
                                        WHERE TC001=@TC001 AND TC002=@TC002

                                        --如果變更單整理指定結案，原PURTC也指定結案

                                        UPDATE [DY].dbo.PURTD
                                        SET TD016='y'
                                        FROM [DY].dbo.PURTE
                                        WHERE TD001=@TD001 AND TD002=@TD002
                                        AND TE012='Y'                                    
                                        AND TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                        --如果變更單單身指定結案，原PURTD也指定結案
                                        UPDATE [DY].dbo.PURTD
                                        SET TD016='y'
                                        FROM [DY].dbo.PURTF
                                        WHERE   TD001=@TD001 AND TD002=@TD002
                                        AND TF001=TD001 AND TF002=TD002 AND TF104=TD003
                                        AND TF014='Y'                                       
                                        AND TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003

                                        --更新PURTE
                                        UPDATE [DY].dbo.PURTE
                                        SET TE017='Y'
                                        ,UDF02=@UDF02
                                        WHERE TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                        --更新PURTF
                                        UPDATE [DY].dbo.PURTF
                                        SET TF016='Y'
                                        WHERE TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003

                                        --更新PURTC
                                        UPDATE [DY].dbo.PURTC
                                        SET UDF03=@UDF03
                                        WHERE TC001=@TC001 AND TC002=@TC002
                                      

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(sqlConn.ConnectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;
                    command.Parameters.Add("@TD001", SqlDbType.NVarChar).Value = TD001;
                    command.Parameters.Add("@TD002", SqlDbType.NVarChar).Value = TD002;
                    command.Parameters.Add("@TE001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TE002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TE003", SqlDbType.NVarChar).Value = TE003;
                    command.Parameters.Add("@TF001", SqlDbType.NVarChar).Value = TF001;
                    command.Parameters.Add("@TF002", SqlDbType.NVarChar).Value = TF002;
                    command.Parameters.Add("@TF003", SqlDbType.NVarChar).Value = TF003;
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value = FORMID;
                    command.Parameters.Add("@UDF03", SqlDbType.NVarChar).Value = FORMID;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }
        }

        #endregion

        #region BUTTON
        private void button15_Click(object sender, EventArgs e)
        {
            //DY採購單>轉入UOF簽核
            NEWPURTCPURTD();

            MessageBox.Show("OK");
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //DY採購變更單>轉入UOF簽核
            NEWPURTEPURTF();

            MessageBox.Show("OK");
        }
        private void button101_Click(object sender, EventArgs e)
        {
            //ERP-PURTCPURTD採購單簽核
            //TKUOF.TRIGGER.PURTCPURTD.EndFormTrigger
            UPDATE_PURTC_PURTD();

            MessageBox.Show("OK");
        }

        private void button103_Click(object sender, EventArgs e)
        {
            //ERP-PURTEPURTF採購變更單簽核
            //TKUOF.TRIGGER.PURTEPURTF.EndFormTrigger
            UPDATE_PURTE_PURTF();

            MessageBox.Show("OK");
        }

        #endregion


    }
}
