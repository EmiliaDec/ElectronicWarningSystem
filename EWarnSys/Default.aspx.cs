using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Web.Services;
using System.Text;
using Microsoft.DirectX.AudioVideoPlayback;
using System.Threading;
namespace warnSysforElec
{
    public partial class _Default : System.Web.UI.Page
    {
        public void Page_Load(object sender, EventArgs e)
        {
            

        }
        public void Button_OK_Click(object sender, EventArgs e)
        {
           string road = "";
           try  
           { 
               string excelFile = this.FileUpload1.PostedFile.FileName;
               string relFilename = excelFile.Substring(excelFile.LastIndexOf("\\"));
               road = Server.MapPath("\\tmp\\tmp.xlsx");
               this.FileUpload1.PostedFile.SaveAs(road);
               
           }
           catch(Exception exx)
           {
               //数据上传失败，请重新导入' 
               Response.Write("<script>alert('数据上传失败，请重新导入')</script>");
               return;
           }
               string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + road + ";Extended Properties=Excel 12.0;"; 
               DataSet dsMin = new DataSet();
               OleDbDataAdapter oada = new OleDbDataAdapter("select 时间,电流,(电流*电流)*3.5*0.000001 as 有功损耗,(电流*电流)*1.5*0.000001 as 无功损耗,温度 from [sheet1$]", strConn);  
               oada.Fill(dsMin);
               this.gridview_W.DataSource = dsMin.Tables[0].DefaultView;
               this.gridview_W.DataBind();
          
        }
        public void ButtonTimelyICheck(object sender, EventArgs e)
        {
            string road = Server.MapPath("\\tmp\\tmp.xlsx");
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + road + ";Extended Properties=Excel 12.0;";
            DataSet dsMin = new DataSet();
            OleDbDataAdapter oada = new OleDbDataAdapter("select 时间,电流,(电流*电流)*3.5*0.000001 as 有功损耗,(电流*电流)*1.5*0.000001 as 无功损耗,温度 from [sheet1$]", strConn);
            oada.Fill(dsMin);
            DataTable dt = dsMin.Tables[0];
            string strName = "";
            double rate = 0.0;
            double temprature = 0.0;
            int flgB = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                strName = dt.Rows[i]["电流"].ToString();
                temprature = Convert.ToDouble(dt.Rows[i]["温度"].ToString());
                double I = Convert.ToDouble(strName);
                if (I > 405 / 2)
                {
                    rate = 10 * I * I * 0.000001 - 0.029;
                    //
                    if (rate > 0.7 && I > 405 * 3 / 5)
                    {
                        flgB = 1;
                        Response.Write("<script>alert('" + dt.Rows[i]["时间"].ToString() + "电流超出阈值')</script>");
                        break;
                    }
                }
            }
            try
            {
                if (flgB == 1)
                {
                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171203093642.mp3"));
                    audio.Play();
                    for (Int64 i = 0; i < 100000; i++)
                    {

                    }

                }

            }
            catch(Exception ex)
            {
               
               Console.Write(ex.Message);
            }
        }

//        public void ButtonTimelyWCheck(object sender, EventArgs e)
//        {
//            string road = Server.MapPath("\\tmp\\tmp.xlsx");
//            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + road + ";Extended Properties=Excel 12.0;";
//            DataSet dsMin = new DataSet();
//            OleDbDataAdapter oada = new OleDbDataAdapter("select 时间,电压,电流,(电流*电流)*93.24 as 有功损耗,(电流*电流)*92.31 as 无功损耗,温度 from [sheet1$]", strConn);
//            oada.Fill(dsMin);
//            DataTable dt = dsMin.Tables[0];
//            double rate = 0.0;
//            double temprature = 0.0;
//            int flgB = 0;
//            for (int i = 0; i < dt.Rows.Count; i++)
//            {
//                temprature = Convert.ToDouble(dt.Rows[i]["温度"].ToString());
//                double W = Convert.ToDouble(dt.Rows[i]["电流"].ToString()) * Convert.ToDouble(dt.Rows[i]["电压"].ToString());
//                if (W > 405 / 2)
//                {
//                    rate = 10 * W * W * 0.000001 - 0.029;
//                    if (rate > 0.7 && W > 10 * 3 / 5)
//                    {
//                        flgB = 1;
//                        Response.Write("<script>alert('" + dt.Rows[i]["时间"].ToString() + "负荷有功超出阈值')</script>");
//                        break;
//                    }
//                }
//            }
//            try
//            {
//                if (flgB == 1)
//                {
//                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171209225930.mp3"));
//                    audio.Play();
//                    for (Int64 i = 0; i < 100000; i++)
//                    {
//
//                    }
//
//                }
//
//            }
//            catch (Exception ex)
//            {
//
//                Console.Write(ex.Message);
//            }
//        }
        public void ButtonTomorrowPredict_I(object sender, EventArgs e)
        {
            string road = Server.MapPath("\\tmp\\tmp.xlsx");
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + road + ";Extended Properties=Excel 12.0;";
            DataSet dsMin = new DataSet();
            OleDbDataAdapter oada = new OleDbDataAdapter("select 时间,电流,(电流*电流)*3.5*0.000001 as 有功损耗,(电流*电流)*1.5*0.000001 as 无功损耗,温度 from [sheet1$]", strConn);
            oada.Fill(dsMin);
            DataTable dt = dsMin.Tables[0];
            string strName = "";
            double temprature = 100;
            double hI = 0.0;
            double I = 0.0;
            int flgA = 0;
            double tomorrowtpr = 100;
            if (TB_temprature.Text.ToString() != "")
            {
                 tomorrowtpr = Convert.ToDouble(TB_temprature.Text.ToString());
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                strName = dt.Rows[i]["电流"].ToString();
                temprature = Convert.ToDouble(dt.Rows[i]["温度"].ToString());
                if (I < Convert.ToDouble(strName))
                {
                    I = Convert.ToDouble(strName);
                }
            }
            string curdate = dt.Rows[dt.Rows.Count -1]["时间"].ToString().Substring(0, 10);
            string hroad = Server.MapPath("\\tmp\\history.xlsx");
            string hstrConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + hroad + ";Extended Properties=Excel 12.0;";
            DataSet hdsMin = new DataSet();
            OleDbDataAdapter hoada = new OleDbDataAdapter("select 时间,电流,负荷有功,温度 from [sheet1$]", hstrConn);
            hoada.Fill(hdsMin);
            DataTable hdt = hdsMin.Tables[0];
            for (int j = 0; j < hdt.Rows.Count; j++)
            {
                if (hdt.Rows[j]["时间"].ToString().IndexOf(curdate) != -1)
                {
                    if (hI < Convert.ToDouble(hdt.Rows[j]["电流"].ToString()))
                    {
                        hI = Convert.ToDouble(hdt.Rows[j]["电流"].ToString());
                    }
                }

            }

            if (temprature >= 5 && temprature <= 25 &&  (tomorrowtpr < 5 || tomorrowtpr > 25) )
            {

                if (I > hI*1.3 && I > 405 * 3 / 5)
                {
                    flgA = 1;
                    Response.Write("<script>alert('预测明天电流超过额定值百分之六十，需要注意')</script>");
                }

            }
            else
            {

                if (I > hI * 1.5 && I > 405 * 3 / 5)
                {
                    flgA = 1;
                    Response.Write("<script>alert('预测明天电流超过额定值百分之六十，需要注意')</script>");
                }
                else
                {
                    flgA = 0;
                }
            }
            try
            {
                if (flgA == 1)
                {
                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171203094954.mp3"));
                    audio.Play();
                    for (Int64 i = 0; i < 100000; i++)
                    {

                    }
                }
                if (flgA == 0)
                {
                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171203094935.mp3"));
                    audio.Play();
                    for (Int64 i = 0; i < 100000; i++)
                    {

                    }
                }

            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
        public void ButtonTimelyWWCheck(object sender, EventArgs e)
        {
            string road = Server.MapPath("\\tmp\\tmp.xlsx");
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + road + ";Extended Properties=Excel 12.0;";
            DataSet dsMin = new DataSet();
            OleDbDataAdapter oada = new OleDbDataAdapter("select 时间,负荷有功,(电流*电流)*3.5*0.000001 as 有功损耗 from [sheet1$]", strConn);
            oada.Fill(dsMin);
            DataTable dt = dsMin.Tables[0];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (Convert.ToDouble(dt.Rows[i]["负荷有功"].ToString()) * 0.2 < Convert.ToDouble(dt.Rows[i]["有功损耗"].ToString()))
                {
                 Response.Write("<script>alert('时间：" + dt.Rows[i]["时间"].ToString() + " 线损超过百分之二，需要注意')</script>");
                }
            }
        }
        public void ButtonTomorrowPredict_W(object sender, EventArgs e)
        {
            string road = Server.MapPath("\\tmp\\tmp.xlsx");
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + road + ";Extended Properties=Excel 12.0;";
            DataSet dsMin = new DataSet();
            OleDbDataAdapter oada = new OleDbDataAdapter("select 时间,电流,负荷有功,(电流*电流)*3.5*0.000001 as 有功损耗,(电流*电流)*1.5*0.000001 as 无功损耗,温度 from [sheet1$]", strConn);
            oada.Fill(dsMin);
            DataTable dt = dsMin.Tables[0];
            double temprature = 100.0;
            double hW = 0.0;
            double W = 0.0;
            int flgA = 0;
            double tomorrowtpr = 100;
            if (TB_temprature.Text.ToString() != "")
            {
                tomorrowtpr = Convert.ToDouble(TB_temprature.Text.ToString());
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                temprature = Convert.ToDouble(dt.Rows[i]["温度"].ToString());
                if (W < Convert.ToDouble(dt.Rows[i]["负荷有功"].ToString()))
                {
                    W = Convert.ToDouble(dt.Rows[i]["负荷有功"].ToString());
                }
            }
            string curdate = dt.Rows[dt.Rows.Count - 1]["时间"].ToString().Substring(0, 10);
            string hroad = Server.MapPath("\\tmp\\history.xlsx");
            string hstrConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + hroad + ";Extended Properties=Excel 12.0;";
            DataSet hdsMin = new DataSet();
            OleDbDataAdapter hoada = new OleDbDataAdapter("select 时间,电流,负荷有功,温度 from [sheet1$]", hstrConn);
            hoada.Fill(hdsMin);
            DataTable hdt = hdsMin.Tables[0];
            for (int j = 0; j < hdt.Rows.Count; j++)
            {
                if (hdt.Rows[j]["时间"].ToString().IndexOf(curdate) != -1)
                {
                    if (hW < Convert.ToDouble(hdt.Rows[j]["负荷有功"].ToString()))
                    {
                        hW = Convert.ToDouble(hdt.Rows[j]["负荷有功"].ToString());
                    }
                }

            }
            if (temprature >= 5 && temprature <= 25 && (tomorrowtpr < 5 || tomorrowtpr > 25))
            {

                if (W > hW * 1.3 && W > 6 * 7/ 10)
                {
                    flgA = 1;
                    Response.Write("<script>alert('预测明天负荷超过额定值百分之七十，需要注意')</script>");
                }

            }
            else
            {

                if (W > hW * 1.5 && W > 6 * 7/ 10)
                {
                    flgA = 1;
                    Response.Write("<script>alert('预测明天负荷超过额定值百分之七十，需要注意')</script>");
                }
                else
                {
                    flgA = 0;
                }
            }
            try
            {
                if (flgA == 1)
                {
                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171209230003.mp3"));
                    audio.Play();
                    for (Int64 i = 0; i < 100000; i++)
                    {

                    }
                }
                if (flgA == 0)
                {
                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171209225954.mp3"));
                    audio.Play();
                    for (Int64 i = 0; i < 100000; i++)
                    {

                    }
                }

            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
        void speech_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            //.Text = "语音试听";
        }
        

    }
}