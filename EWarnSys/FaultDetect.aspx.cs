using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.Web.Services;
using System.Text;
using Microsoft.DirectX.AudioVideoPlayback;
using System.Threading;

namespace warnSysforElec
{
    public partial class FaultDetect : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button_Check_Click(object sender, EventArgs e)
        {
            string road = "";
            try
            {
                string excelFile = this.fileuploadB.PostedFile.FileName;
                string relFilename = excelFile.Substring(excelFile.LastIndexOf("\\"));
                road = Server.MapPath("\\tmp\\check.xlsx");
                this.fileuploadB.PostedFile.SaveAs(road);

            }
            catch (Exception exx)
            {
                Response.Write("<script>alert('数据上传失败，请重新导入')</script>");
                return;
            }
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + road + ";Extended Properties=Excel 12.0;";
            DataSet dsMin = new DataSet();
            OleDbDataAdapter oada = new OleDbDataAdapter("select 时间,[1杆A],[1杆C],[25杆A],[25杆C],[101-6杆A],[101-6杆C],[102杆A],[102杆C] from [sheet1$]", strConn);
            oada.Fill(dsMin);
            this.gridview_C.DataSource = dsMin.Tables[0].DefaultView;
            this.gridview_C.DataBind();
            DataTable dt = dsMin.Tables[0];
            string strname = "";
            string[] breakpoint = new string[12] { "1杆C", "1杆", "1杆A", "25杆C", "25杆", "25杆A", "101-6杆C", "101-6杆", "101-6杆A", "102杆C", "102杆", "102杆A" };
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int flg = 0;
                int s = 0;
                int middle = 0;
                for (int k = 1; k < 9; k=k+2)
                {
                    if (k == 7 || k == 9)
                    {
                        if (Convert.ToDouble(dt.Rows[i][k].ToString()) >= Convert.ToDouble(dt.Rows[i][k - 2].ToString()) * 1.5)
                        {
                            middle = 1;
                        }
                        else
                        {
                            middle = -1;
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(dt.Rows[i][k + 2].ToString()) >= Convert.ToDouble(dt.Rows[i][k].ToString()) * 1.5)
                        {
                            middle = 1;
                        }
                        //(k>= 7 ? k-2:k)
                        else
                        {
                            middle = -1;
                        }
                    }
                    if (dt.Rows[i][k].ToString() == "")
                    {
                        if (dt.Rows[i][k + 1].ToString() == "")
                        {
                            flg = 1 + 6 * s;
                        }
                        else
                        {
                            flg = 2 + 6 * s;
                        }
                    }
                    else
                    {
                        if (dt.Rows[i][k + 1].ToString() == "")
                        {
                            flg = 3 + 6 * s;
                        }
                    }
                    if (dt.Rows[i][k].ToString() == "0")
                    {
                        if (dt.Rows[i][k + 1].ToString() == "0")
                        {
                            flg = 4 + 6 * s;
                        }
                        else
                        {
                            flg = 5 + 6 * s;
                        }
                    }
                    else
                    {
                        if (dt.Rows[i][k + 1].ToString() == "0")
                        {
                            flg = 6 + 6 * s;
                        }
                    }
                    s += 1;
                }
                int point = flg % 3 + ((flg-1)/6)*3;
                if (point == 1 && middle == -1)
                {
                    point = 2;
                }
                if (point == 11 && middle == 1)
                {
                    point = 10;
                }
                if (flg % 6 <= 3 && flg != 0 && point != 0)
                {
                    strname = breakpoint[point] + ((point == 6 ||point ==  7||point ==  8)? "": ((point == 9 ||point ==  10||point ==  11)? "---101主杆处断线":("---" + breakpoint[point + middle])));
                    Response.Write("<script>alert('时间：" + dt.Rows[i]["时间"].ToString() + " " + strname + "处断线')</script>");
                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171209230426.mp3"));
                    audio.Play();
                    for (Int64 a = 0; a < 100000; a++)
                    {

                    }
                }
                if (flg % 6 > 3 || flg % 6 == 0 && point != 0)
                {
                    strname = breakpoint[point] + ((point == 6 || point == 7 || point == 8) ? "" : ((point == 9 || point == 10 || point == 11) ? "---101主杆处接地" : ("---" + breakpoint[point + middle])));
                    Response.Write("<script>alert('时间：" + dt.Rows[i]["时间"].ToString() + " " +  strname + "处接地')</script>");
                    Audio audio = new Audio(Server.MapPath("\\Sound\\REC20171209230434.mp3"));
                    audio.Play();
                    for (Int64 a = 0; a < 100000; a++)
                    {

                    }
                }
                
            }
        }

    }
}