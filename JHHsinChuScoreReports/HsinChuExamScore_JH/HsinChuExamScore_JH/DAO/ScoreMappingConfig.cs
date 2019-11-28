using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;
using System.Xml.Linq;
using System.Windows.Forms;


namespace HsinChuExamScore_JH.DAO
{
    /// <summary>
    /// 等第對照表
    /// </summary>
    public class ScoreMappingConfig
    {
        XElement elmRoot;
        // 處理沒有成績對照最小等第
        string minScoreName = "";
        Dictionary<decimal, string> scoreNameDict = new Dictionary<decimal, string>();
        /// <summary>
        /// 載入資料
        /// </summary>
        public void LoadData()
        {            
            try
            {
                minScoreName = "";
                scoreNameDict.Clear();
                QueryHelper qh = new QueryHelper();
                string query = "SELECT content FROM list WHERE name ='等第對照表';";
                DataTable dt = qh.Select(query);


                if (dt != null && dt.Rows.Count > 0)
                {
                    elmRoot = XElement.Parse(dt.Rows[0]["content"].ToString());
                    if (elmRoot != null)
                    {
                        XElement elmScoreMappingList = XElement.Parse(elmRoot.Element("Configuration").Value);

                        if (elmScoreMappingList != null)
                        {
                            foreach(XElement elm in elmScoreMappingList.Elements("ScoreMapping"))
                            {
                                string scName = "";
                                if (elm.Attribute("Name") != null && elm.Attribute("Name").Value != "")
                                    scName = elm.Attribute("Name").Value;

                                if (elm.Attribute("Score") != null)
                                {
                                    if (elm.Attribute("Score").Value == "")
                                    {
                                        minScoreName = scName;
                                    }
                                    else
                                    {
                                        decimal sc;
                                        if (decimal.TryParse(elm.Attribute("Score").Value, out sc))
                                        {
                                            if (!scoreNameDict.ContainsKey(sc))
                                                scoreNameDict.Add(sc, scName);

                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("解析等第對照發生錯誤" + ex.Message);
                return;
            }
        }

        public string ParseScoreName(decimal score)
        {
            string value = minScoreName;

            foreach( decimal sc in scoreNameDict.Keys)
            {
                if (score >= sc)
                {
                    value = scoreNameDict[sc];
                    break;
                }
            }
            return value;
        }

    }
}
