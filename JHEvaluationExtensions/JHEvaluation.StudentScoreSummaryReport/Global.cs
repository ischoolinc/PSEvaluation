using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using FISCA.Data;
using System.Data;
using Aspose.Words;
using JHSchool.Data;
using K12.Data;
using System.Xml.Linq;

namespace JHEvaluation.StudentScoreSummaryReport
{
    public class Global
    {
        //取得中英文對照
        private static SubjDomainEngNameMapping _SubjDomainEngNameMapping = new SubjDomainEngNameMapping();

        /// <summary>
        /// 匯出合併欄位總表Word(新版 在校成績證明書)
        /// </summary>
        public static void ExportMappingFieldWord(string doc)
        {

            #region 儲存檔案
            string inputReportName = "在校成績證明書合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            Document tempDoc = new Document(new MemoryStream(Properties.Resources.在校成績證明書_2022功能變數));
            if (doc == "英文版")
                tempDoc = new Document(new MemoryStream(Properties.Resources.在校成績證明書_2022英文版功能變數));

            try
            {
                #region 動態產生合併欄位
                // 讀取總表檔案並動態加入合併欄位
                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(tempDoc);
                builder.MoveToDocumentEnd();

                #region 缺曠動態產生合併欄位
                List<string> plist = K12.Data.PeriodMapping.SelectAll().Select(x => x.Type).Distinct().ToList();
                List<string> alist = K12.Data.AbsenceMapping.SelectAll().Select(x => x.Name).ToList();
                builder.Writeln();
                builder.Writeln();
                builder.Writeln("缺曠動態產生合併欄位");
                builder.StartTable();

                builder.InsertCell();
                builder.Write("缺曠名稱");
                builder.InsertCell();
                builder.Write("一上缺曠節數");
                builder.InsertCell();
                builder.Write("一下缺曠節數");
                builder.InsertCell();
                builder.Write("二上缺曠節數");
                builder.InsertCell();
                builder.Write("二下缺曠節數");
                builder.InsertCell();
                builder.Write("三上缺曠節數");
                builder.InsertCell();
                builder.Write("三下缺曠節數");
                builder.InsertCell();
                builder.Write("四上缺曠節數");
                builder.InsertCell();
                builder.Write("四下缺曠節數");
                builder.InsertCell();
                builder.Write("五上缺曠節數");
                builder.InsertCell();
                builder.Write("五下缺曠節數");
                builder.InsertCell();
                builder.Write("六上缺曠節數");
                builder.InsertCell();
                builder.Write("六下缺曠節數");

                builder.EndRow();

                foreach (string pp in plist)
                {
                    foreach (string aa in alist)
                    {

                        string key = pp.Replace(" ", "_") + "_" + aa.Replace(" ", "_");
                        builder.InsertCell();
                        builder.Write(key);

                        for (int i = 1; i <= 12; i++)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + key + i + " \\* MERGEFORMAT ", "«" + key + i + "»");
                        }
                        builder.EndRow();
                    }
                }

                builder.EndTable();
                #endregion

                #region 日常生活表現
                //builder.Writeln();
                //builder.Writeln();
                //builder.Writeln("日常生活表現評量");
                //builder.StartTable();
                //builder.InsertCell();
                //builder.Write("分類");
                //builder.InsertCell();
                //builder.Write("名稱");
                //builder.InsertCell();
                //builder.Write("建議內容");
                //builder.EndRow();

                //builder.EndTable();

                //// 日常生活表現
                //builder.Writeln();
                //builder.Writeln();
                //builder.Writeln("日常生活表現評量子項目");
                //builder.StartTable();
                //builder.InsertCell();
                //builder.Write("項目");
                //builder.InsertCell();
                //builder.Write("指標");
                //builder.InsertCell();
                //builder.Write("表現程度");
                //builder.EndRow();

                //for (int i = 1; i <= 7; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Name" + i + " \\* MERGEFORMAT ", "«項目" + i + "»");
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Index" + i + " \\* MERGEFORMAT ", "«指標" + i + "»");
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD " + "日常生活表現程度_Item_Degree" + i + " \\* MERGEFORMAT ", "«表現" + i + "»");
                //    builder.EndRow();
                //}

                //builder.EndTable();
                #endregion

                //List<string> DomainList = new List<string> { "語文", "數學", "社會", "自然科學", "自然與生活科技", "藝術", "藝術與人文", "健康與體育", "綜合活動", "科技", "彈性課程" };
                List<string> DomainList = Util.GetDomainList();

                #region 學期領域成績
                builder.Writeln();
                builder.Writeln();
                builder.Writeln("學期領域成績");
                builder.StartTable();
                builder.InsertCell();
                builder.Write("領域");
                if (doc == "英文版")
                {
                    builder.InsertCell();
                    builder.Write("領域英文");
                }

                builder.InsertCell();
                builder.Write("一上權數");
                builder.InsertCell();
                builder.Write("一上成績");
                builder.InsertCell();
                builder.Write("一上等第");
                builder.InsertCell();
                builder.Write("一下權數");
                builder.InsertCell();
                builder.Write("一下成績");
                builder.InsertCell();
                builder.Write("一下等第");
                builder.InsertCell();
                builder.Write("二上權數");
                builder.InsertCell();
                builder.Write("二上成績");
                builder.InsertCell();
                builder.Write("二上等第");
                builder.InsertCell();
                builder.Write("二下權數");
                builder.InsertCell();
                builder.Write("二下成績");
                builder.InsertCell();
                builder.Write("二下等第");
                builder.InsertCell();
                builder.Write("三上權數");
                builder.InsertCell();
                builder.Write("三上成績");
                builder.InsertCell();
                builder.Write("三上等第");
                builder.InsertCell();
                builder.Write("三下權數");
                builder.InsertCell();
                builder.Write("三下成績");
                builder.InsertCell();
                builder.Write("三下等第");
                builder.InsertCell();
                builder.Write("平均成績");
                builder.InsertCell();
                builder.Write("平均成績等第");
                builder.EndRow();

                foreach (string domain in DomainList)
                {
                    if (domain == "彈性課程")

                        continue;


                    builder.InsertCell();
                    builder.Write(domain);
                    if (doc == "英文版")
                    {
                        builder.InsertCell();
                        builder.Write(_SubjDomainEngNameMapping.GetDomainEngName(domain));
                    }

                    for (int i = 1; i <= 6; i++)
                    {
                        //領域_語文_成績_1
                        string scoreKey = "領域_" + domain + "_成績_" + i;
                        string levelKey = "領域_" + domain + "_等第_" + i;
                        string creditKey = "領域_" + domain + "_權數_" + i;

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + creditKey + " \\* MERGEFORMAT ", "«C" + i + "»");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + scoreKey + " \\* MERGEFORMAT ", "«S" + i + "»");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + levelKey + " \\* MERGEFORMAT ", "«D" + i + "»");
                    }
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "領域_" + domain + "_平均成績" + " \\* MERGEFORMAT ", "«SA»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "領域_" + domain + "_平均成績等第" + " \\* MERGEFORMAT ", "«LA»");

                    builder.EndRow();
                }

                builder.InsertCell();
                builder.Write("學習領域總成績");
                if (doc == "英文版")
                {
                    builder.InsertCell();
                    builder.Write("Weighted Average Score");
                }
                for (int i = 1; i <= 6; i++)
                {
                    //領域_學習領域總成績_成績_1
                    string scoreKey = "領域_" + "學習領域總成績" + "_成績_" + i;
                    string levelKey = "領域_" + "學習領域總成績" + "_等第_" + i;

                    builder.InsertCell();
                    builder.Write("");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + scoreKey + " \\* MERGEFORMAT ", "«S" + i + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + levelKey + " \\* MERGEFORMAT ", "«D" + i + "»");

                }
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "領域_" + "學習領域總平均成績" + "_成績" + " \\* MERGEFORMAT ", "«SA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + "領域_" + "學習領域總平均成績" + "_等第" + " \\* MERGEFORMAT ", "«LA»");

                builder.EndRow();
                builder.EndTable();
                #endregion

                #region 學期科目成績
                foreach (string domain in DomainList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    builder.Writeln(domain + "領域科目成績");
                    builder.StartTable();

                    builder.InsertCell();
                    builder.Write("科目");
                    builder.InsertCell();
                    builder.Write("一上權數");
                    builder.InsertCell();
                    builder.Write("一上成績");
                    builder.InsertCell();
                    builder.Write("一上等第");
                    builder.InsertCell();
                    builder.Write("一下權數");
                    builder.InsertCell();
                    builder.Write("一下成績");
                    builder.InsertCell();
                    builder.Write("一下等第");
                    builder.InsertCell();
                    builder.Write("二上權數");
                    builder.InsertCell();
                    builder.Write("二上成績");
                    builder.InsertCell();
                    builder.Write("二上等第");
                    builder.InsertCell();
                    builder.Write("二下權數");
                    builder.InsertCell();
                    builder.Write("二下成績");
                    builder.InsertCell();
                    builder.Write("二下等第");
                    builder.InsertCell();
                    builder.Write("三上權數");
                    builder.InsertCell();
                    builder.Write("三上成績");
                    builder.InsertCell();
                    builder.Write("三上等第");
                    builder.InsertCell();
                    builder.Write("三下權數");
                    builder.InsertCell();
                    builder.Write("三下成績");
                    builder.InsertCell();
                    builder.Write("三下等第");
                    builder.InsertCell();
                    builder.Write("平均成績");
                    builder.InsertCell();
                    builder.Write("平均成績等第");
                    builder.EndRow();

                    //1上
                    for (int i = 1; i <= 6; i++)
                    {
                        string subjectKey = domain + "_科目名稱" + i;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + subjectKey + " \\* MERGEFORMAT ", "«N" + i + "»");

                        for (int a = 1; a <= 6; a++)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_權數" + a + " \\* MERGEFORMAT ", "«C" + a + "»");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_成績" + a + " \\* MERGEFORMAT ", "«S" + a + "»");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_等第" + a + " \\* MERGEFORMAT ", "«L" + a + "»");
                        }
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_平均成績" + " \\* MERGEFORMAT ", "«SA»");
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_平均成績等第" + " \\* MERGEFORMAT ", "«LA»");

                        builder.EndRow();
                    }

                    if (domain == "彈性課程")
                        for (int i = 7; i <= 18; i++)
                        {
                            string subjectKey = domain + "_科目名稱" + i;
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + subjectKey + " \\* MERGEFORMAT ", "«N" + i + "»");

                            for (int a = 1; a <= 6; a++)
                            {
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_權數" + a + " \\* MERGEFORMAT ", "«C" + a + "»");
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_成績" + a + " \\* MERGEFORMAT ", "«S" + a + "»");
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_等第" + a + " \\* MERGEFORMAT ", "«L" + a + "»");
                            }
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_平均成績" + " \\* MERGEFORMAT ", "«SA»");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_平均成績等第" + " \\* MERGEFORMAT ", "«LA»");

                            builder.EndRow();

                        }

                    builder.EndTable();
                }
                #endregion

                #region 學期科目原始成績
                foreach (string domain in DomainList)
                {
                    builder.Writeln();
                    builder.Writeln();
                    builder.Writeln(domain + "領域科目原始成績");
                    builder.StartTable();

                    builder.InsertCell();
                    builder.Write("科目");
                    builder.InsertCell();
                    builder.Write("一上權數");
                    builder.InsertCell();
                    builder.Write("一上成績");
                    builder.InsertCell();
                    builder.Write("一上等第");
                    builder.InsertCell();
                    builder.Write("一下權數");
                    builder.InsertCell();
                    builder.Write("一下成績");
                    builder.InsertCell();
                    builder.Write("一下等第");
                    builder.InsertCell();
                    builder.Write("二上權數");
                    builder.InsertCell();
                    builder.Write("二上成績");
                    builder.InsertCell();
                    builder.Write("二上等第");
                    builder.InsertCell();
                    builder.Write("二下權數");
                    builder.InsertCell();
                    builder.Write("二下成績");
                    builder.InsertCell();
                    builder.Write("二下等第");
                    builder.InsertCell();
                    builder.Write("三上權數");
                    builder.InsertCell();
                    builder.Write("三上成績");
                    builder.InsertCell();
                    builder.Write("三上等第");
                    builder.InsertCell();
                    builder.Write("三下權數");
                    builder.InsertCell();
                    builder.Write("三下成績");
                    builder.InsertCell();
                    builder.Write("三下等第");
                    builder.EndRow();
                    //1上
                    for (int i = 1; i <= 6; i++)
                    {
                        string subjectKey = domain + "_科目名稱" + i;
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + subjectKey + " \\* MERGEFORMAT ", "«N" + i + "»");

                        for (int a = 1; a <= 6; a++)
                        {
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_權數" + a + " \\* MERGEFORMAT ", "«C" + a + "»");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_原始成績" + a + " \\* MERGEFORMAT ", "«S" + a + "»");
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_原始等第" + a + " \\* MERGEFORMAT ", "«L" + a + "»");
                        }
                        builder.EndRow();
                    }

                    if (domain == "彈性課程")
                        for (int i = 7; i <= 18; i++)
                        {
                            string subjectKey = domain + "_科目名稱" + i;
                            builder.InsertCell();
                            builder.InsertField("MERGEFIELD " + subjectKey + " \\* MERGEFORMAT ", "«N" + i + "»");

                            for (int a = 1; a <= 6; a++)
                            {
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_權數" + a + " \\* MERGEFORMAT ", "«C" + a + "»");
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_原始成績" + a + " \\* MERGEFORMAT ", "«S" + a + "»");
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + domain + "_科目" + i + "_原始等第" + a + " \\* MERGEFORMAT ", "«L" + a + "»");
                            }

                            builder.EndRow();

                        }

                    builder.EndTable();
                }
                #endregion


                #endregion

                tempDoc.Save(path, SaveFormat.Doc);

                System.Diagnostics.Process.Start(path);


            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        tempDoc.Save(sd.FileName, SaveFormat.Doc);

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }
    }
}
