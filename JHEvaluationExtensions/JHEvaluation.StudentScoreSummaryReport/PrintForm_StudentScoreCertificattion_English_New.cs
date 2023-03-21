using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Aspose.Words;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using JHEvaluation.ScoreCalculation;
using Campus.Rating;
using JHSchool.Data;
using Campus.Report;
using JHSchool.Behavior.BusinessLogic;
using System.IO;
using System.Data;
using System.Globalization;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using FISCA.Data;
using Framework;
using System.Linq;
using ReportHelper;

namespace JHEvaluation.StudentScoreSummaryReport
{
    public partial class PrintForm_StudentScoreCertificattion_English_New : BaseForm, IStatusReporter
    {
        internal const string ConfigName = "StudentScoreSummaryReportEnglish2022";

        private List<string> StudentIDs { get; set; }

        private ReportPreference Preference { get; set; }

        private BackgroundWorker MasterWorker = new BackgroundWorker();

        private BackgroundWorker ConvertToPDF_Worker = new BackgroundWorker();

        private List<K12.Data.StudentRecord> PrintStudents = new List<K12.Data.StudentRecord>();

        // 學生基本資料 [studentID,Data]
        Dictionary<string, K12.Data.StudentRecord> sr_dict = new Dictionary<string, K12.Data.StudentRecord>();

        // 照片
        Dictionary<string, string> _PhotoPDict = new Dictionary<string, string>();

        // 學期歷程 [studentID,Data]
        Dictionary<string, K12.Data.SemesterHistoryRecord> shr_dict = new Dictionary<string, K12.Data.SemesterHistoryRecord>();

        //可以印出來的領域
        List<string> DomainList = Util.GetDomainList();

        // 缺曠 [studentID,List<Data>]
        Dictionary<string, List<K12.Data.AttendanceRecord>> ar_dict = new Dictionary<string, List<K12.Data.AttendanceRecord>>();
        List<string> plist = K12.Data.PeriodMapping.SelectAll().Select(x => x.Type).Distinct().ToList();
        List<string> alist = K12.Data.AbsenceMapping.SelectAll().Select(x => x.Name).ToList();
        Dictionary<string, string> PeriodMappingDic = Util.GetPeriodMappingDict();

        //學期成績(領域、科目) [studentID,List<Data>]
        Dictionary<string, List<JHSemesterScoreRecord>> jssr_dict = new Dictionary<string, List<JHSemesterScoreRecord>>();

        //畢業分數 [studentID,Data]
        Dictionary<string, K12.Data.GradScoreRecord> gsr_dict = new Dictionary<string, K12.Data.GradScoreRecord>();

        // 異動 [studentID,List<Data>]
        Dictionary<string, List<K12.Data.UpdateRecordRecord>> urr_dict = new Dictionary<string, List<K12.Data.UpdateRecordRecord>>();

        // 日常生活表現、校內外特殊表現 [studentID,List<Data>]
        Dictionary<string, List<K12.Data.MoralScoreRecord>> msr_dict = new Dictionary<string, List<K12.Data.MoralScoreRecord>>();

        //取得中英文對照
        private static SubjDomainEngNameMapping _SubjDomainEngNameMapping = new SubjDomainEngNameMapping();

        /// <summary>
        /// 等第對照表
        /// </summary>
        ScoreMappingConfig _ScoreMappingConfig = new ScoreMappingConfig();

        private string fbdPath = "";

        private DoWorkEventArgs e_For_ConvertToPDF_Worker;

        public PrintForm_StudentScoreCertificattion_English_New(List<string> studentIds)
        {
            InitializeComponent();

            StudentIDs = studentIds;
            Preference = new ReportPreference(ConfigName, Properties.Resources.在校成績證明書_2022英文版_範本);
            MasterWorker.DoWork += new DoWorkEventHandler(MasterWorker_DoWork);
            MasterWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(MasterWorker_RunWorkerCompleted);
            MasterWorker.WorkerReportsProgress = true;
            MasterWorker.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage(e.UserState.ToString(), e.ProgressPercentage);
            };


            ConvertToPDF_Worker.DoWork += new DoWorkEventHandler(ConvertToPDF_Worker_DoWork);
            ConvertToPDF_Worker.WorkerReportsProgress = true;
            ConvertToPDF_Worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ConvertToPDF_Worker_RunWorkerCompleted);

            ConvertToPDF_Worker.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage(e.UserState.ToString(), e.ProgressPercentage);
            };

            // 是否列印PDF
            rtnPDF.Checked = Preference.ConvertToPDF;

            // 是否要單檔列印
            OneFileSave.Checked = Preference.OneFileSave;

            _ScoreMappingConfig.LoadData();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            Preference.ConvertToPDF = rtnPDF.Checked;

            Preference.OneFileSave = OneFileSave.Checked;

            Preference.Save(); //儲存設定值。

            //關閉畫面控制項
            Util.DisableControls(this);
            MasterWorker.RunWorkerAsync();
        }

        private void MasterWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (StudentIDs.Count <= 0)
            {
                Feedback("", -1);  //把 Status bar Reset...
                throw new ArgumentException("沒有任何學生資料可列印。");
            }

            #region 抓取學生資料 
            //抓取學生資料 
            //學生基本資料
            List<K12.Data.StudentRecord> sr_list = K12.Data.Student.SelectByIDs(StudentIDs);

            //學期成績(包含領域、科目) (改用 JHSemesterScoreRecord 才抓的到)
            List<JHSemesterScoreRecord> ssr_list = JHSemesterScore.SelectByStudentIDs(StudentIDs);

            //缺曠
            List<K12.Data.AttendanceRecord> ar_list = K12.Data.Attendance.SelectByStudentIDs(StudentIDs);

            //學生異動
            List<K12.Data.UpdateRecordRecord> urr_list = K12.Data.UpdateRecord.SelectByStudentIDs(StudentIDs);

            //畢業分數(暫不處理--舊版國中在校成績證明書沒有，Cynthia)             此處SelectByIDs 確實是SelectByStudentIDs
            List<K12.Data.GradScoreRecord> gsr_list = K12.Data.GradScore.SelectByIDs<K12.Data.GradScoreRecord>(StudentIDs);

            //學期歷程
            List<K12.Data.SemesterHistoryRecord> shr_list = K12.Data.SemesterHistory.SelectByStudentIDs(StudentIDs);

            //日常生活表現、校內外特殊表現(暫不處理--舊版國中在校成績證明書沒有，Cynthia)
            List<K12.Data.MoralScoreRecord> msr_list = K12.Data.MoralScore.SelectByStudentIDs(StudentIDs);

            // 服務學習統計
            Util._SLRDict = Util.GetServiceLearningDetail(StudentIDs);

            //獎懲統計
            Util._DisciplineDict = Util.GetDisciplineDetail(StudentIDs);
            #endregion

            #region 整理學生基本資料
            //整理學生基本資料
            foreach (K12.Data.StudentRecord sr in sr_list)
            {
                if (!sr_dict.ContainsKey(sr.ID))
                {
                    sr_dict.Add(sr.ID, sr);
                }
            }

            //整理學期歷程
            foreach (K12.Data.SemesterHistoryRecord shr in shr_list)
            {
                if (!shr_dict.ContainsKey(shr.RefStudentID))
                {
                    shr_dict.Add(shr.RefStudentID, shr);
                }
            }

            //整理學期成績(包含領域、科目) 紀錄
            foreach (JHSemesterScoreRecord ssr in ssr_list)
            {
                if (!jssr_dict.ContainsKey(ssr.RefStudentID))
                {
                    jssr_dict.Add(ssr.RefStudentID, new List<JHSemesterScoreRecord>());
                    jssr_dict[ssr.RefStudentID].Add(ssr);
                }
                else
                {
                    jssr_dict[ssr.RefStudentID].Add(ssr);
                }
            }

            //整理出缺勤紀錄
            foreach (K12.Data.AttendanceRecord ar in ar_list)
            {
                if (!ar_dict.ContainsKey(ar.RefStudentID))
                {
                    ar_dict.Add(ar.RefStudentID, new List<K12.Data.AttendanceRecord>());
                    ar_dict[ar.RefStudentID].Add(ar);
                }
                else
                {
                    ar_dict[ar.RefStudentID].Add(ar);
                }
            }

            //整理畢業分數
            //foreach (K12.Data.GradScoreRecord gsr in gsr_list)
            //{
            //    if (!gsr_dict.ContainsKey(gsr.RefStudentID))
            //    {
            //        gsr_dict.Add(gsr.RefStudentID, gsr);
            //    }
            //}


            //整理出日常生活表現、校內外特殊表現
            foreach (K12.Data.MoralScoreRecord msr in msr_list)
            {
                if (!msr_dict.ContainsKey(msr.RefStudentID))
                {
                    msr_dict.Add(msr.RefStudentID, new List<K12.Data.MoralScoreRecord>());
                    msr_dict[msr.RefStudentID].Add(msr);
                }
                else
                {
                    msr_dict[msr.RefStudentID].Add(msr);
                }
            }

            //整理出異動紀錄
            foreach (K12.Data.UpdateRecordRecord urr in urr_list)
            {
                if (!urr_dict.ContainsKey(urr.StudentID))
                {
                    urr_dict.Add(urr.StudentID, new List<K12.Data.UpdateRecordRecord>());
                    urr_dict[urr.StudentID].Add(urr);
                }
                else
                {
                    urr_dict[urr.StudentID].Add(urr);
                }
            }
            #endregion

            #region 建立合併欄位總表
            //建立合併欄位總表
            DataTable table = new DataTable();

            List<string> paList = new List<string>();
            // 缺曠欄位
            foreach (string aa in alist)
            {
                foreach (string pp in plist)
                {
                    string key = pp.Replace(" ", "_") + "_" + aa.Replace(" ", "_");
                    for (int i = 1; i <= 12; i++)
                    {
                        if (!paList.Contains(key + i))
                        {
                            table.Columns.Add(key + i);
                            paList.Add(key + i);
                        }

                    }

                }
            }


            #region 基本資料
            //基本資料
            table.Columns.Add("學校名稱");
            table.Columns.Add("學校地址");
            table.Columns.Add("學校電話");
            table.Columns.Add("校長");
            table.Columns.Add("教務主任");
            table.Columns.Add("學生姓名");
            table.Columns.Add("學生英文姓名");
            table.Columns.Add("學生座號");
            table.Columns.Add("學生班級");
            table.Columns.Add("學生性別");
            table.Columns.Add("出生日期");
            table.Columns.Add("入學年月");
            table.Columns.Add("畢業年月");
            table.Columns.Add("入學日期");
            table.Columns.Add("畢業日期");
            table.Columns.Add("學生身分證字號");
            table.Columns.Add("學號");
            table.Columns.Add("照片", typeof(byte[]));
            table.Columns.Add("國籍一");
            table.Columns.Add("國籍一護照名");
            table.Columns.Add("國籍二");
            table.Columns.Add("國籍二護照名");
            table.Columns.Add("國籍一英文");
            table.Columns.Add("國籍二英文");

            #endregion

            #region 異動紀錄
            //異動紀錄
            //table.Columns.Add("異動紀錄1_日期");
            //table.Columns.Add("異動紀錄1_校名");
            //table.Columns.Add("異動紀錄1_學號");
            //table.Columns.Add("異動紀錄2_日期");
            //table.Columns.Add("異動紀錄2_校名");
            //table.Columns.Add("異動紀錄2_學號");
            //table.Columns.Add("異動紀錄3_日期");
            //table.Columns.Add("異動紀錄3_校名");
            //table.Columns.Add("異動紀錄3_學號");
            //table.Columns.Add("異動紀錄4_日期");
            //table.Columns.Add("異動紀錄4_校名");
            //table.Columns.Add("異動紀錄4_學號");
            #endregion

            #region 班級座號資料
            //班級座號資料
            table.Columns.Add("年級1_班級");
            table.Columns.Add("年級1_座號");
            table.Columns.Add("年級2_班級");
            table.Columns.Add("年級2_座號");
            table.Columns.Add("年級3_班級");
            table.Columns.Add("年級3_座號");
            table.Columns.Add("年級4_班級");
            table.Columns.Add("年級4_座號");
            table.Columns.Add("年級5_班級");
            table.Columns.Add("年級5_座號");
            table.Columns.Add("年級6_班級");
            table.Columns.Add("年級6_座號");
            #endregion

            #region 學年度
            //學年度
            table.Columns.Add("學年度1");
            table.Columns.Add("學年度2");
            table.Columns.Add("學年度3");
            table.Columns.Add("學年度4");
            table.Columns.Add("學年度5");
            table.Columns.Add("學年度6");
            #endregion

            #region 領域成績
            //領域成績

            foreach (string domain in DomainList)
            {
                for (int i = 1; i <= 12; i++)
                {
                    table.Columns.Add("領域_" + domain + "_成績_" + i);
                    table.Columns.Add("領域_" + domain + "_等第_" + i);
                    table.Columns.Add("領域_" + domain + "_權數_" + i);
                }
                table.Columns.Add("領域_" + domain + "_平均成績");
                table.Columns.Add("領域_" + domain + "_平均成績等第");
            }

            for (int i = 1; i <= 12; i++)
            {
                table.Columns.Add("領域_學習領域總成績_成績_" + i);
                table.Columns.Add("領域_學習領域總成績_等第_" + i);
            }
            table.Columns.Add("領域_學習領域總平均成績_成績");
            table.Columns.Add("領域_學習領域總平均成績_等第");
            #endregion

            #region 科目成績
            //OO領域 科目成績
            foreach (string domain in DomainList)
            {
                for (int a = 1; a <= 6; a++)
                {
                    table.Columns.Add(domain + "_科目名稱" + a);
                    for (int i = 1; i <= 12; i++)
                    {
                        table.Columns.Add(domain + "_科目" + a + "_權數" + i);
                        table.Columns.Add(domain + "_科目" + a + "_成績" + i);
                        table.Columns.Add(domain + "_科目" + a + "_等第" + i);
                        table.Columns.Add(domain + "_科目" + a + "_原始成績" + i);
                        table.Columns.Add(domain + "_科目" + a + "_原始等第" + i);
                    }
                    table.Columns.Add(domain + "_科目" + a + "_平均成績");
                    table.Columns.Add(domain + "_科目" + a + "_平均成績等第");
                }

                if (domain == "彈性課程")
                    for (int a = 7; a <= 18; a++)
                    {
                        table.Columns.Add(domain + "_科目名稱" + a);
                        for (int i = 1; i <= 12; i++)
                        {
                            table.Columns.Add("彈性課程_科目" + a + "_權數" + i);
                            table.Columns.Add("彈性課程_科目" + a + "_成績" + i);
                            table.Columns.Add("彈性課程_科目" + a + "_等第" + i);
                            table.Columns.Add("彈性課程_科目" + a + "_原始成績" + i);
                            table.Columns.Add("彈性課程_科目" + a + "_原始等第" + i);
                        }
                        table.Columns.Add("彈性課程_科目" + a + "_平均成績");
                        table.Columns.Add("彈性課程_科目" + a + "_平均成績等第");
                    }
            }
            #endregion

            #region 學務相關資料 
            //出缺勤紀錄
            for (int i = 1; i <= 12; i++)
            {
                table.Columns.Add("上課天數_" + i);
            }

            for (int i = 1; i <= 12; i++)
            {
                table.Columns.Add("服務學習時數_" + i);
                table.Columns.Add("大功_" + i);
                table.Columns.Add("小功_" + i);
                table.Columns.Add("嘉獎_" + i);
                table.Columns.Add("大過_" + i);
                table.Columns.Add("小過_" + i);
                table.Columns.Add("警告_" + i);
            }

            #region 前一版康橋使用的 (暫不使用)
            //table.Columns.Add("事假日數_1");
            //table.Columns.Add("事假日數_2");
            //table.Columns.Add("事假日數_3");
            //table.Columns.Add("事假日數_4");
            //table.Columns.Add("事假日數_5");
            //table.Columns.Add("事假日數_6");
            //table.Columns.Add("事假日數_7");
            //table.Columns.Add("事假日數_8");
            //table.Columns.Add("事假日數_9");
            //table.Columns.Add("事假日數_10");
            //table.Columns.Add("事假日數_11");
            //table.Columns.Add("事假日數_12");


            //table.Columns.Add("病假日數_1");
            //table.Columns.Add("病假日數_2");
            //table.Columns.Add("病假日數_3");
            //table.Columns.Add("病假日數_4");
            //table.Columns.Add("病假日數_5");
            //table.Columns.Add("病假日數_6");
            //table.Columns.Add("病假日數_7");
            //table.Columns.Add("病假日數_8");
            //table.Columns.Add("病假日數_9");
            //table.Columns.Add("病假日數_10");
            //table.Columns.Add("病假日數_11");
            //table.Columns.Add("病假日數_12");

            //table.Columns.Add("公假日數_1");
            //table.Columns.Add("公假日數_2");
            //table.Columns.Add("公假日數_3");
            //table.Columns.Add("公假日數_4");
            //table.Columns.Add("公假日數_5");
            //table.Columns.Add("公假日數_6");
            //table.Columns.Add("公假日數_7");
            //table.Columns.Add("公假日數_8");
            //table.Columns.Add("公假日數_9");
            //table.Columns.Add("公假日數_10");
            //table.Columns.Add("公假日數_11");
            //table.Columns.Add("公假日數_12");

            //table.Columns.Add("喪假日數_1");
            //table.Columns.Add("喪假日數_2");
            //table.Columns.Add("喪假日數_3");
            //table.Columns.Add("喪假日數_4");
            //table.Columns.Add("喪假日數_5");
            //table.Columns.Add("喪假日數_6");
            //table.Columns.Add("喪假日數_7");
            //table.Columns.Add("喪假日數_8");
            //table.Columns.Add("喪假日數_9");
            //table.Columns.Add("喪假日數_10");
            //table.Columns.Add("喪假日數_11");
            //table.Columns.Add("喪假日數_12");

            //table.Columns.Add("曠課日數_1");
            //table.Columns.Add("曠課日數_2");
            //table.Columns.Add("曠課日數_3");
            //table.Columns.Add("曠課日數_4");
            //table.Columns.Add("曠課日數_5");
            //table.Columns.Add("曠課日數_6");
            //table.Columns.Add("曠課日數_7");
            //table.Columns.Add("曠課日數_8");
            //table.Columns.Add("曠課日數_9");
            //table.Columns.Add("曠課日數_10");
            //table.Columns.Add("曠課日數_11");
            //table.Columns.Add("曠課日數_12");

            //table.Columns.Add("缺席總日數_1");
            //table.Columns.Add("缺席總日數_2");
            //table.Columns.Add("缺席總日數_3");
            //table.Columns.Add("缺席總日數_4");
            //table.Columns.Add("缺席總日數_5");
            //table.Columns.Add("缺席總日數_6");
            //table.Columns.Add("缺席總日數_7");
            //table.Columns.Add("缺席總日數_8");
            //table.Columns.Add("缺席總日數_9");
            //table.Columns.Add("缺席總日數_10");
            //table.Columns.Add("缺席總日數_11");
            //table.Columns.Add("缺席總日數_12");
            #endregion

            #endregion

            #region 畢業總成績 (暫不使用)
            ////畢業總成績
            //table.Columns.Add("畢業總成績_平均");
            //table.Columns.Add("畢業總成績_等第");
            //table.Columns.Add("准予畢業");
            //table.Columns.Add("發給修業證書");
            #endregion

            #region 日常生活表現及具體建議 (暫不使用)
            //日常生活表現及具體建議
            //table.Columns.Add("日常生活表現及具體建議_1");
            //table.Columns.Add("日常生活表現及具體建議_2");
            //table.Columns.Add("日常生活表現及具體建議_3");
            //table.Columns.Add("日常生活表現及具體建議_4");
            //table.Columns.Add("日常生活表現及具體建議_5");
            //table.Columns.Add("日常生活表現及具體建議_6");
            //table.Columns.Add("日常生活表現及具體建議_7");
            //table.Columns.Add("日常生活表現及具體建議_8");
            //table.Columns.Add("日常生活表現及具體建議_9");
            //table.Columns.Add("日常生活表現及具體建議_10");
            //table.Columns.Add("日常生活表現及具體建議_11");
            //table.Columns.Add("日常生活表現及具體建議_12");
            #endregion

            #region 校內外特殊表現 (暫不使用)
            ////校內外特殊表現
            //table.Columns.Add("校內外特殊表現_1");
            //table.Columns.Add("校內外特殊表現_2");
            //table.Columns.Add("校內外特殊表現_3");
            //table.Columns.Add("校內外特殊表現_4");
            //table.Columns.Add("校內外特殊表現_5");
            //table.Columns.Add("校內外特殊表現_6");
            //table.Columns.Add("校內外特殊表現_7");
            //table.Columns.Add("校內外特殊表現_8");
            //table.Columns.Add("校內外特殊表現_9");
            //table.Columns.Add("校內外特殊表現_10");
            //table.Columns.Add("校內外特殊表現_11");
            //table.Columns.Add("校內外特殊表現_12");
            #endregion

            #region 匯出日期
            table.Columns.Add("匯出日期");
            #endregion

            #endregion

            Aspose.Words.Document document = new Aspose.Words.Document();

            e_For_ConvertToPDF_Worker = e;

            #region 整理所有的獎懲
            List<string> disciplineType_list = new List<string>();

            disciplineType_list.Add("大功");
            disciplineType_list.Add("小功");
            disciplineType_list.Add("嘉獎");
            disciplineType_list.Add("大過");
            disciplineType_list.Add("小過");
            disciplineType_list.Add("警告");
            #endregion

            #region 整理所有的假別  (暫不使用)
            ////整理所有的假別
            //List<string> absenceType_list = new List<string>();

            //absenceType_list.Add("事假");
            //absenceType_list.Add("病假");
            //absenceType_list.Add("公假");
            //absenceType_list.Add("喪假");
            //absenceType_list.Add("曠課");
            //absenceType_list.Add("缺席總");
            #endregion

            #region 整理所有的領域_OO_成績
            //整理所有的領域_OO_成績
            List<string> domainScoreType_list = new List<string>();

            foreach (string domain in DomainList)
            {
                if (domain == "彈性課程")
                    continue;
                domainScoreType_list.Add("領域_" + domain + "_成績_");
                domainScoreType_list.Add("領域_" + domain + "_平均成績");
            }
            //domainScoreType_list.Add("領域_語文_成績_");
            //domainScoreType_list.Add("領域_數學_成績_");
            //domainScoreType_list.Add("領域_自然科學_成績_");
            //domainScoreType_list.Add("領域_自然與生活科技_成績_");
            //domainScoreType_list.Add("領域_藝術_成績_");
            //domainScoreType_list.Add("領域_藝術與人文_成績_");
            //domainScoreType_list.Add("領域_社會_成績_");
            //domainScoreType_list.Add("領域_健康與體育_成績_");
            //domainScoreType_list.Add("領域_綜合活動_成績_");
            domainScoreType_list.Add("領域_學習領域總成績_成績_");
            domainScoreType_list.Add("領域_學習領域總平均成績_成績");
            #endregion

            #region 整理所有的領域_OO_等第
            //整理所有的領域_OO_等第
            List<string> domainLevelType_list = new List<string>();
            foreach (string domain in DomainList)
            {
                if (domain == "彈性課程")
                    continue;
                domainLevelType_list.Add("領域_" + domain + "_等第_");
                domainLevelType_list.Add("領域_" + domain + "_平均成績等第");
            }
            //domainLevelType_list.Add("領域_語文_等第_");
            //domainLevelType_list.Add("領域_數學_等第_");
            //domainLevelType_list.Add("領域_自然科學_等第_");
            //domainLevelType_list.Add("領域_自然與生活科技_等第_");
            //domainLevelType_list.Add("領域_藝術_等第_");
            //domainLevelType_list.Add("領域_藝術與人文_等第_");
            //domainLevelType_list.Add("領域_社會_等第_");
            //domainLevelType_list.Add("領域_健康與體育_等第_");
            //domainLevelType_list.Add("領域_綜合活動_等第_");
            domainLevelType_list.Add("領域_學習領域總成績_等第_");
            domainLevelType_list.Add("領域_學習領域總平均成績_等第");
            #endregion

            #region 整理所有的領域_OO_權數
            //整理所有的領域_OO_等第
            List<string> domainCreditType_list = new List<string>();
            foreach (string domain in DomainList)
            {
                if (domain == "彈性課程")
                    continue;
                domainCreditType_list.Add("領域_" + domain + "_權數_");
            }


            #endregion

            #region 整理所有的科目成績
            List<string> subjectScoreType_list = new List<string>();
            foreach (string domain in DomainList)
            {
                for (int a = 1; a <= 12; a++)
                {
                    subjectScoreType_list.Add(domain + "_科目" + a + "_成績");
                    subjectScoreType_list.Add(domain + "_科目" + a + "_原始成績");
                    subjectScoreType_list.Add(domain + "_科目" + a + "_平均成績");
                }
            }
            #endregion

            #region 整理所有的科目等第
            List<string> subjectLevelType_list = new List<string>();
            foreach (string domain in DomainList)
                for (int a = 1; a <= 12; a++)
                {
                    subjectLevelType_list.Add(domain + "_科目" + a + "_等第");
                    subjectLevelType_list.Add(domain + "_科目" + a + "_原始等第");
                    subjectLevelType_list.Add(domain + "_科目" + a + "_平均成績等第");
                }
            #endregion

            #region 整理所有的科目權數
            List<string> subjectCredit_list = new List<string>();
            foreach (string domain in DomainList)
                for (int a = 1; a <= 12; a++)
                {
                    subjectCredit_list.Add(domain + "_科目" + a + "_權數");
                }
            #endregion

            // 領域分數、等第、權數 的對照
            Dictionary<string, decimal?> domainScore_dict = new Dictionary<string, decimal?>();
            Dictionary<string, string> domainLevel_dict = new Dictionary<string, string>();
            Dictionary<string, decimal?> domainCredit_dict = new Dictionary<string, decimal?>();

            // 科目分數、等第、權數 的對照
            Dictionary<string, decimal?> subjectScore_dict = new Dictionary<string, decimal?>();
            Dictionary<string, string> subjectLevel_dict = new Dictionary<string, string>();
            Dictionary<string, decimal?> subjectCredit_dict = new Dictionary<string, decimal?>();

            // 缺曠節次 、日數 的對照
            //Dictionary<string, decimal> arStatistic_dict = new Dictionary<string, decimal>();
            //Dictionary<string, decimal> arStatistic_dict_days = new Dictionary<string, decimal>();

            //六學期 缺曠節次、節數對照
            Dictionary<string, int> absence_dic = new Dictionary<string, int>();

            //六學期 服務學習時數對照
            Dictionary<string, string> serviceLearning_dic = new Dictionary<string, string>();

            //六學期 獎懲對照
            Dictionary<string, string> discipline_dic = new Dictionary<string, string>();

            //文字評量(日常生活表現及具體建議、校內外特殊表現)的對照
            //Dictionary<string, string> textScore_dict = new Dictionary<string, string>();

            int student_counter = 1;

            foreach (string stuID in StudentIDs)
            {
                //把每一筆資料的字典都清乾淨，避免資料汙染
                //arStatistic_dict.Clear();
                //arStatistic_dict_days.Clear();
                domainScore_dict.Clear();
                domainLevel_dict.Clear();
                subjectScore_dict.Clear();
                subjectLevel_dict.Clear();
                subjectCredit_dict.Clear();
                //textScore_dict.Clear();
                absence_dic.Clear();
                serviceLearning_dic.Clear();
                discipline_dic.Clear();
                domainCredit_dict.Clear();
                // 建立缺曠 對照字典
                //foreach (string ab in absenceType_list)
                //{
                //    for (int i = 1; i <= 12; i++)
                //    {
                //        arStatistic_dict.Add(ab + "日數_" + i, 0);
                //    }
                //}
                // 建立 六學期 缺曠節次字典
                foreach (string p in paList)
                {
                    absence_dic.Add(p, 0);
                }

                // 建立 六學期 獎懲字典
                foreach (string d in disciplineType_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        discipline_dic.Add(d + "_" + i, "0");
                    }
                }

                for (int i = 1; i <= 12; i++)
                {
                    serviceLearning_dic.Add("服務學習時數_" + i, "0");
                }
                // 建立領域成績 對照字典
                foreach (string dst in domainScoreType_list)
                {
                    if (dst.Contains("平均成績"))
                    {
                        domainScore_dict.Add(dst, null);
                        continue;
                    }

                    for (int i = 1; i <= 12; i++)
                    {
                        domainScore_dict.Add(dst + i, null);
                    }
                }

                // 建立領域等第 對照字典
                foreach (string dlt in domainLevelType_list)
                {
                    if (dlt.Contains("平均成績"))
                    {
                        domainLevel_dict.Add(dlt, null);
                        continue;
                    }

                    for (int i = 1; i <= 12; i++)
                    {
                        domainLevel_dict.Add(dlt + i, null);
                    }
                }

                // 建立領域權數 對照字典
                foreach (string dlt in domainCreditType_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        domainCredit_dict.Add(dlt + i, null);
                    }
                }

                // 建立科目成績 對照字典
                foreach (string sst in subjectScoreType_list)
                {
                    if (sst.Contains("平均成績"))
                    {
                        subjectScore_dict.Add(sst, null);
                        continue;
                    }

                    for (int i = 1; i <= 12; i++)
                    {
                        subjectScore_dict.Add(sst + i, null);
                    }
                }

                // 建立科目等第 對照字典
                foreach (string slt in subjectLevelType_list)
                {
                    if (slt.Contains("平均成績"))
                    {
                        subjectLevel_dict.Add(slt, null);
                        continue;
                    }

                    for (int i = 1; i <= 12; i++)
                    {
                        subjectLevel_dict.Add(slt + i, null);
                    }
                }

                // 建立科目等第 對照字典
                foreach (string slt in subjectCredit_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        subjectCredit_dict.Add(slt + i, null);
                    }
                }

                //// 建立文字評量 對照字典
                //for (int i = 1; i <= 6; i++)
                //{
                //    textScore_dict.Add("日常生活表現及具體建議_" + i, null);
                //    textScore_dict.Add("校內外特殊表現_" + i, null);
                //}

                // 存放 各年級與 學年的對照變數
                int schoolyear_grade1 = 0;
                int schoolyear_grade2 = 0;
                int schoolyear_grade3 = 0;
                int schoolyear_grade4 = 0;
                int schoolyear_grade5 = 0;
                int schoolyear_grade6 = 0;


                DataRow row = table.NewRow();

                //row["學校名稱"] = K12.Data.School.ChineseName;
                row["學校名稱"] = K12.Data.School.EnglishName;
                //row["學校地址"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("Address").InnerText;
                row["學校地址"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("EnglishAddress").InnerText;
                row["學校電話"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("Telephone").InnerText;
                //row["校長"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText;
                row["校長"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorEnglishName").InnerText;
                row["教務主任"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("EduDirectorName").InnerText;

                #region 學生照片，若沒有畢業照，則印出入學照
                string graduatePhoto = K12.Data.Photo.SelectGraduatePhoto(stuID);
                string freshmanPhoto = K12.Data.Photo.SelectFreshmanPhoto(stuID);
                if (!_PhotoPDict.ContainsKey(stuID))
                {
                    if (string.IsNullOrEmpty(graduatePhoto))
                        graduatePhoto = freshmanPhoto;
                    _PhotoPDict.Add(stuID, graduatePhoto);
                }
                if (_PhotoPDict.ContainsKey(stuID))
                    row["照片"] = _PhotoPDict[stuID].FromBase64StringToByte();
                #endregion
                QueryHelper queryHelper = new QueryHelper();
                //學生基本資料
                if (sr_dict.ContainsKey(stuID))
                {

                    DateTime birthday = new DateTime();

                    row["學生姓名"] = sr_dict[stuID].Name;
                    row["學生英文姓名"] = sr_dict[stuID].EnglishName;
                    row["學生班級"] = sr_dict[stuID].Class != null ? sr_dict[stuID].Class.Name : "";
                    row["學生座號"] = sr_dict[stuID].SeatNo;
                    //// 英文版格式 Male Female
                    row["學生性別"] = sr_dict[stuID].Gender == "男" ? "Male" : "Female";

                    birthday = (DateTime)sr_dict[stuID].Birthday;
                    // 轉換出生時間 成 2005/09/06 的格式
                    // 英文版格式 August 29, 2016
                    // 如果不加new CultureInfo("en-US") 資訊 會被轉成中文 "八月"
                    row["出生日期"] = birthday.ToString("MMMM dd,yyyy", new CultureInfo("en-US"));

                    row["入學年月"] = "";
                    row["學生身分證字號"] = sr_dict[stuID].IDNumber;
                    row["學號"] = sr_dict[stuID].StudentNumber;

                    #region 國籍護照資料
                    string strSQL = "select nationality1, passport_name1, nat1.eng_name as nat_eng1, nationality2, passport_name2, nat2.eng_name as nat_eng2, nationality3, passport_name3, nat3.eng_name as nat_eng3 from student_info_ext  as stud_info left outer join $ischool.mapping.nationality as nat1 on nat1.name = stud_info.nationality1 left outer join $ischool.mapping.nationality as nat2 on nat2.name = stud_info.nationality2 left outer join $ischool.mapping.nationality as nat3 on nat3.name = stud_info.nationality3 WHERE ref_student_id=" + stuID;
                    DataTable student_info_ext = queryHelper.Select(strSQL);
                    if (student_info_ext.Rows.Count > 0)
                    {
                        row["國籍一"] = student_info_ext.Rows[0]["nationality1"];
                        row["國籍一護照名"] = student_info_ext.Rows[0]["passport_name1"];
                        row["國籍二"] = student_info_ext.Rows[0]["nationality2"];
                        row["國籍二護照名"] = student_info_ext.Rows[0]["passport_name2"];
                        row["國籍一英文"] = student_info_ext.Rows[0]["nat_eng1"];
                        row["國籍二英文"] = student_info_ext.Rows[0]["nat_eng2"];
                    }
                    else
                    {
                        row["國籍一"] = "";
                        row["國籍一護照名"] = "";
                        row["國籍二"] = "";
                        row["國籍二護照名"] = "";
                        row["國籍一英文"] = "";
                        row["國籍二英文"] = "";
                    }

                    #endregion
                    PrintStudents.Add(sr_dict[stuID]);
                }

                //學期歷程 //相容7、8、9
                if (shr_dict.ContainsKey(stuID))
                {
                    foreach (var item in shr_dict[stuID].SemesterHistoryItems)
                    {
                        if (item.GradeYear == 1 || item.GradeYear == 7)
                        {
                            row["學年度1"] = item.SchoolYear;
                            row["年級1_班級"] = item.ClassName;
                            row["年級1_座號"] = item.SeatNo;

                            //為學生的年級與學年配對
                            schoolyear_grade1 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["上課天數_1"] = item.SchoolDayCount;
                            }
                            else
                            {
                                row["上課天數_2"] = item.SchoolDayCount;

                            }

                        }
                        if (item.GradeYear == 2 || item.GradeYear == 8)
                        {
                            row["學年度2"] = item.SchoolYear;
                            row["年級2_班級"] = item.ClassName;
                            row["年級2_座號"] = item.SeatNo;

                            //為學生的年級與學年配對
                            schoolyear_grade2 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["上課天數_3"] = item.SchoolDayCount;
                            }
                            else
                            {
                                row["上課天數_4"] = item.SchoolDayCount;

                            }
                        }
                        if (item.GradeYear == 3 || item.GradeYear == 9)
                        {
                            row["學年度3"] = item.SchoolYear;
                            row["年級3_班級"] = item.ClassName;
                            row["年級3_座號"] = item.SeatNo;

                            //為學生的年級與學年配對
                            schoolyear_grade3 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["上課天數_5"] = item.SchoolDayCount;
                            }
                            else
                            {
                                row["上課天數_6"] = item.SchoolDayCount;

                            }
                        }

                        if (item.GradeYear == 4)
                        {
                            row["學年度4"] = item.SchoolYear;
                            row["年級4_班級"] = item.ClassName;
                            row["年級4_座號"] = item.SeatNo;

                            //為學生的年級與學年配對
                            schoolyear_grade4 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["上課天數_7"] = item.SchoolDayCount;
                            }
                            else
                            {
                                row["上課天數_8"] = item.SchoolDayCount;

                            }
                        }

                        if (item.GradeYear == 5)
                        {
                            row["學年度5"] = item.SchoolYear;
                            row["年級5_班級"] = item.ClassName;
                            row["年級5_座號"] = item.SeatNo;

                            //為學生的年級與學年配對
                            schoolyear_grade5 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["上課天數_9"] = item.SchoolDayCount;
                            }
                            else
                            {
                                row["上課天數_10"] = item.SchoolDayCount;

                            }
                        }

                        if (item.GradeYear == 6)
                        {
                            row["學年度6"] = item.SchoolYear;
                            row["年級6_班級"] = item.ClassName;
                            row["年級6_座號"] = item.SeatNo;

                            //為學生的年級與學年配對
                            schoolyear_grade6 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["上課天數_11"] = item.SchoolDayCount;
                            }
                            else
                            {
                                row["上課天數_12"] = item.SchoolDayCount;

                            }
                        }
                    }

                }

                //學年度與年級的對照字典
                Dictionary<int, int> schoolyear_grade_dict = new Dictionary<int, int>();

                schoolyear_grade_dict.Add(1, schoolyear_grade1);
                schoolyear_grade_dict.Add(2, schoolyear_grade2);
                schoolyear_grade_dict.Add(3, schoolyear_grade3);
                schoolyear_grade_dict.Add(4, schoolyear_grade4);
                schoolyear_grade_dict.Add(5, schoolyear_grade5);
                schoolyear_grade_dict.Add(6, schoolyear_grade6);

                //出缺勤
                if (ar_dict.ContainsKey(stuID))
                {
                    for (int grade = 1; grade <= 6; grade++)
                    {
                        foreach (var ar in ar_dict[stuID])
                        {
                            if (ar.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                if (ar.Semester == 1)
                                {
                                    foreach (var detail in ar.PeriodDetail)
                                    {
                                        //"一般_產假1" 
                                        if (PeriodMappingDic.ContainsKey(detail.Period))
                                        {
                                            string key = PeriodMappingDic[detail.Period] + "_" + detail.AbsenceType + (grade * 2 - 1);
                                            if (absence_dic.ContainsKey(key))
                                            {
                                                absence_dic[key] += 1;

                                            }
                                        }
                                        //if (arStatistic_dict.ContainsKey(detail.AbsenceType + "日數_" + (grade * 2 - 1)))
                                        //{
                                        //    //加一節，整學期節次與日數的關係，再最後再結算
                                        //    arStatistic_dict[detail.AbsenceType + "日數_" + (grade * 2 - 1)] += 1;

                                        //    // 不管是啥缺席，缺席總日數都加一節
                                        //    arStatistic_dict["缺席總日數_" + (grade * 2 - 1)] += 1;
                                        //}
                                    }
                                }
                                else
                                {
                                    foreach (var detail in ar.PeriodDetail)
                                    {
                                        //"一般_產假1" 
                                        if (PeriodMappingDic.ContainsKey(detail.Period))
                                        {
                                            string key = PeriodMappingDic[detail.Period] + "_" + detail.AbsenceType + (grade * 2);
                                            if (absence_dic.ContainsKey(key))
                                            {
                                                absence_dic[key] += 1;

                                            }
                                        }
                                        //if (arStatistic_dict.ContainsKey(detail.AbsenceType + "日數_" + grade * 2))
                                        //{
                                        //    //加一節，整學期節次與日數的關係，再最後再結算
                                        //    arStatistic_dict[detail.AbsenceType + "日數_" + grade * 2] += 1;

                                        //    // 不管是啥缺席，缺席總日數都加一節
                                        //    arStatistic_dict["缺席總日數_" + (grade * 2)] += 1;
                                        //}
                                    }
                                }
                            }
                        }

                        if (Util._SLRDict.ContainsKey(stuID))
                            foreach (var item in Util._SLRDict[stuID])
                            {
                                // 第一學期
                                if (item.Key == schoolyear_grade_dict[grade] + "_1")
                                {
                                    //decimal hours = 0;
                                    //if (decimal.TryParse(item.Value, out hours))
                                    serviceLearning_dic["服務學習時數_" + (grade * 2 - 1)] = item.Value;
                                }
                                // 第二學期
                                if (item.Key == schoolyear_grade_dict[grade] + "_2")
                                {
                                    //decimal hours = 0;
                                    //if (decimal.TryParse(item.Value, out hours))
                                    serviceLearning_dic["服務學習時數_" + (grade * 2)] = item.Value;
                                }
                            }

                        if (Util._DisciplineDict.ContainsKey(stuID))
                            foreach (var item in Util._DisciplineDict[stuID])
                            {
                                // 第一學期
                                if (item.Key == schoolyear_grade_dict[grade] + "_1")
                                {
                                    foreach (var dis in item.Value)
                                    {
                                        discipline_dic[dis.Key + "_" + (grade * 2 - 1)] = dis.Value;
                                    }

                                }
                                // 第二學期
                                if (item.Key == schoolyear_grade_dict[grade] + "_2")
                                {
                                    foreach (var dis in item.Value)
                                    {
                                        discipline_dic[dis.Key + "_" + (grade * 2)] = dis.Value;
                                    }
                                }
                            }
                    }

                    //foreach (string key in arStatistic_dict.Keys)
                    //{
                    //    arStatistic_dict_days.Add(key, arStatistic_dict[key]);
                    //}

                    ////真正的填值，填日數，所以要做節次轉換
                    //foreach (string key in arStatistic_dict_days.Keys)
                    //{
                    //    //康橋一日有九節，多一節缺曠 = 多1/9 日缺曠，先暫時寫死九節設定，日後要去學務作業每日節次抓取
                    //    row[key] = Math.Round(arStatistic_dict_days[key] / 9, 2);
                    //}

                    foreach (string key in absence_dic.Keys)
                    {
                        row[key] = absence_dic[key];
                    }
                    foreach (string key in serviceLearning_dic.Keys)
                    {
                        row[key] = serviceLearning_dic[key];
                    }
                    foreach (string key in discipline_dic.Keys)
                    {
                        row[key] = discipline_dic[key];
                    }
                }


                //一般科目 科目名稱與科目編號對照表
                Dictionary<string, Dictionary<string, int>> SubjectCourseDict = new Dictionary<string, Dictionary<string, int>>();
                //{
                //    { "語文", new Dictionary<string, int>() }
                //    , { "國語文", new Dictionary<string, int>() }
                //    , { "英語文", new Dictionary<string, int>() }
                //    , { "數學", new Dictionary<string, int>() }
                //    , { "社會", new Dictionary<string, int>() }
                //    , { "自然科學", new Dictionary<string, int>() }
                //    , { "自然與生活科技", new Dictionary<string, int>() }
                //    , { "藝術", new Dictionary<string, int>() }
                //    , { "藝術與人文", new Dictionary<string, int>() }
                //    , { "健康與體育", new Dictionary<string, int>() }
                //    , { "綜合活動", new Dictionary<string, int>() }
                //    , { "科技", new Dictionary<string, int>() }
                //    , { "實用語文", new Dictionary<string, int>() }
                //    , { "實用數學", new Dictionary<string, int>() }
                //    , { "社會適應", new Dictionary<string, int>() }
                //    , { "生活教育", new Dictionary<string, int>() }
                //    , { "休閒教育", new Dictionary<string, int>() }
                //    , { "職業教育", new Dictionary<string, int>() }
                //    , { "體育專業", new Dictionary<string, int>() }
                //    , { "藝術才能專長", new Dictionary<string, int>() }
                //};

                foreach (string domain in DomainList)
                {
                    if (!SubjectCourseDict.ContainsKey(domain))
                        SubjectCourseDict.Add(domain, new Dictionary<string, int>());
                }
                // 彈性課程 科目名稱 與彈性課程編號的對照
                Dictionary<string, int> AlternativeCourseDict = new Dictionary<string, int>();


                // 學期成績(包含領域、科目)
                if (jssr_dict.ContainsKey(stuID))
                {
                    // 任一領域的科目數量是否超過
                    bool isExceed = false;

                    for (int grade = 1; grade <= 6; grade++)
                    {
                        foreach (JHSemesterScoreRecord jssr in jssr_dict[stuID])
                        {
                            if (jssr.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                foreach (var subjectscore in jssr.Subjects)
                                {
                                    // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                    if (subjectscore.Value.Domain != "彈性課程" && subjectscore.Value.Domain != "彈性學習" && subjectscore.Value.Domain != "")
                                    {
                                        if (SubjectCourseDict.ContainsKey(subjectscore.Value.Domain))
                                        {
                                            int subjectCourseCount = SubjectCourseDict[subjectscore.Value.Domain].Count;

                                            if (SubjectCourseDict[subjectscore.Value.Domain].ContainsKey(subjectscore.Value.Subject))
                                            {
                                                continue;
                                            }

                                            subjectCourseCount++;

                                            // 目前僅支援 一個學生六學年之中同一領域僅能有 6個科目
                                            if (subjectCourseCount > 6)
                                            {
                                                isExceed = true;
                                                continue;
                                            }

                                            row[subjectscore.Value.Domain + "_科目名稱" + subjectCourseCount] = _SubjDomainEngNameMapping.GetSubjectEngName(subjectscore.Value.Subject);

                                            SubjectCourseDict[subjectscore.Value.Domain].Add(subjectscore.Value.Subject, subjectCourseCount);
                                        }
                                    }
                                }

                            }
                        }
                    }

                    if (isExceed)
                    {
                        MessageBox.Show("科目數超過報表變數可支援數量，超過的將不會顯示於在校成績證明書中");
                    }
                }

                // 先統計 該學生 在全學年間 有的 彈性課程科目
                if (jssr_dict.ContainsKey(stuID))
                {
                    // 彈性課程記數
                    int AlternativeCourse = 0;

                    for (int grade = 1; grade <= 6; grade++)
                    {
                        foreach (JHSemesterScoreRecord jssr in jssr_dict[stuID])
                        {
                            if (jssr.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                foreach (var subjectscore in jssr.Subjects)
                                {
                                    // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                    if (subjectscore.Value.Domain == "彈性課程" || subjectscore.Value.Domain == "彈性學習" || subjectscore.Value.Domain == "")
                                    {
                                        // 對照科目名稱如果已經有，跳過
                                        if (AlternativeCourseDict.ContainsKey(subjectscore.Value.Subject))
                                        {
                                            continue;
                                        }

                                        AlternativeCourse++;

                                        // 目前僅先支援 一個學生在六年之中有 18個 彈性課程
                                        if (AlternativeCourse > 18)
                                        {
                                            MessageBox.Show("彈性科目數超過可支援數量，超過的將不會顯示在在校成績證明書中");
                                            break;
                                        }

                                        row["彈性課程_科目名稱" + AlternativeCourse] = _SubjDomainEngNameMapping.GetSubjectEngName(subjectscore.Value.Subject);

                                        AlternativeCourseDict.Add(subjectscore.Value.Subject, AlternativeCourse);
                                    }
                                }
                            }
                        }
                    }
                }

                Dictionary<string, List<decimal>> domainScoreDic = new Dictionary<string, List<decimal>>();
                Dictionary<string, List<decimal>> subjectScoreDic = new Dictionary<string, List<decimal>>();
                Dictionary<string, List<decimal>> learmingDomainScoreDic = new Dictionary<string, List<decimal>>();

                if (jssr_dict.ContainsKey(stuID))
                {
                    for (int grade = 1; grade <= 6; grade++)
                    {
                        foreach (JHSemesterScoreRecord jssr in jssr_dict[stuID])
                        {
                            if (jssr.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                if (jssr.Semester == 1)
                                {
                                    //領域
                                    foreach (var domainscore in jssr.Domains)
                                    {
                                        // 紀錄成績以計算平均成績
                                        if (!domainScoreDic.ContainsKey("領域_" + domainscore.Value.Domain))
                                        {
                                            domainScoreDic.Add("領域_" + domainscore.Value.Domain, new List<decimal>());
                                        }
                                        domainScoreDic["領域_" + domainscore.Value.Domain].Add(domainscore.Value.Score.Value);

                                        //紀錄成績
                                        if (domainScore_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2 - 1)))
                                        {
                                            domainScore_dict["領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2 - 1)] = domainscore.Value.Score;
                                        }

                                        //換算等第
                                        if (domainLevel_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2 - 1)))
                                        {
                                            //domainLevel_dict["領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2 - 1)] = ScoreTolevel(domainscore.Value.Score);
                                            domainLevel_dict["領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2 - 1)] = _ScoreMappingConfig.ParseScoreEngName(domainscore.Value.Score);
                                        }

                                        //紀錄權數
                                        if (domainCredit_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_權數_" + (grade * 2 - 1)))
                                        {
                                            domainCredit_dict["領域_" + domainscore.Value.Domain + "_權數_" + (grade * 2 - 1)] = domainscore.Value.Credit;
                                        }
                                    }

                                    //科目
                                    foreach (var subjectscore in jssr.Subjects)
                                    {
                                        // 彈性課程記數
                                        int AlternativeCourse = 0;
                                        int SubjectCourseNum = 0;

                                        // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                        if (subjectscore.Value.Domain == "彈性課程" || subjectscore.Value.Domain == "彈性學習" || subjectscore.Value.Domain == "")
                                        {
                                            if (AlternativeCourseDict.ContainsKey(subjectscore.Value.Subject))
                                            {
                                                AlternativeCourse = AlternativeCourseDict[subjectscore.Value.Subject];

                                                // 紀錄成績以計算平均成績
                                                if (!subjectScoreDic.ContainsKey("彈性課程_科目" + AlternativeCourse))
                                                {
                                                    subjectScoreDic.Add("彈性課程_科目" + AlternativeCourse, new List<decimal>());
                                                }
                                                subjectScoreDic["彈性課程_科目" + AlternativeCourse].Add(subjectscore.Value.Score.Value);


                                                //紀錄成績//彈性課程_科目4_成績3
                                                if (subjectScore_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_成績" + (grade * 2 - 1)))
                                                {
                                                    //{[語文_科目1_成績3, null]}
                                                    subjectScore_dict["彈性課程_科目" + AlternativeCourse + "_成績" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                }

                                                //紀錄原始成績
                                                if (subjectScore_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_原始成績" + (grade * 2 - 1)))
                                                {
                                                    subjectScore_dict["彈性課程_科目" + AlternativeCourse + "_原始成績" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                }

                                                //紀錄等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_等第" + (grade * 2 - 1)))
                                                {
                                                    subjectLevel_dict["彈性課程_科目" + AlternativeCourse + "_等第" + (grade * 2 - 1)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                }

                                                //紀錄原始等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_原始等第" + (grade * 2 - 1)))
                                                {
                                                    subjectLevel_dict["彈性課程_科目" + AlternativeCourse + "_原始等第" + (grade * 2 - 1)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                }

                                                //紀錄權數
                                                if (subjectCredit_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_權數" + (grade * 2 - 1)))
                                                {
                                                    subjectCredit_dict["彈性課程_科目" + AlternativeCourse + "_權數" + (grade * 2 - 1)] = subjectscore.Value.Credit;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (SubjectCourseDict.ContainsKey(subjectscore.Value.Domain))
                                            {
                                                if (SubjectCourseDict[subjectscore.Value.Domain].ContainsKey(subjectscore.Value.Subject))
                                                {
                                                    SubjectCourseNum = SubjectCourseDict[subjectscore.Value.Domain][subjectscore.Value.Subject];

                                                    // 紀錄成績以計算平均成績
                                                    if (!subjectScoreDic.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum))
                                                    {
                                                        subjectScoreDic.Add(subjectscore.Value.Domain + "_科目" + SubjectCourseNum, new List<decimal>());
                                                    }
                                                    subjectScoreDic[subjectscore.Value.Domain + "_科目" + SubjectCourseNum].Add(subjectscore.Value.Score.Value);


                                                    //紀錄成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_成績" + (grade * 2 - 1)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_成績" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                    }

                                                    //紀錄原始成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始成績" + (grade * 2 - 1)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始成績" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                    }

                                                    //換算等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_等第" + (grade * 2 - 1)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_等第" + (grade * 2 - 1)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                    }

                                                    //換算原始等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始等第" + (grade * 2 - 1)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始等第" + (grade * 2 - 1)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                    }

                                                    //紀錄權數
                                                    if (subjectCredit_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_權數" + (grade * 2 - 1)))
                                                    {
                                                        subjectCredit_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_權數" + (grade * 2 - 1)] = subjectscore.Value.Credit;
                                                    }

                                                }
                                            }

                                        }

                                    }

                                    //學期學習領域(七大)成績(不包括彈性課程成績)
                                    // 紀錄成績以計算平均成績
                                    if (!learmingDomainScoreDic.ContainsKey("領域_學習領域總平均成績"))
                                    {
                                        learmingDomainScoreDic.Add("領域_學習領域總平均成績", new List<decimal>());
                                    }
                                    learmingDomainScoreDic["領域_學習領域總平均成績"].Add(jssr.LearnDomainScore.Value);

                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_學習領域總成績_成績_" + (grade * 2 - 1)))
                                    {
                                        domainScore_dict["領域_學習領域總成績_成績_" + (grade * 2 - 1)] = jssr.LearnDomainScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_學習領域總成績_等第_" + (grade * 2 - 1)))
                                    {
                                        domainLevel_dict["領域_學習領域總成績_等第_" + (grade * 2 - 1)] = _ScoreMappingConfig.ParseScoreEngName(jssr.LearnDomainScore);
                                    }

                                    //課程學習成績(包括彈性課程成績)
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_課程學習成績_成績_" + (grade * 2 - 1)))
                                    {
                                        domainScore_dict["領域_課程學習成績_成績_" + (grade * 2 - 1)] = jssr.CourseLearnScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_課程學習成績_等第_" + (grade * 2 - 1)))
                                    {
                                        domainLevel_dict["領域_課程學習成績_等第_" + (grade * 2 - 1)] = _ScoreMappingConfig.ParseScoreEngName(jssr.CourseLearnScore);
                                    }


                                }
                                else
                                {
                                    //領域
                                    foreach (var domainscore in jssr.Domains)
                                    {
                                        // 紀錄成績以計算平均成績
                                        if (!domainScoreDic.ContainsKey("領域_" + domainscore.Value.Domain))
                                        {
                                            domainScoreDic.Add("領域_" + domainscore.Value.Domain, new List<decimal>());
                                        }
                                        domainScoreDic["領域_" + domainscore.Value.Domain].Add(domainscore.Value.Score.Value);

                                        //紀錄成績
                                        if (domainScore_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2)))
                                        {
                                            domainScore_dict["領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2)] = domainscore.Value.Score;
                                        }

                                        //換算等第
                                        if (domainLevel_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2)))
                                        {
                                            domainLevel_dict["領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2)] = _ScoreMappingConfig.ParseScoreEngName(domainscore.Value.Score);
                                        }

                                        if (domainCredit_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_權數_" + (grade * 2)))
                                        {
                                            domainCredit_dict["領域_" + domainscore.Value.Domain + "_權數_" + (grade * 2)] = domainscore.Value.Credit;
                                        }
                                    }

                                    //科目
                                    foreach (var subjectscore in jssr.Subjects)
                                    {
                                        // 彈性課程記數
                                        int AlternativeCourse = 0;
                                        int SubjectCourseNum = 0;

                                        // 領域為彈性課程 、或是沒有領域的科目成績 算到彈性課程科目處理
                                        if (subjectscore.Value.Domain == "彈性課程" || subjectscore.Value.Domain == "彈性學習" || subjectscore.Value.Domain == "")
                                        {
                                            if (AlternativeCourseDict.ContainsKey(subjectscore.Value.Subject))
                                            {
                                                AlternativeCourse = AlternativeCourseDict[subjectscore.Value.Subject];

                                                // 紀錄成績以計算平均成績
                                                if (!subjectScoreDic.ContainsKey("彈性課程_科目" + AlternativeCourse))
                                                {
                                                    subjectScoreDic.Add("彈性課程_科目" + AlternativeCourse, new List<decimal>());
                                                }
                                                subjectScoreDic["彈性課程_科目" + AlternativeCourse].Add(subjectscore.Value.Score.Value);

                                                //紀錄成績
                                                if (subjectScore_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_成績" + (grade * 2)))
                                                {
                                                    subjectScore_dict["彈性課程_科目" + AlternativeCourse + "_成績" + (grade * 2)] = subjectscore.Value.Score;
                                                }

                                                //紀錄原始成績
                                                if (subjectScore_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_原始成績" + (grade * 2)))
                                                {
                                                    subjectScore_dict["彈性課程_科目" + AlternativeCourse + "_原始成績" + (grade * 2)] = subjectscore.Value.Score;
                                                }

                                                //紀錄等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_等第" + (grade * 2)))
                                                {
                                                    subjectLevel_dict["彈性課程_科目" + AlternativeCourse + "_等第" + (grade * 2)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                }

                                                //紀錄原始等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_原始等第" + (grade * 2)))
                                                {
                                                    subjectLevel_dict["彈性課程_科目" + AlternativeCourse + "_原始等第" + (grade * 2)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                }

                                                //紀錄權數
                                                if (subjectCredit_dict.ContainsKey("彈性課程_科目" + AlternativeCourse + "_權數" + (grade * 2)))
                                                {
                                                    subjectCredit_dict["彈性課程_科目" + AlternativeCourse + "_權數" + (grade * 2)] = subjectscore.Value.Credit;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (SubjectCourseDict.ContainsKey(subjectscore.Value.Domain))
                                            {
                                                if (SubjectCourseDict[subjectscore.Value.Domain].ContainsKey(subjectscore.Value.Subject))
                                                {
                                                    SubjectCourseNum = SubjectCourseDict[subjectscore.Value.Domain][subjectscore.Value.Subject];

                                                    // 紀錄成績以計算平均成績
                                                    if (!subjectScoreDic.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum))
                                                    {
                                                        subjectScoreDic.Add(subjectscore.Value.Domain + "_科目" + SubjectCourseNum, new List<decimal>());
                                                    }
                                                    subjectScoreDic[subjectscore.Value.Domain + "_科目" + SubjectCourseNum].Add(subjectscore.Value.Score.Value);

                                                    //紀錄成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_成績" + (grade * 2)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_成績" + (grade * 2)] = subjectscore.Value.Score;
                                                    }

                                                    //紀錄原始成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始成績" + (grade * 2)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始成績" + (grade * 2)] = subjectscore.Value.Score;
                                                    }

                                                    //換算等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_等第" + (grade * 2)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_等第" + (grade * 2)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                    }

                                                    //換算原始等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始等第" + (grade * 2)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_原始等第" + (grade * 2)] = _ScoreMappingConfig.ParseScoreEngName(subjectscore.Value.Score);
                                                    }

                                                    //紀錄權數
                                                    if (subjectCredit_dict.ContainsKey(subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_權數" + (grade * 2)))
                                                    {
                                                        subjectCredit_dict[subjectscore.Value.Domain + "_科目" + SubjectCourseNum + "_權數" + (grade * 2)] = subjectscore.Value.Credit;
                                                    }
                                                }
                                            }
                                        }

                                    }

                                    //學期學習領域(七大)成績
                                    // 紀錄成績以計算平均成績
                                    if (!learmingDomainScoreDic.ContainsKey("領域_學習領域總平均成績"))
                                    {
                                        learmingDomainScoreDic.Add("領域_學習領域總平均成績", new List<decimal>());
                                    }
                                    learmingDomainScoreDic["領域_學習領域總平均成績"].Add(jssr.LearnDomainScore.Value);

                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_學習領域總成績_成績_" + (grade * 2)))
                                    {
                                        domainScore_dict["領域_學習領域總成績_成績_" + (grade * 2)] = jssr.LearnDomainScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_學習領域總成績_等第_" + (grade * 2)))
                                    {
                                        domainLevel_dict["領域_學習領域總成績_等第_" + (grade * 2)] = _ScoreMappingConfig.ParseScoreEngName(jssr.LearnDomainScore);
                                    }

                                    //課程學習成績(包括彈性課程成績)
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_課程學習成績_成績_" + (grade * 2)))
                                    {
                                        domainScore_dict["領域_課程學習成績_成績_" + (grade * 2)] = jssr.CourseLearnScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_課程學習成績_等第_" + (grade * 2)))
                                    {
                                        domainLevel_dict["領域_課程學習成績_等第_" + (grade * 2)] = _ScoreMappingConfig.ParseScoreEngName(jssr.CourseLearnScore);
                                    }


                                }
                            }
                        }

                    }

                    foreach (var domainScore in domainScoreDic)
                    {
                        decimal avgScore = Math.Round((domainScore.Value.Sum(x => x) / domainScore.Value.Count), 2);
                        domainScore_dict[domainScore.Key + "_平均成績"] = avgScore;
                        domainLevel_dict[domainScore.Key + "_平均成績等第"] = _ScoreMappingConfig.ParseScoreName(avgScore);
                    }

                    foreach (var subjetScore in subjectScoreDic)
                    {
                        decimal avgScore = Math.Round((subjetScore.Value.Sum(x => x) / subjetScore.Value.Count), 2);
                        subjectScore_dict[subjetScore.Key + "_平均成績"] = avgScore;
                        subjectLevel_dict[subjetScore.Key + "_平均成績等第"] = _ScoreMappingConfig.ParseScoreName(avgScore);
                    }

                    foreach (var learningDomainScore in learmingDomainScoreDic)
                    {
                        decimal avgScore = Math.Round((learningDomainScore.Value.Sum(x => x) / learningDomainScore.Value.Count), 2);
                        subjectScore_dict[learningDomainScore.Key + "_成績"] = avgScore;
                        subjectLevel_dict[learningDomainScore.Key + "_等第"] = _ScoreMappingConfig.ParseScoreName(avgScore);
                    }

                    // 填領域分數
                    foreach (string key in domainScore_dict.Keys)
                    {
                        if (domainScore_dict.ContainsKey(key))
                        {
                            if (!table.Columns.Contains(key))
                                table.Columns.Add(key);
                            
                            row[key] = domainScore_dict[key];
                        }
                            
                    }

                    // 填領域等第
                    foreach (string key in domainLevel_dict.Keys)
                    {
                        if (domainLevel_dict.ContainsKey(key))
                        {
                            if (!table.Columns.Contains(key))
                                table.Columns.Add(key);

                            row[key] = domainLevel_dict[key];
                        }
                            
                    }

                    // 填領域權數
                    foreach (string key in domainCredit_dict.Keys)
                    {
                        if (domainCredit_dict.ContainsKey(key))
                        {
                            if (!table.Columns.Contains(key))
                                table.Columns.Add(key);
                            
                            row[key] = domainCredit_dict[key];
                        }                            
                    }

                    // 填科目分數
                    foreach (string key in subjectScore_dict.Keys)
                    {
                        if (subjectScore_dict.ContainsKey(key))
                        {
                            if (!table.Columns.Contains(key))
                                table.Columns.Add(key);
                            row[key] = subjectScore_dict[key];
                        }                            
                    }

                    // 填科目等第
                    foreach (string key in subjectLevel_dict.Keys)
                    {
                        if (subjectLevel_dict.ContainsKey(key))
                        {
                            if (!table.Columns.Contains(key))
                                table.Columns.Add(key);

                            row[key] = subjectLevel_dict[key];
                        }
                            
                    }

                    // 填科目權數
                    foreach (string key in subjectCredit_dict.Keys)
                    {
                        if (subjectCredit_dict.ContainsKey(key))
                        {
                            if (!table.Columns.Contains(key))
                                table.Columns.Add(key);

                            row[key] = subjectCredit_dict[key];
                        }
                            
                    }

                }

                //畢業分數
                //if (gsr_dict.ContainsKey(stuID))
                //{
                //    row["畢業總成績_平均"] = gsr_dict[stuID].LearnDomainScore;
                //    row["畢業總成績_等第"] = ScoreTolevel(gsr_dict[stuID].LearnDomainScore);

                //    // 60 分 就可以 准予畢業
                //    row["准予畢業"] = gsr_dict[stuID].LearnDomainScore > 60 ? "■" : "□";
                //    row["發給修業證書"] = gsr_dict[stuID].LearnDomainScore > 60 ? "□" : "■";
                //}

                // 異動資料
                if (urr_dict.ContainsKey(stuID))
                {
                    foreach (K12.Data.UpdateRecordRecord urr in urr_dict[stuID])
                    {
                        // 新生異動為1 ，且理論上 一個人 會有1筆新生異動
                        if (urr.UpdateCode == "1")
                        {
                            DateTime enterday = new DateTime();

                            enterday = DateTime.Parse(urr.UpdateDate);
                            // 轉換入學時間 成 2005/09/06 的格式
                            //英文版格式  August 29, 2016
                            // 如果不加new CultureInfo("en-US") 資訊 會被轉成中文 "八月"
                            row["入學日期"] = enterday.ToString("MMMM dd,yyyy", new CultureInfo("en-US"));
                        }
                        if (urr.UpdateCode == "2")
                        {
                            DateTime enterday = new DateTime();

                            enterday = DateTime.Parse(urr.UpdateDate);
                            // 轉換入學時間 成 2005/09/06 的格式
                            //row["畢業日期"] = enterday.ToString("yyyy/MM/dd");

                            row["畢業日期"] = enterday.ToString("MMMM dd,yyyy", new CultureInfo("en-US"));

                        }
                    }
                }

                #region 日常生活表現、校內外特殊表現
                // 日常生活表現、校內外特殊表現
                //if (msr_dict.ContainsKey(stuID))
                //{
                //    for (int grade = 1; grade <= 3; grade++)
                //    {
                //        foreach (var msr in msr_dict[stuID])
                //        {
                //            if (msr.SchoolYear == schoolyear_grade_dict[grade])
                //            {
                //                if (msr.Semester == 1)
                //                {
                //                    if (textScore_dict.ContainsKey("日常生活表現及具體建議_" + (grade * 2 - 1)))
                //                    {
                //                        if (msr.TextScore.SelectSingleNode("DailyLifeRecommend") != null)
                //                        {
                //                            if (msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"] != null)
                //                            {
                //                                textScore_dict["日常生活表現及具體建議_" + (grade * 2 - 1)] = msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"].Value;
                //                            }
                //                        }
                //                    }
                //                    if (textScore_dict.ContainsKey("校內外特殊表現_" + (grade * 2 - 1)))
                //                    {
                //                        if (msr.TextScore.SelectSingleNode("OtherRecommend") != null)
                //                        {
                //                            if (msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"] != null)
                //                            {
                //                                textScore_dict["校內外特殊表現_" + (grade * 2 - 1)] = msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"].Value;
                //                            }
                //                        }

                //                    }
                //                }
                //                else
                //                {
                //                    if (textScore_dict.ContainsKey("日常生活表現及具體建議_" + (grade * 2)))
                //                    {
                //                        if (msr.TextScore.SelectSingleNode("DailyLifeRecommend") != null)
                //                        {
                //                            if (msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"] != null)
                //                            {
                //                                textScore_dict["日常生活表現及具體建議_" + (grade * 2)] = msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"].Value;
                //                            }
                //                        }
                //                    }
                //                    if (textScore_dict.ContainsKey("校內外特殊表現_" + (grade * 2)))
                //                    {
                //                        if (msr.TextScore.SelectSingleNode("OtherRecommend") != null)
                //                        {
                //                            if (msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"] != null)
                //                            {
                //                                textScore_dict["校內外特殊表現_" + (grade * 2)] = msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"].Value;
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }
                //    //填值
                //    foreach (string key in textScore_dict.Keys)
                //    {
                //        row[key] = textScore_dict[key];
                //    }
                //}
                #endregion

                row["匯出日期"] = DateTime.Today.ToString("MMMM dd,yyyy", new CultureInfo("en-US")); ;

                table.Rows.Add(row);

                //回報進度
                int percent = ((student_counter * 100 / StudentIDs.Count));

                MasterWorker.ReportProgress(percent, "學生在校成績證明書產生中...進行到第" + (student_counter) + "/" + StudentIDs.Count + "學生");

                student_counter++;
            }

            //選擇 目前的樣板
            document = new Document(Preference.Template.GetStream());

            //執行 合併列印
            document.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();
            document.MailMerge.Execute(table);

            // 最終產物 .doc
            e.Result = document;

            Feedback("列印完成", -1);
        }

        // 英文版 改為 A B C 
        // 換算分數 與 等第用
        private string ScoreTolevel(decimal? d)
        {
            string level = "";
            if (d >= 90)
            {
                level = "A+";
            }
            else if (d >= 80 && d < 90)
            {
                level = "A";
            }
            else if (d >= 70 && d < 80)
            {
                level = "B";
            }
            else if (d >= 60 && d < 70)
            {
                level = "C";
            }
            else if (d < 60)
            {
                level = "D";
            }
            else
            {
                level = "";
            }
            return level;
        }

        private void MasterWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Util.EnableControls(this);

            if (e.Error == null)
            {
                Document doc = e.Result as Document;

                //單檔列印
                if (OneFileSave.Checked)
                {
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    fbd.Description = "請選擇儲存資料夾";
                    fbd.ShowNewFolderButton = true;

                    if (fbd.ShowDialog() == DialogResult.Cancel) return;

                    fbdPath = fbd.SelectedPath;
                    System.Diagnostics.Process.Start(fbd.SelectedPath);

                    Util.DisableControls(this);
                    ConvertToPDF_Worker.RunWorkerAsync();
                }
                else
                {
                    if (Preference.ConvertToPDF)
                    {
                        MotherForm.SetStatusBarMessage("正在轉換PDF格式... 請耐心等候");
                    }
                    Util.DisableControls(this);
                    ConvertToPDF_Worker.RunWorkerAsync();
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show(e.Error.Message);
            }

            if (Preference.ConvertToPDF)
            {
                MotherForm.SetStatusBarMessage("正在轉換PDF格式", 0);
            }
            else
            {
                MotherForm.SetStatusBarMessage("產生完成", 100);
            }
        }

        private void ConvertToPDF_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Document doc = e_For_ConvertToPDF_Worker.Result as Document;

            if (!OneFileSave.Checked)
            {
                Util.Save(doc, "學生在校成績證明書(英文)", Preference.ConvertToPDF);
            }
            else
            {
                int i = 0;

                foreach (Section section in doc.Sections)
                {
                    // 依照 學號_身分字號_班級_座號_姓名 .doc 來存檔
                    string fileName = "";

                    Document document = new Document();
                    document.Sections.Clear();
                    document.Sections.Add(document.ImportNode(section, true));

                    fileName = PrintStudents[i].StudentNumber;

                    fileName += "_" + PrintStudents[i].IDNumber;

                    if (!string.IsNullOrEmpty(PrintStudents[i].RefClassID))
                    {
                        fileName += "_" + PrintStudents[i].Class.Name;
                    }
                    else
                    {
                        fileName += "_";
                    }

                    fileName += "_" + PrintStudents[i].SeatNo;

                    fileName += "_" + PrintStudents[i].Name;

                    //document.Save(fbd.SelectedPath + "\\" +fileName+ ".doc");

                    if (Preference.ConvertToPDF)
                    {
                        string fPath = fbdPath + "\\" + fileName + ".pdf";

                        document.Save(fPath, SaveFormat.Pdf);

                        int percent = (((i + 1) * 100 / doc.Sections.Count));

                        ConvertToPDF_Worker.ReportProgress(percent, "PDF轉換中...進行到" + (i + 1) + "/" + doc.Sections.Count + "個檔案");
                    }
                    else
                    {
                        document.Save(fbdPath + "\\" + fileName + ".doc");

                        int percent = (((i + 1) * 100 / doc.Sections.Count));

                        ConvertToPDF_Worker.ReportProgress(percent, "Doc存檔...進行到" + (i + 1) + "/" + doc.Sections.Count + "個檔案");
                    }

                    i++;


                }
            }
        }


        private void ConvertToPDF_Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Util.EnableControls(this);

            if (Preference.ConvertToPDF)
            {
                MotherForm.SetStatusBarMessage("PDF轉換完成", 100);

            }
            else
            {
                MotherForm.SetStatusBarMessage("存檔完成", 100);

            }
        }


        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        #region IStatusReporter 成員

        //回報進度條
        public void Feedback(string message, int percentage)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, int>(Feedback), new object[] { message, percentage });
            }
            else
            {
                if (percentage < 0)
                    MotherForm.SetStatusBarMessage(message);
                else
                    MotherForm.SetStatusBarMessage(message, percentage);

                Application.DoEvents();
            }
        }

        #endregion

        //2017/12/19 穎驊特別註解，這邊先採用舊寫法提供使用者設定樣板， 僅能上傳、下載舊版的.doc 格式，若傳.docx 會錯誤
        // 日後有時間再改新寫法
        private void lnkTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportTemplate defaultTemplate = new ReportTemplate(Properties.Resources.在校成績證明書_2022英文版_範本, TemplateType.Word);
            TemplateSettingForm form = new TemplateSettingForm(Preference.Template, defaultTemplate);
            form.DefaultFileName = "在校成績證明書(英文)(樣版).doc";

            if (form.ShowDialog() == DialogResult.OK)
            {
                Preference.Template = (form.Template == defaultTemplate) ? null : form.Template;
                Preference.Save();
            }
        }

        /// <summary>
        /// 下載合併欄位總表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Global.ExportMappingFieldWord("英文版");
        }

        /// <summary>
        /// 處理照片
        /// </summary>
        private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.FieldName == "照片")
                {
                    byte[] photo = e.FieldValue as byte[];
                    if (photo == null)
                        return;
                    DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                    photoBuilder.MoveToField(e.Field, true);
                    e.Field.Remove();

                    Shape photoShape = new Shape(e.Document, ShapeType.Image);
                    photoShape.ImageData.SetImage(photo);
                    double shapeHeight = 0;
                    double shapeWidth = 0;
                    photoShape.WrapType = WrapType.TopBottom;//設定文繞圖

                    //resize

                    //double origSizeRatio = photoShape.ImageData.ImageSize.HeightPoints / photoShape.ImageData.ImageSize.WidthPoints;
                    //Cell curCell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                    //shapeHeight = (curCell.ParentNode as Row).RowFormat.Height;
                    //shapeWidth = curCell.CellFormat.Width;
                    //photoShape.Height = shapeHeight;
                    //photoShape.Width = shapeWidth;

                    // 目前先固定為1吋大小，原本上面動態一表格大小填滿的方法，在在校成績證明書的樣板會被壓縮，暫時不處理。
                    // 1吋
                    photoShape.Width = ConvertUtil.MillimeterToPoint(25);
                    photoShape.Height = ConvertUtil.MillimeterToPoint(35);

                    photoBuilder.InsertNode(photoShape);
                }

            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }
        }
    }
}
