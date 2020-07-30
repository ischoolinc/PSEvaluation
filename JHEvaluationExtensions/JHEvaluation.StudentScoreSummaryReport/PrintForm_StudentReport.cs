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
using System.Linq;
using JHScoreReportDAL;
using JHSchool.Evaluation;

namespace JHEvaluation.StudentScoreSummaryReport
{
    public partial class PrintForm_StudentReport : BaseForm, IStatusReporter
    {
        internal const string ConfigName = "StudentReport";

        private List<string> StudentIDs { get; set; }

        private ReportPreference Preference { get; set; }

        private BackgroundWorker MasterWorker = new BackgroundWorker();

        private BackgroundWorker ConvertToPDF_Worker = new BackgroundWorker();

        //private List<ReportStudent> PrintStudents = new List<ReportStudent>();

        //領域科目資料管理設定的資料
        JHScoreReportDAL.Config _DomainSubjectConfig = new Config();
        List<string> _SujectTemplateList = new List<string>();

        private List<K12.Data.StudentRecord> PrintStudents = new List<K12.Data.StudentRecord>();

        // 學生基本資料 [studentID,Data]
        Dictionary<string, K12.Data.StudentRecord> sr_dict = new Dictionary<string, K12.Data.StudentRecord>();

        // 學生家長基本資料 [studentID,Data]
        Dictionary<string, K12.Data.ParentRecord> spr_dict = new Dictionary<string, K12.Data.ParentRecord>();

        // //學生聯繫資料(住址) [studentID,Data]
        Dictionary<string, K12.Data.AddressRecord> sar_dict = new Dictionary<string, K12.Data.AddressRecord>();

        // 學生聯繫資料(電話) [studentID,Data]
        Dictionary<string, K12.Data.PhoneRecord> sphr_dict = new Dictionary<string, K12.Data.PhoneRecord>();

        // 學期歷程 [studentID,Data]
        Dictionary<string, K12.Data.SemesterHistoryRecord> shr_dict = new Dictionary<string, K12.Data.SemesterHistoryRecord>();


        // 缺曠 [studentID,List<Data>]
        Dictionary<string, List<K12.Data.AttendanceRecord>> ar_dict = new Dictionary<string, List<K12.Data.AttendanceRecord>>();

        //學期成績(領域、科目) [studentID,List<Data>]
        Dictionary<string, List<JHSemesterScoreRecord>> jssr_dict = new Dictionary<string, List<JHSemesterScoreRecord>>();

        //畢業分數 [studentID,Data]
        Dictionary<string, K12.Data.GradScoreRecord> gsr_dict = new Dictionary<string, K12.Data.GradScoreRecord>();

        // 異動 [studentID,List<Data>]
        Dictionary<string, List<K12.Data.UpdateRecordRecord>> urr_dict = new Dictionary<string, List<K12.Data.UpdateRecordRecord>>();

        // 日常生活表現、校內外特殊表現 [studentID,List<Data>]
        Dictionary<string, List<K12.Data.MoralScoreRecord>> msr_dict = new Dictionary<string, List<K12.Data.MoralScoreRecord>>();

        // 上課節次 設定(列入、不列入)
        List<K12.Data.PeriodMappingInfo> periodMappingInfos = K12.Data.PeriodMapping.SelectAll();

        // 缺曠的節次(一般)名稱
        List<string> AbsencePeriod = new List<string>();

        private string fbdPath = "";

        private DoWorkEventArgs e_For_ConvertToPDF_Worker;

        public PrintForm_StudentReport(List<string> studentIds)
        {
            InitializeComponent();

            StudentIDs = studentIds;
            Preference = new ReportPreference(ConfigName, Prc.康橋國小學籍表_樣板_);
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


            #region 清除字典資料
            _SujectTemplateList.Clear();
            sr_dict.Clear();
            spr_dict.Clear();
            sar_dict.Clear();
            sphr_dict.Clear();
            shr_dict.Clear();
            ar_dict.Clear();
            jssr_dict.Clear();
            gsr_dict.Clear();
            urr_dict.Clear();
            msr_dict.Clear();
            #endregion

            //取得領域科目資料管理設定的資料
            _DomainSubjectConfig.GetConfigData();
            foreach (ConfigItem item in _DomainSubjectConfig.GetSubjectItemList())
            {
                _SujectTemplateList.Add(item.Name);
            }

            #region 抓取學生資料 
            //抓取學生資料 
            //學生基本資料
            List<K12.Data.StudentRecord> sr_list = K12.Data.Student.SelectByIDs(StudentIDs);

            //學生聯繫資料(住址)
            List<K12.Data.AddressRecord> sar_list = K12.Data.Address.SelectByStudentIDs(StudentIDs);

            //學生聯繫資料(電話)
            List<K12.Data.PhoneRecord> sphr_list = K12.Data.Phone.SelectByStudentIDs(StudentIDs);

            //學生家長基本資料
            List<K12.Data.ParentRecord> spr_list = K12.Data.Parent.SelectByStudentIDs(StudentIDs);

            //學期成績(包含領域、科目) (K12的被保護起來 Domain 看的到 抓不出來)
            //List<K12.Data.SemesterScoreRecord> ssr_list = K12.Data.SemesterScore.SelectByStudentIDs(StudentIDs);

            //學期成績(包含領域、科目) (改用 JHSemesterScoreRecord 才抓的到)
            List<JHSemesterScoreRecord> ssr_list = JHSemesterScore.SelectByStudentIDs(StudentIDs);

            //缺曠
            List<K12.Data.AttendanceRecord> ar_list = K12.Data.Attendance.SelectByStudentIDs(StudentIDs);


            //學生異動
            List<K12.Data.UpdateRecordRecord> urr_list = K12.Data.UpdateRecord.SelectByStudentIDs(StudentIDs);

            //畢業分數                                                   此處SelectByIDs 確實是SelectByStudentIDs
            List<K12.Data.GradScoreRecord> gsr_list = K12.Data.GradScore.SelectByIDs<K12.Data.GradScoreRecord>(StudentIDs);

            //學期歷程
            List<K12.Data.SemesterHistoryRecord> shr_list = K12.Data.SemesterHistory.SelectByStudentIDs(StudentIDs);

            //日常生活表現、校內外特殊表現
            List<K12.Data.MoralScoreRecord> msr_list = K12.Data.MoralScore.SelectByStudentIDs(StudentIDs);
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

            //整理學家長生基本資料
            foreach (K12.Data.ParentRecord spr in spr_list)
            {
                if (!spr_dict.ContainsKey(spr.RefStudentID))
                {
                    spr_dict.Add(spr.RefStudentID, spr);
                }
            }

            //整理學生聯繫資料(住址)
            foreach (K12.Data.AddressRecord sar in sar_list)
            {
                if (!sar_dict.ContainsKey(sar.RefStudentID))
                {
                    sar_dict.Add(sar.RefStudentID, sar);
                }
            }

            //整理學生聯繫資料(電話)
            foreach (K12.Data.PhoneRecord sphr in sphr_list)
            {
                if (!sphr_dict.ContainsKey(sphr.RefStudentID))
                {
                    sphr_dict.Add(sphr.RefStudentID, sphr);
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
                Dictionary<string, K12.Data.SubjectScore> subjectScoreDic = new Dictionary<string, K12.Data.SubjectScore>();
                List<string> subjectNameList = new List<string>();
                subjectNameList = ssr.Subjects.Keys.ToList();

                subjectNameList.Sort(new PSEvaluation.StudentScoreSummaryReport.StringComparer(_SujectTemplateList.ToArray()));

                foreach (string subjectName in subjectNameList)
                {
                    K12.Data.SubjectScore score = ssr.Subjects[subjectName];
                    subjectScoreDic.Add(subjectName, score);
                }

                ssr.Subjects = subjectScoreDic;
            }

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
            foreach (K12.Data.GradScoreRecord gsr in gsr_list)
            {
                if (!gsr_dict.ContainsKey(gsr.RefStudentID))
                {
                    gsr_dict.Add(gsr.RefStudentID, gsr);
                }
            }


            //整理出異動紀錄
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

            //整理出日常生活表現、校內外特殊表現
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

            #region 基本資料
            //基本資料
            table.Columns.Add("學生姓名");
            table.Columns.Add("學生班級");
            table.Columns.Add("學生性別");
            table.Columns.Add("出生日期");
            table.Columns.Add("入學年月");
            table.Columns.Add("學生身分證字號");
            table.Columns.Add("學號");
            table.Columns.Add("戶籍地址");
            table.Columns.Add("戶籍電話");
            table.Columns.Add("聯絡地址");
            table.Columns.Add("聯絡電話");
            #endregion

            #region 異動紀錄
            // 2020/07/16 俊緯更新，新增核准文號並增加可以記錄的異動資料數量
            //異動紀錄
            table.Columns.Add("異動紀錄1_日期");
            table.Columns.Add("異動紀錄1_校名");
            table.Columns.Add("異動紀錄1_類別");
            table.Columns.Add("異動紀錄1_核准日期");
            table.Columns.Add("異動紀錄1_核准文號");
            table.Columns.Add("異動紀錄2_日期");
            table.Columns.Add("異動紀錄2_校名");
            table.Columns.Add("異動紀錄2_類別");
            table.Columns.Add("異動紀錄2_核准日期");
            table.Columns.Add("異動紀錄2_核准文號");
            table.Columns.Add("異動紀錄3_日期");
            table.Columns.Add("異動紀錄3_校名");
            table.Columns.Add("異動紀錄3_類別");
            table.Columns.Add("異動紀錄3_核准日期");
            table.Columns.Add("異動紀錄3_核准文號");
            table.Columns.Add("異動紀錄4_日期");
            table.Columns.Add("異動紀錄4_校名");
            table.Columns.Add("異動紀錄4_類別");
            table.Columns.Add("異動紀錄4_核准日期");
            table.Columns.Add("異動紀錄4_核准文號");
            table.Columns.Add("異動紀錄5_日期");
            table.Columns.Add("異動紀錄5_校名");
            table.Columns.Add("異動紀錄5_類別");
            table.Columns.Add("異動紀錄5_核准日期");
            table.Columns.Add("異動紀錄5_核准文號");
            table.Columns.Add("異動紀錄6_日期");
            table.Columns.Add("異動紀錄6_校名");
            table.Columns.Add("異動紀錄6_類別");
            table.Columns.Add("異動紀錄6_核准日期");
            table.Columns.Add("異動紀錄6_核准文號");
            table.Columns.Add("異動紀錄7_日期");
            table.Columns.Add("異動紀錄7_校名");
            table.Columns.Add("異動紀錄7_類別");
            table.Columns.Add("異動紀錄7_核准日期");
            table.Columns.Add("異動紀錄7_核准文號");
            table.Columns.Add("異動紀錄8_日期");
            table.Columns.Add("異動紀錄8_校名");
            table.Columns.Add("異動紀錄8_類別");
            table.Columns.Add("異動紀錄8_核准日期");
            table.Columns.Add("異動紀錄8_核准文號");
            table.Columns.Add("異動紀錄9_日期");
            table.Columns.Add("異動紀錄9_校名");
            table.Columns.Add("異動紀錄9_類別");
            table.Columns.Add("異動紀錄9_核准日期");
            table.Columns.Add("異動紀錄9_核准文號");
            table.Columns.Add("異動紀錄10_日期");
            table.Columns.Add("異動紀錄10_校名");
            table.Columns.Add("異動紀錄10_類別");
            table.Columns.Add("異動紀錄10_核准日期");
            table.Columns.Add("異動紀錄10_核准文號");

            #endregion

            #region 家長資料
            table.Columns.Add("監護人姓名");
            table.Columns.Add("監護人關係");
            table.Columns.Add("監護人行動電話");
            table.Columns.Add("父親姓名");
            table.Columns.Add("父親行動電話");
            table.Columns.Add("母親姓名");
            table.Columns.Add("母親行動電話");
            #endregion

            #region 學期歷程 班級、座號、班導師資料
            //班級座號資料
            table.Columns.Add("班級1");
            table.Columns.Add("座號1");
            table.Columns.Add("班導師1");

            table.Columns.Add("班級2");
            table.Columns.Add("座號2");
            table.Columns.Add("班導師2");

            table.Columns.Add("班級3");
            table.Columns.Add("座號3");
            table.Columns.Add("班導師3");

            table.Columns.Add("班級4");
            table.Columns.Add("座號4");
            table.Columns.Add("班導師4");

            table.Columns.Add("班級5");
            table.Columns.Add("座號5");
            table.Columns.Add("班導師5");

            table.Columns.Add("班級6");
            table.Columns.Add("座號6");
            table.Columns.Add("班導師6");

            table.Columns.Add("班級7");
            table.Columns.Add("座號7");
            table.Columns.Add("班導師7");

            table.Columns.Add("班級8");
            table.Columns.Add("座號8");
            table.Columns.Add("班導師8");

            table.Columns.Add("班級9");
            table.Columns.Add("座號9");
            table.Columns.Add("班導師9");

            table.Columns.Add("班級10");
            table.Columns.Add("座號10");
            table.Columns.Add("班導師10");

            table.Columns.Add("班級11");
            table.Columns.Add("座號11");
            table.Columns.Add("班導師11");

            table.Columns.Add("班級12");
            table.Columns.Add("座號12");
            table.Columns.Add("班導師12");
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

            table.Columns.Add("領域_語文_成績_1");
            table.Columns.Add("領域_語文_成績_2");
            table.Columns.Add("領域_語文_成績_3");
            table.Columns.Add("領域_語文_成績_4");
            table.Columns.Add("領域_語文_成績_5");
            table.Columns.Add("領域_語文_成績_6");
            table.Columns.Add("領域_語文_成績_7");
            table.Columns.Add("領域_語文_成績_8");
            table.Columns.Add("領域_語文_成績_9");
            table.Columns.Add("領域_語文_成績_10");
            table.Columns.Add("領域_語文_成績_11");
            table.Columns.Add("領域_語文_成績_12");


            table.Columns.Add("領域_語文_等第_1");
            table.Columns.Add("領域_語文_等第_2");
            table.Columns.Add("領域_語文_等第_3");
            table.Columns.Add("領域_語文_等第_4");
            table.Columns.Add("領域_語文_等第_5");
            table.Columns.Add("領域_語文_等第_6");
            table.Columns.Add("領域_語文_等第_7");
            table.Columns.Add("領域_語文_等第_8");
            table.Columns.Add("領域_語文_等第_9");
            table.Columns.Add("領域_語文_等第_10");
            table.Columns.Add("領域_語文_等第_11");
            table.Columns.Add("領域_語文_等第_12");


            table.Columns.Add("領域_數學_成績_1");
            table.Columns.Add("領域_數學_成績_2");
            table.Columns.Add("領域_數學_成績_3");
            table.Columns.Add("領域_數學_成績_4");
            table.Columns.Add("領域_數學_成績_5");
            table.Columns.Add("領域_數學_成績_6");
            table.Columns.Add("領域_數學_成績_7");
            table.Columns.Add("領域_數學_成績_8");
            table.Columns.Add("領域_數學_成績_9");
            table.Columns.Add("領域_數學_成績_10");
            table.Columns.Add("領域_數學_成績_11");
            table.Columns.Add("領域_數學_成績_12");


            table.Columns.Add("領域_數學_等第_1");
            table.Columns.Add("領域_數學_等第_2");
            table.Columns.Add("領域_數學_等第_3");
            table.Columns.Add("領域_數學_等第_4");
            table.Columns.Add("領域_數學_等第_5");
            table.Columns.Add("領域_數學_等第_6");
            table.Columns.Add("領域_數學_等第_7");
            table.Columns.Add("領域_數學_等第_8");
            table.Columns.Add("領域_數學_等第_9");
            table.Columns.Add("領域_數學_等第_10");
            table.Columns.Add("領域_數學_等第_11");
            table.Columns.Add("領域_數學_等第_12");


            table.Columns.Add("領域_生活課程_成績_1");
            table.Columns.Add("領域_生活課程_成績_2");
            table.Columns.Add("領域_生活課程_成績_3");
            table.Columns.Add("領域_生活課程_成績_4");
            table.Columns.Add("領域_生活課程_成績_5");
            table.Columns.Add("領域_生活課程_成績_6");
            table.Columns.Add("領域_生活課程_成績_7");
            table.Columns.Add("領域_生活課程_成績_8");
            table.Columns.Add("領域_生活課程_成績_9");
            table.Columns.Add("領域_生活課程_成績_10");
            table.Columns.Add("領域_生活課程_成績_11");
            table.Columns.Add("領域_生活課程_成績_12");


            table.Columns.Add("領域_生活課程_等第_1");
            table.Columns.Add("領域_生活課程_等第_2");
            table.Columns.Add("領域_生活課程_等第_3");
            table.Columns.Add("領域_生活課程_等第_4");
            table.Columns.Add("領域_生活課程_等第_5");
            table.Columns.Add("領域_生活課程_等第_6");
            table.Columns.Add("領域_生活課程_等第_7");
            table.Columns.Add("領域_生活課程_等第_8");
            table.Columns.Add("領域_生活課程_等第_9");
            table.Columns.Add("領域_生活課程_等第_10");
            table.Columns.Add("領域_生活課程_等第_11");
            table.Columns.Add("領域_生活課程_等第_12");



            table.Columns.Add("領域_自然與生活科技_成績_1");
            table.Columns.Add("領域_自然與生活科技_成績_2");
            table.Columns.Add("領域_自然與生活科技_成績_3");
            table.Columns.Add("領域_自然與生活科技_成績_4");
            table.Columns.Add("領域_自然與生活科技_成績_5");
            table.Columns.Add("領域_自然與生活科技_成績_6");
            table.Columns.Add("領域_自然與生活科技_成績_7");
            table.Columns.Add("領域_自然與生活科技_成績_8");
            table.Columns.Add("領域_自然與生活科技_成績_9");
            table.Columns.Add("領域_自然與生活科技_成績_10");
            table.Columns.Add("領域_自然與生活科技_成績_11");
            table.Columns.Add("領域_自然與生活科技_成績_12");


            table.Columns.Add("領域_自然與生活科技_等第_1");
            table.Columns.Add("領域_自然與生活科技_等第_2");
            table.Columns.Add("領域_自然與生活科技_等第_3");
            table.Columns.Add("領域_自然與生活科技_等第_4");
            table.Columns.Add("領域_自然與生活科技_等第_5");
            table.Columns.Add("領域_自然與生活科技_等第_6");
            table.Columns.Add("領域_自然與生活科技_等第_7");
            table.Columns.Add("領域_自然與生活科技_等第_8");
            table.Columns.Add("領域_自然與生活科技_等第_9");
            table.Columns.Add("領域_自然與生活科技_等第_10");
            table.Columns.Add("領域_自然與生活科技_等第_11");
            table.Columns.Add("領域_自然與生活科技_等第_12");


            table.Columns.Add("領域_藝術與人文_成績_1");
            table.Columns.Add("領域_藝術與人文_成績_2");
            table.Columns.Add("領域_藝術與人文_成績_3");
            table.Columns.Add("領域_藝術與人文_成績_4");
            table.Columns.Add("領域_藝術與人文_成績_5");
            table.Columns.Add("領域_藝術與人文_成績_6");
            table.Columns.Add("領域_藝術與人文_成績_7");
            table.Columns.Add("領域_藝術與人文_成績_8");
            table.Columns.Add("領域_藝術與人文_成績_9");
            table.Columns.Add("領域_藝術與人文_成績_10");
            table.Columns.Add("領域_藝術與人文_成績_11");
            table.Columns.Add("領域_藝術與人文_成績_12");


            table.Columns.Add("領域_藝術與人文_等第_1");
            table.Columns.Add("領域_藝術與人文_等第_2");
            table.Columns.Add("領域_藝術與人文_等第_3");
            table.Columns.Add("領域_藝術與人文_等第_4");
            table.Columns.Add("領域_藝術與人文_等第_5");
            table.Columns.Add("領域_藝術與人文_等第_6");
            table.Columns.Add("領域_藝術與人文_等第_7");
            table.Columns.Add("領域_藝術與人文_等第_8");
            table.Columns.Add("領域_藝術與人文_等第_9");
            table.Columns.Add("領域_藝術與人文_等第_10");
            table.Columns.Add("領域_藝術與人文_等第_11");
            table.Columns.Add("領域_藝術與人文_等第_12");



            table.Columns.Add("領域_社會_成績_1");
            table.Columns.Add("領域_社會_成績_2");
            table.Columns.Add("領域_社會_成績_3");
            table.Columns.Add("領域_社會_成績_4");
            table.Columns.Add("領域_社會_成績_5");
            table.Columns.Add("領域_社會_成績_6");
            table.Columns.Add("領域_社會_成績_7");
            table.Columns.Add("領域_社會_成績_8");
            table.Columns.Add("領域_社會_成績_9");
            table.Columns.Add("領域_社會_成績_10");
            table.Columns.Add("領域_社會_成績_11");
            table.Columns.Add("領域_社會_成績_12");


            table.Columns.Add("領域_社會_等第_1");
            table.Columns.Add("領域_社會_等第_2");
            table.Columns.Add("領域_社會_等第_3");
            table.Columns.Add("領域_社會_等第_4");
            table.Columns.Add("領域_社會_等第_5");
            table.Columns.Add("領域_社會_等第_6");
            table.Columns.Add("領域_社會_等第_7");
            table.Columns.Add("領域_社會_等第_8");
            table.Columns.Add("領域_社會_等第_9");
            table.Columns.Add("領域_社會_等第_10");
            table.Columns.Add("領域_社會_等第_11");
            table.Columns.Add("領域_社會_等第_12");


            table.Columns.Add("領域_健康與體育_成績_1");
            table.Columns.Add("領域_健康與體育_成績_2");
            table.Columns.Add("領域_健康與體育_成績_3");
            table.Columns.Add("領域_健康與體育_成績_4");
            table.Columns.Add("領域_健康與體育_成績_5");
            table.Columns.Add("領域_健康與體育_成績_6");
            table.Columns.Add("領域_健康與體育_成績_7");
            table.Columns.Add("領域_健康與體育_成績_8");
            table.Columns.Add("領域_健康與體育_成績_9");
            table.Columns.Add("領域_健康與體育_成績_10");
            table.Columns.Add("領域_健康與體育_成績_11");
            table.Columns.Add("領域_健康與體育_成績_12");


            table.Columns.Add("領域_健康與體育_等第_1");
            table.Columns.Add("領域_健康與體育_等第_2");
            table.Columns.Add("領域_健康與體育_等第_3");
            table.Columns.Add("領域_健康與體育_等第_4");
            table.Columns.Add("領域_健康與體育_等第_5");
            table.Columns.Add("領域_健康與體育_等第_6");
            table.Columns.Add("領域_健康與體育_等第_7");
            table.Columns.Add("領域_健康與體育_等第_8");
            table.Columns.Add("領域_健康與體育_等第_9");
            table.Columns.Add("領域_健康與體育_等第_10");
            table.Columns.Add("領域_健康與體育_等第_11");
            table.Columns.Add("領域_健康與體育_等第_12");


            table.Columns.Add("領域_綜合活動_成績_1");
            table.Columns.Add("領域_綜合活動_成績_2");
            table.Columns.Add("領域_綜合活動_成績_3");
            table.Columns.Add("領域_綜合活動_成績_4");
            table.Columns.Add("領域_綜合活動_成績_5");
            table.Columns.Add("領域_綜合活動_成績_6");
            table.Columns.Add("領域_綜合活動_成績_7");
            table.Columns.Add("領域_綜合活動_成績_8");
            table.Columns.Add("領域_綜合活動_成績_9");
            table.Columns.Add("領域_綜合活動_成績_10");
            table.Columns.Add("領域_綜合活動_成績_11");
            table.Columns.Add("領域_綜合活動_成績_12");


            table.Columns.Add("領域_綜合活動_等第_1");
            table.Columns.Add("領域_綜合活動_等第_2");
            table.Columns.Add("領域_綜合活動_等第_3");
            table.Columns.Add("領域_綜合活動_等第_4");
            table.Columns.Add("領域_綜合活動_等第_5");
            table.Columns.Add("領域_綜合活動_等第_6");
            table.Columns.Add("領域_綜合活動_等第_7");
            table.Columns.Add("領域_綜合活動_等第_8");
            table.Columns.Add("領域_綜合活動_等第_9");
            table.Columns.Add("領域_綜合活動_等第_10");
            table.Columns.Add("領域_綜合活動_等第_11");
            table.Columns.Add("領域_綜合活動_等第_12");

            table.Columns.Add("領域_彈性課程_成績_1");
            table.Columns.Add("領域_彈性課程_成績_2");
            table.Columns.Add("領域_彈性課程_成績_3");
            table.Columns.Add("領域_彈性課程_成績_4");
            table.Columns.Add("領域_彈性課程_成績_5");
            table.Columns.Add("領域_彈性課程_成績_6");
            table.Columns.Add("領域_彈性課程_成績_7");
            table.Columns.Add("領域_彈性課程_成績_8");
            table.Columns.Add("領域_彈性課程_成績_9");
            table.Columns.Add("領域_彈性課程_成績_10");
            table.Columns.Add("領域_彈性課程_成績_11");
            table.Columns.Add("領域_彈性課程_成績_12");


            table.Columns.Add("領域_彈性課程_等第_1");
            table.Columns.Add("領域_彈性課程_等第_2");
            table.Columns.Add("領域_彈性課程_等第_3");
            table.Columns.Add("領域_彈性課程_等第_4");
            table.Columns.Add("領域_彈性課程_等第_5");
            table.Columns.Add("領域_彈性課程_等第_6");
            table.Columns.Add("領域_彈性課程_等第_7");
            table.Columns.Add("領域_彈性課程_等第_8");
            table.Columns.Add("領域_彈性課程_等第_9");
            table.Columns.Add("領域_彈性課程_等第_10");
            table.Columns.Add("領域_彈性課程_等第_11");
            table.Columns.Add("領域_彈性課程_等第_12");


            table.Columns.Add("領域_學習領域總成績_成績_1");
            table.Columns.Add("領域_學習領域總成績_成績_2");
            table.Columns.Add("領域_學習領域總成績_成績_3");
            table.Columns.Add("領域_學習領域總成績_成績_4");
            table.Columns.Add("領域_學習領域總成績_成績_5");
            table.Columns.Add("領域_學習領域總成績_成績_6");
            table.Columns.Add("領域_學習領域總成績_成績_7");
            table.Columns.Add("領域_學習領域總成績_成績_8");
            table.Columns.Add("領域_學習領域總成績_成績_9");
            table.Columns.Add("領域_學習領域總成績_成績_10");
            table.Columns.Add("領域_學習領域總成績_成績_11");
            table.Columns.Add("領域_學習領域總成績_成績_12");


            table.Columns.Add("領域_學習領域總成績_等第_1");
            table.Columns.Add("領域_學習領域總成績_等第_2");
            table.Columns.Add("領域_學習領域總成績_等第_3");
            table.Columns.Add("領域_學習領域總成績_等第_4");
            table.Columns.Add("領域_學習領域總成績_等第_5");
            table.Columns.Add("領域_學習領域總成績_等第_6");
            table.Columns.Add("領域_學習領域總成績_等第_7");
            table.Columns.Add("領域_學習領域總成績_等第_8");
            table.Columns.Add("領域_學習領域總成績_等第_9");
            table.Columns.Add("領域_學習領域總成績_等第_10");
            table.Columns.Add("領域_學習領域總成績_等第_11");
            table.Columns.Add("領域_學習領域總成績_等第_12");

            table.Columns.Add("領域_課程學習成績_成績_1");
            table.Columns.Add("領域_課程學習成績_成績_2");
            table.Columns.Add("領域_課程學習成績_成績_3");
            table.Columns.Add("領域_課程學習成績_成績_4");
            table.Columns.Add("領域_課程學習成績_成績_5");
            table.Columns.Add("領域_課程學習成績_成績_6");
            table.Columns.Add("領域_課程學習成績_成績_7");
            table.Columns.Add("領域_課程學習成績_成績_8");
            table.Columns.Add("領域_課程學習成績_成績_9");
            table.Columns.Add("領域_課程學習成績_成績_10");
            table.Columns.Add("領域_課程學習成績_成績_11");
            table.Columns.Add("領域_課程學習成績_成績_12");


            table.Columns.Add("領域_課程學習成績_等第_1");
            table.Columns.Add("領域_課程學習成績_等第_2");
            table.Columns.Add("領域_課程學習成績_等第_3");
            table.Columns.Add("領域_課程學習成績_等第_4");
            table.Columns.Add("領域_課程學習成績_等第_5");
            table.Columns.Add("領域_課程學習成績_等第_6");
            table.Columns.Add("領域_課程學習成績_等第_7");
            table.Columns.Add("領域_課程學習成績_等第_8");
            table.Columns.Add("領域_課程學習成績_等第_9");
            table.Columns.Add("領域_課程學習成績_等第_10");
            table.Columns.Add("領域_課程學習成績_等第_11");
            table.Columns.Add("領域_課程學習成績_等第_12");


            #endregion

            #region 科目成績
            //科目成績
            table.Columns.Add("語文1_科目名稱");
            table.Columns.Add("語文2_科目名稱");
            table.Columns.Add("語文3_科目名稱");
            table.Columns.Add("語文4_科目名稱");
            table.Columns.Add("語文5_科目名稱");
            table.Columns.Add("語文6_科目名稱");
            table.Columns.Add("數學1_科目名稱");
            table.Columns.Add("數學2_科目名稱");
            table.Columns.Add("數學3_科目名稱");
            table.Columns.Add("數學4_科目名稱");
            table.Columns.Add("數學5_科目名稱");
            table.Columns.Add("數學6_科目名稱");
            table.Columns.Add("生活課程1_科目名稱");
            table.Columns.Add("生活課程2_科目名稱");
            table.Columns.Add("生活課程3_科目名稱");
            table.Columns.Add("生活課程4_科目名稱");
            table.Columns.Add("生活課程5_科目名稱");
            table.Columns.Add("生活課程6_科目名稱");
            table.Columns.Add("自然科學1_科目名稱");
            table.Columns.Add("自然科學2_科目名稱");
            table.Columns.Add("自然科學3_科目名稱");
            table.Columns.Add("自然科學4_科目名稱");
            table.Columns.Add("自然科學5_科目名稱");
            table.Columns.Add("自然科學6_科目名稱");
            table.Columns.Add("科技1_科目名稱");
            table.Columns.Add("科技2_科目名稱");
            table.Columns.Add("科技3_科目名稱");
            table.Columns.Add("科技4_科目名稱");
            table.Columns.Add("科技5_科目名稱");
            table.Columns.Add("科技6_科目名稱");
            table.Columns.Add("社會1_科目名稱");
            table.Columns.Add("社會2_科目名稱");
            table.Columns.Add("社會3_科目名稱");
            table.Columns.Add("社會4_科目名稱");
            table.Columns.Add("社會5_科目名稱");
            table.Columns.Add("社會6_科目名稱");
            table.Columns.Add("藝術1_科目名稱");
            table.Columns.Add("藝術2_科目名稱");
            table.Columns.Add("藝術3_科目名稱");
            table.Columns.Add("藝術4_科目名稱");
            table.Columns.Add("藝術5_科目名稱");
            table.Columns.Add("藝術6_科目名稱");
            table.Columns.Add("健康與體育1_科目名稱");
            table.Columns.Add("健康與體育2_科目名稱");
            table.Columns.Add("健康與體育3_科目名稱");
            table.Columns.Add("健康與體育4_科目名稱");
            table.Columns.Add("健康與體育5_科目名稱");
            table.Columns.Add("健康與體育6_科目名稱");
            table.Columns.Add("綜合活動1_科目名稱");
            table.Columns.Add("綜合活動2_科目名稱");
            table.Columns.Add("綜合活動3_科目名稱");
            table.Columns.Add("綜合活動4_科目名稱");
            table.Columns.Add("綜合活動5_科目名稱");
            table.Columns.Add("綜合活動6_科目名稱");
            table.Columns.Add("藝術與人文1_科目名稱");
            table.Columns.Add("藝術與人文2_科目名稱");
            table.Columns.Add("藝術與人文3_科目名稱");
            table.Columns.Add("藝術與人文4_科目名稱");
            table.Columns.Add("藝術與人文5_科目名稱");
            table.Columns.Add("藝術與人文6_科目名稱");
            table.Columns.Add("自然與生活科技1_科目名稱");
            table.Columns.Add("自然與生活科技2_科目名稱");
            table.Columns.Add("自然與生活科技3_科目名稱");
            table.Columns.Add("自然與生活科技4_科目名稱");
            table.Columns.Add("自然與生活科技5_科目名稱");
            table.Columns.Add("自然與生活科技6_科目名稱");

            table.Columns.Add("彈性課程1_科目名稱");
            table.Columns.Add("彈性課程2_科目名稱");
            table.Columns.Add("彈性課程3_科目名稱");
            table.Columns.Add("彈性課程4_科目名稱");
            table.Columns.Add("彈性課程5_科目名稱");
            table.Columns.Add("彈性課程6_科目名稱");
            table.Columns.Add("彈性課程7_科目名稱");
            table.Columns.Add("彈性課程8_科目名稱");
            table.Columns.Add("彈性課程9_科目名稱");
            table.Columns.Add("彈性課程10_科目名稱");

            #endregion

            #region 出缺勤紀錄
            //出缺勤紀錄
            table.Columns.Add("應出席日數_1");
            table.Columns.Add("應出席日數_2");
            table.Columns.Add("應出席日數_3");
            table.Columns.Add("應出席日數_4");
            table.Columns.Add("應出席日數_5");
            table.Columns.Add("應出席日數_6");
            table.Columns.Add("應出席日數_7");
            table.Columns.Add("應出席日數_8");
            table.Columns.Add("應出席日數_9");
            table.Columns.Add("應出席日數_10");
            table.Columns.Add("應出席日數_11");
            table.Columns.Add("應出席日數_12");


            table.Columns.Add("事假日數_1");
            table.Columns.Add("事假日數_2");
            table.Columns.Add("事假日數_3");
            table.Columns.Add("事假日數_4");
            table.Columns.Add("事假日數_5");
            table.Columns.Add("事假日數_6");
            table.Columns.Add("事假日數_7");
            table.Columns.Add("事假日數_8");
            table.Columns.Add("事假日數_9");
            table.Columns.Add("事假日數_10");
            table.Columns.Add("事假日數_11");
            table.Columns.Add("事假日數_12");


            table.Columns.Add("病假日數_1");
            table.Columns.Add("病假日數_2");
            table.Columns.Add("病假日數_3");
            table.Columns.Add("病假日數_4");
            table.Columns.Add("病假日數_5");
            table.Columns.Add("病假日數_6");
            table.Columns.Add("病假日數_7");
            table.Columns.Add("病假日數_8");
            table.Columns.Add("病假日數_9");
            table.Columns.Add("病假日數_10");
            table.Columns.Add("病假日數_11");
            table.Columns.Add("病假日數_12");

            table.Columns.Add("公假日數_1");
            table.Columns.Add("公假日數_2");
            table.Columns.Add("公假日數_3");
            table.Columns.Add("公假日數_4");
            table.Columns.Add("公假日數_5");
            table.Columns.Add("公假日數_6");
            table.Columns.Add("公假日數_7");
            table.Columns.Add("公假日數_8");
            table.Columns.Add("公假日數_9");
            table.Columns.Add("公假日數_10");
            table.Columns.Add("公假日數_11");
            table.Columns.Add("公假日數_12");

            table.Columns.Add("喪假日數_1");
            table.Columns.Add("喪假日數_2");
            table.Columns.Add("喪假日數_3");
            table.Columns.Add("喪假日數_4");
            table.Columns.Add("喪假日數_5");
            table.Columns.Add("喪假日數_6");
            table.Columns.Add("喪假日數_7");
            table.Columns.Add("喪假日數_8");
            table.Columns.Add("喪假日數_9");
            table.Columns.Add("喪假日數_10");
            table.Columns.Add("喪假日數_11");
            table.Columns.Add("喪假日數_12");

            table.Columns.Add("曠課日數_1");
            table.Columns.Add("曠課日數_2");
            table.Columns.Add("曠課日數_3");
            table.Columns.Add("曠課日數_4");
            table.Columns.Add("曠課日數_5");
            table.Columns.Add("曠課日數_6");
            table.Columns.Add("曠課日數_7");
            table.Columns.Add("曠課日數_8");
            table.Columns.Add("曠課日數_9");
            table.Columns.Add("曠課日數_10");
            table.Columns.Add("曠課日數_11");
            table.Columns.Add("曠課日數_12");

            table.Columns.Add("缺席總日數_1");
            table.Columns.Add("缺席總日數_2");
            table.Columns.Add("缺席總日數_3");
            table.Columns.Add("缺席總日數_4");
            table.Columns.Add("缺席總日數_5");
            table.Columns.Add("缺席總日數_6");
            table.Columns.Add("缺席總日數_7");
            table.Columns.Add("缺席總日數_8");
            table.Columns.Add("缺席總日數_9");
            table.Columns.Add("缺席總日數_10");
            table.Columns.Add("缺席總日數_11");
            table.Columns.Add("缺席總日數_12");


            #endregion

            #region 畢業總成績
            //畢業總成績
            table.Columns.Add("畢業總成績_平均");
            table.Columns.Add("畢業總成績_等第");
            table.Columns.Add("准予畢業");
            table.Columns.Add("發給修業證書");
            #endregion

            #region 日常生活表現及具體建議
            //日常生活表現及具體建議
            table.Columns.Add("日常生活表現及具體建議_1");
            table.Columns.Add("日常生活表現及具體建議_2");
            table.Columns.Add("日常生活表現及具體建議_3");
            table.Columns.Add("日常生活表現及具體建議_4");
            table.Columns.Add("日常生活表現及具體建議_5");
            table.Columns.Add("日常生活表現及具體建議_6");
            table.Columns.Add("日常生活表現及具體建議_7");
            table.Columns.Add("日常生活表現及具體建議_8");
            table.Columns.Add("日常生活表現及具體建議_9");
            table.Columns.Add("日常生活表現及具體建議_10");
            table.Columns.Add("日常生活表現及具體建議_11");
            table.Columns.Add("日常生活表現及具體建議_12");


            #endregion

            #region 校內外特殊表現
            //校內外特殊表現
            table.Columns.Add("校內外特殊表現_1");
            table.Columns.Add("校內外特殊表現_2");
            table.Columns.Add("校內外特殊表現_3");
            table.Columns.Add("校內外特殊表現_4");
            table.Columns.Add("校內外特殊表現_5");
            table.Columns.Add("校內外特殊表現_6");
            table.Columns.Add("校內外特殊表現_7");
            table.Columns.Add("校內外特殊表現_8");
            table.Columns.Add("校內外特殊表現_9");
            table.Columns.Add("校內外特殊表現_10");
            table.Columns.Add("校內外特殊表現_11");
            table.Columns.Add("校內外特殊表現_12");


            #endregion 

            #endregion

            Aspose.Words.Document document = new Aspose.Words.Document();

            e_For_ConvertToPDF_Worker = e;

            #region 整理所有的假別
            //整理所有的假別
            List<string> absenceType_list = new List<string>();

            absenceType_list.Add("事假");
            absenceType_list.Add("病假");
            absenceType_list.Add("公假");
            absenceType_list.Add("喪假");
            absenceType_list.Add("曠課");
            absenceType_list.Add("缺席總");
            #endregion

            #region 整理所有的領域_OO_成績
            //整理所有的領域_OO_成績
            List<string> domainScoreType_list = new List<string>();

            domainScoreType_list.Add("領域_語文_成績_");
            domainScoreType_list.Add("領域_數學_成績_");
            domainScoreType_list.Add("領域_生活課程_成績_");
            domainScoreType_list.Add("領域_自然與生活科技_成績_");
            domainScoreType_list.Add("領域_藝術與人文_成績_");
            domainScoreType_list.Add("領域_社會_成績_");
            domainScoreType_list.Add("領域_健康與體育_成績_");
            domainScoreType_list.Add("領域_綜合活動_成績_");
            domainScoreType_list.Add("領域_學習領域總成績_成績_");
            domainScoreType_list.Add("領域_課程學習成績_成績_");
            #endregion

            #region 整理所有的領域_OO_等第
            //整理所有的領域_OO_等第
            List<string> domainLevelType_list = new List<string>();

            domainLevelType_list.Add("領域_語文_等第_");
            domainLevelType_list.Add("領域_數學_等第_");
            domainLevelType_list.Add("領域_生活課程_等第_");
            domainLevelType_list.Add("領域_自然與生活科技_等第_");
            domainLevelType_list.Add("領域_藝術與人文_等第_");
            domainLevelType_list.Add("領域_社會_等第_");
            domainLevelType_list.Add("領域_健康與體育_等第_");
            domainLevelType_list.Add("領域_綜合活動_等第_");
            domainLevelType_list.Add("領域_學習領域總成績_等第_");
            domainLevelType_list.Add("領域_課程學習成績_等第_");
            #endregion

            #region 整理科目_OO_成績
            //整理科目_OO_成績
            List<string> subjectScoreType_list = new List<string>();

            subjectScoreType_list.Add("語文1_科目成績_");
            subjectScoreType_list.Add("語文2_科目成績_");
            subjectScoreType_list.Add("語文3_科目成績_");
            subjectScoreType_list.Add("語文4_科目成績_");
            subjectScoreType_list.Add("語文5_科目成績_");
            subjectScoreType_list.Add("語文6_科目成績_");
            subjectScoreType_list.Add("數學1_科目成績_");
            subjectScoreType_list.Add("數學2_科目成績_");
            subjectScoreType_list.Add("數學3_科目成績_");
            subjectScoreType_list.Add("數學4_科目成績_");
            subjectScoreType_list.Add("數學5_科目成績_");
            subjectScoreType_list.Add("數學6_科目成績_");
            subjectScoreType_list.Add("生活課程1_科目成績_");
            subjectScoreType_list.Add("生活課程2_科目成績_");
            subjectScoreType_list.Add("生活課程3_科目成績_");
            subjectScoreType_list.Add("生活課程4_科目成績_");
            subjectScoreType_list.Add("生活課程5_科目成績_");
            subjectScoreType_list.Add("生活課程6_科目成績_");
            subjectScoreType_list.Add("自然科學1_科目成績_");
            subjectScoreType_list.Add("自然科學2_科目成績_");
            subjectScoreType_list.Add("自然科學3_科目成績_");
            subjectScoreType_list.Add("自然科學4_科目成績_");
            subjectScoreType_list.Add("自然科學5_科目成績_");
            subjectScoreType_list.Add("自然科學6_科目成績_");
            subjectScoreType_list.Add("科技1_科目成績_");
            subjectScoreType_list.Add("科技2_科目成績_");
            subjectScoreType_list.Add("科技3_科目成績_");
            subjectScoreType_list.Add("科技4_科目成績_");
            subjectScoreType_list.Add("科技5_科目成績_");
            subjectScoreType_list.Add("科技6_科目成績_");
            subjectScoreType_list.Add("社會1_科目成績_");
            subjectScoreType_list.Add("社會2_科目成績_");
            subjectScoreType_list.Add("社會3_科目成績_");
            subjectScoreType_list.Add("社會4_科目成績_");
            subjectScoreType_list.Add("社會5_科目成績_");
            subjectScoreType_list.Add("社會6_科目成績_");
            subjectScoreType_list.Add("藝術1_科目成績_");
            subjectScoreType_list.Add("藝術2_科目成績_");
            subjectScoreType_list.Add("藝術3_科目成績_");
            subjectScoreType_list.Add("藝術4_科目成績_");
            subjectScoreType_list.Add("藝術5_科目成績_");
            subjectScoreType_list.Add("藝術6_科目成績_");
            subjectScoreType_list.Add("健康與體育1_科目成績_");
            subjectScoreType_list.Add("健康與體育2_科目成績_");
            subjectScoreType_list.Add("健康與體育3_科目成績_");
            subjectScoreType_list.Add("健康與體育4_科目成績_");
            subjectScoreType_list.Add("健康與體育5_科目成績_");
            subjectScoreType_list.Add("健康與體育6_科目成績_");
            subjectScoreType_list.Add("綜合活動1_科目成績_");
            subjectScoreType_list.Add("綜合活動2_科目成績_");
            subjectScoreType_list.Add("綜合活動3_科目成績_");
            subjectScoreType_list.Add("綜合活動4_科目成績_");
            subjectScoreType_list.Add("綜合活動5_科目成績_");
            subjectScoreType_list.Add("綜合活動6_科目成績_");
            subjectScoreType_list.Add("藝術與人文1_科目成績_");
            subjectScoreType_list.Add("藝術與人文2_科目成績_");
            subjectScoreType_list.Add("藝術與人文3_科目成績_");
            subjectScoreType_list.Add("藝術與人文4_科目成績_");
            subjectScoreType_list.Add("藝術與人文5_科目成績_");
            subjectScoreType_list.Add("藝術與人文6_科目成績_");
            subjectScoreType_list.Add("自然與生活科技1_科目成績_");
            subjectScoreType_list.Add("自然與生活科技2_科目成績_");
            subjectScoreType_list.Add("自然與生活科技3_科目成績_");
            subjectScoreType_list.Add("自然與生活科技4_科目成績_");
            subjectScoreType_list.Add("自然與生活科技5_科目成績_");
            subjectScoreType_list.Add("自然與生活科技6_科目成績_");

            subjectScoreType_list.Add("彈性課程1_科目成績_");
            subjectScoreType_list.Add("彈性課程2_科目成績_");
            subjectScoreType_list.Add("彈性課程3_科目成績_");
            subjectScoreType_list.Add("彈性課程4_科目成績_");
            subjectScoreType_list.Add("彈性課程5_科目成績_");
            subjectScoreType_list.Add("彈性課程6_科目成績_");
            subjectScoreType_list.Add("彈性課程7_科目成績_");
            subjectScoreType_list.Add("彈性課程8_科目成績_");
            subjectScoreType_list.Add("彈性課程9_科目成績_");
            subjectScoreType_list.Add("彈性課程10_科目成績_");

            #endregion

            #region 整理科目_OO_等第
            //整理科目_OO_等第
            List<string> subjectLevelType_list = new List<string>();

            subjectLevelType_list.Add("語文1_科目等第_");
            subjectLevelType_list.Add("語文2_科目等第_");
            subjectLevelType_list.Add("語文3_科目等第_");
            subjectLevelType_list.Add("語文4_科目等第_");
            subjectLevelType_list.Add("語文5_科目等第_");
            subjectLevelType_list.Add("語文6_科目等第_");
            subjectLevelType_list.Add("數學1_科目等第_");
            subjectLevelType_list.Add("數學2_科目等第_");
            subjectLevelType_list.Add("數學3_科目等第_");
            subjectLevelType_list.Add("數學4_科目等第_");
            subjectLevelType_list.Add("數學5_科目等第_");
            subjectLevelType_list.Add("數學6_科目等第_");
            subjectLevelType_list.Add("生活課程1_科目等第_");
            subjectLevelType_list.Add("生活課程2_科目等第_");
            subjectLevelType_list.Add("生活課程3_科目等第_");
            subjectLevelType_list.Add("生活課程4_科目等第_");
            subjectLevelType_list.Add("生活課程5_科目等第_");
            subjectLevelType_list.Add("生活課程6_科目等第_");
            subjectLevelType_list.Add("自然科學1_科目等第_");
            subjectLevelType_list.Add("自然科學2_科目等第_");
            subjectLevelType_list.Add("自然科學3_科目等第_");
            subjectLevelType_list.Add("自然科學4_科目等第_");
            subjectLevelType_list.Add("自然科學5_科目等第_");
            subjectLevelType_list.Add("自然科學6_科目等第_");
            subjectLevelType_list.Add("科技1_科目等第_");
            subjectLevelType_list.Add("科技2_科目等第_");
            subjectLevelType_list.Add("科技3_科目等第_");
            subjectLevelType_list.Add("科技4_科目等第_");
            subjectLevelType_list.Add("科技5_科目等第_");
            subjectLevelType_list.Add("科技6_科目等第_");
            subjectLevelType_list.Add("社會1_科目等第_");
            subjectLevelType_list.Add("社會2_科目等第_");
            subjectLevelType_list.Add("社會3_科目等第_");
            subjectLevelType_list.Add("社會4_科目等第_");
            subjectLevelType_list.Add("社會5_科目等第_");
            subjectLevelType_list.Add("社會6_科目等第_");
            subjectLevelType_list.Add("藝術1_科目等第_");
            subjectLevelType_list.Add("藝術2_科目等第_");
            subjectLevelType_list.Add("藝術3_科目等第_");
            subjectLevelType_list.Add("藝術4_科目等第_");
            subjectLevelType_list.Add("藝術5_科目等第_");
            subjectLevelType_list.Add("藝術6_科目等第_");
            subjectLevelType_list.Add("健康與體育1_科目等第_");
            subjectLevelType_list.Add("健康與體育2_科目等第_");
            subjectLevelType_list.Add("健康與體育3_科目等第_");
            subjectLevelType_list.Add("健康與體育4_科目等第_");
            subjectLevelType_list.Add("健康與體育5_科目等第_");
            subjectLevelType_list.Add("健康與體育6_科目等第_");
            subjectLevelType_list.Add("綜合活動1_科目等第_");
            subjectLevelType_list.Add("綜合活動2_科目等第_");
            subjectLevelType_list.Add("綜合活動3_科目等第_");
            subjectLevelType_list.Add("綜合活動4_科目等第_");
            subjectLevelType_list.Add("綜合活動5_科目等第_");
            subjectLevelType_list.Add("綜合活動6_科目等第_");
            subjectLevelType_list.Add("藝術與人文1_科目等第_");
            subjectLevelType_list.Add("藝術與人文2_科目等第_");
            subjectLevelType_list.Add("藝術與人文3_科目等第_");
            subjectLevelType_list.Add("藝術與人文4_科目等第_");
            subjectLevelType_list.Add("藝術與人文5_科目等第_");
            subjectLevelType_list.Add("藝術與人文6_科目等第_");
            subjectLevelType_list.Add("自然與生活科技1_科目等第_");
            subjectLevelType_list.Add("自然與生活科技2_科目等第_");
            subjectLevelType_list.Add("自然與生活科技3_科目等第_");
            subjectLevelType_list.Add("自然與生活科技4_科目等第_");
            subjectLevelType_list.Add("自然與生活科技5_科目等第_");
            subjectLevelType_list.Add("自然與生活科技6_科目等第_");

            subjectLevelType_list.Add("彈性課程1_科目等第_");
            subjectLevelType_list.Add("彈性課程2_科目等第_");
            subjectLevelType_list.Add("彈性課程3_科目等第_");
            subjectLevelType_list.Add("彈性課程4_科目等第_");
            subjectLevelType_list.Add("彈性課程5_科目等第_");
            subjectLevelType_list.Add("彈性課程6_科目等第_");
            subjectLevelType_list.Add("彈性課程7_科目等第_");
            subjectLevelType_list.Add("彈性課程8_科目等第_");
            subjectLevelType_list.Add("彈性課程9_科目等第_");
            subjectLevelType_list.Add("彈性課程10_科目等第_");

            #endregion

            // 領域分數、等第 的對照
            Dictionary<string, decimal?> domainScore_dict = new Dictionary<string, decimal?>();
            Dictionary<string, string> domainLevel_dict = new Dictionary<string, string>();

            // 科目分數、等第 的對照
            Dictionary<string, decimal?> subjectScore_dict = new Dictionary<string, decimal?>();
            Dictionary<string, string> subjectLevel_dict = new Dictionary<string, string>();

            // 缺曠節次 、日數 的對照
            Dictionary<string, decimal> arStatistic_dict = new Dictionary<string, decimal>();
            Dictionary<string, decimal> arStatistic_dict_days = new Dictionary<string, decimal>();

            //文字評量(日常生活表現及具體建議、校內外特殊表現)的對照
            Dictionary<string, string> textScore_dict = new Dictionary<string, string>();

            foreach (K12.Data.PeriodMappingInfo var in periodMappingInfos)
            {
                if (var.Type == "一般" & !AbsencePeriod.Contains(var.Name))
                    AbsencePeriod.Add(var.Name);
            }

            int student_counter = 1;

            foreach (string stuID in StudentIDs)
            {



                //把每一筆資料的字典都清乾淨，避免資料汙染
                arStatistic_dict.Clear();
                arStatistic_dict_days.Clear();
                domainScore_dict.Clear();
                domainLevel_dict.Clear();
                subjectScore_dict.Clear();
                subjectLevel_dict.Clear();
                textScore_dict.Clear();

                // 建立缺曠 對照字典
                foreach (string ab in absenceType_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        arStatistic_dict.Add(ab + "日數_" + i, 0);
                    }
                }

                // 建立領域成績 對照字典
                foreach (string dst in domainScoreType_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        domainScore_dict.Add(dst + i, null);
                    }
                }

                // 建立領域等第 對照字典
                foreach (string dlt in domainLevelType_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        domainLevel_dict.Add(dlt + i, null);
                    }
                }

                // 建立科目成績 對照字典
                foreach (string sst in subjectScoreType_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        subjectScore_dict.Add(sst + i, null);

                        if (!table.Columns.Contains(sst + i))
                        {
                            table.Columns.Add(sst + i);
                        }

                    }
                }

                // 建立科目等第 對照字典
                foreach (string slt in subjectLevelType_list)
                {
                    for (int i = 1; i <= 12; i++)
                    {
                        subjectLevel_dict.Add(slt + i, null);

                        if (!table.Columns.Contains(slt + i))
                        {
                            table.Columns.Add(slt + i);
                        }
                    }
                }

                // 建立文字評量 對照字典
                for (int i = 1; i <= 12; i++)
                {
                    textScore_dict.Add("日常生活表現及具體建議_" + i, null);
                    textScore_dict.Add("校內外特殊表現_" + i, null);
                }

                // 存放 各年級與 學年的對照變數
                int schoolyear_grade1 = 0;
                int schoolyear_grade2 = 0;
                int schoolyear_grade3 = 0;
                int schoolyear_grade4 = 0;
                int schoolyear_grade5 = 0;
                int schoolyear_grade6 = 0;

                DataRow row = table.NewRow();

                //學生基本資料
                if (sr_dict.ContainsKey(stuID))
                {
                    DateTime birthday = new DateTime();

                    row["學生姓名"] = sr_dict[stuID].Name;
                    row["學生班級"] = sr_dict[stuID].Class != null ? sr_dict[stuID].Class.Name : "";
                    row["學生性別"] = sr_dict[stuID].Gender;

                    if (sr_dict[stuID].Birthday != null)
                    {
                        birthday = (DateTime)sr_dict[stuID].Birthday;
                        // 轉換出生時間 成 2005/09/06 的格式
                        row["出生日期"] = birthday.ToString("yyyy/MM/dd");
                    }
                    else
                    {
                        row["出生日期"] = "";
                    }

                    row["入學年月"] = "";
                    row["學生身分證字號"] = sr_dict[stuID].IDNumber;
                    row["學號"] = sr_dict[stuID].StudentNumber;

                    PrintStudents.Add(sr_dict[stuID]);
                }

                //學生家長基本資料
                if (spr_dict.ContainsKey(stuID))
                {
                    row["監護人姓名"] = spr_dict[stuID].CustodianName;
                    row["監護人關係"] = spr_dict[stuID].CustodianRelationship;
                    row["監護人行動電話"] = spr_dict[stuID].CustodianPhone;
                    row["父親姓名"] = spr_dict[stuID].FatherName;
                    row["父親行動電話"] = spr_dict[stuID].FatherPhone;
                    row["母親姓名"] = spr_dict[stuID].MotherName; ;
                    row["母親行動電話"] = spr_dict[stuID].MotherPhone;
                }

                //學生聯繫資料(住址)
                if (sar_dict.ContainsKey(stuID))
                {
                    row["戶籍地址"] = sar_dict[stuID].PermanentAddress;
                    row["聯絡地址"] = sar_dict[stuID].MailingAddress;
                }

                //學生聯繫資料(電話)
                if (sphr_dict.ContainsKey(stuID))
                {
                    row["戶籍電話"] = sphr_dict[stuID].Permanent;
                    row["聯絡電話"] = sphr_dict[stuID].Contact;
                }

                //學期歷程
                if (shr_dict.ContainsKey(stuID))
                {
                    foreach (var item in shr_dict[stuID].SemesterHistoryItems)
                    {
                        if (item.GradeYear == 1)
                        {
                            row["學年度1"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade1 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_1"] = item.SchoolDayCount;

                                row["班級1"] = item.ClassName;
                                row["座號1"] = item.SeatNo;
                                row["班導師1"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_2"] = item.SchoolDayCount;

                                row["班級2"] = item.ClassName;
                                row["座號2"] = item.SeatNo;
                                row["班導師2"] = item.Teacher;
                            }
                        }
                        if (item.GradeYear == 2)
                        {
                            row["學年度2"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade2 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_3"] = item.SchoolDayCount;

                                row["班級3"] = item.ClassName;
                                row["座號3"] = item.SeatNo;
                                row["班導師3"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_4"] = item.SchoolDayCount;

                                row["班級4"] = item.ClassName;
                                row["座號4"] = item.SeatNo;
                                row["班導師4"] = item.Teacher;
                            }
                        }
                        if (item.GradeYear == 3)
                        {
                            row["學年度3"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade3 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_5"] = item.SchoolDayCount;

                                row["班級5"] = item.ClassName;
                                row["座號5"] = item.SeatNo;
                                row["班導師5"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_6"] = item.SchoolDayCount;

                                row["班級6"] = item.ClassName;
                                row["座號6"] = item.SeatNo;
                                row["班導師6"] = item.Teacher;
                            }
                        }
                        if (item.GradeYear == 4)
                        {
                            row["學年度4"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade4 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_7"] = item.SchoolDayCount;

                                row["班級7"] = item.ClassName;
                                row["座號7"] = item.SeatNo;
                                row["班導師7"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_8"] = item.SchoolDayCount;

                                row["班級8"] = item.ClassName;
                                row["座號8"] = item.SeatNo;
                                row["班導師8"] = item.Teacher;

                            }
                        }
                        if (item.GradeYear == 5)
                        {
                            row["學年度5"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade5 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_9"] = item.SchoolDayCount;

                                row["班級9"] = item.ClassName;
                                row["座號9"] = item.SeatNo;
                                row["班導師9"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_10"] = item.SchoolDayCount;

                                row["班級10"] = item.ClassName;
                                row["座號10"] = item.SeatNo;
                                row["班導師10"] = item.Teacher;
                            }
                        }
                        if (item.GradeYear == 6)
                        {
                            row["學年度6"] = item.SchoolYear;

                            //為學生的年級與學年配對
                            schoolyear_grade6 = item.SchoolYear;

                            if (item.Semester == 1)
                            {
                                row["應出席日數_11"] = item.SchoolDayCount;

                                row["班級11"] = item.ClassName;
                                row["座號11"] = item.SeatNo;
                                row["班導師11"] = item.Teacher;
                            }
                            else
                            {
                                row["應出席日數_12"] = item.SchoolDayCount;

                                row["班級12"] = item.ClassName;
                                row["座號12"] = item.SeatNo;
                                row["班導師12"] = item.Teacher;
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
                                        // 假如該缺曠結束 沒有在 節次管理 設定為一般則跳過
                                        if (!AbsencePeriod.Contains(detail.Period))
                                        {
                                            continue;
                                        }
                                        if (arStatistic_dict.ContainsKey(detail.AbsenceType + "日數_" + (grade * 2 - 1)))
                                        {

                                            //加一節，整學期節次與日數的關係，再最後再結算
                                            arStatistic_dict[detail.AbsenceType + "日數_" + (grade * 2 - 1)] += 1;

                                            // 不管是啥缺席，缺席總日數都加一節
                                            arStatistic_dict["缺席總日數_" + (grade * 2 - 1)] += 1;

                                        }
                                    }
                                }
                                else
                                {
                                    foreach (var detail in ar.PeriodDetail)
                                    {

                                        // 假如該缺曠結束 沒有在 節次管理 設定為一般則跳過
                                        if (!AbsencePeriod.Contains(detail.Period))
                                        {
                                            continue;
                                        }

                                        if (arStatistic_dict.ContainsKey(detail.AbsenceType + "日數_" + grade * 2))
                                        {

                                            //加一節，整學期節次與日數的關係，再最後再結算
                                            arStatistic_dict[detail.AbsenceType + "日數_" + grade * 2] += 1;

                                            // 不管是啥缺席，缺席總日數都加一節
                                            arStatistic_dict["缺席總日數_" + (grade * 2)] += 1;
                                        }
                                    }
                                }
                            }
                        }

                    }

                    foreach (string key in arStatistic_dict.Keys)
                    {
                        arStatistic_dict_days.Add(key, arStatistic_dict[key]);
                    }

                    //真正的填值，填日數，所以要做節次轉換
                    foreach (string key in arStatistic_dict_days.Keys)
                    {
                        // 一天幾節課
                        int periodsADay = AbsencePeriod.Count;

                        row[key] = Math.Round(arStatistic_dict_days[key] / periodsADay, 2);
                    }
                }


                // 學期成績(包含領域、科目)      

                //一般科目 科目名稱與科目編號對照表
                Dictionary<string, Dictionary<string, int>> SubjectCourseDict = new Dictionary<string, Dictionary<string, int>>()
                {
                    { "語文", new Dictionary<string, int>() }
                    , { "數學", new Dictionary<string, int>() }
                    , { "生活課程", new Dictionary<string, int>() }
                    , { "自然科學", new Dictionary<string, int>() }
                    , { "科技", new Dictionary<string, int>() }
                    , { "社會", new Dictionary<string, int>() }
                    , { "藝術", new Dictionary<string, int>() }
                    , { "健康與體育", new Dictionary<string, int>() }
                    , { "綜合活動", new Dictionary<string, int>() }
                    , { "藝術與人文", new Dictionary<string, int>() }
                    , { "自然與生活科技", new Dictionary<string, int>() }
                };

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

                                            row[subjectscore.Value.Domain + subjectCourseCount + "_科目名稱"] = subjectscore.Value.Subject;

                                            SubjectCourseDict[subjectscore.Value.Domain].Add(subjectscore.Value.Subject, subjectCourseCount); 
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (isExceed)
                    {
                        MessageBox.Show("科目數超過報表變數可支援數量，超過的將不會顯示在學籍表中");
                    }
                }


                // 彈性課程 科目名稱 與彈性課程編號的對照
                Dictionary<string, int> AlternativeCourseDict = new Dictionary<string, int>();

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

                                        // 目前僅先支援 一個學生在六年之中有 10個 彈性課程
                                        if (AlternativeCourse > 10)
                                        {
                                            MessageBox.Show("彈性科目數超過可支援數量，超過的將不會顯示在學籍表中");
                                            break;
                                        }

                                        row["彈性課程" + AlternativeCourse + "_科目名稱"] = subjectscore.Value.Subject;

                                        AlternativeCourseDict.Add(subjectscore.Value.Subject, AlternativeCourse);
                                    }
                                }
                            }
                        }
                    }
                }

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
                                        //紀錄成績
                                        if (domainScore_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2 - 1)))
                                        {
                                            domainScore_dict["領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2 - 1)] = domainscore.Value.Score;
                                        }

                                        //換算等第
                                        if (domainLevel_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2 - 1)))
                                        {
                                            domainLevel_dict["領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2 - 1)] = ScoreTolevel(domainscore.Value.Score);
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

                                                //紀錄成績
                                                if (subjectScore_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2 - 1)))
                                                {

                                                    subjectScore_dict["彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                }

                                                //紀錄等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2 - 1)))
                                                {
                                                    subjectLevel_dict["彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2 - 1)] = ScoreTolevel(subjectscore.Value.Score);
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

                                                    //紀錄成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2 - 1)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2 - 1)] = subjectscore.Value.Score;
                                                    }

                                                    //換算等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2 - 1)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2 - 1)] = ScoreTolevel(subjectscore.Value.Score);
                                                    }
                                                } 
                                            }

                                        }

                                    }

                                    //學期學習領域(七大)成績(不包括彈性課程成績)
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_學習領域總成績_成績_" + (grade * 2 - 1)))
                                    {
                                        domainScore_dict["領域_學習領域總成績_成績_" + (grade * 2 - 1)] = jssr.LearnDomainScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_學習領域總成績_等第_" + (grade * 2 - 1)))
                                    {
                                        domainLevel_dict["領域_學習領域總成績_等第_" + (grade * 2 - 1)] = ScoreTolevel(jssr.LearnDomainScore);
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
                                        domainLevel_dict["領域_課程學習成績_等第_" + (grade * 2 - 1)] = ScoreTolevel(jssr.CourseLearnScore);
                                    }


                                }
                                else
                                {
                                    //領域
                                    foreach (var domainscore in jssr.Domains)
                                    {
                                        if (domainScore_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2)))
                                        {
                                            domainScore_dict["領域_" + domainscore.Value.Domain + "_成績_" + (grade * 2)] = domainscore.Value.Score;
                                        }

                                        //換算等第
                                        if (domainLevel_dict.ContainsKey("領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2)))
                                        {
                                            domainLevel_dict["領域_" + domainscore.Value.Domain + "_等第_" + (grade * 2)] = ScoreTolevel(domainscore.Value.Score);
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
                                                //紀錄成績
                                                if (subjectScore_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2)))
                                                {
                                                    subjectScore_dict["彈性課程" + AlternativeCourse + "_科目成績_" + (grade * 2)] = subjectscore.Value.Score;
                                                }

                                                //紀錄等第
                                                if (subjectLevel_dict.ContainsKey("彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2)))
                                                {
                                                    subjectLevel_dict["彈性課程" + AlternativeCourse + "_科目等第_" + (grade * 2)] = ScoreTolevel(subjectscore.Value.Score);
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
                                                    //紀錄成績
                                                    if (subjectScore_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2)))
                                                    {
                                                        subjectScore_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目成績_" + (grade * 2)] = subjectscore.Value.Score;
                                                    }

                                                    //換算等第
                                                    if (subjectLevel_dict.ContainsKey(subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2)))
                                                    {
                                                        subjectLevel_dict[subjectscore.Value.Domain + SubjectCourseNum + "_科目等第_" + (grade * 2)] = ScoreTolevel(subjectscore.Value.Score);
                                                    }
                                                } 
                                            }
                                        }

                                    }

                                    //學期學習領域(七大)成績
                                    //紀錄成績
                                    if (domainScore_dict.ContainsKey("領域_學習領域總成績_成績_" + (grade * 2)))
                                    {
                                        domainScore_dict["領域_學習領域總成績_成績_" + (grade * 2)] = jssr.LearnDomainScore;
                                    }

                                    //換算等第
                                    if (domainLevel_dict.ContainsKey("領域_學習領域總成績_等第_" + (grade * 2)))
                                    {
                                        domainLevel_dict["領域_學習領域總成績_等第_" + (grade * 2)] = ScoreTolevel(jssr.LearnDomainScore);
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
                                        domainLevel_dict["領域_課程學習成績_等第_" + (grade * 2)] = ScoreTolevel(jssr.CourseLearnScore);
                                    }


                                }
                            }
                        }

                    }

                    // 填領域分數
                    foreach (string key in domainScore_dict.Keys)
                    {
                        row[key] = domainScore_dict[key];
                    }

                    // 填領域等第
                    foreach (string key in domainLevel_dict.Keys)
                    {
                        row[key] = domainLevel_dict[key];
                    }

                    // 填科目分數
                    foreach (string key in subjectScore_dict.Keys)
                    {
                        row[key] = subjectScore_dict[key];
                    }

                    // 填科目等第
                    foreach (string key in subjectLevel_dict.Keys)
                    {
                        row[key] = subjectLevel_dict[key];
                    }


                }

                //畢業分數
                if (gsr_dict.ContainsKey(stuID))
                {
                    row["畢業總成績_平均"] = gsr_dict[stuID].LearnDomainScore;
                    row["畢業總成績_等第"] = ScoreTolevel(gsr_dict[stuID].LearnDomainScore);

                    // 60 分 就可以 准予畢業
                    row["准予畢業"] = gsr_dict[stuID].LearnDomainScore > 60 ? "■" : "□";
                    row["發給修業證書"] = gsr_dict[stuID].LearnDomainScore > 60 ? "□" : "■";
                }

                // 異動資料
                if (urr_dict.ContainsKey(stuID))
                {
                    int updateRecordCount = 1;

                    urr_dict[stuID].Sort((x, y) => { return x.UpdateDate.CompareTo(y.UpdateDate); });

                    foreach (K12.Data.UpdateRecordRecord urr in urr_dict[stuID])
                    {
                        // 新生異動為1 ，且理論上 一個人 會有1筆新生異動
                        if (urr.UpdateCode == "1")
                        {
                            DateTime enterday = new DateTime();

                            enterday = DateTime.Parse(urr.UpdateDate);
                            // 轉換入學時間 成 2005/09/06 的格式
                            row["入學年月"] = enterday.ToString("yyyy/MM");

                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = ""; // 2019/03/15 怡芬說 新生異動 校名 空白
                            row["異動紀錄" + updateRecordCount + "_類別"] = "新生";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 畢業（2）
                        if (urr.UpdateCode == "2")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "畢業";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        //  當異動為 轉入 (3) 、轉出(4)時 要依序顯示在 報表上
                        if (urr.UpdateCode == "3" || urr.UpdateCode == "4")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = urr.Attributes["ImportExportSchool"]; // 取得異動校名的方法
                            row["異動紀錄" + updateRecordCount + "_類別"] = urr.UpdateCode == "3" ? "轉入" : "轉出";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;
                            updateRecordCount++;
                        }

                        //因為無法新增休學（5）及續讀（8）這兩種異動，所以這邊先不處理這兩種異動

                        // 復學（6）
                        if (urr.UpdateCode == "6")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "復學";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 中輟（7）
                        if (urr.UpdateCode == "7")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "中輟";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 更正學籍（9）
                        if (urr.UpdateCode == "9")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "更正學籍";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 延長修業年限（10）
                        if (urr.UpdateCode == "10")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "延長修業年限";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }

                        // 死亡（11）
                        if (urr.UpdateCode == "11")
                        {
                            row["異動紀錄" + updateRecordCount + "_日期"] = urr.UpdateDate;
                            row["異動紀錄" + updateRecordCount + "_校名"] = "";
                            row["異動紀錄" + updateRecordCount + "_類別"] = "死亡";
                            row["異動紀錄" + updateRecordCount + "_核准日期"] = urr.ADDate;
                            row["異動紀錄" + updateRecordCount + "_核准文號"] = urr.ADNumber;

                            updateRecordCount++;
                        }
                    }
                }

                // 日常生活表現、校內外特殊表現
                if (msr_dict.ContainsKey(stuID))
                {
                    for (int grade = 1; grade <= 6; grade++)
                    {
                        foreach (var msr in msr_dict[stuID])
                        {
                            if (msr.SchoolYear == schoolyear_grade_dict[grade])
                            {
                                if (msr.Semester == 1)
                                {
                                    if (textScore_dict.ContainsKey("日常生活表現及具體建議_" + (grade * 2 - 1)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("DailyLifeRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["日常生活表現及具體建議_" + (grade * 2 - 1)] = msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"].Value;
                                            }
                                        }
                                    }
                                    if (textScore_dict.ContainsKey("校內外特殊表現_" + (grade * 2 - 1)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("OtherRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["校內外特殊表現_" + (grade * 2 - 1)] = msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"].Value;
                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    if (textScore_dict.ContainsKey("日常生活表現及具體建議_" + (grade * 2)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("DailyLifeRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["日常生活表現及具體建議_" + (grade * 2)] = msr.TextScore.SelectSingleNode("DailyLifeRecommend").Attributes["Description"].Value;
                                            }
                                        }
                                    }
                                    if (textScore_dict.ContainsKey("校內外特殊表現_" + (grade * 2)))
                                    {
                                        if (msr.TextScore.SelectSingleNode("OtherRecommend") != null)
                                        {
                                            if (msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"] != null)
                                            {
                                                textScore_dict["校內外特殊表現_" + (grade * 2)] = msr.TextScore.SelectSingleNode("OtherRecommend").Attributes["Description"].Value;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //填值
                    foreach (string key in textScore_dict.Keys)
                    {
                        row[key] = textScore_dict[key];
                    }
                }

                table.Rows.Add(row);

                //回報進度
                int percent = ((student_counter * 100 / StudentIDs.Count));

                MasterWorker.ReportProgress(percent, "學生學籍表產生中...進行到第" + (student_counter) + "/" + StudentIDs.Count + "學生");

                student_counter++;
            }

            //選擇 目前的樣板
            document = new Document(Preference.Template.GetStream());
            //document = Preference.Template.ToDocument();

            //執行 合併列印
            document.MailMerge.Execute(table);

            // 最終產物 .doc
            e.Result = document;

            Feedback("列印完成", -1);
        }

        // 換算分數 與 等第用
        private string ScoreTolevel(decimal? d)
        {
            string level = "";
            if (d >= 90)
            {
                level = "優";
            }
            else if (d >= 80 && d < 90)
            {
                level = "甲";
            }
            else if (d >= 70 && d < 80)
            {
                level = "乙";
            }
            else if (d >= 60 && d < 70)
            {
                level = "丙";
            }
            else if (d < 60)
            {
                level = "丁";
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
                MsgBox.Show(e.Error.Message);
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
                Util.Save(doc, "學籍表", Preference.ConvertToPDF);
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
                        //string fPath = fbd.SelectedPath + "\\" + fileName + ".pdf";

                        string fPath = fbdPath + "\\" + fileName + ".pdf";

                        FileInfo fi = new FileInfo(fPath);

                        DirectoryInfo folder = new DirectoryInfo(Path.Combine(fi.DirectoryName, Path.GetRandomFileName()));
                        if (!folder.Exists) folder.Create();

                        FileInfo fileinfo = new FileInfo(Path.Combine(folder.FullName, fi.Name));

                        string XmlFileName = fileinfo.FullName.Substring(0, fileinfo.FullName.Length - fileinfo.Extension.Length) + ".xml";
                        string PDFFileName = fileinfo.FullName.Substring(0, fileinfo.FullName.Length - fileinfo.Extension.Length) + ".pdf";

                        document.Save(XmlFileName, Aspose.Words.SaveFormat.Pdf);

                        Aspose.Pdf.Generator.Pdf pdf1 = new Aspose.Pdf.Generator.Pdf();

                        pdf1.BindXML(XmlFileName, null);
                        pdf1.Save(PDFFileName);

                        if (File.Exists(fPath))
                            File.Delete(Path.Combine(fi.DirectoryName, fi.Name));

                        File.Move(PDFFileName, fPath);
                        folder.Delete(true);

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
            ReportTemplate defaultTemplate = new ReportTemplate(Prc.康橋國小學籍表_樣板_, TemplateType.Word);
            TemplateSettingForm form = new TemplateSettingForm(Preference.Template, defaultTemplate);
            form.DefaultFileName = "學籍表(樣版).doc";

            if (form.ShowDialog() == DialogResult.OK)
            {
                Preference.Template = (form.Template == defaultTemplate) ? null : form.Template;
                Preference.Save();
            }
        }

        //供使用者下載學籍表合併欄位總表
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //宣告產生的報表
            Aspose.Words.Document document = new Aspose.Words.Document();

            document = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.康橋新竹國小學籍表功能變數));

            System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "學籍表合併欄位總表" + ".doc";
            sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }
            }
        }
    }
}
