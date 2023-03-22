using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using JHEvaluation.ScoreCalculation;
using JHSchool.Data;
using Aspose.Words;
using Aspose.Words.Tables;
using FISCA.Presentation.Controls;
using Campus.Rating;
using System.Globalization;
using FISCA.Data;
using System.Data;
using System.IO;

namespace JHEvaluation.StudentScoreSummaryReport
{
    internal static class Util
    {
        /// <summary>
        /// 英文日期格式。
        /// </summary>
        public const string EnglishFormat = "MMMM dd, yyyy";

        public static CultureInfo USCulture = new CultureInfo("en-us");

        public static void DisableControls(Control topControl)
        {
            ChangeControlsStatus(topControl, false);
        }

        public static void EnableControls(Control topControl)
        {
            ChangeControlsStatus(topControl, true);
        }

        private static void ChangeControlsStatus(Control topControl, bool status)
        {
            foreach (Control each in topControl.Controls)
            {
                string tag = each.Tag + "";
                if (tag.ToUpper() == "StatusVarying".ToUpper())
                {
                    each.Enabled = status;
                }

                if (each.Controls.Count > 0)
                    ChangeControlsStatus(each, status);
            }
        }

        /// <summary>
        /// 將學生編號轉換成 SCStudent 物件。
        /// </summary>
        /// <remarks>使用指定的學生編號，向 DAL 取得 VO 後轉換成 SCStudent 物件。</remarks>
        public static List<ReportStudent> ToReportStudent(this IEnumerable<string> studentIDs)
        {
            List<ReportStudent> students = new List<ReportStudent>();
            foreach (JHStudentRecord each in JHStudent.SelectByIDs(studentIDs))
                students.Add(new ReportStudent(each));
            return students;
        }

        /// <summary>
        /// 取得全部學生的資料(只包含一般、輟學)。
        /// </summary>
        /// <returns></returns>
        public static List<ReportStudent> GetAllStudents()
        {
            List<ReportStudent> students = new List<ReportStudent>();
            foreach (JHStudentRecord each in JHStudent.SelectAll())
                students.Add(new ReportStudent(each));

            return students;
        }

        /// <summary>
        /// 取得全部學生的資料(只包含一般、輟學)。
        /// </summary>
        /// <returns></returns>
        public static List<ReportStudent> GetRatingStudents(List<ReportStudent> rs)
        {
            List<ReportStudent> students = new List<ReportStudent>();
            foreach (ReportStudent each in rs)
            {
                if (each.StudentStatus == K12.Data.StudentRecord.StudentStatus.一般 ||
                    each.StudentStatus == K12.Data.StudentRecord.StudentStatus.輟學)
                    students.Add(each);
            }
            return students;
        }

        /// <summary>
        /// 轉型成 StudentScore 集合。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static List<StudentScore> ToSC(this IEnumerable<ReportStudent> students)
        {
            List<StudentScore> stus = new List<StudentScore>();
            foreach (ReportStudent each in students)
                stus.Add(each);
            return stus;
        }

        /// <summary>
        /// 轉型成 StudentScore。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static List<ReportStudent> ToSS(this IEnumerable<StudentScore> students)
        {
            List<ReportStudent> stus = new List<ReportStudent>();
            foreach (StudentScore each in students)
                stus.Add(each as ReportStudent);
            return stus;
        }

        /// <summary>
        /// 將指定的 SCStudent 集合轉換成 ID->SCStudent 對照。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static Dictionary<string, StudentScore> ToDictionary(this IEnumerable<StudentScore> students)
        {
            Dictionary<string, StudentScore> dicstuds = new Dictionary<string, StudentScore>();
            foreach (StudentScore each in students)
                dicstuds.Add(each.Id, each);
            return dicstuds;
        }

        public static Dictionary<string, ReportStudent> ToDictionary(this IEnumerable<ReportStudent> students)
        {
            Dictionary<string, ReportStudent> dicstuds = new Dictionary<string, ReportStudent>();
            foreach (ReportStudent each in students)
                dicstuds.Add(each.Id, each);
            return dicstuds;
        }

        /// <summary>
        /// 將 SCStudent 集合轉換成編號的集合。
        /// </summary>
        /// <param name="students"></param>
        /// <returns></returns>
        public static List<string> ToKeys(this IEnumerable<StudentScore> students)
        {
            List<string> keys = new List<string>();
            foreach (StudentScore each in students)
                keys.Add(each.Id);
            return keys;
        }

        public static void Save(Document doc, string fileName, bool convertToPDF)
        {
            try
            {
                if (doc != null)
                {
                    string path = "";
                    if (convertToPDF)
                    {
                        path = $"{Application.StartupPath}\\Reports\\{fileName}.pdf";
                    }
                    else
                    {
                        path = $"{Application.StartupPath}\\Reports\\{fileName}.docx";
                    }

                    int i = 1;
                    while (File.Exists(path))
                    {
                        string newPath = $"{Path.GetDirectoryName(path)}\\{fileName}{i++}{Path.GetExtension(path)}";
                        path = newPath;
                    }

                    doc.Save(path, convertToPDF ? SaveFormat.Pdf : SaveFormat.Docx);

                    DialogResult dialogResult = MessageBox.Show($"{path}\n{fileName}產生完成，是否立即開啟？", "訊息", MessageBoxButtons.YesNo);

                    if (DialogResult.Yes == dialogResult)
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("儲存失敗。" + ex.Message);
                return;
            }
        }

        public static string GetGradeyearString(string gradeYear)
        {
            switch (gradeYear)
            {
                case "1":
                    return "一";
                case "2":
                    return "二";
                case "3":
                    return "三";
                case "4":
                    return "四";
                case "5":
                    return "五";
                case "6":
                    return "六";
                case "7":
                    return "七";
                case "8":
                    return "八";
                case "9":
                    return "九";
                case "10":
                    return "十";
                case "11":
                    return "十一";
                case "12":
                    return "十二";
                default:
                    return gradeYear;
            }
        }

        /// <summary>
        /// 取得下一個 Cell 的 Paragraph。
        /// </summary>
        public static Paragraph NextCell(Paragraph para)
        {
            if (para.ParentNode is Cell)
            {
                Cell cell = para.ParentNode.NextSibling as Cell;

                if (cell == null) return null;

                if (cell.Paragraphs.Count <= 0)
                    cell.Paragraphs.Add(new Paragraph(para.Document));

                return cell.Paragraphs[0];
            }
            else
                return null;
        }

        /// <summary>
        /// 取得前一個 Cell 的 Paragraph。
        /// </summary>
        /// <param name="para"></param>
        /// <returns></returns>
        public static Paragraph PreviousCell(Paragraph para)
        {
            if (para.ParentNode is Cell)
            {
                Cell cell = para.ParentNode.PreviousSibling as Cell;

                if (cell == null) return null;

                if (cell.Paragraphs.Count <= 0)
                    cell.Paragraphs.Add(new Paragraph(para.Document));

                return cell.Paragraphs[0];
            }
            else
                return null;
        }

        public static void Write(this Cell cell, DocumentBuilder builder, string text)
        {
            if (cell.Paragraphs.Count <= 0)
                cell.Paragraphs.Add(new Paragraph(cell.Document));

            builder.MoveTo(cell.Paragraphs[0]);
            builder.Write(text);
        }

        public static List<RatingScope<ReportStudent>> ToGradeYearScopes(this IEnumerable<ReportStudent> students)
        {
            Dictionary<string, RatingScope<ReportStudent>> scopes = new Dictionary<string, RatingScope<ReportStudent>>();

            foreach (ReportStudent each in students)
            {
                string gradeYear = string.Empty;

                if (!string.IsNullOrEmpty(each.RefClassID))
                {
                    int? gy = JHClass.SelectByID(each.RefClassID).GradeYear;
                    if (gy.HasValue) gradeYear = gy.Value.ToString();
                }

                if (!scopes.ContainsKey(gradeYear))
                    scopes.Add(gradeYear, new RatingScope<ReportStudent>(gradeYear, "年排名"));

                scopes[gradeYear].Add(each);
            }

            return new List<RatingScope<ReportStudent>>(scopes.Values);
        }

        //private static List<string> subjOrder = new List<string>(new string[] { "國語文", "語文", "國文", "英文", "英語", "數學", "社會", "歷史", "公民", "地理", "藝術與人文", "自然與生活科技", "理化", "生物", "健康與體育", "綜合活動", "學習領域", "彈性課程" });
        //public static int SortSubject(RowHeader x, RowHeader y)
        //{
        //    int ix = subjOrder.IndexOf(x.Subject);
        //    int iy = subjOrder.IndexOf(y.Subject);

        //    if (ix >= 0 && iy >= 0) //如果都有找到位置。
        //        return ix.CompareTo(iy);
        //    else if (ix >= 0)
        //        return -1;
        //    else if (iy >= 0)
        //        return 1;
        //    else
        //        return x.Subject.CompareTo(y.Subject);
        //}

        //public static int SortDomain(RowHeader x, RowHeader y)
        //{
        //    int ix = subjOrder.IndexOf(x.Domain);
        //    int iy = subjOrder.IndexOf(y.Domain);

        //    if (ix >= 0 && iy >= 0) //如果都有找到位置。
        //        return ix.CompareTo(iy);
        //    else if (ix >= 0)
        //        return -1;
        //    else if (iy >= 0)
        //        return 1;
        //    else
        //        return x.Domain.CompareTo(y.Domain);
        //}

        public static string GetDegree(decimal score)
        {
            if (score >= 90) return "優";
            else if (score >= 80) return "甲";
            else if (score >= 70) return "乙";
            else if (score >= 60) return "丙";
            else return "丁";
        }

        public static string GetDegreeEnglish(decimal score)
        {
            if (score >= 90) return "A";
            else if (score >= 80) return "B";
            else if (score >= 70) return "C";
            else if (score >= 60) return "D";
            else return "E";
        }

        /// <summary>
        ///  例：一般:曠課,事假,病假;集合:曠課,事假,公假
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public static string PeriodOptionsToString(this Dictionary<string, List<string>> setting)
        {
            StringBuilder builder = new StringBuilder();
            foreach (KeyValuePair<string, List<string>> eachType in setting)
            {
                builder.Append(eachType.Key + ":");

                foreach (string each in eachType.Value)
                    builder.Append(each + ",");

                builder.Append(";");
            }
            return builder.ToString();
        }

        /// <summary>
        /// 例：一般:曠課,事假,病假;集合:曠課,事假,公假
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public static Dictionary<string, List<string>> PeriodOptionsFromString(this string setting)
        {
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>();
            //以「;」分割每一個節次類別。
            foreach (string eachType in setting.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                //以「:」分割類別名稱與資料。
                string[] arrTypeData = eachType.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                string typeName, typeData;
                if (arrTypeData.Length >= 2)
                {
                    typeName = arrTypeData[0];
                    typeData = arrTypeData[1];
                }
                else
                    continue;

                result.Add(typeName, new List<string>());
                //以「,」分割每個資料項。
                foreach (string eachEntry in typeData.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    result[typeName].Add(eachEntry);
            }
            return result;
        }

        /// <summary>
        /// 透過學生編號,取得特定學年度學期服務學習時數
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// /// <returns></returns>
        public static Dictionary<string, Dictionary<string, string>> GetServiceLearningDetail(List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, string>> retVal = new Dictionary<string, Dictionary<string, string>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,school_year,semester,sum(hours) as hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') group by ref_student_id,school_year,semester order by school_year,semester;";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["ref_student_id"].ToString();
                    string key1 = dr["school_year"].ToString() + "_" + dr["semester"].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, new Dictionary<string, string>());

                    if (!retVal[sid].ContainsKey(key1))
                        retVal[sid].Add(key1, "0");

                    retVal[sid][key1] = dr["hours"].ToString();
                }
            }
            return retVal;
        }

        /// <summary>
        /// 服務學時數暫存使用
        /// </summary>
        public static Dictionary<string, Dictionary<string, string>> _SLRDict = new Dictionary<string, Dictionary<string, string>>();

        /// <summary>
        /// 取得科目對照
        /// </summary>
        /// <returns></returns>
        public static List<string> GetSubjectList()
        {
            List<string> value = new List<string>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = @"
WITH    Subject_mapping AS 
(
SELECT	
	unnest(xpath('//Subjects/Subject/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS Subject_name
	, unnest(xpath('//Subjects/Subject/@EnglishName',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS Subject_EnglishName
FROM  
    list 
WHERE name  ='JHEvaluation_Subject_Ordinal'
)SELECT
		replace (Subject_name ,'&amp;amp;','&') AS subject_name
		,replace (Subject_EnglishName ,'&amp;amp;','&') AS Subject_EnglishName
	FROM  Subject_mapping";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string name = dr["subject_name"].ToString();
                    if (!value.Contains(name))
                        value.Add(name);
                }

            }
            catch (Exception ex) { }

            return value;
        }


        /// <summary>
        /// 取得領域List
        /// </summary>
        /// <returns></returns>
        public static List<string> GetDomainList()
        {
            List<string> value = new List<string>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = @"
WITH    domain_mapping AS 
(
SELECT	
	unnest(xpath('//Domains/Domain/@Name',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS domain_name
	, unnest(xpath('//Domains/Domain/@EnglishName',  xmlparse(content replace(replace(content ,'&lt;','<'),'&gt;','>'))))::text AS domain_EnglishName
FROM  
    list 
WHERE name  ='JHEvaluation_Subject_Ordinal'
)SELECT
		replace (domain_name ,'&amp;amp;','&') AS domain_name
		,replace (domain_EnglishName ,'&amp;amp;','&') AS domain_EnglishName
	FROM  domain_mapping";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string name = dr["domain_name"].ToString();
                    if (!value.Contains(name))
                        value.Add(name);
                }
                if (!value.Contains("彈性課程"))
                    value.Add("彈性課程");

                if (!value.Contains("語文"))
                    value.Add("語文");

                if (!value.Contains("生活課程"))
                    value.Add("生活課程");
            }
            catch (Exception ex) { }

            return value;
        }

        /// <summary>
        /// 取得節次對照表
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetPeriodMappingDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = @"
SELECT
    array_to_string(xpath('//Period/@Name', each_period.period), '')::text as name
	, array_to_string(xpath('//Period/@Type', each_period.period), '')::text as type
	, row_number() OVER() as period_order
FROM(
    SELECT unnest(xpath('//Periods/Period', xmlparse(content content))) as period
    FROM list
    WHERE name = '節次對照表'
) as each_period";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string name = dr["name"].ToString();
                    string type = dr["type"].ToString();
                    if (!value.ContainsKey(name))
                        value.Add(name, type);
                }

            }
            catch (Exception ex) { }

            return value;
        }
        /// <summary>
        /// 獎懲暫存使用
        /// </summary>
        public static Dictionary<string, Dictionary<string, Dictionary<string, string>>> _DisciplineDict = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

        /// <summary>
        /// 透過學生編號,取得獎懲總計
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// /// <returns></returns>
        public static Dictionary<string, Dictionary<string, Dictionary<string, string>>> GetDisciplineDetail(List<string> StudentIDList)
        {
            //key=id
            //key1=學年度_學期
            //key2= 獎懲type
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> retVal = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = @"WITH data AS(
SELECT 
	ref_student_id 
	, school_year 
	, semester 
	,  ('0'||array_to_string(xpath('//Discipline/Merit/@A', xmlparse(content detail)), '')::text)::decimal  as 大功 
	,  ('0'||array_to_string(xpath('//Discipline/Merit/@B', xmlparse(content detail)), '')::text)::decimal  as 小功 
	,  ('0'||array_to_string(xpath('//Discipline/Merit/@C', xmlparse(content detail)), '')::text)::decimal as 嘉獎 
	,  ('0'||array_to_string(xpath('//Discipline/Demerit/@A', xmlparse(content detail)), '')::text)::decimal  as 大過 
	,  ('0'||array_to_string(xpath('//Discipline/Demerit/@B', xmlparse(content detail)), '')::text)::decimal  as 小過 
	,  ('0'||array_to_string(xpath('//Discipline/Demerit/@C', xmlparse(content detail)), '')::text)::decimal  as 警告 
	, array_to_string(xpath('//Discipline/Demerit/@Cleared', xmlparse(content detail)), '')::text  as 已銷過 
FROM discipline 
WHERE ref_student_id IN ('" + string.Join("','", StudentIDList.ToArray()) + @"')
)
SELECT 
	ref_student_id
	, school_year
	, semester
	, SUM(大功) AS 大功統計 
	, SUM(小功) AS 小功統計 
	, SUM(嘉獎) AS 嘉獎統計 
	, SUM(大過) AS 大過統計 
	, SUM(小過) AS 小過統計 
	, SUM(警告) AS 警告統計 
FROM data 
WHERE 已銷過  <> '是'
GROUP BY  school_year,semester,ref_student_id
ORDER BY school_year,semester";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["ref_student_id"].ToString();
                    string key1 = dr["school_year"].ToString() + "_" + dr["semester"].ToString();

                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, new Dictionary<string, Dictionary<string, string>>());

                    if (!retVal[sid].ContainsKey(key1))
                        retVal[sid].Add(key1, new Dictionary<string, string>());

                    if (!retVal[sid][key1].ContainsKey("大功"))
                        retVal[sid][key1].Add("大功", dr["大功統計"].ToString());
                    if (!retVal[sid][key1].ContainsKey("小功"))
                        retVal[sid][key1].Add("小功", dr["小功統計"].ToString());
                    if (!retVal[sid][key1].ContainsKey("嘉獎"))
                        retVal[sid][key1].Add("嘉獎", dr["嘉獎統計"].ToString());
                    if (!retVal[sid][key1].ContainsKey("大過"))
                        retVal[sid][key1].Add("大過", dr["大過統計"].ToString());
                    if (!retVal[sid][key1].ContainsKey("小過"))
                        retVal[sid][key1].Add("小過", dr["小過統計"].ToString());
                    if (!retVal[sid][key1].ContainsKey("警告"))
                        retVal[sid][key1].Add("警告", dr["警告統計"].ToString());
                }
            }
            return retVal;
        }
    }
}
