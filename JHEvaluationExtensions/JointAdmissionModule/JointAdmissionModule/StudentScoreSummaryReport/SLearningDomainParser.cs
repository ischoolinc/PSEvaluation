﻿using System;
using System.Collections.Generic;
using System.Text;
using JHEvaluation.ScoreCalculation;
using JHEvaluation.ScoreCalculation.ScoreStruct;

namespace JointAdmissionModule.StudentScoreSummaryReport
{
    /// <summary>
    /// 代表單一學期的學習領域成績排名成績解析邏輯。
    /// 回傳 Null 不列入排名
    /// </summary>
    internal class SLearningDomainParser : Campus.Rating.IScoreParser<ReportStudent>
    {
        private SemesterData Semester { get; set; }

        public SLearningDomainParser(int gradeYear, int semester)
        {
            Semester = new SemesterData(gradeYear, 0, semester);
        }

        #region IScoreParser<ReportStudent> 成員

        public decimal? GetScore(ReportStudent student)
        {
            SemesterScore score = null;
            
            foreach (SemesterData each in student.SHistory.GetGradeYearSemester())
            {
                //// 處理轉入生之前不列入排名
                //if (student.LastEnterSchoolyear.HasValue && student.LastEnterSemester.HasValue && student.LastEnterGradeYear.HasValue)
                //{
                //    if (each.GradeYear < student.LastEnterGradeYear.Value)
                //        continue;
                //    else
                //    {
                //        if (each.GradeYear == student.LastEnterGradeYear.Value && each.Semester < student.LastEnterSemester.Value)
                //            continue;
                //    }

                //}

                //// 處理轉入生之前不列入排名
                //if (student.LastEnterSchoolyear.HasValue && student.LastEnterSemester.HasValue && student.LastEnterGradeYear.HasValue)
                //{
                //    if (each.SchoolYear < student.LastEnterSchoolyear.Value)
                //        continue;
                //    else
                //    {
                //        if (each.SchoolYear == student.LastEnterSchoolyear.Value && each.Semester < student.LastEnterSemester.Value)
                //            continue;
                //    }

                //}

                SemesterData gysemester = new SemesterData(each.GradeYear, 0, each.Semester);
                if (gysemester == Semester)
                {
                    SemesterData sd = new SemesterData(0, each.SchoolYear, each.Semester);

                    if (student.SemestersScore.Contains(sd))
                        score = student.SemestersScore[sd];

                    break;
                }
            }


            //if (score == null)
            //    return null;
            //else
            //    return score.LearnDomainScore;

            if (score == null)
                return 0;
            else
                if (score.LearnDomainScore.HasValue)
                {
                    //return score.LearnDomainScore.Value;
                    // 修改讀取成績計算規則
                    if(student.CalculationRule==null)
                        return score.LearnDomainScore.Value;
                    else
                        return student.CalculationRule.ParseLearnDomainScore(score.LearnDomainScore.Value);
                }
                    
                else
                    return 0;
                    
                
        }

        public string Name
        {
            get { return GetSemesterString(Semester); }
        }

        public static string GetSemesterString(SemesterData semester)
        {
            return string.Format("學習領域({0}:{1})", semester.GradeYear, semester.Semester);
        }

        public int Grade { get { return Semester.GradeYear; } }
        public int sm { get { return Semester.Semester; } }

        #endregion
    }
}
