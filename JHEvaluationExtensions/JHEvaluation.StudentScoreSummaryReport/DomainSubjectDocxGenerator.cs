using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

namespace JHEvaluation.StudentScoreSummaryReport
{
    public static class DomainSubjectDocxGenerator
    {
        private static readonly List<string> FixedDomains = new List<string>
{
                "語文",
                "數學",
                "社會",
                "生活課程",
                "自然科學",
                "藝術",
                "綜合活動",
                "科技",
                "健康與體育"
};

        private static readonly string[] SemesterTitles = new string[]
        {
        "一上", "一下",
        "二上", "二下",
        "三上", "三下",
        "四上", "四下",
        "五上", "五下",
        "六上", "六下"
        };

        private const int SubjectCountPerDomain = 10;
        private const int SemesterCount = 12;

        public static void GenerateDomainSubjectMergeFieldDocx(string outputDocxPath)
        {
            if (string.IsNullOrWhiteSpace(outputDocxPath))
                throw new ArgumentException("outputDocxPath 不可空白。", nameof(outputDocxPath));

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "標楷體";
            builder.Font.Size = 10;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            builder.Writeln("領域科目成績合併欄位");
            builder.Writeln();

            WriteDomainThreeTables(builder, FixedDomains, false);

            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("領域科目原始成績合併欄位");
            builder.Writeln();

            WriteDomainThreeTables(builder, FixedDomains, true);

            doc.Save(outputDocxPath, SaveFormat.Docx);
        }

        private static void WriteDomainThreeTables(DocumentBuilder builder, IEnumerable<string> domains, bool isOriginalScore)
        {
            foreach (string domain in domains)
            {
                string blockTitle = domain + (isOriginalScore ? "科目原始成績" : "科目成績");

                builder.Writeln();
                builder.Writeln(blockTitle);

                // 1. 成績表
                builder.Writeln(isOriginalScore ? "原始成績表" : "成績表");
                WriteSingleTypeTable(builder, domain, isOriginalScore, FieldGroupType.Score);

                builder.Writeln();

                // 2. 等第表
                builder.Writeln(isOriginalScore ? "原始等第表" : "等第表");
                WriteSingleTypeTable(builder, domain, isOriginalScore, FieldGroupType.Level);

                builder.Writeln();

                // 3. 權數表
                builder.Writeln("權數表");
                WriteSingleTypeTable(builder, domain, isOriginalScore, FieldGroupType.Weight);

                builder.Writeln();
                builder.Writeln();
            }
        }

        private static void WriteSingleTypeTable(DocumentBuilder builder, string domain, bool isOriginalScore, FieldGroupType fieldGroupType)
        {
            builder.StartTable();

            WriteSingleTypeHeader(builder, isOriginalScore, fieldGroupType);
            WriteSingleTypeRows(builder, domain, isOriginalScore, fieldGroupType);

            Table table = builder.EndTable();
            table.AllowAutoFit = true;
        }

        private static void WriteSingleTypeHeader(DocumentBuilder builder, bool isOriginalScore, FieldGroupType fieldGroupType)
        {
            InsertHeaderCell(builder, "科目");

            foreach (string semesterTitle in SemesterTitles)
            {
                string headerText = semesterTitle;

                switch (fieldGroupType)
                {
                    case FieldGroupType.Score:
                        headerText += isOriginalScore ? "原始成績" : "成績";
                        break;

                    case FieldGroupType.Level:
                        headerText += isOriginalScore ? "原始等第" : "等第";
                        break;

                    case FieldGroupType.Weight:
                        headerText += "權數";
                        break;
                }

                InsertHeaderCell(builder, headerText);
            }

            builder.EndRow();
        }

        private static void WriteSingleTypeRows(DocumentBuilder builder, string domain, bool isOriginalScore, FieldGroupType fieldGroupType)
        {
            for (int i = 1; i <= SubjectCountPerDomain; i++)
            {
                string subjectNameField = domain + "_科目名稱" + i;

                // 科目名稱
                builder.InsertCell();
                ResetCellFormat(builder);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(string.Empty);
                builder.InsertField(
                    "MERGEFIELD " + subjectNameField + " \\* MERGEFORMAT",
                    "S" + i
                );

                for (int semesterIndex = 1; semesterIndex <= SemesterCount; semesterIndex++)
                {
                    string mergeFieldName = GetMergeFieldName(domain, i, semesterIndex, isOriginalScore, fieldGroupType);
                    string displayText = GetDisplayText(i, semesterIndex, fieldGroupType);

                    builder.InsertCell();
                    ResetCellFormat(builder);
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    builder.Write(string.Empty);
                    builder.InsertField(
                        "MERGEFIELD " + mergeFieldName + " \\* MERGEFORMAT",
                        displayText
                    );
                }

                builder.EndRow();
            }
        }

        private static string GetMergeFieldName(string domain, int subjectIndex, int semesterIndex, bool isOriginalScore, FieldGroupType fieldGroupType)
        {
            switch (fieldGroupType)
            {
                case FieldGroupType.Score:
                    return domain + "_科目" + subjectIndex + "_" + (isOriginalScore ? "原始成績" : "成績") + semesterIndex;

                case FieldGroupType.Level:
                    return domain + "_科目" + subjectIndex + "_" + (isOriginalScore ? "原始等第" : "等第") + semesterIndex;

                case FieldGroupType.Weight:
                    return domain + "_科目" + subjectIndex + "_權數" + semesterIndex;

                default:
                    throw new ArgumentOutOfRangeException(nameof(fieldGroupType), fieldGroupType, null);
            }
        }

        private static string GetDisplayText(int subjectIndex, int semesterIndex, FieldGroupType fieldGroupType)
        {
            switch (fieldGroupType)
            {
                case FieldGroupType.Score:
                    return "SC" + subjectIndex + "_" + semesterIndex;

                case FieldGroupType.Level:
                    return "LV" + subjectIndex + "_" + semesterIndex;

                case FieldGroupType.Weight:
                    return "WT" + subjectIndex + "_" + semesterIndex;

                default:
                    return string.Empty;
            }
        }

        private static void InsertHeaderCell(DocumentBuilder builder, string text)
        {
            builder.InsertCell();
            ResetCellFormat(builder);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Bold = true;
            builder.Write(text);
            builder.Bold = false;
        }

        private static void ResetCellFormat(DocumentBuilder builder)
        {
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        }

        private enum FieldGroupType
        {
            Score,
            Level,
            Weight
        }
    }
}
