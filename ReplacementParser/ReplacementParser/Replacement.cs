using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace ReplacementParser
{
    class Lesson
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Teacher { get; set; }
        public string Audience { get; set; }
    }
    class GroupReplacement
    {
        public string Group { get; set; }
        public List<Lesson> Lessons { get; set; }
        public GroupReplacement() => Lessons = new List<Lesson>();
    }
    class Replacement
    {
        public string PathToFile { get; set; }
        public Replacement(string pathToFile) => this.PathToFile = pathToFile;
        public List<GroupReplacement> ResultReplacements { get; set; }
        public string AllText { get; set; }
        public void GetInfo()
        {
            List<GroupReplacement> resultReplacements = new List<GroupReplacement>();
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Open(PathToFile);
            AllText = wordDoc.Range().Text;
            int groupId = 0, lessonId = 0, lessonNameId = 0, teacherId = 0, audienceId = 0;
            foreach (Word.Table table in wordDoc.Tables)
            {
                for (int rowId = 1; rowId < table.Rows.Count; rowId++)
                {
                    GroupReplacement groupReplacement = new GroupReplacement();
                    Word.Row row = table.Rows[rowId];
                    if (rowId == 1)
                        for (int cellId = 1; cellId <= row.Cells.Count; cellId++)
                        {
                            string content = row.Cells[cellId].Range.Text.Replace("\r\a", "").Trim(' ');
                            if (Regex.IsMatch(content, "Группа", RegexOptions.IgnoreCase)) groupId = cellId;
                            else if (Regex.IsMatch(content, "Пара", RegexOptions.IgnoreCase)) lessonId = cellId;
                            else if (Regex.IsMatch(content, "Предмет", RegexOptions.IgnoreCase)) lessonNameId = cellId;
                            else if (Regex.IsMatch(content, "Ауд.", RegexOptions.IgnoreCase)) audienceId = cellId;
                            else if (Regex.IsMatch(content, "Преподаватель", RegexOptions.IgnoreCase)) teacherId = cellId;
                        }
                    else
                    {
                        groupReplacement.Group = row.Cells[groupId].Range.Text.Replace("\r\a", "").Trim(' ');

                        List<string> listLessonsId = new List<string>();
                        List<string> listLessonsName = new List<string>();
                        List<string> listAudience = new List<string>();
                        List<string> listTeacher = new List<string>();

                        foreach (var _lessonId in row.Cells[lessonId].Range.Text.Split(new[] { "\r" }, StringSplitOptions.None))
                            if (_lessonId != "\a") listLessonsId.Add(_lessonId);

                        foreach (var _lessonName in row.Cells[lessonNameId].Range.Text.Split(new[] { "\r" }, StringSplitOptions.None))
                            if (_lessonName != "\a") listLessonsName.Add(_lessonName);
                            else for (int i = 0; i < listLessonsId.Count - listLessonsName.Count; i++)
                                    listLessonsName.Add("");

                        foreach (var _audienceId in row.Cells[audienceId].Range.Text.Split(new[] { "\r" }, StringSplitOptions.None))
                            if (_audienceId != "\a") listAudience.Add(_audienceId);
                            else for (int i = 0; i < listLessonsId.Count - listAudience.Count; i++)
                                    listAudience.Add("");

                        foreach (var _teacher in row.Cells[teacherId].Range.Text.Split(new[] { "\r" }, StringSplitOptions.None))
                            if (_teacher != "\a") listTeacher.Add(_teacher);
                            else for (int i = 0; i < listLessonsId.Count - listTeacher.Count; i++)
                                    listTeacher.Add("");

                        for (int i = 0; i < listLessonsId.Count; i++)
                        {
                            groupReplacement.Lessons.Add(new Lesson()
                            {
                                Id = listLessonsId[i],
                                Name = listLessonsName[i],
                                Teacher = listTeacher[i],
                                Audience = listAudience[i]
                            });
                        }
                        resultReplacements.Add(groupReplacement);
                    }
                }
            }
            wordDoc.Close();
            ResultReplacements = resultReplacements;
        }
        
        public int SendReplacement(string pathToSite, List<GroupReplacement> replacements)
        {


            return 0;
        }

    }
}
