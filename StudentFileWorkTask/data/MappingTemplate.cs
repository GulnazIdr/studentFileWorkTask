using System.Collections.Generic;

namespace StudentFileWorkTask.data
{
    public class MappingTemplate
    {
        public string StudentColumn { get; set; } = "";
        public string GroupColumn { get; set; } = "";
        public string DateColumn { get; set; } = "";
        public List<string> QuestionColumns { get; set; } = new();
        public List<string> ScoreColumns { get; set; } = new();
    }
}