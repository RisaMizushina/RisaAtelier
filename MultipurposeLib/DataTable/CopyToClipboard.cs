using System;
using System.Data;
using System.Text;
using System.Activities;
using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.Collections.Generic;

namespace RisaAtelier.MultipurposeLib.DataTable
{
    public class CopyToClipboard : CodeActivity
    {
        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(CopyToClipboard);

            DesignerMetadata.RegisterHelper(ref builder, t
                                                , Properties.Resources.ActivityName_DataTable_CopyToClipboard
                                                , Properties.Resources.ActivityTree_DataTable_CopyToClipboard
                                                , Properties.Resources.ActivityDesc_DataTable_CopyToClipboard);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(CopyToClipboard.Source)
                                                , Properties.Resources.DataTable_CTC_InArgName1
                                                , Properties.Resources.DataTable_CTC_InArgCategory1
                                                , Properties.Resources.DataTable_CTC_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(CopyToClipboard.DateTimeFormat)
                                                , Properties.Resources.DataTable_CTC_InArgName2
                                                , Properties.Resources.DataTable_CTC_InArgCategory2
                                                , Properties.Resources.DataTable_CTC_InArgDesc2);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(CopyToClipboard.Escape)
                                                , Properties.Resources.DataTable_CTC_InArgName3
                                                , Properties.Resources.DataTable_CTC_InArgCategory3
                                                , Properties.Resources.DataTable_CTC_InArgDesc3);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(CopyToClipboard.SeparateType)
                                                , Properties.Resources.DataTable_CTC_InArgName4
                                                , Properties.Resources.DataTable_CTC_InArgCategory4
                                                , Properties.Resources.DataTable_CTC_InArgDesc4);
        }

        /// <summary>
        /// 区切り文字の設定
        /// </summary>
        public enum SeparateType
        {
            COMMA, TAB
        }

        /// <summary>
        /// 処理対象のDataTable
        /// </summary>
        [RequiredArgument]
        public InArgument<System.Data.DataTable> Source { get; set; }

        /// <summary>
        /// 文字列型の置換フォーマット
        /// </summary>
        public InArgument<String> DateTimeFormat { get; set; } = "yyyy/MM/dd HH:mm:ss";

        /// <summary>
        /// エスケープ処理の有無
        /// </summary>
        public InArgument<Boolean> Escape { get; set; } = true;

        /// <summary>
        /// 区切り文字
        /// </summary>
        public SeparateType Separator { get; set; } = SeparateType.COMMA;

        protected override void Execute(CodeActivityContext context)
        {
            var sb = new StringBuilder();

            var hd = new List<String>();
            foreach(DataColumn c in Source.Get(context).Columns)
            {
                hd.Add(c.ColumnName);
            }
            
            sb.AppendLine(CreateLine(hd.ToArray(), Escape.Get(context)));

            foreach(DataRow row in Source.Get(context).Rows)
            {
                var body = new List<String>();
                for(int i = 0; i < Source.Get(context).Columns.Count; i++)
                {
                    if(Source.Get(context).Columns[i].DataType == typeof(DateTime) || row[i].GetType() == typeof(DateTime))
                    {
                        body.Add(row[i] != null ? ((DateTime)row[i]).ToString(DateTimeFormat.Get(context)) : null);
                    }
                    else
                    {
                        body.Add(row[i].ToString());
                    }
                }
                sb.AppendLine(CreateLine(body.ToArray(), Escape.Get(context)));
            }

            System.Windows.Clipboard.SetText(sb.ToString());

        }

        private string CreateLine(string[] array, Boolean doEscape)
        {
            string ret = string.Empty;

            for(int i=0; i < array.Length; i++)
            {
                ret += (i == 0 ? string.Empty : (Separator == SeparateType.COMMA ? "," : "\t"));
                ret += doEscape ? ExecEscape(array[i]) : array[i];
            }
            return ret;
        }

        private string ExecEscape(string input)
        {
            // " は "" にエスケープ
            if(input.IndexOf("\"") >= 0)
            {
                input = input.Replace("\"", "\"\"");
            }

            // 区切り文字と、改行があったら、エスケープします
            if (input.IndexOf(Separator == SeparateType.COMMA ? "," : "\t") >= 0 | input.IndexOf("\r") >= 0 | input.IndexOf("\n") >= 0)
            {
                input = "\"" + input + "\"";
            }

            return input;
        }

    }
}
