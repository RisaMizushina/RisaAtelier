using System;
using System.Data;
using System.Activities;
using System.Activities.Presentation.Metadata;
using System.Collections.Generic;
using System.ComponentModel;

namespace RisaAtelier.MultipurposeLib.DataTable
{
    public class JointTable : CodeActivity
    {

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(JointTable);

            DesignerMetadata.RegisterHelper(ref builder, t
                                                , Properties.Resources.ActivityName_DataTable_JointTable
                                                , Properties.Resources.ActivityTree_DataTable_JointTable
                                                , Properties.Resources.ActivityDesc_DataTable_JointTable);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(JointTable.LeftTable)
                                                , Properties.Resources.DataTable_JT_InArgName1
                                                , Properties.Resources.DataTable_JT_InArgCategory1
                                                , Properties.Resources.DataTable_JT_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(JointTable.RightTable)
                                                , Properties.Resources.DataTable_JT_InArgName2
                                                , Properties.Resources.DataTable_JT_InArgCategory2
                                                , Properties.Resources.DataTable_JT_InArgDesc2);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(JointTable.LeftSort)
                                                , Properties.Resources.DataTable_JT_InArgName3
                                                , Properties.Resources.DataTable_JT_InArgCategory3
                                                , Properties.Resources.DataTable_JT_InArgDesc3);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(JointTable.RightSort)
                                                , Properties.Resources.DataTable_JT_InArgName4
                                                , Properties.Resources.DataTable_JT_InArgCategory4
                                                , Properties.Resources.DataTable_JT_InArgDesc4);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(JointTable.CheckRecordsCount)
                                                , Properties.Resources.DataTable_JT_InArgName5
                                                , Properties.Resources.DataTable_JT_InArgCategory5
                                                , Properties.Resources.DataTable_JT_InArgDesc5);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(JointTable.Result)
                                                , Properties.Resources.DataTable_JT_OutArgName1
                                                , Properties.Resources.DataTable_JT_OutArgCategory1
                                                , Properties.Resources.DataTable_JT_OutArgDesc1);

        }

        [RequiredArgument]
        public InArgument<System.Data.DataTable> LeftTable { get; set; }

        [RequiredArgument]
        public InArgument<System.Data.DataTable> RightTable { get; set; }

        public InArgument<String> LeftSort { get; set; }

        public InArgument<String> RightSort { get; set; }

        public InArgument<Boolean> CheckRecordsCount { get; set; } = true;

        public OutArgument<System.Data.DataTable> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            System.Data.DataTable LTable = LeftTable.Get(context).Select(string.Empty, LeftSort.Get(context)).CopyToDataTable();
            System.Data.DataTable RTable = RightTable.Get(context).Select(string.Empty, RightSort.Get(context)).CopyToDataTable();


            if(CheckRecordsCount.Get(context) && LTable.Rows.Count != RTable.Rows.Count)
            {
                throw new ArgumentException(Properties.Resources.Exception_Records_Count_Invalid);
            }

            System.Data.DataTable ret = new System.Data.DataTable();

            var convDic = new Dictionary<DataColumn, String>();
            
            // 左側テーブルは、列名を、そのまま使います
            foreach(DataColumn c in LTable.Columns)
            {
                convDic.Add(c, c.ColumnName);
                ret.Columns.Add(new DataColumn(c.ColumnName, c.DataType));
            }

            // 右側テーブルは、列名が、重複したら、 _0 形式で、数値を足します
            foreach (DataColumn c in RTable.Columns)
            {
                var newColName = string.Empty;
                for (int i = -1; convDic.ContainsValue(newColName = (i < 0 ? c.ColumnName : string.Format("{0}_{1}", newColName, i.ToString().Trim()))) ; i++) ;

                convDic.Add(c, newColName);
                ret.Columns.Add(new DataColumn(newColName, c.DataType));
            }

            int maxRecord = LTable.Rows.Count > RTable.Rows.Count ? LTable.Rows.Count : RTable.Rows.Count;

            for(int i = 0; i < maxRecord; i++)
            {
                DataRow row = ret.NewRow();

                if(LTable.Rows.Count > i)
                {
                    foreach (DataColumn c in LTable.Columns)
                    {
                        row[convDic[c]] = LTable.Rows[i][c];
                    }
                }

                if(RTable.Rows.Count > i)
                {
                    foreach (DataColumn c in RTable.Columns)
                    {
                        row[convDic[c]] = RTable.Rows[i][c];
                    }
                }

                ret.Rows.Add(row);
            }

            Result.Set(context, ret);

        }
    }
}
