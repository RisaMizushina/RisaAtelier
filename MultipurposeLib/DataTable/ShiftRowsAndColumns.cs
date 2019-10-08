using System;
using System.Data;
using System.Activities;
using System.Activities.Presentation.Metadata;
using System.ComponentModel;

namespace RisaAtelier.MultipurposeLib.DataTable
{
    /// <summary>
    /// DataRowの列と行を入れ替えるアクティビティ
    /// </summary>
    public class ShiftRowsAndColumns : CodeActivity
    {

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(ShiftRowsAndColumns);

            DesignerMetadata.RegisterHelper(ref builder, t
                                                , Properties.Resources.ActivityName_DataTable_SwapRowsAndColumns
                                                , Properties.Resources.ActivityTree_DataTable_SwapRowsAndColumns
                                                , Properties.Resources.ActivityDesc_DataTable_SwapRowsAndColumns);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(ShiftRowsAndColumns.Source)
                                                , Properties.Resources.DataTable_SRAC_InArgName1
                                                , Properties.Resources.DataTable_SRAC_InArgCategory1
                                                , Properties.Resources.DataTable_SRAC_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(ShiftRowsAndColumns.Result)
                                                , Properties.Resources.DataTable_SRAC_OutArgName1
                                                , Properties.Resources.DataTable_SRAC_OutArgCategory1
                                                , Properties.Resources.DataTable_SRAC_OutArgDesc1);
        }

        /// <summary>
        /// 入力値
        /// </summary>
        [RequiredArgument]
        public InArgument<System.Data.DataTable> Source { get; set; }

        /// <summary>
        /// 出力値
        /// </summary>
        public OutArgument<System.Data.DataTable> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            // objectの、配列で、再構築します
            var xSize = Source.Get(context).Columns.Count;
            var ySize = Source.Get(context).Rows.Count;
            object[,] map = new object[xSize ,ySize + 1];

            // 0行目として、列名をセット
            for(var i = 0; i < xSize; i++)
            {
                map[i, 0] = Source.Get(context).Columns[i].ColumnName;
            }

            // 1行目以降
            for(var x = 0; x < xSize; x++)
            {
                for(var y = 0; y < ySize; y++)
                {
                    map[x, y + 1] = Source.Get(context).Rows[y][x];
                }
            }

            var ret = new System.Data.DataTable();

            // 縦方向に、読んで、列を作ります
            for(var y = 0; y <= ySize; y++)
            {
                ret.Columns.Add(new DataColumn(map[0, y].ToString()));
            }

            // 値を、セットします
            for (var x = 1; x < xSize; x++)
            {
                DataRow row = ret.NewRow();
                for(var y = 0; y <= ySize; y++)
                {
                    row[y] = map[x, y];
                }
                ret.Rows.Add(row);
            }

            Result.Set(context, ret);

        }
    }
}
