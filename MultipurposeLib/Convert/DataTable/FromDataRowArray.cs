using System;
using System.Activities;
using System.Activities.Presentation.Metadata;
using System.Data;

namespace RisaAtelier.MultipurposeLib.Convert.DataTable
{
    public class FromDataRowArray : CodeActivity
    {

        [RequiredArgument()]
        public InArgument<DataRow[]> DataRowArray { get; set; }

        public OutArgument<System.Data.DataTable> Result { get; set; }

        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            var t = typeof(FromDataRowArray);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Convert_DataTable_FromDataRowArray
                                                           , Properties.Resources.ActivityTree_Convert_DataTable_FromDataRowArray
                                                           , Properties.Resources.ActivityTree_Convert_DataTable_FromDataRowArray);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(DataRowArray)
                                                            , Properties.Resources.Convert_DataTable_FDArr_InArgName1
                                                            , Properties.Resources.Convert_DataTable_FDArr_InArgCategory1
                                                            , Properties.Resources.Convert_DataTable_FDArr_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(Result)
                                                , Properties.Resources.Convert_DataTable_FDArr_OutArgName1
                                                , Properties.Resources.Convert_DataTable_FDArr_OutArgCategory1
                                                , Properties.Resources.Convert_DataTable_FDArr_OutArgDesc1);

        }
        protected override void Execute(CodeActivityContext context)
        {
            Result.Set(context, DataRowArray.Get(context).CopyToDataTable());
        }
    }
}
