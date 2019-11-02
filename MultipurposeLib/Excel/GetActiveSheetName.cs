using System;
using System.Activities;
using System.Activities.Presentation.Metadata;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace RisaAtelier.MultipurposeLib.Excel
{
    public class GetActiveSheetName : CodeActivity
    {

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(GetActiveSheetName);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Excel_GetActiveSheetName
                                                            , Properties.Resources.ActivityTree_Excel_GetActiveSheetName
                                                            , Properties.Resources.ActivityDesc_Excel_GetActiveSheetName);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(GetActiveSheetName.ActiveSheetName)
                                                            , Properties.Resources.Excel_GASN_OutArgName1
                                                            , Properties.Resources.Excel_GASN_OutArgCategory1
                                                            , Properties.Resources.Excel_GASN_OutArgDesc1);

        }

        public OutArgument<String> ActiveSheetName { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            // Excelの、プロセスがあるか、確認します
            try
            {
                if (System.Diagnostics.Process.GetProcessesByName("excel").Length == 0) throw new Exception();
            }
            catch
            {
                throw new ApplicationException(Properties.Resources.Exception_Excel_Not_Started);
            }

            object xlApp = null;

            try
            {
                xlApp = Marshal.GetActiveObject("Excel.Application");


                object xlActiveSheet = null;

                try
                {
                    xlActiveSheet = COMUtil.InvokeGetProperty(ref xlApp, "ActiveSheet");

                    ActiveSheetName.Set(context, COMUtil.InvokeGetProperty(ref xlActiveSheet, "Name").ToString());
                }
                finally
                {
                    COMUtil.COMRelease(ref xlActiveSheet, false);
                }

            }
            finally
            {
                COMUtil.COMRelease(ref xlApp, true);
            }
        }

    }
}
