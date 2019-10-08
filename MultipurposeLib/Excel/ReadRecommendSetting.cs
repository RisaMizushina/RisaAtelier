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
    public class ReadRecommendSetting : CodeActivity
    {

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(ReadRecommendSetting);

            DesignerMetadata.RegisterHelper(ref builder, t
                                                , Properties.Resources.ActivityName_Excel_ReadReccomendSetting
                                                , Properties.Resources.ActivityTree_Excel_ReadRecommendSetting
                                                , Properties.Resources.ActivityDesc_Excel_ReadReccomendSetting);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(ReadRecommendSetting.SetReadOnlyRecommended)
                                                , Properties.Resources.Excel_RRS_InArgName1
                                                , Properties.Resources.Excel_RRS_InArgCategory1
                                                , Properties.Resources.Excel_RRS_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(ReadRecommendSetting.TargetExcelFile)
                                                , Properties.Resources.Excel_RRS_InArgName2
                                                , Properties.Resources.Excel_RRS_InArgCategory2
                                                , Properties.Resources.Excel_RRS_InArgDesc2);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(ReadRecommendSetting.ExcelBookPassword)
                                                , Properties.Resources.Excel_RRS_InArgName3
                                                , Properties.Resources.Excel_RRS_InArgCategory3
                                                , Properties.Resources.Excel_RRS_InArgDesc3);
        }

        [RequiredArgument()]
        public InArgument<bool> SetReadOnlyRecommended { get; set; } = false;

        [RequiredArgument()]
        public InArgument<string> TargetExcelFile { get; set; }

        public InArgument<string> ExcelBookPassword { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            object xlApp = null;
            string targetFile = TargetExcelFile.Get(context);
            if(!System.IO.Path.IsPathRooted(targetFile))
            {
                targetFile = System.IO.Path.Combine(Environment.CurrentDirectory, targetFile);
            }

            try
            {
                xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                COMUtil.InvokeSetProperty(ref xlApp, "DisplayAlerts", new object[] { false });

                object xlWorkBooks = null;
                try
                {
                    xlWorkBooks = COMUtil.InvokeGetProperty(ref xlApp, "Workbooks");

                    object xlWorkBookObject = null;

                    try
                    {
                        xlWorkBookObject = COMUtil.InvokeByMethod(ref xlWorkBooks, "Open", new object[] {
                                targetFile                      // FileName
                                , Type.Missing                  // UpdateLinks
                                , false                         // ReadOnly
                                , Type.Missing                  // Format
                                , !String.IsNullOrEmpty(ExcelBookPassword.Get(context)) ? ExcelBookPassword.Get(context) : Type.Missing // Password
                                , !String.IsNullOrEmpty(ExcelBookPassword.Get(context)) ? ExcelBookPassword.Get(context) : Type.Missing // WriteResPassword
                                , true                          // Ignorereadonlyrecommended
                        
                        });

                        COMUtil.InvokeSetProperty(ref xlWorkBookObject, "ReadOnlyRecommended", new object[] { SetReadOnlyRecommended.Get(context) });

                        COMUtil.InvokeByMethod(ref xlWorkBookObject, "Save", null);
                    }
                    finally
                    {
                        COMUtil.COMRelease(ref xlWorkBookObject, false);
                    }
                }
                finally
                {
                    COMUtil.COMRelease(ref xlWorkBooks, false);
                }

                COMUtil.InvokeSetProperty(ref xlApp, "DisplayAlerts", new object[] { true });

                COMUtil.InvokeByMethod(ref xlApp, "Quit", null) ;
            }
            finally
            {
                COMUtil.COMRelease(ref xlApp, true);
            }
        }



    }
}
