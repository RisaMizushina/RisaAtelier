using System;
using System.Activities;
using System.IO;
using System.Activities.Presentation.Metadata;

namespace RisaAtelier.MultipurposeLib.File
{
    public class WaitForFileGrowthCompleted : CodeActivity
    {

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(WaitForFileGrowthCompleted);

            DesignerMetadata.RegisterHelper(ref builder, t
                                                , Properties.Resources.ActivityName_File_WaitForFileGrowthCompleted
                                                , Properties.Resources.ActivityTree_File_WaitForFileGrowthCompleted
                                                , Properties.Resources.ActivityDesc_File_WaitForFileGrowthCompleted);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(WaitForFileGrowthCompleted.FileName)
                                                , Properties.Resources.File_WFFGC_InArgName1
                                                , Properties.Resources.File_WFFGC_InArgCategory1
                                                , Properties.Resources.File_WFFGC_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(WaitForFileGrowthCompleted.IntervalMS)
                                                , Properties.Resources.File_WFFGC_InArgName2
                                                , Properties.Resources.File_WFFGC_InArgCategory2
                                                , Properties.Resources.File_WFFGC_InArgDesc2);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(WaitForFileGrowthCompleted.TimeOutSeconds)
                                                , Properties.Resources.File_WFFGC_InArgName3
                                                , Properties.Resources.File_WFFGC_InArgCategory3
                                                , Properties.Resources.File_WFFGC_InArgDesc3);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(WaitForFileGrowthCompleted.FinalFileSize)
                                                , Properties.Resources.File_WFFGC_OutArgName1
                                                , Properties.Resources.File_WFFGC_OutArgCategory1
                                                , Properties.Resources.File_WFFGC_OutArgDesc1);
        }

        public InArgument<String> FileName { get; set; }

        public InArgument<Int32> IntervalMS { get; set; } = 2000;

        public InArgument<Int32> TimeOutSeconds { get; set; }

        public OutArgument<long> FinalFileSize { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var fName = FileName.Get(context);
            if (!System.IO.File.Exists(fName)) throw new FileNotFoundException();

            if (IntervalMS.Get(context) <= 0) throw new ArgumentOutOfRangeException();

            long lastFileSize = 0;
            DateTime timeLimit = TimeOutSeconds.Get(context) > 0 ? DateTime.Now.AddSeconds(TimeOutSeconds.Get(context)) : DateTime.MaxValue;

            do
            {
                lastFileSize = new FileInfo(fName).Length;
                System.Threading.Thread.Sleep(IntervalMS.Get(context));
                if (timeLimit.Ticks > DateTime.Now.Ticks) throw new TimeoutException();
                
            } while (lastFileSize != new FileInfo(fName).Length);

            FinalFileSize.Set(context, new FileInfo(fName).Length);
        }
    }
}
