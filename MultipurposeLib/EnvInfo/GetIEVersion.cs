using System;
using System.Activities;
using System.Activities.Presentation.Metadata;
using Microsoft.Win32;

namespace RisaAtelier.MultipurposeLib.EnvInfo
{
    public class GetIEVersion : CodeActivity
    {

        public OutArgument<string> Result { get; set; }


        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(GetIEVersion);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Env_IEVer
                                                            , Properties.Resources.ActivityTree_Env_IEVer
                                                            , Properties.Resources.ActivityDesc_Env_IEVer);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(GetIEVersion.Result)
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgName1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgCategory1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgDesc1);
        }

        protected override void Execute(CodeActivityContext context)
        {
            var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer", false);
            var ret = string.Empty;

            try
            {
                ret = reg.GetValue("svcVersion").ToString();

                if (ret.Equals(string.Empty)) throw new NullReferenceException();

            }
            catch(NullReferenceException)
            {
                ret = reg.GetValue("Version").ToString();
            }

            Result.Set(context, ret);
        }
    }
}
