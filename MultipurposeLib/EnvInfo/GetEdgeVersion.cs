using System;
using System.Activities;
using System.Activities.Presentation.Metadata;

namespace RisaAtelier.MultipurposeLib.EnvInfo
{
    public class GetEdgeVersion : CodeActivity
    {
        public OutArgument<string> Result { get; set; }

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(GetEdgeVersion);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Env_EdgeVer
                                                            , Properties.Resources.ActivityTree_Env_EdgeVer
                                                            , Properties.Resources.ActivityDesc_Env_EdgeVer);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(GetEdgeVersion.Result)
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgName1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgCategory1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgDesc1);
        }

        protected override void Execute(CodeActivityContext context)
        {
            var ret = string.Empty;
            try
            {
                foreach (var k in Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(@"\ActivatableClasses\Package", false).GetSubKeyNames())
                {
                    if (k.Contains("Microsoft.MicrosoftEdge_"))
                    {
                        ret = k.Replace("Microsoft.MicrosoftEdge_", string.Empty).Split("_".ToCharArray())[0];
                    }
                }
            }
            catch
            {
            }

            Result.Set(context, ret);
        }
    }
}
