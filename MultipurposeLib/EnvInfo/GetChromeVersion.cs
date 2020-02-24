using System;
using System.Activities;
using System.Activities.Presentation.Metadata;

namespace RisaAtelier.MultipurposeLib.EnvInfo
{
    public class GetChromeVersion : CodeActivity
    {
        public OutArgument<string> Result { get; set; }

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(GetChromeVersion);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Env_ChromeVer
                                                            , Properties.Resources.ActivityTree_Env_ChromeVer
                                                            , Properties.Resources.ActivityDesc_Env_ChromeVer);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(GetChromeVersion.Result)
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgName1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgCategory1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgDesc1);
        }


        protected override void Execute(CodeActivityContext context)
        {
            // Chromeの、インストール先を、探します。
            var targetDir = System.IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\Application");

            if (!System.IO.Directory.Exists(targetDir))
            {
                targetDir = @"C:\Program Files\Google\Chrome\Application";

                if (!System.IO.Directory.Exists(targetDir))
                {
                    targetDir = @"C:\Program Files (x86)\Google\Chrome\Application";

                    if (!System.IO.Directory.Exists(targetDir))
                    {
                        throw new SystemException(Properties.Resources.Exception_Chrome_Not_Found);
                    }
                }
            }

            // 仮値として、0.0を、バージョンとします
            var ret = "0.0";

            // バージョンの、大小比較
            bool IsNewerThan(string s1, string s2)
            {
                try
                {
                    return Version.Parse(s1) < Version.Parse(s2);
                }
                catch { }
                return false;
            }

            foreach (var d in System.IO.Directory.GetDirectories(targetDir))
            {
                if (IsNewerThan(ret, System.IO.Path.GetFileName(d)))
                {
                    ret = System.IO.Path.GetFileName(d);
                }
            }

            Result.Set(context, ret);
        }
    }
}
