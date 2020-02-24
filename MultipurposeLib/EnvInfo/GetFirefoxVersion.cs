using System;
using System.Activities;
using System.Activities.Presentation.Metadata;

namespace RisaAtelier.MultipurposeLib.EnvInfo
{


    public class GetFirefoxVersion : CodeActivity
    {
        public OutArgument<string> Result { get; set; }

        public InArgument<string> FireFoxPath { get; set; } = @"C:\Program Files\Mozilla Firefox";

        /// <summary>
        /// プロパティの設定
        /// </summary>
        /// <param name="builder"></param>
        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            Type t = typeof(GetFirefoxVersion);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Env_FirefoxVer
                                                            , Properties.Resources.ActivityTree_Env_FirefoxVer
                                                            , Properties.Resources.ActivityDesc_Env_FirefoxVer);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(GetFirefoxVersion.Result)
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgName1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgCategory1
                                                            , Properties.Resources.Environment_Env_GetBrowserVersion_OutArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(GetFirefoxVersion.FireFoxPath)
                                                            , Properties.Resources.Environment_Env_GetFirefoxVersion_InArgName1
                                                            , Properties.Resources.Environment_Env_GetFirefoxVersion_InArgCategory1
                                                            , Properties.Resources.Environment_Env_GetFirefoxVersion_InArgDesc1);
        }

        protected override void Execute(CodeActivityContext context)
        {
            var fxDir = FireFoxPath.Get(context);

            if (!System.IO.Directory.Exists(fxDir))
            {
                throw new SystemException(Properties.Resources.Exception_Firefox_Not_Found);
            }

            var arg = System.IO.Path.Combine(fxDir, "firefox.exe");
            arg = arg.Contains(" ") ? "\"" + arg + "\"" : arg;

            var pi = new System.Diagnostics.ProcessStartInfo("cmd.exe", "/c " + arg + " -v | more");
            pi.RedirectStandardOutput = true;
            pi.CreateNoWindow = true;
            pi.UseShellExecute = false;

            var p = System.Diagnostics.Process.Start(pi);

            var ret = p.StandardOutput.ReadToEnd().Replace("Mozilla", string.Empty).Replace("Firefox", string.Empty).Trim();
            p.WaitForExit();

            Result.Set(context,ret);

            
        }
    }
}
