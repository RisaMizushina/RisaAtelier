using System;
using System.Activities;
using System.IO;
using System.Activities.Presentation.Metadata;

namespace RisaAtelier.MultipurposeLib.Convert.SecureString
{
    /// <summary>
    /// Stringから、SecureStringに、変換を行います
    /// </summary>
    public class FromString : CodeActivity
    {

        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            var t = typeof(FromString);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Convert_SecureString_FromString
                                                           , Properties.Resources.ActivityTree_Convert_SecureString_FromString
                                                           , Properties.Resources.ActivityTree_Convert_SecureString_FromString);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(InputValue)
                                                            , Properties.Resources.Convert_SecureString_FS_InArgName1
                                                            , Properties.Resources.Convert_SecureString_FS_InArgCategory1
                                                            , Properties.Resources.Convert_SecureString_FS_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(Result)
                                                , Properties.Resources.Convert_SecureString_FS_OutArgName1
                                                , Properties.Resources.Convert_SecureString_FS_OutArgCategory1
                                                , Properties.Resources.Convert_SecureString_FS_OutArgDesc1);

        }

        [RequiredArgument()]
        public InArgument<String> InputValue { get; set; }


        public OutArgument<System.Security.SecureString> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var ret = new System.Security.SecureString();
            foreach(var c in InputValue.Get(context).ToCharArray())
            {
                ret.AppendChar(c);
            }

            Result.Set(context, ret.Copy());

            ret.Dispose();
        }
    }
}
