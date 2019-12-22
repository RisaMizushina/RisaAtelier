using System;
using System.Activities;
using System.IO;
using System.Activities.Presentation.Metadata;
using System.Runtime.InteropServices;

namespace RisaAtelier.MultipurposeLib.Convert.SecureString
{
    public class ToString : CodeActivity
    {

        [RequiredArgument()]
        public InArgument<System.Security.SecureString> InputValue { get; set; }

        public OutArgument<string> Result { get; set; }


        public static void SetMetaData(ref AttributeTableBuilder builder)
        {
            var t = typeof(ToString);

            DesignerMetadata.RegisterHelper(ref builder, t, Properties.Resources.ActivityName_Convert_SecureString_ToString
                                                           , Properties.Resources.ActivityTree_Convert_SecureString_ToString
                                                           , Properties.Resources.ActivityTree_Convert_SecureString_ToString);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(InputValue)
                                                            , Properties.Resources.Convert_SecureString_TS_InArgName1
                                                            , Properties.Resources.Convert_SecureString_TS_InArgCategory1
                                                            , Properties.Resources.Convert_SecureString_TS_InArgDesc1);

            DesignerMetadata.RegisterHelper(ref builder, t, nameof(Result)
                                                , Properties.Resources.Convert_SecureString_TS_OutArgName1
                                                , Properties.Resources.Convert_SecureString_TS_OutArgCategory1
                                                , Properties.Resources.Convert_SecureString_TS_OutArgDesc1);

        }

        protected override void Execute(CodeActivityContext context)
        {
            Result.Set(context, Marshal.PtrToStringUni(Marshal.SecureStringToGlobalAllocUnicode(InputValue.Get(context))));
        }
    }
}
