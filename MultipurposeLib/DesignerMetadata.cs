using System;
using System.Activities.Presentation.Metadata;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RisaAtelier.MultipurposeLib
{
    /// <summary>
    /// 文字列リソースファイルの読み込み・設定
    /// </summary>
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();

            // TODO: クラス（Activity）を追加したら、ここから初期設定をさせます
            DataTable.CopyToClipboard.SetMetaData(ref builder);
            DataTable.JointTable.SetMetaData(ref builder);
            DataTable.ShiftRowsAndColumns.SetMetaData(ref builder);

            Excel.ReadRecommendSetting.SetMetaData(ref builder);

            File.WaitForFileGrowthCompleted.SetMetaData(ref builder);

            MetadataStore.AddAttributeTable(builder.CreateTable());
        }

        /// <summary>
        /// Category、Displayname、Descriptionをセットで設定します
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="t"></param>
        /// <param name="nameOf"></param>
        /// <param name="DisplayName"></param>
        /// <param name="Category"></param>
        /// <param name="Description"></param>
        public static void RegisterHelper(ref AttributeTableBuilder builder, Type t, string nameOf, string DisplayName, string Category, string Description)
        {
            // \nのReplaceを、こちらで、アドホックに対応しているのですが、リソースファイルに、きれいに埋め込む方法は・・・？

            builder.AddCustomAttributes(t, nameOf, new DisplayNameAttribute(DisplayName));
            builder.AddCustomAttributes(t, nameOf, new CategoryAttribute(Category));
            builder.AddCustomAttributes(t, nameOf, new DescriptionAttribute(Description.Replace(@"\n", Environment.NewLine)));
        }

        /// <summary>
        /// Category、Displayname、Descriptionをセットで設定します
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="t"></param>
        /// <param name="DisplayName"></param>
        /// <param name="Category"></param>
        /// <param name="Description"></param>
        public static void RegisterHelper(ref AttributeTableBuilder builder, Type t, string DisplayName, string Category, string Description)
        {
            builder.AddCustomAttributes(t, new DisplayNameAttribute(DisplayName));
            builder.AddCustomAttributes(t, new CategoryAttribute(Category));
            builder.AddCustomAttributes(t, new DescriptionAttribute(Description.Replace(@"\n", Environment.NewLine)));
        }
    }


}
