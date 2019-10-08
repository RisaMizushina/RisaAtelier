using System;
using System.Runtime.InteropServices;

namespace RisaAtelier.MultipurposeLib
{
    /// <summary>
    /// COM操作の、汎用クラスです
    /// </summary>
    internal static class COMUtil
    {
        /// <summary>
        /// COM相互運用の、Unmanaged Resourceを、解放します
        /// </summary>
        /// <param name="target"></param>
        /// <param name="isFinal"></param>
        internal static void COMRelease(ref object target, bool isFinal)
        {
            try
            {
                if (isFinal)
                {
                    Marshal.FinalReleaseComObject(target);
                }
                else
                {
                    Marshal.ReleaseComObject(target);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (NullReferenceException) { }
            catch (ObjectDisposedException) { }
            catch (InvalidComObjectException) { }
        }

        /// <summary>
        /// COMオブジェクトの、メソッドを呼びます
        /// </summary>
        /// <param name="target"></param>
        /// <param name="nameOf"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        internal static object InvokeByMethod(ref object target, string nameOf, object[] args)
        {
            return target.GetType().InvokeMember(nameOf, System.Reflection.BindingFlags.InvokeMethod, null, target, args);
        }

        /// <summary>
        /// COMオブジェクトの、プロパティを設定します
        /// </summary>
        /// <param name="target"></param>
        /// <param name="nameOf"></param>
        /// <param name="value"></param>
        internal static void InvokeSetProperty(ref object target, string nameOf, object[] value)
        {
            target.GetType().InvokeMember(nameOf, System.Reflection.BindingFlags.SetProperty, null, target, value);
        }

        /// <summary>
        /// COMオブジェクトの、プロパティを取得します
        /// </summary>
        /// <param name="target"></param>
        /// <param name="nameOf"></param>
        /// <returns></returns>
        internal static object InvokeGetProperty(ref object target, string nameOf)
        {
            return target.GetType().InvokeMember(nameOf, System.Reflection.BindingFlags.GetProperty, null, target, null);
        }
    }
}
