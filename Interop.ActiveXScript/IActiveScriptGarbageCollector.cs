namespace Interop.ActiveXScript
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Provides a method to start garbage collection. 
    /// This interface should be implemented by Active Script engines that want to clean up their resources.
    /// <see cref="http://msdn.microsoft.com/en-us/library/hh769816(v=vs.94).aspx"/>
    /// </summary>
    [ComImport]
    [Guid("6aa2c4a0-2b53-11d4-a2a0-00104bd35090")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptGarbageCollector
    {
        /// <summary>
        /// The Active Script host calls this method to start garbage collection.
        /// <see cref="http://msdn.microsoft.com/en-us/library/hh769822(v=vs.94).aspx"/>
        /// </summary>
        /// <param name="type">The SCRIPTGCTYPE Enumeration that specifies whether to do normal
        /// or exhaustive garbage collection.</param>
        void CollectGarbage(
            [In] ScriptGCType type
        );
    }
}
