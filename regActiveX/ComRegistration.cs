using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace regActiveX
{
    /// <summary>
    ///Регистрация  сборки как COM
    /// </summary>
    class ComRegistration
    {
        
        /// <summary>
        /// Отменить регистрацию сборки
        /// </summary>
        /// <param name="filepath">путь к сборке .dll</param>
        public void UnregisterCOM(string filepath)
        {
            Assembly assembly = Assembly.LoadFrom(filepath);
            Log("\nCancel registration of Assemblies " + filepath);
            foreach (Type type in assembly.GetExportedTypes())
            {
                foreach (RegistryHive baseKey in GetBaseKeys())
                {
                    var key = Unregister(Utility.GetProgID(type), baseKey, Utility.IsComVisible(type));
                    var d = type.GetMethods().Where(o => o.GetCustomAttributes(typeof(ComUnregisterFunctionAttribute), false).Any()).ToArray();
                    if (d.Any())
                    {
                        var str = $"{key}\\CLSID\\{type.GUID:B}";
                        foreach (MethodInfo methodInfo in d)
                        {
                            methodInfo.Invoke(type, new object[1] { str });
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Регистрация сборки 
        /// </summary>
        /// <param name="filepath">путь к сборке .dll</param>
        public void RegisterCOM(string filepath)
        {
            var assembly = Assembly.LoadFrom(filepath);
            var types = assembly.GetExportedTypes();
            Log("\nRegister Assembly " + filepath);
            Log("");
            foreach (var type in types)//.Where(Utility.IsComVisible))
            {
                bool isControl = Utility.IsComVisible(type) && (type.BaseType == typeof(System.Windows.Forms.Control) || type.BaseType == typeof(System.Windows.Forms.UserControl));
                var key = RegisterCOM(type, isControl);
                var d = type.GetMethods().Where(o => o.GetCustomAttributes(typeof(ComRegisterFunctionAttribute), false).Any()).ToArray();
                if (d.Any())
                {
                    var str = $"{key}\\CLSID\\{type.GUID:B}";
                    foreach (MethodInfo methodInfo in d)
                    {
                        methodInfo.Invoke(type, new object[1] { str });
                    }
                }
            }
            Log("");
        }
        /// <summary>
        ///Проверка регистрации сборки
        /// </summary>
        /// <param name="filepath"></param>
        public bool CheckClassCOM(string filepath)
        {
            var assembly = Assembly.LoadFrom(filepath);
            var types = assembly.GetExportedTypes().Where(o => !o.IsInterface && Utility.IsComVisible(o) && !o.IsDelegate() && !o.IsNested).ToArray();

            Log("\nCheck assembly registration " + filepath);
            Log("");
            bool[] regAllTypes = new bool[types.Length];
            for (int i = 0; i < types.Length; i++)
            {
                regAllTypes[i] = CheckClass(types[i]);
            }
            Log("");
            return regAllTypes.All(o => o);
        }

        private bool CheckClass(Type type)
        {
            var progID = Utility.GetProgID(type);
            var basekey = Utility.IsAdministrator() ? RegistryHive.LocalMachine : RegistryHive.ClassesRoot;
            var keyPath = GetKeyPath(basekey);

            var regularx86View = RegistryKey.OpenBaseKey(basekey, RegistryView.Registry32);
            var keys = new RegistryKey[Environment.Is64BitOperatingSystem ? 2 : 1];
            keys[0] = string.IsNullOrWhiteSpace(keyPath) ? regularx86View : regularx86View.OpenSubKey(keyPath);

            if (Environment.Is64BitOperatingSystem)
            {
                var regularx64View = RegistryKey.OpenBaseKey(basekey, RegistryView.Registry64);
                keys[1] = string.IsNullOrWhiteSpace(keyPath) ? regularx64View : regularx64View.OpenSubKey(keyPath);
            }

            bool[] regAllplatform = new bool[keys.Length];
            for (int i = 0; i < regAllplatform.Length; i++)
            {
                string logmsg = string.Empty;
                RegistryKey keyProgID = keys[i].OpenSubKey(progID);
                regAllplatform[i] = keyProgID != null;
                logmsg = regAllplatform[i] ? $"{progID} REGISTERED IN {keyProgID?.View}:[{keyProgID?.Name}]": $"{progID} IN {keys[i]?.View} UNREGISTERED";
                Log(logmsg);
            }
            Log("");
            return regAllplatform.All(o => o);
        }
        private static RegistryHive[] GetBaseKeys()
        {
            RegistryHive[] basekey = new RegistryHive[Utility.IsAdministrator() ? 2 : 1];
            basekey[0] = RegistryHive.CurrentUser;
            if (Utility.IsAdministrator())
            {
                basekey[0] = RegistryHive.LocalMachine;
                basekey[1] = RegistryHive.CurrentUser;
            }
            return basekey;
        }

        private static string GetKeyPath(RegistryHive key)
        {
            string path = "Software\\Classes\\";
            if (key == RegistryHive.ClassesRoot) path = string.Empty;
            return path;
        }
        private static string RegisterCOM(Type type, bool isControl = true)
        {
            var progID = Utility.GetProgID(type);
            var guidStr = "{" + type.GUID.ToString() + "}";

            var basekey = GetBaseKeys()[0];
            var keyPath = GetKeyPath(basekey);

            Log("Register " + progID);

            string typeLibGuid;

            var regularx86View = RegistryKey.OpenBaseKey(basekey, RegistryView.Registry32);
            var keys = new RegistryKey[Environment.Is64BitOperatingSystem ? 2 : 1];
            keys[0] = regularx86View.OpenSubKey(keyPath, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
            var retKeyStr = keys[0]?.Name;
            if (Environment.Is64BitOperatingSystem)
            {
                var regularx64View = RegistryKey.OpenBaseKey(basekey, RegistryView.Registry64);
                keys[1] = regularx64View.OpenSubKey(keyPath, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                retKeyStr = keys[1]?.Name;
                typeLibGuid = SetTypeLib(keys[1], type);
            }
            else
            {
                typeLibGuid = SetTypeLib(keys[0], type);
            }

            foreach (RegistryKey rootKey in keys)
            {
                //var comSourceI = Utility.GetComSourceInterfaces(type);
                //if (comSourceI != null)
                //{
                //    var eventType = type.Assembly.GetType(comSourceI.Value);
                //    SetInterface(rootKey, eventType, typeLibGUID);
                //}

                if (!Utility.IsComVisible(type))
                {
                    if (type.IsInterface || type.IsDelegate() || type.IsNested)
                    {
                        SetInterface(rootKey, type, typeLibGuid);
                    }
                    continue;
                }

                /*****************************************************************
                * [HKEY_CURRENT_USER\Software\Classes\Prog.ID]="Namespace.Class" *
                *****************************************************************/
                var keyProgID = rootKey.CreateSubKey(progID);
                keyProgID.SetValue(null, type.FullName);

                //[HKEY_CURRENT_USER\Software\Classes\Prog.ID\CLSID]
                //@="{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"
                RegistryKey prtidclsid = keyProgID.CreateSubKey(@"CLSID");
                prtidclsid.SetValue(null, guidStr);

                //[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}]
                //@="Namespace.Class
               
                RegistryKey clsid = rootKey.CreateSubKey("CLSID");
                RegistryKey keyCLSID = clsid.CreateSubKey(guidStr);
                //keyCLSID.SetValue(null, type.FullName); 
                keyCLSID.SetValue(null, progID);

                if (isControl)
                {
                    //[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\Control]
                    // для того что бы видно было как OleControl 
                    var keycontrol = keyCLSID.CreateSubKey("Control");
                    keycontrol.Close();
                }

                //[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\TypeLib]
                // @="typeLibGUID"
                var keyTypeLibGuid = keyCLSID.CreateSubKey("TypeLib");
                keyTypeLibGuid.SetValue(null, typeLibGuid);
                keyTypeLibGuid.Close();

                //[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\ProgId]
                //@="Prog.ID"
                keyCLSID.CreateSubKey("ProgId").SetValue(null, progID);

                //[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\InprocServer32]
                //@="mscoree.dll"
                //"ThreadingModel"="Both"
                //"Class"="Namespace.Class"
                //"Assembly"="AssemblyName, Version=1.0.0.0, Culture=neutral, PublicKeyToken=71c72075855a359a"
                //"RuntimeVersion"="v4.0.30319"
                //"CodeBase"="file:///Drive:/Full/Image/Path/file.dll"
                RegistryKey InprocServer32 = keyCLSID.CreateSubKey("InprocServer32");
                SetInprocServer(InprocServer32, type, false);

                //[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\InprocServer32\1.0.0.0]
                //"Class"="Namespace.Class"
                //"Assembly"="AssemblyName, Version=1.0.0.0, Culture=neutral, PublicKeyToken=71c72075855a359a"
                //"RuntimeVersion"="v4.0.30319"
                //"CodeBase"="file:///Drive:/Full/Image/Path/file.dll"
                SetInprocServer(InprocServer32.CreateSubKey("Version"), type, true);

                //[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}]
                keyCLSID.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");
                Log($"Class {progID} registered in to {rootKey.View}:[{keyCLSID.Name}]");
                keyCLSID.Close();
            }
            Log("");
            return retKeyStr;
        }
        private static string Unregister(string progId, RegistryHive basekey, bool withLog=true)
        {
            var retkeyStr = string.Empty;
            var keyPath = GetKeyPath(basekey);
            var keys = new RegistryKey[Environment.Is64BitOperatingSystem ? 2 : 1];

            var regularx86View = RegistryKey.OpenBaseKey(basekey, RegistryView.Registry32);
            keys[0] = regularx86View;
            retkeyStr = keys[0].Name;

            if (Environment.Is64BitOperatingSystem)
            {
                var regularx64View = RegistryKey.OpenBaseKey(basekey, RegistryView.Registry64);
                keys[1] = regularx64View;//.OpenSubKey(keyPath, RegistryKeyPermissionCheck.ReadWriteSubTree,System.Security.AccessControl.RegistryRights.FullControl);
                retkeyStr = keys[1].Name;
            }
            string clsid = string.Empty;
            foreach (RegistryKey key in keys)
            {
                var rootKey = key.OpenSubKey(keyPath, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                if (withLog)
                {
                    Log("");
                    Log($"Open key {key.View} [{rootKey.Name}\\{progId}]");
                }
                RegistryKey keyClass = null;
                try
                {
                    keyClass = rootKey.OpenSubKey(progId, true);
                }
                catch
                {
                    continue;
                }
                //******************* удалив ключ из ветки HKEY_CURRENT_USER\Software\Classes (progID) сохраняем clsid в переменной 
                //*******************для удаления этого clsid в других ветках реестра (HKCU,HKLM и.т.д , а также в других ветках разрядности (x64))
                if (keyClass == null && string.IsNullOrWhiteSpace(clsid))
                {
                    if (withLog) Log(string.Format("CLSID for {0} not found...", progId));
                    continue;
                }

                if (keyClass != null)
                {
                    clsid = (string)keyClass.OpenSubKey("CLSID").GetValue("");
                    if (string.IsNullOrWhiteSpace(clsid)) continue;
                }
                //***************************

                var keyCLSID = rootKey.OpenSubKey("CLSID", true);
                if (keyCLSID == null) continue;

                Log($"Open key [{keyCLSID.Name}\\{clsid}]");
                var key_clsid = keyCLSID.OpenSubKey(clsid);

                if (key_clsid != null)
                {
                    var typeLibkey = key_clsid.OpenSubKey("TypeLib");
                    if (typeLibkey != null)
                    {

                        string typeLib = (string)typeLibkey.GetValue("");
                        
                        Log(string.Format("TypeLib {1} from [{0}]", key_clsid.Name, typeLib));
                        UnregisterInterface(keys, typeLib, keyPath);
                        UnregistrTypeLib(keys, typeLib, keyPath);
                    }
                    else
                    {
                        Log($"TypeLib from [{key_clsid.Name}] not found");
                    }

                    if (!DeleteKeyAndCheck(keyCLSID, clsid))
                        Log(string.Format("not deleted key {1} from [{0}]", keyCLSID.Name, clsid));
                }

                if (!DeleteKeyAndCheck(rootKey, progId))
                    Log(string.Format("not deleted key {1} from [{0}]", rootKey.Name, progId));

                rootKey.Close();
            }

            if (withLog) Log("");
            return retkeyStr;
        }

        private static void UnregisterInterface(RegistryKey[] keys, string typLibGUID, string keyPath)
        {
            foreach (var key in keys)
            {
                var rootKey = key.OpenSubKey(keyPath, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                RegistryKey keyInterface = null;
                try
                {

                    keyInterface = rootKey.OpenSubKey("INTERFACE", true);
                }
                catch
                {
                    continue;
                }

                foreach (string name in keyInterface.GetSubKeyNames())
                {
                    RegistryKey keyint = null;
                    try
                    {
                        keyint = keyInterface.OpenSubKey(name, true);
                    }
                    catch
                    {
                        continue;
                    }

                    if (keyint == null)
                    {
                        continue;
                    }

                    RegistryKey _Ityplib = null;
                    try
                    {
                        _Ityplib = keyint.OpenSubKey("TypeLib", true);
                    }
                    catch 
                    {
                        continue;
                    }

                    if (_Ityplib != null)
                    {
                        var Ityplib = (string)_Ityplib.GetValue(null);
                        if (Ityplib != null && Ityplib == typLibGUID)
                        {
                            keyint.Close();
                            Log(string.Format("delete key {1} from {2}:[{0}]", keyInterface.Name, name, key.View));
                            keyInterface.DeleteSubKeyTree(name);
                            continue;
                        }
                    }
                    keyint.Close();
                }
                keyInterface.Close();
            }
        }
        private static void UnregistrTypeLib(RegistryKey[] keys, string typeLib, string keyPath)
        {
            foreach (var key in keys)
            {
                var rootKey = key.OpenSubKey(keyPath, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                RegistryKey keyTypeLib = null;
                try
                {
                    keyTypeLib = rootKey.OpenSubKey("TypeLib", true);
                }
                catch
                {
                    continue;
                }

                if (keyTypeLib == null) continue;

                Log(string.Format("delete key {1} from {2}:[{0}]", keyTypeLib.Name, typeLib, key.View));
                keyTypeLib.DeleteSubKeyTree(typeLib, false);
                keyTypeLib.Close();
            }
        }

        private static void SetInterface(RegistryKey key, Type type, string typeLibguid)
        {
            var interfaceName = type.Name.StartsWith("I") ? type.Name : "_" + type.Name;
            Log(string.Format("Open key [{0}\\{1}]", key.Name, "Interface"));
            var keyInteface = key.CreateSubKey("Interface");

            if (CheckInterface(keyInteface, interfaceName, typeLibguid))
            {
                Log(string.Format("Found registered Interface {0} in to {2}[{1}] and be deleted", interfaceName, keyInteface.Name, key.View));
            }

            var guidIntrface = Guid.NewGuid().ToString("B");
            keyInteface = keyInteface.CreateSubKey(guidIntrface);
            keyInteface.SetValue(null, interfaceName);
            var ProxyStubClsid32 = keyInteface.CreateSubKey("ProxyStubClsid32");

            var intefaceTypeAttr = type.GetCustomAttributes(typeof(InterfaceTypeAttribute), false).OfType<InterfaceTypeAttribute>().ToList();
            if (intefaceTypeAttr != null && intefaceTypeAttr.Any())
            {
                ProxyStubClsid32.SetValue(null, "{00020420-0000-0000-C000-000000000046}");
            }
            else
            {
                ProxyStubClsid32.SetValue(null, "{00020424-0000-0000-C000-000000000046}");
            }
            ProxyStubClsid32.Close();

            var TypeLib = keyInteface.CreateSubKey("TypeLib");
            TypeLib.SetValue(null, typeLibguid);
            TypeLib.SetValue("Version", type.Assembly.GetName().Version.ToString(2));
            TypeLib.Close();

            Log(string.Format("Interface {0} registered in to [{1}]", interfaceName, keyInteface.Name));
            keyInteface.Close();

        }
        private static bool CheckInterface(RegistryKey interfaceKey, string name, string typeLibGuid)
        {
            foreach (var key in interfaceKey.GetSubKeyNames())
            {
                var subkey = interfaceKey.OpenSubKey(key);
                if (subkey == null) continue;

                var nm = (string)subkey.GetValue(null);
                if (string.IsNullOrWhiteSpace(nm))
                {
                    subkey.Close();
                    continue;
                }


                var subkey_typlib = subkey.OpenSubKey("TypeLib");
                if (subkey_typlib == null) continue;

                var tl = (string)subkey_typlib.GetValue(null);
                if (string.IsNullOrWhiteSpace(tl))
                {
                    subkey_typlib.Close();
                    continue;
                }

                subkey_typlib.Close();
                subkey.Close();
                if (nm.ToLower().Trim() == name.ToLower().Trim() && tl.ToLower().Trim() == typeLibGuid.ToLower().Trim())
                {
                    interfaceKey.DeleteSubKeyTree(key);
                    return true;
                }
            }
            return false;
        }

        private static string SetTypeLib(RegistryKey key, Type type)
        {
            var tlbdir = System.IO.Path.GetDirectoryName(type.Assembly.Location);
            var tlbFullname = tlbdir + "\\" + System.IO.Path.GetFileNameWithoutExtension(type.Assembly.Location) + ".tlb";
            if (!System.IO.File.Exists(tlbFullname))
            {
                //throw new FileNotFoundException(string.Format("невозможно найти файл: {0}", tlbFullname));
                var prc = Process.Start("TlbExp.exe", $"{type.Assembly.Location} /out:{tlbFullname}");
            }

            if (!System.IO.File.Exists(tlbFullname))
            {
                throw new FileNotFoundException(string.Format("невозможно найти файл: {0}", tlbFullname));
            }

            var oldguid = CheckTypeLib(key, type, tlbFullname);
            if (!string.IsNullOrWhiteSpace(oldguid))
            {
                Log(string.Format("Found registered TypeLib {0} key in to {2}[{1}\\TypeLib\\{3}]", tlbFullname, key.Name, key.View, oldguid));
                return oldguid;
            }

            var guidTypeLib = Guid.NewGuid().ToString("B");
            RegistryKey keyTypeLib = key.CreateSubKey(string.Format("TypeLib\\{0}\\1.0", guidTypeLib));
            keyTypeLib.SetValue(null, type.Assembly.GetName().Name);
            keyTypeLib.CreateSubKey("FLAGS").SetValue(null, "0");
            keyTypeLib.CreateSubKey("HELPDIR").SetValue(null, tlbdir);
            keyTypeLib.CreateSubKey("0\\win32").SetValue(null, tlbFullname);

            Log(string.Format("TypeLib {0} registered in to {3}[{1}\\{2}]", guidTypeLib, key.Name, "TypeLib", key.View));
            keyTypeLib.Close();
            return guidTypeLib;
        }
        private static string CheckTypeLib(RegistryKey key, Type type, string fullpath)
        {
            RegistryKey keyTypeLib = key.CreateSubKey("TypeLib");
            foreach (var sname in keyTypeLib.GetSubKeyNames())
            {
                var subkey = keyTypeLib.OpenSubKey(sname + "\\1.0");
                if (subkey == null) continue;

                var win32key = subkey.OpenSubKey("0\\win32");
                if (win32key == null) continue;

                var assmName = (string)subkey.GetValue(null);
                var assFulPat = (string)win32key.GetValue(null);

                if (type.Assembly.GetName().Name == assmName && assFulPat == fullpath)
                {
                    //keyTypeLib.DeleteSubKeyTree(sname);
                    return sname;
                }
            }
            return null;
        }

        private static void SetInprocServer(RegistryKey key, Type type, bool versionNode)
        {
            if (!versionNode)
            {
                key.SetValue(null, "mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
            }

            key.SetValue("Class", type.FullName);
            key.SetValue("Assembly", type.Assembly.FullName);
            key.SetValue("RuntimeVersion", type.Assembly.ImageRuntimeVersion);
            key.SetValue("CodeBase", type.Assembly.CodeBase);
        }

        private static bool DeleteKeyAndCheck(RegistryKey key, string subkeyName)
        {
            key.DeleteSubKeyTree(subkeyName, false);
            Log(string.Format("delete key {1} from [{0}]", key.Name, subkeyName));
            RegistryKey chekkey = null;
            try
            {
                chekkey = key.OpenSubKey(subkeyName);
            }
            catch
            {
                return false;
            }
            return chekkey == null;
        }
        private static void Log(string msg)
        {
            if (Debugger.IsAttached)
                Debugger.Log(1, "RegComActiveX", msg + "\n");
            else
                Console.WriteLine(msg);
        }
    }

}
