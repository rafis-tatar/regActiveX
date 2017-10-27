using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

namespace regActiveX
{
    static class Utility
    {
        /// <summary>
        /// Проверяет , является ли текущий пользователь администратором
        /// </summary>
        /// <returns></returns>
        public static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        /// <summary>
        ///проверяет является ли данный тип доступным для COM
        /// </summary>
        /// <param name="type">Тип</param>
        /// <returns></returns>
        public static bool IsComVisible(Type type)
        {
            ComVisibleAttribute[] attributes = (ComVisibleAttribute[])type.GetCustomAttributes(typeof(ComVisibleAttribute), false);
            //InterfaceTypeAttribute[] attributes2 = (InterfaceTypeAttribute[]) type.GetCustomAttributes(typeof(InterfaceTypeAttribute), false);
            return (attributes.Length > 0 && attributes[0].Value);// || attributes2.Length > 0;
        }

        public static ComSourceInterfacesAttribute GetComSourceInterfaces(Type type)
        {
            ComSourceInterfacesAttribute comSourceInterfacesAttribute = null;
            var attributes = (ComSourceInterfacesAttribute[])type.GetCustomAttributes(typeof(ComSourceInterfacesAttribute), false);
            if (attributes.Length > 0)
            {
                comSourceInterfacesAttribute = attributes[0];
            }
            return comSourceInterfacesAttribute;
        }
        public static string GetProgID(Type type)
        {
            string progId = type.FullName;
            if (IsComVisible(type))
            {
                var attributes = type.GetCustomAttributes(typeof(ProgIdAttribute), true);
                if (attributes.Length > 0)
                {
                    var progIdAttrib = attributes[0] as ProgIdAttribute;
                    progId = progIdAttrib?.Value;
                }
            }
            return progId;
        }
        /// <summary>
        ///проверяет является ли данный тип делегатом
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsDelegate(this Type type)
        {
            return typeof(MulticastDelegate).IsAssignableFrom(type.BaseType);
        }

    }
}
