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
            return (attributes.Length > 0 && attributes[0].Value);
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
