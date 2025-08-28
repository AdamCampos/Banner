using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace OcxProbe
{
    internal static class RcwInspector
    {
        public static string DumpType(Type t)
        {
            var sb = new StringBuilder();

            sb.AppendLine("=== " + t.FullName + " ===");
            sb.AppendLine("Kind: " + (t.IsInterface ? "Interface" : (t.IsClass ? "Class" : t.MemberType.ToString())));
            if (t.GUID != Guid.Empty) sb.AppendLine("GUID: " + t.GUID);

            var it = (InterfaceTypeAttribute)Attribute.GetCustomAttribute(t, typeof(InterfaceTypeAttribute));
            if (it != null) sb.AppendLine("InterfaceType: " + it.Value);

            var ifaces = t.GetInterfaces();
            if (ifaces.Length > 0)
                sb.AppendLine("Implements: " + string.Join(", ", ifaces.Select(i => i.FullName)));

            var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.FlattenHierarchy);
            if (props.Length > 0)
            {
                sb.AppendLine("\n-- Properties --");
                foreach (var p in props.OrderBy(p => p.Name))
                {
                    var get = p.GetGetMethod();
                    var set = p.GetSetMethod();
                    var disp = (DispIdAttribute)(get != null
                                ? Attribute.GetCustomAttribute(get, typeof(DispIdAttribute))
                                : (set != null ? Attribute.GetCustomAttribute(set, typeof(DispIdAttribute)) : null));

                    sb.Append("  ");
                    if (disp != null) sb.Append("[DispId " + disp.Value + "] ");
                    sb.AppendLine(p.PropertyType.Name + " " + p.Name + " { " +
                                  (get != null ? "get; " : "") + (set != null ? "set; " : "") + "}");
                }
            }

            var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.FlattenHierarchy)
                           .Where(m => !m.IsSpecialName) // ignora get_/set_/add_/remove_
                           .OrderBy(m => m.Name)
                           .ToArray();
            if (methods.Length > 0)
            {
                sb.AppendLine("\n-- Methods --");
                foreach (var m in methods)
                {
                    var disp = (DispIdAttribute)Attribute.GetCustomAttribute(m, typeof(DispIdAttribute));
                    var pars = m.GetParameters().Select(FormatParam).ToArray();

                    sb.Append("  ");
                    if (disp != null) sb.Append("[DispId " + disp.Value + "] ");
                    sb.AppendLine(m.ReturnType.Name + " " + m.Name + "(" + string.Join(", ", pars) + ")");
                }
            }

            var evs = t.GetEvents(BindingFlags.Public | BindingFlags.Instance | BindingFlags.FlattenHierarchy)
                       .OrderBy(e => e.Name)
                       .ToArray();
            if (evs.Length > 0)
            {
                sb.AppendLine("\n-- Events --");
                foreach (var e in evs)
                    sb.AppendLine("  event " + e.EventHandlerType.Name + " " + e.Name);
            }

            sb.AppendLine();
            return sb.ToString();
        }

        private static string FormatParam(ParameterInfo p)
        {
            var sb = new StringBuilder();
            if (p.ParameterType.IsByRef) sb.Append(p.IsOut ? "out " : "ref ");
            if (p.IsOptional) sb.Append("optional ");

            var t = p.ParameterType.IsByRef ? p.ParameterType.GetElementType() : p.ParameterType;
            if (t == null) t = typeof(object);

            sb.Append(t.Name).Append(" ").Append(p.Name);
            return sb.ToString();
        }
    }
}
