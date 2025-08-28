using System;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using FTAlarmEventSummary;
// using static System.Windows.Forms.VisualStyles.VisualStyleElement; // não é necessário

namespace OcxProbe
{
    internal static class RcwEventProbe
    {
        public static string Run()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== RCW Event/Call Probe ===");
            sb.AppendLine("Process x86? " + (IntPtr.Size == 4));
            sb.AppendLine("CLR Version: " + Environment.Version);
            sb.AppendLine();

            ProbeOne("Summary", new AlarmEventSummaryClass(), sb);
            sb.AppendLine();
            ProbeOne("Banner", new AlarmEventBannerClass(), sb);

            // Exporta mapas DISPIDs (Summary/Banner) para apoiar o hook de IDispatch no alvo real
            try
            {
                var (j1, c1) = RcwDispatchMap.ExportSummary();
                var (j2, c2) = RcwDispatchMap.ExportBanner();

                sb.AppendLine();
                sb.AppendLine("+ DISPIDs exportados:");
                sb.AppendLine("  - " + j1);
                sb.AppendLine("  - " + c1);
                sb.AppendLine("  - " + j2);
                sb.AppendLine("  - " + c2);

            }
            catch (Exception ex)
            {
                sb.AppendLine("! Falha ao exportar DISPIDs: " + ex.Message);
            }

            // Grava após adicionar tudo ao buffer
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rcw_run.txt");
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            return path;
        }

        private static void ProbeOne(string tag, object comObj, StringBuilder sb)
        {
            sb.AppendLine("== " + tag + " ==");
            // 1) Assina todos os eventos dinamicamente
            AttachAllEvents(comObj, (evName, args) =>
            {
                sb.Append(tag).Append("::").Append(evName).Append(" (");
                for (int i = 0; i < args.Length; i++)
                {
                    if (i > 0) sb.Append(", ");
                    sb.Append(args[i] == null ? "null" : args[i] + " : " + args[i].GetType().FullName);
                }
                sb.AppendLine(")");
            });
            sb.AppendLine("+ OK: eventos assinados.");

            // 2) Propriedades e flags seguras (não exigem UI host)
            TryGetProp(comObj, "AutomaticUpdate", tag, sb);
            TrySetProp(comObj, "AutomaticUpdate", false, tag, sb);
            TrySetProp(comObj, "OptionsChanged", true, tag, sb); // setter-only em alguns casos

            // 3) Métodos “seguros” (podem falhar se faltar configuração – loga HRESULT)
            TryInvoke(comObj, "LoadMessages", tag, sb);       // pode requerer ConnectionString
            TryInvoke(comObj, "Refresh", tag, sb);            // comum no Summary/Banner
            TryInvoke(comObj, "ShowPropertyPages", tag, sb);  // pode abrir UI/do COM
        }

        private static void TryGetProp(object obj, string prop, string tag, StringBuilder sb)
        {
            try
            {
                var pi = obj.GetType().GetProperty(prop);
                if (pi != null && pi.CanRead)
                {
                    var val = pi.GetValue(obj, null);
                    sb.AppendLine($"+ GET {tag}.{prop} => {Format(val)}");
                }
                else
                {
                    sb.AppendLine($"- GET {tag}.{prop} não disponível.");
                }
            }
            catch (COMException ex)
            {
                sb.AppendLine($"! GET {tag}.{prop} COM 0x{ex.HResult:X8} {ex.Message}");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"! GET {tag}.{prop} EXC {ex.GetType().Name}: {ex.Message}");
            }
        }

        private static void TrySetProp(object obj, string prop, object value, string tag, StringBuilder sb)
        {
            try
            {
                var pi = obj.GetType().GetProperty(prop);
                if (pi != null && pi.CanWrite)
                {
                    pi.SetValue(obj, value, null);
                    sb.AppendLine($"+ SET {tag}.{prop} = {Format(value)}");
                }
                else
                {
                    // há props “set-only” expostas via DispId; tenta via setter mesmo se CanWrite=false
                    if (pi != null && !pi.CanWrite)
                        sb.AppendLine($"- SET {tag}.{prop} aparenta ser read-only (metadata).");
                    else
                        sb.AppendLine($"- SET {tag}.{prop} inexistente.");
                }
            }
            catch (COMException ex)
            {
                sb.AppendLine($"! SET {tag}.{prop} COM 0x{ex.HResult:X8} {ex.Message}");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"! SET {tag}.{prop} EXC {ex.GetType().Name}: {ex.Message}");
            }
        }

        private static void TryInvoke(object obj, string method, string tag, StringBuilder sb)
        {
            try
            {
                var mi = obj.GetType().GetMethod(method, new Type[0]);
                if (mi == null)
                {
                    sb.AppendLine($"- CALL {tag}.{method}() não encontrado.");
                    return;
                }
                mi.Invoke(obj, null);
                sb.AppendLine($"+ CALL {tag}.{method}()");
            }
            catch (TargetInvocationException tie) when (tie.InnerException is COMException cex)
            {
                sb.AppendLine($"! CALL {tag}.{method}() COM 0x{cex.HResult:X8} {cex.Message}");
            }
            catch (COMException ex)
            {
                sb.AppendLine($"! CALL {tag}.{method}() COM 0x{ex.HResult:X8} {ex.Message}");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"! CALL {tag}.{method}() EXC {ex.GetType().Name}: {ex.Message}");
            }
        }

        private static string Format(object v)
        {
            return v == null ? "null" : (v + " : " + v.GetType().FullName);
        }

        /// <summary>
        /// Assina todos os eventos de um RCW criando delegates com a assinatura exata do evento,
        /// mas encaminhando os argumentos para um sink (nome do evento, object[] args).
        /// Compatível com C# 7.3.
        /// </summary>
        public static void AttachAllEvents(object comObj, Action<string, object[]> sink)
        {
            var t = comObj.GetType();
            var events = t.GetEvents();
            foreach (var ev in events)
            {
                try
                {
                    var handlerType = ev.EventHandlerType;
                    var invoke = handlerType.GetMethod("Invoke");
                    var parms = invoke.GetParameters();

                    // Parâmetros tipados exatamente como o delegate do evento:
                    var paramExprs = parms.Select(p => Expression.Parameter(p.ParameterType, p.Name)).ToArray();

                    // Constrói object[] com boxing de cada parâmetro:
                    var boxed = paramExprs
                        .Select(p => (Expression)Expression.Convert(p, typeof(object)))
                        .ToArray();
                    var newArr = Expression.NewArrayInit(typeof(object), boxed);

                    // Chama sink.Invoke(evName, object[])
                    var sinkConst = Expression.Constant(sink);
                    var sinkInvoke = sink.GetType().GetMethod("Invoke");
                    var callSink = Expression.Call(
                        sinkConst,
                        sinkInvoke,
                        Expression.Constant(ev.Name),
                        newArr
                    );

                    // Lambda com o tipo EXATO do evento
                    var lambda = Expression.Lambda(handlerType, callSink, paramExprs).Compile();

                    try
                    {
                        ev.AddEventHandler(comObj, lambda);
                    }
                    catch (TargetInvocationException tie) // first-chance comum sem host específico
                    {
                        var inner = tie.InnerException;
                        Debug.WriteLine(
                            $"! EVENT {ev.DeclaringType?.Name}.{ev.Name} AddHandler TIE 0x{(inner as COMException)?.HResult:X8} {inner?.Message}");
                    }
                    catch (COMException cex)
                    {
                        Debug.WriteLine(
                            $"! EVENT {ev.DeclaringType?.Name}.{ev.Name} AddHandler COM 0x{cex.HResult:X8} {cex.Message}");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(
                            $"! EVENT {ev.DeclaringType?.Name}.{ev.Name} AddHandler EXC {ex.GetType().Name}: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"! EVENT {ev.DeclaringType?.Name}.{ev.Name} PrepareHandler EXC {ex.GetType().Name}: {ex.Message}");
                }
            }
        }
    }
}
