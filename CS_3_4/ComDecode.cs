// ComHookLib - ComDecode.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ComHookLib
{
    /// <summary>
    /// Utilitários de normalização/legibilidade.
    /// </summary>
    public static class ComDecode
    {
        private static readonly Regex SepRegex = new Regex(@"[\s\|\+\,/\\]+", RegexOptions.Compiled);

        /// <summary>
        /// Normaliza CLSCTX para 1 linha, tokens únicos e ordenados.
        /// Entrada tolerante (null, vazio, quebras, "|", "+", ",", etc).
        /// </summary>
        public static string NormalizeClsctx(string clsctx)
        {
            if (string.IsNullOrWhiteSpace(clsctx))
                return clsctx; // não força nada; só evita exceção

            // remove quebras e normaliza separadores
            var cleaned = clsctx.Replace("\r", " ").Replace("\n", " ");

            // tokeniza por vários separadores
            var tokens = SepRegex
                .Split(cleaned)
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Select(t => t.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(t => t, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            return string.Join("|", tokens);
        }

        /// <summary>
        /// Reduz múltiplos espaços/quebras para uma linha só.
        /// </summary>
        public static string OneLine(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return Regex.Replace(s, @"\s+", " ").Trim();
        }
    }
}
