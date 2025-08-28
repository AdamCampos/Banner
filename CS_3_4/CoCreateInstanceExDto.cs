using System.Collections.Generic;

namespace ComHookLib.Dto
{
    internal class CoCreateInstanceExDto : ComEventBase
    {
        public string api { get; set; } = "CoCreateInstanceEx";

        public string clsid { get; set; }
        public string progId { get; set; }
        public string clsctx { get; set; }
        public int count { get; set; }                // número de IIDs
        public List<string> iids { get; set; }        // lista de "{...}"
        public List<int> multiqi_hr { get; set; }     // HRESULTs de cada IID (opcional)
        public string hr { get; set; }
        public string hr_name { get; set; }
        public double elapsed_ms { get; set; }
        public string kind { get; set; }
        // opcionalmente: uma lista paralela com "iid_names" se quiser popular também
        public List<string> iid_names { get; set; }   // preenchida no enrichment, se desejar
    }
}
