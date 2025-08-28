using System.Collections.Generic;

namespace ComHookLib.Dto
{
    internal class CoGetClassObjectDto : ComEventBase
    {
        public string api { get; set; } = "CoGetClassObject";

        public string clsid { get; set; }         // "{...}"
        public string progId { get; set; }        // pode vir null e será resolvido
        public string iid { get; set; }           // "{...}"
        public string iid_name { get; set; }      // preenchido no EnrichIfComEvent
        public string clsctx { get; set; }        // "INPROC_SERVER|..."; normalizado pelo enrichment
        public string hr { get; set; }            // "0x........"
        public string hr_name { get; set; }       // preenchido no enrichment
        public double elapsed_ms { get; set; }
        public string kind { get; set; }          // "ftaerel"/"maybe" via enrichment
    }
}
