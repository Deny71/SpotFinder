using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace SpotFinder.Properties

{
    public class FormatCheck
    {
        public string val;

        public static string check52(List<Format52>ilosci)
        {
            string val="not_empty";
            if (
                ilosci[0].tv_Sam <= 0 &&
                ilosci[0].tv_LG <= 0 &&
                ilosci[0].tv_Sony <= 0 &&
                ilosci[0].tv_Sha <= 0 &&
                ilosci[0].laptop <= 0 &&
                ilosci[0].tel_Sam <= 0 &&
                ilosci[0].tel_Mot <= 0 &&
                ilosci[0].pra_Sam <= 0 &&
                ilosci[0].pra_Whi <= 0 &&
                ilosci[0].kuc_Ami <= 0 &&
                ilosci[0].lod_Sam <= 0 &&
                ilosci[0].lod_Bek <= 0 &&
                ilosci[0].susz <= 0 &&
                ilosci[0].oczysz <= 0 &&
                ilosci[0].odk <= 0 &&
                ilosci[0].eksp <= 0 &&
                ilosci[0].szczot <= 0 &&
                ilosci[0].paro <= 0)
            {
                val = "empty";
            }

            return val;
        }

        public static string check63(List<Format52>ilosci)
        {
            string val="not_empty";
            if (
                ilosci[1].tv_Sam <= 0 &&
                ilosci[1].tv_LG <= 0 &&
                ilosci[1].tv_Sony <= 0 &&
                ilosci[1].tv_Sha <= 0 &&
                ilosci[1].laptop <= 0 &&
                ilosci[1].tel_Sam <= 0 &&
                ilosci[1].tel_Mot <= 0 &&
                ilosci[1].pra_Sam <= 0 &&
                ilosci[1].pra_Whi <= 0 &&
                ilosci[1].kuc_Ami <= 0 &&
                ilosci[1].lod_Sam <= 0 &&
                ilosci[1].lod_Bek <= 0 &&
                ilosci[1].susz <= 0 &&
                ilosci[1].oczysz <= 0 &&
                ilosci[1].odk <= 0 &&
                ilosci[1].eksp <= 0 &&
                ilosci[1].szczot <= 0 &&
                ilosci[1].paro <= 0)
            {
                val = "empty";
            }

            return val;
        }
        public static string check1234(List<Format52>ilosci)
        {
            string val="not_empty";
            if (
                ilosci[2].tv_Sam <= 0 &&
                ilosci[2].tv_LG <= 0 &&
                ilosci[2].tv_Sony <= 0 &&
                ilosci[2].tv_Sha <= 0 &&
                ilosci[2].laptop <= 0 &&
                ilosci[2].tel_Sam <= 0 &&
                ilosci[2].tel_Mot <= 0 &&
                ilosci[2].pra_Sam <= 0 &&
                ilosci[2].pra_Whi <= 0 &&
                ilosci[2].kuc_Ami <= 0 &&
                ilosci[2].lod_Sam <= 0 &&
                ilosci[2].lod_Bek <= 0 &&
                ilosci[2].susz <= 0 &&
                ilosci[2].oczysz <= 0 &&
                ilosci[2].odk <= 0 &&
                ilosci[2].eksp <= 0 &&
                ilosci[2].szczot <= 0 &&
                ilosci[2].paro <= 0)
            {
                val = "empty";
            }

            return val;
        }
    }
}