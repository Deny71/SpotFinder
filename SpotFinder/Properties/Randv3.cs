using System;
using System.Collections.Generic;

namespace SpotFinder.Properties
{
    public class Randv3
    {

        private static Random variable = new Random();
        private static int value;

        public static string losowacz63(int value, List<Format52> ilosci, List<Record> formatka, int i)
        {

            string napis = "test";
            if (value == 1)
            {
                ilosci[1].tv_Sam--;

                if (ilosci[1].tv_Sam <= 0)
                {
                    return "again";
                }
                else
                    return "Tv_Samsung";

            }
            else if (value == 2)
            {
                ilosci[1].tv_LG--;

                if (ilosci[1].tv_LG <= 0)
                {
                    return "again";
                }
                else
                    return "Tv_LG";
            }
            else if (value == 3)
            {
                ilosci[1].tv_Sony--;

                if (ilosci[1].tv_Sony <= 0)
                {
                    return "again";
                }
                else
                    return "Tv_Sony";
            }
            else if (value == 4)
            {
                ilosci[1].tv_Sha--;

                if (ilosci[1].tv_Sha <= 0)
                {
                    return "again";
                }
                else
                    return "Tv_Sharp";
            }
            else if (value == 5)
            {
                ilosci[1].laptop--;

                if (ilosci[1].laptop <= 0)
                {
                    return "again";
                }
                else
                    return "Laptop";
            }
            else if (value == 6)
            {
                ilosci[1].tel_Sam--;

                if (ilosci[1].tel_Sam <= 0)
                {
                    return "again";
                }
                else
                    return "Tel_Samsung";
            }
            else if (value == 7)
            {
                ilosci[1].tel_Mot--;

                if (ilosci[1].tel_Mot <= 0)
                {
                    return "again";
                }
                else
                    return "Tel_Motorola";
            }
            else if (value == 8)
            {
                ilosci[1].pra_Sam--;

                if (ilosci[1].pra_Sam <= 0)
                {
                    return "again";
                }
                else
                    return "Pralka_Samsung";
            }
            else if (value == 9)
            {
                ilosci[1].pra_Whi--;

                if (ilosci[1].pra_Whi <= 0)
                {
                    return "again";
                }
                else
                    return "Pralka_Whirpool";
            }
            else if (value == 10)
            {
                ilosci[1].kuc_Ami--;

                if (ilosci[1].kuc_Ami <= 0)
                {
                    return "again";
                }
                else
                    return "Kuchenka_Amica";
            }
            else if (value == 11)
            {
                ilosci[1].lod_Sam--;

                if (ilosci[1].lod_Sam <= 0)
                {
                    return "again";
                }
                else
                    return "Lodowka_Samsung";
            }
            else if (value == 12)
            {
                ilosci[1].lod_Bek--;

                if (ilosci[1].lod_Bek <= 0)
                {
                    return "again";
                }
                else
                    return "Lodowka_Beko";
            }
            else if (value == 13)
            {
                ilosci[1].susz--;

                if (ilosci[1].susz <= 0)
                {
                    return "again";
                }
                else
                    return "Suszarka";
            }
            else if (value == 14)
            {
                ilosci[1].oczysz--;

                if (ilosci[1].oczysz <= 0)
                {
                    return "again";
                }
                else
                    return "Oczyszczacz";
            }
            else if (value == 15)
            {
                ilosci[1].odk--;

                if (ilosci[1].odk <= 0)
                {
                    return "again";
                }
                else
                    return "Odkurzacz";
            }
            else if (value == 16)
            {
                ilosci[1].eksp--;

                if (ilosci[1].eksp <= 0)
                {
                    return "again";
                }
                else
                    return "Ekspres";
            }
            else if (value == 17)
            {
                ilosci[1].szczot--;

                if (ilosci[1].szczot <= 0)
                {
                    return "again";
                }
                else
                    return "Szczoteczka";
            }
            else if (value == 18)
            {
                ilosci[1].paro--;

                if (ilosci[1].paro <= 0)
                {
                    return "again";
                }
                else
                    return "Parownica";
            }
            else if (value == 19)
            {
                System.Console.WriteLine("Pusty if");
            }
            else if (value == 20)
            {
                System.Console.WriteLine("Pusty if");
            }

            return napis;
        }
    }

}