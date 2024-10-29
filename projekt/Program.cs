// Importowanie przestrzeni nazw
using System;
using System.Collections.Generic;
using System.Windows.Forms;

// Definicja przestrzeni nazw projektu
namespace projekt
{
    // Definicja statycznej klasy Program, która zawiera g³ówn¹ logikê aplikacji
    internal static class Program
    {
        // S³ownik przechowuj¹cy mapowanie identyfikatora formularza na numer formularza
        public static readonly Dictionary<int, int> FormNumberMapping = new Dictionary<int, int>
        {
            // Inicjalizacja s³ownika z okreœlonymi wartoœciami
        {0,10},
        {1,11},
        {2,12},
        {3,13},
        {4,14},
        {5,15},
        {6,16},
        {7,17},
        {8,18},
        {9,19},
        {10,20},
        {11,25},
        {12,26},
        {13,27},
        {14,28},
        {15,29},
        {16,30},
        {17,31},
        {18,32},
        {19,33},
        {20,34},
        {21,35},
        {22,36},
        {23,37},
        {24,38},
        {25,39},
        {26,40},
        {27,43},
        {28,44},
        {29,45},
        {30,46},
        {31,47},
        {32,48},
        {33,49},
        {34,50},
        {35,51},
        {36,52},
        {37,53},
        {38,54},
        {39,57},
        {40,58},
        {41,59},
        {42,60},
        {43,61},
        {44,62},
        {45,63},
        {46,64},
        {47,65},
        {48,66},
        {49,67},
        {50,68},
        {51,69},
        {52,71},
        {53,72},
        {54,73},
        {55,74},
        {56,75},
        {57,76},
        {58,77},
        {59,78},
        {60,79},
        {61,80},
        {62,81},
        {63,82},
        {64,84},
        {65,85},
        {66,87},
        {100,10},
        {101,11},
        {102,12},
        {103,13},
        {104,14},
        {105,15},
        {106,16},
        {107,17},
        {108,18},
        {109,19},
        {110,24},
        {111,25},
        {112,26},
        {113,27},
        {114,28},
        {115,29},
        {116,30},
        {117,31},
        {118,32},
        {119,33},
        {120,34},
        {121,35},
        {122,36},
        {123,37},
        {124,38},
        {125,39},
        {126,42},
        {127,43},
        {128,44},
        {129,45},
        {130,46},
        {131,47},
        {132,48},
        {133,49},
        {134,50},
        {135,51},
        {136,52},
        {137,53},
        {138,56},
        {139,57},
        {140,58},
        {141,59},
        {142,60},
        {143,61},
        {144,62},
        {145,63},
        {146,64},
        {147,65},
        {148,66},
        {149,67},
        {150,68},
        {151,70},
        {152,71},
        {153,72},
        {154,73},
        {155,74},
        {156,75},
        {157,76},
        {158,77},
        {159,78},
        {160,79},
        {161,80},
        {162,81},
        {163,83},
        {164,84},
        {165,86}
        };

        // Metoda zwracaj¹ca numer formularza na podstawie podanego identyfikatora formularza
        public static int GetFormNumber(int formId)
        {
            // Próba pobrania wartoœci z mapowania, jeœli istnieje
            if (FormNumberMapping.TryGetValue(formId, out int formNumber))
            {
                return formNumber; // Zwrócenie numeru formularza
            }
            else
            {
                // Rzucenie wyj¹tku, jeœli identyfikator formularza nie istnieje w mapowaniu
                throw new KeyNotFoundException($"Formularz o identyfikatorze {formId} nie zosta³ znaleziony.");
            }
        }

        // G³ówna metoda aplikacji
        [STAThread]
        static void Main()
        {
            // Konfiguracja stylów wizualnych aplikacji
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ApplicationConfiguration.Initialize(); // Inicjalizacja konfiguracji aplikacji
            Application.Run(new Form1()); // Uruchomienie g³ównego okna aplikacji
        }
    }
}
